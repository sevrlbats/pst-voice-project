"""
PST Voice Booth - Compact recording tool for Planescape: Torment voiceover mod.
Records WAV files at 22050 Hz / 16-bit / mono (native IE engine format).
Files are named by strref for direct mod integration.

Keyboard shortcuts:
  Space    - Record / Stop recording
  R        - Record / Stop recording
  P        - Play back current recording
  Left / A - Previous line
  Right / D - Next line
  Escape   - Stop playback
"""

import csv
import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

import numpy as np
import sounddevice as sd
import soundfile as sf

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
SAMPLE_RATE  = 22050
CHANNELS     = 1
SUBTYPE      = "PCM_16"
OUTPUT_DIR   = Path(__file__).parent / "voice_recordings"
CSV_PATH     = Path(__file__).parent / "unvoiced_dialog.csv"
CONFIG_PATH  = Path(__file__).parent / "voice_booth.cfg"
CREDITS_PATH = Path(__file__).parent / "voice_recordings" / "CREDITS.txt"

# CSV columns (voice_actor is appended if missing)
CSV_FIELDS   = ["area", "dlg_file", "character", "line_type",
                "conv_position", "strref", "voice_file", "voice_actor", "text"]

# UI sizing
WIN_WIDTH    = 520
WIN_HEIGHT   = 500
TEXT_HEIGHT  = 7       # lines visible in text box

# Colors
BG           = "#1a1a2e"
BG_MID       = "#16213e"
BG_DARK      = "#0f3460"
FG           = "#e0e0e0"
FG_DIM       = "#8888aa"
ACCENT       = "#e94560"
ACCENT_GLOW  = "#ff6b81"
GREEN        = "#2ecc71"
GOLD         = "#f1c40f"

# ---------------------------------------------------------------------------
# Area ordering — loose game progression
# ---------------------------------------------------------------------------
AREA_ORDER = [
    # --- Act 1: Mortuary ---
    "Mortuary - 1st Floor",
    "Mortuary - 2nd Floor",
    "Mortuary - 3rd Floor",
    "Mortuary",
    # --- Act 1: Hive ---
    "Hive - Mortuary Area",
    "Mortuary / Gathering Dust Bar",
    "Gathering Dust Bar",
    "Mausoleum - 1st Section",
    "Mausoleum - 2nd Section",
    "Mausoleum - Inner Sanctum",
    "Hive - Marketplace",
    "Hive (Generic)",
    "Hive (Various)",
    "Small Dwelling (Marketplace)",
    "Arlo's Kip",
    "Fell's Tattoo Parlor",
    "Hive - Smoldering Corpse Area",
    "Smoldering Corpse Bar",
    "Alley of Dangerous Angles",
    "Alley of Lingering Sighs",
    "Tenement of Thugs",
    "Vermin & Disease Control",
    # --- Act 1: Ragpicker's / Buried Village ---
    "Hive - NW / Ragpicker's Area",
    "Ragpicker's Square",
    "Midwife's Hut (Mebbeth)",
    "Angyar's Kip",
    "Ojo's House",
    "Seamstress' House",
    "Ragpicker's Square / Buried Village",
    "Trash Warrens",
    "Buried Village",
    "Pharod's Court",
    # --- Act 1: Catacombs ---
    "Weeping Stone Catacombs",
    "Dead Nations",
    "Drowned Nations",
    "Warrens of Thought",
    "Warrens of Thought - Many-as-One",
    # --- Act 2: Lower Ward ---
    "Lower Ward",
    "Lower Ward - Marketplace",
    "Siege Tower",
    "Godsmen Foundry",
    "Godsmen Foundry - Worker Area",
    "Godsmen Foundry - Assembly Hall",
    "Godsmen Foundry - Sandoz's Area",
    "Godsmen Foundry - Cannon Room",
    "Pawn Shop",
    "Coffin Maker's Shop",
    "Anarchists' Printing Shop",
    "Vault of the Ninth World",
    # --- Act 2: Clerk's Ward ---
    "Clerk's Ward",
    "Civic Festhall - Main Hall",
    "Civic Festhall - Resting Quarters",
    "Civic Festhall - Public Sensorium",
    "Civic Festhall - Private Sensorium",
    "Civic Festhall",
    "Brothel of Slaking Intellectual Lusts",
    "Brothel - Cellar",
    "Advocate's Home",
    "Art Store",
    "Curiosity Shoppe (Vrischika's)",
    "Tailor",
    "Lothar's Home - Skull Room",
    # --- Act 3: Ravel ---
    "Ravel's Maze",
    # --- Act 3: Curst / Outlands ---
    "Curst",
    "Curst - Outer",
    "Curst - Inner",
    "Traitor's Gate Tavern",
    "Curst - Prison",
    "Curst - Underground",
    "Carceri",
    # --- Act 3: Baator ---
    "Baator - Pillar of Skulls",
    # --- Act 4: Fortress ---
    "Fortress of Regrets - Entrance",
    "Fortress of Regrets - Roof",
    "Fortress of Regrets",
    "Fortress of Regrets - Maze of Reflections",
    # --- Companion / global ---
    "Companion Dialog (Any Area)",
    "Companion Dialog (Circle of Zerthimon)",
    # --- Unmapped raw area codes & system (end of list) ---
]

# Companion DLG files that belong to a specific area rather than "Any Area"
COMPANION_DLG_AREA = {
    "DMORTE1": "Mortuary",
    "DMORTE2": "Mortuary",
    "DANNAHF": "Fortress of Regrets",
    "DGRACEF": "Fortress of Regrets",
}


def _reclassify_companion_dlgs(rows: list[dict]):
    """Move area-specific companion dialogues out of the catch-all bucket."""
    for r in rows:
        if r["area"] == "Companion Dialog (Any Area)":
            new_area = COMPANION_DLG_AREA.get(r["dlg_file"])
            if new_area:
                r["area"] = new_area


def _build_area_key(areas_in_data: set[str]) -> list[str]:
    """Return areas in game order. Any areas in the data not listed in
    AREA_ORDER are appended alphabetically at the end."""
    ordered = [a for a in AREA_ORDER if a in areas_in_data]
    remainder = sorted(areas_in_data - set(ordered))
    return ordered + remainder


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
def load_csv(path: Path) -> list[dict]:
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    # Ensure voice_actor column exists (older CSVs won't have it)
    for r in rows:
        r.setdefault("voice_actor", "")
    return rows


def save_csv(path: Path, rows: list[dict]):
    """Write rows back to CSV, preserving column order."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def load_voice_name(path: Path) -> str:
    """Load the last-used voice actor name from config."""
    try:
        return path.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return ""


def save_voice_name(path: Path, name: str):
    path.write_text(name.strip(), encoding="utf-8")


def generate_credits(rows: list[dict], path: Path):
    """Generate a CREDITS.txt from voice_actor data in the CSV."""
    # Collect actor -> set of characters voiced
    actor_roles: dict[str, set[str]] = {}
    for r in rows:
        actor = r.get("voice_actor", "").strip()
        if not actor:
            continue
        actor_roles.setdefault(actor, set()).add(r["character"])

    if not actor_roles:
        return

    lines = []
    lines.append("=" * 60)
    lines.append("  PLANESCAPE: TORMENT — VOICE MOD CREDITS")
    lines.append("=" * 60)
    lines.append("")
    lines.append("This mod adds voiceover to previously unvoiced dialogue")
    lines.append("in Planescape: Torment. The following voice actors")
    lines.append("generously contributed their talents:")
    lines.append("")

    for actor in sorted(actor_roles):
        chars = sorted(actor_roles[actor])
        lines.append(f"  {actor}")
        # Group into lines of ~3 characters each for readability
        for i in range(0, len(chars), 3):
            chunk = ", ".join(chars[i:i+3])
            prefix = "    as " if i == 0 else "       "
            lines.append(f"{prefix}{chunk}")
        lines.append("")

    total_voiced = sum(1 for r in rows if r.get("voice_actor", "").strip())
    lines.append(f"  Total lines voiced: {total_voiced:,}")
    lines.append("")
    lines.append("=" * 60)

    path.parent.mkdir(exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
class VoiceBooth(tk.Tk):
    def __init__(self, rows: list[dict]):
        super().__init__()
        self.all_rows = rows
        self.filtered_rows = list(rows)
        self.idx = 0

        # Recording state
        self.is_recording = False
        self.rec_frames: list[np.ndarray] = []
        self.rec_stream = None
        self.rec_start_time = 0.0
        self.timer_id = None

        # Playback state
        self.is_playing = False

        # Unique filter values — areas in game order, characters per area
        all_areas = {r["area"] for r in rows}
        self.areas = _build_area_key(all_areas)

        # Map area -> sorted character list (for cascading filter)
        self.area_characters: dict[str, list[str]] = {}
        for r in rows:
            self.area_characters.setdefault(r["area"], set()).add(r["character"])
        for k in self.area_characters:
            self.area_characters[k] = sorted(self.area_characters[k])
        self.all_characters = sorted({r["character"] for r in rows})

        self._build_ui()
        self._bind_keys()
        self._update_display()
        self._update_progress()

    # -----------------------------------------------------------------------
    # UI construction
    # -----------------------------------------------------------------------
    def _build_ui(self):
        self.title("PST Voice Booth")
        self.configure(bg=BG)
        self.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # --- Top bar: pin toggle + title ---
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=10, pady=(8, 2))

        tk.Label(top, text="\U0001f399  PST Voice Booth", font=("Segoe UI", 13, "bold"),
                 bg=BG, fg=ACCENT).pack(side="left")

        self.pin_var = tk.BooleanVar(value=True)
        self.pin_btn = tk.Checkbutton(
            top, text="\U0001f4cc On Top", variable=self.pin_var,
            command=self._toggle_pin, font=("Segoe UI", 9),
            bg=BG, fg=FG_DIM, selectcolor=BG_MID,
            activebackground=BG, activeforeground=FG)
        self.pin_btn.pack(side="right")

        # --- Voice actor name ---
        voice_row = tk.Frame(self, bg=BG)
        voice_row.pack(fill="x", padx=10, pady=(2, 2))

        tk.Label(voice_row, text="Voice:", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=FG_DIM).pack(side="left")

        self.voice_var = tk.StringVar(value=load_voice_name(CONFIG_PATH))
        self.voice_entry = tk.Entry(
            voice_row, textvariable=self.voice_var, font=("Segoe UI", 10),
            bg=BG_MID, fg=GOLD, insertbackground=GOLD, relief="flat",
            width=30)
        self.voice_entry.pack(side="left", padx=(6, 0), ipady=2)
        # Save name whenever it changes
        self.voice_var.trace_add("write", lambda *_: save_voice_name(
            CONFIG_PATH, self.voice_var.get()))

        # --- Info bar: area | character | type ---
        info = tk.Frame(self, bg=BG_MID)
        info.pack(fill="x", padx=10, pady=(4, 2))

        self.lbl_area = tk.Label(info, text="", font=("Segoe UI", 9),
                                 bg=BG_MID, fg=FG_DIM, anchor="w")
        self.lbl_area.pack(side="left", padx=(6, 10))

        self.lbl_char = tk.Label(info, text="", font=("Segoe UI", 10, "bold"),
                                 bg=BG_MID, fg=GOLD, anchor="w")
        self.lbl_char.pack(side="left", padx=(0, 10))

        self.lbl_type = tk.Label(info, text="", font=("Segoe UI", 9),
                                 bg=BG_MID, fg=FG_DIM, anchor="e")
        self.lbl_type.pack(side="right", padx=6)

        # --- Text display ---
        txt_frame = tk.Frame(self, bg=BG_DARK, bd=1, relief="solid")
        txt_frame.pack(fill="both", expand=True, padx=10, pady=4)

        self.txt = tk.Text(
            txt_frame, wrap="word", font=("Georgia", 11), height=TEXT_HEIGHT,
            bg=BG_DARK, fg=FG, insertbackground=FG, relief="flat",
            padx=10, pady=8, state="disabled", cursor="arrow")
        self.txt.pack(fill="both", expand=True)

        # --- Strref + timer row ---
        meta = tk.Frame(self, bg=BG)
        meta.pack(fill="x", padx=10, pady=2)

        self.lbl_strref = tk.Label(meta, text="strref: —", font=("Consolas", 10),
                                    bg=BG, fg=FG_DIM, anchor="w")
        self.lbl_strref.pack(side="left")

        self.lbl_recorded = tk.Label(meta, text="", font=("Segoe UI", 9),
                                      bg=BG, fg=GREEN, anchor="w")
        self.lbl_recorded.pack(side="left", padx=(12, 0))

        self.lbl_timer = tk.Label(meta, text="", font=("Consolas", 11, "bold"),
                                   bg=BG, fg=ACCENT, anchor="e")
        self.lbl_timer.pack(side="right")

        # --- Transport controls ---
        transport = tk.Frame(self, bg=BG)
        transport.pack(fill="x", padx=10, pady=4)

        btn_style = dict(font=("Segoe UI", 10, "bold"), bd=0, relief="flat",
                         cursor="hand2", padx=12, pady=4)

        self.btn_prev = tk.Button(
            transport, text="\u23ee Prev", bg=BG_MID, fg=FG,
            activebackground=BG_DARK, activeforeground=FG,
            command=self._prev, **btn_style)
        self.btn_prev.pack(side="left", padx=(0, 4))

        self.btn_rec = tk.Button(
            transport, text="\u23fa  Record", bg=ACCENT, fg="white",
            activebackground=ACCENT_GLOW, activeforeground="white",
            command=self._toggle_record, width=12, **btn_style)
        self.btn_rec.pack(side="left", padx=4)

        self.btn_next = tk.Button(
            transport, text="Next \u23ed", bg=BG_MID, fg=FG,
            activebackground=BG_DARK, activeforeground=FG,
            command=self._next, **btn_style)
        self.btn_next.pack(side="left", padx=4)

        self.btn_play = tk.Button(
            transport, text="\u25b6 Play", bg=BG_MID, fg=FG,
            activebackground=BG_DARK, activeforeground=FG,
            command=self._play, **btn_style)
        self.btn_play.pack(side="right")

        self.instant_rec_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            transport, text="Instant rec", variable=self.instant_rec_var,
            font=("Segoe UI", 8),
            bg=BG, fg=FG_DIM, selectcolor=BG_MID,
            activebackground=BG, activeforeground=FG
        ).pack(side="right", padx=(0, 6))

        # --- Filter row ---
        filt = tk.Frame(self, bg=BG)
        filt.pack(fill="x", padx=10, pady=(2, 2))

        tk.Label(filt, text="Filter:", font=("Segoe UI", 9),
                 bg=BG, fg=FG_DIM).pack(side="left")

        self.area_var = tk.StringVar(value="All Areas")
        self.area_cb = ttk.Combobox(filt, textvariable=self.area_var, width=22,
                                    values=["All Areas"] + self.areas, state="readonly")
        self.area_cb.pack(side="left", padx=(4, 6))
        self.area_cb.bind("<<ComboboxSelected>>", lambda e: self._on_area_changed())
        self._bind_key_jump(self.area_cb)

        self.char_var = tk.StringVar(value="All Characters")
        self.char_cb = ttk.Combobox(filt, textvariable=self.char_var, width=18,
                                    values=["All Characters"] + self.all_characters,
                                    state="readonly")
        self.char_cb.pack(side="left", padx=(0, 6))
        self.char_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        self._bind_key_jump(self.char_cb)

        self.unrecorded_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            filt, text="Unrecorded only", variable=self.unrecorded_var,
            command=self._apply_filter, font=("Segoe UI", 8),
            bg=BG, fg=FG_DIM, selectcolor=BG_MID,
            activebackground=BG, activeforeground=FG
        ).pack(side="right")

        # --- Progress bar ---
        prog = tk.Frame(self, bg=BG)
        prog.pack(fill="x", padx=10, pady=(2, 8))

        self.lbl_progress = tk.Label(prog, text="", font=("Segoe UI", 9),
                                      bg=BG, fg=FG_DIM, anchor="w")
        self.lbl_progress.pack(side="left")

        self.lbl_nav = tk.Label(prog, text="", font=("Consolas", 9),
                                 bg=BG, fg=FG_DIM, anchor="e")
        self.lbl_nav.pack(side="right")

    # -----------------------------------------------------------------------
    # Key bindings
    # -----------------------------------------------------------------------
    def _bind_keys(self):
        # Guard: don't fire shortcuts while typing in the voice name field
        def _guard(action):
            def wrapper(event):
                if event.widget is self.voice_entry:
                    return  # let the Entry handle it normally
                action()
                return "break"
            return wrapper

        self.bind("<space>",  _guard(self._toggle_record))
        self.bind("r",        _guard(self._toggle_record))
        self.bind("R",        _guard(self._toggle_record))
        self.bind("<Left>",   _guard(self._prev))
        self.bind("<Right>",  _guard(self._next))
        self.bind("a",        _guard(self._prev))
        self.bind("A",        _guard(self._prev))
        self.bind("d",        _guard(self._next))
        self.bind("D",        _guard(self._next))
        self.bind("p",        _guard(self._play))
        self.bind("P",        _guard(self._play))
        self.bind("<Escape>", lambda e: self._stop_playback())

    # -----------------------------------------------------------------------
    # Display
    # -----------------------------------------------------------------------
    def _cur(self) -> dict | None:
        if not self.filtered_rows:
            return None
        return self.filtered_rows[self.idx]

    def _wav_path(self, row: dict) -> Path:
        return OUTPUT_DIR / f"{row['strref']}.wav"

    def _update_display(self):
        row = self._cur()
        if row is None:
            self.lbl_area.config(text="No lines match filter")
            self.lbl_char.config(text="")
            self.lbl_type.config(text="")
            self.txt.config(state="normal")
            self.txt.delete("1.0", "end")
            self.txt.config(state="disabled")
            self.lbl_strref.config(text="strref: —")
            self.lbl_recorded.config(text="")
            self.lbl_nav.config(text="0 / 0")
            return

        self.lbl_area.config(text=row["area"])
        self.lbl_char.config(text=row["character"])
        self.lbl_type.config(text=row["line_type"])
        self.lbl_strref.config(text=f"strref: {row['strref']}")

        # Show text
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", row["text"])
        self.txt.config(state="disabled")

        # Recorded indicator (show who voiced it if known)
        if self._wav_path(row).exists():
            actor = row.get("voice_actor", "").strip()
            if actor:
                self.lbl_recorded.config(text=f"\u2714 voiced by {actor}", fg=GREEN)
            else:
                self.lbl_recorded.config(text="\u2714 recorded", fg=GREEN)
        else:
            self.lbl_recorded.config(text="", fg=GREEN)

        # Nav counter
        self.lbl_nav.config(text=f"{self.idx + 1} / {len(self.filtered_rows)}")

    def _update_progress(self):
        total = len(self.all_rows)
        done = sum(1 for r in self.all_rows if self._wav_path(r).exists())
        pct = (done / total * 100) if total else 0
        self.lbl_progress.config(text=f"Progress: {done:,} / {total:,}  ({pct:.1f}%)")

    # -----------------------------------------------------------------------
    # Navigation
    # -----------------------------------------------------------------------
    def _prev(self):
        if self.is_recording:
            if self.instant_rec_var.get():
                self._stop_record_internal()
            else:
                return
        if self.filtered_rows and self.idx > 0:
            self.idx -= 1
            self._update_display()
            if self.instant_rec_var.get():
                self.after(50, self._start_record)

    def _next(self):
        if self.is_recording:
            if self.instant_rec_var.get():
                self._stop_record_internal()
            else:
                return
        if self.filtered_rows and self.idx < len(self.filtered_rows) - 1:
            self.idx += 1
            self._update_display()
            if self.instant_rec_var.get():
                self.after(50, self._start_record)

    # -----------------------------------------------------------------------
    # Filtering
    # -----------------------------------------------------------------------
    def _on_area_changed(self):
        """When area filter changes, update character dropdown to show only
        characters present in that area (cascading filter)."""
        area = self.area_var.get()
        if area == "All Areas":
            chars = self.all_characters
        else:
            chars = self.area_characters.get(area, [])

        self.char_cb.config(values=["All Characters"] + chars)

        # Reset character selection if current choice isn't in the new area
        if self.char_var.get() != "All Characters" and self.char_var.get() not in chars:
            self.char_var.set("All Characters")

        self._apply_filter()

    def _bind_key_jump(self, combobox: ttk.Combobox):
        """Bind key presses on a combobox so typing a letter jumps to the
        first item starting with that letter."""
        combobox.bind("<KeyPress>", lambda e: self._key_jump(e, combobox))

    def _key_jump(self, event, combobox: ttk.Combobox):
        """Jump to the first combobox value starting with the pressed key."""
        ch = event.char.upper()
        if not ch or not ch.isalpha():
            return
        values = combobox.cget("values")
        if not values:
            return
        for i, v in enumerate(values):
            # Skip "All Areas" / "All Characters" header entries
            if v.startswith("All "):
                continue
            if v.upper().startswith(ch):
                combobox.current(i)
                combobox.event_generate("<<ComboboxSelected>>")
                return

    def _apply_filter(self):
        if self.is_recording:
            return

        area = self.area_var.get()
        char = self.char_var.get()
        unrecorded = self.unrecorded_var.get()

        self.filtered_rows = [
            r for r in self.all_rows
            if (area == "All Areas" or r["area"] == area)
            and (char == "All Characters" or r["character"] == char)
            and (not unrecorded or not self._wav_path(r).exists())
        ]
        self.idx = 0
        self._update_display()

    # -----------------------------------------------------------------------
    # Recording
    # -----------------------------------------------------------------------
    def _toggle_record(self):
        if self._cur() is None:
            return
        if self.is_recording:
            self._stop_record()
        else:
            self._start_record()

    def _start_record(self):
        if self.is_playing:
            self._stop_playback()

        self.rec_frames = []
        self.is_recording = True
        self.rec_start_time = time.time()

        self.btn_rec.config(text="\u23f9  Stop", bg="#c0392b")
        self.lbl_timer.config(text="\u25cf 0:00.0", fg=ACCENT)
        self._tick_timer()

        try:
            self.rec_stream = sd.InputStream(
                samplerate=SAMPLE_RATE, channels=CHANNELS,
                dtype="int16", callback=self._audio_callback,
                blocksize=1024)
            self.rec_stream.start()
        except Exception as exc:
            messagebox.showerror("Audio Error", f"Could not open microphone:\n{exc}")
            self.is_recording = False
            self.btn_rec.config(text="\u23fa  Record", bg=ACCENT)
            self.lbl_timer.config(text="")

    def _audio_callback(self, indata, frames, time_info, status):
        self.rec_frames.append(indata.copy())

    def _tick_timer(self):
        if not self.is_recording:
            return
        elapsed = time.time() - self.rec_start_time
        mins = int(elapsed) // 60
        secs = elapsed % 60
        self.lbl_timer.config(text=f"\u25cf {mins}:{secs:04.1f}")
        self.timer_id = self.after(100, self._tick_timer)

    def _stop_record(self):
        """Stop recording and save. If instant record mode is on, auto-advance."""
        self._stop_record_internal()
        # In instant record mode, stopping via Space/button auto-advances
        if self.instant_rec_var.get():
            if self.filtered_rows and self.idx < len(self.filtered_rows) - 1:
                self.idx += 1
                self._update_display()
                self.after(50, self._start_record)

    def _stop_record_internal(self):
        """Stop recording and save the WAV file (no auto-advance)."""
        self.is_recording = False
        if self.timer_id:
            self.after_cancel(self.timer_id)
            self.timer_id = None

        # Tear down the audio stream safely
        stream = self.rec_stream
        self.rec_stream = None
        if stream:
            try:
                stream.stop()
                stream.close()
            except Exception:
                pass  # stream may already be closed

        self.btn_rec.config(text="\u23fa  Record", bg=ACCENT)

        # Grab frames collected so far
        frames = self.rec_frames
        self.rec_frames = []

        if not frames:
            self.lbl_timer.config(text="\u2716 too short", fg=ACCENT)
            return

        try:
            audio = np.concatenate(frames, axis=0)
            row = self._cur()
            OUTPUT_DIR.mkdir(exist_ok=True)
            path = self._wav_path(row)
            sf.write(str(path), audio, SAMPLE_RATE, subtype=SUBTYPE)

            # Stamp voice_file and voice_actor into the CSV row
            row["voice_file"] = f"{row['strref']}.wav"
            actor = self.voice_var.get().strip()
            if actor:
                row["voice_actor"] = actor

            # Persist CSV and regenerate credits
            save_csv(CSV_PATH, self.all_rows)
            generate_credits(self.all_rows, CREDITS_PATH)

            elapsed = time.time() - self.rec_start_time
            mins = int(elapsed) // 60
            secs = elapsed % 60
            self.lbl_timer.config(text=f"\u2714 {mins}:{secs:04.1f}", fg=GREEN)
            self._update_display()
            self._update_progress()

        except Exception as exc:
            self.lbl_timer.config(text=f"\u2716 save failed", fg=ACCENT)
            messagebox.showerror("Save Error",
                f"Recording captured but could not save:\n{exc}\n\n"
                f"Check write permissions to:\n{OUTPUT_DIR}\n{CSV_PATH}")

    # -----------------------------------------------------------------------
    # Playback
    # -----------------------------------------------------------------------
    def _play(self):
        row = self._cur()
        if row is None:
            return
        path = self._wav_path(row)
        if not path.exists():
            return

        if self.is_playing:
            self._stop_playback()
            return

        try:
            data, sr = sf.read(str(path), dtype="int16")
            self.is_playing = True
            self.btn_play.config(text="\u23f9 Stop", bg=BG_DARK)
            sd.play(data, sr)

            # Schedule reset when done
            duration_ms = int(len(data) / sr * 1000) + 100
            self.after(duration_ms, self._on_playback_done)
        except Exception as exc:
            messagebox.showerror("Playback Error", str(exc))

    def _stop_playback(self):
        sd.stop()
        self.is_playing = False
        self.btn_play.config(text="\u25b6 Play", bg=BG_MID)

    def _on_playback_done(self):
        if self.is_playing:
            self.is_playing = False
            self.btn_play.config(text="\u25b6 Play", bg=BG_MID)

    # -----------------------------------------------------------------------
    # Misc
    # -----------------------------------------------------------------------
    def _toggle_pin(self):
        self.attributes("-topmost", self.pin_var.get())

    def _on_close(self):
        if self.is_recording:
            self._stop_record()
        if self.is_playing:
            self._stop_playback()
        self.destroy()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not CSV_PATH.exists():
        print(f"ERROR: Cannot find {CSV_PATH}")
        sys.exit(1)

    rows = load_csv(CSV_PATH)
    _reclassify_companion_dlgs(rows)
    print(f"Loaded {len(rows):,} unvoiced lines from {CSV_PATH.name}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"Recordings will be saved to: {OUTPUT_DIR}")

    app = VoiceBooth(rows)
    app.mainloop()


if __name__ == "__main__":
    main()
