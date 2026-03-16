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
import json
import os
import shutil
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from collections import defaultdict
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
WIN_WIDTH    = 580
WIN_HEIGHT   = 700
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


def load_config(path: Path) -> dict:
    """Load config (JSON). Falls back to old plain-text voice-name format."""
    try:
        text = path.read_text(encoding="utf-8").strip()
        if text.startswith("{"):
            return json.loads(text)
        return {"voice_name": text}  # backward compat
    except FileNotFoundError:
        return {}


def save_config(path: Path, cfg: dict):
    path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")


def get_input_devices() -> list[tuple[int, str]]:
    """Return (index, name) for each input-capable audio device."""
    result = []
    for i, d in enumerate(sd.query_devices()):
        if d["max_input_channels"] > 0:
            result.append((i, d["name"]))
    return result


# ---------------------------------------------------------------------------
# Audio post-processing
# ---------------------------------------------------------------------------
def apply_noise_gate(audio: np.ndarray, sample_rate: int) -> np.ndarray:
    """Silence low-level segments (breaths, ambient noise) between speech.

    Adaptive threshold set -30 dB below the peak RMS, with a 150 ms hold
    time to avoid chopping word tails and a 30 ms look-ahead so consonant
    onsets aren't clipped.  Gate boundaries get a 5 ms crossfade.
    """
    WINDOW_MS  = 20      # analysis window length
    THRESH_DB  = -30     # dB below peak RMS
    HOLD_MS    = 150     # keep gate open after last loud frame
    ATTACK_MS  = 30      # open gate this far before speech onset
    FADE_MS    = 5       # crossfade at open / close boundaries

    orig_shape = audio.shape
    audio = audio.ravel()                    # flatten to 1-D

    win = int(sample_rate * WINDOW_MS / 1000)
    n_frames = len(audio) // win
    if n_frames < 3:
        return audio.reshape(orig_shape)

    af = audio.astype(np.float64)

    # RMS per analysis window
    rms = np.array([
        np.sqrt(np.mean(af[i * win:(i + 1) * win] ** 2))
        for i in range(n_frames)
    ])

    peak = np.max(rms)
    if peak < 1.0:
        return audio.reshape(orig_shape)  # entire clip is silence

    thresh = peak * 10 ** (THRESH_DB / 20.0)

    # --- gate open / closed per frame ---
    gate = rms >= thresh

    # Hold: keep open for HOLD_MS after last above-threshold frame
    hold_n = max(1, int(HOLD_MS / WINDOW_MS))
    held = gate.copy()
    countdown = 0
    for i in range(n_frames):
        if gate[i]:
            countdown = hold_n
        elif countdown > 0:
            held[i] = True
            countdown -= 1

    # Attack look-back: open gate a few frames before each speech onset
    attack_n = max(1, int(ATTACK_MS / WINDOW_MS))
    for i in range(1, n_frames):
        if held[i] and not held[i - 1]:
            for j in range(max(0, i - attack_n), i):
                held[j] = True

    # --- apply gate with crossfades ---
    fade_n = max(1, int(sample_rate * FADE_MS / 1000))
    result = audio.copy()

    for i in range(n_frames):
        if not held[i]:
            s = i * win
            e = min(s + win, len(result))
            result[s:e] = 0

    # smooth open/close boundaries
    for i in range(1, n_frames):
        boundary = i * win
        if held[i] == held[i - 1]:
            continue
        fn = min(fade_n, boundary, len(result) - boundary)
        if fn <= 0:
            continue
        if held[i]:                                  # gate opening
            ramp = np.linspace(0.0, 1.0, fn)
            result[boundary:boundary + fn] = (
                audio[boundary:boundary + fn].astype(np.float64) * ramp
            ).astype(audio.dtype)
        else:                                        # gate closing
            ramp = np.linspace(1.0, 0.0, fn)
            result[boundary - fn:boundary] = (
                audio[boundary - fn:boundary].astype(np.float64) * ramp
            ).astype(audio.dtype)

    return result.reshape(orig_shape)


def trim_tail_click(audio: np.ndarray, sample_rate: int) -> np.ndarray:
    """Detect and remove a keypress / mouse-click transient in the last 0.5 s.

    Scans the tail of the recording in 5 ms windows.  If a window's peak
    amplitude exceeds the local median by 6 ×, everything from that point
    onward is trimmed and a 10 ms fade-out is applied.
    """
    SEARCH_MS = 500
    RATIO     = 6.0       # click must be this much louder than the local level
    FADE_MS   = 10

    orig_shape = audio.shape
    audio = audio.ravel()                            # flatten to 1-D

    search_n = int(sample_rate * SEARCH_MS / 1000)
    if len(audio) < search_n * 2:
        return audio.reshape(orig_shape)             # too short to trim safely

    tail = np.abs(audio[-search_n:].astype(np.float64))

    # 5 ms peak-amplitude windows
    win = max(1, int(sample_rate * 0.005))
    n = len(tail) // win
    if n < 4:
        return audio.reshape(orig_shape)

    peaks = np.array([np.max(tail[i * win:(i + 1) * win]) for i in range(n)])

    # Reference level: median of first half of search region (speech tail)
    ref = float(np.median(peaks[: n // 2]))
    if ref < 100:
        ref = 100.0                                  # floor for very quiet tails

    # Scan backwards for a transient
    for i in range(n - 1, n // 2, -1):
        if peaks[i] > ref * RATIO:
            onset = max(0, i - 2)                    # back up a couple of windows
            cut = len(audio) - search_n + onset * win
            trimmed = audio[:cut].copy()
            # apply fade-out
            fade_n = min(int(sample_rate * FADE_MS / 1000), len(trimmed))
            if fade_n > 0:
                ramp = np.linspace(1.0, 0.0, fade_n)
                trimmed[-fade_n:] = (
                    trimmed[-fade_n:].astype(np.float64) * ramp
                ).astype(audio.dtype)
            return trimmed.reshape(-1, *orig_shape[1:])

    return audio.reshape(orig_shape)                 # no click detected


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

        # Config
        self.cfg = load_config(CONFIG_PATH)

        # Audio devices
        self.input_devices = get_input_devices()  # [(index, name), ...]

        # Recording state
        self.is_recording = False
        self.is_paused = False
        self.rec_frames: list[np.ndarray] = []
        self.rec_stream = None
        self.rec_start_time = 0.0
        self.rec_pause_start = 0.0       # when current pause began
        self.rec_paused_total = 0.0      # accumulated paused seconds
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
        self.all_dlg_files = sorted({r["dlg_file"] for r in rows})

        # Conversation order: dlg_file -> rows sorted by conv_position
        self.dlg_conv_order: dict[str, list[dict]] = {}
        for r in rows:
            self.dlg_conv_order.setdefault(r["dlg_file"], []).append(r)
        for k in self.dlg_conv_order:
            self.dlg_conv_order[k].sort(
                key=lambda r: int(r.get("conv_position", 0)))

        # Duplicate-text index: text -> list of rows (only for duplicated lines)
        _text_groups: dict[str, list[dict]] = defaultdict(list)
        for r in rows:
            _text_groups[r["text"]].append(r)
        self.text_duplicates = {t: rs for t, rs in _text_groups.items()
                                if len(rs) > 1}

        self._build_ui()
        self._bind_keys()
        self._rebuild_area_dropdown()
        self._rebuild_char_dropdown()
        self._update_dlg_dropdown()
        self._apply_filter()
        # Start on the first unrecorded line when skip mode is on
        if self.unrecorded_var.get():
            for i, r in enumerate(self.filtered_rows):
                if not self._wav_path(r).exists():
                    self.idx = i
                    self._update_line_browser()
                    self._update_display()
                    break
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

        self.voice_var = tk.StringVar(value=self.cfg.get("voice_name", ""))
        self.voice_entry = tk.Entry(
            voice_row, textvariable=self.voice_var, font=("Segoe UI", 10),
            bg=BG_MID, fg=GOLD, insertbackground=GOLD, relief="flat",
            width=30)
        self.voice_entry.pack(side="left", padx=(6, 0), ipady=2)
        # Save name whenever it changes
        self.voice_var.trace_add("write", lambda *_: self._save_cfg())

        # --- Input device selector ---
        dev_row = tk.Frame(self, bg=BG)
        dev_row.pack(fill="x", padx=10, pady=(2, 2))

        tk.Label(dev_row, text="Mic:", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=FG_DIM).pack(side="left")

        device_names = [name for _, name in self.input_devices]
        self.dev_var = tk.StringVar()
        self.dev_cb = ttk.Combobox(dev_row, textvariable=self.dev_var,
                                   values=device_names, state="readonly",
                                   width=45)
        self.dev_cb.pack(side="left", padx=(6, 0))
        self.dev_cb.bind("<<ComboboxSelected>>", lambda e: self._save_cfg())

        # Restore saved device or fall back to system default
        saved_dev = self.cfg.get("input_device", "")
        if saved_dev in device_names:
            self.dev_var.set(saved_dev)
        elif device_names:
            self.dev_var.set(device_names[0])

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

        # --- Context: previous line ---
        self.lbl_ctx_prev = tk.Label(
            self, text="", font=("Segoe UI", 8), bg=BG, fg=FG_DIM,
            anchor="w", wraplength=WIN_WIDTH - 30)
        self.lbl_ctx_prev.pack(fill="x", padx=14, pady=(2, 0))

        # --- Text display ---
        txt_frame = tk.Frame(self, bg=BG_DARK, bd=1, relief="solid")
        txt_frame.pack(fill="both", expand=True, padx=10, pady=2)

        self.txt = tk.Text(
            txt_frame, wrap="word", font=("Georgia", 11), height=TEXT_HEIGHT,
            bg=BG_DARK, fg=FG, insertbackground=FG, relief="flat",
            padx=10, pady=8, state="disabled", cursor="arrow")
        self.txt.pack(fill="both", expand=True)

        # --- Context: next line ---
        self.lbl_ctx_next = tk.Label(
            self, text="", font=("Segoe UI", 8), bg=BG, fg=FG_DIM,
            anchor="w", wraplength=WIN_WIDTH - 30)
        self.lbl_ctx_next.pack(fill="x", padx=14, pady=(0, 2))

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

        # --- Processing options ---
        proc = tk.Frame(self, bg=BG)
        proc.pack(fill="x", padx=10, pady=(0, 2))

        _cb = dict(font=("Segoe UI", 8), bg=BG, fg=FG_DIM,
                   selectcolor=BG_MID, activebackground=BG,
                   activeforeground=FG)

        self.gate_var = tk.BooleanVar(
            value=self.cfg.get("noise_gate", True))
        tk.Checkbutton(
            proc, text="Noise gate", variable=self.gate_var,
            command=self._save_cfg, **_cb
        ).pack(side="left")

        self.trim_var = tk.BooleanVar(
            value=self.cfg.get("trim_click", True))
        tk.Checkbutton(
            proc, text="Trim click", variable=self.trim_var,
            command=self._save_cfg, **_cb
        ).pack(side="left", padx=(8, 0))

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
        self.char_cb.bind("<<ComboboxSelected>>", lambda e: self._on_char_changed())
        self._bind_key_jump(self.char_cb)

        self.unrecorded_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            filt, text="Skip recorded \u2192", variable=self.unrecorded_var,
            command=self._update_display, font=("Segoe UI", 8),
            bg=BG, fg=FG_DIM, selectcolor=BG_MID,
            activebackground=BG, activeforeground=FG
        ).pack(side="right")

        # --- Dialogue filter ---
        dlg_row = tk.Frame(self, bg=BG)
        dlg_row.pack(fill="x", padx=10, pady=(0, 2))

        tk.Label(dlg_row, text="DLG:", font=("Segoe UI", 9),
                 bg=BG, fg=FG_DIM).pack(side="left")

        self.dlg_var = tk.StringVar(value="All Dialogues")
        self.dlg_cb = ttk.Combobox(dlg_row, textvariable=self.dlg_var,
                                   values=["All Dialogues"] + self.all_dlg_files,
                                   state="readonly", width=52)
        self.dlg_cb.pack(side="left", padx=(4, 0))
        self.dlg_cb.bind("<<ComboboxSelected>>", lambda e: self._on_dlg_changed())
        self._bind_key_jump(self.dlg_cb)

        # --- Text search ---
        search_row = tk.Frame(self, bg=BG)
        search_row.pack(fill="x", padx=10, pady=(0, 2))

        tk.Label(search_row, text="Find:", font=("Segoe UI", 9),
                 bg=BG, fg=FG_DIM).pack(side="left")

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            search_row, textvariable=self.search_var, font=("Segoe UI", 9),
            bg=BG_MID, fg=FG, insertbackground=FG, relief="flat", width=40)
        self.search_entry.pack(side="left", padx=(4, 0), ipady=2)
        self.search_var.trace_add("write", lambda *_: self._apply_filter())

        self.lbl_search_count = tk.Label(
            search_row, text="", font=("Segoe UI", 8),
            bg=BG, fg=FG_DIM, anchor="w")
        self.lbl_search_count.pack(side="left", padx=(6, 0))

        # --- Line browser ---
        line_row = tk.Frame(self, bg=BG)
        line_row.pack(fill="x", padx=10, pady=(0, 2))

        tk.Label(line_row, text="Line:", font=("Segoe UI", 9),
                 bg=BG, fg=FG_DIM).pack(side="left")

        self.line_var = tk.StringVar()
        self.line_cb = ttk.Combobox(line_row, textvariable=self.line_var,
                                    state="readonly", width=52)
        self.line_cb.pack(side="left", padx=(4, 0))
        self.line_cb.bind("<<ComboboxSelected>>", lambda e: self._on_line_selected())

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
                if event.widget in (self.voice_entry, self.search_entry):
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
        self.bind("<Control_L>",  lambda e: self._toggle_pause())
        self.bind("<Control_R>",  lambda e: self._toggle_pause())

    # -----------------------------------------------------------------------
    # Display
    # -----------------------------------------------------------------------
    def _cur(self) -> dict | None:
        if not self.filtered_rows:
            return None
        return self.filtered_rows[self.idx]

    def _wav_path(self, row: dict) -> Path:
        return OUTPUT_DIR / f"{row['strref']}.wav"

    def _strip_check(self, text: str) -> str:
        """Remove the leading ✔ completion mark from a dropdown value."""
        return text[2:] if text.startswith("\u2714 ") else text

    def _get_context(self, row: dict) -> tuple[dict | None, dict | None]:
        """Return (prev_line, next_line) from the same dialogue."""
        conv = self.dlg_conv_order.get(row["dlg_file"], [])
        strref = row["strref"]
        idx = None
        for i, r in enumerate(conv):
            if r["strref"] == strref:
                idx = i
                break
        if idx is None:
            return None, None
        prev_r = conv[idx - 1] if idx > 0 else None
        next_r = conv[idx + 1] if idx < len(conv) - 1 else None
        return prev_r, next_r

    def _fmt_context(self, row: dict | None, arrow: str) -> str:
        """Format a context line for display."""
        if row is None:
            return ""
        text = row["text"].replace("\r", "").replace("\n", " ")[:65]
        return f"{arrow} [{row['character']}]  {text}"

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
            self.lbl_ctx_prev.config(text="")
            self.lbl_ctx_next.config(text="")
            return

        self.lbl_area.config(text=row["area"])
        self.lbl_char.config(text=row["character"])
        self.lbl_type.config(text=row["line_type"])
        dupes = self.text_duplicates.get(row["text"], [])
        dupe_tag = f"  (\u00d7{len(dupes)} identical)" if len(dupes) > 1 else ""
        self.lbl_strref.config(text=f"strref: {row['strref']}{dupe_tag}")

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
        if self.unrecorded_var.get():
            remaining = sum(1 for r in self.filtered_rows
                            if not self._wav_path(r).exists())
            if remaining == 0:
                extra = "  \u2714 all recorded"
            else:
                extra = f"  ({remaining} unrecorded)"
        else:
            extra = ""
        self.lbl_nav.config(
            text=f"{self.idx + 1} / {len(self.filtered_rows)}{extra}")

        # Sync line browser selection
        if hasattr(self, "line_cb") and self.filtered_rows and self.line_cb.cget("values"):
            self.line_cb.current(self.idx)

        # Conversation context
        prev_r, next_r = self._get_context(row)
        self.lbl_ctx_prev.config(text=self._fmt_context(prev_r, "\u25b2"))
        self.lbl_ctx_next.config(text=self._fmt_context(next_r, "\u25bc"))

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
                self.is_paused = False
                self._stop_record_internal()
            else:
                return
        if self.filtered_rows and self.idx > 0:
            self.idx -= 1
            self._update_display()

    def _next(self):
        if self.is_recording:
            if self.instant_rec_var.get():
                self.is_paused = False
                self._stop_record_internal()
            else:
                return
        if not self.filtered_rows:
            return
        if self.unrecorded_var.get():
            # Skip forward to the next unrecorded line
            for i in range(self.idx + 1, len(self.filtered_rows)):
                if not self._wav_path(self.filtered_rows[i]).exists():
                    self.idx = i
                    self._update_display()
                    if self.instant_rec_var.get():
                        self.after(50, self._start_record)
                    return
            # No unrecorded lines ahead — stay put
        else:
            if self.idx < len(self.filtered_rows) - 1:
                self.idx += 1
                self._update_display()
                if self.instant_rec_var.get():
                    self.after(50, self._start_record)

    # -----------------------------------------------------------------------
    # Filtering
    # -----------------------------------------------------------------------
    def _rebuild_area_dropdown(self):
        """Refresh area dropdown, adding ✔ to fully-recorded areas."""
        decorated = []
        for area in self.areas:
            rows = [r for r in self.all_rows if r["area"] == area]
            if rows and all(self._wav_path(r).exists() for r in rows):
                decorated.append(f"\u2714 {area}")
            else:
                decorated.append(area)
        cur_raw = self._strip_check(self.area_var.get())
        self.area_cb.config(values=["All Areas"] + decorated)
        for v in ["All Areas"] + decorated:
            if self._strip_check(v) == cur_raw:
                self.area_var.set(v)
                break

    def _rebuild_char_dropdown(self):
        """Refresh character dropdown, adding ✔ to fully-recorded characters."""
        area = self._strip_check(self.area_var.get())
        if area == "All Areas":
            chars = self.all_characters
            relevant = self.all_rows
        else:
            chars = self.area_characters.get(area, [])
            relevant = [r for r in self.all_rows if r["area"] == area]
        decorated = []
        for ch in chars:
            ch_rows = [r for r in relevant if r["character"] == ch]
            if ch_rows and all(self._wav_path(r).exists() for r in ch_rows):
                decorated.append(f"\u2714 {ch}")
            else:
                decorated.append(ch)
        cur_raw = self._strip_check(self.char_var.get())
        self.char_cb.config(values=["All Characters"] + decorated)
        for v in ["All Characters"] + decorated:
            if self._strip_check(v) == cur_raw:
                self.char_var.set(v)
                break

    def _on_area_changed(self):
        """When area filter changes, cascade: characters → dialogues → filter."""
        area = self._strip_check(self.area_var.get())
        if area == "All Areas":
            chars = self.all_characters
        else:
            chars = self.area_characters.get(area, [])

        # Reset character selection if current choice isn't in the new area
        cur_char = self._strip_check(self.char_var.get())
        if cur_char != "All Characters" and cur_char not in chars:
            self.char_var.set("All Characters")

        self._rebuild_char_dropdown()
        self._on_char_changed()

    def _on_char_changed(self):
        """When character filter changes, cascade: dialogues → filter."""
        self._update_dlg_dropdown()
        self._apply_filter()

    def _on_dlg_changed(self):
        """When dialogue filter changes, re-apply filter."""
        self._apply_filter()

    def _update_dlg_dropdown(self):
        """Rebuild DLG dropdown from current area/character selection."""
        area = self._strip_check(self.area_var.get())
        char = self._strip_check(self.char_var.get())
        relevant = [
            r for r in self.all_rows
            if (area == "All Areas" or r["area"] == area)
            and (char == "All Characters" or r["character"] == char)
        ]
        dlgs = sorted({r["dlg_file"] for r in relevant})
        decorated = []
        for d in dlgs:
            d_rows = [r for r in relevant if r["dlg_file"] == d]
            if d_rows and all(self._wav_path(r).exists() for r in d_rows):
                decorated.append(f"\u2714 {d}")
            else:
                decorated.append(d)
        cur_raw = self._strip_check(self.dlg_var.get())
        self.dlg_cb.config(values=["All Dialogues"] + decorated)
        if cur_raw != "All Dialogues" and cur_raw not in dlgs:
            self.dlg_var.set("All Dialogues")
        else:
            for v in ["All Dialogues"] + decorated:
                if self._strip_check(v) == cur_raw:
                    self.dlg_var.set(v)
                    break

    def _update_line_browser(self):
        """Rebuild the line dropdown from current filtered_rows."""
        previews = []
        for r in self.filtered_rows:
            mark = "\u2714" if self._wav_path(r).exists() else "  "
            text = r["text"].replace("\r", "").replace("\n", " ")[:45]
            previews.append(f"{mark} #{r['strref']}  {text}")
        self.line_cb.config(values=previews)
        if previews and self.idx < len(previews):
            self.line_cb.current(self.idx)
        elif not previews:
            self.line_var.set("")

    def _on_line_selected(self):
        """Jump to the line chosen in the line browser."""
        idx = self.line_cb.current()
        if 0 <= idx < len(self.filtered_rows):
            self.idx = idx
            self._update_display()

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
            if self._strip_check(v).upper().startswith(ch):
                combobox.current(i)
                combobox.event_generate("<<ComboboxSelected>>")
                return

    def _apply_filter(self):
        if self.is_recording:
            return

        area = self._strip_check(self.area_var.get())
        char = self._strip_check(self.char_var.get())
        dlg  = self._strip_check(self.dlg_var.get())
        query = self.search_var.get().strip().lower() if hasattr(self, "search_var") else ""

        self.filtered_rows = [
            r for r in self.all_rows
            if (area == "All Areas" or r["area"] == area)
            and (char == "All Characters" or r["character"] == char)
            and (dlg == "All Dialogues" or r["dlg_file"] == dlg)
            and (not query or query in r["text"].lower()
                 or query in r["character"].lower()
                 or query in r.get("strref", ""))
        ]
        self.idx = 0
        self._update_line_browser()
        self._update_display()

        # Update search hit count
        if hasattr(self, "lbl_search_count"):
            if query:
                self.lbl_search_count.config(
                    text=f"{len(self.filtered_rows)} hits")
            else:
                self.lbl_search_count.config(text="")

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
        self.is_paused = False
        self.rec_start_time = time.time()
        self.rec_pause_start = 0.0
        self.rec_paused_total = 0.0

        self.btn_rec.config(text="\u23f9  Stop", bg="#c0392b")
        self.lbl_timer.config(text="\u25cf 0:00.0", fg=ACCENT)
        self._tick_timer()

        try:
            self.rec_stream = sd.InputStream(
                samplerate=SAMPLE_RATE, channels=CHANNELS,
                dtype="int16", callback=self._audio_callback,
                blocksize=1024, device=self._selected_device_index())
            self.rec_stream.start()
        except Exception as exc:
            messagebox.showerror("Audio Error", f"Could not open microphone:\n{exc}")
            self.is_recording = False
            self.btn_rec.config(text="\u23fa  Record", bg=ACCENT)
            self.lbl_timer.config(text="")

    def _audio_callback(self, indata, frames, time_info, status):
        if not self.is_paused:
            self.rec_frames.append(indata.copy())

    def _tick_timer(self):
        if not self.is_recording:
            return
        if self.is_paused:
            elapsed = self.rec_pause_start - self.rec_start_time - self.rec_paused_total
        else:
            elapsed = time.time() - self.rec_start_time - self.rec_paused_total
        mins = int(elapsed) // 60
        secs = elapsed % 60
        if self.is_paused:
            self.lbl_timer.config(text=f"\u23f8 {mins}:{secs:04.1f}", fg=GOLD)
        else:
            self.lbl_timer.config(text=f"\u25cf {mins}:{secs:04.1f}", fg=ACCENT)
        self.timer_id = self.after(100, self._tick_timer)

    def _toggle_pause(self):
        """Pause / unpause the current recording (Ctrl key)."""
        if not self.is_recording:
            return
        if self.is_paused:
            # Resume: accumulate how long this pause lasted
            self.rec_paused_total += time.time() - self.rec_pause_start
            self.is_paused = False
            self.btn_rec.config(text="\u23f9  Stop", bg="#c0392b")
        else:
            # Pause: note the moment
            self.rec_pause_start = time.time()
            self.is_paused = True
            self.btn_rec.config(text="\u23f8 Paused", bg=BG_DARK)

    def _stop_record(self):
        """Stop recording and save."""
        self.is_paused = False
        self._stop_record_internal()

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

            # --- Post-processing ---
            if self.trim_var.get():
                audio = trim_tail_click(audio, SAMPLE_RATE)
            if self.gate_var.get():
                audio = apply_noise_gate(audio, SAMPLE_RATE)

            row = self._cur()
            OUTPUT_DIR.mkdir(exist_ok=True)
            path = self._wav_path(row)
            sf.write(str(path), audio, SAMPLE_RATE, subtype=SUBTYPE)

            # Stamp voice_file and voice_actor into the CSV row
            row["voice_file"] = f"{row['strref']}.wav"
            actor = self.voice_var.get().strip()
            if actor:
                row["voice_actor"] = actor

            # Propagate to rows with identical text
            dupe_count = 0
            dupes = self.text_duplicates.get(row["text"], [])
            for dupe_row in dupes:
                if dupe_row is row:
                    continue
                if self._wav_path(dupe_row).exists():
                    continue
                shutil.copy2(str(path), str(self._wav_path(dupe_row)))
                dupe_row["voice_file"] = f"{dupe_row['strref']}.wav"
                if actor:
                    dupe_row["voice_actor"] = actor
                dupe_count += 1

            # Persist CSV and regenerate credits
            save_csv(CSV_PATH, self.all_rows)
            generate_credits(self.all_rows, CREDITS_PATH)

            elapsed = time.time() - self.rec_start_time
            mins = int(elapsed) // 60
            secs = elapsed % 60
            dupe_tag = f"  (+{dupe_count} dupes)" if dupe_count else ""
            self.lbl_timer.config(
                text=f"\u2714 {mins}:{secs:04.1f}{dupe_tag}", fg=GREEN)
            self._update_display()
            self._update_progress()
            self._update_line_browser()
            self._rebuild_area_dropdown()
            self._rebuild_char_dropdown()
            self._update_dlg_dropdown()

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

    def _save_cfg(self):
        self.cfg["voice_name"] = self.voice_var.get().strip()
        self.cfg["input_device"] = self.dev_var.get()
        self.cfg["noise_gate"] = self.gate_var.get()
        self.cfg["trim_click"] = self.trim_var.get()
        save_config(CONFIG_PATH, self.cfg)

    def _selected_device_index(self) -> int | None:
        """Return the sounddevice index for the currently selected mic."""
        name = self.dev_var.get()
        for idx, n in self.input_devices:
            if n == name:
                return idx
        return None

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
