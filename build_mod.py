"""
build_mod.py - Build a WeiDU mod that patches voice into unvoiced dialog.

Reads unvoiced_dialog.csv for lines with voice_file filled in,
copies WAV files from voice_recordings/ into the mod package,
and generates a self-contained WeiDU .tp2 (no external includes).

Usage:
  python build_mod.py          # build the mod
  setup-PST_Voice_Mod.exe      # install it (WeiDU)
"""

import csv
import shutil
import sys
from pathlib import Path

import numpy as np
import soundfile as sf

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
PROJECT_DIR   = Path(__file__).parent
GAME_DIR      = Path(r"C:\Program Files (x86)\Steam\steamapps\common\Project P")
CSV_PATH      = PROJECT_DIR / "unvoiced_dialog.csv"
RECORDINGS    = PROJECT_DIR / "voice_recordings"
CREDITS_SRC   = RECORDINGS / "CREDITS.txt"

# Mod output goes to game dir (where WeiDU needs it)
MOD_DIR       = GAME_DIR / "PST_Voice_Mod"
MOD_AUDIO     = MOD_DIR / "audio"
TP2_PATH      = GAME_DIR / "setup-PST_Voice_Mod.tp2"

# Volume boost applied at build time (original recordings are untouched)
GAIN          = 9.0


def main():
    # ------------------------------------------------------------------
    # 1. Read CSV — find rows where voice_file is populated
    # ------------------------------------------------------------------
    with open(CSV_PATH, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    entries = []
    for r in rows:
        vf = r.get("voice_file", "").strip()
        if vf:
            strref = int(r["strref"])
            # Resref = filename without extension, max 8 chars
            resref = Path(vf).stem[:8]
            entries.append({
                "strref":   strref,
                "resref":   resref,
                "wav_name": vf if vf.endswith(".wav") else vf + ".wav",
                "char":     r["character"],
                "text":     r["text"],
                "preview":  r["text"][:60],
            })

    if not entries:
        print("ERROR: No voice_file entries found in unvoiced_dialog.csv!")
        print("Record some lines in Voice Booth first, then re-run.")
        sys.exit(1)

    print(f"Found {len(entries)} voiced lines in CSV.")

    # ------------------------------------------------------------------
    # 2. Clean and create mod directory structure
    # ------------------------------------------------------------------
    # Clean old generated files (but leave backup/ alone — WeiDU owns that)
    for subdir in ["audio", "lib"]:
        d = MOD_DIR / subdir
        if d.exists():
            shutil.rmtree(d, ignore_errors=True)
    for f in MOD_DIR.glob("*.tp2"):
        f.unlink(missing_ok=True)
    for f in MOD_DIR.glob("*.txt"):
        f.unlink(missing_ok=True)
    MOD_AUDIO.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # 3. Copy WAV files from voice_recordings/ into mod audio folder
    # ------------------------------------------------------------------
    copied = 0
    clipped = 0
    missing = []
    for e in entries:
        src = RECORDINGS / e["wav_name"]
        dst = MOD_AUDIO / f"{e['resref']}.wav"
        if src.exists():
            data, sr = sf.read(str(src), dtype="int16")
            boosted = data.astype(np.float64) * GAIN
            boosted = np.clip(boosted, -32768, 32767).astype(np.int16)
            if np.any(np.abs(data.astype(np.float64) * GAIN) > 32767):
                clipped += 1
            sf.write(str(dst), boosted, sr, subtype="PCM_16")
            copied += 1
        else:
            missing.append(f"  {e['wav_name']} (strref {e['strref']}, {e['char']})")

    if missing:
        print(f"\nWARNING: {len(missing)} audio files not found in {RECORDINGS}:")
        for m in missing[:20]:
            print(m)
        if len(missing) > 20:
            print(f"  ... and {len(missing) - 20} more")
        print()

    print(f"Copied {copied} audio files to {MOD_AUDIO} (gain: {GAIN}x)")
    if clipped:
        print(f"  Note: {clipped} file(s) had peaks clipped at {GAIN}x gain")

    # ------------------------------------------------------------------
    # 4. Generate self-contained .tp2
    # ------------------------------------------------------------------
    # We use STRING_SET to patch dialog.tlk — the proper WeiDU way.
    # dialog.tlk is NOT a chitin.key resource, so COPY_EXISTING won't
    # work on it. STRING_SET operates on the TLK directly.
    #
    # Syntax: STRING_SET <strref> ~text~ [SOUND_RESREF]
    #
    # We use ACTION_GET_STRREF to read the current text first, then
    # STRING_SET_EVALUATE to write it back with the sound resref added.
    # This preserves any text changes from other mods.

    tp2_lines = []
    tp2_lines.append(f"BACKUP ~PST_Voice_Mod/backup~")
    tp2_lines.append(f"AUTHOR ~PST Voice Mod~")
    tp2_lines.append(f"VERSION ~1.0~")
    tp2_lines.append(f"")
    tp2_lines.append(f"BEGIN ~PST Voice Mod - {len(entries)} voiced lines~")
    tp2_lines.append(f"")
    tp2_lines.append(f"  // Copy audio files to override")
    tp2_lines.append(f"  COPY ~PST_Voice_Mod/audio~ ~override~")
    tp2_lines.append(f"")
    tp2_lines.append(f"  // Patch dialog.tlk: add sound resrefs for {len(entries)} strings")

    for e in entries:
        safe = e["preview"].replace("\r", "").replace("\n", " ").replace("*/", "* /").replace("~", "")
        resref = e["resref"].upper()
        tp2_lines.append(f"")
        tp2_lines.append(f"  // strref {e['strref']}: {e['char']} - {safe}")
        # Read current text, then re-set it with the sound resref attached
        tp2_lines.append(f"  ACTION_GET_STRREF {e['strref']} str_{e['strref']}")
        tp2_lines.append(f"  STRING_SET_EVALUATE {e['strref']} ~%str_{e['strref']}%~ [{resref}]")

    tp2_lines.append(f"")

    with open(TP2_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(tp2_lines))

    print(f"Generated TP2: {TP2_PATH}")

    # ------------------------------------------------------------------
    # 5. Copy credits if available
    # ------------------------------------------------------------------
    if CREDITS_SRC.exists():
        shutil.copy2(CREDITS_SRC, MOD_DIR / "CREDITS.txt")
        print(f"Copied CREDITS.txt into mod package")

    # ------------------------------------------------------------------
    # Done
    # ------------------------------------------------------------------
    print(f"\n{'='*50}")
    print(f"  Mod built: {MOD_DIR}")
    print(f"  {len(entries)} dialog lines | {copied} audio files")
    print(f"{'='*50}")
    print()
    print("To install:")
    print("  1. Download WeiDU from https://github.com/WeiDUorg/weidu/releases")
    print(f"  2. Copy weidu.exe into: {GAME_DIR}")
    print(f"  3. Rename it to: setup-PST_Voice_Mod.exe")
    print(f"  4. Double-click setup-PST_Voice_Mod.exe")
    print()
    print("The tp2 is named setup-PST_Voice_Mod.tp2 so WeiDU finds it")
    print("automatically when you run setup-PST_Voice_Mod.exe.")


if __name__ == "__main__":
    main()
