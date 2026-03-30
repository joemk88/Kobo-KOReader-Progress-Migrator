# Kobo → KOReader Progress Migrator

A Windows-based tool to migrate reading progress from Kobo (Nickel) into KOReader, including accurate EPUB positioning via a sandboxed KOReader engine.

---

<img width="437" height="565" alt="Screenshot 2026-03-30 112807" src="https://github.com/user-attachments/assets/365839da-9485-47c7-be3b-cf28772ffd7c" />

---

## What This Does

This tool performs a one-time migration of your reading data:

- Copies reading progress (percent + position)
- Marks books as finished / reading
- Rebuilds KOReader metadata (`.sdr` sidecars)
- Updates KOReader history
- Uses KOReader itself to generate accurate EPUB positions

---

## Important: What This Tool Is (and Is NOT)

**This tool is:**

- A migration tool
- Designed to copy your library state from Nickel to KOReader
- Best used as a one-off or occasional sync

**This tool is NOT:**

- A live sync service
- A background sync tool

---

## Requirements

- **Python 3.10 or newer**
  Install from [python.org](https://www.python.org/downloads/) — ensure "Add Python to PATH" is checked during installation

- **WSL (Windows Subsystem for Linux)** with a Linux distribution installed
  Install via PowerShell:
```
  wsl --install
```

- **KOReader AppImage** — `koreader-v2026.03-x86_64.AppImage` placed in the **same folder** as the script
  Download from the [KOReader releases page](https://github.com/koreader/koreader/releases)

---

## Best Use Case (Recommended)

### Kobo device to KOReader on the same Kobo device

This is the most reliable setup:

- No file transfer issues
- No permission problems
- Direct access to KOReader folders
- Highest accuracy and success rate

---

## Using With Android Devices

This tool can migrate from a Kobo to another device (e.g. Android), but there are some additional considerations.

### Requirements

- You must first copy your book library to the target device
- File structure should remain consistent

### Connection Limitations

| Method | Reliability |
|---|---|
| USB (MTP) | Unreliable (not a real filesystem) |
| Sync tools (CrossDevice, etc.) | Can cause file locks |
| Network mounts (WebDAV, etc.) | Better |
| Manual copy | Most reliable |

### Manual Transfer Mode (Recommended for Android)

If your device connection is unreliable, enable:

> **Manual Transfer Mode (Android / unstable connection)**

**What it does:**

- Writes all output locally into an `OUTPUT` folder
- Preserves your exact folder structure
- Generates all `.sdr` metadata correctly

**Then you:**

1. Copy everything from `OUTPUT/BOOKS/` into your device Books folder
2. Copy `OUTPUT/KOREADER/history.lua` into `<device>/koreader/history.lua`
3. Overwrite when prompted
4. Restart KOReader


---

## How It Works

Kobo and KOReader store progress in completely different formats.

Instead of guessing positions, this tool runs KOReader in a sandbox and lets KOReader calculate the correct resume position itself. This is why EPUB support is significantly more accurate than typical conversion tools.

### Format Handling

| Format | Method |
|---|---|
| EPUB | KOReader sandbox (accurate positioning) |
| PDF | Percent to page |
| CBZ / CBR | Percent to image / page |
| Finished books | Instantly marked complete |

---

## Limitations

- EPUB positioning is approximate, but usually very close
- EPUB accuracy relies on the KOReader sandbox running correctly
- Book matching is filename-based — consistent naming between your Kobo and KOReader libraries improves results

---

## Development Note

This tool was developed with the assistance of AI.

AI was used to help:

- Reverse-engineer KOReader metadata formats
- Design the sandbox execution system
- Refine EPUB position handling
- Iterate through edge cases
