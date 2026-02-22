# Python Auto Capture Tools

A collection of simple Python productivity tools for capturing screenshots and keystrokes, designed to help with report-making and command review.

## Tools

### 1. Screen Capture (`main.py`)

Captures screenshots via global hotkey and saves them into a landscape Word document (.docx). Each session creates a new file.

- **Global hotkey** (`Ctrl+Shift+S`) to capture screenshots
- Full-quality PNG, no compression — ready for reports
- Auto-save on every capture
- Audio beep feedback

### 2. Keystroke Logger (`keystroke_logger.py`)

Records all keystrokes globally and saves them to a `.txt` log file. Useful for reviewing commands you've typed during a session to check for mistakes.

- Captures all keyboard input across all applications
- Detects `Ctrl+V` (paste) and records clipboard content
- Lines are flushed on every `Enter` press with timestamp
- Special keys displayed as tags (e.g. `[ESC]`, `[F1]`, `[UP]`)
- Session summary at the end of the log

## Requirements

- Python 3.8+
- Windows OS

## Installation

```bash
git clone https://github.com/<your-username>/python-auto-capture.git
cd python-auto-capture
pip install -r requirements.txt
```

## Usage

### Screen Capture

```bash
python main.py
```

| Shortcut | Action |
|---|---|
| `Ctrl+Shift+S` | Capture screenshot |
| `Ctrl+C` | Stop and save |

Output: `captures/capture_YYYY-MM-DD_HHMMSS.docx`

### Keystroke Logger

```bash
python keystroke_logger.py
```

| Action | How |
|---|---|
| Start logging | Run the script |
| Stop logging | `Ctrl+C` |

Output: `logs/keylog_YYYY-MM-DD_HHMMSS.txt`

**Example log content:**

```
Keystroke Log — Started at 2026-02-13 14:30:00
============================================================

[14:30:05]  cd projects
[14:30:08]  git status
[14:30:15]  docker run [PASTE: nginx:latest -p 8080:80 -d]
[14:31:02]  kubectl get pods

============================================================
Session ended at 2026-02-13 14:31:10
Total lines captured: 4
```

## License

MIT
