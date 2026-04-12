import ctypes
import os
import socket
import threading
import traceback
import webbrowser

import tkinter as tk
from tkinter import messagebox

from waitress import create_server

os.environ.setdefault("PCR_DESKTOP_MODE", "1")

from app import APP_DATA_DIR, app


LOG_FILE = APP_DATA_DIR / "desktop_startup.log"


def _log(message: str) -> None:
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    with LOG_FILE.open("a", encoding="utf-8") as fp:
        fp.write(message.rstrip() + "\n")


def _show_error(title: str, message: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)
    except Exception:
        pass


def _run_browser_fallback(url: str) -> None:
    webbrowser.open(url)

    root = tk.Tk()
    root.title("PCR BATO Fallback")
    root.geometry("520x170")
    root.resizable(False, False)

    label = tk.Label(
        root,
        text=(
            "Desktop runtime is unavailable on this device.\n"
            "The app has been opened in your default browser.\n"
            "Keep this window open while using the app."
        ),
        justify="left",
        padx=16,
        pady=16,
    )
    label.pack(fill="both", expand=True)

    def stop_server() -> None:
        root.destroy()

    tk.Button(root, text="Close App", command=stop_server, width=16).pack(pady=(0, 16))
    root.mainloop()


class ServerThread(threading.Thread):
    def __init__(self, host: str, port: int) -> None:
        super().__init__(daemon=True)
        self.server = create_server(app, host=host, port=port, threads=8)

    def run(self) -> None:
        self.server.run()

    def shutdown(self) -> None:
        self.server.close()


def _find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        sock.listen(1)
        return sock.getsockname()[1]


def main() -> None:
    host = "127.0.0.1"
    port = _find_free_port()
    url = f"http://{host}:{port}/records"

    server_thread = ServerThread(host, port)
    server_thread.start()

    try:
        try:
            import webview

            webview.create_window(
                "PCR BATO",
                url,
                width=1280,
                height=820,
                min_size=(980, 680),
            )
            webview.start()
        except Exception:
            _log("Desktop runtime failed. Falling back to browser mode.")
            _log(traceback.format_exc())
            _show_error(
                "PCR BATO Desktop",
                "Desktop runtime failed on this device. Falling back to browser mode.\n\n"
                f"Details are saved to: {LOG_FILE}",
            )
            _run_browser_fallback(url)
    finally:
        server_thread.shutdown()


if __name__ == "__main__":
    main()
