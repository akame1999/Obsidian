"""
OpsCenter Windows Patching — Desktop Launcher
Runs Streamlit in the main thread (required for signal handlers)
and tkinter control window in a background thread.
"""

import sys
import os
import time
import socket
import threading
import tkinter as tk
from tkinter import messagebox
import logging
import datetime

# Configure logging
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    filename=os.path.join(LOG_DIR, f'launcher_{datetime.datetime.now():%Y%m%d}.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ── Single-instance lock ───────────────────────────────────────────────────────
LOCK_PORT = 47823

def acquire_instance_lock():
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
        sock.bind(("127.0.0.1", LOCK_PORT))
        sock.listen(1)
        logging.info("Instance lock acquired")
        return sock
    except OSError as e:
        logging.warning(f"Could not acquire instance lock: {e}")
        return None

# ── Paths ──────────────────────────────────────────────────────────────────────
FROZEN = getattr(sys, "frozen", False)

if FROZEN:
    BASE_DIR = sys._MEIPASS
    EXE_DIR  = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR  = BASE_DIR

APP_SCRIPT = os.path.join(BASE_DIR, "patcher_app.py")
LOG_PATH   = os.path.join(EXE_DIR, "opcenter_start.log")
PORT       = 8501

def draw_shield(canvas, cx, cy, size):
    s = size / 2
    points = [
        cx - s, cy - s*0.9, cx - s, cy - s*0.1,
        cx,     cy + s,     cx + s, cy - s*0.1,
        cx + s, cy - s*0.9, cx,     cy - s*1.05,
    ]
    canvas.create_polygon(points, fill="#4f8bf9", outline="#2a5fcc", width=2, smooth=True)
    inner = [cx-s*0.6, cy-s*0.7, cx-s*0.6, cy, cx, cy+s*0.7, cx, cy-s*0.8]
    canvas.create_polygon(inner, fill="#3a72e8", outline="", smooth=True)
    canvas.create_line(
        cx-s*0.28, cy+s*0.05, cx-s*0.05, cy+s*0.32, cx+s*0.35, cy-s*0.22,
        width=max(2, int(size*0.07)), fill="white", capstyle="round", joinstyle="round"
    )

def run_tkinter_window(lock_sock):
    """
    Runs the tkinter control window in a background thread.
    Shows status, URL, and "Open Dashboard" button once server is ready.
    """
    root = tk.Tk()
    root.title("OpsCenter | Windows Patching")
    root.geometry("480x420")
    root.resizable(False, False)
    root.configure(bg="#0d1117")

    # Center window
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"480x420+{(sw-480)//2}+{(sh-420)//2}")

    # Shield icon
    c = tk.Canvas(root, width=80, height=80, bg="#0d1117", highlightthickness=0)
    c.pack(pady=(28, 0))
    draw_shield(c, 40, 40, 60)

    tk.Label(root, text="OpsCenter | Windows Patching",
             font=("Segoe UI", 15, "bold"), bg="#0d1117", fg="#ffffff").pack(pady=(10, 2))

    status_var = tk.StringVar(value="Starting server…")
    status_lbl = tk.Label(root, textvariable=status_var,
                          font=("Segoe UI", 10), bg="#0d1117", fg="#888888")
    status_lbl.pack()

    dot_var = tk.StringVar(value="●  ○  ○")
    dot_lbl = tk.Label(root, textvariable=dot_var,
                       font=("Segoe UI", 13), bg="#0d1117", fg="#4f8bf9")
    dot_lbl.pack(pady=8)

    url = f"http://127.0.0.1:{PORT}"

    url_lbl = tk.Label(root, text=url, font=("Segoe UI", 9, "underline"),
                       bg="#0d1117", fg="#4f8bf9", cursor="hand2")
    url_lbl.bind("<Button-1>", lambda e: __import__("webbrowser").open(url))

    def open_browser():
        import webbrowser
        webbrowser.open(url)

    open_btn = tk.Button(
        root, text="Open Dashboard",
        font=("Segoe UI", 11, "bold"), bg="#4f8bf9", fg="#ffffff",
        activebackground="#3a72e8", activeforeground="#ffffff",
        relief="flat", padx=24, pady=7, cursor="hand2",
        command=open_browser
    )

    running_lbl = tk.Label(
        root, text="● Server running — keep this window open",
        font=("Segoe UI", 9), bg="#0d1117", fg="#00c853"
    )

    # Animation
    anim_id = None
    def animate(frame=0):
        nonlocal anim_id
        patterns = ["●  ○  ○", "○  ●  ○", "○  ○  ●", "○  ●  ○"]
        dot_var.set(patterns[frame % len(patterns)])
        anim_id = root.after(420, animate, frame + 1)
    animate()

    # Wait for server to be ready
    def check_server():
        deadline = time.time() + 90
        while time.time() < deadline:
            try:
                with socket.create_connection(("127.0.0.1", PORT), timeout=1):
                    # Server is ready
                    if anim_id:
                        root.after_cancel(anim_id)
                    dot_lbl.pack_forget()
                    status_var.set("Dashboard is ready")
                    status_lbl.config(fg="#00c853")
                    url_lbl.pack(pady=2)
                    open_btn.pack(pady=10)
                    running_lbl.pack(pady=2)
                    open_browser()
                    logging.info("Server ready, opening browser")
                    return
            except OSError:
                time.sleep(0.5)
        # Timeout
        if anim_id:
            root.after_cancel(anim_id)
        dot_lbl.pack_forget()
        status_var.set("Failed to start — check log file")
        status_lbl.config(fg="#ff1744")
        tk.Label(root, text=f"Log: {LOG_PATH}",
                 font=("Segoe UI", 8), bg="#0d1117", fg="#ff4444").pack(pady=4)
        logging.error("Server failed to start within timeout")

    threading.Thread(target=check_server, daemon=True).start()

    def on_close():
        if messagebox.askokcancel("Exit OpsCenter",
                                  "Close OpsCenter?\nThis will stop the patching server."):
            try:
                lock_sock.close()
            except Exception:
                pass
            root.destroy()
            os._exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

def run_streamlit():
    """
    Runs Streamlit in the main thread (required for signal handlers).
    This function blocks until the server is stopped.
    """
    import logging

    log = open(LOG_PATH, "w", encoding="utf-8")

    try:
        # ── Production mode env vars ──────────────────────────────────────────
        os.environ["STREAMLIT_DEVELOPMENT_MODE"]           = "false"
        os.environ["STREAMLIT_SERVER_HEADLESS"]            = "true"
        os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
        os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"]    = "false"
        os.environ["STREAMLIT_SERVER_RUN_ON_SAVE"]         = "false"
        os.environ["STREAMLIT_THEME_BASE"]                 = "dark"
        os.environ["STREAMLIT_THEME_PRIMARY_COLOR"]        = "#4f8bf9"
        os.environ["STREAMLIT_SERVER_MAX_UPLOAD_SIZE"]     = "200"  # MB

        # ── Patch static path for PyInstaller ─────────────────────────────────
        st_static = os.path.join(BASE_DIR, "streamlit", "static")
        log.write(f"Static dir: {st_static} exists={os.path.isdir(st_static)}\n")
        os.environ["STREAMLIT_STATIC_FOLDER"] = st_static

        try:
            import streamlit
            import streamlit.file_util as _fu
            _original_get = _fu.get_streamlit_file_path
            def _patched_get(*args):
                path = os.path.join(BASE_DIR, "streamlit", *args)
                if os.path.exists(path):
                    return path
                return _original_get(*args)
            _fu.get_streamlit_file_path = _patched_get

            import streamlit.web.server.server_util as _su
            if hasattr(_su, "get_static_dir"):
                _su.get_static_dir = lambda: st_static
            log.write("Streamlit path patched OK\n")
            logging.info("Streamlit path patched successfully")
        except Exception as e:
            log.write(f"Patch warning: {e}\n")
            logging.warning(f"Streamlit patch warning: {e}")

        # Logging
        handler = logging.StreamHandler(log)
        logging.getLogger("streamlit").addHandler(handler)

        # Set argv
        sys.argv = [
            "streamlit", "run", APP_SCRIPT,
            "--server.port",     str(PORT),
            "--server.headless", "true",
            "--browser.gatherUsageStats", "false",
            "--server.maxUploadSize", "200",
        ]

        os.chdir(BASE_DIR)

        from streamlit.web import cli as stcli
        logging.info("Starting Streamlit server")
        stcli.main(standalone_mode=False)

    except Exception as e:
        log.write(f"STREAMLIT ERROR: {e}\n")
        logging.error(f"Streamlit error: {e}")
        import traceback
        traceback.print_exc(file=log)
    finally:
        log.flush()

def main():
    lock_sock = acquire_instance_lock()
    if lock_sock is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "OpsCenter Already Running",
            "OpsCenter is already open.\n\nCheck your taskbar or Alt+Tab to find it."
        )
        root.destroy()
        logging.info("Instance already running, exiting")
        return

    # Start tkinter window in background thread
    threading.Thread(target=run_tkinter_window, args=(lock_sock,), daemon=True).start()

    # Run Streamlit in main thread (blocks until server stops)
    run_streamlit()

if __name__ == "__main__":
    main()