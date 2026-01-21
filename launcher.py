import subprocess
import os
import sys
import tkinter as tk
from tkinter import messagebox
import logging

# Get base directory
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Setup logging to an absolute path
log_file = os.path.join(base_dir, 'launcher_log.txt')
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def get_python_executable():
    """Tries to find the pythonw or python executable."""
    logging.debug("Starting python discovery")
    candidates = [
        os.path.join(os.path.dirname(sys.executable), "pythonw.exe"),
        os.path.join(os.path.dirname(sys.executable), "python.exe"),
        "pythonw.exe",
        "python.exe",
        r"C:\Users\Dell\AppData\Local\Programs\Python\Python312\pythonw.exe",
        r"C:\Users\Dell\AppData\Local\Programs\Python\Python312\python.exe",
    ]
    
    for candidate in candidates:
        logging.debug(f"Checking candidate: {candidate}")
        if os.path.exists(candidate):
            logging.info(f"Found python at: {candidate}")
            return candidate
            
    import shutil
    res = shutil.which("pythonw") or shutil.which("python")
    logging.info(f"Shutil which found: {res}")
    return res

def launch():
    logging.info("--- Launching Office Tools ---")
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
        logging.debug("Running as frozen EXE")
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logging.debug("Running as script")
    
    logging.info(f"Base Directory: {base_dir}")
    
    script_path = os.path.join(base_dir, "Office Tools.py")
    python_exe = get_python_executable()
    
    logging.info(f"Target script: {script_path}")
    logging.info(f"Python EXE: {python_exe}")

    if not os.path.exists(script_path):
        msg = f"Missing script file:\n{script_path}"
        logging.error(msg)
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", msg)
        return

    if not python_exe:
        msg = "Could not locate Python interpreter.\nPlease ensure Python is installed and in your PATH."
        logging.error(msg)
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", msg)
        return

    try:
        logging.info(f"Executing: {python_exe} {script_path}")
        # Use pythonw if possible to avoid terminal flash
        # If we found python.exe but not pythonw.exe, we might still want to use it
        subprocess.Popen([python_exe, script_path], cwd=base_dir, start_new_session=True)
        logging.info("Process started successfully")
    except Exception as e:
        msg = f"Failed to launch:\n{str(e)}\n\nCommand: {python_exe} {script_path}"
        logging.error(f"Execution failed: {str(e)}")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", msg)

if __name__ == "__main__":
    try:
        launch()
    except Exception as e:
        logging.critical(f"Unhandled exception: {str(e)}")
