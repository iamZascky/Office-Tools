import sys
from gui import CompressorApp
import tkinter as tk

def main():
    root = tk.Tk()
    app = CompressorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
