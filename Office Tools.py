import sys
from gui import OfficeToolsApp
import tkinter as tk

def main():
    root = tk.Tk()
    app = OfficeToolsApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
