from gui import CompressorApp
import tkinter as tk

def main():
    root = tk.Tk()
    # Initially hide the main window
    root.withdraw() 
    app = CompressorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
