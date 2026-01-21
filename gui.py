import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import threading
import sys
from compressor import iterative_compress, get_file_size_mb

class CompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Office Tools")
        self.root.geometry("400x150")
        self.root.configure(bg="#f8f9fa")
        
        # Center the window
        self.root.eval('tk::PlaceWindow . center')

        self.create_widgets()
        
        # Start the process immediately (Bypassing main menu)
        self.root.after(100, self.start_process)

    def create_widgets(self):
        # Mini Header
        header_frame = tk.Frame(self.root, bg="#212529", height=40)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="Compressing...", font=("Segoe UI", 10, "bold"), 
                 bg="#212529", fg="white").pack(expand=True)

        # Content
        main_content = tk.Frame(self.root, bg="#f8f9fa", padx=20, pady=20)
        main_content.pack(fill="both", expand=True)

        # Progress
        self.progress = ttk.Progressbar(main_content, orient="horizontal", mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 10))
        
        # Status
        self.status_label = tk.Label(main_content, text="Initializing...", 
                                    font=("Segoe UI", 9), bg="#f8f9fa", fg="#6c757d")
        self.status_label.pack()

    def start_process(self):
        # 1. Select File
        input_path = filedialog.askopenfilename(
            title="Step 1: Select File to Compress",
            filetypes=[("Compressible files", "*.pdf *.docx *.xlsx")]
        )
        if not input_path:
            self.root.destroy()
            return

        # 2. Ask for Target Size
        target_kb = simpledialog.askfloat(
            "Step 2: Target Size", 
            "Enter target size in KB:", 
            initialvalue=500.0, 
            minvalue=10.0
        )
        if target_kb is None:
            self.root.destroy()
            return

        # 3. Select Output Location
        base_name = os.path.basename(input_path)
        name, ext = os.path.splitext(base_name)
        suggested_name = f"{name}_compressed{ext}"
        
        output_path = filedialog.asksaveasfilename(
            title="Step 3: Save Compressed File As",
            initialdir=os.path.dirname(input_path),
            initialfile=suggested_name,
            defaultextension=ext,
            filetypes=[("Compressed files", f"*{ext}")]
        )
        
        if not output_path:
            self.root.destroy()
            return

        # Show the progress window
        self.root.deiconify() 
        self.progress.start()
        self.status_label.config(text=f"Compressing {base_name}...", fg="#007bff")
        
        target_mb = target_kb / 1024.0
        threading.Thread(target=self.run_compression, 
                         args=(input_path, output_path, target_mb, ext[1:].lower())).start()

    def run_compression(self, input_path, output_path, target_mb, file_type):
        try:
            res_msg = iterative_compress(input_path, output_path, target_mb, file_type)
            final_size_mb = get_file_size_mb(output_path)
            original_size_mb = get_file_size_mb(input_path)
            
            message = f"Success!\n\nOriginal: {original_size_mb*1024:.1f} KB\nCompressed: {final_size_mb*1024:.1f} KB\n({res_msg})"
            self.root.after(0, lambda: self.on_complete(True, message))
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda msg=err_msg: self.on_complete(False, f"Error: {msg}"))

    def on_complete(self, success, message):
        self.progress.stop()
        if success:
            messagebox.showinfo("Done", message)
        else:
            messagebox.showerror("Error", message)
        
        # Close the app after completion
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    # Initially hide the main window until the dialogs are done
    root.withdraw() 
    app = CompressorApp(root)
    root.mainloop()
