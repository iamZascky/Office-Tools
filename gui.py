import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import threading
import sys
from compressor import iterative_compress, get_file_size_mb, convert_pdf_to_word

class OfficeToolsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Office Tools")
        self.root.geometry("500x400")
        self.root.configure(bg="#f8f9fa")
        
        # Modern Styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TProgressbar", thickness=10, troughcolor='#e9ecef', background='#007bff')
        
        self.main_container = tk.Frame(self.root, bg="#f8f9fa")
        self.main_container.pack(fill="both", expand=True)

        self.show_home()

    def clear_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    def show_home(self):
        self.clear_container()
        self.root.geometry("500x400")
        
        # Header
        header = tk.Frame(self.main_container, bg="#212529", height=80)
        header.pack(fill="x")
        tk.Label(header, text="OFFICE TOOLS", font=("Segoe UI", 16, "bold"), 
                 bg="#212529", fg="#ffffff", pady=20).pack()

        # Menu Content
        menu_frame = tk.Frame(self.main_container, bg="#f8f9fa", pady=40)
        menu_frame.pack(expand=True)

        btn_style = {"font": ("Segoe UI", 11, "bold"), "width": 25, "pady": 15, "bd": 0, "cursor": "hand2"}
        
        # Compress Button
        comp_btn = tk.Button(menu_frame, text="📄 FILE COMPRESSOR", bg="#007bff", fg="white",
                            activebackground="#0056b3", activeforeground="white",
                            command=self.show_compressor, **btn_style)
        comp_btn.pack(pady=10)

        # PDF to Word Button
        conv_btn = tk.Button(menu_frame, text="📝 PDF TO WORD", bg="#28a745", fg="white",
                            activebackground="#1e7e34", activeforeground="white",
                            command=self.show_converter, **btn_style)
        conv_btn.pack(pady=10)

        # Footer
        tk.Label(self.main_container, text="v2.0 | High Performance Tools", 
                 font=("Segoe UI", 8), bg="#f8f9fa", fg="#adb5bd").pack(side="bottom", pady=10)

    def show_compressor(self):
        self.clear_container()
        self.setup_action_view("File Compressor", "#007bff", self.run_compress_workflow)

    def show_converter(self):
        self.clear_container()
        self.setup_action_view("PDF to Word Converter", "#28a745", self.run_convert_workflow)

    def setup_action_view(self, title, color, action_cmd):
        self.root.geometry("500x300")
        
        # Header
        header = tk.Frame(self.main_container, bg=color, height=60)
        header.pack(fill="x")
        tk.Label(header, text=title.upper(), font=("Segoe UI", 12, "bold"), 
                 bg=color, fg="white", pady=15).pack()

        # Back Button
        back_btn = tk.Button(header, text="←", font=("Segoe UI", 12, "bold"), bg=color, fg="white",
                            bd=0, cursor="hand2", command=self.show_home)
        back_btn.place(x=10, y=10)

        # Content
        content = tk.Frame(self.main_container, bg="#f8f9fa", padx=30, pady=30)
        content.pack(fill="both", expand=True)

        self.status_label = tk.Label(content, text="Ready to start...", font=("Segoe UI", 10), 
                                    bg="#f8f9fa", fg="#495057")
        self.status_label.pack(pady=(0, 20))

        self.progress = ttk.Progressbar(content, orient="horizontal", length=400, mode="indeterminate")
        self.progress.pack(pady=10)

        self.start_btn = tk.Button(content, text="START PROCESS", bg=color, fg="white",
                                  font=("Segoe UI", 10, "bold"), padx=20, pady=10, bd=0,
                                  cursor="hand2", command=action_cmd)
        self.start_btn.pack(pady=20)

    # --- Compression Workflow ---
    def run_compress_workflow(self):
        input_path = filedialog.askopenfilename(
            title="Select File to Compress",
            filetypes=[("Compressible files", "*.pdf *.docx *.xlsx")]
        )
        if not input_path: return

        target_kb = simpledialog.askfloat("Target Size", "Enter target size in KB:", initialvalue=500.0, minvalue=10.0)
        if target_kb is None: return

        ext = os.path.splitext(input_path)[1]
        output_path = filedialog.asksaveasfilename(
            title="Save Compressed File As",
            initialfile=f"{os.path.splitext(os.path.basename(input_path))[0]}_compressed{ext}",
            defaultextension=ext,
            filetypes=[("Compressed files", f"*{ext}")]
        )
        if not output_path: return

        self.start_btn.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="Compressing... Please wait.", fg="#007bff")
        
        threading.Thread(target=self.execute_compression, 
                         args=(input_path, output_path, target_kb/1024.0, ext[1:].lower()), daemon=True).start()

    def execute_compression(self, input_path, output_path, target_mb, file_type):
        try:
            res_msg = iterative_compress(input_path, output_path, target_mb, file_type)
            final_size = get_file_size_mb(output_path) * 1024
            orig_size = get_file_size_mb(input_path) * 1024
            msg = f"Success!\n\nOriginal: {orig_size:.1f} KB\nCompressed: {final_size:.1f} KB\n({res_msg})"
            self.root.after(0, lambda: self.on_action_complete(True, msg))
        except Exception as e:
            self.root.after(0, lambda: self.on_action_complete(False, str(e)))

    # --- Conversion Workflow ---
    def run_convert_workflow(self):
        input_path = filedialog.askopenfilename(
            title="Select PDF to Convert",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not input_path: return

        output_path = filedialog.asksaveasfilename(
            title="Save Word File As",
            initialfile=f"{os.path.splitext(os.path.basename(input_path))[0]}.docx",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")]
        )
        if not output_path: return

        self.start_btn.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="Converting PDF to Word... This may take a moment.", fg="#28a745")
        
        threading.Thread(target=self.execute_conversion, args=(input_path, output_path), daemon=True).start()

    def execute_conversion(self, input_path, output_path):
        try:
            convert_pdf_to_word(input_path, output_path)
            self.root.after(0, lambda: self.on_action_complete(True, "PDF successfully converted to Word!"))
        except Exception as e:
            self.root.after(0, lambda: self.on_action_complete(False, str(e)))

    def on_action_complete(self, success, message):
        self.progress.stop()
        self.start_btn.config(state="normal")
        self.status_label.config(text="Ready to start...", fg="#495057")
        if success:
            messagebox.showinfo("Done", message)
        else:
            messagebox.showerror("Error", message)

if __name__ == "__main__":
    root = tk.Tk()
    app = OfficeToolsApp(root)
    # Center window
    root.update_idletasks()
    width, height = root.winfo_width(), root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    root.mainloop()
