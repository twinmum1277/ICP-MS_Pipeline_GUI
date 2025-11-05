"""
icpms_gui.py

Simple Tkinter front-end so you don't have to type long file paths.
Select:
  - SORT file
  - DIGEST file
  - ICV file
Then click "Run Batch" → calls process_batch(...) and writes Excel next to SORT.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import os
import sys
from io import StringIO

from process_icpms_batch import process_batch  # make sure this file is in same folder


class ICPMS_GUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ICP-MS Batch Processor")
        self.geometry("700x650")  # Sized to fit all controls comfortably
        self.minsize(700, 650)  # Prevent resizing smaller

        self.sort_path = ""
        self.digest_path = ""
        self.icv_path = ""
        self.ref_path = ""
        self.working_dir = os.path.expanduser("~")  # Default to home directory

        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 10, "pady": 4}
        
        # Working Folder section
        ttk.Label(self, text="Working Folder:", font=("TkDefaultFont", 11, "bold")).grid(row=0, column=0, sticky="w", **pad)
        self.working_dir_var = tk.StringVar(value="No folder selected (using Home)")
        ttk.Label(self, textvariable=self.working_dir_var, foreground="blue", font=("TkDefaultFont", 10)).grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(self, text="Set Folder…", command=self._set_working_folder).grid(row=0, column=2, **pad)

        # Separator
        ttk.Separator(self, orient='horizontal').grid(row=1, column=0, columnspan=3, sticky='ew', padx=10, pady=8)

        # SORT
        ttk.Label(self, text="SORT file (MassHunter export):").grid(row=2, column=0, sticky="w", **pad)
        self.sort_entry = ttk.Entry(self, width=50)
        self.sort_entry.grid(row=2, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_sort).grid(row=2, column=2, **pad)

        # DIGEST
        ttk.Label(self, text="DIGEST file (dilution factors):").grid(row=3, column=0, sticky="w", **pad)
        self.digest_entry = ttk.Entry(self, width=50)
        self.digest_entry.grid(row=3, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_digest).grid(row=3, column=2, **pad)

        # ICV
        ttk.Label(self, text="ICV file (ICV/SRM targets):").grid(row=4, column=0, sticky="w", **pad)
        self.icv_entry = ttk.Entry(self, width=50)
        self.icv_entry.grid(row=4, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_icv).grid(row=4, column=2, **pad)

        # REF
        ttk.Label(self, text="REF values file:").grid(row=5, column=0, sticky="w", **pad)
        self.ref_entry = ttk.Entry(self, width=50)
        self.ref_entry.grid(row=5, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_ref).grid(row=5, column=2, **pad)

        # Settings section
        ttk.Label(self, text="Settings:", font=("TkDefaultFont", 11, "bold")).grid(row=6, column=0, sticky="w", **pad)
        
        # Output units selection
        ttk.Label(self, text="Output units:").grid(row=6, column=1, sticky="w", **pad)
        self.output_units_var = tk.StringVar(value="ppb")  # Default to ppb (no division)
        units_frame = ttk.Frame(self)
        units_frame.grid(row=6, column=1, sticky="w", padx=(90, 0), pady=4)
        ttk.Radiobutton(units_frame, text="ppb (no conversion)", variable=self.output_units_var, value="ppb").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(units_frame, text="ppm (÷1000)", variable=self.output_units_var, value="ppm").pack(side=tk.LEFT, padx=5)

        # Run
        self.run_btn = ttk.Button(self, text="Run Batch", command=self._run_batch, width=20)
        self.run_btn.grid(row=8, column=1, **pad)

        # === Processing Summary Panel (below Run button) ===
        summary_frame = ttk.LabelFrame(self, text="Processing Summary", padding=10)
        summary_frame.grid(row=9, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)
        
        # Configure grid to expand summary panel
        self.grid_rowconfigure(9, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        self.summary_text = scrolledtext.ScrolledText(summary_frame, height=12, wrap=tk.WORD, 
                                                      font=("Arial", 13), state=tk.DISABLED)
        self.summary_text.pack(fill=tk.BOTH, expand=True)
        
        # Button frame for summary panel
        summary_btn_frame = ttk.Frame(summary_frame)
        summary_btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(summary_btn_frame, text="Clear", command=self._clear_summary).pack(side=tk.LEFT, padx=5)
        
        # Status at very bottom
        self.status_var = tk.StringVar(value="Select files, then click Run.")
        ttk.Label(self, textvariable=self.status_var, foreground="gray", font=("TkDefaultFont", 10)).grid(row=10, column=0, columnspan=3, sticky="w", padx=10, pady=6)
    
    def _clear_summary(self):
        """Clear the processing summary"""
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.config(state=tk.DISABLED)
    
    def _append_to_summary(self, text, color=None, bold=False):
        """Append text to the processing summary"""
        self.summary_text.config(state=tk.NORMAL)
        if color or bold:
            # Create unique tag name
            tag_name = f"{color or 'normal'}{'_bold' if bold else ''}"
            font_config = ("Arial", 13, "bold") if bold else ("Arial", 13)
            
            if color == "red":
                self.summary_text.tag_config(tag_name, foreground=color, font=("Arial", 14, "bold"))
            elif color == "green":
                self.summary_text.tag_config(tag_name, foreground=color, font=("Arial", 15, "bold"))
            elif color:
                self.summary_text.tag_config(tag_name, foreground=color, font=font_config)
            else:
                self.summary_text.tag_config(tag_name, font=font_config)
            
            self.summary_text.insert(tk.END, text, tag_name)
        else:
            self.summary_text.insert(tk.END, text)
        self.summary_text.see(tk.END)  # Auto-scroll to bottom
        self.summary_text.config(state=tk.DISABLED)
        self.update_idletasks()

    def _autosize_to_content(self):
        # Expand the window to at least the size required by its widgets
        self.update_idletasks()
        req_w = self.winfo_reqwidth()
        req_h = self.winfo_reqheight()
        self.minsize(req_w, req_h)
        self.geometry(f"{req_w}x{req_h}")

    # ---- folder picker ----
    def _set_working_folder(self):
        folder = filedialog.askdirectory(
            parent=self,
            initialdir=self.working_dir,
            title="Select Working Folder (where your input files are located)"
        )
        if folder:
            self.working_dir = folder
            # Show abbreviated path if too long
            display_path = folder
            if len(display_path) > 60:
                display_path = "..." + display_path[-57:]
            self.working_dir_var.set(display_path)
            self.status_var.set(f"Working folder set: {os.path.basename(folder)}")

    # ---- file pickers ----
    def _pick_sort(self):
        # Bring window to front so dialog isn't hidden on macOS
        self.lift()
        self.status_var.set("Opening file dialog…")
        self.update_idletasks()
        self.attributes('-topmost', True)
        self.after(0, lambda: self.attributes('-topmost', False))

        try:
            path = filedialog.askopenfilename(
                parent=self,
                initialdir=os.path.dirname(self.sort_path) if self.sort_path else self.working_dir,
                title="Select SORT file",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
        except Exception as e:
            messagebox.showerror("File dialog error", str(e))
            path = simpledialog.askstring("Enter path", "Paste full path to SORT .csv:") or ""
        if path:
            self.sort_path = path
            self.sort_entry.delete(0, tk.END)
            self.sort_entry.insert(0, path)

    def _pick_digest(self):
        # Bring window to front so dialog isn't hidden on macOS
        self.lift()
        self.status_var.set("Opening file dialog…")
        self.update_idletasks()
        self.attributes('-topmost', True)
        self.after(0, lambda: self.attributes('-topmost', False))

        try:
            path = filedialog.askopenfilename(
                parent=self,
                initialdir=os.path.dirname(self.digest_path) if self.digest_path else self.working_dir,
                title="Select DIGEST file",
                filetypes=[("CSV/Excel files", "*.csv *.xlsx *.xls"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
        except Exception as e:
            messagebox.showerror("File dialog error", str(e))
            path = simpledialog.askstring("Enter path", "Paste full path to DIGEST file:") or ""
        if path:
            self.digest_path = path
            self.digest_entry.delete(0, tk.END)
            self.digest_entry.insert(0, path)

    def _pick_icv(self):
        # Bring window to front so dialog isn't hidden on macOS
        self.lift()
        self.status_var.set("Opening file dialog…")
        self.update_idletasks()
        self.attributes('-topmost', True)
        self.after(0, lambda: self.attributes('-topmost', False))

        try:
            path = filedialog.askopenfilename(
                parent=self,
                initialdir=os.path.dirname(self.icv_path) if self.icv_path else self.working_dir,
                title="Select ICV file",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
        except Exception as e:
            messagebox.showerror("File dialog error", str(e))
            path = simpledialog.askstring("Enter path", "Paste full path to ICV .csv:") or ""
        if path:
            self.icv_path = path
            self.icv_entry.delete(0, tk.END)
            self.icv_entry.insert(0, path)

    def _pick_ref(self):
        # Bring window to front so dialog isn't hidden on macOS
        self.lift()
        self.status_var.set("Opening file dialog…")
        self.update_idletasks()
        self.attributes('-topmost', True)
        self.after(0, lambda: self.attributes('-topmost', False))

        try:
            path = filedialog.askopenfilename(
                parent=self,
                initialdir=os.path.dirname(self.ref_path) if self.ref_path else self.working_dir,
                title="Select REF values file",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
        except Exception as e:
            messagebox.showerror("File dialog error", str(e))
            path = simpledialog.askstring("Enter path", "Paste full path to REF .csv (optional):") or ""
        if path:
            self.ref_path = path
            self.ref_entry.delete(0, tk.END)
            self.ref_entry.insert(0, path)

    # ---- run ----
    def _run_batch(self):
        if not self.sort_path or not self.digest_path or not self.icv_path:
            messagebox.showerror("Missing file", "Please select SORT, DIGEST, and ICV files.")
            return

        # build default output name next to SORT
        out_dir = os.path.dirname(self.sort_path)
        out_path = os.path.join(out_dir, "Batch_Results.xlsx")

        self.status_var.set("Running…")
        self.update_idletasks()

        try:
            # Capture debug output
            old_stdout = sys.stdout
            sys.stdout = captured_output = StringIO()
            
            # Run batch and get statistics
            stats = process_batch(
                sort_path=self.sort_path,
                digest_path=self.digest_path,
                icv_path=self.icv_path,
                output_path=out_path,
                ref_values_path=self.ref_path if self.ref_path else None,
                apply_divide1000=(self.output_units_var.get() == "ppm")
            )
            
            # Restore stdout
            sys.stdout = old_stdout
            debug_output = captured_output.getvalue()
            
            self.status_var.set(f"Done. Wrote: {out_path}")
            
            # Write clean summary to panel
            self._append_to_summary("="*55 + "\n", "blue")
            self._append_to_summary("✓ PROCESSING COMPLETE\n", "green")
            self._append_to_summary("="*55 + "\n\n", "blue")
            
            # WARNINGS FIRST - VERY PROMINENT
            unmatched_samples = stats.get('unmatched_samples', [])
            if unmatched_samples and len(unmatched_samples) > 0:
                self._append_to_summary("⚠️  " + "="*45 + "\n", "red")
                self._append_to_summary(f"⚠️  MISSING DIGEST DATA: {len(unmatched_samples)} SAMPLES\n", "red")
                self._append_to_summary("⚠️  " + "="*45 + "\n\n", "red")
                for s in unmatched_samples:
                    self._append_to_summary(f"  ❌ {s}\n", "red")
                self._append_to_summary("\nThese samples used df=1.0 (NO correction applied!)\n", "red")
                self._append_to_summary("→ Highlighted in YELLOW in output Excel.\n\n", "red")
                self._append_to_summary("="*45 + "\n\n", "red")
            
            # SUMMARY STATISTICS
            self._append_to_summary("📊 BATCH SUMMARY\n", "blue", bold=True)
            self._append_to_summary(f"  Samples processed:    {stats.get('total_samples', 0)}\n")
            self._append_to_summary(f"  ICV samples:          {stats.get('total_icv', 0)}\n")
            self._append_to_summary(f"  REF samples:          {stats.get('total_ref', 0)}\n")
            self._append_to_summary(f"  Blank samples:        {stats.get('total_blanks', 0)}\n")
            self._append_to_summary(f"  Elements analyzed:    {stats.get('elements_analyzed', 0)}\n\n")
            
            # QC PASS RATES
            self._append_to_summary("✓ QC PASS RATES\n", "blue", bold=True)
            icv_pass = stats.get('icv_pass_rate', 0)
            ref_pass = stats.get('ref_pass_rate', 0)
            
            icv_color = "green" if icv_pass >= 80 else "red"
            ref_color = "green" if ref_pass >= 80 else ("orange" if ref_pass >= 60 else "red")
            
            self._append_to_summary(f"  ICV (90-110%):        ", None)
            self._append_to_summary(f"{icv_pass:.0f}% passed\n", icv_color, bold=True)
            
            self._append_to_summary(f"  REF (80-120%):        ", None)
            self._append_to_summary(f"{ref_pass:.0f}% passed\n", ref_color, bold=True)
            
            self._append_to_summary(f"\n📄 Output: {os.path.basename(out_path)}\n", "blue")
            self._append_to_summary("="*55 + "\n", "blue")
            
            # Show simple popup
            if unmatched_samples and len(unmatched_samples) > 0:
                messagebox.showwarning("Complete with Warnings", 
                                      f"⚠️ Processing complete with warnings!\n\n{len(unmatched_samples)} samples missing DIGEST data.\n\nCheck summary panel for details.")
            else:
                messagebox.showinfo("Complete", f"✓ Processing complete!\n\nOutput: {os.path.basename(out_path)}")
            
        except Exception as e:
            sys.stdout = old_stdout  # Restore stdout on error
            self.status_var.set("Error.")
            self._append_to_summary(f"\n\n❌ ERROR: {str(e)}\n\n", "red")
            messagebox.showerror("Error during processing", str(e))


if __name__ == "__main__":
    app = ICPMS_GUI()
    app.mainloop()
