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
from tkinter import ttk, filedialog, messagebox, simpledialog
import os

from process_icpms_batch import process_batch  # make sure this file is in same folder


class ICPMS_GUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ICP-MS Batch Processor")
        self.geometry("560x240")

        self.sort_path = ""
        self.digest_path = ""
        self.icv_path = ""
        self.crm_path = ""

        self._build_widgets()

        # Ensure initial window is large enough to show all controls
        self.after(0, self._autosize_to_content)

    def _build_widgets(self):
        pad = {"padx": 10, "pady": 4}

        # SORT
        ttk.Label(self, text="SORT file (MassHunter export):").grid(row=0, column=0, sticky="w", **pad)
        self.sort_entry = ttk.Entry(self, width=50)
        self.sort_entry.grid(row=0, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_sort).grid(row=0, column=2, **pad)

        # DIGEST
        ttk.Label(self, text="DIGEST file (dilution factors):").grid(row=1, column=0, sticky="w", **pad)
        self.digest_entry = ttk.Entry(self, width=50)
        self.digest_entry.grid(row=1, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_digest).grid(row=1, column=2, **pad)

        # ICV
        ttk.Label(self, text="ICV file (ICV/SRM targets):").grid(row=2, column=0, sticky="w", **pad)
        self.icv_entry = ttk.Entry(self, width=50)
        self.icv_entry.grid(row=2, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_icv).grid(row=2, column=2, **pad)

        # CRM (optional)
        ttk.Label(self, text="CRM values file (optional):").grid(row=3, column=0, sticky="w", **pad)
        self.crm_entry = ttk.Entry(self, width=50)
        self.crm_entry.grid(row=3, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self._pick_crm).grid(row=3, column=2, **pad)

        # Run
        self.run_btn = ttk.Button(self, text="Run Batch", command=self._run_batch)
        self.run_btn.grid(row=5, column=1, **pad)

        # Status
        self.status_var = tk.StringVar(value="Select files, then click Run.")
        ttk.Label(self, textvariable=self.status_var, foreground="gray").grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=6)

        # Let the entry column expand if user resizes
        self.grid_columnconfigure(1, weight=1)

    def _autosize_to_content(self):
        # Expand the window to at least the size required by its widgets
        self.update_idletasks()
        req_w = self.winfo_reqwidth()
        req_h = self.winfo_reqheight()
        self.minsize(req_w, req_h)
        self.geometry(f"{req_w}x{req_h}")

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
                initialdir=os.path.dirname(self.sort_path) if self.sort_path else os.path.expanduser("~"),
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
                initialdir=os.path.dirname(self.digest_path) if self.digest_path else os.path.expanduser("~"),
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
                initialdir=os.path.dirname(self.icv_path) if self.icv_path else os.path.expanduser("~"),
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

    def _pick_crm(self):
        # Bring window to front so dialog isn't hidden on macOS
        self.lift()
        self.status_var.set("Opening file dialog…")
        self.update_idletasks()
        self.attributes('-topmost', True)
        self.after(0, lambda: self.attributes('-topmost', False))

        try:
            path = filedialog.askopenfilename(
                parent=self,
                initialdir=os.path.dirname(self.crm_path) if self.crm_path else os.path.expanduser("~"),
                title="Select CRM values file",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
        except Exception as e:
            messagebox.showerror("File dialog error", str(e))
            path = simpledialog.askstring("Enter path", "Paste full path to CRM .csv (optional):") or ""
        if path:
            self.crm_path = path
            self.crm_entry.delete(0, tk.END)
            self.crm_entry.insert(0, path)

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
            process_batch(
                sort_path=self.sort_path,
                digest_path=self.digest_path,
                icv_path=self.icv_path,
                output_path=out_path,
                crm_values_path=self.crm_path if self.crm_path else None,
                apply_divide1000=True
            )
            self.status_var.set(f"Done. Wrote: {out_path}")
            messagebox.showinfo("Done", f"Processing complete.\nOutput saved to:\n{out_path}")
        except Exception as e:
            self.status_var.set("Error.")
            messagebox.showerror("Error during processing", str(e))


if __name__ == "__main__":
    app = ICPMS_GUI()
    app.mainloop()
