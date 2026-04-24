import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import os
import sys

try:
    from docx2pdf import convert
except ImportError:
    messagebox.showerror(
        "Fehlende Abhängigkeit",
        "Bitte installiere docx2pdf:\n\npip install docx2pdf",
    )
    sys.exit(1)


def pdf_path(docx_path):
    return os.path.splitext(docx_path)[0] + ".pdf"


class DuplicateDialog(tk.Toplevel):
    """Modal dialog shown when existing PDFs are detected."""

    REPLACE_ALL = "replace_all"
    SKIP_EXISTING = "skip_existing"
    CANCEL = "cancel"

    def __init__(self, parent, existing_names):
        super().__init__(parent)
        self.title("Vorhandene PDFs gefunden")
        self.resizable(False, False)
        self.grab_set()
        self.result = self.CANCEL

        msg = (
            f"{len(existing_names)} PDF(s) mit gleichem Namen existieren bereits:\n\n"
            + "\n".join(f"  • {n}" for n in existing_names[:10])
            + ("\n  ..." if len(existing_names) > 10 else "")
            + "\n\nWas soll passieren?"
        )

        tk.Label(self, text=msg, font=("Segoe UI", 10), justify="left", wraplength=380).pack(
            padx=20, pady=(18, 14)
        )

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(0, 16))

        tk.Button(
            btn_frame,
            text="Alle ersetzen",
            width=16,
            font=("Segoe UI", 10, "bold"),
            bg="#2563EB",
            fg="white",
            relief="flat",
            cursor="hand2",
            command=self._replace_all,
        ).pack(side="left", padx=6)

        tk.Button(
            btn_frame,
            text="Nur neue konvertieren",
            width=20,
            font=("Segoe UI", 10),
            bg="#E5E7EB",
            relief="flat",
            cursor="hand2",
            command=self._skip_existing,
        ).pack(side="left", padx=6)

        tk.Button(
            btn_frame,
            text="Abbrechen",
            width=12,
            font=("Segoe UI", 10),
            bg="#E5E7EB",
            relief="flat",
            cursor="hand2",
            command=self._cancel,
        ).pack(side="left", padx=6)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self._center(parent)
        self.wait_window()

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _replace_all(self):
        self.result = self.REPLACE_ALL
        self.destroy()

    def _skip_existing(self):
        self.result = self.SKIP_EXISTING
        self.destroy()

    def _cancel(self):
        self.result = self.CANCEL
        self.destroy()


class DrDocxApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dr. Docx")
        self.resizable(False, False)
        self._build_ui()
        self._center()

    def _build_ui(self):
        pad = {"padx": 16, "pady": 8}

        tk.Label(self, text="Dr. Docx", font=("Segoe UI", 22, "bold"), fg="#2563EB").grid(
            row=0, column=0, columnspan=3, pady=(20, 4)
        )
        tk.Label(self, text="DOCX → PDF Konverter", font=("Segoe UI", 10), fg="#6B7280").grid(
            row=1, column=0, columnspan=3, pady=(0, 16)
        )

        tk.Label(self, text="Verzeichnis:", font=("Segoe UI", 10)).grid(
            row=2, column=0, sticky="w", **pad
        )

        self.folder_var = tk.StringVar()
        tk.Entry(self, textvariable=self.folder_var, width=46, font=("Segoe UI", 10)).grid(
            row=2, column=1, padx=(0, 4), pady=8
        )

        tk.Button(
            self,
            text="Durchsuchen",
            command=self._browse,
            font=("Segoe UI", 10),
            bg="#E5E7EB",
            relief="flat",
            cursor="hand2",
        ).grid(row=2, column=2, padx=(0, 16), pady=8)

        self.status_label = tk.Label(
            self, text="", font=("Segoe UI", 9), fg="#374151", wraplength=460
        )
        self.status_label.grid(row=3, column=0, columnspan=3, pady=(0, 4))

        self.progress = ttk.Progressbar(self, length=460, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, padx=16, pady=(0, 12))

        self.convert_btn = tk.Button(
            self,
            text="Konvertieren",
            command=self._start_conversion,
            font=("Segoe UI", 11, "bold"),
            bg="#2563EB",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=24,
            pady=8,
        )
        self.convert_btn.grid(row=5, column=0, columnspan=3, pady=(0, 20))

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    def _browse(self):
        folder = filedialog.askdirectory(title="Verzeichnis wählen")
        if folder:
            self.folder_var.set(folder)
            self._update_status(folder)

    def _update_status(self, folder):
        files = [f for f in os.listdir(folder) if f.lower().endswith(".docx")]
        count = len(files)
        if count == 0:
            self.status_label.config(text="Keine .docx-Dateien gefunden.", fg="#EF4444")
        else:
            self.status_label.config(text=f"{count} .docx-Datei(en) gefunden.", fg="#059669")

    def _start_conversion(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Kein Verzeichnis", "Bitte wähle zuerst ein Verzeichnis.")
            return

        all_files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".docx")
        ]
        if not all_files:
            messagebox.showinfo("Keine Dateien", "Keine .docx-Dateien im gewählten Verzeichnis.")
            return

        existing = [p for p in all_files if os.path.exists(pdf_path(p))]

        files_to_convert = all_files
        if existing:
            dlg = DuplicateDialog(self, [os.path.basename(pdf_path(p)) for p in existing])
            if dlg.result == DuplicateDialog.CANCEL:
                return
            if dlg.result == DuplicateDialog.SKIP_EXISTING:
                files_to_convert = [p for p in all_files if not os.path.exists(pdf_path(p))]
                if not files_to_convert:
                    messagebox.showinfo(
                        "Nichts zu tun",
                        "Alle PDFs existieren bereits. Keine neuen Dateien zu konvertieren.",
                    )
                    return

        self.convert_btn.config(state="disabled")
        self.progress["maximum"] = len(files_to_convert)
        self.progress["value"] = 0

        threading.Thread(target=self._convert, args=(files_to_convert,), daemon=True).start()

    def _convert(self, files):
        errors = []
        for i, path in enumerate(files, 1):
            name = os.path.basename(path)
            self.status_label.config(
                text=f"Konvertiere ({i}/{len(files)}): {name}", fg="#374151"
            )
            try:
                convert(path, pdf_path(path))
            except Exception as e:
                errors.append(f"{name}: {e}")
            self.progress["value"] = i
            self.update_idletasks()

        self.convert_btn.config(state="normal")
        if errors:
            self.status_label.config(text=f"Fertig mit {len(errors)} Fehler(n).", fg="#EF4444")
            messagebox.showerror(
                "Fehler",
                "Folgende Dateien konnten nicht konvertiert werden:\n\n" + "\n".join(errors),
            )
        else:
            self.status_label.config(
                text=f"Alle {len(files)} Datei(en) erfolgreich konvertiert!", fg="#059669"
            )
            messagebox.showinfo(
                "Fertig", f"{len(files)} Datei(en) wurden als PDF gespeichert."
            )


if __name__ == "__main__":
    app = DrDocxApp()
    app.mainloop()
