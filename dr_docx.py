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


class DrDocxApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dr. Docx")
        self.resizable(False, False)
        self._build_ui()
        self._center()

    def _build_ui(self):
        pad = {"padx": 16, "pady": 8}

        header = tk.Label(
            self,
            text="Dr. Docx",
            font=("Segoe UI", 22, "bold"),
            fg="#2563EB",
        )
        header.grid(row=0, column=0, columnspan=3, pady=(20, 4))

        sub = tk.Label(
            self,
            text="DOCX → PDF Konverter",
            font=("Segoe UI", 10),
            fg="#6B7280",
        )
        sub.grid(row=1, column=0, columnspan=3, pady=(0, 16))

        tk.Label(self, text="Verzeichnis:", font=("Segoe UI", 10)).grid(
            row=2, column=0, sticky="w", **pad
        )

        self.folder_var = tk.StringVar()
        self.folder_entry = tk.Entry(
            self, textvariable=self.folder_var, width=46, font=("Segoe UI", 10)
        )
        self.folder_entry.grid(row=2, column=1, padx=(0, 4), pady=8)

        browse_btn = tk.Button(
            self,
            text="Durchsuchen",
            command=self._browse,
            font=("Segoe UI", 10),
            bg="#E5E7EB",
            relief="flat",
            cursor="hand2",
        )
        browse_btn.grid(row=2, column=2, padx=(0, 16), pady=8)

        self.status_label = tk.Label(
            self, text="", font=("Segoe UI", 9), fg="#374151", wraplength=460
        )
        self.status_label.grid(row=3, column=0, columnspan=3, pady=(0, 4))

        self.progress = ttk.Progressbar(
            self, length=460, mode="determinate"
        )
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
            self.status_label.config(
                text=f"{count} .docx-Datei(en) gefunden.", fg="#059669"
            )

    def _start_conversion(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Kein Verzeichnis", "Bitte wähle zuerst ein Verzeichnis.")
            return

        files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".docx")
        ]
        if not files:
            messagebox.showinfo("Keine Dateien", "Keine .docx-Dateien im gewählten Verzeichnis.")
            return

        self.convert_btn.config(state="disabled")
        self.progress["maximum"] = len(files)
        self.progress["value"] = 0

        threading.Thread(
            target=self._convert, args=(files,), daemon=True
        ).start()

    def _convert(self, files):
        errors = []
        for i, path in enumerate(files, 1):
            name = os.path.basename(path)
            self.status_label.config(
                text=f"Konvertiere ({i}/{len(files)}): {name}", fg="#374151"
            )
            try:
                convert(path)
            except Exception as e:
                errors.append(f"{name}: {e}")
            self.progress["value"] = i
            self.update_idletasks()

        self.convert_btn.config(state="normal")
        if errors:
            self.status_label.config(
                text=f"Fertig mit {len(errors)} Fehler(n).", fg="#EF4444"
            )
            messagebox.showerror(
                "Fehler",
                "Folgende Dateien konnten nicht konvertiert werden:\n\n"
                + "\n".join(errors),
            )
        else:
            self.status_label.config(
                text=f"Alle {len(files)} Datei(en) erfolgreich konvertiert!", fg="#059669"
            )
            messagebox.showinfo("Fertig", f"{len(files)} Datei(en) wurden als PDF gespeichert.")


if __name__ == "__main__":
    app = DrDocxApp()
    app.mainloop()
