import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import subprocess
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


def find_docx(root):
    """Recursively find all .docx files under root."""
    found = []
    for dirpath, _, filenames in os.walk(root):
        for f in filenames:
            if f.lower().endswith(".docx"):
                found.append(os.path.join(dirpath, f))
    return found


def open_in_explorer(path):
    subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])


class DuplicateDialog(tk.Toplevel):
    REPLACE_ALL = "replace_all"
    SKIP_EXISTING = "skip_existing"
    CANCEL = "cancel"

    def __init__(self, parent, existing_paths, root_folder):
        super().__init__(parent)
        self.title("Vorhandene PDFs gefunden")
        self.resizable(False, False)
        self.grab_set()
        self.result = self.CANCEL

        names = [os.path.relpath(pdf_path(p), root_folder) for p in existing_paths]
        preview = "\n".join(f"  • {n}" for n in names[:10])
        if len(names) > 10:
            preview += f"\n  ... und {len(names) - 10} weitere"

        msg = (
            f"{len(names)} PDF(s) mit gleichem Namen existieren bereits:\n\n"
            + preview
            + "\n\nWas soll passieren?"
        )
        tk.Label(self, text=msg, font=("Segoe UI", 10), justify="left", wraplength=420).pack(
            padx=20, pady=(18, 14)
        )

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(0, 16))

        tk.Button(
            btn_frame, text="Alle ersetzen", width=16,
            font=("Segoe UI", 10, "bold"), bg="#2563EB", fg="white",
            relief="flat", cursor="hand2", command=self._replace_all,
        ).pack(side="left", padx=6)

        tk.Button(
            btn_frame, text="Nur neue konvertieren", width=20,
            font=("Segoe UI", 10), bg="#E5E7EB",
            relief="flat", cursor="hand2", command=self._skip_existing,
        ).pack(side="left", padx=6)

        tk.Button(
            btn_frame, text="Abbrechen", width=12,
            font=("Segoe UI", 10), bg="#E5E7EB",
            relief="flat", cursor="hand2", command=self._cancel,
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


class ErrorDialog(tk.Toplevel):
    """Shows failed files as clickable links that open the file in Explorer."""

    def __init__(self, parent, errors):
        super().__init__(parent)
        self.title("Fehler beim Konvertieren")
        self.resizable(True, True)
        self.grab_set()

        tk.Label(
            self,
            text=f"{len(errors)} Datei(en) konnten nicht konvertiert werden.\n"
                 "Klicke auf einen Dateinamen um ihn im Explorer zu öffnen:",
            font=("Segoe UI", 10),
            justify="left",
        ).pack(padx=16, pady=(14, 6), anchor="w")

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")

        text = tk.Text(
            frame,
            font=("Segoe UI", 10),
            wrap="none",
            cursor="arrow",
            yscrollcommand=scrollbar.set,
            width=70,
            height=min(len(errors) + 2, 16),
            relief="flat",
            bg=self.cget("bg"),
        )
        text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text.yview)

        text.tag_config("link", foreground="#2563EB", underline=True)
        text.tag_config("err", foreground="#6B7280")

        for path, exc in errors:
            tag = f"link_{id(path)}"
            text.tag_config(tag, foreground="#2563EB", underline=True)
            text.tag_bind(tag, "<Button-1>", lambda _e, p=path: open_in_explorer(p))
            text.tag_bind(tag, "<Enter>", lambda _e: text.config(cursor="hand2"))
            text.tag_bind(tag, "<Leave>", lambda _e: text.config(cursor="arrow"))

            text.insert("end", os.path.basename(path), (tag,))
            text.insert("end", f"  –  {exc}\n", ("err",))

        text.config(state="disabled")

        tk.Button(
            self, text="Schließen", command=self.destroy,
            font=("Segoe UI", 10), bg="#E5E7EB", relief="flat", cursor="hand2",
        ).pack(pady=(0, 14))

        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")


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
        tk.Label(self, text="DOCX → PDF Konverter  ·  inkl. Unterordner", font=("Segoe UI", 10), fg="#6B7280").grid(
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
            self, text="Durchsuchen", command=self._browse,
            font=("Segoe UI", 10), bg="#E5E7EB", relief="flat", cursor="hand2",
        ).grid(row=2, column=2, padx=(0, 16), pady=8)

        self.status_label = tk.Label(
            self, text="", font=("Segoe UI", 9), fg="#374151", wraplength=480
        )
        self.status_label.grid(row=3, column=0, columnspan=3, pady=(0, 4))

        self.progress = ttk.Progressbar(self, length=480, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, padx=16, pady=(0, 12))

        self.convert_btn = tk.Button(
            self, text="Konvertieren", command=self._start_conversion,
            font=("Segoe UI", 11, "bold"), bg="#2563EB", fg="white",
            relief="flat", cursor="hand2", padx=24, pady=8,
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
            self._refresh_status(folder)

    def _refresh_status(self, folder):
        files = find_docx(folder)
        count = len(files)
        if count == 0:
            self.status_label.config(text="Keine .docx-Dateien gefunden (inkl. Unterordner).", fg="#EF4444")
        else:
            self.status_label.config(text=f"{count} .docx-Datei(en) gefunden (inkl. Unterordner).", fg="#059669")

    def _start_conversion(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Kein Verzeichnis", "Bitte wähle zuerst ein Verzeichnis.")
            return

        all_files = find_docx(folder)
        if not all_files:
            messagebox.showinfo("Keine Dateien", "Keine .docx-Dateien gefunden (auch nicht in Unterordnern).")
            return

        existing = [p for p in all_files if os.path.exists(pdf_path(p))]
        files_to_convert = all_files

        if existing:
            dlg = DuplicateDialog(self, existing, folder)
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

        threading.Thread(
            target=self._convert, args=(files_to_convert, folder), daemon=True
        ).start()

    def _convert(self, files, root_folder):
        errors = []
        for i, path in enumerate(files, 1):
            rel = os.path.relpath(path, root_folder)
            self.status_label.config(
                text=f"({i}/{len(files)})  {rel}", fg="#374151"
            )
            try:
                convert(path, pdf_path(path))
            except Exception as e:
                errors.append((path, str(e)))
            self.progress["value"] = i
            self.update_idletasks()

        self.convert_btn.config(state="normal")
        done = len(files) - len(errors)

        if errors:
            self.status_label.config(
                text=f"{done} konvertiert, {len(errors)} Fehler.", fg="#EF4444"
            )
            self.after(0, lambda: ErrorDialog(self, errors))
        else:
            self.status_label.config(
                text=f"Alle {len(files)} Datei(en) erfolgreich konvertiert!", fg="#059669"
            )
            messagebox.showinfo("Fertig", f"{len(files)} Datei(en) wurden als PDF gespeichert.")


if __name__ == "__main__":
    app = DrDocxApp()
    app.mainloop()
