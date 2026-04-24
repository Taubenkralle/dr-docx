import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import subprocess
import shutil
import tempfile
import os
import sys
import math

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
    found = []
    for dirpath, _, filenames in os.walk(root):
        for f in filenames:
            if f.lower().endswith(".docx"):
                found.append(os.path.join(dirpath, f))
    return found


def open_in_explorer(path):
    subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])


def convert_via_temp(docx_path, out_pdf_path):
    """Copy to temp dir first so OneDrive/cloud-only files are downloaded."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_docx = os.path.join(tmp, os.path.basename(docx_path))
        tmp_pdf  = os.path.splitext(tmp_docx)[0] + ".pdf"
        shutil.copy2(docx_path, tmp_docx)   # forces OneDrive to sync file locally
        convert(tmp_docx, tmp_pdf)
        shutil.move(tmp_pdf, out_pdf_path)


class DoctorCanvas(tk.Canvas):
    W, H = 140, 168

    def __init__(self, parent, **kwargs):
        kwargs.setdefault("bg", parent.cget("bg"))
        super().__init__(parent, width=self.W, height=self.H,
                         highlightthickness=0, **kwargs)
        self._frame = 0
        self._animating = False
        self._job = None
        self._idle_job = self.after(0, self._idle_tick)

    # ------------------------------------------------------------------
    def start(self):
        if self._idle_job:
            self.after_cancel(self._idle_job)
            self._idle_job = None
        self._animating = True
        self._frame = 0
        self._tick()

    def stop(self):
        self._animating = False
        if self._job:
            self.after_cancel(self._job)
            self._job = None
        self._frame = 0
        self._idle_job = self.after(0, self._idle_tick)

    def _tick(self):
        if not self._animating:
            return
        self._frame += 1
        self._draw(self._frame, active=True)
        self._job = self.after(55, self._tick)

    def _idle_tick(self):
        self._frame += 1
        self._draw(self._frame, active=False)
        self._idle_job = self.after(80, self._idle_tick)

    # ------------------------------------------------------------------
    def _draw(self, t, active):
        self.delete("all")
        cx = self.W // 2  # 70

        speed = 0.15 if active else 0.04
        breath = math.sin(t * speed) * 2

        # ── HEAD ──────────────────────────────────────────────────────
        hy = 10 + breath
        self.create_oval(cx - 24, hy, cx + 24, hy + 46,
                         fill="#FDDCB5", outline="#C9966A", width=2)

        # Doctor's head mirror / reflector band
        self.create_rectangle(cx - 26, hy + 4, cx + 26, hy + 13,
                               fill="#D1D5DB", outline="#9CA3AF", width=1)
        self.create_oval(cx - 8, hy + 2, cx + 8, hy + 15,
                         fill="#C0C0C0", outline="#9CA3AF", width=2)
        self.create_oval(cx - 3, hy + 6, cx + 3, hy + 11,
                         fill="#444", outline="")

        # Eyes
        ey = hy + 24
        blink = (t % 55 < 2)
        if blink:
            self.create_line(cx - 13, ey, cx - 6, ey, fill="#555", width=2)
            self.create_line(cx + 6,  ey, cx + 13, ey, fill="#555", width=2)
        else:
            self.create_oval(cx - 14, ey - 4, cx - 6,  ey + 4, fill="#333", outline="")
            self.create_oval(cx + 6,  ey - 4, cx + 14, ey + 4, fill="#333", outline="")
            self.create_oval(cx - 12, ey - 3, cx - 10, ey - 1, fill="white", outline="")
            self.create_oval(cx + 8,  ey - 3, cx + 10, ey - 1, fill="white", outline="")

        # Smile – wider when busy
        ex = 160 if active else 140
        self.create_arc(cx - 12, hy + 31, cx + 12, hy + 44,
                        start=200, extent=ex, style="arc",
                        outline="#C9966A", width=2)

        # ── BODY ──────────────────────────────────────────────────────
        bt = hy + 44          # body top y
        self.create_rectangle(cx - 30, bt, cx + 30, bt + 68,
                               fill="white", outline="#9CA3AF", width=2)

        # Lapels
        self.create_polygon(cx, bt + 2, cx - 22, bt + 20, cx - 8, bt + 20,
                            fill="#F3F4F6", outline="#9CA3AF", width=1)
        self.create_polygon(cx, bt + 2, cx + 22, bt + 20, cx + 8, bt + 20,
                            fill="#F3F4F6", outline="#9CA3AF", width=1)

        # Red cross
        self.create_rectangle(cx - 6,  bt + 32, cx + 6,  bt + 46,
                               fill="#EF4444", outline="")
        self.create_rectangle(cx - 12, bt + 37, cx + 12, bt + 41,
                               fill="#EF4444", outline="")

        # Breast pocket
        self.create_rectangle(cx + 12, bt + 20, cx + 28, bt + 36,
                               fill="#F9FAFB", outline="#9CA3AF", width=1)
        # Pen in pocket
        self.create_line(cx + 18, bt + 20, cx + 18, bt + 28,
                         fill="#2563EB", width=2)

        # ── STETHOSCOPE ──────────────────────────────────────────────
        swing_speed = 0.20 if active else 0.05
        swing_amp   = 26    if active else 6
        swing = math.sin(t * swing_speed) * swing_amp

        self.create_arc(cx - 20, bt - 8, cx + 20, bt + 22,
                        start=215, extent=110,
                        style="arc", outline="#374151", width=3)

        rad = math.radians(swing)
        end_x = cx + math.sin(rad) * 20
        end_y = bt + 54 + math.cos(rad) * 10
        self.create_line(cx, bt + 14, end_x, end_y,
                         fill="#374151", width=3, smooth=True)
        self.create_oval(end_x - 7, end_y - 7, end_x + 7, end_y + 7,
                         fill="#1F2937", outline="#374151", width=2)

        # ── ARMS ─────────────────────────────────────────────────────
        arm_speed = 0.13 if active else 0.03
        arm_amp   = 9    if active else 2

        # Left arm
        lax = math.sin(t * arm_speed) * arm_amp
        self.create_line(cx - 30, bt + 16, cx - 58, bt + 34 + lax,
                         fill="#FDDCB5", width=11, capstyle="round")
        self.create_oval(cx - 65, bt + 28 + lax, cx - 51, bt + 42 + lax,
                         fill="#FDDCB5", outline="#C9966A", width=1)

        # Right arm + document
        rax = math.sin(t * arm_speed + 1.2) * arm_amp
        self.create_line(cx + 30, bt + 16, cx + 58, bt + 30 + rax,
                         fill="#FDDCB5", width=11, capstyle="round")

        # Document / DOCX file the doctor is examining
        dx = cx + 56
        dy = bt + 24 + rax
        # Paper
        self.create_rectangle(dx - 12, dy, dx + 12, dy + 18,
                               fill="#FEF9C3", outline="#CA8A04", width=1)
        # Folded corner
        self.create_polygon(dx + 5, dy, dx + 12, dy + 7, dx + 12, dy, dx + 5, dy,
                            fill="#FDE68A", outline="#CA8A04", width=1)
        # Lines of text
        for i, w in enumerate([9, 9, 6]):
            yy = dy + 8 + i * 3
            self.create_line(dx - 8, yy, dx - 8 + w, yy,
                             fill="#92400E", width=1)
        # "DOCX" label
        self.create_text(dx, dy + 22, text=".docx",
                         font=("Segoe UI", 6, "bold"), fill="#92400E")

        # Pulse / heartbeat line during active conversion
        if active:
            pulse_x = cx - 28
            pulse_y = bt + 50
            pw = 56
            phase = (t * 5) % (pw * 2)
            pts = []
            for px_ in range(pw):
                rel = (px_ + phase) % (pw * 2)
                if 20 < rel < 22:
                    py_ = pulse_y - 12
                elif 22 < rel < 24:
                    py_ = pulse_y + 8
                elif 24 < rel < 26:
                    py_ = pulse_y
                else:
                    py_ = pulse_y
                pts.extend([pulse_x + px_, py_])
            if len(pts) >= 4:
                self.create_line(*pts, fill="#EF4444", width=2, smooth=False)

        # ── LEGS ─────────────────────────────────────────────────────
        self.create_rectangle(cx - 22, bt + 68, cx - 8,  bt + 100,
                               fill="#1E3A5F", outline="#111827", width=1)
        self.create_rectangle(cx + 8,  bt + 68, cx + 22, bt + 100,
                               fill="#1E3A5F", outline="#111827", width=1)

        # Shoes
        self.create_oval(cx - 26, bt + 95, cx - 4,  bt + 108,
                         fill="#111827", outline="")
        self.create_oval(cx + 4,  bt + 95, cx + 26, bt + 108,
                         fill="#111827", outline="")


# ──────────────────────────────────────────────────────────────────────


class DuplicateDialog(tk.Toplevel):
    REPLACE_ALL   = "replace_all"
    SKIP_EXISTING = "skip_existing"
    CANCEL        = "cancel"

    def __init__(self, parent, existing_paths, root_folder):
        super().__init__(parent)
        self.title("Vorhandene PDFs gefunden")
        self.resizable(False, False)
        self.grab_set()
        self.result = self.CANCEL

        names = [os.path.relpath(pdf_path(p), root_folder) for p in existing_paths]
        preview = "\n".join(f"  • {n}" for n in names[:10])
        if len(names) > 10:
            preview += f"\n  … und {len(names) - 10} weitere"

        tk.Label(self,
                 text=f"{len(names)} PDF(s) mit gleichem Namen existieren bereits:\n\n"
                      + preview + "\n\nWas soll passieren?",
                 font=("Segoe UI", 10), justify="left", wraplength=420).pack(
            padx=20, pady=(18, 14))

        bf = tk.Frame(self)
        bf.pack(pady=(0, 16))

        for text, cmd, bold, color, fg in [
            ("Alle ersetzen",       self._replace_all,   True,  "#2563EB", "white"),
            ("Nur neue konvertieren", self._skip_existing, False, "#E5E7EB", "black"),
            ("Abbrechen",           self._cancel,         False, "#E5E7EB", "black"),
        ]:
            tk.Button(bf, text=text, command=cmd, relief="flat", cursor="hand2",
                      font=("Segoe UI", 10, "bold" if bold else "normal"),
                      bg=color, fg=fg, padx=10, pady=4).pack(side="left", padx=6)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self._center(parent)
        self.wait_window()

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _replace_all(self):   self.result = self.REPLACE_ALL;   self.destroy()
    def _skip_existing(self): self.result = self.SKIP_EXISTING; self.destroy()
    def _cancel(self):        self.result = self.CANCEL;        self.destroy()


class ErrorDialog(tk.Toplevel):
    def __init__(self, parent, errors):
        super().__init__(parent)
        self.title("Fehler beim Konvertieren")
        self.resizable(True, True)
        self.grab_set()

        tk.Label(self,
                 text=f"{len(errors)} Datei(en) konnten nicht konvertiert werden.\n"
                      "Klick auf den Dateinamen öffnet ihn im Explorer:",
                 font=("Segoe UI", 10), justify="left").pack(
            padx=16, pady=(14, 6), anchor="w")

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        sb = tk.Scrollbar(frame)
        sb.pack(side="right", fill="y")

        text = tk.Text(frame, font=("Segoe UI", 10), wrap="none",
                       cursor="arrow", yscrollcommand=sb.set,
                       width=70, height=min(len(errors) + 2, 16),
                       relief="flat", bg=self.cget("bg"))
        text.pack(side="left", fill="both", expand=True)
        sb.config(command=text.yview)

        text.tag_config("err", foreground="#6B7280")

        for path, exc in errors:
            tag = f"lnk_{id(path)}"
            text.tag_config(tag, foreground="#2563EB", underline=True)
            text.tag_bind(tag, "<Button-1>", lambda _e, p=path: open_in_explorer(p))
            text.tag_bind(tag, "<Enter>",    lambda _e: text.config(cursor="hand2"))
            text.tag_bind(tag, "<Leave>",    lambda _e: text.config(cursor="arrow"))
            text.insert("end", os.path.basename(path), (tag,))
            text.insert("end", f"  –  {exc}\n", ("err",))

        text.config(state="disabled")

        tk.Button(self, text="Schließen", command=self.destroy,
                  font=("Segoe UI", 10), bg="#E5E7EB", relief="flat",
                  cursor="hand2").pack(pady=(0, 14))

        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")


# ──────────────────────────────────────────────────────────────────────


class DrPdfApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dr. PDF")
        self.resizable(False, False)
        self.configure(bg="#F9FAFB")
        self._build_ui()
        self._center()

    def _build_ui(self):
        bg = "#F9FAFB"

        # ── Title ──
        tk.Label(self, text="Dr. PDF", font=("Segoe UI", 26, "bold"),
                 fg="#2563EB", bg=bg).pack(pady=(20, 2))
        tk.Label(self, text="Doc. PDF  ·  DOCX → PDF  ·  inkl. Unterordner",
                 font=("Segoe UI", 9), fg="#9CA3AF", bg=bg).pack(pady=(0, 10))

        # ── Doctor ──
        self.doctor = DoctorCanvas(self, bg=bg)
        self.doctor.pack(pady=(0, 10))

        # ── Folder selection ──
        row = tk.Frame(self, bg=bg)
        row.pack(padx=20, pady=(0, 4), fill="x")

        tk.Label(row, text="Verzeichnis:", font=("Segoe UI", 10), bg=bg).pack(side="left")

        self.folder_var = tk.StringVar()
        tk.Entry(row, textvariable=self.folder_var, width=40,
                 font=("Segoe UI", 10)).pack(side="left", padx=(8, 6))

        tk.Button(row, text="Durchsuchen", command=self._browse,
                  font=("Segoe UI", 10), bg="#E5E7EB", relief="flat",
                  cursor="hand2", pady=2).pack(side="left")

        # ── Status ──
        self.status_label = tk.Label(self, text="", font=("Segoe UI", 9),
                                     fg="#374151", bg=bg, wraplength=480)
        self.status_label.pack(pady=(6, 2))

        # ── Progress ──
        self.progress = ttk.Progressbar(self, length=480, mode="determinate")
        self.progress.pack(padx=20, pady=(2, 10))

        # ── Convert button ──
        self.convert_btn = tk.Button(
            self, text="Konvertieren", command=self._start_conversion,
            font=("Segoe UI", 12, "bold"), bg="#2563EB", fg="white",
            relief="flat", cursor="hand2", padx=28, pady=9,
        )
        self.convert_btn.pack(pady=(0, 22))

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    def _browse(self):
        folder = filedialog.askdirectory(title="Verzeichnis wählen")
        if folder:
            self.folder_var.set(folder)
            files = find_docx(folder)
            if files:
                self.status_label.config(
                    text=f"{len(files)} .docx-Datei(en) gefunden.", fg="#059669")
            else:
                self.status_label.config(
                    text="Keine .docx-Dateien gefunden.", fg="#EF4444")

    def _start_conversion(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Kein Verzeichnis",
                                   "Bitte wähle zuerst ein Verzeichnis.")
            return

        all_files = find_docx(folder)
        if not all_files:
            messagebox.showinfo("Keine Dateien",
                                "Keine .docx-Dateien gefunden (auch nicht in Unterordnern).")
            return

        existing = [p for p in all_files if os.path.exists(pdf_path(p))]
        files_to_convert = all_files

        if existing:
            dlg = DuplicateDialog(self, existing, folder)
            if dlg.result == DuplicateDialog.CANCEL:
                return
            if dlg.result == DuplicateDialog.SKIP_EXISTING:
                files_to_convert = [p for p in all_files
                                    if not os.path.exists(pdf_path(p))]
                if not files_to_convert:
                    messagebox.showinfo("Nichts zu tun",
                                        "Alle PDFs existieren bereits.")
                    return

        self.convert_btn.config(state="disabled")
        self.progress["maximum"] = len(files_to_convert)
        self.progress["value"] = 0
        self.doctor.start()

        threading.Thread(
            target=self._convert, args=(files_to_convert, folder), daemon=True
        ).start()

    def _convert(self, files, root_folder):
        errors = []
        for i, path in enumerate(files, 1):
            rel = os.path.relpath(path, root_folder)
            self.status_label.config(
                text=f"({i}/{len(files)})  {rel}", fg="#374151")
            try:
                convert_via_temp(path, pdf_path(path))
            except Exception as e:
                errors.append((path, str(e)))
            self.progress["value"] = i
            self.update_idletasks()

        self.doctor.stop()
        self.convert_btn.config(state="normal")
        done = len(files) - len(errors)

        if errors:
            self.status_label.config(
                text=f"{done} konvertiert, {len(errors)} Fehler.", fg="#EF4444")
            self.after(0, lambda: ErrorDialog(self, errors))
        else:
            self.status_label.config(
                text=f"Alle {len(files)} Datei(en) erfolgreich konvertiert!",
                fg="#059669")
            messagebox.showinfo("Fertig",
                                f"{len(files)} Datei(en) wurden als PDF gespeichert.")


if __name__ == "__main__":
    try:
        app = DrPdfApp()
        app.mainloop()
    except Exception as exc:
        import traceback
        tk.Tk().withdraw()
        messagebox.showerror(
            "Dr. PDF – Startfehler",
            f"{exc}\n\n{traceback.format_exc()}"
        )
