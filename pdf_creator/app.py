"""Interfaz gráfica (Tkinter) para combinar 2 documentos de Word en un PDF."""

from __future__ import annotations

import os
import platform
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .core import combine_word_documents

WORD_FILETYPES = [
    ("Documentos de Word", "*.docx *.doc *.docm *.rtf *.odt"),
    ("Todos los archivos", "*.*"),
]


def open_file(path: Path) -> None:
    """Abre un archivo con la aplicación predeterminada del sistema."""
    system = platform.system()
    if system == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
    elif system == "Darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class PdfCreatorApp(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=16)
        self.master = master
        master.title("PDF Creator – Combinar Word en PDF")
        master.minsize(560, 0)

        self.doc1 = tk.StringVar()
        self.doc2 = tk.StringVar()
        self.output = tk.StringVar()
        self.open_after = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Elegí los dos documentos de Word.")

        self._build()
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

    # ------------------------------------------------------------------ UI
    def _build(self) -> None:
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="Combinar 2 documentos de Word en un PDF",
                  font=("TkDefaultFont", 12, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        self._file_row(1, "Documento 1:", self.doc1, self._pick_doc1)
        self._file_row(2, "Documento 2:", self.doc2, self._pick_doc2)

        ttk.Button(self, text="⇅ Invertir orden", command=self._swap).grid(
            row=3, column=1, sticky="w", pady=(0, 8))

        ttk.Separator(self).grid(row=4, column=0, columnspan=3, sticky="ew", pady=8)

        self._file_row(5, "PDF de salida:", self.output, self._pick_output, button_text="Guardar como…")

        ttk.Checkbutton(self, text="Abrir el PDF al terminar", variable=self.open_after).grid(
            row=6, column=1, sticky="w", pady=(4, 8))

        self.generate_btn = ttk.Button(self, text="Generar PDF", command=self._generate)
        self.generate_btn.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(4, 8), ipady=4)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.grid(row=8, column=0, columnspan=3, sticky="ew")

        ttk.Label(self, textvariable=self.status, foreground="#555", wraplength=520).grid(
            row=9, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _file_row(self, row, label, var, command, button_text="Examinar…") -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(self, textvariable=var).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(self, text=button_text, command=command).grid(row=row, column=2, padx=(8, 0), pady=4)

    # ------------------------------------------------------------ actions
    def _pick_word(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(title="Elegir documento de Word", filetypes=WORD_FILETYPES)
        if path:
            var.set(path)
            if not self.output.get():
                first = Path(self.doc1.get() or path)
                self.output.set(str(first.with_name(first.stem + "_combinado.pdf")))

    def _pick_doc1(self) -> None:
        self._pick_word(self.doc1)

    def _pick_doc2(self) -> None:
        self._pick_word(self.doc2)

    def _pick_output(self) -> None:
        initial = Path(self.output.get()) if self.output.get() else None
        path = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialdir=str(initial.parent) if initial else None,
            initialfile=initial.name if initial else "combinado.pdf",
        )
        if path:
            self.output.set(path)

    def _swap(self) -> None:
        a, b = self.doc1.get(), self.doc2.get()
        self.doc1.set(b)
        self.doc2.set(a)

    def _generate(self) -> None:
        doc1, doc2, out = self.doc1.get().strip(), self.doc2.get().strip(), self.output.get().strip()
        for label, value in (("Documento 1", doc1), ("Documento 2", doc2)):
            if not value:
                messagebox.showwarning("Falta un archivo", f"Elegí el {label}.")
                return
            if not Path(value).is_file():
                messagebox.showerror("Archivo no encontrado", f"No existe el {label}:\n{value}")
                return
        if not out:
            messagebox.showwarning("Falta la salida", "Elegí dónde guardar el PDF.")
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
            self.output.set(out)

        self._set_busy(True)
        threading.Thread(target=self._worker, args=([Path(doc1), Path(doc2)], Path(out)),
                         daemon=True).start()

    def _worker(self, docs: list[Path], out: Path) -> None:
        try:
            result = combine_word_documents(docs, out, progress=self._report)
        except Exception as exc:
            self.after(0, self._on_error, exc)
        else:
            self.after(0, self._on_success, result)

    def _report(self, message: str) -> None:
        self.after(0, self.status.set, message)

    def _on_success(self, result: Path) -> None:
        self._set_busy(False)
        self.status.set(f"PDF generado: {result}")
        if self.open_after.get():
            try:
                open_file(result)
            except Exception:
                pass
        else:
            messagebox.showinfo("Listo", f"PDF generado correctamente:\n{result}")

    def _on_error(self, exc: Exception) -> None:
        self._set_busy(False)
        self.status.set("Ocurrió un error.")
        messagebox.showerror("Error", str(exc))

    def _set_busy(self, busy: bool) -> None:
        if busy:
            self.generate_btn.state(["disabled"])
            self.progress.start(12)
        else:
            self.generate_btn.state(["!disabled"])
            self.progress.stop()


def main() -> None:
    root = tk.Tk()
    try:
        ttk.Style(root).theme_use("vista" if platform.system() == "Windows" else "clam")
    except tk.TclError:
        pass
    PdfCreatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
