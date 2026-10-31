"""Lógica de conversión Word -> PDF y combinación de PDFs (sin interfaz)."""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable, Iterable

from pypdf import PdfWriter

WORD_EXTENSIONS = {".docx", ".doc", ".docm", ".rtf", ".odt"}


class ConversionError(Exception):
    """Error al convertir un documento a PDF."""


def _find_libreoffice() -> str | None:
    """Busca el ejecutable de LibreOffice en el PATH y en rutas habituales."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found

    candidates: list[str] = []
    system = platform.system()
    if system == "Windows":
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if base:
                candidates.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    elif system == "Darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")

    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def _convert_with_word_windows(src: Path, dst: Path) -> None:
    """Convierte usando Microsoft Word en Windows vía COM (pywin32)."""
    import pythoncom  # import diferido: solo existe en Windows
    import win32com.client

    pythoncom.CoInitialize()  # necesario: se llama desde un hilo secundario
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(src), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        )
        try:
            doc.ExportAsFixedFormat(str(dst), 17)  # 17 = wdExportFormatPDF
        finally:
            doc.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def _convert_with_word(src: Path, dst: Path) -> None:
    """Convierte usando Microsoft Word (Windows por COM, macOS vía docx2pdf)."""
    if platform.system() == "Windows":
        _convert_with_word_windows(src, dst)
    else:
        from docx2pdf import convert  # import diferido: dependencia opcional

        convert(str(src), str(dst))
    if not dst.exists():
        raise ConversionError("Microsoft Word no generó el PDF.")


def _convert_with_libreoffice(soffice: str, src: Path, out_dir: Path) -> Path:
    """Convierte usando LibreOffice en modo headless."""
    # Perfil de usuario aislado para no chocar con una instancia abierta.
    profile_dir = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    try:
        cmd = [
            soffice,
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(src),
        ]
        extra = {"creationflags": subprocess.CREATE_NO_WINDOW} if os.name == "nt" else {}
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300, **extra)
    except subprocess.TimeoutExpired as exc:
        raise ConversionError(f"LibreOffice tardó demasiado en convertir {src.name}.") from exc
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    produced = out_dir / (src.stem + ".pdf")
    if result.returncode != 0 or not produced.exists():
        detail = (result.stderr or result.stdout or "").strip()
        raise ConversionError(f"LibreOffice no pudo convertir {src.name}. {detail}")
    return produced


def convert_to_pdf(src: Path, out_dir: Path) -> Path:
    """Convierte un documento de Word a PDF dentro de ``out_dir``.

    Intenta primero con Microsoft Word (si está instalado y docx2pdf está
    disponible) y, si no, con LibreOffice.
    """
    src = Path(src).resolve()
    if not src.is_file():
        raise ConversionError(f"No se encontró el archivo: {src}")
    if src.suffix.lower() not in WORD_EXTENSIONS:
        raise ConversionError(f"Formato no soportado: {src.name}")

    errors: list[str] = []

    if platform.system() in ("Windows", "Darwin") and src.suffix.lower() in {".docx", ".doc", ".docm", ".rtf"}:
        dst = out_dir / (src.stem + ".pdf")
        try:
            _convert_with_word(src, dst)
            return dst
        except ImportError:
            errors.append("falta pywin32/docx2pdf para usar Microsoft Word")
        except Exception as exc:  # Word no instalado, error de COM, etc.
            errors.append(f"Microsoft Word: {exc}")

    soffice = _find_libreoffice()
    if soffice:
        try:
            return _convert_with_libreoffice(soffice, src, out_dir)
        except ConversionError as exc:
            errors.append(str(exc))
    else:
        errors.append("LibreOffice no está instalado")

    raise ConversionError(
        "No se pudo convertir el documento a PDF.\n"
        "Instalá Microsoft Word o LibreOffice (https://www.libreoffice.org).\n\n"
        + "\n".join(f"• {e}" for e in errors)
    )


def merge_pdfs(pdfs: Iterable[Path], output: Path) -> None:
    """Combina varios PDFs, en orden, en un único archivo."""
    writer = PdfWriter()
    for pdf in pdfs:
        writer.append(str(pdf))
    output.parent.mkdir(parents=True, exist_ok=True)
    with open(output, "wb") as fh:
        writer.write(fh)
    writer.close()


def combine_word_documents(
    documents: list[Path],
    output: Path,
    progress: Callable[[str], None] | None = None,
) -> Path:
    """Convierte cada documento de Word a PDF y los combina en ``output``."""
    if len(documents) < 2:
        raise ValueError("Se necesitan al menos 2 documentos.")

    report = progress or (lambda _msg: None)
    output = Path(output)

    with tempfile.TemporaryDirectory(prefix="pdf_creator_") as tmp:
        tmp_dir = Path(tmp)
        pdfs: list[Path] = []
        for index, doc in enumerate(documents, start=1):
            report(f"Convirtiendo {index}/{len(documents)}: {Path(doc).name}…")
            # Subcarpeta por documento para evitar choques si tienen el mismo nombre.
            sub = tmp_dir / str(index)
            sub.mkdir()
            pdfs.append(convert_to_pdf(Path(doc), sub))

        report("Combinando documentos…")
        merge_pdfs(pdfs, output)

    report(f"Listo: {output}")
    return output
