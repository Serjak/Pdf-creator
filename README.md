# PDF Creator

App de escritorio para elegir **2 documentos de Word** y generar **un único PDF** con ambos combinados (primero el documento 1, después el documento 2).

## Requisitos

- Python 3.9 o superior (con Tkinter, que viene incluido en el instalador oficial de Windows/macOS).
- Para convertir Word → PDF hace falta **una** de estas dos opciones:
  - **Microsoft Word** instalado (Windows o macOS), o
  - **LibreOffice** (gratis): https://www.libreoffice.org/download/

La app usa Word si está disponible; si no, usa LibreOffice automáticamente.

## Instalación y uso

```bash
pip install -r requirements.txt
python run.py
```

1. Hacé clic en **Examinar…** para elegir el *Documento 1* y el *Documento 2* (`.docx`, `.doc`, `.rtf`, `.odt`).
2. Usá **⇅ Invertir orden** si querés cambiar cuál va primero.
3. Elegí dónde guardar el PDF (por defecto `<documento1>_combinado.pdf`).
4. Hacé clic en **Generar PDF**.

## Crear un ejecutable (.exe) para Windows

Doble clic en `build_exe.bat` (o ejecutalo desde una consola). El ejecutable queda en `dist\PDFCreator.exe` y se puede usar sin abrir Python.

## Estructura

- `pdf_creator/core.py` – conversión Word → PDF (Word o LibreOffice) y combinación con `pypdf`.
- `pdf_creator/app.py` – interfaz gráfica en Tkinter.
- `run.py` – punto de entrada.
