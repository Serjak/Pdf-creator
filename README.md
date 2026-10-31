# PDF Creator

App de escritorio para elegir **2 documentos (PDF o Word)** y generar **un único PDF** con ambos combinados (primero el documento 1, después el documento 2). Se pueden combinar 2 PDF, 2 Word o un PDF y un Word.

## Instalar en una PC sin Python (Windows)

No hace falta instalar Python. Solo hay que descargar `PDFCreator.exe`:

1. Entrá a la pestaña **Releases** del repositorio y descargá `PDFCreator.exe` de la última versión.
   (Otra opción: en **Actions → Compilar ejecutable de Windows**, abrí la última ejecución exitosa y descargá el artefacto `PDFCreator-windows`, que es un .zip con el .exe adentro.)
2. Copiá el `.exe` a la PC (escritorio, pendrive, donde quieras) y abrilo con doble clic.
   Si Windows muestra *"Windows protegió su PC"*, hacé clic en **Más información → Ejecutar de todas formas** (aparece porque el ejecutable no está firmado digitalmente).

Funciona en Windows 7, 8, 10 y 11 (32 o 64 bits). En Windows 7, si aparece un error por falta de `api-ms-win-crt-runtime-l1-1-0.dll`, instalá el *Visual C++ Redistributable 2015-2022 (x86)* de Microsoft.

**Para combinar solo PDF no hace falta nada más.** Para usar documentos de Word, la PC necesita **Microsoft Word** o **LibreOffice** (gratis: https://www.libreoffice.org/download/) para convertirlos. La app usa Word si está; si no, LibreOffice.

## Uso

1. Hacé clic en **Examinar…** para elegir el *Documento 1* y el *Documento 2* (`.pdf`, `.docx`, `.doc`, `.rtf`, `.odt`). Los PDF protegidos con contraseña para abrirlos no se pueden combinar.
2. Usá **⇅ Invertir orden** si querés cambiar cuál va primero.
3. Elegí dónde guardar el PDF (por defecto `<documento1>_combinado.pdf`).
4. Hacé clic en **Generar PDF**.

## Para desarrolladores

Con Python 3.9+:

```bash
pip install -r requirements.txt
python run.py
```

El `.exe` se compila automáticamente en GitHub Actions con cada push (`.github/workflows/build-windows.yml`). Además se publica un Release con el `.exe` al lanzar el workflow a mano (**Actions → Compilar ejecutable de Windows → Run workflow**), con cada push a `main` o al subir un tag `v*`. La versión sale de `pdf_creator/__init__.py`. También se puede compilar localmente en Windows con `build_exe.bat`.

## Estructura

- `pdf_creator/core.py` – conversión Word → PDF (Word o LibreOffice) y combinación de PDFs con `pypdf`.
- `pdf_creator/app.py` – interfaz gráfica en Tkinter.
- `run.py` – punto de entrada.
