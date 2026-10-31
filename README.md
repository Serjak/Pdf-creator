# PDF Creator

App de escritorio para elegir **2 documentos de Word** y generar **un único PDF** con ambos combinados (primero el documento 1, después el documento 2).

## Instalar en una PC sin Python (Windows)

No hace falta instalar Python. Solo hay que descargar `PDFCreator.exe`:

1. Entrá a la pestaña **Releases** del repositorio y descargá `PDFCreator.exe` de la última versión.
   (Otra opción: en **Actions → Compilar ejecutable de Windows**, abrí la última ejecución exitosa y descargá el artefacto `PDFCreator-windows`, que es un .zip con el .exe adentro.)
2. Copiá el `.exe` a la PC (escritorio, pendrive, donde quieras) y abrilo con doble clic.
   Si Windows muestra *"Windows protegió su PC"*, hacé clic en **Más información → Ejecutar de todas formas** (aparece porque el ejecutable no está firmado digitalmente).

**Importante:** la PC necesita **Microsoft Word** o **LibreOffice** (gratis: https://www.libreoffice.org/download/) para convertir los documentos. La app usa Word si está; si no, LibreOffice.

## Uso

1. Hacé clic en **Examinar…** para elegir el *Documento 1* y el *Documento 2* (`.docx`, `.doc`, `.rtf`, `.odt`).
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

- `pdf_creator/core.py` – conversión Word → PDF (Word o LibreOffice) y combinación con `pypdf`.
- `pdf_creator/app.py` – interfaz gráfica en Tkinter.
- `run.py` – punto de entrada.
