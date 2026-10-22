
from pathlib import Path
from io import BytesIO

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
from pdf2docx import Converter


# ==================================================
# CONFIGURACIÓN
# ==================================================

BASE_DIR = Path(__file__).resolve().parent
PDF_DIR = BASE_DIR / "PDF"
WORD_DIR = BASE_DIR / "Word"

# Idiomas de OCR: español + inglés.
OCR_LANG = "spa+eng"

# Resolución del escaneo. 300 DPI suele dar buenos resultados.
OCR_DPI = 300

# Umbral para determinar si el PDF tiene texto digital.
MIN_TEXT_CHARS = 30

# Si Tesseract no está en el PATH de Windows, descomenta
# la siguiente línea y ajusta la ruta a tu instalación.
#
# pytesseract.pytesseract.tesseract_cmd = (
#     r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# )


# ==================================================
# DETECTAR SI EL PDF TIENE TEXTO DIGITAL
# ==================================================

def tiene_texto_digital(pdf_path: Path) -> bool:
    """Comprueba si el PDF contiene suficiente texto extraíble."""

    with fitz.open(pdf_path) as pdf:
        texto_total = sum(
            len(pagina.get_text("text").strip())
            for pagina in pdf
        )

    return texto_total >= MIN_TEXT_CHARS


# ==================================================
# CONVERSIÓN DE PDF DIGITAL A WORD
# ==================================================

def convertir_pdf_digital(
    pdf_path: Path,
    docx_path: Path
) -> None:
    """Convierte un PDF digital conservando su estructura."""

    converter = Converter(str(pdf_path))
    try:
        converter.convert(str(docx_path))
    finally:
        converter.close()


# ==================================================
# OCR: PDF ESCANEADO A WORD EDITABLE
# ==================================================

def convertir_pdf_ocr(
    pdf_path: Path,
    docx_path: Path
) -> None:
    """
    Reconoce el texto de cada página mediante OCR
    y crea un DOCX con texto editable.
    """

    documento = Document()

    with fitz.open(pdf_path) as pdf:
        total_paginas = len(pdf)

        for numero, pagina in enumerate(pdf, start=1):
            print(
                f"    OCR página {numero}/{total_paginas}..."
            )

            # Renderizar la página como imagen de alta resolución.
            pix = pagina.get_pixmap(
                dpi=OCR_DPI,
                colorspace=fitz.csRGB,
                alpha=False
            )

            imagen = Image.open(
                BytesIO(pix.tobytes("png"))
            )

            # Reconocer el texto de la imagen.
            texto = pytesseract.image_to_string(
                imagen,
                lang=OCR_LANG,
                config="--psm 3"
            )

            # Agregar el texto reconocido al documento.
            if numero > 1:
                documento.add_page_break()

            for linea in texto.splitlines():
                linea = linea.strip()

                if linea:
                    documento.add_paragraph(linea)

    documento.save(docx_path)


# ==================================================
# PROCESAMIENTO PRINCIPAL
# ==================================================

def convertir_pdf(
    pdf_path: Path,
    docx_path: Path
) -> str:
    """Selecciona el método de conversión apropiado."""

    if tiene_texto_digital(pdf_path):
        convertir_pdf_digital(pdf_path, docx_path)
        return "PDF digital"

    convertir_pdf_ocr(pdf_path, docx_path)
    return "PDF escaneado con OCR"


def main() -> int:
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    WORD_DIR.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(
        PDF_DIR.glob("*.pdf")
    )

    if not pdf_files:
        print(f"No se encontraron PDF en: {PDF_DIR}")
        return 0

    failures = 0

    for pdf_path in pdf_files:
        docx_path = WORD_DIR / f"{pdf_path.stem}.docx"

        if docx_path.exists():
            print(f"Se omite (ya existe): {docx_path.name}")
            continue

        try:
            print(f"\nProcesando: {pdf_path.name}")

            metodo = convertir_pdf(
                pdf_path,
                docx_path
            )

            print(
                f"Convertido: {pdf_path.name} -> "
                f"{docx_path.name} [{metodo}]"
            )

        except Exception as error:
            failures += 1

            # Evitar dejar un DOCX incompleto si falla.
            if docx_path.exists():
                docx_path.unlink()

            print(
                f"Error al convertir {pdf_path.name}: {error}"
            )

    print("\nProceso terminado.")
    print(f"Archivos encontrados: {len(pdf_files)}")
    print(f"Errores: {failures}")

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
