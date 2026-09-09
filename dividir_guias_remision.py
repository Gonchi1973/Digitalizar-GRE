#!/usr/bin/env python3
"""
Divide uno o varios PDF de guias de remision en archivos de una pagina.

El nombre de cada salida usa el numero de la GRE y el destinatario:
    EG07-00000038 - Whan Li.pdf

Funciona con PDF digitales y con PDF escaneados. Para estos ultimos usa OCR.

Dependencias Python:
    python -m pip install pypdf pymupdf pillow pytesseract

En desarrollo, para documentos escaneados se requiere Tesseract OCR:
    Ubuntu/WSL: sudo apt update && sudo apt install -y tesseract-ocr tesseract-ocr-spa

La compilacion portable realizada con compilar_portable.ps1 incluye Tesseract
y no requiere instalar Python ni Tesseract en la computadora de destino.

Uso con ventanas de seleccion:
    python dividir_guias_remision.py

Uso desde terminal:
    python dividir_guias_remision.py archivo1.pdf archivo2.pdf -o Guias_separadas

Se pueden seleccionar varios PDF en una sola ejecucion. Si una misma guia se
repite dentro de un PDF, entre varios PDF o ya existe en la carpeta de salida,
se conserva un solo archivo y la repeticion se registra en el reporte CSV.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path

try:
    import pymupdf
    import pytesseract
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps
    from pypdf import PdfReader, PdfWriter
except ImportError as exc:
    paquete = getattr(exc, "name", "una dependencia")
    raise SystemExit(
        f"Falta instalar {paquete}. Ejecuta:\n"
        "python -m pip install pypdf pymupdf pillow pytesseract"
    ) from exc


# Serie usada actualmente por Patocentro. Si SUNAT asigna otra, modificala aqui
# o usa el parametro --serie al ejecutar el programa.
SERIE_PREDETERMINADA = "EG07"
MINIMO_TEXTO_DIGITAL = 80


@dataclass
class Resultado:
    pdf_origen: str
    pagina: int
    numero_guia: str
    destinatario: str
    archivo_salida: str
    estado: str


def carpeta_recursos() -> Path:
    """Devuelve la carpeta real o la carpeta temporal creada por PyInstaller."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def localizar_tesseract() -> str | None:
    """Localiza Tesseract empaquetado, instalado o disponible en PATH."""
    base = carpeta_recursos()
    junto_al_ejecutable = Path(sys.executable).resolve().parent
    candidatos = (
        base / "Tesseract-OCR" / "tesseract.exe",
        junto_al_ejecutable / "Tesseract-OCR" / "tesseract.exe",
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    )
    for candidato in candidatos:
        if candidato.is_file():
            return str(candidato)

    encontrado = shutil.which("tesseract")
    if encontrado:
        return encontrado
    return None


def configurar_tesseract() -> str:
    """Configura el ejecutable y los idiomas, incluidos los empaquetados."""
    ejecutable = localizar_tesseract()
    if not ejecutable:
        raise RuntimeError(
            "No se encontro Tesseract OCR dentro del programa ni instalado en la PC."
        )

    pytesseract.pytesseract.tesseract_cmd = ejecutable
    tessdata = Path(ejecutable).resolve().parent / "tessdata"
    if tessdata.is_dir():
        os.environ["TESSDATA_PREFIX"] = str(tessdata)
    return ejecutable


def quitar_acentos(texto: str) -> str:
    return "".join(
        caracter
        for caracter in unicodedata.normalize("NFD", texto)
        if unicodedata.category(caracter) != "Mn"
    )


def limpiar_espacios(texto: str) -> str:
    return re.sub(r"\s+", " ", texto).strip(" \t\r\n:;-_")


def normalizar_numero(texto: str) -> str:
    traduccion = str.maketrans({"O": "0", "I": "1", "L": "1"})
    return re.sub(r"\D", "", texto.upper().translate(traduccion))


def extraer_numero_guia(texto: str, serie_preferida: str) -> str | None:
    plano = quitar_acentos(texto).upper()
    serie_preferida = serie_preferida.upper().replace("-", "").strip()

    # En los escaneos de Patocentro, OCR suele leer EG07 como EGO7.
    variantes = {
        serie_preferida,
        serie_preferida.replace("0", "O"),
        serie_preferida.replace("0", "Q"),
    }
    patron_series = "|".join(re.escape(x) for x in sorted(variantes, key=len, reverse=True))
    coincidencia = re.search(
        rf"(?:N\s*[°ºO]?\s*)?(?:{patron_series})\s*[-–—:]?\s*([0-9OIL]{{6,10}})",
        plano,
    )
    if coincidencia:
        correlativo = normalizar_numero(coincidencia.group(1))
        return f"{serie_preferida}-{correlativo.zfill(8)}"

    # OCR puede deformar un caracter de la serie (p. ej., EG07 -> EGO?).
    # Si la linea esta rotulada con N°, conserva la serie configurada y toma
    # solamente el correlativo, que suele reconocerse con mucha mas precision.
    coincidencia = re.search(
        r"N\s*[°ºO]?\s*[A-Z0-9OIL?]{3,6}\s*[-–—:]\s*([0-9OIL]{6,10})",
        plano,
    )
    if coincidencia:
        correlativo = normalizar_numero(coincidencia.group(1))
        return f"{serie_preferida}-{correlativo.zfill(8)}"

    # Respaldo para otras series SUNAT: letra(s) + digitos, seguidos del correlativo.
    coincidencia = re.search(
        r"(?:N\s*[°ºO]?\s*)?([A-Z]{1,3}[0-9OIL]{1,3})\s*[-–—:]\s*([0-9OIL]{6,10})",
        plano,
    )
    if not coincidencia:
        return None

    serie = coincidencia.group(1)
    correlativo = normalizar_numero(coincidencia.group(2))
    return f"{serie}-{correlativo.zfill(8)}"


def quitar_tipo_societario(nombre: str) -> str:
    """Elimina SAC, SRL o EIRL al final, con o sin puntos y espacios."""
    tipo_societario = (
        r"(?:"
        # En documentos escaneados OCR puede confundir S con $/5 y C con G.
        r"[S$5]\s*\.?\s*A\s*\.?\s*[CG]\s*\.?"  # SAC, $.A.C., S.AG.
        r"|[S$5]\s*\.?\s*R\s*\.?\s*L\s*\.?"    # SRL, S.R.L.
        r"|E\s*\.?\s*I\s*\.?\s*R\s*\.?\s*L\s*\.?"  # EIRL, E.I.R.L.
        r")"
    )
    resultado = nombre.strip()
    resultado = re.sub(
        rf"(?:\s*[-,]?\s*{tipo_societario})+\s*[.,-]*\s*$",
        "",
        resultado,
        flags=re.IGNORECASE,
    )
    return resultado.strip(" .,-")


def capitalizar_nombre(nombre: str) -> str:
    palabras_mayusculas = {"II", "III", "IV"}
    resultado: list[str] = []
    for palabra in nombre.lower().split():
        candidata = "-".join(parte.capitalize() for parte in palabra.split("-"))
        if candidata.upper() in palabras_mayusculas:
            candidata = candidata.upper()
        resultado.append(candidata)
    return " ".join(resultado)


def extraer_destinatario(texto: str) -> str | None:
    coincidencia = re.search(
        r"Datos\s+d[e3][l\]1Ii|]?\s+Destinatar[iIl1]o\s*[:;,.]?\s*(.+?)\s*[-–—]\s*"
        r"REGISTRO\s+[ÚU]NICO\s+DE\s+CONTRIBUYENTES",
        texto,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not coincidencia:
        # Respaldo cuando OCR pierde parte de la etiqueta o del separador.
        coincidencia = re.search(
            r"Destinatar[iIl1]o\s*[:;,.]?\s*(.+?)(?:\s*[-–—]\s*)?REGISTRO\s+[ÚU]NICO",
            texto,
            flags=re.IGNORECASE | re.DOTALL,
        )
    if not coincidencia:
        return None

    nombre = limpiar_espacios(coincidencia.group(1))
    nombre = quitar_tipo_societario(nombre)
    return capitalizar_nombre(nombre) if nombre else None


def nombre_seguro(texto: str, maximo: int = 120) -> str:
    texto = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", texto)
    texto = limpiar_espacios(texto).rstrip(". ")
    return texto[:maximo].rstrip(". ") or "Sin_nombre"


def ruta_sin_colision(carpeta: Path, base: str) -> Path:
    candidato = carpeta / f"{base}.pdf"
    contador = 2
    while candidato.exists():
        candidato = carpeta / f"{base} ({contador}).pdf"
        contador += 1
    return candidato


def obtener_guias_existentes(carpeta: Path, serie: str) -> dict[str, str]:
    """Indexa por numero las guias que ya existen en la carpeta de salida."""
    existentes: dict[str, str] = {}
    for archivo in sorted(carpeta.glob("*.pdf")):
        numero = extraer_numero_guia(archivo.stem, serie)
        if numero and numero not in existentes:
            existentes[numero] = archivo.name
    return existentes


def texto_ocr_pagina(documento: pymupdf.Document, indice: int) -> str:
    configurar_tesseract()

    pagina = documento.load_page(indice)
    # Ambos datos requeridos aparecen en la mitad superior de la GRE.
    area = pymupdf.Rect(0, 0, pagina.rect.width, pagina.rect.height * 0.48)
    pixmap = pagina.get_pixmap(
        matrix=pymupdf.Matrix(2.5, 2.5), clip=area, alpha=False
    )
    imagen = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    imagen = ImageOps.grayscale(imagen)
    imagen = ImageOps.autocontrast(imagen)
    imagen = ImageEnhance.Contrast(imagen).enhance(1.25)
    imagen = imagen.filter(ImageFilter.SHARPEN)

    idiomas = set(pytesseract.get_languages(config=""))
    if not idiomas.intersection({"spa", "eng"}):
        raise RuntimeError(
            "Tesseract fue encontrado, pero no contiene los idiomas spa o eng."
        )
    idioma = "spa+eng" if "spa" in idiomas and "eng" in idiomas else ("spa" if "spa" in idiomas else "eng")
    return pytesseract.image_to_string(imagen, lang=idioma, config="--psm 3")


def obtener_datos_pagina(
    lector: PdfReader,
    documento_ocr: pymupdf.Document,
    indice: int,
    serie: str,
) -> tuple[str | None, str | None, bool]:
    texto = lector.pages[indice].extract_text() or ""
    numero = extraer_numero_guia(texto, serie)
    destinatario = extraer_destinatario(texto)
    uso_ocr = False

    if len(texto.strip()) < MINIMO_TEXTO_DIGITAL or not numero or not destinatario:
        texto_ocr = texto_ocr_pagina(documento_ocr, indice)
        uso_ocr = True
        numero = numero or extraer_numero_guia(texto_ocr, serie)
        destinatario = destinatario or extraer_destinatario(texto_ocr)

    return numero, destinatario, uso_ocr


def dividir_pdf(
    pdf: Path,
    carpeta_salida: Path,
    serie: str,
    guias_conservadas: dict[str, str],
) -> list[Resultado]:
    lector = PdfReader(str(pdf))
    if lector.is_encrypted:
        try:
            lector.decrypt("")
        except Exception as exc:
            raise RuntimeError(f"El PDF esta protegido y no pudo abrirse: {pdf.name}") from exc

    documento_ocr = pymupdf.open(str(pdf))
    resultados: list[Resultado] = []

    try:
        for indice, pagina in enumerate(lector.pages):
            numero, destinatario, uso_ocr = obtener_datos_pagina(
                lector, documento_ocr, indice, serie
            )
            faltantes: list[str] = []
            if not numero:
                faltantes.append("numero")
            if not destinatario:
                faltantes.append("destinatario")

            if faltantes:
                base = f"REVISAR - {pdf.stem} - pagina {indice + 1:03d}"
                estado = "Revisar: no se detecto " + " y ".join(faltantes)
            else:
                base = f"{numero} - {destinatario}"
                estado = "Correcto (OCR)" if uso_ocr else "Correcto"

                # El numero de GRE es la clave unica. Se conserva la primera
                # copia encontrada, aunque otra repeticion tenga distinto nombre.
                if numero in guias_conservadas:
                    resultados.append(
                        Resultado(
                            pdf_origen=pdf.name,
                            pagina=indice + 1,
                            numero_guia=numero,
                            destinatario=destinatario,
                            archivo_salida=guias_conservadas[numero],
                            estado="Duplicada - omitida",
                        )
                    )
                    continue

            salida = ruta_sin_colision(carpeta_salida, nombre_seguro(base))
            escritor = PdfWriter()
            escritor.add_page(pagina)
            with salida.open("wb") as archivo_salida:
                escritor.write(archivo_salida)

            if numero:
                guias_conservadas[numero] = salida.name

            resultados.append(
                Resultado(
                    pdf_origen=pdf.name,
                    pagina=indice + 1,
                    numero_guia=numero or "",
                    destinatario=destinatario or "",
                    archivo_salida=salida.name,
                    estado=estado,
                )
            )
    finally:
        documento_ocr.close()

    return resultados


def guardar_reporte(resultados: list[Resultado], carpeta: Path) -> Path:
    ruta = carpeta / "resultado_division.csv"
    with ruta.open("w", encoding="utf-8-sig", newline="") as archivo:
        escritor = csv.writer(archivo, delimiter=";")
        escritor.writerow(
            ["PDF origen", "Pagina", "Numero de guia", "Destinatario", "Archivo generado", "Estado"]
        )
        for item in resultados:
            escritor.writerow(
                [
                    item.pdf_origen,
                    item.pagina,
                    item.numero_guia,
                    item.destinatario,
                    item.archivo_salida,
                    item.estado,
                ]
            )
    return ruta


def seleccionar_con_ventanas() -> tuple[list[Path], Path] | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None

    raiz = tk.Tk()
    raiz.withdraw()
    archivos = filedialog.askopenfilenames(
        title="Selecciona uno o varios PDF con guias",
        filetypes=[("Archivos PDF", "*.pdf")],
    )
    if not archivos:
        raiz.destroy()
        return None

    carpeta = filedialog.askdirectory(title="Selecciona la carpeta de salida")
    raiz.destroy()
    if not carpeta:
        return None
    return [Path(x) for x in archivos], Path(carpeta)


def crear_argumentos() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Divide PDF de GRE y los nombra con numero de guia y destinatario."
    )
    parser.add_argument("pdf", nargs="*", type=Path, help="Uno o varios PDF de entrada")
    parser.add_argument("-o", "--salida", type=Path, help="Carpeta donde guardar las guias")
    parser.add_argument(
        "--serie",
        default=SERIE_PREDETERMINADA,
        help=f"Serie esperada de la GRE (predeterminada: {SERIE_PREDETERMINADA})",
    )
    return parser


def main() -> int:
    argumentos = crear_argumentos().parse_args()
    archivos = argumentos.pdf
    carpeta = argumentos.salida

    if not archivos:
        seleccion = seleccionar_con_ventanas()
        if not seleccion:
            print("No se seleccionaron archivos.")
            return 1
        archivos, carpeta = seleccion

    archivos = [ruta.expanduser().resolve() for ruta in archivos]
    inexistentes = [str(ruta) for ruta in archivos if not ruta.is_file()]
    if inexistentes:
        print("No se encontraron estos archivos:\n- " + "\n- ".join(inexistentes), file=sys.stderr)
        return 2

    if carpeta is None:
        carpeta = archivos[0].parent / "Guias_separadas"
    carpeta = carpeta.expanduser().resolve()
    carpeta.mkdir(parents=True, exist_ok=True)

    resultados: list[Resultado] = []
    guias_conservadas = obtener_guias_existentes(carpeta, argumentos.serie)
    try:
        for pdf in archivos:
            print(f"Procesando: {pdf.name}")
            resultados.extend(
                dividir_pdf(
                    pdf,
                    carpeta,
                    argumentos.serie,
                    guias_conservadas,
                )
            )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 3

    reporte = guardar_reporte(resultados, carpeta)
    generados = sum(item.estado.startswith("Correcto") for item in resultados)
    duplicados = sum(item.estado.startswith("Duplicada") for item in resultados)
    revisar = sum(item.estado.startswith("Revisar") for item in resultados)

    print(f"\nListo: {len(resultados)} pagina(s) procesada(s).")
    print(
        f"Archivos nuevos: {generados} | "
        f"Duplicadas omitidas: {duplicados} | Por revisar: {revisar}"
    )
    print(f"Carpeta: {carpeta}")
    print(f"Reporte: {reporte.name}")
    return 0 if revisar == 0 else 4


if __name__ == "__main__":
    raise SystemExit(main())
