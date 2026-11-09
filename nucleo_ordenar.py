# -*- coding: utf-8 -*-
"""Ordena los archivos sueltos de la carpeta base en inicial / primaria / secundaria.
Los que no se pueden clasificar van a REVISAR_A_MANO.
Solo MUEVE archivos: nunca borra ni sobrescribe. Las carpetas se crean solo si hacen falta."""

import codecs
import html
import re
import shutil
import unicodedata
import zipfile
from pathlib import Path

NIVELES = ["inicial", "primaria", "secundaria"]
REVISAR = "REVISAR_A_MANO"
NOMBRE_EXCEL = "casos-niños.xlsx"
CLAVES = {
    "inicial": {"inicial"},
    "primaria": {"primaria", "prim"},
    "secundaria": {"secundaria", "sec"},
}


def normalizar(texto):
    t = unicodedata.normalize("NFD", texto.lower())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


def nivel_por_titulo(stem):
    """Devuelve 'inicial'/'primaria'/'secundaria', 'VARIOS' si hay mas de uno, o None."""
    pal = {p for p in re.split(r"[^a-z0-9]+", normalizar(stem)) if p}
    if "6to" in pal:  # sexto grado es siempre primaria
        return "primaria"
    coinc = [n for n, claves in CLAVES.items() if pal & claves]
    if len(coinc) == 1:
        return coinc[0]
    return "VARIOS" if len(coinc) > 1 else None


def es_informe(stem):
    """True si el titulo (sin extension, sin tildes/mayusculas) empieza con 'informe'."""
    return normalizar(stem).lstrip(" _-.([").startswith("informe")


def decodificar(crudo):
    """Bytes -> texto. Soporta BOM UTF-8 y UTF-16; si no, UTF-8 y luego latin-1."""
    if crudo.startswith(codecs.BOM_UTF8):
        return crudo.decode("utf-8-sig", "replace")
    if crudo.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return crudo.decode("utf-16", "replace")
    try:
        return crudo.decode("utf-8")
    except UnicodeDecodeError:
        return crudo.decode("latin-1")


_ADJ_ANTES = r"(?:entrevista|evaluacion|sesion|diagnostico|reunion|consulta|atencion|fase|etapa|observacion|primera|valoracion|contacto|cita|visita|intervencion|caso|informe|derivacion|nota|llamada)"
_INICIAL_ADJ = re.compile(r"\b" + _ADJ_ANTES + r"\s+inicial\b|\binicial\s+(?:con|de|del|a|al|y|en|por|para)\b")
_INICIAL_NIVEL = re.compile(r"\b(?:nivel|educacion)\s*:?\s*inicial\b")
_EDAD_INICIAL = re.compile(r"\b[345]\s*anos\b")


def _niveles_en_lineas(lineas):
    """Todos los niveles distintos que aparecen. 'inicial' como adjetivo no cuenta; como nivel solo
    en 'nivel/educacion inicial' o suelto en las primeras 8 lineas junto a 3/4/5 anos."""
    norm = [normalizar(l) for l in lineas]
    texto = "\n".join(norm)
    niveles = {m.group(1) for m in re.finditer(r"\b(primaria|secundaria)\b", texto)}
    if _INICIAL_NIVEL.search(texto):
        niveles.add("inicial")
    else:
        cab = norm[:8]
        if any(_EDAD_INICIAL.search(l) for l in cab):
            for l in cab:
                if (re.search(r"\binicial\b", l) and len(l.split()) <= 6
                        and not _INICIAL_ADJ.search(l)):
                    niveles.add("inicial")
    return niveles


def _lineas_xml(zf, nombre):
    x = zf.read(nombre).decode("utf-8", "ignore")
    x = re.sub(r"</w:p>", "\n", x)
    x = re.sub(r"<[^>]+>", "", x)
    return [l.strip() for l in html.unescape(x).splitlines() if l.strip()]


def nivel_por_contenido(ruta):
    """Solo para archivos sin palabra clave en el titulo. Devuelve el nivel, 'VARIOS' si aparece mas
    de uno, o None si no se puede saber."""
    ext = ruta.suffix.lower()
    try:
        if ext in (".docx", ".dotx", ".docm"):
            with zipfile.ZipFile(ruta) as zf:
                nombres = zf.namelist()
                lineas = _lineas_xml(zf, "word/document.xml") if "word/document.xml" in nombres else []
                for n in sorted(n for n in nombres if re.match(r"word/(header|footer)\d*\.xml", n)):
                    lineas += _lineas_xml(zf, n)
        elif ext == ".pdf":
            from pypdf import PdfReader
            lector = PdfReader(str(ruta))
            if lector.is_encrypted:
                return None
            lineas = "\n".join((pg.extract_text() or "") for pg in lector.pages[:3]).splitlines()
            lineas = [l.strip() for l in lineas if l.strip()]
        elif ext in (".txt", ".md", ".csv"):
            with open(ruta, "rb") as f:
                crudo = f.read(2_000_000)
            lineas = [l.strip() for l in decodificar(crudo).splitlines() if l.strip()][:400]
        else:
            return None
    except Exception:
        return None
    niveles = _niveles_en_lineas(lineas)
    if len(niveles) == 1:
        return next(iter(niveles))
    return "VARIOS" if niveles else None


def es_ignorado(p):
    n = p.name
    return (n.startswith((".", "~$")) or n.lower() in ("thumbs.db", "desktop.ini")
            or unicodedata.normalize("NFC", n) in (NOMBRE_EXCEL, "~tmp-casos.xlsx"))


def planificar(base):
    """Devuelve lista [(ruta, destino)] con destino = nivel o REVISAR. Solo archivos sueltos."""
    plan = []
    for p in sorted(Path(base).iterdir()):
        try:
            if not p.is_file() or es_ignorado(p):
                continue
        except OSError:
            continue
        nivel = nivel_por_titulo(p.stem)
        if es_informe(p.stem):  # informes: no se lee su contenido; sin nivel en el titulo se dejan donde estan
            if nivel in NIVELES:
                plan.append((p, nivel))
            continue
        if nivel is None:
            nivel = nivel_por_contenido(p)
        plan.append((p, nivel if nivel in NIVELES else REVISAR))  # None o VARIOS -> REVISAR
    return plan


def _destino_libre(carpeta, nombre):
    d = carpeta / nombre
    if not d.exists():
        return d
    stem, suf = Path(nombre).stem, Path(nombre).suffix
    n = 1
    while (carpeta / f"{stem}_({n}){suf}").exists():
        n += 1
    return carpeta / f"{stem}_({n}){suf}"


def mover(base, plan):
    """Mueve los archivos. Devuelve ({destino: [nombres]}, [errores])."""
    base = Path(base)
    movidos = {n: [] for n in NIVELES + [REVISAR]}
    errores = []
    for origen, destino_nombre in plan:
        try:
            carpeta = base / destino_nombre
            carpeta.mkdir(exist_ok=True)  # solo se crea cuando hay algo que mover
            destino = _destino_libre(carpeta, origen.name)
            shutil.move(str(origen), str(destino))
            if origen.exists() or not destino.exists():
                raise RuntimeError("no se pudo verificar el movimiento")
            movidos[destino_nombre].append(destino.name)
        except Exception as e:
            errores.append(f"{origen.name}: {e}")
    return movidos, errores
