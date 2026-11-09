# -*- coding: utf-8 -*-
"""Crea (la primera vez) y actualiza el unico Excel casos-niños.xlsx en la carpeta base.
- Una hoja por nivel + hoja Resumen.
- El Estado de cada fila nunca se toca despues de creada; lo escrito a mano tampoco.
- Archivos renombrados/movidos se detectan por su contenido (huella).
- Filas cuyo archivo ya no existe se marcan 'Archivo no encontrado' (no cuentan en el Resumen)."""

import json
import os
import unicodedata
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

import lector
from nucleo_ordenar import NIVELES, NOMBRE_EXCEL, es_informe, normalizar

TEMP_EXCEL = "~tmp-casos.xlsx"

ENCABEZADOS = ["Nombre", "Grado", "Edad", "Fecha", "Responsable", "Caso / Motivo", "Ficha",
               "Estado", "Archivo", "Por qué del Estado", "Archivo no encontrado", "Ruta", "Huella", "Auto"]
CAMPOS = ["nombre", "grado", "edad", "fecha", "responsable", "motivo", "ficha", "estado"]  # columnas A-H
AUTO_CAMPOS = CAMPOS + ["porque"]  # 'porque' = columna J
FILAS_MAX = 5000


class ExcelAbierto(Exception):
    pass


class ErrorCarpeta(Exception):
    pass


class ExcelInvalido(Exception):
    pass


# ------------------------------------------------------------------ utilidades
def _nfc(s):
    return unicodedata.normalize("NFC", s)


def excel_abierto(base):
    """True si Excel tiene el archivo abierto (archivo de bloqueo ~$...)."""
    objetivo = _nfc(NOMBRE_EXCEL)
    for p in Path(base).iterdir():
        n = _nfc(p.name)
        if n.startswith("~$") and (n[2:] == objetivo or n[2:] == objetivo[2:]):
            return True
    return False


def _archivos_de_nivel(base, nivel):
    carpeta = Path(base) / nivel
    if not carpeta.is_dir():
        return []
    res = []
    for p in sorted(carpeta.rglob("*")):
        try:
            if not p.is_file():
                continue
        except OSError:
            continue
        partes = p.relative_to(base).parts
        if any(x.startswith((".", "~$")) for x in partes) or p.name.lower() in ("thumbs.db", "desktop.ini", TEMP_EXCEL):
            continue
        if es_informe(p.stem):  # los "informe..." no se leen ni entran al Excel
            continue
        res.append(p)
    return res


def _texto(v):
    return "" if v is None else str(v)


def _parse_fecha(v):
    """Fecha de la celda -> ordinal (int) o None. Acepta dd/mm/aaaa, 'Mes aaaa' y fechas reales."""
    if isinstance(v, (datetime, date)):
        return (v.date() if isinstance(v, datetime) else v).toordinal()
    s = normalizar(_texto(v)).strip()
    if not s:
        return None
    import re
    m = re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", s)
    try:
        if m:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1))).toordinal()
        m = re.fullmatch(r"([a-z]+)\s+(\d{4})", s)
        if m and m.group(1) in lector.MESES:
            return date(int(m.group(2)), lector.MESES.index(m.group(1)) + 1, 1).toordinal()
    except ValueError:
        return None
    return None


# ------------------------------------------------------------------ lectura del Excel existente
def _cargar_filas(ruta):
    try:
        wb = load_workbook(ruta)
    except Exception as e:
        raise ExcelInvalido(f"No se pudo abrir el Excel existente: {e}")
    filas = []
    cab_extra = {}  # {nivel: {columna: (valor, formato, relleno)}} columnas que la persona agrego a la derecha
    n = len(ENCABEZADOS)
    for nivel in NIVELES:
        if nivel not in wb.sheetnames:
            raise ExcelInvalido(f"Al Excel le falta la hoja «{nivel}». No se modificó nada.")
        ws = wb[nivel]
        cab = [c.value for c in ws[1]][:n]
        if cab != ENCABEZADOS:
            raise ExcelInvalido(f"La hoja «{nivel}» no tiene el formato esperado. No se modificó nada.")
        cab_extra[nivel] = {c.column: _estilo(c) for c in ws[1][n:] if c.value not in (None, "")}
        for celdas in ws.iter_rows(min_row=2):
            r = [c.value for c in celdas[:n]]
            r += [None] * (n - len(r))
            extra = {c.column: _estilo(c) for c in celdas[n:] if c.value not in (None, "")}
            if all(v in (None, "") for v in r) and not extra:
                continue
            fila = {c: r[i] for i, c in enumerate(CAMPOS)}
            fila.update(archivo=_texto(r[8]), porque=r[9], noenc=_texto(r[10]),
                        ruta=_nfc(_texto(r[11])) or None, huella=_texto(r[12]) or None, nivel=nivel,
                        extra=extra,
                        estilo={c.column: _estilo(c)[1:] for c in celdas[:n]
                                if c.number_format != "General" or c.fill.fill_type == "solid"})
            try:
                fila["auto"] = json.loads(_texto(r[13])) if r[13] else {}
            except ValueError:
                fila["auto"] = {}
            filas.append(fila)
    return filas, cab_extra


def _estilo(c):
    """(valor, formato de numero, relleno solido o None) de una celda."""
    relleno = copy(c.fill) if c.fill is not None and c.fill.fill_type == "solid" else None
    return (c.value, c.number_format, relleno)


def _aplicar_auto(fila, nuevo, nueva):
    """Escribe valores automaticos sin pisar lo editado a mano. El Estado nunca se toca en filas existentes."""
    if nueva:
        for c in AUTO_CAMPOS:
            fila[c] = nuevo[c]
    else:
        auto_viejo = fila.get("auto") or {}
        for c in AUTO_CAMPOS:
            if c == "estado" and _texto(fila.get(c)).strip():
                continue  # el Estado con valor nunca se toca; vacio (no se pudo leer antes) si se completa
            actual = _texto(fila.get(c))
            if c in auto_viejo:
                sin_tocar = actual == _texto(auto_viejo[c])
            else:
                sin_tocar = actual == ""  # fila sin historial: solo se completan vacios
            if sin_tocar:
                fila[c] = nuevo[c]
    fila["auto"] = {c: nuevo[c] for c in AUTO_CAMPOS}


# ------------------------------------------------------------------ escritura
def _orden(fila):
    if fila.get("noenc") == "Sí":
        rank = 3
    else:
        rank = {"No derivado": 0, "Derivado": 1}.get(_texto(fila.get("estado")).strip(), 2)
    f = _parse_fecha(fila.get("fecha"))
    return (rank, 0 if f is not None else 1, -(f or 0), normalizar(_texto(fila.get("nombre"))))


def _limpio(v):
    """Quita caracteres de control que openpyxl no acepta (IllegalCharacterError)."""
    return ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v


def _escribir(filas, destino, origen=None, cab_extra=None):
    """Si hay un Excel previo (origen), se conservan las hojas extra del usuario: solo se recrean
    inicial/primaria/secundaria/Resumen (al inicio, en ese orden)."""
    cab_fill = PatternFill("solid", start_color="D9E2F3", end_color="D9E2F3")
    rojo = PatternFill("solid", start_color="F4B6B6", end_color="F4B6B6")
    verde = PatternFill("solid", start_color="B7E1B5", end_color="B7E1B5")
    gris = PatternFill("solid", start_color="E7E6E6", end_color="E7E6E6")
    if origen is not None:
        wb = load_workbook(origen)
        for n in NIVELES + ["Resumen"]:
            if n in wb.sheetnames:
                del wb[n]
    else:
        wb = Workbook()
        wb.remove(wb.active)
    for idx, nivel in enumerate(NIVELES):
        ws = wb.create_sheet(nivel, idx)
        ws.append(ENCABEZADOS)
        for c in ws[1]:
            c.font, c.fill = Font(bold=True), cab_fill
        for f in sorted((x for x in filas if x["nivel"] == nivel), key=_orden):
            ed = f.get("edad")
            ed = int(ed) if _texto(ed).isdigit() else (ed if ed not in (None, "") else None)
            ws.append([_limpio(v) for v in (
                f.get("nombre"), f.get("grado"), ed, f.get("fecha"), f.get("responsable"),
                f.get("motivo"), f.get("ficha"), f.get("estado"), f.get("archivo"),
                f.get("porque") or None, f.get("noenc") or None, f.get("ruta"), f.get("huella"),
                json.dumps(f.get("auto") or {}, ensure_ascii=False))])
            for c in ws[ws.max_row]:  # texto que empieza con = + - @ se escribe como texto, no como formula
                if isinstance(c.value, str) and c.value[:1] in "=+-@":
                    c.data_type = "s"
            r = ws.max_row  # lo que la persona agrego: formatos, colores y columnas extra a la derecha
            for col, (fmt, relleno) in (f.get("estilo") or {}).items():
                if fmt != "General":
                    ws.cell(r, col).number_format = fmt
                if relleno is not None:
                    ws.cell(r, col).fill = copy(relleno)
            for col, (v, fmt, relleno) in (f.get("extra") or {}).items():
                cel = ws.cell(r, col, _limpio(v))
                cel.number_format = fmt
                if relleno is not None:
                    cel.fill = copy(relleno)
        for col, (v, fmt, relleno) in (cab_extra or {}).get(nivel, {}).items():
            cel = ws.cell(1, col, _limpio(v))
            cel.number_format = fmt
            if relleno is not None:
                cel.fill = copy(relleno)
        for fila in ws.iter_rows(min_row=2):
            for c in fila:
                c.alignment = Alignment(wrap_text=(c.column == 6), vertical="top")
        for col, w in zip("ABCDEFGHIJK", (36, 9, 7, 16, 22, 80, 8, 15, 44, 60, 14)):
            ws.column_dimensions[col].width = w
        for col in "LMN":
            ws.column_dimensions[col].hidden = True
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:K{max(ws.max_row, 2)}"
        dv = DataValidation(type="list", formula1='"Derivado,No derivado"', allow_blank=True)
        dv.showErrorMessage = True
        dv.error, dv.errorTitle = "Elige Derivado o No derivado", "Valor no válido"
        ws.add_data_validation(dv)
        dv.add(f"H2:H{FILAS_MAX}")
        ws.conditional_formatting.add(f"H2:H{FILAS_MAX}", CellIsRule(
            operator="equal", formula=['"Derivado"'], fill=rojo, font=Font(color="9C0006", bold=True)))
        ws.conditional_formatting.add(f"H2:H{FILAS_MAX}", CellIsRule(
            operator="equal", formula=['"No derivado"'], fill=verde, font=Font(color="1E5E20", bold=True)))
        ws.conditional_formatting.add(f"A2:G{FILAS_MAX}", FormulaRule(
            formula=['$K2="Sí"'], fill=gris, font=Font(color="7F7F7F", italic=True)))
    rs = wb.create_sheet("Resumen", len(NIVELES))
    rs.append(["Nivel", "Casos", "Derivados", "No derivados", "Archivo no encontrado (no cuenta)"])
    for c in rs[1]:
        c.font, c.fill = Font(bold=True), cab_fill
    for i, nivel in enumerate(NIVELES, start=2):
        nf = f'{nivel}!$K$2:$K${FILAS_MAX}'
        rs.append([nivel.capitalize(),
                   f'=COUNTIFS({nivel}!$A$2:$A${FILAS_MAX},"<>",{nf},"<>Sí")',
                   f'=COUNTIFS({nivel}!$H$2:$H${FILAS_MAX},"Derivado",{nf},"<>Sí")',
                   f'=COUNTIFS({nivel}!$H$2:$H${FILAS_MAX},"No derivado",{nf},"<>Sí")',
                   f'=COUNTIF({nf},"Sí")'])
    rs.append(["TOTAL"] + [f"=SUM({c}2:{c}4)" for c in "BCDE"])
    borde = Border(top=Side(style="thin"))
    for c in rs[5]:
        c.font, c.border = Font(bold=True), borde
    for fila in rs.iter_rows(min_row=2, min_col=2, max_col=5):
        for c in fila:
            c.alignment = Alignment(horizontal="center")
    for r in range(2, 6):
        rs[f"C{r}"].fill, rs[f"C{r}"].font = rojo, Font(color="9C0006", bold=r == 5)
        rs[f"D{r}"].fill, rs[f"D{r}"].font = verde, Font(color="1E5E20", bold=r == 5)
    for col, w in zip("ABCDE", (18, 12, 14, 16, 34)):
        rs.column_dimensions[col].width = w
    wb.active = 0
    for s in wb:
        s.sheet_view.tabSelected = (s is wb.active)
    wb.save(destino)


def _guardar(filas, ruta, cab_extra=None):
    """Escribe en un temporal junto al Excel y lo reemplaza al final (evita dejarlo a medias)."""
    tmp = ruta.with_name(TEMP_EXCEL)
    try:
        _escribir(filas, tmp, ruta if ruta.exists() else None, cab_extra)
        try:
            os.replace(tmp, ruta)
        except PermissionError:
            raise ExcelAbierto()
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


# ------------------------------------------------------------------ proceso principal
def actualizar(base, preguntar_manual, progreso=None):
    """Crea o actualiza el Excel.
    preguntar_manual(nombre, ruta) -> 'continuar' | 'todos' | 'abrir'  (la app abre el archivo si 'abrir').
    Devuelve dict con el resultado."""
    base = Path(base)
    ruta_xlsx = base / NOMBRE_EXCEL
    if excel_abierto(base):
        raise ExcelAbierto()
    existe = ruta_xlsx.exists()
    escaneo = {n: _archivos_de_nivel(base, n) for n in NIVELES}
    total = sum(len(v) for v in escaneo.values())
    hay_carpetas = any((base / n).is_dir() for n in NIVELES)
    if not hay_carpetas:
        raise ErrorCarpeta("En esta carpeta no están las carpetas inicial, primaria o secundaria.\n"
                           "¿Elegiste la carpeta correcta? Si no están, primero usa «Ordenar archivos».")
    filas, cab_extra = _cargar_filas(ruta_xlsx) if existe else ([], {})
    res = {"creado": not existe, "nuevos": 0, "actualizados": 0, "movidos_hoja": 0,
           "no_encontrados": 0, "huerfanos": [], "detenido": False, "total": total, "errores": []}

    rel = lambda p: _nfc(p.relative_to(base).as_posix())
    en_escaneo = {rel(p) for v in escaneo.values() for p in v}
    por_ruta = {f["ruta"]: f for f in filas if f["ruta"]}
    huerfanas_previas = [f for f in filas if f["ruta"] and f["ruta"] not in en_escaneo and f["huella"]]
    reclamadas = set()
    responder_todos = False
    hecho = 0

    for nivel in NIVELES:
        for p in escaneo[nivel]:
            hecho += 1
            if progreso:
                progreso(hecho, total, p.name)
            try:
                h = lector.huella(p)
                r = rel(p)
                fila, nueva, cand = por_ruta.get(r), False, None
                if fila is not None:
                    fila["noenc"] = ""
                    if fila["huella"] == h:
                        fila["archivo"] = p.name
                        continue  # sin cambios: no se relee
                else:
                    for c in huerfanas_previas:  # renombrado o movido: misma huella
                        if c["huella"] == h and id(c) not in reclamadas:
                            cand = c
                            break
                # primero se lee y extrae; la fila solo cambia (y la huella solo se guarda) si todo salio bien
                estado_lectura, lineas = lector.leer(p)
                nuevo = lector.extraer(p, estado_lectura, lineas)
                if fila is None:
                    if cand is not None:
                        fila = cand
                        reclamadas.add(id(cand))
                        por_ruta.pop(fila["ruta"], None)
                        if fila["nivel"] != nivel:
                            res["movidos_hoja"] += 1
                        fila.update(ruta=r, nivel=nivel, noenc="")
                    else:
                        fila = {"nivel": nivel, "ruta": r, "noenc": "", "auto": {}, "huella": None}
                        nueva = True
                    por_ruta[r] = fila
                fila["archivo"] = p.name
                _aplicar_auto(fila, nuevo, nueva)
                fila["huella"] = h
                if nueva:
                    filas.append(fila)
                res["nuevos" if nueva else "actualizados"] += 1
                if estado_lectura == lector.MANUAL:
                    res["huerfanos"].append(p.name)
                    if not responder_todos:
                        resp = preguntar_manual(p.name, p)
                        if resp == "todos":
                            responder_todos = True
                        elif resp == "abrir":
                            res["detenido"] = True
                            break
            except ExcelAbierto:
                raise
            except Exception as e:
                res["errores"].append(f"{p.name}: {e}")
        if res["detenido"]:
            break

    if not res["detenido"]:
        for f in filas:
            if f["ruta"] and f["ruta"] not in en_escaneo and id(f) not in reclamadas:
                f["noenc"] = "Sí"
    res["no_encontrados"] = sum(1 for f in filas if f.get("noenc") == "Sí")
    _guardar(filas, ruta_xlsx, cab_extra)
    res["ruta"] = ruta_xlsx
    return res
