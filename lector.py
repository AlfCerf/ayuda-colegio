# -*- coding: utf-8 -*-
"""Lee un documento y extrae: nombre, grado, edad, fecha, responsable, motivo, ficha, derivado.
Regla de oro: si un dato no aparece, queda vacio. Nunca se inventa."""

import hashlib
import re
import zipfile
from datetime import date
from pathlib import Path

from nucleo_ordenar import decodificar, normalizar

OK, VACIO, MANUAL, NO_LEIDO = "ok", "vacio", "manual", "no_leido"
EXT_TEXTO = {".txt", ".md", ".csv", ".rtf"}
EXT_IMAGEN = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".heic", ".webp"}
EXT_MANUAL = EXT_IMAGEN | {".doc", ".pages"}  # un humano si los puede leer; la app no

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MES_ABREV = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7,
             "ago": 8, "set": 9, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12}


def huella(ruta):
    h = hashlib.sha1()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _xml_a_lineas(x):
    x = re.sub(r"</(w:p|text:p|text:h)>", "\n", x)
    x = re.sub(r"</w:tc>", " | ", x)
    x = re.sub(r"<w:tab/>", " ", x)
    x = re.sub(r"<[^>]+>", "", x)
    x = (x.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
          .replace("&quot;", '"').replace("&apos;", "'"))
    return x


def _lineas(texto):
    return [re.sub(r"[ \t]+", " ", l).strip() for l in texto.splitlines() if l.strip()]


def leer(ruta):
    """Devuelve (estado, lineas). estado: ok / vacio / manual / no_leido."""
    ext = Path(ruta).suffix.lower()
    try:
        if ext in (".docx", ".dotx", ".docm"):
            with zipfile.ZipFile(ruta) as z:
                texto = _xml_a_lineas(z.read("word/document.xml").decode("utf-8", "ignore"))
        elif ext == ".odt":
            with zipfile.ZipFile(ruta) as z:
                texto = _xml_a_lineas(z.read("content.xml").decode("utf-8", "ignore"))
        elif ext == ".pdf":
            from pypdf import PdfReader
            lector = PdfReader(str(ruta))
            if lector.is_encrypted:
                try:
                    if not lector.decrypt(""):
                        return MANUAL, []
                except Exception:
                    return MANUAL, []
            texto = "\n".join((pg.extract_text() or "") for pg in lector.pages)
            if len(re.sub(r"\s+", "", texto)) < 25:  # sin texto = escaneado
                return MANUAL, []
        elif ext in EXT_TEXTO:
            texto = decodificar(Path(ruta).read_bytes())  # soporta BOM UTF-8 / UTF-16
            if ext == ".rtf":
                texto = re.sub(r"\\par[d]?", "\n", texto)
                texto = re.sub(r"\\'[0-9a-f]{2}|\\[a-z]+-?\d* ?|[{}]", "", texto)
        elif ext in EXT_MANUAL:
            return MANUAL, []
        else:
            return NO_LEIDO, []
    except Exception:
        return MANUAL, []  # dañado, protegido o ilegible: que lo lea una persona
    lineas = _lineas(texto)
    return (OK, lineas) if lineas else (VACIO, [])


# ------------------------------------------------------------------ campos
TEMAS = [
    (r"violencia|golpe|agresi|maltrato", "Violencia / agresión"),
    (r"bullying|bulling|acoso|hostiga", "Bullying / acoso"),
    (r"inasisten|no quiere (venir|ir)|faltas|ausent|deserci", "Inasistencias"),
    (r"ansiedad|nervios|angustia|ataque de", "Ansiedad"),
    (r"depresi|triste|llora|llanto|animo bajo", "Estado de ánimo"),
    (r"autolesi|se corta|cortes|suicid|ideacion", "Autolesión / riesgo"),
    (r"duelo|falleci|murio|fallecimiento", "Duelo"),
    (r"conducta|desafiante|agresiv|berrinche|rebeld|no obedece|groseria", "Conducta"),
    (r"no escribe|no lee|rendimiento|notas bajas|desaprob|retencion|dificultad(es)? (de|en) (aprendizaje|lectura|escritura)|no avanza", "Aprendizaje / rendimiento"),
    (r"separad|divorcio|abandon|violencia familiar|padres? (se )?separ|vive solo", "Situación familiar"),
    (r"enamorad|pareja|novi[oa]|ex ?(pareja|novi)", "Relaciones / pareja"),
    (r"sustancia|alcohol|drog|fuma", "Consumo de sustancias"),
    (r"autoestima|inseguridad|se siente (mal|solo|sola)", "Autoestima"),
]
CUE = re.compile(r"motivo|deriv|solicita|refiere|preocup|queja|problema|dificultad|conducta|inasisten|"
                 r"no quiere|llora|agred|golpe|bullying|acoso|ansied|trist|no (escribe|lee|avanza)|"
                 r"violencia|separad")
# ---- derivacion a algo EXTERNO (reglas de contexto, todo en local) ----
# Entidades siempre externas
_EXT_FIJAS = (r"centro de salud|\bposta|\bhospital|psiquiatr\w*|psicolog\w* clinic\w*|(?:a|en|una|la) clinica|essalud|minsa|\bcem\b|demuna|"
              r"fiscal\w*|policia\b|\bpnp|comisaria|neurolog(?:o|a|os|as|ia)\b|pediatr(?:a|as|ia)\b|"
              r"terapia (?:de lenguaje|ocupacional|fisica|especializada)|\bcai\b|juzgado|"
              r"(?<!certificado )medic[oa]s?\b|especialista\w*|\bsis\b|"
              r"seguro integral|seguro social|seguro de salud|seguro medico|seguro privado|"
              r"\bcsmc\b|puesto de salud|establecimiento de salud|\bugel\b|psicolog\w* (?:particular|privad\w+)|"
              r"neuropsicolog\w+|centro (?:de )?emergencia mujer|poder judicial|defensoria|\binabif\b|\bmimp\b|"
              r"(?:institucion(?:es)?|centro|psicolog[oa]s?|profesional(?:es)?|especialista\w*|derivacion|"
              r"atencion|seguimiento) extern[oa]s?")
# "seguro" como entidad: solo "al/del/el seguro" (no el adjetivo "no esta seguro")
_EXT_SEGURO = r"(?:al|del|en el|el) seguro\b"
# "institucion" solo si se dice que es externa
_EXT_INSTITUCION = (r"institucion(?:es)? (?:externa|externas|de salud|especializada|especializadas|"
                    r"publica|publicas|privada|privadas|medica|de apoyo)")
EXTERNA = "|".join([_EXT_FIJAS, _EXT_SEGURO, _EXT_INSTITUCION])
ACCION = (r"deriv\w*|referid[oa]s? (?:a|al|con|para)\b|referenci\w+|remit\w+|transfer\w+|recom(?:e|ie)nd\w+|sugier\w+|sugerid\w+|llevar\w*|llevo|llevaron|"
          r"acud\w+|solicit\w+|seguimiento|canaliz\w+|pas[oó] a|enviar\w*|envi[oó]|requiere|necesita|"
          r"coordin(?!ador)\w+|vincul\w+|articul\w+|atencion (?:en|por|con)|atendid[oa]s? (?:en|por)")
DERIV = re.compile(r"\b(" + ACCION + r")\b[^.\n]{0,90}?(?:\b|^)(" + EXTERNA + r")|"
                   r"(" + EXTERNA + r")[^.\n]{0,60}?\b(" + ACCION + r")\b|"
                   r"\b(?:derivacion|atencion|seguimiento) extern[oa]s?\b")
# Destinos INTERNOS (el propio colegio / Psicologia): se borran antes de buscar
_INTERNO = re.compile(r"psicolog\w* (?:del|de la|de) (?:colegio|institucion|ie)|area de psicologia|"
                      r"departamento de psicologia|(?:la|nuestra|esta|dicha) institucion(?: educativa)?|"
                      r"institucion educativa|el colegio|tutor\w*|subdireccion|direccion|coordinacion|"
                      r"docente\w*|auxiliar\w*|ficha de derivacion|quien deriva|promotor pnp|derivacion de casos|"
                      r"motivo por el que se solicita derivacion|solicitud de derivacion|"
                      r"(?:especialista|medic[oa]|psicolog\w*)s? (?:del|de la|de el) (?:colegio|institucion)")
# Frases sobre un familiar (no sobre el estudiante): "la mama acudio al hospital por su embarazo"
_FAMILIAR = re.compile(r"\b(?:mama|papa|madre|padre|abuela|abuelo|tia|tio|hermana|hermano|familiar\w*|apoderad\w+)\b")
_ESTUDIANTE = re.compile(r"\b(?:estudiante|alumn\w+|nino|nina|hij[oa]|menor|deriv\w*|recom(?:e|ie)nd\w*|solicit\w*|sugier\w*|"
                         r"sugerid\w*|seguimiento|remit\w*|referid\w*)\b")
# Frase con verbo de derivar + destino interno (para la razon de "No derivado")
_INTERNO_DEST = re.compile(r"(?:deriv\w*|refer\w*|remit\w*|pas[oó]|envi\w*|coordin\w*|informo|informad\w*)[^.\n]{0,60}?"
                           r"(?:psicolog\w*|tutor\w*|direccion|subdireccion|coordinacion|colegio|institucion)")
# Derivaciones que VIENEN de afuera hacia Psicologia (promotor PNP) y giros que no son destino
_ORIGEN = re.compile(r"deriv\w* (?:realizada |hecha )?(?:por|del|de la|de el)(?: parte)?(?: de)?(?: el| la)? ?(?:\w+ )?"
                     r"(?:promotor\w*|pnp|policia\w*|oficial)(?: de(?: la)?)?(?: policia(?: nacional)?(?: del peru)?| pnp)?|deriv\w* pnp|personal de (?:la )?policia\w*|derivacion oportuna|"
                     r"promotor\w*(?: de(?: la)?)? (?:policia(?: nacional)?(?: del peru)?|pnp)|"
                     r"derivad[oa]s? de(?:l)? \w+|acus\w+[^.\n]*")
_COORD = re.compile(r"coordin(?!ador)\w+|articul\w+|vincul\w+")
_FUERTE = re.compile(r"\b(?:deriv\w*|referid\w*|recom(?:e|ie)nd\w*|solicit\w*|acud\w*|llevar\w*|llevo|llevaron|enviar\w*|envi[oó])")
_GENERAL = re.compile(r"trabajo (?:articulado|conjunto|coordinado)|articulad|prevencion")
_ALIADO = re.compile(r"aliad\w*[^.\n]{0,40}(?:seguimiento|atencion|atend\w+|deriv\w+)|"
                     r"(?:seguimiento|atencion|atend\w+|deriv\w+)[^.\n]{0,40}aliad\w*")
# Negaciones: "no se derivo", "no requiere derivacion", "sin necesidad de derivar"
_NEGACION = re.compile(r"\bno (?:se )?(?:requiere|necesita|amerita|fue necesari\w+|es necesari\w+|hay necesidad|"
                       r"deriv\w+|refir\w+|derivar\w*)|\bsin (?:necesidad de )?deriv\w+|\bno existe (?:necesidad|derivacion)|\bno se evidencia\w*[^.\n]{0,90}(?:requier\w+|deriv\w+)")


# Etiqueta de escala: "Severo (derivacion externa)" = derivado, y la etiqueta es el por que
_ETIQUETA = re.compile(r"\b(?:leve|moderad[oa]|severo|grave)\s*\(?\s*derivaci(?:o|\u00f3|o\u0301)n\s+externa\s*\)?", re.I)
# Frase en PRIMERA PERSONA del estudiante sobre atencion externa: no es derivacion
_PRIMERA = re.compile(r"\byo\s+(?:\w+\s+){0,2}?(?:llevo|voy|estoy|tomo|asisto|recibo|tengo)\b|\bme\s+(?:\w+\s+){0,2}?(?:llev\w+|mand\w+|envi\w+|trae\w*|atiend\w+|inscrib\w+)")
_ATENCION_EXT = re.compile(EXTERNA + r"|psicolog\w*|terapia\w*|tratamiento\w*|medicacion|pastillas")


def _primera_persona(n):
    """n = frase normalizada. True si el propio estudiante cuenta que recibe atencion externa."""
    n = _INTERNO.sub(" ", n)
    return bool(_PRIMERA.search(n) and _ATENCION_EXT.search(n) and not _NEGACION.search(n))


def _recorte(frase):
    f = re.sub(r"\s+", " ", frase).strip(" |")
    return (f[:160] + "…") if len(f) > 160 else f


def evidencia_derivacion(texto):
    """Devuelve la frase que indica derivacion a algo externo, o '' si no hay. Reglas de contexto:
    ignora destinos internos (colegio/Psicologia), negaciones y palabras ambiguas ('seguro', 'institucion')."""
    for frase in re.split(r"(?<=[.;!?])\s+|\n", texto):
        n = normalizar(frase)
        m = _ETIQUETA.search(frase)
        if m and not re.search(r"\bno (?:se )?evidencia", normalizar(frase[:m.start()])):
            return re.sub(r"\s+", " ", m.group(0)).strip()
        if _NEGACION.search(n) or _primera_persona(n):
            continue
        n = _ORIGEN.sub(" ", _INTERNO.sub(" ", n))
        if _FAMILIAR.search(n) and not _ESTUDIANTE.search(n):
            continue  # habla de un familiar, no del estudiante
        if _COORD.search(n) and not _FUERTE.search(n):  # "coordinar con profesionales externos" no es derivar
            n = re.sub(r"\bexterno\w*|\bexterna\w*", " ", n)
        # "aliado/a" solo cuenta si seguimiento, atencion o derivacion estan junto a el
        if DERIV.search(n) or (_ALIADO.search(n) and not _GENERAL.search(n)):
            f = re.sub(r"\s+", " ", frase).strip(" |")
            return (f[:160] + "…") if len(f) > 160 else f
    return ""


def estado_y_porque(texto):
    """Devuelve (estado, porque). 'porque' explica el Estado en ambos casos."""
    ev = evidencia_derivacion(texto)
    if ev:
        return "Derivado", ev
    neg = fp = ""
    interno = False
    for frase in re.split(r"(?<=[.;!?])\s+|\n", texto):
        n = normalizar(frase)
        if not fp and _primera_persona(n):
            fp = _recorte(frase)
        if not neg and _NEGACION.search(n) and re.search(r"deriv|refer|extern|especializ", n):
            f = re.sub(r"\s+", " ", frase).strip(" |")
            neg = (f[:160] + "…") if len(f) > 160 else f
        if _INTERNO_DEST.search(n):
            interno = True
    if fp:  # lo cuenta el propio estudiante: no es una derivacion
        return "No derivado", fp
    if neg:
        return "No derivado", neg
    if interno:
        return "No derivado", "Solo se menciona atención interna (colegio/Psicología)"
    return "No derivado", "No se menciona derivación a algo externo"


def _solo_subrayado(s):
    return not re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ0-9]", s)


def _valor(patron, texto):
    m = re.search(patron, texto, re.I)
    if not m:
        return ""
    v = m.group(1).strip(" _:.-|\t")
    return "" if _solo_subrayado(v) else re.sub(r"\s+", " ", v)


def _nombre_valido(l):
    return 3 <= len(l) <= 70 and sum(c.isdigit() for c in l) <= 2 and len(l.split()) >= 2


def _mes_token(tok):
    """Mes (1-12) si el token es un mes completo (cualquier caso) o una abreviatura en MAYUSCULAS
    (AGO, SET, JUN...). 'Mar' o 'del' en un nombre no cuentan."""
    limpio = tok.strip(".,-_()")
    t = normalizar(limpio)
    if t in MESES:
        return MESES.index(t) + 1
    if t in MES_ABREV and limpio.isupper():
        return MES_ABREV[t]
    return None


def _limpiar_titulo(stem):
    """Nombre a partir del titulo: se corta en el primer mes, grado o nivel."""
    out = []
    for tok in stem.split():
        t = normalizar(tok).strip(".,-_()")
        if (_mes_token(tok) or t in {"prim", "primaria", "sec", "secundaria", "inicial", "empty"}
                or re.fullmatch(r"[1-6]\s?(ro|do|er|to|mo|vo|ero)", t)):
            break
        out.append(tok)
    return " ".join(out).strip(" -_.")


def mes_en_titulo(stem):
    """Ultimo mes del titulo (1-12) o None: mes completo, o abreviatura solo si esta en MAYUSCULAS."""
    mes = None
    for tok in re.split(r"[^A-Za-zÁÉÍÓÚáéíóúñÑ]+", stem):
        m = _mes_token(tok)
        if m:
            mes = m
    return mes


def _fecha_en_texto(lineas, anio):
    """Primera fecha COMPLETA que sea del caso: 'Fecha: ...' o con el año actual.
    Se ignoran fechas de nacimiento (otro año). Devuelve (dia, mes, anio) o None."""
    for l in lineas[:60]:
        n = normalizar(l)
        for m in re.finditer(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", n):
            d, mm, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
            a = a + 2000 if a < 100 else a
            if a == anio and 1 <= d <= 31 and 1 <= mm <= 12:
                return d, mm, a
        m = re.search(r"fecha\s*[:.]?\s*(\d{1,2})\s*(?:de)?\s*([a-z]+)", n)
        if m:
            mes = MES_ABREV.get(m.group(2)) or (MESES.index(m.group(2)) + 1 if m.group(2) in MESES else None)
            if mes and 1 <= int(m.group(1)) <= 31:
                return int(m.group(1)), mes, anio
    return None


def formatear_fecha(dia, mes, anio):
    return f"{dia:02d}/{mes:02d}/{anio}" if dia else f"{MESES[mes - 1].capitalize()} {anio}"


def _grado(lineas, stem):
    pat = re.compile(r"^\s*([1-6])\s?(ro|do|er|to|mo|vo|ero|°|º)\s*(?:grado)?\s*([A-Ea-e])?\b", re.I)
    for l in lineas[:8]:
        m = pat.match(l)
        if m and len(l) <= 40:
            sec = f" {m.group(3).upper()}" if m.group(3) else ""
            return f"{m.group(1)}{m.group(2).lower().replace('ero', 'ro')}{sec}"
    toks = stem.split()
    for i, tok in enumerate(toks):
        m = re.fullmatch(r"([1-6])(ro|do|er|to|mo|vo|ero)", normalizar(tok).strip(".,"))
        if m:
            sec = ""
            if i + 1 < len(toks) and re.fullmatch(r"[A-Ea-e]", toks[i + 1].strip(".,")):
                sec = " " + toks[i + 1].strip(".,").upper()
            return f"{m.group(1)}{m.group(2).replace('ero', 'ro')}{sec}"
    return ""


def _edad(lineas):
    for l in lineas[:8]:
        n = normalizar(l)
        m = re.fullmatch(r"\s*(\d{1,2})\s*(anos|ano)\s*", n) or re.search(r"\bedad\s*[:.]?\s*(\d{1,2})\b", n)
        if m:
            return int(m.group(1))
    return ""


def _motivo_ficha(lineas):
    for i, l in enumerate(lineas):
        if "motivo por el que se solicita" in normalizar(l):
            partes = []
            for x in lineas[i + 1:i + 8]:
                nx = normalizar(x)
                if "acuerdos" in nx or "firma" in nx:
                    break
                if not _solo_subrayado(x):
                    partes.append(x.strip(" _"))
            return " ".join(partes)
    return ""


def _responsable_ficha(lineas):
    for i, l in enumerate(lineas):
        if "quien deriva" in normalizar(l):
            resto = l.split(":", 1)[1] if ":" in l else ""
            if not _solo_subrayado(resto):
                return re.sub(r"\s+", " ", resto.strip(" _"))
            if i + 1 < len(lineas):
                sig = lineas[i + 1]
                if not _solo_subrayado(sig) and "motivo" not in normalizar(sig):
                    return re.sub(r"\s+", " ", sig.strip(" _"))
    return ""


def extraer(ruta, estado, lineas, anio=None):
    """Devuelve dict con los campos automaticos. Lo que no aparece, vacio."""
    anio = anio or date.today().year
    stem = Path(ruta).stem
    base = {"nombre": "", "grado": "", "edad": "", "fecha": "", "responsable": "",
            "motivo": "", "ficha": "No", "estado": "No derivado", "porque": ""}
    mes_t = mes_en_titulo(stem)
    nombre_t = _limpiar_titulo(stem)

    if estado != OK:
        base["nombre"] = nombre_t or stem
        base["grado"] = _grado([], stem)
        if mes_t:
            base["fecha"] = formatear_fecha(None, mes_t, anio)
        base["motivo"] = {
            MANUAL: "LEER MANUALMENTE (documento escaneado o ilegible)",
            VACIO: "(archivo vacío)",
            NO_LEIDO: f"(formato {Path(ruta).suffix.lower()} no se lee: ver el archivo)",
        }.get(estado, "")
        base["estado"] = ""  # no se leyo: no se asume ni Derivado ni No derivado
        base["porque"] = "No se pudo leer: revisar a mano"
        return base

    completo = "\n".join(lineas)
    tn = normalizar(completo)
    ficha = "ficha de derivacion" in tn
    base["ficha"] = "Sí" if ficha else "No"

    nombre = ""
    if ficha:
        nombre = _valor(r"nombres? y apellidos?\s*:\s*([^\n|]{3,80}?)(?=\s{2,}|_{2,}|$|\n)", completo)
    if not nombre and lineas:  # sin ficha: el titulo manda, salvo que la 1ra linea diga "Nombre: ..."
        m = re.match(r"\s*(?:nombres?(?: y apellidos?)?|apellidos? y nombres?|alumn[oa]|estudiante)\s*[:.\-]\s*(.+)",
                     lineas[0], re.I)
        if m and _nombre_valido(m.group(1).strip(" :-|_")):
            nombre = re.sub(r"\s+", " ", m.group(1)).strip(" :-|_")
    base["nombre"] = nombre or nombre_t or stem

    g = _valor(r"grado y secci[oó]n\s*:\s*([^\n|]*?)(?=\s*(?:1\.6|edad|_{2,}|$|\n))", completo) if ficha else ""
    base["grado"] = g or _grado(lineas, stem)
    e = _valor(r"edad\s*[:.]\s*(\d{1,2})\b", completo) if ficha else ""
    base["edad"] = int(e) if e else _edad(lineas)
    if ficha:
        base["responsable"] = _responsable_ficha(lineas)  # opcion A: solo la ficha

    # fecha: titulo primero; la fecha completa del texto solo si coincide el mes o no hay mes en el titulo
    f_txt = _fecha_en_texto(lineas, anio)
    if mes_t and f_txt and f_txt[1] == mes_t:
        base["fecha"] = formatear_fecha(*f_txt)
    elif mes_t:
        base["fecha"] = formatear_fecha(None, mes_t, anio)
    elif f_txt:
        base["fecha"] = formatear_fecha(*f_txt)

    temas = [et for pat, et in TEMAS if re.search(pat, tn)][:4]
    narr = _motivo_ficha(lineas) if ficha else ""
    if not narr:
        util = []
        for l in lineas[1:]:
            ln = normalizar(l)
            if re.fullmatch(r"[\d/\-\. :]+", ln) or re.match(r"\d+\s*(anos|ano)", ln):
                continue
            if re.match(r"(fecha|edad|apertura|vive con|\d(ro|do|to|er|mo|vo)\b)", ln) or len(l) < 20:
                continue
            util.append(l)
        narr = next((l for l in util if CUE.search(normalizar(l))), util[0] if util else "")
    narr = (narr[:220] + "…") if len(narr) > 220 else narr
    partes = [p for p in (", ".join(temas), narr) if p]
    base["motivo"] = ". ".join(partes) or "(sin información suficiente)"
    base["estado"], por_que = estado_y_porque(completo)  # Derivado solo si es algo EXTERNO
    base["porque"] = por_que
    return base
