import copy, re
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.text.paragraph import Paragraph
from docx.table import Table

plantilla = r"C:\Users\Marco\Desktop\DOCS\plantilla.docx"
origen = r"C:\Users\Marco\Desktop\DOCS\PRUEBA.docx"
salida = r"C:\Users\Marco\Desktop\DOCS\RESULTADO_PRUEBA.docx"

ESTILO_TABLA_TEXTO = "Sin espaciado"   # si falla, el script prueba "No Spacing"

MAPA = {
    "Heading 1": "Heading 1", "Heading 2": "Heading 2",
    "Heading 3": "Heading 3", "Heading 4": "Heading 4",
    "Normal": "Normal", "Caption": "Caption",
}

# Numeración manual al inicio: "1. ", "1.1 ", "1.1.1. ", "2.3.4 "
NUM_RE = re.compile(r"^\s*\d{1,2}(?:\.\d{1,2})*\.?\s+")

# Etiquetas que generan título con campo SEQ (y su propio índice).
# Para agregar otra en el futuro, solo añádela aquí.
ETIQUETAS = ["Tabla", "Figura", "Código"]
ALIAS = {"Codigo": "Código"}   # variantes sin tilde -> nombre oficial

_nombres = "|".join(ETIQUETAS + list(ALIAS.keys()))
CAP_RE = re.compile(rf"^\s*({_nombres})\s+(\d+)\s*[.:\-–—]?\s*(.*)$", re.DOTALL)

src = Document(origen)
dst = Document(plantilla)
body = dst.element.body

for el in list(body):
    if not el.tag.endswith("sectPr"):
        body.remove(el)
sectPr = body.find(qn("w:sectPr"))

def get_style(nombres):
    for n in nombres:
        try:
            return dst.styles[n]
        except KeyError:
            continue
    return None

estilo_celda = get_style([ESTILO_TABLA_TEXTO, "No Spacing"])
ESTILO_CAPTION = get_style(["Descripción", "Caption"])   # mismo para todas las etiquetas

R_EMBED = qn("r:embed")
R_ID = qn("r:id")
_cache = {}

def copy_image_part(rid):
    if rid in _cache:
        return _cache[rid]
    part = src.part.related_parts[rid]
    n = len(_cache) + 1
    partname = PackURI(f"/word/media/copia_{n}.{part.partname.ext}")
    new_part = Part(partname, part.content_type, part.blob, dst.part.package)
    new_rid = dst.part.relate_to(new_part, RT.IMAGE)
    _cache[rid] = new_rid
    return new_rid

def remap_imgs(el):
    for node in el.iter():
        rid = node.get(R_EMBED)
        if rid and rid in src.part.related_parts:
            node.set(R_EMBED, copy_image_part(rid))
        if node.tag.endswith("}imagedata"):
            rid = node.get(R_ID)
            if rid and rid in src.part.related_parts:
                node.set(R_ID, copy_image_part(rid))

def add_field(par, instr, resultado):
    def fld(tipo):
        r = OxmlElement("w:r")
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), tipo)
        r.append(f)
        return r
    par._p.append(fld("begin"))
    r = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r.append(it)
    par._p.append(r)
    par._p.append(fld("separate"))
    par.add_run(resultado)
    par._p.append(fld("end"))

def add_caption(texto):
    m = CAP_RE.match(texto)
    if not m:
        return False
    etiqueta, num, resto = m.groups()
    etiqueta = ALIAS.get(etiqueta, etiqueta)   # variantes sin tilde -> oficial
    par = dst.add_paragraph(style=ESTILO_CAPTION)
    par.add_run(f"{etiqueta} ")
    add_field(par, f"SEQ {etiqueta} \\* ARABIC", num)
    par.add_run(f". {resto.strip()}")
    return True

def add_par(p_src):
    # Títulos de tabla/figura/código -> campo SEQ (para el índice)
    if add_caption(p_src.text):
        return

    nombre = MAPA.get(p_src.style.name, "Normal")
    nuevo = dst.add_paragraph(style=nombre)

    # Títulos: quitar numeración manual y dejar solo el texto
    if nombre.startswith("Heading"):
        texto = NUM_RE.sub("", p_src.text, count=1)
        nuevo.add_run(texto)
        return

    for r in p_src.runs:
        if r._r.findall(".//" + qn("w:drawing")) or r._r.findall(".//" + qn("w:pict")):
            r_copy = copy.deepcopy(r._r)
            remap_imgs(r_copy)
            nuevo._p.append(r_copy)
        else:
            run = nuevo.add_run(r.text)
            run.bold, run.italic, run.underline = r.bold, r.italic, r.underline

def add_table(t_src):
    tbl = copy.deepcopy(t_src._tbl)
    remap_imgs(tbl)
    sectPr.addprevious(tbl)
    try:
        Table(tbl, dst._body).style = "Table Grid"
    except Exception:
        pass

    if estilo_celda is None:
        print("AVISO: no se encontró el estilo 'Sin espaciado' en la plantilla.")
        return

    QUITAR_RUN = ("w:rFonts", "w:sz", "w:szCs", "w:color", "w:highlight",
                  "w:rStyle", "w:lang", "w:shd", "w:spacing", "w:w", "w:kern")

    for p in tbl.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            p.remove(pPr)
        Paragraph(p, dst._body).style = estilo_celda

        for r in p.iter(qn("w:r")):
            rPr = r.find(qn("w:rPr"))
            if rPr is not None:
                for tag in QUITAR_RUN:
                    for e in rPr.findall(qn(tag)):
                        rPr.remove(e)

for child in src.element.body.iterchildren():
    if child.tag == qn("w:p"):
        add_par(Paragraph(child, src._body))
    elif child.tag == qn("w:tbl"):
        add_table(Table(child, src._body))

dst.save(salida)
print("Listo:", salida)