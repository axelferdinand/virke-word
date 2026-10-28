"""
Virke Word design system v3 — ekte stilark basert på Figma-tokens
fra seksjonen "A4 Word-template Design system" (node 2653:218736).

Registrerer navngitte paragraph-, character- og tabell-stiler i styles.xml,
slik at åpnet dokument viser "Virke Tittel", "Virke Overskrift 1" osv.
som valgbare stiler i Word's Stiler-panel — en ekte mal, ikke overstyling.

Figma-tokens:
  Word/H1         Bold  24pt  line-height 1.5
  Word/H2         Bold  18pt  line-height 1.5
  Word/H3         Bold  14pt  line-height 1.6
  Word/Paragraph  Reg   12pt  line-height 1.6
  Word/Caption    Reg   10pt  line-height 1.6

  Family/Sans              Clan Offc
  Family/Sans (Fallback)   Arial          ← brukes her

  Text/Text Black     #000000
  Text/Text Gray      #616161
  Text/Text Link      #C9007F   (Virke-magenta)
  Gray scale/Grå 25   #EDEDED
  Gray scale/Grå 75   #CACACA
"""
from docx.shared import Pt, Cm, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ─── TOKENS ───────────────────────────────────────────────
FONT_FAMILY = "Arial"  # Fallback iht. Figma (Family/Sans Fallback)

PT_H1 = 24
PT_H2 = 18
PT_H3 = 14
PT_BODY = 12
PT_CAPTION = 10
PT_HELP = 11  # litt mindre enn body, men fortsatt lesbart

COLOR_INK = RGBColor(0x00, 0x00, 0x00)
COLOR_MUTED = RGBColor(0x61, 0x61, 0x61)
COLOR_LINK = RGBColor(0xC9, 0x00, 0x7F)         # Virke-magenta
COLOR_PLACEHOLDER = RGBColor(0x61, 0x61, 0x61)  # muted grå, ikke magenta

COLOR_GRAY_25_HEX = "EDEDED"   # lys tabell-header / subtil skille
COLOR_GRAY_75_HEX = "CACACA"   # table border, divider
COLOR_LINK_HEX = "C9007F"


# ─── LAVNIVÅ-HJELPERE ─────────────────────────────────────
def _set_rfonts(rPr, family=FONT_FAMILY):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(qn(f"w:{attr}"), family)


def _configure_para_style(style, size_pt, bold=False, italic=False,
                          color=COLOR_INK, space_before=0, space_after=120,
                          line_rule="auto", line=288, keep_with_next=False):
    """Konfigurer font + spacing på en paragraph-stil.
       line er i twentieths of a point: 240 = single, 288 = 1.2, 360 = 1.5, 384 = 1.6."""
    font = style.font
    font.name = FONT_FAMILY
    font.size = Pt(size_pt)
    font.bold = bold
    font.italic = italic
    font.color.rgb = color
    rPr = style.element.get_or_add_rPr()
    _set_rfonts(rPr)

    pPr = style.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:spacing")):
        pPr.remove(old)
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(space_before))
    spacing.set(qn("w:after"), str(space_after))
    spacing.set(qn("w:line"), str(line))
    spacing.set(qn("w:lineRule"), line_rule)
    pPr.append(spacing)
    if keep_with_next:
        pPr.append(OxmlElement("w:keepNext"))


def _configure_char_style(style, size_pt=None, bold=False, italic=False,
                          color=COLOR_INK, underline=False):
    font = style.font
    font.name = FONT_FAMILY
    if size_pt:
        font.size = Pt(size_pt)
    font.bold = bold
    font.italic = italic
    font.underline = underline
    font.color.rgb = color
    rPr = style.element.get_or_add_rPr()
    _set_rfonts(rPr)


def _add_left_border(style, color_hex=COLOR_LINK_HEX, sz=24):
    """Legg en tykk venstre-kantlinje på en paragraph-stil (callout-stil).
       sz er i eighths of a point — 24 = 3pt."""
    pPr = style.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:pBdr")):
        pPr.remove(old)
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(sz))
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color_hex)
    pBdr.append(left)
    pPr.append(pBdr)
    # venstre-indent for å løfte teksten vekk fra borderen
    for old in pPr.findall(qn("w:ind")):
        pPr.remove(old)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "240")  # 12pt indent
    pPr.append(ind)


def _add_or_get_style(doc, name, wd_type):
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, wd_type)


# ─── STYLESHEET INSTALLER ─────────────────────────────────
def install_virke_stylesheet(doc):
    """Registrer alle Virke-stiler i dokumentets styles.xml."""
    styles = doc.styles

    # Normal (basis for alle andre)
    normal = styles["Normal"]
    _configure_para_style(normal, PT_BODY,
                          space_after=160, line=int(PT_BODY * 20 * 1.6))

    # ── PARAGRAPH STYLES ──────────────────────────────────
    tittel = _add_or_get_style(doc, "Virke Tittel", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(tittel, PT_H1, bold=True,
                          space_before=0, space_after=300,
                          line=int(PT_H1 * 20 * 1.5),
                          keep_with_next=True)
    tittel.base_style = normal
    tittel.quick_style = True

    h1 = _add_or_get_style(doc, "Virke Overskrift 1", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(h1, PT_H2, bold=True,
                          space_before=360, space_after=160,
                          line=int(PT_H2 * 20 * 1.5),
                          keep_with_next=True)
    h1.base_style = normal
    h1.quick_style = True

    h2 = _add_or_get_style(doc, "Virke Overskrift 2", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(h2, PT_H3, bold=True,
                          space_before=280, space_after=120,
                          line=int(PT_H3 * 20 * 1.6),
                          keep_with_next=True)
    h2.base_style = normal
    h2.quick_style = True

    h3 = _add_or_get_style(doc, "Virke Overskrift 3", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(h3, PT_BODY, bold=True,
                          space_before=200, space_after=80,
                          line=int(PT_BODY * 20 * 1.6),
                          keep_with_next=True)
    h3.base_style = normal
    h3.quick_style = True

    body = _add_or_get_style(doc, "Virke Brødtekst", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(body, PT_BODY,
                          space_before=0, space_after=160,
                          line=int(PT_BODY * 20 * 1.6))
    body.base_style = normal
    body.quick_style = True

    helptext = _add_or_get_style(doc, "Virke Hjelpetekst", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(helptext, PT_HELP, italic=True, color=COLOR_MUTED,
                          space_before=160, space_after=200,
                          line=int(PT_HELP * 20 * 1.6))
    _add_left_border(helptext, COLOR_LINK_HEX, sz=24)
    helptext.base_style = normal
    helptext.quick_style = True

    caption = _add_or_get_style(doc, "Virke Caption", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(caption, PT_CAPTION, color=COLOR_MUTED,
                          space_before=0, space_after=0,
                          line=int(PT_CAPTION * 20 * 1.6))
    caption.base_style = normal
    caption.quick_style = True

    signature = _add_or_get_style(doc, "Virke Signaturlinje", WD_STYLE_TYPE.PARAGRAPH)
    _configure_para_style(signature, PT_BODY,
                          space_before=240, space_after=60,
                          line=int(PT_BODY * 20 * 1.6))
    signature.base_style = normal

    # ── CHARACTER STYLES ──────────────────────────────────
    variabel = _add_or_get_style(doc, "Virke Variabel", WD_STYLE_TYPE.CHARACTER)
    _configure_char_style(variabel, italic=True, color=COLOR_PLACEHOLDER)
    variabel.quick_style = True

    fremhevet = _add_or_get_style(doc, "Virke Fremhevet", WD_STYLE_TYPE.CHARACTER)
    _configure_char_style(fremhevet, bold=True)
    fremhevet.quick_style = True

    lenke = _add_or_get_style(doc, "Virke Lenke", WD_STYLE_TYPE.CHARACTER)
    _configure_char_style(lenke, color=COLOR_LINK, underline=True)
    lenke.quick_style = True

    # ── TABLE STYLE (bare registrert som stilnavn) ────────
    try:
        doc.styles.add_style("Virke Tabell", WD_STYLE_TYPE.TABLE)
    except (KeyError, ValueError):
        pass

    return {
        "tittel": "Virke Tittel",
        "h1": "Virke Overskrift 1",
        "h2": "Virke Overskrift 2",
        "h3": "Virke Overskrift 3",
        "body": "Virke Brødtekst",
        "helptext": "Virke Hjelpetekst",
        "caption": "Virke Caption",
        "signature": "Virke Signaturlinje",
        "variabel": "Virke Variabel",
        "fremhevet": "Virke Fremhevet",
        "lenke": "Virke Lenke",
        "table": "Virke Tabell",
    }


# ─── BUILDING BLOCKS ─────────────────────────────────────
def add_tittel(doc, text):
    p = doc.add_paragraph(style="Virke Tittel")
    p.add_run(text)
    return p


def add_h1(doc, text):
    p = doc.add_paragraph(style="Virke Overskrift 1")
    p.add_run(text)
    return p


def add_h2(doc, text):
    p = doc.add_paragraph(style="Virke Overskrift 2")
    p.add_run(text)
    return p


def add_h3(doc, text):
    p = doc.add_paragraph(style="Virke Overskrift 3")
    p.add_run(text)
    return p


def add_body(doc, text=""):
    p = doc.add_paragraph(style="Virke Brødtekst")
    if text:
        p.add_run(text)
    return p


def add_helptext(doc, text, prefix="Hjelpetekst – fjernes ved utfylling: "):
    p = doc.add_paragraph(style="Virke Hjelpetekst")
    if prefix:
        r = p.add_run(prefix)
        r.bold = True
    p.add_run(text)
    return p


def add_placeholder_to(paragraph, text):
    run = paragraph.add_run(text)
    run.style = "Virke Variabel"
    return run


def add_bold_to(paragraph, text):
    run = paragraph.add_run(text)
    run.style = "Virke Fremhevet"
    return run


def add_body_with_placeholder(doc, parts):
    """parts: liste av ('text', str), ('placeholder', str), ('bold', str)."""
    p = doc.add_paragraph(style="Virke Brødtekst")
    for kind, text in parts:
        if kind == "placeholder":
            add_placeholder_to(p, text)
        elif kind == "bold":
            add_bold_to(p, text)
        else:
            p.add_run(text)
    return p


# ─── LAYOUT: marger, header, footer ──────────────────────
def setup_page(doc, margin_h_cm=2.25, margin_v_cm=2.0):
    """Figma viser 2,27 cm horisontal marg — vi runder til 2,25.
       Topp/bunn: 2,0 cm gir rom til header/footer uten å stjele for mye innhold."""
    for section in doc.sections:
        section.top_margin = Cm(margin_v_cm)
        section.bottom_margin = Cm(margin_v_cm)
        section.left_margin = Cm(margin_h_cm)
        section.right_margin = Cm(margin_h_cm)
        section.page_height = Cm(29.7)
        section.page_width = Cm(21.0)
        section.header_distance = Cm(1.0)
        section.footer_distance = Cm(1.0)


def _set_tab_right(paragraph, pos_twips=9360):
    pPr = paragraph._element.get_or_add_pPr()
    for old in pPr.findall(qn("w:tabs")):
        pPr.remove(old)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(pos_twips))
    tabs.append(tab)
    pPr.append(tabs)


def _set_para_border(paragraph, edge, color_hex="CACACA", sz=4, space=2):
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    for old in pBdr.findall(qn(f"w:{edge}")):
        pBdr.remove(old)
    b = OxmlElement(f"w:{edge}")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(sz))
    b.set(qn("w:space"), str(space))
    b.set(qn("w:color"), color_hex)
    pBdr.append(b)


def _add_page_field(paragraph, instr):
    run = paragraph.add_run()
    for tag, val in [
        ("w:fldChar", {"w:fldCharType": "begin"}),
        ("w:instrText", {}),
        ("w:fldChar", {"w:fldCharType": "end"}),
    ]:
        el = OxmlElement(tag)
        for k, v in val.items():
            el.set(qn(k), v)
        if tag == "w:instrText":
            el.text = instr
        run._element.append(el)
    return run


def add_header_footer(doc, section_label, doc_title):
    """Header: Virke-logo-plass (venstre) + seksjonslabel (høyre).
       Footer: Side X av Y (venstre) + doktittel (høyre).
       Per Figma: seksjon og dokumenttittel er *grunnteksten* — Virke-logoet er vanligvis
       skjult i eksemplene, men vi beholder 'Virke' som tekst-markør."""
    for section in doc.sections:
        # HEADER
        header = section.header
        hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        hp.style = doc.styles["Virke Caption"]
        _set_tab_right(hp)
        r = hp.add_run("Virke"); r.bold = True
        hp.add_run("\t")
        hp.add_run(section_label)
        _set_para_border(hp, "bottom", COLOR_GRAY_75_HEX, sz=4, space=1)

        # FOOTER
        footer = section.footer
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        fp.style = doc.styles["Virke Caption"]
        _set_tab_right(fp)
        fp.add_run("Side ")
        _add_page_field(fp, "PAGE")
        fp.add_run(" av ")
        _add_page_field(fp, "NUMPAGES")
        fp.add_run("\t")
        fp.add_run(doc_title)
        _set_para_border(fp, "top", COLOR_GRAY_75_HEX, sz=4, space=1)


# ─── TABELLER ────────────────────────────────────────────
def set_cell_shading(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def set_cell_margins(cell, top=60, bottom=60, left=120, right=120):
    tcPr = cell._element.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    mar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom),
                      ("left", left), ("right", right)]:
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tcPr.append(mar)


def set_table_borders(table, color=COLOR_GRAY_75_HEX, sz=4,
                      edges=("top", "left", "bottom", "right", "insideH", "insideV")):
    tbl = table._element
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        if edge in edges and sz > 0:
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), str(sz))
            b.set(qn("w:color"), color)
        else:
            b.set(qn("w:val"), "none")
            b.set(qn("w:sz"), "0")
            b.set(qn("w:color"), "auto")
        b.set(qn("w:space"), "0")
        borders.append(b)
    tblPr.append(borders)


def style_virke_table(table, header=True, borders=True, zebra=False):
    """Gi tabellen Virke-utseende: header-rad med Grå 25-bakgrunn, data-rader hvit.
       zebra=True legger Grå 25 på annenhver datarad (subtilt)."""
    if borders:
        set_table_borders(table, color=COLOR_GRAY_75_HEX, sz=4)
    else:
        set_table_borders(table, sz=0)
    for ridx, row in enumerate(table.rows):
        for cell in row.cells:
            set_cell_margins(cell)
            for p in cell.paragraphs:
                p.style = "Virke Brødtekst"
            if ridx == 0 and header:
                set_cell_shading(cell, COLOR_GRAY_25_HEX)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.bold = True
            elif zebra and header and ridx >= 1 and ridx % 2 == 0:
                set_cell_shading(cell, "F7F7F7")


# ─── DOTX-LAGRING ────────────────────────────────────────
def save_as_dotx(doc, path):
    """Lagre som template (.dotx) ved å endre [Content_Types].xml."""
    import zipfile, os
    tmp_docx = path + ".tmp.docx"
    doc.save(tmp_docx)
    tmp_dotx = path + ".tmp2.dotx"
    with zipfile.ZipFile(tmp_docx, "r") as zin, \
         zipfile.ZipFile(tmp_dotx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                )
            zout.writestr(item, data)
    os.replace(tmp_dotx, path)
    os.remove(tmp_docx)
