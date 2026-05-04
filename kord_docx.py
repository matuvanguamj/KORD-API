"""
KORD DOCX — Pré-rapport Word éditable
Facile à modifier dans Word ou Google Docs.
"""
import io
from datetime import datetime
from typing import Dict, Any

from docx import Document
from docx.shared import Pt, Mm, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import base64


def set_cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.replace('#',''))
    tcPr.append(shd)


def add_image_from_bytes(doc, img_bytes: bytes, width_cm: float) -> None:
    buf = io.BytesIO(img_bytes)
    doc.add_picture(buf, width=Cm(width_cm))


def generate_prereport_docx(
    consolidated: Dict[str, Any],
    recommendations: Dict[str, Any],
    all_results: list,
    client_name: str = "Client",
    company_name: str = "",
    trimestre: str = "",
    radar_png: bytes = None,
    evol_png: bytes  = None,
    bar_png: bytes   = None,
    gauge_png: bytes = None,
) -> bytes:

    doc = Document()

    # ── Marges ──
    for section in doc.sections:
        section.page_width  = Mm(210)
        section.page_height = Mm(297)
        section.top_margin    = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin   = Mm(20)
        section.right_margin  = Mm(20)

    # ── Styles de base ──
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    score    = consolidated.get("score_total", 0)
    now      = datetime.now()
    date_str = now.strftime("%d %B %Y")
    trim_str = trimestre or f"T{(now.month-1)//3+1} {now.year}"

    # ══════════════════════════════════
    # PAGE DE COUVERTURE
    # ══════════════════════════════════
    # Entête noire
    cover_table = doc.add_table(rows=1, cols=1)
    cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cover_cell = cover_table.rows[0].cells[0]
    set_cell_bg(cover_cell, '#000000')
    cover_cell.width = Mm(170)

    p = cover_cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(16)
    run = p.add_run("KORD")
    run.bold = True
    run.font.size = Pt(32)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run.font.name = 'Arial'

    p2 = cover_cell.add_paragraph()
    r2 = p2.add_run("PRÉ-RAPPORT D'AUDIT OPÉRATIONNEL")
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    r2.font.name = 'Arial'

    p3 = cover_cell.add_paragraph()
    p3.paragraph_format.space_before = Pt(20)
    r3 = p3.add_run((company_name or client_name).upper())
    r3.bold = True
    r3.font.size = Pt(20)
    r3.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r3.font.name = 'Arial'

    if company_name and client_name:
        p4 = cover_cell.add_paragraph()
        r4 = p4.add_run(client_name)
        r4.font.size = Pt(11)
        r4.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
        r4.font.name = 'Arial'

    p5 = cover_cell.add_paragraph()
    r5 = p5.add_run(f"{trim_str}  ·  {date_str}  ·  {len(all_results)} fichier(s) analysé(s)")
    r5.font.size = Pt(8)
    r5.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    r5.font.name = 'Arial'

    p6 = cover_cell.add_paragraph()
    p6.paragraph_format.space_before = Pt(16)
    p6.paragraph_format.space_after  = Pt(16)
    r6 = p6.add_run(f"SCORE KORD : {score}/100  [INDICATIF]")
    r6.bold = True
    r6.font.size = Pt(14)
    r6.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r6.font.name = 'Arial'

    # Jauge score
    if gauge_png:
        p_gauge = cover_cell.add_paragraph()
        p_gauge.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_gauge = p_gauge.add_run()
        run_gauge.add_picture(io.BytesIO(gauge_png), width=Cm(10))
        p_gauge.paragraph_format.space_after = Pt(8)

    doc.add_paragraph()

    # Note indicative
    note_table = doc.add_table(rows=1, cols=1)
    note_cell  = note_table.rows[0].cells[0]
    set_cell_bg(note_cell, '#F4F3F0')
    pn = note_cell.paragraphs[0]
    rn = pn.add_run("⚠  DOCUMENT DE TRAVAIL — USAGE INTERNE KORD")
    rn.bold = True
    rn.font.size = Pt(8)
    rn.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    rn.font.name = 'Arial'
    pn2 = note_cell.add_paragraph()
    rn2 = pn2.add_run(
        "Ce pré-rapport est généré automatiquement. "
        "Il doit être relu et validé par l'équipe KORD avant toute restitution client. "
        "Le score est indicatif et sera ajusté après analyse approfondie."
    )
    rn2.font.size = Pt(8)
    rn2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    rn2.font.name = 'Arial'
    pn2.paragraph_format.space_after = Pt(6)

    doc.add_paragraph()

    # ══════════════════════════════════
    # RÉSUMÉ EXÉCUTIF
    # ══════════════════════════════════
    h1 = doc.add_paragraph()
    r_h1 = h1.add_run("01 — RÉSUMÉ EXÉCUTIF")
    r_h1.bold = True
    r_h1.font.size = Pt(11)
    r_h1.font.name = 'Arial'
    r_h1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    h1.paragraph_format.border_bottom = True
    _add_border_bottom(h1, '000000')

    summary = recommendations.get("resume_executif", "")
    if summary:
        p_sum = doc.add_paragraph()
        r_sum = p_sum.add_run(summary)
        r_sum.font.size = Pt(10)
        r_sum.font.name = 'Arial'
        r_sum.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    doc.add_paragraph()

    # ══════════════════════════════════
    # SCORES PAR PILIER + GRAPHIQUES
    # ══════════════════════════════════
    h2 = doc.add_paragraph()
    r_h2 = h2.add_run("02 — SCORES PAR PILIER")
    r_h2.bold = True
    r_h2.font.size = Pt(11)
    r_h2.font.name = 'Arial'
    _add_border_bottom(h2, '000000')

    pilier_labels = {
        "stock_cash":          ("Stock et Cash immobilisé",    30),
        "transport_service":   ("Transport et Taux de service", 20),
        "achats_fournisseurs": ("Achats et Fournisseurs",       20),
        "marges_retours":      ("Marges et Retours",            15),
        "donnees_pilotage":    ("Données et Pilotage",          15),
    }

    analyses = consolidated.get("analyses", {})
    scores_table = doc.add_table(rows=1, cols=4)
    scores_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Header
    hdr = scores_table.rows[0]
    for i, txt in enumerate(["Pilier", "Score", "Max", "Niveau"]):
        c = hdr.cells[i]
        set_cell_bg(c, '#000000')
        p = c.paragraphs[0]
        r = p.add_run(txt)
        r.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.name = 'Arial'

    for key, (label, max_pts) in pilier_labels.items():
        a   = analyses.get(key, {})
        s   = a.get("score", 0)
        pct = round(s / max_pts * 100) if max_pts > 0 else 0
        niveau = "Bon" if pct >= 70 else ("Moyen" if pct >= 45 else "Critique")
        bg = 'F4F3F0'

        row = scores_table.add_row()
        data = [label, str(s), f"/{max_pts}", niveau]
        for i, val in enumerate(data):
            c = row.cells[i]
            set_cell_bg(c, bg)
            p = c.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9)
            r.font.name = 'Arial'
            r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            if i == 1:
                r.bold = True

    doc.add_paragraph()

    # Graphiques
    if bar_png or radar_png:
        pg = doc.add_paragraph()
        pg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        chart_table = doc.add_table(rows=1, cols=2 if (bar_png and radar_png) else 1)
        chart_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        if bar_png:
            c1 = chart_table.rows[0].cells[0]
            p1 = c1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run1 = p1.add_run()
            run1.add_picture(io.BytesIO(bar_png), width=Cm(9))

        if radar_png:
            c2 = chart_table.rows[0].cells[-1]
            p2 = c2.paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run2 = p2.add_run()
            run2.add_picture(io.BytesIO(radar_png), width=Cm(7))

    if evol_png:
        pe = doc.add_paragraph()
        pe.alignment = WD_ALIGN_PARAGRAPH.CENTER
        re = pe.add_run()
        re.add_picture(io.BytesIO(evol_png), width=Cm(16))

    doc.add_paragraph()

    # ══════════════════════════════════
    # PRIORITÉS D'ACTION
    # ══════════════════════════════════
    h3 = doc.add_paragraph()
    r_h3 = h3.add_run("03 — PRIORITÉS D'ACTION")
    r_h3.bold = True
    r_h3.font.size = Pt(11)
    r_h3.font.name = 'Arial'
    _add_border_bottom(h3, '000000')

    for p_item in recommendations.get("priorites", [])[:5]:
        rang   = p_item.get("rang", "")
        titre  = p_item.get("titre", "")
        desc   = p_item.get("description", "")
        action = p_item.get("action", "")
        impact = p_item.get("impact", "")
        delai  = p_item.get("delai", "")

        pt = doc.add_paragraph()
        rt = pt.add_run(f"{str(rang).zfill(2)}  {titre.upper()}")
        rt.bold = True
        rt.font.size = Pt(10)
        rt.font.name = 'Arial'
        rt.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        set_paragraph_bg(pt, 'F4F3F0')

        if desc:
            pd = doc.add_paragraph()
            rd = pd.add_run(desc)
            rd.font.size = Pt(9)
            rd.font.name = 'Arial'
            rd.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            pd.paragraph_format.left_indent = Mm(8)

        if action:
            pa = doc.add_paragraph()
            ra = pa.add_run(f"Action : {action}")
            ra.font.size = Pt(9)
            ra.font.name = 'Arial'
            ra.italic = True
            ra.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            pa.paragraph_format.left_indent = Mm(8)

        if impact or delai:
            pi = doc.add_paragraph()
            ri = pi.add_run(f"Impact estimé : {impact}  |  {delai}")
            ri.bold = True
            ri.font.size = Pt(9)
            ri.font.name = 'Arial'
            ri.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            pi.paragraph_format.left_indent = Mm(8)
            pi.paragraph_format.space_after = Pt(8)

    doc.add_paragraph()

    # ══════════════════════════════════
    # ALERTES
    # ══════════════════════════════════
    alertes = consolidated.get("alertes", [])
    if alertes:
        h4 = doc.add_paragraph()
        r_h4 = h4.add_run("04 — ANOMALIES DÉTECTÉES")
        r_h4.bold = True
        r_h4.font.size = Pt(11)
        r_h4.font.name = 'Arial'
        _add_border_bottom(h4, '000000')

        pilier_short = {
            "stock_cash":          "Stock",
            "transport_service":   "Transport",
            "achats_fournisseurs": "Achats",
            "marges_retours":      "Marges",
            "donnees_pilotage":    "Données",
        }
        for a in alertes[:10]:
            pa = doc.add_paragraph()
            ra1 = pa.add_run(f"[{pilier_short.get(a.get('pilier',''), a.get('pilier','').upper())}]  ")
            ra1.bold = True
            ra1.font.size = Pt(9)
            ra1.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
            ra1.font.name = 'Arial'
            ra2 = pa.add_run(a.get("message", ""))
            ra2.font.size = Pt(9)
            ra2.font.name = 'Arial'
            ra2.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

        doc.add_paragraph()

    # ══════════════════════════════════
    # POINTS DE VIGILANCE
    # ══════════════════════════════════
    vigilance = recommendations.get("points_vigilance", [])
    if vigilance:
        h5 = doc.add_paragraph()
        r_h5 = h5.add_run("05 — POINTS DE VIGILANCE")
        r_h5.bold = True
        r_h5.font.size = Pt(11)
        r_h5.font.name = 'Arial'
        _add_border_bottom(h5, '000000')

        for i, pt in enumerate(vigilance, 1):
            pv = doc.add_paragraph()
            rv1 = pv.add_run(f"{i:02d}.  ")
            rv1.bold = True
            rv1.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
            rv1.font.name = 'Arial'
            rv1.font.size = Pt(9)
            rv2 = pv.add_run(pt)
            rv2.font.size = Pt(9)
            rv2.font.name = 'Arial'
            rv2.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

        doc.add_paragraph()

    # ══════════════════════════════════
    # PROCHAINE ÉTAPE
    # ══════════════════════════════════
    next_step = recommendations.get("prochaine_etape", "")
    if next_step:
        h6 = doc.add_paragraph()
        r_h6 = h6.add_run("06 — PROCHAINE ÉTAPE")
        r_h6.bold = True
        r_h6.font.size = Pt(11)
        r_h6.font.name = 'Arial'
        _add_border_bottom(h6, '000000')

        pns = doc.add_table(rows=1, cols=2)
        set_cell_bg(pns.rows[0].cells[0], 'F4F3F0')
        set_cell_bg(pns.rows[0].cells[1], 'F4F3F0')
        c0 = pns.rows[0].cells[0]
        r0 = c0.paragraphs[0].add_run("ACTION PRIORITAIRE")
        r0.bold = True
        r0.font.size = Pt(8)
        r0.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        r0.font.name = 'Arial'
        c1 = pns.rows[0].cells[1]
        r1 = c1.paragraphs[0].add_run(next_step)
        r1.bold = True
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        r1.font.name = 'Arial'

        doc.add_paragraph()

    # ══════════════════════════════════
    # SECTION GRAPHIQUES LIBRES
    # ══════════════════════════════════
    h7 = doc.add_paragraph()
    r_h7 = h7.add_run("07 — GRAPHIQUES ET ANNEXES")
    r_h7.bold = True
    r_h7.font.size = Pt(11)
    r_h7.font.name = 'Arial'
    _add_border_bottom(h7, '000000')

    p_note = doc.add_paragraph()
    r_note = p_note.add_run(
        "[ Insérez ici vos graphiques, analyses complémentaires et commentaires d'expert. "
        "Cette section est libre et peut être enrichie avant envoi au client. ]"
    )
    r_note.font.size = Pt(9)
    r_note.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    r_note.italic = True
    r_note.font.name = 'Arial'
    p_note.paragraph_format.space_before = Pt(12)
    p_note.paragraph_format.space_after  = Pt(40)

    # Footer disclaimer
    doc.add_paragraph()
    pf = doc.add_paragraph()
    _add_border_bottom(pf, 'CCCCCC')
    rf = pf.add_run(
        f"KORD — Pré-rapport d'audit opérationnel — {trim_str} — "
        "Document de travail confidentiel — Score indicatif à valider par les experts KORD"
    )
    rf.font.size = Pt(7)
    rf.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    rf.font.name = 'Arial'

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def _add_border_bottom(paragraph, color_hex: str = '000000'):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)


def set_paragraph_bg(paragraph, color_hex: str):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    pPr.append(shd)
