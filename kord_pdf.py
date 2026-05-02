"""
KORD PDF — Génération automatique du pré-rapport brandé KORD
Document de travail pour l'équipe KORD avant restitution au client.
"""

import io
from datetime import datetime
from typing import Dict, Any

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable

# ── COULEURS KORD ──
BLACK  = colors.HexColor("#000000")
WHITE  = colors.HexColor("#FFFFFF")
OFF    = colors.HexColor("#F4F3F0")
GRAY   = colors.HexColor("#6B6B6B")
LGRAY  = colors.HexColor("#E8E8E8")
DGRAY  = colors.HexColor("#1A1A1A")

W, H = A4  # 210mm x 297mm


# ── LIGNE NOIRE SÉPARATRICE ──
class BlackLine(Flowable):
    def __init__(self, width=None, thickness=0.5):
        super().__init__()
        self._width = width
        self.thickness = thickness

    def wrap(self, availWidth, availHeight):
        self._width = self._width or availWidth
        return self._width, self.thickness + 2

    def draw(self):
        self.canv.setStrokeColor(BLACK)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 1, self._width, 1)


class GrayLine(Flowable):
    def __init__(self, width=None):
        super().__init__()
        self._width = width

    def wrap(self, availWidth, availHeight):
        self._width = self._width or availWidth
        return self._width, 1.5

    def draw(self):
        self.canv.setStrokeColor(LGRAY)
        self.canv.setLineWidth(0.5)
        self.canv.line(0, 1, self._width, 1)


# ── STYLES ──
def get_styles():
    return {
        "kord_title": ParagraphStyle(
            "kord_title",
            fontName="Helvetica-Bold",
            fontSize=28,
            leading=32,
            textColor=WHITE,
            spaceAfter=4,
            spaceBefore=0,
        ),
        "kord_sub": ParagraphStyle(
            "kord_sub",
            fontName="Helvetica",
            fontSize=9,
            leading=13,
            textColor=colors.HexColor("#999999"),
            spaceAfter=0,
        ),
        "section_label": ParagraphStyle(
            "section_label",
            fontName="Helvetica-Bold",
            fontSize=7,
            leading=10,
            textColor=GRAY,
            spaceAfter=6,
            spaceBefore=18,
            wordWrap='CJK',
        ),
        "section_title": ParagraphStyle(
            "section_title",
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=17,
            textColor=BLACK,
            spaceAfter=10,
            spaceBefore=4,
        ),
        "body": ParagraphStyle(
            "body",
            fontName="Helvetica",
            fontSize=9,
            leading=14,
            textColor=DGRAY,
            spaceAfter=6,
        ),
        "body_gray": ParagraphStyle(
            "body_gray",
            fontName="Helvetica",
            fontSize=9,
            leading=14,
            textColor=GRAY,
            spaceAfter=4,
        ),
        "score_big": ParagraphStyle(
            "score_big",
            fontName="Helvetica-Bold",
            fontSize=52,
            leading=56,
            textColor=BLACK,
            alignment=TA_CENTER,
        ),
        "score_label": ParagraphStyle(
            "score_label",
            fontName="Helvetica",
            fontSize=8,
            leading=11,
            textColor=GRAY,
            alignment=TA_CENTER,
        ),
        "tag": ParagraphStyle(
            "tag",
            fontName="Helvetica-Bold",
            fontSize=7,
            leading=10,
            textColor=GRAY,
        ),
        "priority_num": ParagraphStyle(
            "priority_num",
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=20,
            textColor=LGRAY,
        ),
        "priority_title": ParagraphStyle(
            "priority_title",
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=14,
            textColor=BLACK,
            spaceAfter=3,
        ),
        "priority_desc": ParagraphStyle(
            "priority_desc",
            fontName="Helvetica",
            fontSize=8.5,
            leading=13,
            textColor=DGRAY,
            spaceAfter=3,
        ),
        "priority_impact": ParagraphStyle(
            "priority_impact",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=12,
            textColor=BLACK,
        ),
        "alert_text": ParagraphStyle(
            "alert_text",
            fontName="Helvetica",
            fontSize=8.5,
            leading=13,
            textColor=DGRAY,
        ),
        "footer_text": ParagraphStyle(
            "footer_text",
            fontName="Helvetica",
            fontSize=7,
            leading=10,
            textColor=GRAY,
        ),
        "indicatif": ParagraphStyle(
            "indicatif",
            fontName="Helvetica",
            fontSize=8,
            leading=12,
            textColor=GRAY,
            borderPad=8,
        ),
    }


def generate_prereport_pdf(
    consolidated: Dict[str, Any],
    recommendations: Dict[str, Any],
    all_results: list,
    client_name: str = "Client",
    company_name: str = "",
    trimestre: str = "",
) -> bytes:
    """
    Génère le pré-rapport PDF KORD.
    Retourne les bytes du PDF.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18*mm,
        rightMargin=18*mm,
        topMargin=0,
        bottomMargin=14*mm,
    )

    S = get_styles()
    story = []
    score = consolidated.get("score_total", 0)
    now   = datetime.now()
    date_str = now.strftime("%d %B %Y").upper()
    trim_str  = trimestre or f"T{(now.month-1)//3+1} {now.year}"

    # ── PAGE DE COUVERTURE ──
    story += _cover_page(S, score, client_name, company_name, trim_str, date_str,
                         len(all_results), consolidated)

    # ── RÉSUMÉ EXÉCUTIF ──
    story += _executive_summary(S, recommendations, consolidated)

    # ── SCORES PAR PILIER ──
    story += _scores_section(S, consolidated)

    # ── PRIORITÉS D'ACTION ──
    story += _priorities_section(S, recommendations)

    # ── ALERTES DÉTECTÉES ──
    story += _alerts_section(S, consolidated)

    # ── POINTS DE VIGILANCE ──
    story += _vigilance_section(S, recommendations)

    # ── PROCHAINE ÉTAPE ──
    story += _next_step_section(S, recommendations)

    # ── NOTE INDICATIVE ──
    story += _disclaimer_section(S)

    doc.build(story, onFirstPage=_add_header_footer, onLaterPages=_add_header_footer)
    buf.seek(0)
    return buf.read()


def _cover_page(S, score, client, company, trimestre, date, nb_files, consolidated):
    story = []

    # Bloc couverture noir
    cover_data = [[
        Table([
            [Paragraph("KORD", S["kord_title"])],
            [Paragraph("PRÉ-RAPPORT D'AUDIT OPÉRATIONNEL", ParagraphStyle(
                "cover_sub", fontName="Helvetica", fontSize=9, leading=13,
                textColor=colors.HexColor("#888888"), spaceAfter=20
            ))],
            [Spacer(1, 16*mm)],
            [Paragraph(company.upper() if company else client.upper(), ParagraphStyle(
                "cover_company", fontName="Helvetica-Bold", fontSize=18, leading=22,
                textColor=WHITE, spaceAfter=6
            ))],
            [Paragraph(client, ParagraphStyle(
                "cover_client", fontName="Helvetica", fontSize=11, leading=15,
                textColor=colors.HexColor("#BBBBBB"), spaceAfter=4
            ))],
            [Paragraph(f"{trimestre}  ·  {date}  ·  {nb_files} fichier{'s' if nb_files>1 else ''} analysé{'s' if nb_files>1 else ''}", ParagraphStyle(
                "cover_meta", fontName="Helvetica", fontSize=8, leading=12,
                textColor=colors.HexColor("#666666"),
            ))],
            [Spacer(1, 12*mm)],
            [Table([
                [Paragraph(str(score), ParagraphStyle(
                    "cover_score_num", fontName="Helvetica-Bold", fontSize=48,
                    leading=52, textColor=WHITE, alignment=TA_CENTER
                ))],
                [Paragraph("/100 — SCORE KORD INDICATIF", ParagraphStyle(
                    "cover_score_lbl", fontName="Helvetica", fontSize=7,
                    leading=10, textColor=colors.HexColor("#777777"), alignment=TA_CENTER
                ))],
            ], colWidths=[80*mm],
            style=TableStyle([
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                ("BACKGROUND", (0,0), (-1,-1), DGRAY),
                ("TOPPADDING", (0,0), (-1,-1), 12),
                ("BOTTOMPADDING", (0,0), (-1,-1), 12),
            ]))],
        ], colWidths=[170*mm],
        style=TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), BLACK),
            ("LEFTPADDING", (0,0), (-1,-1), 16*mm),
            ("RIGHTPADDING", (0,0), (-1,-1), 16*mm),
            ("TOPPADDING", (0,0), (0,0), 18*mm),
            ("TOPPADDING", (0,1), (-1,-1), 0),
            ("BOTTOMPADDING", (0,-1), (-1,-1), 18*mm),
        ]))
    ]]

    cover_table = Table(cover_data, colWidths=[174*mm])
    cover_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BLACK),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ]))
    story.append(cover_table)
    story.append(Spacer(1, 8*mm))

    # Note document de travail
    note = Table([[
        Paragraph("DOCUMENT DE TRAVAIL — USAGE INTERNE KORD", ParagraphStyle(
            "note_tag", fontName="Helvetica-Bold", fontSize=7, leading=10,
            textColor=GRAY
        )),
        Paragraph("Ce pré-rapport est généré automatiquement par le moteur KORD. Il doit être relu et validé par l'équipe avant toute restitution client.", ParagraphStyle(
            "note_body", fontName="Helvetica", fontSize=7.5, leading=11,
            textColor=GRAY, alignment=TA_RIGHT
        )),
    ]], colWidths=[70*mm, 100*mm])
    note.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), OFF),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(note)
    story.append(Spacer(1, 6*mm))

    return story


def _executive_summary(S, recommendations, consolidated):
    story = []
    summary = recommendations.get("resume_executif", "")
    if not summary:
        return story

    story.append(Paragraph("01  —  RÉSUMÉ EXÉCUTIF", S["section_label"]))
    story.append(BlackLine())
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(summary, S["body"]))
    story.append(Spacer(1, 6*mm))
    return story


def _scores_section(S, consolidated):
    story = []
    analyses = consolidated.get("analyses", {})
    if not analyses:
        return story

    story.append(Paragraph("02  —  SCORES PAR PILIER", S["section_label"]))
    story.append(BlackLine())
    story.append(Spacer(1, 4*mm))

    pilier_labels = {
        "stock_cash":          ("Stock et Cash immobilisé",    "30 pts"),
        "transport_service":   ("Transport et Taux de service", "20 pts"),
        "achats_fournisseurs": ("Achats et Fournisseurs",       "20 pts"),
        "marges_retours":      ("Marges et Retours",            "15 pts"),
        "donnees_pilotage":    ("Données et Pilotage",          "15 pts"),
    }

    rows = []
    for key, (label, max_label) in pilier_labels.items():
        if key not in analyses:
            continue
        a   = analyses[key]
        s   = a.get("score", 0)
        m   = a.get("max", 0)
        pct = round(s / m * 100) if m > 0 else 0
        etat = "BON" if pct >= 70 else ("MOYEN" if pct >= 45 else "CRITIQUE")
        etat_color = colors.HexColor("#222222") if pct >= 70 else (
            colors.HexColor("#555555") if pct >= 45 else BLACK
        )
        bar_filled = int(pct / 100 * 80)

        rows.append([
            Paragraph(label, ParagraphStyle(
                "pn", fontName="Helvetica-Bold", fontSize=9, leading=13, textColor=BLACK
            )),
            Paragraph(f"{s}/{m}", ParagraphStyle(
                "ps", fontName="Helvetica-Bold", fontSize=11, leading=14,
                textColor=BLACK, alignment=TA_CENTER
            )),
            Paragraph(etat, ParagraphStyle(
                "pe", fontName="Helvetica-Bold", fontSize=7, leading=10,
                textColor=etat_color, alignment=TA_CENTER
            )),
            Paragraph(f"{pct}%", ParagraphStyle(
                "pp", fontName="Helvetica", fontSize=8, leading=11,
                textColor=GRAY, alignment=TA_RIGHT
            )),
        ])

    if rows:
        t = Table(rows, colWidths=[90*mm, 28*mm, 28*mm, 24*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), WHITE),
            ("BACKGROUND",    (0, 0), (-1,  0), OFF),
            ("ROWBACKGROUNDS",(0, 0), (-1, -1), [WHITE, OFF]),
            ("LEFTPADDING",   (0, 0), (-1, -1), 8),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
            ("TOPPADDING",    (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("LINEBELOW",     (0, 0), (-1, -1), 0.5, LGRAY),
            ("BOX",           (0, 0), (-1, -1), 0.5, LGRAY),
        ]))
        story.append(t)

    story.append(Spacer(1, 6*mm))
    return story


def _priorities_section(S, recommendations):
    story = []
    priorites = recommendations.get("priorites", [])
    if not priorites:
        return story

    story.append(Paragraph("03  —  PRIORITÉS D'ACTION", S["section_label"]))
    story.append(BlackLine())
    story.append(Spacer(1, 4*mm))

    for p in priorites[:5]:
        rang   = p.get("rang", "")
        titre  = p.get("titre", "")
        desc   = p.get("description", "")
        action = p.get("action", "")
        impact = p.get("impact", "")
        delai  = p.get("delai", "")

        block = KeepTogether([
            Table([[
                Paragraph(str(rang).zfill(2), S["priority_num"]),
                Table([
                    [Paragraph(titre, S["priority_title"])],
                    [Paragraph(desc, S["priority_desc"])],
                    [Paragraph(f"Action : {action}", S["priority_desc"])],
                    [Table([[
                        Paragraph(f"Impact : {impact}", S["priority_impact"]),
                        Paragraph(delai, ParagraphStyle(
                            "delai", fontName="Helvetica", fontSize=7, leading=10,
                            textColor=GRAY, alignment=TA_RIGHT
                        )),
                    ]], colWidths=[90*mm, 40*mm],
                    style=TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))],
                ], colWidths=[134*mm],
                style=TableStyle([
                    ("LEFTPADDING",(0,0),(-1,-1),0),
                    ("RIGHTPADDING",(0,0),(-1,-1),0),
                    ("TOPPADDING",(0,0),(-1,-1),1),
                    ("BOTTOMPADDING",(0,0),(-1,-1),1),
                ])),
            ]], colWidths=[18*mm, 136*mm],
            style=TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), WHITE),
                ("LEFTPADDING",   (0, 0), (-1, -1), 10),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
                ("TOPPADDING",    (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
                ("BOX",           (0, 0), (-1, -1), 0.5, LGRAY),
                ("LINEBELOW",     (0, 0), (-1, -1), 0.5, LGRAY),
            ])),
            Spacer(1, 3*mm),
        ])
        story.append(block)

    story.append(Spacer(1, 4*mm))
    return story


def _alerts_section(S, consolidated):
    story = []
    alertes = consolidated.get("alertes", [])
    if not alertes:
        return story

    story.append(Paragraph("04  —  ANOMALIES DÉTECTÉES", S["section_label"]))
    story.append(BlackLine())
    story.append(Spacer(1, 4*mm))

    pilier_labels = {
        "stock_cash":          "Stock",
        "transport_service":   "Transport",
        "achats_fournisseurs": "Achats",
        "marges_retours":      "Marges",
        "donnees_pilotage":    "Données",
    }

    rows = []
    for a in alertes[:10]:
        pilier = pilier_labels.get(a.get("pilier",""), a.get("pilier","").upper())
        msg    = a.get("message", "")
        rows.append([
            Paragraph(pilier, ParagraphStyle(
                "pilier_badge", fontName="Helvetica-Bold", fontSize=7, leading=10,
                textColor=GRAY
            )),
            Paragraph(msg, S["alert_text"]),
        ])

    if rows:
        t = Table(rows, colWidths=[26*mm, 144*mm])
        t.setStyle(TableStyle([
            ("ROWBACKGROUNDS", (0, 0), (-1, -1), [WHITE, OFF]),
            ("LEFTPADDING",    (0, 0), (-1, -1), 8),
            ("RIGHTPADDING",   (0, 0), (-1, -1), 8),
            ("TOPPADDING",     (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING",  (0, 0), (-1, -1), 7),
            ("VALIGN",         (0, 0), (-1, -1), "TOP"),
            ("LINEBELOW",      (0, 0), (-1, -1), 0.5, LGRAY),
            ("BOX",            (0, 0), (-1, -1), 0.5, LGRAY),
        ]))
        story.append(t)

    story.append(Spacer(1, 6*mm))
    return story


def _vigilance_section(S, recommendations):
    story = []
    points = recommendations.get("points_vigilance", [])
    if not points:
        return story

    story.append(Paragraph("05  —  POINTS DE VIGILANCE", S["section_label"]))
    story.append(BlackLine())
    story.append(Spacer(1, 4*mm))

    rows = [[
        Paragraph(f"{i+1:02d}", ParagraphStyle(
            "vnum", fontName="Helvetica-Bold", fontSize=9, leading=13, textColor=LGRAY
        )),
        Paragraph(pt, S["body"]),
    ] for i, pt in enumerate(points)]

    if rows:
        t = Table(rows, colWidths=[12*mm, 158*mm])
        t.setStyle(TableStyle([
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("LINEBELOW",     (0, 0), (-1, -1), 0.5, LGRAY),
        ]))
        story.append(t)

    story.append(Spacer(1, 6*mm))
    return story


def _next_step_section(S, recommendations):
    story = []
    next_step = recommendations.get("prochaine_etape", "")
    if not next_step:
        return story

    story.append(Paragraph("06  —  PROCHAINE ÉTAPE", S["section_label"]))
    story.append(BlackLine())
    story.append(Spacer(1, 4*mm))

    t = Table([[
        Paragraph("ACTION PRIORITAIRE", ParagraphStyle(
            "ap", fontName="Helvetica-Bold", fontSize=7, leading=10,
            textColor=GRAY
        )),
        Paragraph(next_step, ParagraphStyle(
            "ap_body", fontName="Helvetica-Bold", fontSize=10, leading=14,
            textColor=BLACK
        )),
    ]], colWidths=[36*mm, 134*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), OFF),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("TOPPADDING",    (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LINEAFTER",     (0, 0), (0, -1), 1, BLACK),
    ]))
    story.append(t)
    story.append(Spacer(1, 6*mm))
    return story


def _disclaimer_section(S):
    story = []
    story.append(GrayLine())
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "Score indicatif — Ce score est calculé automatiquement par le moteur KORD sur la base des données transmises. "
        "Il constitue une première estimation et sera ajusté après analyse approfondie par les experts KORD. "
        "Ne pas communiquer ce score au client avant validation.",
        ParagraphStyle(
            "disc", fontName="Helvetica", fontSize=7.5, leading=11,
            textColor=GRAY
        )
    ))
    return story


def _add_header_footer(canvas, doc):
    """Ajoute le header et le footer sur chaque page."""
    canvas.saveState()
    W, H = A4

    # Footer
    canvas.setFillColor(GRAY)
    canvas.setFont("Helvetica", 7)
    canvas.drawString(18*mm, 8*mm, "KORD — Pré-rapport d'audit opérationnel — Document de travail confidentiel")
    canvas.drawRightString(W - 18*mm, 8*mm, f"Page {doc.page}")

    canvas.restoreState()
