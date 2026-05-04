"""
KORD PDF — Rapport d'audit opérationnel
Style premium inspiré JLL — Format paysage A4
"""

import io
from datetime import datetime
from typing import Dict, Any

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    KeepTogether, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import Flowable

PW, PH = landscape(A4)
MW = PW - 32*mm   # Largeur utile

BLACK = colors.HexColor("#000000")
WHITE = colors.HexColor("#FFFFFF")
OFF   = colors.HexColor("#F4F3F0")
GRAY  = colors.HexColor("#6B6B6B")
LGRAY = colors.HexColor("#DEDEDE")
DGRAY = colors.HexColor("#1A1A1A")
MGRAY = colors.HexColor("#444444")


class VLine(Flowable):
    def __init__(self, height=20*mm, color=BLACK, thickness=0.5):
        super().__init__()
        self.height = height; self.color = color; self.thickness = thickness
    def wrap(self, aw, ah):
        return self.thickness + 2, self.height
    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(1, 0, 1, self.height)


class HLine(Flowable):
    def __init__(self, w=None, color=BLACK, thickness=0.5):
        super().__init__()
        self._w = w; self.color = color; self.thickness = thickness
    def wrap(self, aw, ah):
        self._w = self._w or aw
        return self._w, self.thickness + 3
    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 2, self._w, 2)


def PS(name, **kw):
    d = dict(fontName="Helvetica", fontSize=10, leading=15, textColor=DGRAY, spaceAfter=2)
    d.update(kw)
    return ParagraphStyle(name, **d)


def stat_box(number, label, sub="", bg=BLACK):
    """Encadré statistique style JLL — grand nombre + contexte."""
    rows = [[Paragraph(str(number), PS("sn", fontName="Helvetica-Bold",
             fontSize=36, leading=38, textColor=WHITE, alignment=TA_CENTER))]]
    if label:
        rows.append([Paragraph(label, PS("sl", fontName="Helvetica", fontSize=8.5,
                    leading=12, textColor=colors.HexColor("#BBBBBB"),
                    alignment=TA_CENTER))])
    if sub:
        rows.append([Paragraph(sub, PS("ss", fontName="Helvetica-Bold", fontSize=7,
                    leading=10, textColor=colors.HexColor("#888888"),
                    alignment=TA_CENTER, charSpace=0.5))])
    return Table(rows, colWidths=[38*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),bg),
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
        ("TOPPADDING",(0,0),(-1,-1),10),
        ("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
    ]))


def pull_quote(text, bg=OFF):
    """Citation/insight mis en valeur — style JLL pull quote."""
    return Table([[
        Table([[Spacer(3*mm, 1)]], colWidths=[3*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLACK),
        ])),
        Paragraph(text, PS("pq", fontName="Helvetica-Bold", fontSize=10.5,
                 leading=15, textColor=BLACK, alignment=TA_JUSTIFY)),
    ]], colWidths=[5*mm, MW-5*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),bg),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("RIGHTPADDING",(0,0),(-1,-1),12),
        ("TOPPADDING",(0,0),(-1,-1),12),
        ("BOTTOMPADDING",(0,0),(-1,-1),12),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))


def section_header(num, title, subtitle=""):
    items = [
        Spacer(1, 4*mm),
        Table([[
            Paragraph(num, PS("sh_n", fontName="Helvetica-Bold", fontSize=8,
                     textColor=WHITE, charSpace=1, alignment=TA_CENTER)),
            Spacer(4*mm, 1),
            Paragraph(title.upper(), PS("sh_t", fontName="Helvetica-Bold",
                     fontSize=10, textColor=WHITE, charSpace=1.5)),
        ]], colWidths=[14*mm, 4*mm, MW-18*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLACK),
            ("LEFTPADDING",(0,0),(-1,-1),10),
            ("RIGHTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),9),
            ("BOTTOMPADDING",(0,0),(-1,-1),9),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ])),
    ]
    if subtitle:
        items.append(Paragraph(subtitle, PS("sh_s", fontName="Helvetica",
                    fontSize=9, textColor=GRAY, spaceAfter=4)))
    items.append(Spacer(1, 4*mm))
    return items


def niveau_chip(niveau):
    cfg = {"BON": MGRAY, "MOYEN": GRAY, "CRITIQUE": BLACK}
    bg = cfg.get(niveau.upper(), GRAY)
    return Table([[Paragraph(niveau.upper(), PS("nc", fontName="Helvetica-Bold",
              fontSize=6.5, textColor=WHITE, alignment=TA_CENTER, charSpace=0.8))]],
        colWidths=[16*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),bg),
            ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))


def breadcrumb(sections, current, page):
    """Fil d'Ariane style JLL en bas de page."""
    items = []
    for s in sections:
        bold = s == current
        items.append(Paragraph(s, PS("bc", fontName="Helvetica-Bold" if bold else "Helvetica",
                    fontSize=7, textColor=BLACK if bold else GRAY,
                    charSpace=0.3)))
    row = [[i] for i in items]
    return None  # géré dans footer


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(GRAY)
    canvas.setFont("Helvetica", 7)
    # Breadcrumb
    sections = ["Résumé Exécutif", "Performance", "Anomalies", "Priorités", "Opportunités"]
    x = 16*mm
    for s in sections:
        canvas.setFont("Helvetica", 7)
        canvas.drawString(x, 6*mm, s)
        x += 52*mm
    # Ligne séparatrice
    canvas.setStrokeColor(LGRAY)
    canvas.setLineWidth(0.4)
    canvas.line(16*mm, 9.5*mm, PW-16*mm, 9.5*mm)
    # Page + copyright
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(GRAY)
    canvas.drawRightString(PW-16*mm, 6*mm, f"{doc.page}")
    canvas.drawString(16*mm, 3.5*mm, "KORD — Pré-rapport confidentiel — Ne pas transmettre sans validation")
    canvas.restoreState()


def generate_prereport_pdf(consolidated, recommendations, all_results,
                            client_name="Client", company_name="",
                            trimestre="", gauge_png=None, bar_png=None,
                            radar_png=None, evol_png=None) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
        leftMargin=16*mm, rightMargin=16*mm, topMargin=0, bottomMargin=16*mm)

    score = consolidated.get("score_total", 0)
    now = datetime.now()
    date_str = now.strftime("%d %B %Y").upper()
    trim_str = trimestre or f"T{(now.month-1)//3+1} {now.year}"
    nb = len(all_results)
    reco = recommendations
    story = []

    # ═══════════════════════════════════════════════════════
    # PAGE 1 — COUVERTURE PREMIUM
    # ═══════════════════════════════════════════════════════
    LW = 118*mm  # colonne gauche
    RW = MW - LW - 6*mm

    # Colonne gauche noire
    left = Table([
        [Paragraph("KORD", PS("ck", fontName="Helvetica-Bold", fontSize=72,
                  leading=68, textColor=WHITE))],
        [Spacer(1, 2*mm)],
        [Paragraph("AUDIT DE PERFORMANCE\nOPÉRATIONNELLE",
                  PS("cs", fontName="Helvetica", fontSize=9, leading=13,
                     textColor=colors.HexColor("#777777"), charSpace=1.5))],
        [Spacer(1, 14*mm)],
        [HLine(color=colors.HexColor("#333333"), thickness=0.5)],
        [Spacer(1, 8*mm)],
        [Paragraph((company_name or client_name).upper(),
                  PS("cco", fontName="Helvetica-Bold", fontSize=22,
                     leading=24, textColor=WHITE))],
        [Paragraph(client_name if company_name else " ",
                  PS("ccl", fontName="Helvetica", fontSize=11,
                     textColor=colors.HexColor("#AAAAAA")))],
        [Spacer(1, 10*mm)],
        [Paragraph(f"{trim_str}  ·  {date_str}",
                  PS("cm", fontName="Helvetica", fontSize=8,
                     textColor=colors.HexColor("#666666")))],
        [Paragraph(f"{nb} fichier{'s' if nb>1 else ''} analysé{'s' if nb>1 else ''} — Moteur KORD v1",
                  PS("cm2", fontName="Helvetica", fontSize=8,
                     textColor=colors.HexColor("#555555")))],
        [Spacer(1, 16*mm)],
        [Table([[
            Paragraph("⚠  DOCUMENT DE TRAVAIL — USAGE INTERNE",
                     PS("cw", fontName="Helvetica-Bold", fontSize=7,
                        textColor=colors.HexColor("#555555"), charSpace=0.5)),
        ]], colWidths=[LW-32*mm], style=TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),
        ]))],
    ], colWidths=[LW-32*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BLACK),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))

    left_block = Table([[left]], colWidths=[LW], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BLACK),
        ("LEFTPADDING",(0,0),(-1,-1),16*mm),
        ("RIGHTPADDING",(0,0),(-1,-1),8*mm),
        ("TOPPADDING",(0,0),(-1,-1),20*mm),
        ("BOTTOMPADDING",(0,0),(-1,-1),20*mm),
    ]))

    # Colonne droite claire — Score + piliers résumé
    piliers_data = consolidated.get("analyses", {})
    pilier_defs = [
        ("stock_cash",          "Stock & Cash",     30),
        ("transport_service",   "Transport",        20),
        ("achats_fournisseurs", "Achats",           20),
        ("marges_retours",      "Marges",           15),
        ("donnees_pilotage",    "Données",          15),
    ]

    right_rows = [
        [Paragraph("SCORE KORD", PS("rsl", fontName="Helvetica-Bold", fontSize=8,
                  textColor=GRAY, charSpace=2, alignment=TA_CENTER))],
        [Paragraph(str(score), PS("rsn", fontName="Helvetica-Bold", fontSize=80,
                  leading=78, textColor=BLACK, alignment=TA_CENTER))],
        [Paragraph("/100  —  INDICATIF", PS("rss", fontName="Helvetica", fontSize=8,
                  textColor=GRAY, charSpace=1, alignment=TA_CENTER))],
        [Spacer(1, 4*mm)],
        [HLine(color=LGRAY, thickness=0.5)],
        [Spacer(1, 4*mm)],
    ]

    for key, label, max_pts in pilier_defs:
        a = piliers_data.get(key, {})
        s = a.get("score", 0) if isinstance(a, dict) else 0
        pct = round(s/max_pts*100) if max_pts > 0 else 0
        niveau = "BON" if pct >= 70 else "MOYEN" if pct >= 45 else "CRITIQUE"
        bar_w = round(pct / 100 * (RW - 24*mm))
        right_rows.append([Table([[
            Paragraph(label, PS("rl", fontName="Helvetica", fontSize=8.5, textColor=DGRAY)),
            Paragraph(f"{s}/{max_pts}", PS("rs", fontName="Helvetica-Bold",
                     fontSize=9, textColor=BLACK, alignment=TA_RIGHT)),
        ]], colWidths=[RW-30*mm, 24*mm], style=TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))])
        # Mini barre de progression
        right_rows.append([Table([[
            Table([[]], colWidths=[max(bar_w, 2)], style=TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),BLACK if pct >= 70 else MGRAY if pct >= 45 else GRAY),
                ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
            ])),
            Table([[]], colWidths=[max(RW-24*mm-bar_w, 1)], style=TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),LGRAY),
                ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
            ])),
        ]], colWidths=[max(bar_w, 2), max(RW-24*mm-bar_w, 1)], style=TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]))])

    right = Table(right_rows, colWidths=[RW], style=TableStyle([
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))

    right_block = Table([[right]], colWidths=[RW+6*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),OFF),
        ("LEFTPADDING",(0,0),(-1,-1),10*mm),
        ("RIGHTPADDING",(0,0),(-1,-1),8*mm),
        ("TOPPADDING",(0,0),(-1,-1),20*mm),
        ("BOTTOMPADDING",(0,0),(-1,-1),20*mm),
    ]))

    story.append(Table([[left_block, right_block]],
        colWidths=[LW, RW+6*mm], style=TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ])))

    # ═══════════════════════════════════════════════════════
    # PAGE 2 — RÉSUMÉ EXÉCUTIF
    # ═══════════════════════════════════════════════════════
    story.append(PageBreak())
    story += section_header("01", "Résumé Exécutif",
                           "Situation globale et principaux constats")

    resume = reco.get("resume_executif", {})
    p1 = resume.get("paragraphe_1","") if isinstance(resume, dict) else str(resume)
    p2 = resume.get("paragraphe_2","") if isinstance(resume, dict) else ""
    p3 = resume.get("paragraphe_3","") if isinstance(resume, dict) else ""

    # Layout JLL : stats à gauche, texte à droite
    col_stats = 42*mm
    col_text  = MW - col_stats - 6*mm

    # Stats tirées du score
    alertes_count = len(consolidated.get("alertes", []))
    opps_count    = len(reco.get("opportunites_cachees", []))
    prios_count   = len(reco.get("priorites", []))

    stats_col = Table([
        [stat_box(f"{score}/100", "Score KORD", "INDICATIF")],
        [Spacer(1, 3*mm)],
        [stat_box(str(alertes_count), "Anomalie" + ("s" if alertes_count > 1 else ""),
                 "DÉTECTÉE" + ("S" if alertes_count > 1 else ""), bg=DGRAY)],
        [Spacer(1, 3*mm)],
        [stat_box(str(prios_count), "Priorité" + ("s" if prios_count > 1 else ""),
                 "D'ACTION", bg=MGRAY)],
    ], colWidths=[col_stats], style=TableStyle([
        ("LEFTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))

    text_col = Table([
        [Paragraph(p1, PS("rp1", fontName="Helvetica", fontSize=10, leading=16,
                  textColor=DGRAY, alignment=TA_JUSTIFY, spaceAfter=6))],
        [Paragraph(p2, PS("rp2", fontName="Helvetica", fontSize=10, leading=16,
                  textColor=DGRAY, alignment=TA_JUSTIFY, spaceAfter=6))],
        [Paragraph(p3, PS("rp3", fontName="Helvetica", fontSize=10, leading=16,
                  textColor=DGRAY, alignment=TA_JUSTIFY))],
    ], colWidths=[col_text], style=TableStyle([
        ("LEFTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))

    story.append(Table([[stats_col, Spacer(6*mm, 1), text_col]],
        colWidths=[col_stats, 6*mm, col_text], style=TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ])))

    story.append(Spacer(1, 5*mm))

    # Message dirigeant — pull quote style JLL
    msg = reco.get("message_dirigeant","")
    if msg:
        story.append(pull_quote(msg))
        story.append(Spacer(1, 4*mm))

    # Benchmark
    bench = reco.get("benchmark","")
    if bench:
        story.append(Table([[
            Paragraph("BENCHMARK SECTORIEL", PS("bml", fontName="Helvetica-Bold",
                     fontSize=7, textColor=GRAY, charSpace=1.5)),
            Paragraph(bench, PS("bmv", fontName="Helvetica", fontSize=9,
                     leading=13, textColor=DGRAY)),
        ]], colWidths=[40*mm, MW-40*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),OFF),
            ("LEFTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEAFTER",(0,0),(0,-1),1,LGRAY),
        ])))

    # ═══════════════════════════════════════════════════════
    # PAGE 3 — PERFORMANCE PAR PILIER
    # ═══════════════════════════════════════════════════════
    story.append(PageBreak())
    story += section_header("02", "Performance par Pilier",
                           "Analyse détaillée — Score KORD vs Benchmark marché")

    # Graphiques : barres (large) + radar (compact)
    chart_row = []
    if bar_png:
        try: chart_row.append(RLImage(io.BytesIO(bar_png), width=155*mm, height=75*mm))
        except: pass
    if radar_png:
        try: chart_row.append(RLImage(io.BytesIO(radar_png), width=MW-155*mm-4*mm, height=75*mm))
        except: pass
    if chart_row:
        widths = [155*mm, MW-155*mm-4*mm, 4*mm] if len(chart_row)==2 else [MW]
        if len(chart_row)==2:
            story.append(Table([chart_row + [Spacer(4*mm,1)]], colWidths=widths,
                style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),
                                  ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                  ("BACKGROUND",(0,0),(-1,-1),OFF),
                                  ("TOPPADDING",(0,0),(-1,-1),6),
                                  ("BOTTOMPADDING",(0,0),(-1,-1),6)])))
        else:
            story.append(Table([chart_row], colWidths=[MW],
                style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),
                                  ("BACKGROUND",(0,0),(-1,-1),OFF),
                                  ("TOPPADDING",(0,0),(-1,-1),6),
                                  ("BOTTOMPADDING",(0,0),(-1,-1),6)])))
        story.append(Paragraph("Source : Moteur d'analyse KORD — Données client. Benchmark : PME secteur distribution/logistique.",
                    PS("src", fontName="Helvetica", fontSize=7, textColor=GRAY,
                       spaceAfter=6)))
        story.append(Spacer(1, 3*mm))

    # Analyse piliers — 2 colonnes, style JLL
    analyse_piliers = reco.get("analyse_piliers", {})
    pilier_order = ["stock_cash","transport_service","achats_fournisseurs","marges_retours","donnees_pilotage"]
    col_w = (MW - 4*mm) / 2

    cells = []
    for key in pilier_order:
        dp = analyse_piliers.get(key, {})
        if not isinstance(dp, dict):
            continue
        titre   = dp.get("titre", key)
        sc      = dp.get("score", consolidated.get("analyses",{}).get(key,{}).get("score",0))
        mx      = dp.get("max", 20)
        niveau  = dp.get("niveau","moyen")
        analyse = dp.get("analyse","")
        chiffres= dp.get("chiffres_cles",[])
        risque  = dp.get("risque_principal","")

        pct = round(sc/mx*100) if mx > 0 else 0

        header_bg = BLACK if pct < 45 else DGRAY if pct < 70 else MGRAY

        inner = [
            [Table([[
                Paragraph(titre.upper(), PS("pt", fontName="Helvetica-Bold",
                         fontSize=8.5, textColor=WHITE, charSpace=0.5)),
                Paragraph(f"{sc}/{mx}", PS("ps", fontName="Helvetica-Bold",
                         fontSize=13, textColor=WHITE, alignment=TA_RIGHT)),
            ]], colWidths=[col_w-30*mm, 24*mm], style=TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),header_bg),
                ("LEFTPADDING",(0,0),(-1,-1),8),
                ("RIGHTPADDING",(0,0),(-1,-1),8),
                ("TOPPADDING",(0,0),(-1,-1),7),
                ("BOTTOMPADDING",(0,0),(-1,-1),7),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ]))],
        ]
        if analyse:
            inner.append([Paragraph(analyse, PS("pa", fontName="Helvetica",
                         fontSize=8.5, leading=13, textColor=DGRAY, alignment=TA_JUSTIFY))])
        for cf in chiffres[:2]:
            inner.append([Table([[
                Paragraph("▸", PS("cbul", fontName="Helvetica-Bold",
                         fontSize=10, textColor=BLACK)),
                Paragraph(cf, PS("cf", fontName="Helvetica", fontSize=8,
                         leading=12, textColor=GRAY)),
            ]], colWidths=[5*mm, col_w-22*mm], style=TableStyle([
                ("LEFTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),1),
                ("BOTTOMPADDING",(0,0),(-1,-1),1),("VALIGN",(0,0),(-1,-1),"TOP"),
            ]))])
        if risque:
            inner.append([Table([[
                Paragraph("RISQUE PRINCIPAL", PS("rql", fontName="Helvetica-Bold",
                         fontSize=7, textColor=GRAY, charSpace=0.8)),
                Paragraph(risque, PS("rqv", fontName="Helvetica-Bold", fontSize=8.5,
                         textColor=BLACK, alignment=TA_RIGHT)),
            ]], colWidths=[36*mm, col_w-44*mm], style=TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),OFF),
                ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
                ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ]))])

        cell = Table([[Table(inner, colWidths=[col_w-12],
            style=TableStyle([
                ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
            ]))]], colWidths=[col_w], style=TableStyle([
            ("BOX",(0,0),(-1,-1),0.5,LGRAY),
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),4),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ]))
        cells.append(cell)

    rows_2 = []
    for i in range(0, len(cells), 2):
        r = [cells[i], cells[i+1] if i+1 < len(cells) else Spacer(col_w,1)]
        rows_2.append(r)
    if rows_2:
        story.append(Table(rows_2, colWidths=[col_w, col_w], style=TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ])))

    # ═══════════════════════════════════════════════════════
    # PAGE 4 — ANOMALIES + CROISEMENTS
    # ═══════════════════════════════════════════════════════
    story.append(PageBreak())

    # Graphique évolution
    if evol_png:
        story += section_header("03", "Trajectoire de Performance",
                               "Score KORD dans le temps vs moyenne sectorielle")
        try:
            story.append(RLImage(io.BytesIO(evol_png), width=MW, height=62*mm))
            story.append(Paragraph("Source : Moteur KORD — Premier audit. La courbe d'évolution sera enrichie à chaque trimestre.",
                        PS("src", fontName="Helvetica", fontSize=7, textColor=GRAY, spaceAfter=5)))
        except: pass
        story.append(Spacer(1, 3*mm))

    # Croisements
    croisements = reco.get("croisements_cles", [])
    if croisements:
        story += section_header("04", "Croisements Clés",
                               "Ce que révèle l'analyse croisée de l'ensemble de vos données")
        cr_rows = [[
            Paragraph("FICHIERS CROISÉS", PS("th", fontName="Helvetica-Bold",
                     fontSize=7, textColor=WHITE, charSpace=1)),
            Paragraph("OBSERVATION", PS("th", fontName="Helvetica-Bold",
                     fontSize=7, textColor=WHITE, charSpace=1)),
            Paragraph("IMPACT BUSINESS", PS("th", fontName="Helvetica-Bold",
                     fontSize=7, textColor=WHITE, charSpace=1, alignment=TA_RIGHT)),
        ]]
        for cr in croisements[:4]:
            if not isinstance(cr, dict): continue
            cr_rows.append([
                Paragraph(" × ".join(cr.get("fichiers",[])), PS("crf", fontName="Helvetica-Bold",
                         fontSize=8.5, textColor=DGRAY)),
                Paragraph(cr.get("observation",""), PS("cro", fontName="Helvetica",
                         fontSize=8.5, leading=13, textColor=DGRAY)),
                Paragraph(cr.get("impact",""), PS("cri", fontName="Helvetica-Bold",
                         fontSize=9, textColor=BLACK, alignment=TA_RIGHT)),
            ])
        t = Table(cr_rows, colWidths=[48*mm, MW-48*mm-60*mm, 60*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),BLACK),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,OFF]),
            ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
            ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEBELOW",(0,0),(-1,-1),0.4,LGRAY),
            ("LINEAFTER",(0,0),(1,-1),0.4,LGRAY),
        ]))
        story.append(t)
        story.append(Spacer(1, 5*mm))

    # Anomalies
    anomalies = reco.get("anomalies", consolidated.get("alertes", []))
    if anomalies:
        story += section_header("05", "Anomalies Détectées",
                               "Classées par niveau d'urgence — Impact business estimé")
        u_order = {"CRITIQUE":0,"MOYEN":1,"FAIBLE":2}
        if anomalies and isinstance(anomalies[0], dict) and "urgence" in anomalies[0]:
            anomalies = sorted(anomalies, key=lambda x: u_order.get(x.get("urgence","FAIBLE"),2))

        a_rows = [[
            Paragraph("URGENCE", PS("th", fontName="Helvetica-Bold", fontSize=7,
                     textColor=WHITE, charSpace=1, alignment=TA_CENTER)),
            Paragraph("ANOMALIE DÉTECTÉE", PS("th", fontName="Helvetica-Bold",
                     fontSize=7, textColor=WHITE, charSpace=1)),
            Paragraph("IMPACT BUSINESS", PS("th", fontName="Helvetica-Bold",
                     fontSize=7, textColor=WHITE, charSpace=1)),
            Paragraph("IMPACT €", PS("th", fontName="Helvetica-Bold", fontSize=7,
                     textColor=WHITE, charSpace=1, alignment=TA_RIGHT)),
        ]]

        for al in anomalies[:10]:
            if not isinstance(al, dict): continue
            urg = al.get("urgence","MOYEN")
            u_bg = BLACK if urg=="CRITIQUE" else MGRAY if urg=="MOYEN" else GRAY
            titre = al.get("titre", al.get("message",""))
            detect = al.get("detection", al.get("message",""))
            impact = al.get("impact_business","")
            imp_fin = al.get("impact_financier","")

            a_rows.append([
                Table([[Paragraph(urg, PS("ub", fontName="Helvetica-Bold", fontSize=6.5,
                        textColor=WHITE, alignment=TA_CENTER, charSpace=0.5))]],
                    colWidths=[20*mm], style=TableStyle([
                        ("BACKGROUND",(0,0),(-1,-1),u_bg),
                        ("TOPPADDING",(0,0),(-1,-1),3),
                        ("BOTTOMPADDING",(0,0),(-1,-1),3),
                    ])),
                Paragraph(f"<b>{titre}</b><br/>{detect}",
                         PS("ad", fontName="Helvetica", fontSize=8.5, leading=12, textColor=DGRAY)),
                Paragraph(impact, PS("ai", fontName="Helvetica", fontSize=8.5,
                         leading=12, textColor=DGRAY)),
                Paragraph(imp_fin, PS("ae", fontName="Helvetica-Bold", fontSize=9,
                         textColor=BLACK, alignment=TA_RIGHT)),
            ])

        ta = Table(a_rows, colWidths=[22*mm, MW-22*mm-68*mm-42*mm, 68*mm, 42*mm])
        ta.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),BLACK),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,OFF]),
            ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
            ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEBELOW",(0,0),(-1,-1),0.3,LGRAY),
            ("LINEAFTER",(0,0),(2,-1),0.3,LGRAY),
        ]))
        story.append(ta)

    # ═══════════════════════════════════════════════════════
    # PAGE 5 — PRIORITÉS D'ACTION
    # ═══════════════════════════════════════════════════════
    story.append(PageBreak())
    story += section_header("06", "Priorités d'Action",
                           "Top actions classées par impact financier — Ce qui rapporte le plus vite")

    priorites = reco.get("priorites", [])
    for p in priorites[:5]:
        if not isinstance(p, dict): continue
        rang = p.get("rang","")
        qw   = p.get("quick_win", False)

        story.append(KeepTogether([
            Table([[
                # Numéro grand
                Table([[
                    Paragraph(f"0{rang}", PS("pnum", fontName="Helvetica-Bold",
                             fontSize=28, leading=30, textColor=LGRAY, alignment=TA_CENTER)),
                    Paragraph("⚡ QW" if qw else "", PS("qwl", fontName="Helvetica-Bold",
                             fontSize=7, textColor=WHITE, alignment=TA_CENTER, charSpace=0.5)),
                ]], colWidths=[18*mm], style=TableStyle([
                    ("BACKGROUND",(0,0),(0,-1),OFF),
                    ("BACKGROUND",(1,0),(-1,-1),DGRAY if qw else OFF),
                    ("LEFTPADDING",(0,0),(-1,-1),2),("TOPPADDING",(0,0),(-1,-1),4),
                    ("BOTTOMPADDING",(0,0),(-1,-1),4),("ALIGN",(0,0),(-1,-1),"CENTER"),
                ])),
                # Contenu
                Table([
                    # Header
                    [Table([[
                        Paragraph(p.get("titre","").upper(), PS("ptt", fontName="Helvetica-Bold",
                                 fontSize=10, textColor=WHITE, charSpace=0.3)),
                        Paragraph(p.get("delai",""), PS("ptd", fontName="Helvetica",
                                 fontSize=8, textColor=colors.HexColor("#AAAAAA"),
                                 alignment=TA_RIGHT)),
                    ]], colWidths=[MW-18*mm-50*mm, 44*mm], style=TableStyle([
                        ("BACKGROUND",(0,0),(-1,-1),BLACK),
                        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
                        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
                        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ]))],
                    # Corps 2 colonnes
                    [Table([[
                        Table([
                            [Paragraph(f"<b>Problème :</b> {p.get('probleme','')}",
                                      PS("pp", fontName="Helvetica", fontSize=9, leading=13, textColor=DGRAY))],
                            [Paragraph(f"<b>Action :</b> {p.get('action','')}",
                                      PS("pa2", fontName="Helvetica", fontSize=9, leading=13, textColor=DGRAY))],
                        ], colWidths=[(MW-18*mm-4*mm)//2-2], style=TableStyle([
                            ("LEFTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),2),
                            ("BOTTOMPADDING",(0,0),(-1,-1),2),
                        ])),
                        # Gain + impact
                        Table([
                            [Paragraph(f"<b>Impact attendu</b>", PS("pil", fontName="Helvetica-Bold",
                                      fontSize=8, textColor=GRAY, charSpace=0.5))],
                            [Paragraph(p.get("impact_attendu",""), PS("piv",
                                      fontName="Helvetica-Bold", fontSize=9.5,
                                      leading=13, textColor=BLACK))],
                            [Spacer(1, 3*mm)],
                            [Paragraph(f"Gain estimé : {p.get('gain_potentiel','')}",
                                      PS("pg", fontName="Helvetica-Bold", fontSize=9,
                                         textColor=WHITE))],
                            [Paragraph(f"Complexité : {p.get('complexite','')}",
                                      PS("pc", fontName="Helvetica", fontSize=8, textColor=GRAY))],
                        ], colWidths=[(MW-18*mm-4*mm)//2-2], style=TableStyle([
                            ("BACKGROUND",(3,0),(-1,-1),OFF),
                            ("BACKGROUND",(0,3),(-1,3),DGRAY),
                            ("LEFTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),2),
                            ("BOTTOMPADDING",(0,0),(-1,-1),2),
                        ])),
                    ]], colWidths=[(MW-18*mm-4*mm)//2, (MW-18*mm-4*mm)//2], style=TableStyle([
                        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
                        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
                        ("VALIGN",(0,0),(-1,-1),"TOP"),
                        ("LINEAFTER",(0,0),(0,-1),0.4,LGRAY),
                    ]))],
                ], colWidths=[MW-18*mm], style=TableStyle([
                    ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                    ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
                ])),
            ]], colWidths=[18*mm, MW-18*mm], style=TableStyle([
                ("BOX",(0,0),(-1,-1),0.5,LGRAY),
                ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
            ])),
            Spacer(1, 4*mm),
        ]))

    # ═══════════════════════════════════════════════════════
    # PAGE 6 — OPPORTUNITÉS + SUITE
    # ═══════════════════════════════════════════════════════
    story.append(PageBreak())

    # Opportunités cachées
    opps = reco.get("opportunites_cachees", [])
    if opps:
        story += section_header("07", "Opportunités Cachées",
                               "Leviers additionnels identifiés dans l'analyse des données")
        for opp in opps:
            if not isinstance(opp, dict): continue
            story.append(Table([[
                Table([[
                    Paragraph(opp.get("titre",""), PS("ot", fontName="Helvetica-Bold",
                             fontSize=10, textColor=WHITE)),
                    Paragraph(opp.get("gain_estime",""), PS("og", fontName="Helvetica-Bold",
                             fontSize=12, textColor=WHITE, alignment=TA_RIGHT)),
                ]], colWidths=[MW-60*mm, 58*mm], style=TableStyle([
                    ("BACKGROUND",(0,0),(-1,-1),DGRAY),
                    ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
                    ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ])),
            ]], colWidths=[MW], style=TableStyle([
                ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),2),
            ])))
            story.append(Table([[
                Paragraph(opp.get("description",""), PS("od", fontName="Helvetica",
                         fontSize=9, leading=13, textColor=DGRAY, alignment=TA_JUSTIFY)),
            ]], colWidths=[MW], style=TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),OFF),
                ("LEFTPADDING",(0,0),(-1,-1),10),("TOPPADDING",(0,0),(-1,-1),7),
                ("BOTTOMPADDING",(0,0),(-1,-1),7),
            ])))
            story.append(Spacer(1, 3*mm))
        story.append(Spacer(1, 3*mm))

    # Vigilance + Questions — 2 colonnes
    col_w2 = (MW - 5*mm) / 2
    vigilance = reco.get("points_vigilance",[])
    questions = reco.get("questions_restitution",[])

    if vigilance or questions:
        vcol = []
        if vigilance:
            vcol += section_header("08", "Points de Vigilance", "")
            for i, pt in enumerate(vigilance, 1):
                vcol.append(Table([[
                    Paragraph(f"{i:02d}", PS("vn", fontName="Helvetica-Bold",
                             fontSize=9, textColor=LGRAY, alignment=TA_CENTER)),
                    Paragraph(pt, PS("vt", fontName="Helvetica", fontSize=9,
                             leading=13, textColor=DGRAY)),
                ]], colWidths=[12*mm, col_w2-18*mm], style=TableStyle([
                    ("LEFTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),4),
                    ("BOTTOMPADDING",(0,0),(-1,-1),4),("LINEBELOW",(0,0),(-1,-1),0.3,LGRAY),
                    ("VALIGN",(0,0),(-1,-1),"TOP"),
                ])))

        qcol = []
        if questions:
            qcol += section_header("09", "Questions — Restitution Client", "")
            for i, q in enumerate(questions, 1):
                qcol.append(Table([[
                    Paragraph(f"Q{i}", PS("qn", fontName="Helvetica-Bold",
                             fontSize=9, textColor=WHITE, alignment=TA_CENTER)),
                    Paragraph(q, PS("qt", fontName="Helvetica", fontSize=9,
                             leading=13, textColor=DGRAY)),
                ]], colWidths=[12*mm, col_w2-18*mm], style=TableStyle([
                    ("BACKGROUND",(0,0),(0,-1),BLACK),
                    ("LEFTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),5),
                    ("BOTTOMPADDING",(0,0),(-1,-1),5),("LINEBELOW",(0,0),(-1,-1),0.3,LGRAY),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ])))

        max_l = max(len(vcol), len(qcol))
        while len(vcol) < max_l: vcol.append(Spacer(1,4))
        while len(qcol) < max_l: qcol.append(Spacer(1,4))

        story.append(Table([[v, Spacer(5*mm,1), q] for v, q in zip(vcol, qcol)],
            colWidths=[col_w2, 5*mm, col_w2], style=TableStyle([
                ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
            ])))
        story.append(Spacer(1, 5*mm))

    # Prochaine étape
    next_step = reco.get("prochaine_etape","")
    if next_step:
        story += section_header("10", "Prochaine Étape", "")
        story.append(pull_quote(next_step, bg=OFF))
        story.append(Spacer(1, 6*mm))

    # Espace annotations + footer
    story.append(HLine(color=LGRAY, thickness=0.3))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("ANNEXE — ESPACE ANNOTATIONS EXPERTS KORD",
                PS("an", fontName="Helvetica-Bold", fontSize=7, textColor=GRAY, charSpace=1)))
    story.append(Spacer(1, 18*mm))
    story.append(HLine(color=LGRAY, thickness=0.3))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f"KORD — Audit de performance opérationnelle — {trim_str} — Confidentiel — Score indicatif avant validation experts",
        PS("ft", fontName="Helvetica", fontSize=7, textColor=GRAY)))

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buf.seek(0)
    return buf.read()
