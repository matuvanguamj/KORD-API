"""
KORD PDF — Pré-rapport d'audit opérationnel
Document de travail interne — Ne pas envoyer au client sans validation
"""

import io
from datetime import datetime
from typing import Dict, Any

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    KeepTogether, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable

W, H = A4
BLACK  = colors.HexColor("#000000")
WHITE  = colors.HexColor("#FFFFFF")
OFF    = colors.HexColor("#F4F3F0")
GRAY   = colors.HexColor("#6B6B6B")
LGRAY  = colors.HexColor("#E8E8E8")
DGRAY  = colors.HexColor("#1A1A1A")
MGRAY  = colors.HexColor("#333333")


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


def S(name, **kw):
    defaults = dict(fontName="Helvetica", fontSize=10, leading=15, textColor=DGRAY, spaceAfter=4)
    defaults.update(kw)
    return ParagraphStyle(name, **defaults)


def generate_prereport_pdf(consolidated, recommendations, all_results,
                            client_name="Client", company_name="",
                            trimestre="", gauge_png=None, bar_png=None, radar_png=None) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm, topMargin=0, bottomMargin=14*mm)

    score = consolidated.get("score_total", 0)
    now = datetime.now()
    date_str = now.strftime("%d %B %Y").upper()
    trim_str = trimestre or f"T{(now.month-1)//3+1} {now.year}"
    nb = len(all_results)

    story = []

    # ══════════════════════════════════════════════
    # PAGE DE COUVERTURE
    # ══════════════════════════════════════════════
    # Bloc noir couverture
    cover_inner = []

    # KORD + sous-titre
    cover_inner.append(Table([[
        Paragraph("KORD", S("cov_kord", fontName="Helvetica-Bold", fontSize=52,
                  leading=52, textColor=WHITE, spaceAfter=0)),
    ]], colWidths=[170*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BLACK),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ])))

    cover_inner.append(Table([[
        Paragraph("AUDIT DE PERFORMANCE OPÉRATIONNELLE", S("cov_sub",
                  fontName="Helvetica", fontSize=9, leading=12,
                  textColor=colors.HexColor("#888888"), charSpace=2)),
    ]], colWidths=[170*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BLACK),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),28),
    ])))

    # Nom entreprise
    cover_inner.append(Table([[
        Paragraph((company_name or client_name).upper(), S("cov_company",
                  fontName="Helvetica-Bold", fontSize=24, leading=26,
                  textColor=WHITE, spaceAfter=4)),
    ]], colWidths=[170*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BLACK),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
    ])))

    if company_name and client_name:
        cover_inner.append(Table([[
            Paragraph(client_name, S("cov_client", fontName="Helvetica",
                      fontSize=12, leading=15, textColor=colors.HexColor("#BBBBBB"))),
        ]], colWidths=[170*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLACK),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ])))

    cover_inner.append(Table([[
        Paragraph(f"{trim_str}  ·  {date_str}  ·  {nb} fichier{'s' if nb>1 else ''} analysé{'s' if nb>1 else ''}",
                  S("cov_meta", fontName="Helvetica", fontSize=8, leading=11,
                    textColor=colors.HexColor("#666666"))),
    ]], colWidths=[170*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BLACK),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),24),
    ])))

    # Score + jauge
    score_color = DGRAY
    score_row = [[
        Table([[
            [Paragraph(str(score), S("cov_score", fontName="Helvetica-Bold",
                      fontSize=64, leading=64, textColor=WHITE, alignment=TA_CENTER))],
            [Paragraph("/100 — SCORE KORD INDICATIF", S("cov_score_lbl",
                      fontName="Helvetica", fontSize=7, leading=10,
                      textColor=colors.HexColor("#777777"), alignment=TA_CENTER,
                      charSpace=1))],
        ]], colWidths=[80*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),score_color),
            ("TOPPADDING",(0,0),(-1,-1),14),
            ("BOTTOMPADDING",(0,0),(-1,-1),14),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ])),
    ]]

    # Ajout jauge si disponible
    if gauge_png:
        try:
            gauge_img = RLImage(io.BytesIO(gauge_png), width=80*mm, height=40*mm)
            score_row[0].append(gauge_img)
        except:
            pass

    cover_inner.append(Table(score_row,
        colWidths=[80*mm, 86*mm] if gauge_png else [170*mm],
        style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLACK),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ])))

    # Assembler le bloc noir
    cover_table = Table([[item] for item in cover_inner],
        colWidths=[170*mm],
        style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLACK),
            ("LEFTPADDING",(0,0),(-1,-1),16*mm),
            ("RIGHTPADDING",(0,0),(-1,-1),16*mm),
            ("TOPPADDING",(0,0),(0,0),22*mm),
            ("TOPPADDING",(0,1),(-1,-1),0),
            ("BOTTOMPADDING",(0,-1),(-1,-1),22*mm),
            ("BOTTOMPADDING",(0,0),(-1,-2),0),
        ]))
    story.append(cover_table)
    story.append(Spacer(1, 8*mm))

    # Note document de travail
    story.append(Table([[
        Paragraph("⚠  DOCUMENT DE TRAVAIL — USAGE INTERNE KORD UNIQUEMENT",
                  S("note_tag", fontName="Helvetica-Bold", fontSize=7,
                    textColor=GRAY, charSpace=0.5)),
        Paragraph("Ce pré-rapport est généré automatiquement. Ne pas transmettre au client avant validation par l'équipe KORD.",
                  S("note_body", fontName="Helvetica", fontSize=7.5,
                    leading=11, textColor=GRAY, alignment=TA_RIGHT)),
    ]], colWidths=[90*mm, 80*mm], style=TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),OFF),
        ("LEFTPADDING",(0,0),(-1,-1),8),
        ("RIGHTPADDING",(0,0),(-1,-1),8),
        ("TOPPADDING",(0,0),(-1,-1),7),
        ("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ])))
    story.append(Spacer(1, 8*mm))

    # ══════════════════════════════════════════════
    # SECTION 01 — MESSAGE AU DIRIGEANT
    # ══════════════════════════════════════════════
    story += _section_header("01", "MESSAGE AU DIRIGEANT")
    msg = recommendations.get("message_dirigeant", "")
    if msg:
        story.append(Table([[
            Paragraph(msg, S("msg_dir", fontName="Helvetica", fontSize=10,
                     leading=16, textColor=DGRAY, leftIndent=6))
        ]], colWidths=[170*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),OFF),
            ("LEFTPADDING",(0,0),(-1,-1),14),
            ("RIGHTPADDING",(0,0),(-1,-1),14),
            ("TOPPADDING",(0,0),(-1,-1),12),
            ("BOTTOMPADDING",(0,0),(-1,-1),12),
            ("LINEBEFORE",(0,0),(0,-1),3,BLACK),
        ])))
        story.append(Spacer(1, 6*mm))

    # ══════════════════════════════════════════════
    # SECTION 02 — RÉSUMÉ EXÉCUTIF
    # ══════════════════════════════════════════════
    story += _section_header("02", "RÉSUMÉ EXÉCUTIF")
    summary = recommendations.get("resume_executif", "")
    if summary:
        story.append(Paragraph(summary, S("body", fontName="Helvetica",
                     fontSize=10, leading=16, textColor=DGRAY)))
        story.append(Spacer(1, 6*mm))

    # Benchmark
    benchmark = recommendations.get("benchmark", "")
    if benchmark:
        story.append(Table([[
            Paragraph("BENCHMARK MARCHÉ", S("bench_lbl", fontName="Helvetica-Bold",
                      fontSize=8, textColor=GRAY, charSpace=1)),
            Paragraph(benchmark, S("bench_val", fontName="Helvetica",
                      fontSize=9, leading=14, textColor=DGRAY)),
        ]], colWidths=[38*mm, 132*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),OFF),
            ("LEFTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LINEAFTER",(0,0),(0,-1),1,LGRAY),
        ])))
        story.append(Spacer(1, 8*mm))

    # ══════════════════════════════════════════════
    # SECTION 03 — SCORES PAR PILIER + GRAPHIQUES
    # ══════════════════════════════════════════════
    story += _section_header("03", "PERFORMANCE PAR PILIER")

    analyses = consolidated.get("analyses", {})
    pilier_meta = {
        "stock_cash":          ("Stock et Cash immobilisé",    30, "Dormance, sur-stockage, cash bloqué dans les références non-vendantes."),
        "transport_service":   ("Transport et Taux de service", 20, "Délais, avaries, taux de service et coûts logistiques."),
        "achats_fournisseurs": ("Achats et Fournisseurs",       20, "Négociation, délais fournisseurs, concentration et dépendance."),
        "marges_retours":      ("Marges et Retours",            15, "Rentabilité des produits, taux de retour et coût de traitement."),
        "donnees_pilotage":    ("Données et Pilotage",          15, "Qualité des données, couverture et capacité de décision."),
    }

    rows = [[
        Paragraph("PILIER", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=GRAY, charSpace=1)),
        Paragraph("SCORE", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=GRAY, charSpace=1, alignment=TA_CENTER)),
        Paragraph("NIVEAU", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=GRAY, charSpace=1, alignment=TA_CENTER)),
        Paragraph("ANALYSE", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=GRAY, charSpace=1)),
    ]]

    synth = recommendations.get("synthese_piliers", {})

    for key, (label, max_pts, desc) in pilier_meta.items():
        a = analyses.get(key, {})
        s = a.get("score", 0)
        pct = round(s/max_pts*100) if max_pts > 0 else 0
        niveau = "BON" if pct >= 70 else ("MOYEN" if pct >= 45 else "CRITIQUE")
        bg = OFF if rows.index(rows[-1]) % 2 == 0 else colors.HexColor("#FFFFFF")

        # Commentaire IA si disponible
        commentaire = synth.get(key, {}).get("commentaire", desc)

        score_txt = Paragraph(f"<b>{s}/{max_pts}</b>", S("score_cell",
                   fontName="Helvetica-Bold", fontSize=14, leading=16,
                   textColor=BLACK, alignment=TA_CENTER))
        niveau_color = BLACK if niveau != "BON" else GRAY
        niveau_txt = Paragraph(niveau, S("niveau_cell", fontName="Helvetica-Bold",
                   fontSize=8, charSpace=0.5, textColor=niveau_color, alignment=TA_CENTER))

        rows.append([
            Table([[
                [Paragraph(label, S("p_label", fontName="Helvetica-Bold", fontSize=9, textColor=BLACK))],
            ]], colWidths=[72*mm], style=TableStyle([
                ("TOPPADDING",(0,0),(-1,-1),0),
                ("BOTTOMPADDING",(0,0),(-1,-1),2),
            ])),
            score_txt,
            niveau_txt,
            Paragraph(commentaire, S("p_comment", fontName="Helvetica", fontSize=8.5,
                     leading=13, textColor=DGRAY)),
        ])

    t = Table(rows, colWidths=[72*mm, 20*mm, 22*mm, 56*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),DGRAY),
        ("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE, OFF]),
        ("LEFTPADDING",(0,0),(-1,-1),8),
        ("RIGHTPADDING",(0,0),(-1,-1),8),
        ("TOPPADDING",(0,0),(-1,-1),8),
        ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LINEBELOW",(0,0),(-1,-1),0.4,LGRAY),
        ("BOX",(0,0),(-1,-1),0.5,LGRAY),
    ]))
    story.append(t)
    story.append(Spacer(1, 8*mm))

    # Graphiques
    if bar_png or radar_png:
        chart_cols = []
        if bar_png:
            try:
                chart_cols.append(RLImage(io.BytesIO(bar_png), width=95*mm, height=55*mm))
            except:
                pass
        if radar_png:
            try:
                chart_cols.append(RLImage(io.BytesIO(radar_png), width=65*mm, height=55*mm))
            except:
                pass
        if chart_cols:
            widths = [95*mm, 75*mm] if len(chart_cols) == 2 else [170*mm]
            story.append(Table([chart_cols], colWidths=widths, style=TableStyle([
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("BACKGROUND",(0,0),(-1,-1),OFF),
                ("TOPPADDING",(0,0),(-1,-1),10),
                ("BOTTOMPADDING",(0,0),(-1,-1),10),
            ])))
            story.append(Spacer(1, 8*mm))

    # ══════════════════════════════════════════════
    # SECTION 04 — PRIORITÉS D'ACTION
    # ══════════════════════════════════════════════
    story.append(PageBreak())
    story += _section_header("04", "PRIORITÉS D'ACTION — CE QUI COÛTE DE L'ARGENT")

    for p in recommendations.get("priorites", [])[:5]:
        rang = p.get("rang", "")
        titre = p.get("titre", "")
        probleme = p.get("probleme", "")
        impact = p.get("impact_financier", "")
        action = p.get("action", "")
        exemple = p.get("exemple", "")
        gain = p.get("gain_potentiel", "")
        delai = p.get("delai", "")
        complexite = p.get("complexite", "")

        block = KeepTogether([
            Table([[
                # Numéro
                Paragraph(f"0{rang}", S("rang_num", fontName="Helvetica-Bold",
                         fontSize=20, leading=22, textColor=LGRAY)),
                # Contenu
                Table([
                    [Paragraph(titre.upper(), S("p_titre", fontName="Helvetica-Bold",
                              fontSize=11, leading=14, textColor=BLACK, charSpace=0.3))],
                    [Spacer(1, 2*mm)],
                    [Paragraph(f"<b>Problème détecté :</b> {probleme}",
                              S("p_body", fontName="Helvetica", fontSize=9, leading=14, textColor=DGRAY))],
                    [Paragraph(f"<b>Impact financier :</b> {impact}",
                              S("p_impact", fontName="Helvetica-Bold", fontSize=9.5,
                                leading=14, textColor=BLACK))],
                    [Paragraph(f"<b>Exemple :</b> {exemple}",
                              S("p_ex", fontName="Helvetica", fontSize=9, leading=13,
                                textColor=GRAY, leftIndent=4))],
                    [Paragraph(f"<b>Action recommandée :</b> {action}",
                              S("p_action", fontName="Helvetica", fontSize=9, leading=13, textColor=DGRAY))],
                    [Table([[
                        Paragraph(f"Gain potentiel : {gain}", S("gain_txt",
                                 fontName="Helvetica-Bold", fontSize=9, textColor=BLACK)),
                        Paragraph(f"{delai}  ·  {complexite}", S("delai_txt",
                                 fontName="Helvetica", fontSize=8, textColor=GRAY,
                                 alignment=TA_RIGHT)),
                    ]], colWidths=[90*mm, 60*mm], style=TableStyle([
                        ("BACKGROUND",(0,0),(-1,-1),OFF),
                        ("LEFTPADDING",(0,0),(-1,-1),8),
                        ("TOPPADDING",(0,0),(-1,-1),5),
                        ("BOTTOMPADDING",(0,0),(-1,-1),5),
                    ]))],
                ], colWidths=[150*mm], style=TableStyle([
                    ("LEFTPADDING",(0,0),(-1,-1),0),
                    ("RIGHTPADDING",(0,0),(-1,-1),0),
                    ("TOPPADDING",(0,0),(-1,-1),1),
                    ("BOTTOMPADDING",(0,0),(-1,-1),1),
                ])),
            ]], colWidths=[18*mm, 152*mm], style=TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),WHITE),
                ("LEFTPADDING",(0,0),(-1,-1),10),
                ("RIGHTPADDING",(0,0),(-1,-1),10),
                ("TOPPADDING",(0,0),(-1,-1),12),
                ("BOTTOMPADDING",(0,0),(-1,-1),12),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("BOX",(0,0),(-1,-1),0.5,LGRAY),
                ("LINEBEFORE",(0,0),(0,-1),3,BLACK),
            ])),
            Spacer(1, 4*mm),
        ])
        story.append(block)

    story.append(Spacer(1, 6*mm))

    # ══════════════════════════════════════════════
    # SECTION 05 — OPPORTUNITÉS CACHÉES
    # ══════════════════════════════════════════════
    opps = recommendations.get("opportunites_cachees", [])
    if opps:
        story += _section_header("05", "OPPORTUNITÉS CACHÉES DANS VOS DONNÉES")
        for opp in opps:
            story.append(KeepTogether([
                Table([[
                    Paragraph(opp.get("titre",""), S("opp_titre", fontName="Helvetica-Bold",
                             fontSize=10, textColor=BLACK)),
                    Paragraph(opp.get("gain_estime",""), S("opp_gain",
                             fontName="Helvetica-Bold", fontSize=10, textColor=BLACK,
                             alignment=TA_RIGHT)),
                ]], colWidths=[120*mm, 50*mm], style=TableStyle([
                    ("BACKGROUND",(0,0),(-1,-1),DGRAY),
                    ("LEFTPADDING",(0,0),(-1,-1),12),
                    ("RIGHTPADDING",(0,0),(-1,-1),12),
                    ("TOPPADDING",(0,0),(-1,-1),8),
                    ("BOTTOMPADDING",(0,0),(-1,-1),8),
                    ("TEXTCOLOR",(0,0),(-1,-1),WHITE),
                ])),
                Table([[
                    Paragraph(opp.get("description",""), S("opp_desc", fontName="Helvetica",
                             fontSize=9, leading=14, textColor=DGRAY)),
                ]], colWidths=[170*mm], style=TableStyle([
                    ("BACKGROUND",(0,0),(-1,-1),OFF),
                    ("LEFTPADDING",(0,0),(-1,-1),12),
                    ("TOPPADDING",(0,0),(-1,-1),8),
                    ("BOTTOMPADDING",(0,0),(-1,-1),8),
                ])),
                Spacer(1, 4*mm),
            ]))
        story.append(Spacer(1, 4*mm))

    # ══════════════════════════════════════════════
    # SECTION 06 — ANOMALIES DÉTECTÉES
    # ══════════════════════════════════════════════
    alertes = consolidated.get("alertes", [])
    if alertes:
        story += _section_header("06", "ANOMALIES DÉTECTÉES DANS LES DONNÉES")
        pilier_labels = {
            "stock_cash":"Stock","transport_service":"Transport",
            "achats_fournisseurs":"Achats","marges_retours":"Marges","donnees_pilotage":"Données"
        }
        rows_a = [[
            Paragraph("PILIER", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=GRAY, charSpace=1)),
            Paragraph("ANOMALIE DÉTECTÉE", S("th", fontName="Helvetica-Bold", fontSize=8, textColor=GRAY, charSpace=1)),
        ]]
        for al in alertes[:12]:
            p = pilier_labels.get(al.get("pilier",""), al.get("pilier","").upper())
            rows_a.append([
                Paragraph(p, S("al_p", fontName="Helvetica-Bold", fontSize=8.5, textColor=GRAY)),
                Paragraph(al.get("message",""), S("al_m", fontName="Helvetica", fontSize=9, leading=13, textColor=DGRAY)),
            ])
        ta = Table(rows_a, colWidths=[28*mm, 142*mm])
        ta.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),DGRAY),
            ("TEXTCOLOR",(0,0),(-1,0),WHITE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,OFF]),
            ("LEFTPADDING",(0,0),(-1,-1),8),
            ("TOPPADDING",(0,0),(-1,-1),7),
            ("BOTTOMPADDING",(0,0),(-1,-1),7),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEBELOW",(0,0),(-1,-1),0.4,LGRAY),
        ]))
        story.append(ta)
        story.append(Spacer(1, 8*mm))

    # ══════════════════════════════════════════════
    # SECTION 07 — POINTS DE VIGILANCE
    # ══════════════════════════════════════════════
    vigilance = recommendations.get("points_vigilance", [])
    if vigilance:
        story += _section_header("07", "POINTS DE VIGILANCE")
        for i, pt in enumerate(vigilance, 1):
            story.append(Table([[
                Paragraph(f"{i:02d}", S("vig_n", fontName="Helvetica-Bold",
                         fontSize=9, textColor=LGRAY, alignment=TA_CENTER)),
                Paragraph(pt, S("vig_t", fontName="Helvetica", fontSize=9,
                         leading=14, textColor=DGRAY)),
            ]], colWidths=[14*mm, 156*mm], style=TableStyle([
                ("LEFTPADDING",(0,0),(-1,-1),6),
                ("TOPPADDING",(0,0),(-1,-1),5),
                ("BOTTOMPADDING",(0,0),(-1,-1),5),
                ("LINEBELOW",(0,0),(-1,-1),0.4,LGRAY),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
            ])))
        story.append(Spacer(1, 8*mm))

    # ══════════════════════════════════════════════
    # SECTION 08 — PROCHAINE ÉTAPE
    # ══════════════════════════════════════════════
    next_step = recommendations.get("prochaine_etape", "")
    if next_step:
        story += _section_header("08", "ACTION PRIORITAIRE — 7 PROCHAINS JOURS")
        story.append(Table([[
            Paragraph("À FAIRE MAINTENANT", S("ns_lbl", fontName="Helvetica-Bold",
                     fontSize=8, textColor=WHITE, charSpace=1)),
            Paragraph(next_step, S("ns_val", fontName="Helvetica-Bold",
                     fontSize=11, leading=15, textColor=WHITE)),
        ]], colWidths=[38*mm, 132*mm], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLACK),
            ("LEFTPADDING",(0,0),(-1,-1),12),
            ("TOPPADDING",(0,0),(-1,-1),14),
            ("BOTTOMPADDING",(0,0),(-1,-1),14),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEAFTER",(0,0),(0,-1),1,MGRAY),
        ])))
        story.append(Spacer(1, 8*mm))

    # ══════════════════════════════════════════════
    # ANNEXE — ESPACE ANNOTATIONS EXPERT
    # ══════════════════════════════════════════════
    story.append(HLine(color=LGRAY, thickness=0.4))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "ANNEXE — NOTES ET COMMENTAIRES EXPERTS",
        S("annex_title", fontName="Helvetica-Bold", fontSize=8, textColor=GRAY, charSpace=1)))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "[ Espace réservé à l'équipe KORD pour ajouter ses observations, données complémentaires et recommandations avant restitution au client. ]",
        S("annex_note", fontName="Helvetica", fontSize=8.5, leading=12,
          textColor=colors.HexColor("#BBBBBB"), leftIndent=8)))
    story.append(Spacer(1, 30*mm))
    story.append(HLine(color=LGRAY, thickness=0.3))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f"KORD — Pré-rapport d'audit opérationnel — {trim_str} — Document de travail confidentiel — Score indicatif non communicable au client",
        S("footer", fontName="Helvetica", fontSize=7, textColor=GRAY)))

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buf.seek(0)
    return buf.read()


def _section_header(num, title):
    return [
        Paragraph(f"{num}  —  {title}", ParagraphStyle(
            "sec_lbl", fontName="Helvetica-Bold", fontSize=8,
            textColor=GRAY, leading=11, spaceBefore=4, spaceAfter=6, charSpace=1)),
        HLine(color=BLACK, thickness=1),
        Spacer(1, 4*mm),
    ]


def _footer(canvas, doc):
    canvas.saveState()
    W, H = A4
    canvas.setFillColor(GRAY)
    canvas.setFont("Helvetica", 7)
    canvas.drawString(18*mm, 8*mm, "KORD — Pré-rapport d'audit opérationnel — Document confidentiel")
    canvas.drawRightString(W - 18*mm, 8*mm, f"Page {doc.page}")
    canvas.restoreState()
