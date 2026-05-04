"""
KORD HTML — Rapport premium modifiable
Ouvrir dans Chrome → modifier → Imprimer → PDF
"""
from datetime import datetime
from typing import Dict, Any
import base64


def img_to_b64(png_bytes):
    if not png_bytes:
        return None
    return "data:image/png;base64," + base64.b64encode(png_bytes).decode()


def generate_prereport_html(
    consolidated, recommendations, all_results,
    client_name="Client", company_name="",
    trimestre="", gauge_png=None, bar_png=None,
    radar_png=None, evol_png=None,
) -> str:

    score    = consolidated.get("score_total", 0)
    now      = datetime.now()
    date_str = now.strftime("%d %B %Y").upper()
    trim_str = trimestre or f"T{(now.month-1)//3+1} {now.year}"
    nb       = len(all_results)
    reco     = recommendations

    # Images en base64
    gauge_b64 = img_to_b64(gauge_png)
    bar_b64   = img_to_b64(bar_png)
    radar_b64 = img_to_b64(radar_png)
    evol_b64  = img_to_b64(evol_png)

    # Résumé
    resume = reco.get("resume_executif", {})
    if isinstance(resume, dict):
        p1 = resume.get("paragraphe_1","")
        p2 = resume.get("paragraphe_2","")
        p3 = resume.get("paragraphe_3","")
    else:
        p1, p2, p3 = str(resume), "", ""

    msg_dir   = reco.get("message_dirigeant","")
    benchmark = reco.get("benchmark","")

    # Piliers
    pilier_order = ["stock_cash","transport_service","achats_fournisseurs","marges_retours","donnees_pilotage"]
    pilier_meta  = {
        "stock_cash":          ("Stock et Cash Immobilisé",      30),
        "transport_service":   ("Transport et Taux de Service",  20),
        "achats_fournisseurs": ("Achats et Fournisseurs",        20),
        "marges_retours":      ("Marges et Retours Clients",     15),
        "donnees_pilotage":    ("Données et Pilotage",           15),
    }
    analyse_piliers = reco.get("analyse_piliers", {})

    def pilier_html():
        out = ""
        for key in pilier_order:
            label, max_pts = pilier_meta[key]
            sc = consolidated.get("analyses",{}).get(key,{}).get("score",0)
            dp = analyse_piliers.get(key, {})
            if isinstance(dp, dict):
                sc      = dp.get("score", sc)
                titre   = dp.get("titre", label)
                niveau  = dp.get("niveau","moyen")
                analyse = dp.get("analyse","")
                chiffres= dp.get("chiffres_cles",[])
                risque  = dp.get("risque_principal","")
            else:
                titre, niveau, analyse, chiffres, risque = label, "moyen", "", [], ""
            pct = round(sc/max_pts*100) if max_pts > 0 else 0
            niveau_cls = "niveau-bon" if pct >= 70 else "niveau-moyen" if pct >= 45 else "niveau-critique"

            chiffres_html = "".join(f'<div class="chiffre-cle">▸ {c}</div>' for c in chiffres if c and c != "...")
            risque_html = f'<div class="risque-box"><span class="risque-label">RISQUE PRINCIPAL</span><span class="risque-val">{risque}</span></div>' if risque and risque != "..." else ""

            out += f'''
            <div class="pilier-card">
              <div class="pilier-header {niveau_cls}">
                <span class="pilier-titre">{titre.upper()}</span>
                <span class="pilier-score">{sc}/{max_pts}</span>
              </div>
              <div class="pilier-progress"><div class="pilier-bar" style="width:{pct}%"></div></div>
              <div class="pilier-body">
                <p class="pilier-analyse">{analyse}</p>
                {chiffres_html}
                {risque_html}
              </div>
            </div>'''
        return out

    # Anomalies
    anomalies = reco.get("anomalies", consolidated.get("alertes",[]))
    u_order = {"CRITIQUE":0,"MOYEN":1,"FAIBLE":2}
    if anomalies and isinstance(anomalies[0], dict) and "urgence" in anomalies[0]:
        anomalies = sorted(anomalies, key=lambda x: u_order.get(x.get("urgence","FAIBLE"),2))

    def anomalies_html():
        if not anomalies: return "<p class='empty'>Aucune anomalie détectée.</p>"
        rows = ""
        for i, al in enumerate(anomalies[:12]):
            if not isinstance(al, dict): continue
            urg = al.get("urgence","MOYEN")
            urg_cls = f"urg-{urg.lower()}"
            titre   = al.get("titre", al.get("message",""))
            detect  = al.get("detection", al.get("message",""))
            impact  = al.get("impact_business","")
            imp_fin = al.get("impact_financier","")
            bg = "#fff" if i%2==0 else "#fafafa"
            rows += f'''
            <tr style="background:{bg}">
              <td><span class="urg-badge {urg_cls}">{urg}</span></td>
              <td><strong>{titre}</strong><br><span class="small">{detect}</span></td>
              <td>{impact}</td>
              <td class="impact-fin">{imp_fin}</td>
            </tr>'''
        return f'<table class="data-table"><thead><tr><th>URGENCE</th><th>ANOMALIE DÉTECTÉE</th><th>IMPACT BUSINESS</th><th>IMPACT €</th></tr></thead><tbody>{rows}</tbody></table>'

    # Priorités
    priorites = reco.get("priorites",[])
    def priorites_html():
        out = ""
        for p in priorites[:5]:
            if not isinstance(p, dict): continue
            rang   = p.get("rang","")
            titre  = p.get("titre","")
            prob   = p.get("probleme","")
            action = p.get("action","")
            impact = p.get("impact_attendu","")
            gain   = p.get("gain_potentiel","")
            delai  = p.get("delai","")
            compl  = p.get("complexite","")
            qw     = p.get("quick_win", False)
            qw_badge = '<span class="qw-badge">⚡ QUICK WIN</span>' if qw else ""
            out += f'''
            <div class="priorite-card">
              <div class="prio-header">
                <span class="prio-num">0{rang}</span>
                <span class="prio-titre">{titre.upper()}</span>
                {qw_badge}
                <span class="prio-delai">{delai}</span>
              </div>
              <div class="prio-body">
                <div class="prio-col">
                  <div class="prio-field"><span class="prio-lbl">Problème détecté</span><p>{prob}</p></div>
                  <div class="prio-field"><span class="prio-lbl">Action recommandée</span><p>{action}</p></div>
                </div>
                <div class="prio-col prio-right">
                  <div class="prio-impact-box">
                    <span class="prio-lbl">IMPACT ATTENDU</span>
                    <p class="impact-big">{impact}</p>
                  </div>
                  <div class="gain-box">
                    <span class="prio-lbl">GAIN ESTIMÉ</span>
                    <p class="gain-val">{gain}</p>
                    <span class="compl-badge">{compl}</span>
                  </div>
                </div>
              </div>
            </div>'''
        return out

    # Croisements
    croisements = reco.get("croisements_cles",[])
    def croisements_html():
        if not croisements: return ""
        rows = ""
        for cr in croisements[:4]:
            if not isinstance(cr, dict): continue
            f = " × ".join(cr.get("fichiers",[]))
            obs = cr.get("observation","")
            imp = cr.get("impact","")
            if obs and obs != "...":
                rows += f'<tr><td class="cr-fichiers">{f}</td><td>{obs}</td><td class="cr-impact">{imp}</td></tr>'
        if not rows: return ""
        return f'''
        <section class="report-section">
          <div class="section-header"><span class="section-num">04</span><span class="section-title">CROISEMENTS CLÉS</span></div>
          <p class="section-sub">Ce que révèle l'analyse croisée de l'ensemble de vos données</p>
          <table class="data-table"><thead><tr><th>FICHIERS CROISÉS</th><th>OBSERVATION</th><th>IMPACT</th></tr></thead><tbody>{rows}</tbody></table>
        </section>'''

    # Opportunités
    opps = reco.get("opportunites_cachees",[])
    def opps_html():
        if not opps: return ""
        items = ""
        for opp in opps:
            if not isinstance(opp, dict): continue
            t = opp.get("titre","")
            d = opp.get("description","")
            g = opp.get("gain_estime","")
            if t and t != "...":
                items += f'<div class="opp-card"><div class="opp-head"><span>{t}</span><span class="opp-gain">{g}</span></div><p>{d}</p></div>'
        if not items: return ""
        return f'''
        <section class="report-section">
          <div class="section-header"><span class="section-num">07</span><span class="section-title">OPPORTUNITÉS CACHÉES</span></div>
          {items}
        </section>'''

    # Vigilance + Questions
    vigilance = reco.get("points_vigilance",[])
    questions = reco.get("questions_restitution",[])
    next_step = reco.get("prochaine_etape","")

    def vigilance_html():
        items = "".join(f'<div class="vig-item"><span class="vig-num">{i+1:02d}</span><p>{pt}</p></div>'
                       for i, pt in enumerate(vigilance) if pt and pt != "...")
        return items

    def questions_html():
        items = "".join(f'<div class="q-item"><span class="q-num">Q{i+1}</span><p>{q}</p></div>'
                       for i, q in enumerate(questions) if q and q != "...")
        return items

    gauge_img_tag  = f'<img src="{gauge_b64}" class="gauge-img" alt="Score">' if gauge_b64 else ""
    bar_img_tag    = f'<img src="{bar_b64}" class="chart-img" alt="Barres">' if bar_b64 else ""
    radar_img_tag  = f'<img src="{radar_b64}" class="chart-img radar-img" alt="Radar">' if radar_b64 else ""
    evol_img_tag   = f'<img src="{evol_b64}" class="chart-full" alt="Evolution">' if evol_b64 else ""

    score_color = "#000" if score < 45 else "#1a1a1a" if score < 70 else "#333"

    html = f'''<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>KORD — Pré-rapport {company_name or client_name} — {trim_str}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=Inter:wght@300;400;500&display=swap');

  :root {{
    --black: #000000;
    --white: #ffffff;
    --off: #F4F3F0;
    --gray: #6B6B6B;
    --lgray: #DEDEDE;
    --dgray: #1A1A1A;
  }}

  * {{ margin:0; padding:0; box-sizing:border-box; }}

  body {{
    font-family: 'Inter', sans-serif;
    font-weight: 300;
    background: #fff;
    color: var(--dgray);
    font-size: 10pt;
    line-height: 1.6;
  }}

  /* ── COVER ── */
  .cover {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    min-height: 100vh;
    page-break-after: always;
  }}
  .cover-left {{
    background: var(--black);
    padding: 60px 52px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }}
  .cover-kord {{
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 80pt;
    color: #fff;
    line-height: 0.9;
    letter-spacing: -0.02em;
  }}
  .cover-subtitle {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 9pt;
    color: #666;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    margin-top: 8px;
  }}
  .cover-divider {{
    border: none;
    border-top: 0.5px solid #333;
    margin: 28px 0;
  }}
  .cover-company {{
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 22pt;
    color: #fff;
    text-transform: uppercase;
    letter-spacing: -0.01em;
  }}
  .cover-client {{
    font-size: 12pt;
    color: #aaa;
    margin-top: 4px;
  }}
  .cover-meta {{
    font-size: 8pt;
    color: #555;
    margin-top: 12px;
  }}
  .cover-warning {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    color: #444;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    margin-top: 28px;
    padding-top: 16px;
    border-top: 0.5px solid #333;
  }}
  .cover-right {{
    background: var(--off);
    padding: 60px 48px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
  }}
  .score-label {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 8pt;
    font-weight: 600;
    color: var(--gray);
    letter-spacing: 0.2em;
    text-transform: uppercase;
    margin-bottom: 8px;
  }}
  .score-big {{
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 88pt;
    line-height: 1;
    color: var(--black);
    letter-spacing: -0.02em;
  }}
  .score-sub {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 8pt;
    color: var(--gray);
    letter-spacing: 0.15em;
    margin-bottom: 20px;
  }}
  .gauge-img {{ width: 100%; max-width: 240px; margin: 12px 0; }}

  /* Piliers mini sur cover */
  .cover-piliers {{ width: 100%; margin-top: 16px; }}
  .cover-pilier-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 5px 0;
    border-bottom: 0.5px solid var(--lgray);
    font-size: 8.5pt;
  }}
  .cover-pilier-name {{ color: var(--dgray); }}
  .cover-pilier-score {{ font-weight: 600; color: var(--black); }}
  .cover-pilier-bar-wrap {{ width: 60px; height: 3px; background: var(--lgray); margin: 0 8px; }}
  .cover-pilier-bar {{ height: 3px; background: var(--black); }}

  /* ── PAGE LAYOUT ── */
  .page {{ padding: 48px 52px; page-break-after: always; }}
  .page:last-child {{ page-break-after: auto; }}

  /* ── SECTION HEADER ── */
  .section-header {{
    display: flex;
    align-items: center;
    gap: 12px;
    background: var(--black);
    padding: 10px 16px;
    margin-bottom: 4px;
    margin-top: 28px;
  }}
  .section-header:first-child {{ margin-top: 0; }}
  .section-num {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 8pt;
    font-weight: 700;
    color: #666;
    letter-spacing: 0.1em;
    min-width: 24px;
  }}
  .section-title {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 10pt;
    font-weight: 600;
    color: #fff;
    letter-spacing: 0.1em;
    text-transform: uppercase;
  }}
  .section-sub {{
    font-size: 8.5pt;
    color: var(--gray);
    margin-bottom: 12px;
    font-style: italic;
  }}

  /* ── STAT BOXES ── */
  .stat-row {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
    margin: 12px 0;
  }}
  .stat-box {{
    background: var(--black);
    padding: 16px 12px;
    text-align: center;
  }}
  .stat-num {{
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 24pt;
    color: #fff;
    line-height: 1;
  }}
  .stat-label {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-top: 4px;
  }}

  /* ── PULL QUOTE ── */
  .pull-quote {{
    border-left: 4px solid var(--black);
    background: var(--off);
    padding: 14px 18px;
    margin: 12px 0;
  }}
  .pull-quote p {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 10.5pt;
    font-weight: 600;
    color: var(--black);
    line-height: 1.5;
  }}

  /* ── BENCHMARK ── */
  .benchmark-box {{
    display: flex;
    background: var(--off);
    padding: 10px 14px;
    gap: 16px;
    margin: 8px 0;
  }}
  .bench-label {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    font-weight: 700;
    color: var(--gray);
    letter-spacing: 0.1em;
    text-transform: uppercase;
    min-width: 70px;
    padding-top: 1px;
  }}
  .bench-val {{ font-size: 9pt; color: var(--dgray); line-height: 1.5; }}

  /* ── BODY TEXT ── */
  .body-text {{ font-size: 9.5pt; line-height: 1.7; margin: 6px 0; color: var(--dgray); text-align: justify; }}

  /* ── CHARTS ── */
  .charts-row {{ display: flex; gap: 12px; align-items: center; margin: 8px 0; background: var(--off); padding: 10px; }}
  .chart-img {{ max-height: 200px; flex: 2; object-fit: contain; }}
  .radar-img {{ flex: 1; max-height: 180px; }}
  .chart-full {{ width: 100%; margin: 8px 0; }}
  .chart-source {{ font-size: 7pt; color: var(--gray); font-style: italic; margin: 2px 0 10px; }}

  /* ── PILIERS GRID ── */
  .piliers-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin: 8px 0; }}
  .pilier-card {{ border: 0.5px solid var(--lgray); overflow: hidden; }}
  .pilier-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 12px;
    color: #fff;
  }}
  .niveau-critique {{ background: #000; }}
  .niveau-moyen {{ background: #1A1A1A; }}
  .niveau-bon {{ background: #444; }}
  .pilier-titre {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 8pt;
    font-weight: 600;
    letter-spacing: 0.05em;
  }}
  .pilier-score {{
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 13pt;
  }}
  .pilier-progress {{ height: 3px; background: var(--lgray); }}
  .pilier-bar {{ height: 3px; background: var(--black); }}
  .pilier-body {{ padding: 10px 12px; background: #fff; }}
  .pilier-analyse {{ font-size: 9pt; line-height: 1.6; color: var(--dgray); text-align: justify; }}
  .chiffre-cle {{
    font-size: 8.5pt;
    color: var(--gray);
    padding: 2px 0;
    padding-left: 8px;
  }}
  .risque-box {{
    display: flex;
    gap: 10px;
    background: var(--off);
    padding: 5px 8px;
    margin-top: 6px;
    align-items: baseline;
  }}
  .risque-label {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 6.5pt;
    font-weight: 700;
    color: var(--gray);
    letter-spacing: 0.1em;
    text-transform: uppercase;
    min-width: 55px;
  }}
  .risque-val {{ font-size: 9pt; font-weight: 600; color: var(--black); }}

  /* ── DATA TABLE ── */
  .data-table {{ width: 100%; border-collapse: collapse; margin: 6px 0; font-size: 9pt; }}
  .data-table thead tr {{ background: var(--black); }}
  .data-table th {{
    padding: 7px 10px;
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    font-weight: 600;
    color: #fff;
    text-align: left;
    letter-spacing: 0.1em;
    text-transform: uppercase;
  }}
  .data-table td {{ padding: 7px 10px; border-bottom: 0.5px solid var(--lgray); vertical-align: middle; }}
  .data-table .small {{ font-size: 8pt; color: var(--gray); }}
  .impact-fin {{ font-weight: 600; font-family: 'Space Grotesk', sans-serif; text-align: right; }}

  /* Urgence badges */
  .urg-badge {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 6.5pt;
    font-weight: 700;
    padding: 2px 6px;
    color: #fff;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    display: inline-block;
  }}
  .urg-critique {{ background: #000; }}
  .urg-moyen    {{ background: #444; }}
  .urg-faible   {{ background: #888; }}

  /* ── PRIORITÉS ── */
  .priorite-card {{ border: 0.5px solid var(--lgray); margin: 6px 0; overflow: hidden; }}
  .prio-header {{
    display: flex;
    align-items: center;
    gap: 10px;
    background: var(--black);
    padding: 9px 14px;
  }}
  .prio-num {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 18pt;
    font-weight: 700;
    color: #444;
    line-height: 1;
    min-width: 36px;
  }}
  .prio-titre {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 9.5pt;
    font-weight: 600;
    color: #fff;
    flex: 1;
    letter-spacing: 0.05em;
  }}
  .qw-badge {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    font-weight: 700;
    color: #fff;
    background: #333;
    padding: 2px 7px;
  }}
  .prio-delai {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7.5pt;
    color: #888;
  }}
  .prio-body {{ display: flex; }}
  .prio-col {{ flex: 1; padding: 12px 14px; background: #fafafa; }}
  .prio-right {{ background: var(--off); border-left: 0.5px solid var(--lgray); flex: 0 0 220px; }}
  .prio-field {{ margin-bottom: 8px; }}
  .prio-lbl {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    font-weight: 700;
    color: var(--gray);
    text-transform: uppercase;
    letter-spacing: 0.1em;
    display: block;
    margin-bottom: 3px;
  }}
  .prio-field p {{ font-size: 9pt; color: var(--dgray); line-height: 1.5; }}
  .prio-impact-box {{ padding-bottom: 10px; margin-bottom: 10px; border-bottom: 0.5px solid var(--lgray); }}
  .impact-big {{ font-size: 10pt; font-weight: 600; color: var(--black); line-height: 1.4; margin-top: 3px; }}
  .gain-box {{ }}
  .gain-val {{ font-family: 'Space Grotesk', sans-serif; font-size: 12pt; font-weight: 700; color: var(--black); margin: 3px 0; }}
  .compl-badge {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    color: var(--gray);
    border: 0.5px solid var(--lgray);
    padding: 2px 6px;
    display: inline-block;
    margin-top: 4px;
  }}

  /* ── CROISEMENTS ── */
  .cr-fichiers {{ font-weight: 600; font-size: 8pt; color: var(--gray); }}
  .cr-impact {{ font-weight: 600; text-align: right; }}

  /* ── OPPORTUNITÉS ── */
  .opp-card {{ margin: 6px 0; border: 0.5px solid var(--lgray); }}
  .opp-head {{
    display: flex;
    justify-content: space-between;
    background: var(--dgray);
    padding: 9px 14px;
    font-family: 'Space Grotesk', sans-serif;
    font-size: 10pt;
    font-weight: 600;
    color: #fff;
  }}
  .opp-gain {{ font-size: 11pt; }}
  .opp-card p {{ padding: 8px 14px; font-size: 9pt; color: var(--dgray); background: var(--off); }}

  /* ── VIGILANCE ── */
  .vig-item {{ display: flex; gap: 10px; padding: 7px 0; border-bottom: 0.5px solid var(--lgray); }}
  .vig-num {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 9pt;
    font-weight: 700;
    color: var(--lgray);
    min-width: 26px;
  }}
  .vig-item p {{ font-size: 9pt; color: var(--dgray); line-height: 1.5; }}

  /* ── QUESTIONS ── */
  .questions-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }}
  .q-item {{ display: flex; gap: 8px; background: var(--off); padding: 8px 10px; }}
  .q-num {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 9pt;
    font-weight: 700;
    color: #fff;
    background: var(--black);
    min-width: 26px;
    display: flex;
    align-items: center;
    justify-content: center;
  }}
  .q-item p {{ font-size: 9pt; color: var(--dgray); line-height: 1.5; }}

  /* ── NEXT STEP ── */
  .next-step-box {{
    background: var(--black);
    padding: 16px 20px;
    border-left: 4px solid #555;
    margin: 8px 0;
  }}
  .next-step-box p {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 10.5pt;
    font-weight: 500;
    color: #fff;
    line-height: 1.6;
  }}

  /* ── ANNEXE ── */
  .annexe-box {{
    border: 1px dashed var(--lgray);
    padding: 20px;
    min-height: 80px;
    margin: 8px 0;
    color: #ccc;
    font-style: italic;
    font-size: 9pt;
  }}

  /* ── BREADCRUMB ── */
  .breadcrumb {{
    position: fixed;
    bottom: 0; left: 0; right: 0;
    background: #fff;
    border-top: 0.5px solid var(--lgray);
    padding: 4px 52px;
    display: flex;
    gap: 28px;
    font-family: 'Space Grotesk', sans-serif;
    font-size: 7pt;
    color: var(--gray);
    z-index: 100;
  }}
  .bc-item {{ cursor: pointer; }}
  .bc-item:hover {{ color: var(--black); }}

  /* ── FOOTER ── */
  .report-footer {{
    border-top: 0.5px solid var(--lgray);
    padding: 10px 0 0;
    margin-top: 20px;
    font-size: 7pt;
    color: var(--gray);
  }}

  /* ── PRINT ── */
  @media print {{
    body {{ font-size: 9pt; }}
    .cover {{ min-height: auto; height: 100vh; }}
    .page {{ padding: 36px 44px; }}
    .breadcrumb {{ display: none; }}
    .no-print {{ display: none; }}
    @page {{ size: A4 landscape; margin: 0; }}
    h1, h2 {{ page-break-after: avoid; }}
    .priorite-card, .pilier-card {{ page-break-inside: avoid; }}
  }}
</style>
</head>
<body>

<!-- ══ COVER ══ -->
<div class="cover">
  <div class="cover-left">
    <div>
      <div class="cover-kord">KORD</div>
      <div class="cover-subtitle">Audit de Performance Opérationnelle</div>
    </div>
    <div>
      <hr class="cover-divider">
      <div class="cover-company">{company_name or client_name}</div>
      {'<div class="cover-client">' + client_name + '</div>' if company_name else ''}
      <div class="cover-meta">{trim_str}  ·  {date_str}  ·  {nb} fichier{"s" if nb>1 else ""} analysé{"s" if nb>1 else ""}</div>
      <div class="cover-warning">⚠  Document de travail — Usage interne KORD — Ne pas transmettre sans validation</div>
    </div>
  </div>
  <div class="cover-right">
    <div class="score-label">Score KORD</div>
    <div class="score-big">{score}</div>
    <div class="score-sub">/100 — INDICATIF</div>
    {gauge_img_tag}
    <div class="cover-piliers">
      {cover_pilier_rows}
    </div>
  </div>
</div>

<!-- ══ PAGE 2 — RÉSUMÉ ══ -->
<div class="page">

  <div class="section-header"><span class="section-num">01</span><span class="section-title">Résumé Exécutif</span></div>
  <p class="section-sub">Situation globale et principaux constats issus de l'analyse de vos données</p>

  <div class="stat-row">
    <div class="stat-box"><div class="stat-num">{score}/100</div><div class="stat-label">Score KORD indicatif</div></div>
    <div class="stat-box"><div class="stat-num">{len(consolidated.get("alertes",[]))}</div><div class="stat-label">Anomalie{"s" if len(consolidated.get("alertes",[]))>1 else ""} détectée{"s" if len(consolidated.get("alertes",[]))>1 else ""}</div></div>
    <div class="stat-box"><div class="stat-num">{len(reco.get("priorites",[]))}</div><div class="stat-label">Priorité{"s" if len(reco.get("priorites",[]))>1 else ""} d'action</div></div>
    <div class="stat-box"><div class="stat-num">{nb}</div><div class="stat-label">Fichier{"s" if nb>1 else ""} analysé{"s" if nb>1 else ""}</div></div>
  </div>

  {'<p class="body-text">' + p1 + '</p>' if p1 and p1 not in ["...",""] else ""}
  {'<p class="body-text">' + p2 + '</p>' if p2 and p2 not in ["...",""] else ""}
  {'<p class="body-text">' + p3 + '</p>' if p3 and p3 not in ["...",""] else ""}

  {'<div class="pull-quote"><p>' + msg_dir + '</p></div>' if msg_dir and msg_dir not in ["...",""] else ""}

  {'<div class="benchmark-box"><span class="bench-label">Benchmark sectoriel</span><span class="bench-val">' + benchmark + '</span></div>' if benchmark and benchmark not in ["...",""] else ""}

  <div class="section-header"><span class="section-num">02</span><span class="section-title">Graphiques d'Analyse</span></div>
  <p class="section-sub">Score KORD vs Benchmark marché PME</p>

  <div class="charts-row">
    {bar_img_tag}
    {radar_img_tag}
  </div>
  <p class="chart-source">Source : Moteur d'analyse KORD — Données client — Benchmark PME secteur distribution/logistique</p>

  {('<div>' + evol_img_tag + '</div><p class="chart-source">Source : Trajectoire indicative — enrichie à chaque trimestre</p>') if evol_img_tag else ""}
</div>

<!-- ══ PAGE 3 — PILIERS ══ -->
<div class="page">
  <div class="section-header"><span class="section-num">03</span><span class="section-title">Performance par Pilier</span></div>
  <p class="section-sub">Analyse détaillée — Chiffres issus de vos données — Risques identifiés</p>

  <div class="piliers-grid">
    {pilier_html()}
  </div>
</div>

<!-- ══ PAGE 4 — ANOMALIES ══ -->
<div class="page">
  {croisements_html()}

  <div class="section-header"><span class="section-num">05</span><span class="section-title">Anomalies Détectées</span></div>
  <p class="section-sub">Classées par urgence — Impact business estimé sur votre activité</p>
  {anomalies_html()}
</div>

<!-- ══ PAGE 5 — PRIORITÉS ══ -->
<div class="page">
  <div class="section-header"><span class="section-num">06</span><span class="section-title">Priorités d'Action</span></div>
  <p class="section-sub">Top actions classées par impact financier — Ce qui rapporte le plus vite</p>
  {priorites_html()}
</div>

<!-- ══ PAGE 6 — SUITE ══ -->
<div class="page">
  {opps_html()}

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
    <div>
      <div class="section-header"><span class="section-num">08</span><span class="section-title">Points de Vigilance</span></div>
      {vigilance_html()}
    </div>
    <div>
      <div class="section-header"><span class="section-num">09</span><span class="section-title">Questions Restitution</span></div>
      <div class="questions-grid">
        {questions_html()}
      </div>
    </div>
  </div>

  {'<div class="section-header" style="margin-top:20px"><span class="section-num">10</span><span class="section-title">Prochaine Étape</span></div><div class="next-step-box"><p>' + next_step + '</p></div>' if next_step and next_step not in ["...",""] else ""}

  <div class="section-header" style="margin-top:20px"><span class="section-num">—</span><span class="section-title">Annexe — Notes Experts KORD</span></div>
  <div class="annexe-box" contenteditable="true">[ Cliquez ici pour ajouter vos observations, données complémentaires et recommandations avant envoi au client. ]</div>

  <div class="report-footer">
    KORD — Audit de performance opérationnelle — {trim_str} — Document de travail confidentiel — Score indicatif avant validation experts KORD
  </div>
</div>

<!-- Breadcrumb navigation -->
<div class="breadcrumb no-print">
  <span class="bc-item" onclick="document.querySelector('.cover').scrollIntoView()">Couverture</span>
  <span class="bc-item">Résumé Exécutif</span>
  <span class="bc-item">Piliers</span>
  <span class="bc-item">Anomalies</span>
  <span class="bc-item">Priorités</span>
  <span class="bc-item">Suite</span>
</div>

</body>
</html>'''

    return html
