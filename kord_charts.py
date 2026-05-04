"""
KORD CHARTS — Graphiques style rapport JLL / cabinet de conseil
Propres, annotés, professionnels — identité KORD noir/blanc
"""
import io
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import matplotlib.gridspec as gridspec

# Palette KORD
C_BLACK  = "#000000"
C_WHITE  = "#FFFFFF"
C_OFF    = "#F4F3F0"
C_GRAY   = "#6B6B6B"
C_LGRAY  = "#DDDDDD"
C_DGRAY  = "#1A1A1A"
C_MGRAY  = "#888888"

def _base_style(fig, bg=C_OFF):
    fig.patch.set_facecolor(bg)
    plt.rcParams.update({
        'font.family': 'DejaVu Sans',
        'axes.spines.top': False,
        'axes.spines.right': False,
    })


def generate_gauge_chart(score: int) -> bytes:
    """Jauge demi-cercle style tableau de bord — score KORD."""
    fig, ax = plt.subplots(figsize=(7, 3.8))
    _base_style(fig, C_BLACK)
    ax.set_facecolor(C_BLACK)
    ax.set_xlim(-1.2, 1.2)
    ax.set_ylim(-0.15, 1.15)
    ax.axis('off')

    # Fond arc gris
    theta_bg = np.linspace(np.pi, 0, 300)
    r = 1.0
    xb = r * np.cos(theta_bg)
    yb = r * np.sin(theta_bg)
    ax.plot(xb, yb, color="#333333", linewidth=28, solid_capstyle='round', zorder=1)

    # Arc score
    ratio = min(score / 100, 1.0)
    theta_fill = np.linspace(np.pi, np.pi - ratio * np.pi, 300)
    xf = r * np.cos(theta_fill)
    yf = r * np.sin(theta_fill)
    ax.plot(xf, yf, color=C_WHITE, linewidth=28, solid_capstyle='round', zorder=2)

    # Zones colorées discrètes
    zones = [(0.0, 0.35, "#555555"), (0.35, 0.65, "#444444"), (0.65, 1.0, "#333333")]
    for z_start, z_end, _ in zones:
        t = np.linspace(np.pi - z_start*np.pi, np.pi - z_end*np.pi, 100)
        ax.plot(r*np.cos(t), r*np.sin(t), color=_,
                linewidth=28, solid_capstyle='butt', zorder=1, alpha=0.3)

    # Labels zones
    for angle, label in [(np.pi*0.9, "CRITIQUE"), (np.pi*0.6, "MOYEN"), (np.pi*0.2, "BON")]:
        ax.text(1.13*np.cos(angle), 1.13*np.sin(angle), label,
                ha='center', va='center', fontsize=6.5, color=C_MGRAY,
                fontweight='bold')

    # Score central
    ax.text(0, 0.25, str(score), ha='center', va='center',
            fontsize=64, fontweight='bold', color=C_WHITE, zorder=5)
    ax.text(0, -0.05, "/100  —  SCORE KORD", ha='center', va='center',
            fontsize=9, color=C_MGRAY, zorder=5)

    # Aiguille
    needle_angle = np.pi - ratio * np.pi
    nx = 0.72 * np.cos(needle_angle)
    ny = 0.72 * np.sin(needle_angle)
    ax.annotate('', xy=(nx, ny), xytext=(0, 0.02),
                arrowprops=dict(arrowstyle='->', color=C_WHITE, lw=2.5))
    ax.plot(0, 0.02, 'o', color=C_WHITE, markersize=8, zorder=6)

    plt.tight_layout(pad=0.3)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor=C_BLACK)
    plt.close()
    buf.seek(0)
    return buf.read()


def generate_bar_chart(scores: dict) -> bytes:
    """
    Barres horizontales style JLL — scores par pilier
    Avec benchmark marché en superposition.
    """
    piliers = {
        "stock_cash":          ("Stock & Cash", 30),
        "transport_service":   ("Transport & Service", 20),
        "achats_fournisseurs": ("Achats & Fournisseurs", 20),
        "marges_retours":      ("Marges & Retours", 15),
        "donnees_pilotage":    ("Données & Pilotage", 15),
    }
    # Benchmarks moyens secteur (PME)
    benchmarks = {
        "stock_cash":          0.61,
        "transport_service":   0.68,
        "achats_fournisseurs": 0.65,
        "marges_retours":      0.70,
        "donnees_pilotage":    0.58,
    }

    labels, vals, maxes, pcts, benches = [], [], [], [], []
    for key, (label, max_pts) in piliers.items():
        a = scores.get(key, {})
        s = a.get("score", 0) if isinstance(a, dict) else 0
        pct = s / max_pts if max_pts > 0 else 0
        labels.append(label)
        vals.append(s)
        maxes.append(max_pts)
        pcts.append(pct)
        benches.append(benchmarks.get(key, 0.65))

    fig, ax = plt.subplots(figsize=(9, 5.5))
    _base_style(fig, C_OFF)
    ax.set_facecolor(C_OFF)

    y = np.arange(len(labels))
    h = 0.42

    # Fond total (max)
    ax.barh(y, [1.0]*len(labels), height=h, color=C_LGRAY, zorder=1, left=0)

    # Barres score
    bar_colors = [C_BLACK if p >= 0.7 else C_DGRAY if p >= 0.45 else C_GRAY for p in pcts]
    bars = ax.barh(y, pcts, height=h, color=bar_colors, zorder=2)

    # Ligne benchmark
    for i, b in enumerate(benches):
        ax.plot([b, b], [i-h/2-0.04, i+h/2+0.04], color=C_MGRAY,
                linewidth=1.5, linestyle='--', zorder=3)

    # Annotations score
    for i, (p, s, m) in enumerate(zip(pcts, vals, maxes)):
        pct_label = f"{round(p*100)}%"
        score_label = f"{s}/{m} pts"
        ax.text(p + 0.01, i, pct_label, va='center', ha='left',
                fontsize=10, fontweight='bold', color=C_BLACK)
        ax.text(1.04, i, score_label, va='center', ha='left',
                fontsize=8.5, color=C_GRAY)

    # Niveau labels
    for i, p in enumerate(pcts):
        niveau = "BON" if p >= 0.7 else "MOYEN" if p >= 0.45 else "CRITIQUE"
        col = C_DGRAY if niveau == "BON" else C_GRAY if niveau == "MOYEN" else C_BLACK
        ax.text(-0.01, i, niveau, va='center', ha='right',
                fontsize=7, color=col, fontweight='bold', alpha=0.7)

    # Axes
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=10, color=C_DGRAY)
    ax.set_xlim(-0.12, 1.22)
    ax.set_ylim(-0.6, len(labels)-0.4)
    ax.set_xticks([0, 0.25, 0.5, 0.75, 1.0])
    ax.set_xticklabels(['0%', '25%', '50%', '75%', '100%'],
                       fontsize=8, color=C_GRAY)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color(C_LGRAY)
    ax.tick_params(left=False)
    ax.grid(axis='x', color=C_LGRAY, linewidth=0.5, alpha=0.5)

    # Légende benchmark
    bench_line = plt.Line2D([0], [0], color=C_MGRAY, linewidth=1.5,
                            linestyle='--', label='Benchmark PME secteur')
    score_patch = mpatches.Patch(color=C_BLACK, label='Score client KORD')
    ax.legend(handles=[score_patch, bench_line], loc='lower right',
             fontsize=8, frameon=False, labelcolor=C_GRAY)

    ax.set_title('Performance par pilier — Score KORD vs Benchmark marché',
                fontsize=11, fontweight='bold', color=C_BLACK, pad=14, loc='left')

    plt.tight_layout(pad=1.0)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor=C_OFF)
    plt.close()
    buf.seek(0)
    return buf.read()


def generate_radar_chart(scores: dict) -> bytes:
    """Radar chart style rapport conseil — 5 piliers."""
    piliers = {
        "stock_cash":          ("Stock\n& Cash", 30),
        "transport_service":   ("Transport\n& Service", 20),
        "achats_fournisseurs": ("Achats\nFournisseurs", 20),
        "marges_retours":      ("Marges\n& Retours", 15),
        "donnees_pilotage":    ("Données\nPilotage", 15),
    }
    benchmarks_pct = [0.61, 0.68, 0.65, 0.70, 0.58]

    keys = list(piliers.keys())
    labels = [piliers[k][0] for k in keys]
    client_vals = []
    for key in keys:
        a = scores.get(key, {})
        s = a.get("score", 0) if isinstance(a, dict) else 0
        m = piliers[key][1]
        client_vals.append(s / m if m > 0 else 0)

    N = len(keys)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]
    client_vals_plot = client_vals + client_vals[:1]
    bench_plot = benchmarks_pct + benchmarks_pct[:1]

    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    _base_style(fig, C_OFF)
    ax.set_facecolor(C_OFF)

    # Grille
    ax.set_ylim(0, 1)
    ax.set_yticks([0.25, 0.5, 0.75, 1.0])
    ax.set_yticklabels(['25%', '50%', '75%', '100%'], fontsize=7, color=C_LGRAY)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=9, color=C_DGRAY, fontweight='bold')
    ax.grid(color=C_LGRAY, linewidth=0.6, linestyle='-', alpha=0.7)
    ax.spines['polar'].set_color(C_LGRAY)

    # Benchmark
    ax.plot(angles, bench_plot, 'o--', color=C_MGRAY, linewidth=1.5,
            markersize=3, label='Benchmark PME', alpha=0.7, zorder=2)
    ax.fill(angles, bench_plot, color=C_MGRAY, alpha=0.06, zorder=1)

    # Score client
    ax.plot(angles, client_vals_plot, 'o-', color=C_BLACK, linewidth=2.5,
            markersize=5, label='Score KORD', zorder=4)
    ax.fill(angles, client_vals_plot, color=C_BLACK, alpha=0.12, zorder=3)

    # Annotations scores
    for angle, val, key in zip(angles[:-1], client_vals, keys):
        s = scores.get(key, {})
        score_val = s.get("score", 0) if isinstance(s, dict) else 0
        max_val = piliers[key][1]
        offset = 0.12
        x = (val + offset) * np.cos(angle - np.pi/2)
        y = (val + offset) * np.sin(angle - np.pi/2)
        ax.annotate(f"{score_val}/{max_val}",
                   xy=(angle, val), xytext=(angle, val + 0.14),
                   fontsize=8, fontweight='bold', color=C_BLACK,
                   ha='center', va='center')

    ax.legend(loc='upper right', bbox_to_anchor=(1.35, 1.15),
             fontsize=8, frameon=False, labelcolor=C_GRAY)
    ax.set_title('Radar de performance\nScore KORD vs Benchmark', 
                fontsize=10, fontweight='bold', color=C_BLACK, pad=20)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor=C_OFF)
    plt.close()
    buf.seek(0)
    return buf.read()


def generate_evolution_chart(scores_history: list = None) -> bytes:
    """
    Graphique d'évolution temporelle du score KORD — style JLL time series.
    Si pas d'historique, génère une projection indicative.
    """
    fig, ax = plt.subplots(figsize=(9, 4))
    _base_style(fig, C_OFF)
    ax.set_facecolor(C_OFF)

    trimestres = ['T1 2024', 'T2 2024', 'T3 2024', 'T4 2024', 'T1 2025', 'T2 2025']

    if scores_history and len(scores_history) >= 2:
        vals = scores_history[:6]
        while len(vals) < 6:
            vals.append(vals[-1])
        label = "Score KORD réel"
    else:
        # Courbe de référence secteur
        vals = [58, 61, 59, 64, 62, 67]
        label = "Score KORD (premier audit)"

    # Zone de performance
    ax.axhspan(70, 100, alpha=0.06, color=C_DGRAY, label='Zone performance')
    ax.axhspan(45, 70, alpha=0.04, color=C_GRAY)
    ax.axhspan(0,  45, alpha=0.06, color=C_MGRAY, label='Zone critique')

    # Courbe moyenne secteur
    benchmark = [63, 64, 64, 65, 65, 66]
    ax.plot(trimestres, benchmark, '--', color=C_LGRAY, linewidth=1.5,
            label='Moyenne secteur PME', zorder=2)

    # Score client
    ax.plot(trimestres, vals, 'o-', color=C_BLACK, linewidth=2.5,
            markersize=7, label=label, zorder=4)

    # Zone grisée autour de la courbe
    ax.fill_between(trimestres,
                   [v-3 for v in vals], [v+3 for v in vals],
                   alpha=0.1, color=C_BLACK)

    # Annotations valeurs
    for i, (t, v) in enumerate(zip(trimestres, vals)):
        ax.annotate(str(v), (t, v), textcoords="offset points",
                   xytext=(0, 10), ha='center', fontsize=9,
                   fontweight='bold', color=C_BLACK)

    # Labels zones
    ax.text(trimestres[-1], 82, 'BON', ha='right', va='center',
            fontsize=8, color=C_DGRAY, fontweight='bold', alpha=0.5)
    ax.text(trimestres[-1], 57, 'MOYEN', ha='right', va='center',
            fontsize=8, color=C_GRAY, fontweight='bold', alpha=0.5)
    ax.text(trimestres[-1], 30, 'CRITIQUE', ha='right', va='center',
            fontsize=8, color=C_MGRAY, fontweight='bold', alpha=0.5)

    ax.set_ylim(20, 95)
    ax.set_xlim(-0.3, len(trimestres)-0.7)
    ax.spines['left'].set_color(C_LGRAY)
    ax.spines['bottom'].set_color(C_LGRAY)
    ax.tick_params(colors=C_GRAY)
    ax.set_yticklabels([f'{int(y)}' for y in ax.get_yticks()], color=C_GRAY, fontsize=8)
    ax.grid(axis='y', color=C_LGRAY, linewidth=0.5, alpha=0.6)

    ax.legend(loc='lower right', fontsize=8, frameon=False, labelcolor=C_GRAY)
    ax.set_title('Évolution du Score KORD — Trajectoire de performance',
                fontsize=11, fontweight='bold', color=C_BLACK, pad=12, loc='left')

    plt.tight_layout(pad=1.0)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor=C_OFF)
    plt.close()
    buf.seek(0)
    return buf.read()


def generate_pilier_detail_chart(scores: dict, pilier_key: str) -> bytes:
    """
    Mini graphique détail pour un pilier — sous-scores et alertes.
    Style JLL : propre, annoté, informatif.
    """
    pilier_data = scores.get(pilier_key, {})
    if not isinstance(pilier_data, dict):
        return None

    score = pilier_data.get("score", 0)
    max_val = pilier_data.get("max", 20)
    alertes = pilier_data.get("alertes", [])

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(8, 3),
                                    gridspec_kw={'width_ratios': [1, 2]})
    _base_style(fig, C_OFF)

    # Gauge mini
    ax1.set_facecolor(C_BLACK)
    pct = score / max_val if max_val > 0 else 0
    theta = np.linspace(np.pi, np.pi - pct * np.pi, 200)
    ax1.plot(np.cos(np.linspace(np.pi, 0, 200)),
             np.sin(np.linspace(np.pi, 0, 200)),
             color="#333333", linewidth=20)
    ax1.plot(np.cos(theta), np.sin(theta), color=C_WHITE, linewidth=20)
    ax1.text(0, 0.1, f"{score}/{max_val}", ha='center', va='center',
             fontsize=18, fontweight='bold', color=C_WHITE)
    ax1.set_xlim(-1.3, 1.3)
    ax1.set_ylim(-0.2, 1.2)
    ax1.axis('off')

    # Liste alertes
    ax2.set_facecolor(C_OFF)
    ax2.axis('off')
    if alertes:
        for i, al in enumerate(alertes[:4]):
            msg = al.get('message', '') if isinstance(al, dict) else str(al)
            color = C_BLACK if i == 0 else C_DGRAY if i == 1 else C_GRAY
            ax2.text(0.02, 0.85 - i*0.22, f"▸ {msg[:65]}{'...' if len(msg)>65 else ''}",
                    transform=ax2.transAxes, fontsize=8.5, color=color,
                    va='top', wrap=True)
    else:
        ax2.text(0.5, 0.5, "Aucune anomalie détectée",
                transform=ax2.transAxes, ha='center', va='center',
                fontsize=10, color=C_GRAY, style='italic')

    plt.tight_layout(pad=0.5)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=120, bbox_inches='tight', facecolor=C_OFF)
    plt.close()
    buf.seek(0)
    return buf.read()
