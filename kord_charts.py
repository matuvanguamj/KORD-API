"""
KORD CHARTS — Génération des graphiques brandés KORD
"""
import io
import numpy as np

def generate_radar_chart(scores: dict) -> bytes:
    """Graphique radar des 5 piliers."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches

    piliers = {
        "stock_cash":          "Stock\nCash",
        "transport_service":   "Transport\nService",
        "achats_fournisseurs": "Achats\nFournisseurs",
        "marges_retours":      "Marges\nRetours",
        "donnees_pilotage":    "Données\nPilotage",
    }

    labels = list(piliers.values())
    vals   = []
    maxes  = []
    for key in piliers:
        a = scores.get(key, {})
        s = a.get("score", 0)
        m = a.get("max", 20)
        vals.append(s / m * 100)
        maxes.append(100)

    N = len(labels)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]
    vals   += vals[:1]

    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor('#F4F3F0')
    ax.set_facecolor('#F4F3F0')

    ax.plot(angles, vals, 'o-', linewidth=2, color='#000000')
    ax.fill(angles, vals, alpha=0.15, color='#000000')

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, size=8, fontfamily='sans-serif')
    ax.set_ylim(0, 100)
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(['25', '50', '75', '100'], size=6, color='#999999')
    ax.grid(color='#DDDDDD', linestyle='-', linewidth=0.5)
    ax.spines['polar'].set_color('#DDDDDD')

    plt.title('Score par pilier', size=10, fontweight='bold',
              fontfamily='sans-serif', pad=20, color='#000000')

    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                facecolor='#F4F3F0')
    plt.close()
    buf.seek(0)
    return buf.read()


def generate_bar_chart(scores: dict) -> bytes:
    """Graphique barres horizontales des scores par pilier."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    piliers = {
        "stock_cash":          "Stock et Cash immobilisé",
        "transport_service":   "Transport et Service",
        "achats_fournisseurs": "Achats et Fournisseurs",
        "marges_retours":      "Marges et Retours",
        "donnees_pilotage":    "Données et Pilotage",
    }

    labels, vals, maxes = [], [], []
    for key, label in piliers.items():
        a = scores.get(key, {})
        s = a.get("score", 0)
        m = a.get("max", 20)
        labels.append(label)
        vals.append(s)
        maxes.append(m)

    fig, ax = plt.subplots(figsize=(7, 3.5))
    fig.patch.set_facecolor('#F4F3F0')
    ax.set_facecolor('#F4F3F0')

    y = np.arange(len(labels))
    # Barre fond gris
    ax.barh(y, maxes, color='#E8E8E8', height=0.5, zorder=1)
    # Barre score noir
    colors = ['#000000' if v/m >= 0.7 else '#555555' if v/m >= 0.45 else '#999999'
              for v, m in zip(vals, maxes)]
    bars = ax.barh(y, vals, color=colors, height=0.5, zorder=2)

    # Labels
    for i, (v, m) in enumerate(zip(vals, maxes)):
        ax.text(m + 0.3, i, f'{v}/{m}', va='center', ha='left',
                fontsize=8, fontfamily='sans-serif', color='#333333', fontweight='bold')

    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8, fontfamily='sans-serif', color='#333333')
    ax.set_xlim(0, max(maxes) * 1.15)
    ax.set_xlabel('Score obtenu', fontsize=8, color='#666666')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#DDDDDD')
    ax.spines['bottom'].set_color('#DDDDDD')
    ax.tick_params(colors='#666666')
    ax.set_title('Scores par pilier KORD', fontsize=10, fontweight='bold',
                 fontfamily='sans-serif', color='#000000', pad=10)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                facecolor='#F4F3F0')
    plt.close()
    buf.seek(0)
    return buf.read()


def generate_gauge_chart(score: int) -> bytes:
    """Jauge du score global KORD."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches

    fig, ax = plt.subplots(figsize=(5, 3))
    fig.patch.set_facecolor('#000000')
    ax.set_facecolor('#000000')
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis('off')

    # Score
    ax.text(5, 3.5, str(score), fontsize=52, fontweight='bold',
            ha='center', va='center', color='#FFFFFF', fontfamily='sans-serif')
    ax.text(5, 1.8, '/100  —  SCORE KORD', fontsize=9,
            ha='center', va='center', color='#888888', fontfamily='sans-serif')

    # Barre de progression
    bar_bg = patches.FancyBboxPatch((1, 0.8), 8, 0.6,
                                     boxstyle="round,pad=0.05",
                                     facecolor='#333333', edgecolor='none')
    ax.add_patch(bar_bg)
    bar_fill = patches.FancyBboxPatch((1, 0.8), 8 * (score/100), 0.6,
                                       boxstyle="round,pad=0.05",
                                       facecolor='#FFFFFF', edgecolor='none')
    ax.add_patch(bar_fill)

    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                facecolor='#000000')
    plt.close()
    buf.seek(0)
    return buf.read()
