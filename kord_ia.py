"""
KORD IA — Moteur de recommandations expert supply chain
Claude Sonnet 4.6 — Analyse globale consolidée
"""

import os
import json
from typing import Dict, Any

SYSTEM_PROMPT = """Tu es un expert en performance opérationnelle et supply chain avec 20 ans d'expérience en conseil pour des PME et ETI.
Tu analyses des données opérationnelles (stock, transport, achats, marges, retours) et tu rédiges des rapports d'audit professionnels.

TON STYLE :
- Ton d'expert mais accessible — tu parles à un dirigeant, pas à un data scientist
- Chaque anomalie est traduite en impact financier concret (€ ou %)
- Tu utilises des exemples concrets et des comparaisons de marché
- Tu priorises par urgence et impact financier
- Tu es direct, précis, sans jargon inutile
- Tu commences toujours par "Ce qu'on a trouvé dans vos données" pas par des généralités

STRUCTURE JSON OBLIGATOIRE (retourne UNIQUEMENT le JSON, rien d'autre) :
{
  "resume_executif": "3-4 phrases percutantes. Commence par le fait le plus important. Mentionne le score et ce qu'il signifie concrètement. Donne une estimation du potentiel de gain en €.",
  "message_dirigeant": "1 paragraphe personnel au dirigeant. Ton conseil principal. Ce que tu ferais en priorité si tu étais à sa place.",
  "priorites": [
    {
      "rang": 1,
      "titre": "Titre court et percutant",
      "pilier": "nom_pilier",
      "probleme": "Description précise du problème détecté dans les données. Citez les chiffres trouvés.",
      "impact_financier": "Estimation chiffrée de l'impact en € ou % de CA/marge. Soyez précis.",
      "action": "Action concrète à mener. Qui fait quoi, dans quel délai.",
      "exemple": "Exemple concret ou analogie pour illustrer le problème.",
      "gain_potentiel": "Fourchette de gain en € si l'action est menée.",
      "delai": "Court terme (0-3 mois) / Moyen terme (3-6 mois) / Long terme (6-12 mois)",
      "complexite": "Facile / Modéré / Complexe"
    }
  ],
  "synthese_piliers": {
    "stock_cash": {"commentaire": "Analyse détaillée du pilier avec chiffres précis", "niveau": "bon/moyen/critique"},
    "transport_service": {"commentaire": "...", "niveau": "..."},
    "achats_fournisseurs": {"commentaire": "...", "niveau": "..."},
    "marges_retours": {"commentaire": "...", "niveau": "..."},
    "donnees_pilotage": {"commentaire": "...", "niveau": "..."}
  },
  "points_vigilance": ["Point 1 avec chiffre", "Point 2 avec chiffre", "Point 3 avec chiffre"],
  "opportunites_cachees": [
    {
      "titre": "Titre de l'opportunité",
      "description": "Description détaillée de l'opportunité cachée dans les données",
      "gain_estime": "Estimation du gain en €"
    }
  ],
  "prochaine_etape": "Action #1 à faire dans les 7 prochains jours. Très concrète.",
  "benchmark": "Comparaison avec les standards du secteur pour contextualiser le score."
}"""


def generate_recommendations_global(consolidated: Dict[str, Any], all_results: list) -> Dict[str, Any]:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        return _fallback(consolidated)

    import anthropic
    client = anthropic.Anthropic(api_key=api_key)

    pilier_labels = {
        "stock_cash":          "Stock et Cash immobilisé",
        "transport_service":   "Transport et Taux de service",
        "achats_fournisseurs": "Achats et Fournisseurs",
        "marges_retours":      "Marges et Retours",
        "donnees_pilotage":    "Données et Pilotage",
    }

    # S'assurer que all_results contient des dicts
    safe_results = [r for r in all_results if isinstance(r, dict)]
    if not safe_results:
        safe_results = all_results if all_results else []

    lines = [
        "=== ANALYSE KORD — DOSSIER CLIENT COMPLET ===",
        "",
        f"Fichiers analysés : {len(safe_results)}",
        f"Types : {', '.join(r.get('file_type', r.get('file_name','')) if isinstance(r,dict) else str(r) for r in safe_results)}",
        f"Score KORD global : {consolidated['score_total']}/100",
        f"Diagnostic : {consolidated.get('interpretation', '')}",
        "",
        "=== SCORES PAR PILIER ===",
    ]

    for key, label in pilier_labels.items():
        a = consolidated.get("analyses", {}).get(key, {})
        s, m = a.get("score", 0), a.get("max", 20)
        pct = round(s/m*100) if m > 0 else 0
        etat = "BON" if pct >= 70 else ("MOYEN" if pct >= 45 else "CRITIQUE")
        lines.append(f"  {label} : {s}/{m} pts ({pct}%) — {etat}")

    alertes = consolidated.get("alertes", [])
    if alertes:
        lines += ["", f"=== ANOMALIES DÉTECTÉES ({len(alertes)}) ==="]
        for a in alertes[:15]:
            p = pilier_labels.get(a.get("pilier", ""), a.get("pilier", ""))
            lines.append(f"  [{p.upper()}] {a.get('message', '')}")

    opps = consolidated.get("opportunites", [])
    if opps:
        lines += ["", f"=== OPPORTUNITÉS ({len(opps)}) ==="]
        for o in opps[:8]:
            msg = o.get("message", "")
            if "non fournies" not in msg and "analysé" not in msg:
                lines.append(f"  {msg}")

    # Données brutes par fichier
    lines += ["", "=== DÉTAIL PAR FICHIER ==="]
    for r in safe_results:
        if not isinstance(r, dict): continue
        lines.append(f"\nFichier : {r.get('file_name', '')} — score {r.get('score_total', 0)}/100")
        for pilier, data in r.get("analyses", {}).items():
            if data.get("alertes"):
                for al in data["alertes"][:3]:
                    lines.append(f"  → {al.get('message', '')}")

    context = "\n".join(lines)
    context += "\n\nSur la base de cette analyse complète, génère le rapport expert KORD. Traduis chaque anomalie en impact financier concret. Sois précis sur les chiffres et les gains potentiels. Retourne UNIQUEMENT le JSON demandé."

    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=3000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": context}]
        )
        raw = message.content[0].text.strip()
        # Nettoyer
        if "```" in raw:
            for part in raw.split("```"):
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("{"):
                    raw = part
                    break
        start, end = raw.find("{"), raw.rfind("}") + 1
        if start >= 0 and end > start:
            raw = raw[start:end]
        return json.loads(raw)
    except Exception as e:
        print(f"Erreur IA : {e}")
        return _fallback(consolidated)


def _fallback(consolidated: Dict[str, Any]) -> Dict[str, Any]:
    score = consolidated.get("score_total", 0)
    return {
        "resume_executif": f"L'analyse de vos données opérationnelles révèle un Score KORD de {score}/100. Plusieurs leviers de performance ont été identifiés dans vos flux stock et transport.",
        "message_dirigeant": "Notre équipe a identifié des opportunités concrètes dans vos données. Nous vous recommandons de prioriser les actions sur le pilier stock qui représente le levier le plus important.",
        "priorites": [{
            "rang": 1, "titre": "Optimisation prioritaire identifiée", "pilier": "stock_cash",
            "probleme": "Anomalies détectées dans vos données de stock.",
            "impact_financier": "À quantifier lors de la session de restitution.",
            "action": "Session de restitution à planifier avec l'équipe KORD.",
            "exemple": "Des références dormantes représentent en moyenne 15-20% du stock valorisé.",
            "gain_potentiel": "À estimer selon votre volume de stock.",
            "delai": "Court terme", "complexite": "Modéré"
        }],
        "synthese_piliers": {},
        "points_vigilance": ["Analyse en cours de finalisation."],
        "opportunites_cachees": [],
        "prochaine_etape": "Planifier la session de restitution KORD.",
        "benchmark": "Score en cours d'étalonnage."
    }


def generate_recommendations(audit_results: Dict[str, Any]) -> Dict[str, Any]:
    """Compatibilité route legacy."""
    return generate_recommendations_global(audit_results, [audit_results])
