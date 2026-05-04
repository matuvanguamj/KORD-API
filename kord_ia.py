"""
KORD IA — Moteur de recommandations expert supply chain
Analyse approfondie — Rapport de qualité consultant senior
"""

import os
import json
from typing import Dict, Any, List

SYSTEM_PROMPT = """Tu es un consultant senior en performance opérationnelle et supply chain.
Tu rédiges des rapports dans le style des grands cabinets de conseil (JLL, McKinsey, Roland Berger) : 
narratifs, précis, fondés sur des données, accessibles à un dirigeant non-spécialiste.

MODÈLE DE STYLE À SUIVRE — Rapport JLL Marché Logistique :
"Une économie mondialisée de plus en plus interconnectée, associée aux progrès technologiques, a donné naissance 
à de nouveaux domaines d'activité. En 2021, il y a eu environ 2,6 fois plus de marchandises importées qu'exportées. 
Toutefois, la valeur des marchandises exportées était 1,3 fois supérieure à celle des importations."
→ Ce style : narratif fluide + chiffres précis intégrés dans le texte + contexte marché + implication business.

RÈGLES DE RÉDACTION ABSOLUES :
1. Chaque paragraphe commence par le fait le plus important, pas par une introduction générale
2. Les chiffres sont TOUJOURS intégrés dans le texte narratif, jamais isolés en bullet points
3. Chaque anomalie est connectée à une conséquence business concrète (trésorerie, marge, délais)
4. Les analyses croisent les fichiers entre eux — un chiffre du stock doit résonner avec les commandes
5. Le ton est celui d'un expert qui parle à un dirigeant : direct, respectueux, sans jargon inutile
6. Chaque section doit répondre à la question : "Qu'est-ce que ça coûte concrètement ?"
7. Les estimations financières sont réalistes — une fourchette vaut mieux qu'un chiffre faux
8. JAMAIS de phrases comme "il convient de", "il est recommandé de", "nous vous suggérons"
   → À la place : "La priorité est de...", "Le levier le plus immédiat est...", "Ce chiffre révèle..."

RETOURNE UNIQUEMENT CE JSON (sans markdown, sans texte avant ou après) :
{"resume_executif":{"paragraphe_1":"...","paragraphe_2":"...","paragraphe_3":"..."},"message_dirigeant":"...","analyse_piliers":{"stock_cash":{"titre":"Stock et Cash Immobilisé","score":0,"max":30,"niveau":"moyen","analyse":"5-8 lignes avec chiffres précis","chiffres_cles":["chiffre 1","chiffre 2"],"risque_principal":"..."},"transport_service":{"titre":"Transport et Taux de Service","score":0,"max":20,"niveau":"moyen","analyse":"...","chiffres_cles":[],"risque_principal":"..."},"achats_fournisseurs":{"titre":"Achats et Fournisseurs","score":0,"max":20,"niveau":"moyen","analyse":"...","chiffres_cles":[],"risque_principal":"..."},"marges_retours":{"titre":"Marges et Retours","score":0,"max":15,"niveau":"moyen","analyse":"...","chiffres_cles":[],"risque_principal":"..."},"donnees_pilotage":{"titre":"Données et Pilotage","score":0,"max":15,"niveau":"moyen","analyse":"...","chiffres_cles":[],"risque_principal":"..."}},"croisements_cles":[{"fichiers":["f1.csv","f2.csv"],"observation":"...","impact":"...€"}],"anomalies":[{"titre":"...","pilier":"stock_cash","detection":"...","impact_business":"...","impact_financier":"...€","urgence":"CRITIQUE"}],"priorites":[{"rang":1,"titre":"...","pilier":"stock_cash","probleme":"...","action":"...","impact_attendu":"...","gain_potentiel":"...€","delai":"Sous 30 jours","complexite":"Facile","quick_win":true}],"questions_restitution":["Q1","Q2","Q3"],"prochaine_etape":"...","benchmark":"...","points_vigilance":["pt1","pt2","pt3"],"opportunites_cachees":[{"titre":"...","description":"...","gain_estime":"...€"}]}"""


def generate_recommendations_global(consolidated: Dict[str, Any], all_results: list) -> Dict[str, Any]:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        return _fallback(consolidated)

    import anthropic
    client = anthropic.Anthropic(api_key=api_key)

    # S'assurer que all_results contient des dicts
    safe_results = [r for r in all_results if isinstance(r, dict)]

    pilier_labels = {
        "stock_cash":          "Stock et Cash immobilisé",
        "transport_service":   "Transport et Taux de service",
        "achats_fournisseurs": "Achats et Fournisseurs",
        "marges_retours":      "Marges et Retours",
        "donnees_pilotage":    "Données et Pilotage",
    }

    lines = [
        "=== DOSSIER CLIENT COMPLET — ANALYSE KORD ===",
        "",
        f"Nombre de fichiers : {len(safe_results)}",
        f"Fichiers reçus : {', '.join(r.get('file_name', '') for r in safe_results)}",
        f"Score KORD consolidé : {consolidated['score_total']}/100",
        f"Diagnostic initial : {consolidated.get('interpretation', '')}",
        "",
        "=== SCORES DÉTAILLÉS PAR PILIER ===",
    ]

    for key, label in pilier_labels.items():
        a = consolidated.get("analyses", {}).get(key, {})
        s, m = a.get("score", 0), a.get("max", 20)
        pct = round(s/m*100) if m > 0 else 0
        etat = "BON" if pct >= 70 else ("MOYEN" if pct >= 45 else "CRITIQUE")
        lines.append(f"  {label} : {s}/{m} pts ({pct}%) — {etat}")

    # Anomalies détaillées
    alertes = consolidated.get("alertes", [])
    if alertes:
        lines += ["", f"=== {len(alertes)} ANOMALIES DÉTECTÉES ==="]
        for a in alertes:
            p = pilier_labels.get(a.get("pilier", ""), a.get("pilier", "").upper())
            lines.append(f"  [{p}] {a.get('message', '')}")

    # Opportunités
    opps = [o for o in consolidated.get("opportunites", [])
            if "non fournies" not in o.get("message","") and "analysé" not in o.get("message","")]
    if opps:
        lines += ["", f"=== {len(opps)} OPPORTUNITÉS IDENTIFIÉES ==="]
        for o in opps:
            lines.append(f"  {o.get('message', '')}")

    # Données brutes par fichier
    lines += ["", "=== DONNÉES BRUTES PAR FICHIER ==="]
    for r in safe_results:
        fname = r.get('file_name', '')
        ftype = r.get('file_type', '')
        score = r.get('score_total', 0)
        lines.append(f"\n--- {fname} (type: {ftype}, score: {score}/100) ---")

        for pilier, data in r.get("analyses", {}).items():
            if not isinstance(data, dict):
                continue
            ps = data.get("score", 0)
            pm = data.get("max", 20)
            label = pilier_labels.get(pilier, pilier)
            lines.append(f"  {label} : {ps}/{pm}")

            # Alertes du fichier
            for al in data.get("alertes", [])[:5]:
                if isinstance(al, dict):
                    lines.append(f"    → ALERTE : {al.get('message', '')}")

            # Données brutes si disponibles
            raw = data.get("raw_data", {})
            if raw:
                for k, v in list(raw.items())[:5]:
                    lines.append(f"    • {k} : {v}")

        # KPIs calculés
        kpis = r.get("kpis", {})
        if kpis:
            lines.append("  KPIs calculés :")
            for k, v in list(kpis.items())[:8]:
                lines.append(f"    {k} : {v}")

    context = "\n".join(lines)
    context += f"""

=== INSTRUCTIONS SPÉCIFIQUES ===
Score global : {consolidated['score_total']}/100
Tu dois produire un rapport de qualité consultant senior.
Chaque analyse doit citer les chiffres réels trouvés dans les données.
Croise les fichiers entre eux pour identifier des corrélations.
Estime les impacts financiers de manière réaliste.
Retourne UNIQUEMENT le JSON demandé, sans markdown ni texte avant ou après."""

    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": context}]
        )
        raw = message.content[0].text.strip()

        # Parser robuste
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

        # Nettoyage agressif avant parsing
        # Supprimer les caractères de contrôle sauf 
 	
        import re as _re
        raw = _re.sub(r'[--]', '', raw)
        # Trouver le JSON valide
        start = raw.find('{')
        end   = raw.rfind('}') + 1
        if start >= 0 and end > start:
            raw = raw[start:end]
        result = json.loads(raw)
        print("Recommandations IA générées avec succès")
        return result

    except json.JSONDecodeError as e:
        print(f"Erreur JSON IA : {e} — utilisation du fallback")
        return _fallback(consolidated)
    except Exception as e:
        print(f"Erreur IA : {e}")
        return _fallback(consolidated)


def generate_recommendations(audit_results: Dict[str, Any]) -> Dict[str, Any]:
    """Compatibilité route legacy."""
    if not isinstance(audit_results, dict):
        return _fallback({})
    return generate_recommendations_global(audit_results, [audit_results])


def _fallback(consolidated: Dict[str, Any]) -> Dict[str, Any]:
    score = consolidated.get("score_total", 0) if isinstance(consolidated, dict) else 0
    return {
        "resume_executif": {
            "paragraphe_1": f"Le Score KORD de {score}/100 révèle des axes d'amélioration significatifs dans votre performance opérationnelle.",
            "paragraphe_2": "L'analyse de vos données a permis d'identifier plusieurs opportunités de gain sur les piliers stock et transport.",
            "paragraphe_3": "Notre équipe estime un potentiel de récupération à quantifier lors de la session de restitution."
        },
        "message_dirigeant": "Vos données révèlent des leviers concrets. Nous vous recommandons de prioriser le pilier stock qui représente le potentiel le plus important.",
        "analyse_piliers": {},
        "croisements_cles": [],
        "anomalies": [],
        "priorites": [{
            "rang": 1, "titre": "Analyse à approfondir",
            "pilier": "stock_cash",
            "probleme": "Données en cours d'analyse approfondie.",
            "action": "Session de restitution KORD à planifier.",
            "impact_attendu": "À quantifier lors de la restitution.",
            "gain_potentiel": "À estimer",
            "delai": "Sous 30 jours",
            "complexite": "Modéré",
            "quick_win": False
        }],
        "questions_restitution": [
            "Quelle est votre valeur de stock actuelle ?",
            "Quel est votre taux de retour moyen ?",
            "Quels sont vos principaux fournisseurs ?"
        ],
        "prochaine_etape": "Planifier la session de restitution KORD pour approfondir l'analyse.",
        "benchmark": "Score en cours d'étalonnage selon votre secteur.",
        "points_vigilance": ["Analyse en cours de finalisation."],
        "opportunites_cachees": []
    }
