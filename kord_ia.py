"""
KORD IA — Génération de recommandations supply chain par Claude (Anthropic)
Nourri avec la méthode KORD et l'expertise supply chain.
"""

import os
import json
from typing import Dict, Any
import anthropic


# ════════════════════════════════════════════════════════════════
#  SYSTÈME — Cerveau supply chain de KORD
# ════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """
Tu es le cerveau analytique de KORD, une plateforme d'audit en performance opérationnelle et supply chain.

Tu es un expert senior en :
- Gestion des stocks et optimisation du BFR (Besoin en Fonds de Roulement)
- Transport, logistique et pilotage des flux
- Achats stratégiques et relations fournisseurs
- Analyse des marges et rentabilité produit
- Data quality et pilotage par les KPIs opérationnels

Ta mission est de transformer des résultats d'analyse quantitative en recommandations business concrètes, priorisées et actionnables pour des dirigeants de PME et ETI.

RÈGLES ABSOLUES :
1. Tu parles toujours en termes business, jamais en termes techniques ou informatiques.
2. Tu ne mentionnes jamais Python, pandas, algorithmes, API, ou tout autre outil technique.
3. Tu quantifies les gains potentiels quand les données le permettent.
4. Tu priorises les actions par impact financier décroissant.
5. Tu es direct, précis, sans jargon inutile.
6. Tu ne fais jamais de promesses impossibles. Tu parles de "potentiel estimé" et "levier à confirmer".
7. Tes recommandations sont réalistes et ancrées dans la réalité terrain.
8. Tu écris en français professionnel, neutre, sans formules creuses.
9. Tu ne dis jamais "il est important de" ou "il convient de". Tu dis ce qu'il faut faire.
10. Aucun tiret (-) dans tes réponses. Utilise des points ou des virgules.
11. Chaque recommandation doit être directement actionnable par un directeur supply chain ou un DAF.

FORMAT DE RÉPONSE :
Tu retournes uniquement un objet JSON valide avec cette structure exacte, sans markdown, sans backticks :
{
  "resume_executif": "3 à 5 phrases résumant les conclusions principales et l'enjeu financier global",
  "priorites": [
    {
      "rang": 1,
      "titre": "Titre court et percutant",
      "description": "Description précise de la problématique identifiée dans les données",
      "action": "Action concrète et immédiate à mener",
      "impact": "Impact financier estimé ou bénéfice attendu chiffré si possible",
      "delai": "Court terme / Moyen terme / Long terme"
    }
  ],
  "points_vigilance": ["Point 1", "Point 2", "Point 3"],
  "prochaine_etape": "Une seule action à mener dans les 30 prochains jours"
}
"""


# ════════════════════════════════════════════════════════════════
#  GÉNÉRATION DES RECOMMANDATIONS
# ════════════════════════════════════════════════════════════════

def generate_recommendations(audit_results: Dict[str, Any]) -> Dict[str, Any]:
    """
    Génère les recommandations via Claude.
    Transforme les résultats quantitatifs en insights business actionnables.
    """

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("Clé Anthropic manquante — recommandations de secours utilisées.")
        return _fallback_recommendations(audit_results)

    client = anthropic.Anthropic(api_key=api_key)
    context = _build_context(audit_results)

    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2000,
            system=SYSTEM_PROMPT,
            messages=[
                {"role": "user", "content": context}
            ]
        )

        raw = message.content[0].text.strip()

        # Nettoyer si Claude ajoute des backticks malgré les instructions
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()

        recommendations = json.loads(raw)
        return recommendations

    except json.JSONDecodeError as e:
        print(f"Erreur parsing JSON Claude : {e}")
        return _fallback_recommendations(audit_results)

    except Exception as e:
        print(f"Erreur Claude API : {e}")
        return _fallback_recommendations(audit_results)


# ════════════════════════════════════════════════════════════════
#  CONSTRUCTION DU CONTEXTE
# ════════════════════════════════════════════════════════════════

def _build_context(results: Dict[str, Any]) -> str:
    """
    Construit le prompt contextuel enrichi pour Claude.
    Plus le contexte est précis, meilleures sont les recommandations.
    """

    score     = results.get("score_total", 0)
    filename  = results.get("filename", "")
    nb_lignes = results.get("nb_lignes", 0)
    file_type = results.get("file_type", "")
    interp    = results.get("interpretation", "")
    analyses  = results.get("analyses", {})
    alertes   = results.get("alertes", [])
    opps      = results.get("opportunites", [])

    pilier_labels = {
        "stock_cash":          "Stock et Cash immobilisé",
        "transport_service":   "Transport et Taux de service",
        "achats_fournisseurs": "Achats et Fournisseurs",
        "marges_retours":      "Marges et Retours clients",
        "donnees_pilotage":    "Données et Pilotage",
    }

    lines = [
        "RAPPORT D'ANALYSE KORD — À TRAITER",
        "",
        f"Fichier : {filename}",
        f"Catégorie détectée : {file_type}",
        f"Volume de données : {nb_lignes:,} lignes",
        f"Score KORD global : {score}/100",
        f"Diagnostic : {interp}",
        "",
        "SCORES PAR PILIER :",
    ]

    for key, label in pilier_labels.items():
        if key in analyses:
            a     = analyses[key]
            s     = a.get("score", 0)
            m     = a.get("max", 0)
            pct   = round(s / m * 100) if m > 0 else 0
            etat  = "Bon" if pct >= 70 else ("Moyen" if pct >= 45 else "Critique")
            lines.append(f"  {label} : {s}/{m} pts ({pct}%) — {etat}")

            # Données détaillées par pilier
            if key == "stock_cash":
                nb_sku  = a.get("nb_sku", 0)
                dormants = a.get("pct_dormants", 0)
                rupt    = a.get("nb_ruptures", 0)
                val_dorm = a.get("valeur_dormants", 0)
                if nb_sku:    lines.append(f"    Références analysées : {nb_sku:,}")
                if dormants:  lines.append(f"    Part dormante estimée : {dormants}%")
                if rupt:      lines.append(f"    Ruptures détectées : {rupt}")
                if val_dorm:  lines.append(f"    Valeur stock dormant : {val_dorm:,.0f}")

            elif key == "transport_service":
                nb_exp  = a.get("nb_expeditions", 0)
                anom    = a.get("pct_anomalies", 0)
                cout_t  = a.get("cout_total", 0)
                cout_m  = a.get("cout_moyen", 0)
                if nb_exp:  lines.append(f"    Expéditions analysées : {nb_exp:,}")
                if anom:    lines.append(f"    Expéditions anormales : {anom}%")
                if cout_t:  lines.append(f"    Coût transport total : {cout_t:,.0f}")
                if cout_m:  lines.append(f"    Coût moyen par expédition : {cout_m:.2f}")

            elif key == "achats_fournisseurs":
                nb_cmd  = a.get("nb_commandes", 0)
                retards = a.get("pct_retards", 0)
                crit    = a.get("fournisseurs_critiques", [])
                if nb_cmd:  lines.append(f"    Commandes analysées : {nb_cmd:,}")
                if retards: lines.append(f"    Taux de retard fournisseurs : {retards}%")
                if crit:    lines.append(f"    Fournisseurs concentrant les risques : {', '.join(str(f) for f in crit[:3])}")

            elif key == "marges_retours":
                nb_prod = a.get("nb_produits", 0)
                neg     = a.get("pct_marges_neg", 0)
                moy     = a.get("marge_moyenne", 0)
                if nb_prod: lines.append(f"    Références analysées : {nb_prod:,}")
                if neg:     lines.append(f"    Références à marge négative : {neg}%")
                if moy:     lines.append(f"    Marge moyenne : {moy:.2f}")

            elif key == "donnees_pilotage":
                vides  = a.get("pct_vides", 0)
                dup    = a.get("nb_doublons", 0)
                if vides: lines.append(f"    Taux de données manquantes : {vides}%")
                if dup:   lines.append(f"    Doublons détectés : {dup}")

    if alertes:
        lines.append("")
        lines.append(f"ANOMALIES DÉTECTÉES ({len(alertes)}) :")
        for a in alertes:
            lines.append(f"  [{pilier_labels.get(a['pilier'], a['pilier'])}] {a['message']}")

    if opps:
        lines.append("")
        lines.append("OPPORTUNITÉS IDENTIFIÉES :")
        for o in opps:
            if "Données non fournies" not in o["message"] and "Fichier analysé" not in o["message"]:
                lines.append(f"  {o['message']}")

    lines.append("")
    lines.append(
        "Sur la base de cette analyse, génère le rapport de recommandations KORD complet. "
        "Priorise par impact financier. Sois précis sur les enjeux chiffrés. "
        "Retourne uniquement le JSON demandé, sans aucun texte avant ou après."
    )

    return "\n".join(lines)


# ════════════════════════════════════════════════════════════════
#  FALLBACK
# ════════════════════════════════════════════════════════════════

def _fallback_recommendations(results: Dict[str, Any]) -> Dict[str, Any]:
    """Recommandations de secours si Claude n'est pas disponible."""

    score   = results.get("score_total", 0)
    alertes = results.get("alertes", [])
    opps    = results.get("opportunites", [])
    interp  = results.get("interpretation", "")

    priorites = []
    for i, a in enumerate(alertes[:4], 1):
        pilier_clean = a["pilier"].replace("_", " ").title()
        priorites.append({
            "rang":        i,
            "titre":       f"Anomalie détectée sur {pilier_clean}",
            "description": a["message"],
            "action":      "Analyse approfondie recommandée en priorité.",
            "impact":      "À quantifier lors de la restitution.",
            "delai":       "Court terme"
        })

    if not priorites:
        priorites = [{
            "rang":        1,
            "titre":       "Rapport en cours de préparation",
            "description": "Les données ont été analysées. La restitution KORD est en cours.",
            "action":      "Attendre la restitution détaillée.",
            "impact":      "À définir.",
            "delai":       "Court terme"
        }]

    points = [
        o["message"] for o in opps
        if "Données non fournies" not in o["message"] and "Fichier analysé" not in o["message"]
    ][:3]
    if not points:
        points = ["Analyse complète disponible lors de la restitution."]

    return {
        "resume_executif":  (
            f"L'analyse KORD a produit un score de {score}/100. "
            f"{interp} "
            f"{len(alertes)} anomalie(s) ont été identifiées dans les données transmises. "
            f"La restitution permettra de quantifier précisément chaque levier."
        ),
        "priorites":        priorites,
        "points_vigilance": points,
        "prochaine_etape":  "Planifier la session de restitution avec l'équipe KORD."
    }


# ════════════════════════════════════════════════════════════════
#  RECOMMANDATIONS GLOBALES — Tous fichiers consolidés
# ════════════════════════════════════════════════════════════════

def generate_recommendations_global(consolidated: dict, all_results: list) -> dict:
    """
    Génère les recommandations IA sur l'ensemble des fichiers d'un client.
    Claude analyse tout en croisant les données de tous les fichiers.
    """
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        return _fallback_recommendations(consolidated)

    import anthropic
    client = anthropic.Anthropic(api_key=api_key)

    # Construire un contexte enrichi avec TOUS les fichiers
    lines = [
        "ANALYSE GLOBALE KORD — DOSSIER CLIENT COMPLET",
        "",
        f"Nombre de fichiers analysés : {len(all_results)}",
        f"Types de données : {', '.join(r.get('file_type', '') for r in all_results)}",
        f"Score KORD global : {consolidated['score_total']}/100",
        f"Diagnostic : {consolidated.get('interpretation', '')}",
        "",
        "SCORES CONSOLIDÉS PAR PILIER :",
    ]

    pilier_labels = {
        "stock_cash":          "Stock et Cash immobilisé",
        "transport_service":   "Transport et Taux de service",
        "achats_fournisseurs": "Achats et Fournisseurs",
        "marges_retours":      "Marges et Retours",
        "donnees_pilotage":    "Données et Pilotage",
    }

    for key, label in pilier_labels.items():
        if key in consolidated.get("analyses", {}):
            a = consolidated["analyses"][key]
            s, m = a.get("score", 0), a.get("max", 0)
            pct = round(s/m*100) if m > 0 else 0
            etat = "Bon" if pct >= 70 else ("Moyen" if pct >= 45 else "Critique")
            lines.append(f"  {label} : {s}/{m} pts ({pct}%) — {etat}")

    if consolidated.get("alertes"):
        lines.append("")
        lines.append(f"ANOMALIES DÉTECTÉES SUR L'ENSEMBLE DES DONNÉES ({len(consolidated['alertes'])}) :")
        for a in consolidated["alertes"][:8]:
            lines.append(f"  [{a.get('pilier','').upper()}] {a.get('message','')}")

    if consolidated.get("opportunites"):
        lines.append("")
        lines.append("OPPORTUNITÉS IDENTIFIÉES :")
        for o in consolidated["opportunites"][:5]:
            if "non fournies" not in o.get("message","") and "analysé" not in o.get("message",""):
                lines.append(f"  {o.get('message','')}")

    lines.append("")
    lines.append(
        "Sur la base de cette analyse complète du dossier client (tous fichiers croisés), "
        "génère le rapport de recommandations KORD consolidé. "
        "Croise les données entre les différents fichiers pour identifier les corrélations. "
        "Priorise par impact financier. Retourne uniquement le JSON demandé."
    )

    context = "\n".join(lines)

    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": context}]
        )
        raw = message.content[0].text.strip()
        # Nettoyer les backticks
        if "```" in raw:
            parts = raw.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("{"):
                    raw = part
                    break
        raw = raw.strip()
        # Trouver le JSON entre { et }
        start = raw.find("{")
        end   = raw.rfind("}") + 1
        if start >= 0 and end > start:
            raw = raw[start:end]
        import json as json_module
        return json_module.loads(raw)
    except Exception as e:
        print(f"Erreur IA globale : {e}")
        return _fallback_recommendations(consolidated)
