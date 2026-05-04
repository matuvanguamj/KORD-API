"""
KORD ENGINE — Moteur d'analyse supply chain
Analyse les données opérationnelles et calcule le Score KORD sur 5 piliers.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Optional


# ════════════════════════════════════════════════════════
#  CONSTANTES — Pondération des 5 piliers
# ════════════════════════════════════════════════════════

PILIERS = {
    "stock_cash":          30,   # Stock & Cash immobilisé
    "transport_service":   20,   # Transport & Taux de service
    "achats_fournisseurs": 20,   # Achats & Fournisseurs
    "marges_retours":      15,   # Marges & Retours
    "donnees_pilotage":    15,   # Données & Pilotage
}


# ════════════════════════════════════════════════════════
#  CHARGEMENT DES DONNÉES
# ════════════════════════════════════════════════════════

def load_data(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Charge un fichier CSV ou Excel en DataFrame."""
    ext = filename.lower().split(".")[-1]
    try:
        if ext == "csv":
            # Essaie différents séparateurs
            for sep in [",", ";", "\t"]:
                try:
                    df = pd.read_csv(
                        pd.io.common.BytesIO(file_bytes),
                        sep=sep,
                        encoding="utf-8",
                        low_memory=False
                    )
                    if len(df.columns) > 1:
                        return df
                except Exception:
                    continue
            # Fallback
            return pd.read_csv(pd.io.common.BytesIO(file_bytes), encoding="latin-1")

        elif ext in ["xlsx", "xls"]:
            return pd.read_excel(pd.io.common.BytesIO(file_bytes))

    except Exception as e:
        raise ValueError(f"Impossible de lire le fichier {filename} : {str(e)}")

    return pd.DataFrame()


def detect_file_type(df: pd.DataFrame, filename: str) -> str:
    """Détecte le type de fichier selon les colonnes présentes."""
    cols = [c.lower().strip() for c in df.columns]
    cols_str = " ".join(cols)

    if any(k in cols_str for k in ["stock", "quantite", "qte", "inventory", "sku"]):
        return "stock"
    elif any(k in cols_str for k in ["expedition", "livraison", "transport", "colis", "poids"]):
        return "expeditions"
    elif any(k in cols_str for k in ["commande", "order", "achat", "fournisseur"]):
        return "commandes"
    elif any(k in cols_str for k in ["marge", "prix", "tarif", "ca", "chiffre"]):
        return "marges"
    elif any(k in cols_str for k in ["retour", "return", "avoir"]):
        return "retours"
    else:
        return "generique"


# ════════════════════════════════════════════════════════
#  ANALYSE PAR PILIER
# ════════════════════════════════════════════════════════

def analyse_stock(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Pilier 1 — Stock & Cash immobilisé (30 pts)
    Détecte : surstock, dormants, ruptures, rotation faible
    """
    results = {
        "score": 0,
        "max": 30,
        "nb_sku": 0,
        "nb_dormants": 0,
        "pct_dormants": 0,
        "nb_ruptures": 0,
        "valeur_dormants": 0,
        "rotation_moyenne": 0,
        "alertes": [],
        "opportunites": [],
    }

    cols = {c.lower().strip(): c for c in df.columns}

    # Identifier les colonnes clés
    qty_col   = next((cols[c] for c in cols if "quant" in c or "qte" in c or "stock" in c or "qty" in c), None)
    val_col   = next((cols[c] for c in cols if "valeur" in c or "value" in c or "montant" in c or "prix" in c), None)
    mvt_col   = next((cols[c] for c in cols if "mouvement" in c or "sortie" in c or "vente" in c or "mvt" in c), None)
    sku_col   = next((cols[c] for c in cols if "sku" in c or "ref" in c or "article" in c or "code" in c), None)

    if qty_col is None:
        results["alertes"].append("Colonne quantité non identifiée dans le fichier stock.")
        results["score"] = 12
        return results

    try:
        df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
        results["nb_sku"] = len(df)

        # Ruptures (stock = 0)
        ruptures = df[df[qty_col] <= 0]
        results["nb_ruptures"] = len(ruptures)

        # Dormants (aucun mouvement ou stock > seuil sans rotation)
        if mvt_col:
            df[mvt_col] = pd.to_numeric(df[mvt_col], errors="coerce").fillna(0)
            dormants = df[(df[qty_col] > 0) & (df[mvt_col] == 0)]
        else:
            # Sans colonne mouvement : top 20% des stocks les plus élevés = potentiellement dormants
            seuil = df[qty_col].quantile(0.80)
            dormants = df[df[qty_col] >= seuil]

        results["nb_dormants"] = len(dormants)
        if results["nb_sku"] > 0:
            results["pct_dormants"] = round(len(dormants) / results["nb_sku"] * 100, 1)

        # Valeur immobilisée
        if val_col:
            df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0)
            results["valeur_dormants"] = round(dormants[val_col].sum(), 0)

        # Alertes
        if results["pct_dormants"] > 25:
            results["alertes"].append(
                f"{results['pct_dormants']}% des références présentent un risque de dormance — cash potentiellement immobilisé."
            )
        if results["nb_ruptures"] > 0:
            pct_rupt = round(results["nb_ruptures"] / results["nb_sku"] * 100, 1)
            results["alertes"].append(
                f"{results['nb_ruptures']} références en rupture ({pct_rupt}% du catalogue)."
            )

        # Opportunités
        if results["valeur_dormants"] > 0:
            results["opportunites"].append(
                f"Libération potentielle estimée : {round(results['valeur_dormants'] * 0.3):,} à {round(results['valeur_dormants'] * 0.6):,} selon les arbitrages de déstockage."
            )

        # Scoring
        score = 30
        if results["pct_dormants"] > 30:  score -= 12
        elif results["pct_dormants"] > 15: score -= 7
        elif results["pct_dormants"] > 5:  score -= 3

        if results["nb_ruptures"] > 0:
            pct = results["nb_ruptures"] / results["nb_sku"] * 100
            if pct > 10:   score -= 8
            elif pct > 5:  score -= 4
            elif pct > 1:  score -= 2

        results["score"] = max(0, score)

    except Exception as e:
        results["alertes"].append(f"Erreur analyse stock : {str(e)}")
        results["score"] = 15

    return results


def analyse_expeditions(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Pilier 2 — Transport & Service (20 pts)
    Détecte : surfacturations, écarts poids, délais, anomalies
    """
    results = {
        "score": 0,
        "max": 20,
        "nb_expeditions": 0,
        "nb_anomalies": 0,
        "pct_anomalies": 0,
        "cout_total": 0,
        "cout_moyen": 0,
        "alertes": [],
        "opportunites": [],
    }

    cols = {c.lower().strip(): c for c in df.columns}

    cout_col   = next((cols[c] for c in cols if "cout" in c or "cost" in c or "montant" in c or "tarif" in c), None)
    poids_col  = next((cols[c] for c in cols if "poids" in c or "weight" in c or "kg" in c), None)
    delai_col  = next((cols[c] for c in cols if "delai" in c or "delay" in c or "jour" in c or "day" in c), None)
    statut_col = next((cols[c] for c in cols if "statut" in c or "status" in c or "livr" in c), None)

    results["nb_expeditions"] = len(df)

    try:
        if cout_col:
            df[cout_col] = pd.to_numeric(df[cout_col], errors="coerce").fillna(0)
            results["cout_total"] = round(df[cout_col].sum(), 0)
            results["cout_moyen"] = round(df[cout_col].mean(), 2)

            # Anomalies : coûts > moyenne + 2 écarts-types
            mean  = df[cout_col].mean()
            std   = df[cout_col].std()
            anomalies = df[df[cout_col] > mean + 2 * std]
            results["nb_anomalies"] = len(anomalies)

            if results["nb_expeditions"] > 0:
                results["pct_anomalies"] = round(len(anomalies) / results["nb_expeditions"] * 100, 1)

            if results["pct_anomalies"] > 5:
                surplus = round(anomalies[cout_col].sum() - (mean * len(anomalies)), 0)
                results["alertes"].append(
                    f"{results['pct_anomalies']}% des expéditions présentent des coûts anormalement élevés."
                )
                results["opportunites"].append(
                    f"Surcoût estimé sur expéditions atypiques : {surplus:,} à requalifier avec les transporteurs."
                )

        if poids_col and cout_col:
            df[poids_col] = pd.to_numeric(df[poids_col], errors="coerce").fillna(0)
            df["cout_par_kg"] = np.where(df[poids_col] > 0, df[cout_col] / df[poids_col], 0)
            cout_kg_mean = df[df["cout_par_kg"] > 0]["cout_par_kg"].mean()
            cout_kg_std  = df[df["cout_par_kg"] > 0]["cout_par_kg"].std()
            ecarts = df[df["cout_par_kg"] > cout_kg_mean + 2 * cout_kg_std]
            if len(ecarts) > 0:
                results["alertes"].append(
                    f"{len(ecarts)} expéditions présentent un coût au kg significativement supérieur à la moyenne."
                )

        # Scoring
        score = 20
        if results["pct_anomalies"] > 10:  score -= 10
        elif results["pct_anomalies"] > 5:  score -= 6
        elif results["pct_anomalies"] > 2:  score -= 3

        results["score"] = max(0, score)

    except Exception as e:
        results["alertes"].append(f"Erreur analyse expéditions : {str(e)}")
        results["score"] = 10

    return results


def analyse_commandes(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Pilier 3 — Achats & Fournisseurs (20 pts)
    Détecte : retards, erreurs, dépendances, délais anormaux
    """
    results = {
        "score": 0,
        "max": 20,
        "nb_commandes": 0,
        "nb_retards": 0,
        "pct_retards": 0,
        "fournisseurs_critiques": [],
        "alertes": [],
        "opportunites": [],
    }

    cols = {c.lower().strip(): c for c in df.columns}

    statut_col    = next((cols[c] for c in cols if "statut" in c or "status" in c or "etat" in c), None)
    fourn_col     = next((cols[c] for c in cols if "fourn" in c or "supplier" in c or "vendor" in c), None)
    montant_col   = next((cols[c] for c in cols if "montant" in c or "amount" in c or "valeur" in c or "total" in c), None)
    retard_col    = next((cols[c] for c in cols if "retard" in c or "delay" in c or "late" in c), None)

    results["nb_commandes"] = len(df)

    try:
        # Retards
        if retard_col:
            df[retard_col] = pd.to_numeric(df[retard_col], errors="coerce").fillna(0)
            retards = df[df[retard_col] > 0]
            results["nb_retards"] = len(retards)
        elif statut_col:
            retards = df[df[statut_col].astype(str).str.lower().str.contains("retard|late|delay|overdue", na=False)]
            results["nb_retards"] = len(retards)
        else:
            retards = pd.DataFrame()

        if results["nb_commandes"] > 0 and results["nb_retards"] > 0:
            results["pct_retards"] = round(results["nb_retards"] / results["nb_commandes"] * 100, 1)

        # Fournisseurs critiques (concentration des commandes)
        if fourn_col and montant_col:
            df[montant_col] = pd.to_numeric(df[montant_col], errors="coerce").fillna(0)
            total = df[montant_col].sum()
            par_fourn = df.groupby(fourn_col)[montant_col].sum().sort_values(ascending=False)
            top3_pct  = round(par_fourn.head(3).sum() / total * 100, 1) if total > 0 else 0
            if top3_pct > 70:
                results["fournisseurs_critiques"] = par_fourn.head(3).index.tolist()
                results["alertes"].append(
                    f"Concentration fournisseur élevée : les 3 premiers représentent {top3_pct}% du volume d'achats."
                )

        if results["pct_retards"] > 10:
            results["alertes"].append(
                f"{results['pct_retards']}% des commandes présentent des retards fournisseurs."
            )
            results["opportunites"].append(
                "Renégociation des conditions de livraison et mise en place de pénalités contractuelles recommandée."
            )

        # Scoring
        score = 20
        if results["pct_retards"] > 20:  score -= 12
        elif results["pct_retards"] > 10: score -= 7
        elif results["pct_retards"] > 5:  score -= 3

        if len(results["fournisseurs_critiques"]) > 0:
            score -= 4

        results["score"] = max(0, score)

    except Exception as e:
        results["alertes"].append(f"Erreur analyse commandes : {str(e)}")
        results["score"] = 10

    return results


def analyse_marges(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Pilier 4 — Marges & Retours (15 pts)
    Détecte : marges négatives, produits à perte, incohérences
    """
    results = {
        "score": 0,
        "max": 15,
        "nb_produits": 0,
        "nb_marges_neg": 0,
        "pct_marges_neg": 0,
        "marge_moyenne": 0,
        "alertes": [],
        "opportunites": [],
    }

    cols = {c.lower().strip(): c for c in df.columns}

    marge_col  = next((cols[c] for c in cols if "marge" in c or "margin" in c or "benefice" in c), None)
    prix_col   = next((cols[c] for c in cols if "prix" in c or "price" in c or "tarif" in c or "pvt" in c), None)
    cout_col   = next((cols[c] for c in cols if "cout" in c or "cost" in c or "revient" in c), None)
    retour_col = next((cols[c] for c in cols if "retour" in c or "return" in c or "avoir" in c), None)

    results["nb_produits"] = len(df)

    try:
        # Si colonne marge directe
        if marge_col:
            df[marge_col] = pd.to_numeric(df[marge_col], errors="coerce").fillna(0)
            neg = df[df[marge_col] < 0]
            results["nb_marges_neg"] = len(neg)
            results["marge_moyenne"]  = round(df[marge_col].mean(), 2)

        # Sinon calcul prix - coût
        elif prix_col and cout_col:
            df[prix_col] = pd.to_numeric(df[prix_col], errors="coerce").fillna(0)
            df[cout_col] = pd.to_numeric(df[cout_col], errors="coerce").fillna(0)
            df["_marge_calc"] = df[prix_col] - df[cout_col]
            neg = df[df["_marge_calc"] < 0]
            results["nb_marges_neg"] = len(neg)
            results["marge_moyenne"]  = round(df["_marge_calc"].mean(), 2)

        if results["nb_produits"] > 0 and results["nb_marges_neg"] > 0:
            results["pct_marges_neg"] = round(results["nb_marges_neg"] / results["nb_produits"] * 100, 1)

        if results["pct_marges_neg"] > 5:
            results["alertes"].append(
                f"{results['pct_marges_neg']}% des références présentent une marge négative ou nulle."
            )
            results["opportunites"].append(
                "Révision des tarifs ou des conditions d'achat sur les références à marge négative recommandée."
            )

        # Retours
        if retour_col:
            df[retour_col] = pd.to_numeric(df[retour_col], errors="coerce").fillna(0)
            total_retours = df[retour_col].sum()
            if total_retours > 0:
                results["alertes"].append(
                    f"Volume de retours clients détecté : {int(total_retours)} unités à analyser."
                )

        # Scoring
        score = 15
        if results["pct_marges_neg"] > 15:  score -= 10
        elif results["pct_marges_neg"] > 8:  score -= 6
        elif results["pct_marges_neg"] > 3:  score -= 3

        results["score"] = max(0, score)

    except Exception as e:
        results["alertes"].append(f"Erreur analyse marges : {str(e)}")
        results["score"] = 8

    return results


def analyse_donnees(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Pilier 5 — Données & Pilotage (15 pts)
    Évalue : qualité des données, complétude, cohérence
    """
    results = {
        "score": 0,
        "max": 15,
        "nb_lignes": len(df),
        "nb_colonnes": len(df.columns),
        "pct_vides": 0,
        "nb_doublons": 0,
        "alertes": [],
        "opportunites": [],
    }

    try:
        # Valeurs manquantes
        total_cells  = df.shape[0] * df.shape[1]
        missing      = df.isnull().sum().sum()
        results["pct_vides"] = round(missing / total_cells * 100, 1) if total_cells > 0 else 0

        # Doublons
        results["nb_doublons"] = df.duplicated().sum()

        if results["pct_vides"] > 20:
            results["alertes"].append(
                f"{results['pct_vides']}% des cellules sont vides — qualité des données à améliorer."
            )
        if results["nb_doublons"] > 0:
            results["alertes"].append(
                f"{results['nb_doublons']} lignes dupliquées détectées dans le fichier."
            )

        # Scoring
        score = 15
        if results["pct_vides"] > 30:   score -= 8
        elif results["pct_vides"] > 15: score -= 4
        elif results["pct_vides"] > 5:  score -= 2

        if results["nb_doublons"] > 0:
            pct_dup = round(results["nb_doublons"] / len(df) * 100, 1)
            if pct_dup > 10:  score -= 4
            elif pct_dup > 2: score -= 2

        results["score"] = max(0, score)
        results["opportunites"].append(
            f"Fichier analysé : {results['nb_lignes']:,} lignes × {results['nb_colonnes']} colonnes."
        )

    except Exception as e:
        results["alertes"].append(f"Erreur analyse données : {str(e)}")
        results["score"] = 8

    return results


# ════════════════════════════════════════════════════════
#  MOTEUR PRINCIPAL
# ════════════════════════════════════════════════════════

def run_audit(file_bytes: bytes, filename: str) -> Dict[str, Any]:
    """
    Point d'entrée principal.
    Analyse un fichier et retourne les résultats complets.
    """

    # 1. Chargement
    df = load_data(file_bytes, filename)
    if df.empty:
        return {"error": "Fichier vide ou non lisible."}

    # 2. Détection du type de fichier
    file_type = detect_file_type(df, filename)

    # 3. Analyse selon le type
    analyses = {}

    if file_type == "stock":
        analyses["stock_cash"]          = analyse_stock(df)
        analyses["donnees_pilotage"]     = analyse_donnees(df)
        # Piliers non couverts → score neutre
        analyses["transport_service"]    = _score_neutre(20)
        analyses["achats_fournisseurs"]  = _score_neutre(20)
        analyses["marges_retours"]       = _score_neutre(15)

    elif file_type == "expeditions":
        analyses["transport_service"]    = analyse_expeditions(df)
        analyses["donnees_pilotage"]     = analyse_donnees(df)
        analyses["stock_cash"]           = _score_neutre(30)
        analyses["achats_fournisseurs"]  = _score_neutre(20)
        analyses["marges_retours"]       = _score_neutre(15)

    elif file_type == "commandes":
        analyses["achats_fournisseurs"]  = analyse_commandes(df)
        analyses["donnees_pilotage"]     = analyse_donnees(df)
        analyses["stock_cash"]           = _score_neutre(30)
        analyses["transport_service"]    = _score_neutre(20)
        analyses["marges_retours"]       = _score_neutre(15)

    elif file_type == "marges":
        analyses["marges_retours"]       = analyse_marges(df)
        analyses["donnees_pilotage"]     = analyse_donnees(df)
        analyses["stock_cash"]           = _score_neutre(30)
        analyses["transport_service"]    = _score_neutre(20)
        analyses["achats_fournisseurs"]  = _score_neutre(20)

    else:
        # Fichier générique → analyse données + tentative multi
        analyses["stock_cash"]           = analyse_stock(df)
        analyses["transport_service"]    = analyse_expeditions(df)
        analyses["achats_fournisseurs"]  = analyse_commandes(df)
        analyses["marges_retours"]       = analyse_marges(df)
        analyses["donnees_pilotage"]     = analyse_donnees(df)

    # 4. Score global
    score_total = sum(a.get("score", 0) for a in analyses.values())
    score_total = min(100, max(0, score_total))

    # 5. Toutes les alertes
    toutes_alertes = []
    toutes_opportunites = []
    for pilier, data in analyses.items():
        for alerte in data.get("alertes", []):
            toutes_alertes.append({"pilier": pilier, "message": alerte})
        for opp in data.get("opportunites", []):
            toutes_opportunites.append({"pilier": pilier, "message": opp})

    # 6. Résultat complet
    return {
        "filename":      filename,
        "file_type":     file_type,
        "nb_lignes":     len(df),
        "nb_colonnes":   len(df.columns),
        "score_total":   score_total,
        "analyses":      analyses,
        "alertes":       toutes_alertes,
        "opportunites":  toutes_opportunites,
        "interpretation": interpret_score(score_total),
    }


def _score_neutre(max_pts: int) -> Dict[str, Any]:
    """Score neutre pour les piliers non couverts par le fichier."""
    return {
        "score":        round(max_pts * 0.6),   # 60% par défaut
        "max":          max_pts,
        "alertes":      [],
        "opportunites": ["Données non fournies pour ce pilier — analyse partielle."],
    }


def interpret_score(score: int) -> str:
    """Interprétation textuelle du score KORD."""
    if score >= 80:
        return "Performance opérationnelle solide. Peu de leviers majeurs détectés. Optimisation à la marge."
    elif score >= 65:
        return "Performance correcte avec des axes d'amélioration identifiés. Opportunités ciblées à saisir."
    elif score >= 50:
        return "Niveau de performance moyen. Plusieurs leviers significatifs détectés — actions prioritaires recommandées."
    elif score >= 35:
        return "Tensions opérationnelles importantes. Cash immobilisé et inefficiences à traiter en priorité."
    else:
        return "Niveau critique. Failles structurelles multiples détectées — intervention recommandée sans délai."
