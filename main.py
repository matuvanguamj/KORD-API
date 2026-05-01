"""
KORD API — Serveur FastAPI
Reçoit les fichiers, lance le moteur d'analyse et sauvegarde les résultats.
"""

import os
import json
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv
from supabase import create_client, Client

from kord_engine import run_audit, load_data
from kord_ia import generate_recommendations

load_dotenv()

# ════════════════════════════════════════════════════════
#  CONFIGURATION
# ════════════════════════════════════════════════════════

SUPABASE_URL         = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")

sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

app = FastAPI(title="KORD API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],   # En prod : mettre ton domaine Vercel
    allow_methods=["*"],
    allow_headers=["*"],
)


# ════════════════════════════════════════════════════════
#  MODÈLES
# ════════════════════════════════════════════════════════

class AnalyzeRequest(BaseModel):
    submission_id: str
    file_path:     str
    file_name:     str
    user_id:       str


# ════════════════════════════════════════════════════════
#  ROUTES
# ════════════════════════════════════════════════════════

@app.get("/")
def root():
    return {"status": "KORD API en ligne", "version": "1.0.0"}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/analyze")
async def analyze(req: AnalyzeRequest):
    """
    Déclenche l'analyse d'un fichier client.
    1. Télécharge le fichier depuis Supabase Storage
    2. Lance le moteur d'analyse KORD
    3. Génère les recommandations IA
    4. Sauvegarde les résultats en base
    5. Met à jour le statut du dépôt
    """

    print(f"Analyse démarrée : {req.file_name} (user: {req.user_id})")

    # ── 1. Mettre le statut en "en_analyse" ──
    sb.table("submissions").update(
        {"status": "en_analyse"}
    ).eq("id", req.submission_id).execute()

    # ── 2. Télécharger le fichier depuis Supabase Storage ──
    try:
        response = sb.storage.from_("client-files").download(req.file_path)
        file_bytes = response
    except Exception as e:
        print(f"Erreur téléchargement : {e}")
        sb.table("submissions").update(
            {"status": "erreur"}
        ).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=500, detail=f"Impossible de télécharger le fichier : {str(e)}")

    # ── 3. Lancer le moteur d'analyse KORD ──
    try:
        audit_results = run_audit(file_bytes, req.file_name)
    except Exception as e:
        print(f"Erreur moteur : {e}")
        sb.table("submissions").update(
            {"status": "erreur"}
        ).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=500, detail=f"Erreur moteur d'analyse : {str(e)}")

    if "error" in audit_results:
        sb.table("submissions").update(
            {"status": "erreur"}
        ).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=400, detail=audit_results["error"])

    # ── 4. Générer les recommandations IA ──
    try:
        recommendations = generate_recommendations(audit_results)
    except Exception as e:
        print(f"Erreur IA : {e}")
        recommendations = {"resume_executif": "Recommandations en cours de génération.", "priorites": []}

    # ── 5. Sauvegarder le rapport en base ──
    score    = audit_results.get("score_total", 0)
    summary  = recommendations.get("resume_executif", "")

    # Données complètes (pour consultation interne)
    full_data = {
        "audit":           audit_results,
        "recommendations": recommendations,
    }

    try:
        sb.table("reports").insert({
            "submission_id": req.submission_id,
            "user_id":       req.user_id,
            "score_kord":    score,
            "summary":       summary,
            "report_url":    None,   # Sera rempli manuellement par l'équipe KORD
        }).execute()
    except Exception as e:
        print(f"Erreur sauvegarde rapport : {e}")

    # ── 6. Sauvegarder les données complètes (optionnel) ──
    # Tu peux sauvegarder le JSON complet si tu veux y accéder plus tard
    try:
        sb.table("submissions").update({
            "status":    "en_analyse",   # Reste en_analyse — toi tu passes à "terminé" manuellement
        }).eq("id", req.submission_id).execute()
    except Exception as e:
        print(f"Erreur mise à jour statut : {e}")

    print(f"Analyse terminée : score {score}/100 — {len(audit_results.get('alertes', []))} alertes")

    return {
        "status":       "ok",
        "score":        score,
        "nb_alertes":   len(audit_results.get("alertes", [])),
        "interpretation": audit_results.get("interpretation", ""),
        "resume":       summary,
    }


@app.post("/webhook/supabase")
async def webhook(payload: dict):
    """
    Webhook Supabase — déclenché automatiquement quand un fichier est déposé.
    Configure dans Supabase : Database > Webhooks > table submissions > INSERT
    """
    try:
        record = payload.get("record", {})

        submission_id = record.get("id")
        file_path     = record.get("file_path")
        file_name     = record.get("file_name")
        user_id       = record.get("user_id")

        if not all([submission_id, file_path, file_name, user_id]):
            return {"status": "ignored", "reason": "Données manquantes"}

        # Déclencher l'analyse
        req = AnalyzeRequest(
            submission_id=submission_id,
            file_path=file_path,
            file_name=file_name,
            user_id=user_id
        )
        result = await analyze(req)
        return {"status": "ok", "result": result}

    except Exception as e:
        print(f"Erreur webhook : {e}")
        return {"status": "error", "detail": str(e)}
