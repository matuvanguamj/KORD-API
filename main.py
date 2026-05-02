import os
import json
import httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv
from supabase import create_client, Client

from kord_engine import run_audit
from kord_ia import generate_recommendations

load_dotenv()

SUPABASE_URL         = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")

sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

app = FastAPI(title="KORD API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class AnalyzeRequest(BaseModel):
    submission_id: str
    file_path:     str
    file_name:     str
    user_id:       str

@app.get("/")
def root():
    return {"status": "KORD API en ligne", "version": "1.0.0"}

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/analyze")
async def analyze(req: AnalyzeRequest):
    print(f"Analyse démarrée : {req.file_name} (user: {req.user_id})")

    # Statut en analyse
    sb.table("submissions").update({"status": "en_analyse"}).eq("id", req.submission_id).execute()

    # Télécharger via URL signée (plus fiable)
    file_bytes = None
    try:
        signed_data = sb.storage.from_("client-files").create_signed_url(req.file_path, 300)
        signed_url  = signed_data.get("signedURL") or signed_data.get("signedUrl") or signed_data.get("data", {}).get("signedURL")
        if not signed_url:
            raise Exception(f"URL signée non générée : {signed_data}")
        async with httpx.AsyncClient() as client:
            response = await client.get(signed_url, timeout=60)
            response.raise_for_status()
            file_bytes = response.content
        print(f"Fichier téléchargé : {len(file_bytes)} bytes")
    except Exception as e:
        print(f"Erreur téléchargement : {e}")
        sb.table("submissions").update({"status": "erreur"}).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=500, detail=f"Impossible de télécharger le fichier : {str(e)}")

    # Moteur d'analyse
    try:
        audit_results = run_audit(file_bytes, req.file_name)
    except Exception as e:
        print(f"Erreur moteur : {e}")
        sb.table("submissions").update({"status": "erreur"}).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=500, detail=f"Erreur moteur : {str(e)}")

    if "error" in audit_results:
        sb.table("submissions").update({"status": "erreur"}).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=400, detail=audit_results["error"])

    # Recommandations IA
    try:
        recommendations = generate_recommendations(audit_results)
    except Exception as e:
        print(f"Erreur IA : {e}")
        recommendations = {"resume_executif": "Analyse terminée. Restitution en cours.", "priorites": []}

    score   = audit_results.get("score_total", 0)
    summary = recommendations.get("resume_executif", "")

    # Sauvegarder rapport
    sb.table("reports").insert({
        "submission_id": req.submission_id,
        "user_id":       req.user_id,
        "score_kord":    score,
        "summary":       summary,
        "report_url":    None,
    }).execute()

    # Statut en_analyse (toi tu passes à terminé manuellement)
    sb.table("submissions").update({"status": "en_analyse"}).eq("id", req.submission_id).execute()

    print(f"Analyse terminée : score {score}/100 — {len(audit_results.get('alertes', []))} alertes")

    return {
        "status":  "ok",
        "score":   score,
        "summary": summary,
        "alertes": len(audit_results.get("alertes", [])),
    }
