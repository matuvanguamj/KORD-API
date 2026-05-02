import os
import json
import httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from dotenv import load_dotenv
from supabase import create_client, Client
import anthropic

from kord_engine import run_audit
from kord_ia import generate_recommendations

load_dotenv()

SUPABASE_URL         = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")
ANTHROPIC_API_KEY    = os.getenv("ANTHROPIC_API_KEY")

sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

app = FastAPI(title="KORD API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── MODÈLES ──
class AnalyzeRequest(BaseModel):
    submission_id: str
    file_path:     str
    file_name:     str
    user_id:       str

class ChatMessage(BaseModel):
    role:    str
    content: str

class ChatRequest(BaseModel):
    message:  str
    history:  List[ChatMessage] = []
    context:  Optional[str] = "commercial"  # "commercial" ou "client"

# ── ROUTES DE BASE ──
@app.get("/")
def root():
    return {"status": "KORD API en ligne", "version": "1.0.0"}

@app.get("/health")
def health():
    return {"status": "ok"}

# ── CHATBOT ──
@app.post("/chat")
async def chat(req: ChatRequest):
    """Route chatbot — appelée par le dashboard et la page index."""

    if req.context == "commercial":
        system = """Tu es l'assistant virtuel KORD, une plateforme d'audit en performance opérationnelle et supply chain pour PME et ETI.

Tu guides les visiteurs sur ce qu'est KORD, comment ça marche et ce qu'on peut faire pour leur entreprise.
Réponds en français, de façon concise et professionnelle. Maximum 3 phrases par réponse.
Ne donne pas de prix précis. Invite à créer un espace ou contacter l'équipe pour aller plus loin.
Si on te demande si tu es une IA, réponds que tu es l'assistant virtuel KORD."""

    else:
        system = """Tu es l'assistant virtuel KORD, disponible dans l'espace client.

Tu guides les clients sur l'utilisation de la plateforme :
- Comment déposer leurs fichiers (CSV ou Excel depuis leur ERP)
- Quels fichiers déposer et pourquoi (stock, commandes, expéditions, achats, marges, retours)
- Comment suivre l'état de leur dossier
- Comment lire leur rapport KORD

Tu peux aussi répondre à des questions générales sur la supply chain, la logistique et la performance opérationnelle.
Réponds en français, de façon concise et utile. Maximum 4 phrases par réponse.
Si on te demande si tu es une IA, réponds que tu es l'assistant virtuel KORD."""

    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        messages = [{"role": m.role, "content": m.content} for m in req.history[-8:]]
        messages.append({"role": "user", "content": req.message})

        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=400,
            system=system,
            messages=messages
        )

        return {"response": response.content[0].text}

    except Exception as e:
        print(f"Erreur chat : {e}")
        raise HTTPException(status_code=500, detail=str(e))

# ── ANALYSE ──
@app.post("/analyze")
async def analyze(req: AnalyzeRequest):
    print(f"Analyse démarrée : {req.file_name} (user: {req.user_id})")

    # Statut en analyse
    sb.table("submissions").update({"status": "en_analyse"}).eq("id", req.submission_id).execute()

    # Télécharger le fichier
    file_bytes = None
    try:
        # Méthode 1 — URL signée
        result = sb.storage.from_("client-files").create_signed_url(req.file_path, 300)
        print(f"Signed URL result : {result}")

        signed_url = None
        if isinstance(result, dict):
            signed_url = (result.get("signedURL") or result.get("signedUrl") or
                         (result.get("data") or {}).get("signedURL") or
                         (result.get("data") or {}).get("signedUrl"))

        if not signed_url:
            raise Exception(f"URL signée vide : {result}")

        async with httpx.AsyncClient(timeout=60) as client:
            r = await client.get(signed_url)
            r.raise_for_status()
            file_bytes = r.content
            print(f"Fichier téléchargé : {len(file_bytes)} bytes")

    except Exception as e:
        print(f"Erreur téléchargement : {e}")
        # Méthode 2 — download direct
        try:
            file_bytes = sb.storage.from_("client-files").download(req.file_path)
            print(f"Download direct : {len(file_bytes)} bytes")
        except Exception as e2:
            print(f"Erreur download direct : {e2}")
            sb.table("submissions").update({"status": "erreur"}).eq("id", req.submission_id).execute()
            raise HTTPException(status_code=500, detail=f"Impossible de télécharger : {e2}")

    # Moteur d'analyse
    try:
        audit_results = run_audit(file_bytes, req.file_name)
    except Exception as e:
        print(f"Erreur moteur : {e}")
        sb.table("submissions").update({"status": "erreur"}).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=500, detail=f"Erreur moteur : {e}")

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

    # Sauvegarder
    try:
        sb.table("reports").insert({
            "submission_id": req.submission_id,
            "user_id":       req.user_id,
            "score_kord":    score,
            "summary":       summary,
            "report_url":    None,
        }).execute()
    except Exception as e:
        print(f"Erreur sauvegarde rapport : {e}")

    print(f"Analyse terminée : score {score}/100 — {len(audit_results.get('alertes', []))} alertes")

    return {
        "status":  "ok",
        "score":   score,
        "summary": summary,
        "alertes": len(audit_results.get("alertes", [])),
    }

# ── WEBHOOK SUPABASE ──
@app.post("/webhook/supabase")
async def webhook(payload: dict):
    """Déclenché automatiquement par Supabase quand un fichier est déposé."""
    try:
        record = payload.get("record", {})
        submission_id = record.get("id")
        file_path     = record.get("file_path")
        file_name     = record.get("file_name")
        user_id       = record.get("user_id")

        if not all([submission_id, file_path, file_name, user_id]):
            return {"status": "ignored", "reason": "Données manquantes"}

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
