import os
import json
import base64
import anthropic
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from dotenv import load_dotenv
from supabase import create_client, Client

from kord_engine import run_audit
from kord_ia import generate_recommendations

load_dotenv()

SUPABASE_URL         = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")
ANTHROPIC_API_KEY    = os.getenv("ANTHROPIC_API_KEY")

sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

app = FastAPI(title="KORD API", version="1.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── MODÈLES ──
class AnalyzeRequest(BaseModel):
    submission_id: str
    file_name:     str
    user_id:       str
    file_b64:      str  # fichier encodé base64 — plus besoin de Storage

class ChatMessage(BaseModel):
    role:    str
    content: str

class ChatRequest(BaseModel):
    message:  str
    history:  List[ChatMessage] = []
    context:  Optional[str] = "commercial"

# ── ROUTES ──
@app.get("/")
def root():
    return {"status": "KORD API en ligne", "version": "1.0.0"}

@app.get("/health")
def health():
    return {"status": "ok"}

# ── CHATBOT ──
@app.post("/chat")
async def chat(req: ChatRequest):
    if req.context == "commercial":
        system = """Tu es l'assistant virtuel KORD, une plateforme d'audit en performance opérationnelle et supply chain pour PME et ETI.
Tu guides les visiteurs sur ce qu'est KORD, comment ça marche et ce qu'on peut faire pour leur entreprise.
Réponds en français, de façon concise et professionnelle. Maximum 3 phrases.
Ne donne pas de prix précis. Invite à créer un espace ou contacter l'équipe."""
    else:
        system = """Tu es l'assistant virtuel KORD, disponible dans l'espace client.
Tu guides les clients sur l'utilisation de la plateforme et réponds aux questions sur la supply chain, la logistique et la performance opérationnelle.
Réponds en français, de façon concise et utile. Maximum 4 phrases."""
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
        raise HTTPException(status_code=500, detail=str(e))

# ── ANALYSE ──
@app.post("/analyze")
async def analyze(req: AnalyzeRequest):
    print(f"Analyse démarrée : {req.file_name} (user: {req.user_id})")

    # Mettre en analyse
    sb.table("submissions").update({"status": "en_analyse"}).eq("id", req.submission_id).execute()

    # Décoder le fichier base64
    try:
        file_bytes = base64.b64decode(req.file_b64)
        print(f"Fichier reçu : {len(file_bytes)} bytes")
    except Exception as e:
        sb.table("submissions").update({"status": "erreur"}).eq("id", req.submission_id).execute()
        raise HTTPException(status_code=400, detail=f"Erreur décodage fichier : {e}")

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

    # Sauvegarder rapport
    try:
        sb.table("reports").insert({
            "submission_id": req.submission_id,
            "user_id":       req.user_id,
            "score_kord":    score,
            "summary":       summary,
            "report_url":    None,
        }).execute()
    except Exception as e:
        print(f"Erreur sauvegarde : {e}")

    print(f"Analyse terminée : score {score}/100 — {len(audit_results.get('alertes', []))} alertes")
    return {"status": "ok", "score": score, "summary": summary}
