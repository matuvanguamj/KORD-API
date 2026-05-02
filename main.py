import os
import json
import base64
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from dotenv import load_dotenv
from supabase import create_client, Client

from kord_engine import run_audit
from kord_ia import generate_recommendations_global
from kord_pdf import generate_prereport_pdf

load_dotenv()

SUPABASE_URL         = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")
ANTHROPIC_API_KEY    = os.getenv("ANTHROPIC_API_KEY")

sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

app = FastAPI(title="KORD API", version="1.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── MODÈLES ──
class FileData(BaseModel):
    submission_id: str
    file_name:     str
    file_b64:      str

class AnalyzeGlobalRequest(BaseModel):
    user_id:   str
    trimestre: str
    files:     List[FileData]

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
    import anthropic
    if req.context == "commercial":
        system = """Tu es l'assistant virtuel KORD, une plateforme d'audit en performance opérationnelle et supply chain.
Tu guides les visiteurs sur ce qu'est KORD et ce qu'on peut faire pour leur entreprise.
Réponds en français, concis et professionnel. Maximum 3 phrases. Ne donne pas de prix précis."""
    else:
        system = """Tu es l'assistant virtuel KORD dans l'espace client.
Tu guides les clients sur la plateforme et réponds aux questions supply chain, logistique et performance.
Réponds en français, concis et utile. Maximum 4 phrases."""
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        messages = [{"role": m.role, "content": m.content} for m in req.history[-8:]]
        messages.append({"role": "user", "content": req.message})
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=400,
            system=system,
            messages=messages
        )
        return {"response": response.content[0].text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ── ANALYSE GLOBALE ──
@app.post("/analyze-global")
async def analyze_global(req: AnalyzeGlobalRequest):
    """
    Analyse TOUS les fichiers d'un client ensemble.
    Retourne UN score global et UN rapport consolidé.
    """
    print(f"Analyse globale démarrée : {len(req.files)} fichiers (user: {req.user_id})")

    # Mettre tous les fichiers en analyse
    for f in req.files:
        sb.table("submissions").update({"status": "en_analyse"}).eq("id", f.submission_id).execute()

    # Analyser chaque fichier
    all_results = []
    for file_data in req.files:
        try:
            file_bytes = base64.b64decode(file_data.file_b64)
            print(f"Analyse fichier : {file_data.file_name} ({len(file_bytes)} bytes)")
            result = run_audit(file_bytes, file_data.file_name)
            if "error" not in result:
                all_results.append(result)
        except Exception as e:
            print(f"Erreur fichier {file_data.file_name} : {e}")

    if not all_results:
        raise HTTPException(status_code=400, detail="Aucun fichier analysable")

    # Consolider tous les résultats en un score global
    consolidated = consolidate_results(all_results)

    # Générer recommandations IA sur l'ensemble
    try:
        recommendations = generate_recommendations_global(consolidated, all_results)
    except Exception as e:
        print(f"Erreur IA : {e}")
        recommendations = {
            "resume_executif": f"Analyse de {len(all_results)} fichiers terminée. Score global : {consolidated['score_total']}/100.",
            "priorites": []
        }

    score   = consolidated["score_total"]
    summary = recommendations.get("resume_executif", "")

    # Générer le pré-rapport PDF
    pdf_url = None
    try:
        from datetime import datetime
        now = datetime.now()
        pdf_bytes = generate_prereport_pdf(
            consolidated=consolidated,
            recommendations=recommendations,
            all_results=all_results,
            client_name=req.user_id[:8],
            company_name="",
            trimestre=req.trimestre,
        )
        pdf_path = f"rapports/{req.user_id}/{now.strftime('%Y%m%d_%H%M%S')}_prereport.pdf"
        sb.storage.from_("client-files").upload(pdf_path, pdf_bytes, {"content-type": "application/pdf"})
        signed = sb.storage.from_("client-files").create_signed_url(pdf_path, 60*60*24*30)
        pdf_url = signed.get("signedURL") or signed.get("signedUrl") or (signed.get("data") or {}).get("signedURL")
        print(f"Pré-rapport PDF généré : {pdf_path}")
    except Exception as e:
        print(f"Erreur génération PDF : {e}")

    # Sauvegarder UN seul rapport global avec lien PDF
    try:
        sb.table("reports").insert({
            "user_id":       req.user_id,
            "submission_id": req.files[0].submission_id,
            "score_kord":    score,
            "summary":       summary,
            "report_url":    pdf_url,
        }).execute()
        print(f"Rapport global sauvegardé : score {score}/100")
    except Exception as e:
        print(f"Erreur sauvegarde rapport : {e}")

    # Marquer tous les fichiers en analyse (toi tu passes à terminé manuellement)
    for f in req.files:
        sb.table("submissions").update({"status": "en_analyse"}).eq("id", f.submission_id).execute()

    return {
        "status":    "ok",
        "score":     score,
        "summary":   summary,
        "nb_files":  len(all_results),
        "alertes":   len(consolidated.get("alertes", [])),
    }

def consolidate_results(results: list) -> dict:
    """Consolide les résultats de plusieurs fichiers en un score global."""

    # Agréger les scores par pilier
    piliers = {
        "stock_cash":          {"score": 0, "max": 30, "count": 0},
        "transport_service":   {"score": 0, "max": 20, "count": 0},
        "achats_fournisseurs": {"score": 0, "max": 20, "count": 0},
        "marges_retours":      {"score": 0, "max": 15, "count": 0},
        "donnees_pilotage":    {"score": 0, "max": 15, "count": 0},
    }

    all_alertes      = []
    all_opportunites = []

    for r in results:
        for pilier, data in r.get("analyses", {}).items():
            if pilier in piliers:
                piliers[pilier]["score"] += data.get("score", 0)
                piliers[pilier]["count"] += 1

        for a in r.get("alertes", []):
            if a not in all_alertes:
                all_alertes.append(a)
        for o in r.get("opportunites", []):
            if o not in all_opportunites:
                all_opportunites.append(o)

    # Moyenne des scores par pilier
    consolidated_analyses = {}
    score_total = 0
    for pilier, data in piliers.items():
        if data["count"] > 0:
            avg = round(data["score"] / data["count"])
        else:
            avg = round(data["max"] * 0.6)
        consolidated_analyses[pilier] = {
            "score": avg,
            "max":   data["max"],
            "alertes": [],
            "opportunites": []
        }
        score_total += avg

    score_total = min(100, max(0, score_total))

    return {
        "score_total":    score_total,
        "nb_fichiers":    len(results),
        "analyses":       consolidated_analyses,
        "alertes":        all_alertes,
        "opportunites":   all_opportunites,
        "interpretation": interpret_global(score_total, len(results)),
    }

def interpret_global(score: int, nb_fichiers: int) -> str:
    base = f"Analyse consolidée de {nb_fichiers} fichier{'s' if nb_fichiers > 1 else ''}. "
    if score >= 80:
        return base + "Performance opérationnelle solide. Peu de leviers majeurs détectés."
    elif score >= 65:
        return base + "Performance correcte avec des axes d'amélioration identifiés."
    elif score >= 50:
        return base + "Niveau moyen. Plusieurs leviers significatifs à actionner."
    elif score >= 35:
        return base + "Tensions opérationnelles importantes. Actions prioritaires recommandées."
    else:
        return base + "Niveau critique. Failles structurelles multiples détectées."

# ── GARDE L'ANCIENNE ROUTE POUR COMPATIBILITÉ ──
class AnalyzeRequest(BaseModel):
    submission_id: str
    file_name:     str
    user_id:       str
    file_b64:      str

@app.post("/analyze")
async def analyze(req: AnalyzeRequest):
    """Route legacy — un fichier à la fois."""
    print(f"Analyse démarrée : {req.file_name}")
    sb.table("submissions").update({"status": "en_analyse"}).eq("id", req.submission_id).execute()
    try:
        file_bytes = base64.b64decode(req.file_b64)
        audit_results = run_audit(file_bytes, req.file_name)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    if "error" in audit_results:
        raise HTTPException(status_code=400, detail=audit_results["error"])

    score = audit_results.get("score_total", 0)

    try:
        sb.table("reports").insert({
            "submission_id": req.submission_id,
            "user_id":       req.user_id,
            "score_kord":    score,
            "summary":       f"Analyse de {req.file_name} terminée. Score : {score}/100.",
            "report_url":    None,
        }).execute()
    except Exception as e:
        print(f"Erreur sauvegarde : {e}")

    print(f"Analyse terminée : score {score}/100")
    return {"status": "ok", "score": score}
