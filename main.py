import os
import base64
import anthropic
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from dotenv import load_dotenv
from supabase import create_client, Client

from kord_engine import run_audit
from kord_ia import generate_recommendations_global
from kord_pdf import generate_prereport_pdf
from kord_docx import generate_prereport_docx
from kord_charts import generate_radar_chart, generate_bar_chart, generate_gauge_chart

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
class AnalyzeGlobalRequest(BaseModel):
    user_id:        str
    trimestre:      str
    submission_ids: List[str]  # Juste les IDs — l'API télécharge elle-même

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
        system = """Tu es l'assistant virtuel KORD, une plateforme d'audit en performance opérationnelle et supply chain.
Tu guides les visiteurs. Réponds en français, concis. Maximum 3 phrases. Pas de prix précis."""
    else:
        system = """Tu es l'assistant virtuel KORD dans l'espace client.
Tu guides les clients sur la plateforme et réponds aux questions supply chain et logistique.
Réponds en français, concis. Maximum 4 phrases."""
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
    print(f"Analyse globale : {len(req.submission_ids)} fichiers (user: {req.user_id})")

    # 1. Récupérer les infos des submissions
    submissions = []
    for sid in req.submission_ids:
        result = sb.table("submissions").select("*").eq("id", sid).execute()
        if result.data:
            submissions.append(result.data[0])
            sb.table("submissions").update({"status": "en_analyse"}).eq("id", sid).execute()

    if not submissions:
        raise HTTPException(status_code=400, detail="Aucun fichier trouvé")

    # 2. Télécharger et analyser chaque fichier
    all_results = []
    for sub in submissions:
        file_path = sub.get("file_path")
        file_name = sub.get("file_name")
        try:
            # Téléchargement direct avec service_key — bypass RLS total
            file_bytes = sb.storage.from_("client-files").download(file_path)
            if not file_bytes:
                raise Exception("Fichier vide")
            print(f"Téléchargé : {file_name} ({len(file_bytes)} bytes)")
            result = run_audit(file_bytes, file_name)
            if "error" not in result:
                all_results.append(result)
                print(f"Analysé : {file_name} — score partiel {result.get('score_total',0)}")
        except Exception as e:
            print(f"Erreur {file_name} : {e}")

    if not all_results:
        raise HTTPException(status_code=400, detail="Aucun fichier analysable")

    # 3. Consolider
    consolidated = consolidate_results(all_results)
    score = consolidated["score_total"]
    print(f"Score consolidé : {score}/100 sur {len(all_results)} fichiers")

    # 4. Recommandations IA
    try:
        recommendations = generate_recommendations_global(consolidated, all_results)
        print("Recommandations IA générées")
    except Exception as e:
        print(f"Erreur IA : {e}")
        recommendations = {
            "resume_executif": f"Analyse de {len(all_results)} fichiers terminée. Score global KORD : {score}/100.",
            "priorites": [],
            "points_vigilance": [],
            "prochaine_etape": "Planifier la session de restitution."
        }

    summary = recommendations.get("resume_executif", "")

    # 5. Générer les graphiques + PDF + DOCX
    pdf_url  = None
    docx_url = None
    try:
        from datetime import datetime

        # Récupérer le profil client
        profile = sb.table("profiles").select("*").eq("id", req.user_id).execute()
        client_name  = profile.data[0].get("contact_name", "") if profile.data else ""
        company_name = profile.data[0].get("company_name", "") if profile.data else ""

        # Graphiques
        radar_png = None
        bar_png   = None
        gauge_png = None
        try:
            radar_png = generate_radar_chart(consolidated.get("analyses", {}))
            bar_png   = generate_bar_chart(consolidated.get("analyses", {}))
            gauge_png = generate_gauge_chart(score)
            print("Graphiques générés")
        except Exception as e:
            print(f"Erreur graphiques : {e}")

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')

        # PDF
        try:
            pdf_bytes = generate_prereport_pdf(
                consolidated=consolidated,
                recommendations=recommendations,
                all_results=all_results,
                client_name=client_name or "Client",
                company_name=company_name or "",
                trimestre=req.trimestre,
            )
            pdf_path = f"rapports/{req.user_id}/{ts}_prereport.pdf"
            sb.storage.from_("client-files").upload(pdf_path, pdf_bytes, {"content-type": "application/pdf"})
            signed = sb.storage.from_("client-files").create_signed_url(pdf_path, 60*60*24*30)
            pdf_url = (signed.get("signedURL") or signed.get("signedUrl") or
                      (signed.get("data") or {}).get("signedURL") or
                      (signed.get("data") or {}).get("signedUrl"))
            print(f"PDF généré : {pdf_path}")
        except Exception as e:
            print(f"Erreur PDF : {e}")

        # DOCX éditable
        try:
            docx_bytes = generate_prereport_docx(
                consolidated=consolidated,
                recommendations=recommendations,
                all_results=all_results,
                client_name=client_name or "Client",
                company_name=company_name or "",
                trimestre=req.trimestre,
                radar_png=radar_png,
                bar_png=bar_png,
                gauge_png=gauge_png,
            )
            docx_path = f"rapports/{req.user_id}/{ts}_prereport.docx"
            sb.storage.from_("client-files").upload(docx_path, docx_bytes, {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"})
            signed2 = sb.storage.from_("client-files").create_signed_url(docx_path, 60*60*24*30)
            docx_url = (signed2.get("signedURL") or signed2.get("signedUrl") or
                       (signed2.get("data") or {}).get("signedURL") or
                       (signed2.get("data") or {}).get("signedUrl"))
            print(f"DOCX généré : {docx_path}")
        except Exception as e:
            print(f"Erreur DOCX : {e}")

    except Exception as e:
        print(f"Erreur génération documents : {e}")

    # 6. Sauvegarder rapport
    try:
        sb.table("reports").insert({
            "submission_id": submissions[0]["id"],
            "user_id":       req.user_id,
            "score_kord":    score,
            "summary":       summary,
            "report_url":    pdf_url,
            "prereport_docx": docx_url,
        }).execute()
        print(f"Rapport sauvegardé — score {score}/100 — PDF : {pdf_url is not None}")
    except Exception as e:
        print(f"Erreur sauvegarde : {e}")

    return {
        "status":   "ok",
        "score":    score,
        "summary":  summary,
        "nb_files": len(all_results),
        "pdf":      pdf_url is not None,
    }


def consolidate_results(results: list) -> dict:
    piliers = {
        "stock_cash":          {"score": 0, "max": 30, "count": 0, "alertes": [], "opportunites": []},
        "transport_service":   {"score": 0, "max": 20, "count": 0, "alertes": [], "opportunites": []},
        "achats_fournisseurs": {"score": 0, "max": 20, "count": 0, "alertes": [], "opportunites": []},
        "marges_retours":      {"score": 0, "max": 15, "count": 0, "alertes": [], "opportunites": []},
        "donnees_pilotage":    {"score": 0, "max": 15, "count": 0, "alertes": [], "opportunites": []},
    }
    all_alertes, all_opps = [], []

    for r in results:
        for pilier, data in r.get("analyses", {}).items():
            if pilier in piliers:
                piliers[pilier]["score"] += data.get("score", 0)
                piliers[pilier]["count"] += 1
        for a in r.get("alertes", []):
            if a["message"] not in [x["message"] for x in all_alertes]:
                all_alertes.append(a)
        for o in r.get("opportunites", []):
            if o["message"] not in [x["message"] for x in all_opps]:
                all_opps.append(o)

    consolidated_analyses = {}
    score_total = 0
    for pilier, data in piliers.items():
        avg = round(data["score"] / data["count"]) if data["count"] > 0 else round(data["max"] * 0.6)
        consolidated_analyses[pilier] = {"score": avg, "max": data["max"], "alertes": [], "opportunites": []}
        score_total += avg

    score_total = min(100, max(0, score_total))

    interp = (
        "Performance opérationnelle solide." if score_total >= 80 else
        "Performance correcte avec des axes d'amélioration identifiés." if score_total >= 65 else
        "Niveau moyen. Plusieurs leviers significatifs à actionner." if score_total >= 50 else
        "Tensions opérationnelles importantes. Actions prioritaires recommandées." if score_total >= 35 else
        "Niveau critique. Failles structurelles multiples détectées."
    )

    return {
        "score_total":    score_total,
        "nb_fichiers":    len(results),
        "analyses":       consolidated_analyses,
        "alertes":        all_alertes,
        "opportunites":   all_opps,
        "interpretation": f"Analyse consolidée de {len(results)} fichier(s). {interp}",
    }
