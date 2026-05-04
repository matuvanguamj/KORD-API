import os
import base64
import asyncio
from datetime import datetime
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
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── MODÈLES ──
class AnalyzeGlobalRequest(BaseModel):
    user_id:        str
    trimestre:      str
    submission_ids: List[str]

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
    system = (
        "Tu es l'assistant virtuel KORD, plateforme d'audit supply chain. "
        "Réponds en français, concis, max 3 phrases."
        if req.context == "commercial"
        else
        "Tu es l'assistant virtuel KORD dans l'espace client. "
        "Guide le client sur la plateforme et réponds aux questions supply chain. Max 4 phrases."
    )
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        messages = [{"role": m.role, "content": m.content} for m in req.history[-8:]]
        messages.append({"role": "user", "content": req.message})
        response = client.messages.create(model="claude-sonnet-4-6", max_tokens=400, system=system, messages=messages)
        return {"response": response.content[0].text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ── WEBHOOK SUPABASE ──
@app.post("/webhook/supabase")
async def webhook(payload: dict):
    """
    Déclenché par Supabase à chaque INSERT sur submissions.
    Attend que tous les fichiers arrivent puis lance l'analyse globale.
    """
    try:
        record  = payload.get("record", {})
        user_id = record.get("user_id")
        if not user_id:
            return {"status": "ignored"}

        print(f"Webhook reçu — user: {user_id[:8]}")

        # Attendre 8 secondes que le client finisse de déposer tous ses fichiers
        await asyncio.sleep(8)

        # Vérifier qu'on n'a pas déjà un rapport récent pour ce client (< 2 min)
        recent = sb.table("reports").select("id,created_at").eq("user_id", user_id).order("created_at", desc=True).limit(1).execute()
        if recent.data:
            last_report_time = datetime.fromisoformat(recent.data[0]["created_at"].replace("Z",""))
            diff = (datetime.utcnow() - last_report_time).total_seconds()
            if diff < 120:
                print(f"Rapport récent trouvé ({int(diff)}s) — analyse ignorée pour éviter doublon")
                return {"status": "skipped", "reason": "recent_report"}

        # Récupérer tous les fichiers en_analyse de ce client
        subs = sb.table("submissions").select("id").eq("user_id", user_id).eq("status", "en_analyse").execute()
        submission_ids = [r["id"] for r in (subs.data or [])]

        if not submission_ids:
            print(f"Aucun fichier en_analyse pour {user_id[:8]}")
            return {"status": "ignored", "reason": "no_files"}

        print(f"Lancement analyse globale : {len(submission_ids)} fichiers")

        now = datetime.now()
        trimestre = f"T{(now.month-1)//3+1} {now.year}"

        req = AnalyzeGlobalRequest(user_id=user_id, trimestre=trimestre, submission_ids=submission_ids)
        result = await analyze_global(req)
        return {"status": "ok", "result": result}

    except Exception as e:
        print(f"Erreur webhook : {e}")
        return {"status": "error", "detail": str(e)}

# ── ANALYSE GLOBALE ──
@app.post("/analyze-global")
async def analyze_global(req: AnalyzeGlobalRequest):
    print(f"Analyse globale : {len(req.submission_ids)} fichiers (user: {req.user_id[:8]})")

    # Récupérer les submissions
    submissions = []
    for sid in req.submission_ids:
        r = sb.table("submissions").select("*").eq("id", sid).execute()
        if r.data:
            submissions.append(r.data[0])
            sb.table("submissions").update({"status": "en_analyse"}).eq("id", sid).execute()

    if not submissions:
        raise HTTPException(status_code=400, detail="Aucun fichier trouvé")

    # Analyser chaque fichier
    all_results = []
    for sub in submissions:
        try:
            file_bytes = sb.storage.from_("client-files").download(sub["file_path"])
            if not file_bytes:
                raise Exception("Vide")
            print(f"Téléchargé : {sub['file_name']} ({len(file_bytes)} bytes)")
            result = run_audit(file_bytes, sub["file_name"])
            if "error" not in result:
                all_results.append(result)
                print(f"Analysé : {sub['file_name']} — score {result.get('score_total',0)}")
        except Exception as e:
            print(f"Erreur {sub['file_name']} : {e}")

    if not all_results:
        raise HTTPException(status_code=400, detail="Aucun fichier analysable")

    # Consolider
    consolidated = consolidate(all_results)
    score = consolidated["score_total"]
    print(f"Score consolidé : {score}/100 sur {len(all_results)} fichiers")

    # Recommandations IA
    try:
        reco = generate_recommendations_global(consolidated, all_results)
    except Exception as e:
        print(f"Erreur IA : {e}")
        reco = {"resume_executif": f"Analyse de {len(all_results)} fichier(s). Score : {score}/100.", "priorites": [], "points_vigilance": [], "prochaine_etape": "Session de restitution à planifier."}

    summary = reco.get("resume_executif", "")

    # Profil client
    profile = sb.table("profiles").select("*").eq("id", req.user_id).execute()
    client_name  = profile.data[0].get("contact_name","Client") if profile.data else "Client"
    company_name = profile.data[0].get("company_name","") if profile.data else ""

    # Graphiques
    radar_png = bar_png = gauge_png = evol_png = None
    try:
        from kord_charts import generate_evolution_chart
        radar_png = generate_radar_chart(consolidated.get("analyses",{}))
        bar_png   = generate_bar_chart(consolidated.get("analyses",{}))
        gauge_png = generate_gauge_chart(score)
        evol_png  = generate_evolution_chart(None)
        print("Graphiques générés")
    except Exception as e:
        print(f"Erreur graphiques : {e}")

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_url = docx_url = None

    # PDF
    try:
        pdf_bytes = generate_prereport_pdf(consolidated=consolidated, recommendations=reco, all_results=all_results,
            client_name=client_name, company_name=company_name, trimestre=req.trimestre,
            gauge_png=gauge_png, bar_png=bar_png, radar_png=radar_png, evol_png=evol_png)
        pdf_path = f"rapports/{req.user_id}/{ts}_prereport.pdf"
        sb.storage.from_("client-files").upload(pdf_path, pdf_bytes, {"content-type":"application/pdf"})
        signed = sb.storage.from_("client-files").create_signed_url(pdf_path, 60*60*24*30)
        pdf_url = signed.get("signedURL") or signed.get("signedUrl") or (signed.get("data") or {}).get("signedURL")
        print(f"PDF généré")
    except Exception as e:
        print(f"Erreur PDF : {e}")

    # DOCX
    try:
        docx_bytes = generate_prereport_docx(consolidated=consolidated, recommendations=reco, all_results=all_results,
            client_name=client_name, company_name=company_name, trimestre=req.trimestre,
            radar_png=radar_png, bar_png=bar_png, gauge_png=gauge_png, evol_png=evol_png)
        docx_path = f"rapports/{req.user_id}/{ts}_prereport.docx"
        sb.storage.from_("client-files").upload(docx_path, docx_bytes,
            {"content-type":"application/vnd.openxmlformats-officedocument.wordprocessingml.document"})
        signed2 = sb.storage.from_("client-files").create_signed_url(docx_path, 60*60*24*30)
        docx_url = signed2.get("signedURL") or signed2.get("signedUrl") or (signed2.get("data") or {}).get("signedURL")
        print(f"DOCX généré")
    except Exception as e:
        print(f"Erreur DOCX : {e}")

    # Sauvegarder rapport (sans report_url — toi tu l'envoies manuellement au client)
    try:
        sb.table("reports").insert({
            "submission_id": submissions[0]["id"],
            "user_id":       req.user_id,
            "score_kord":    score,
            "summary":       summary,
            "report_url":    None,      # NULL — client ne voit rien avant validation admin
            "prereport_docx": docx_url,
            "prereport_pdf":  pdf_url,   # PDF admin seulement
        }).execute()
        print(f"Rapport sauvegardé — score {score}/100")
    except Exception as e:
        print(f"Erreur sauvegarde : {e}")

    return {"status":"ok","score":score,"pdf":pdf_url is not None,"docx":docx_url is not None}


def consolidate(results):
    piliers = {
        "stock_cash":          {"score":0,"max":30,"count":0},
        "transport_service":   {"score":0,"max":20,"count":0},
        "achats_fournisseurs": {"score":0,"max":20,"count":0},
        "marges_retours":      {"score":0,"max":15,"count":0},
        "donnees_pilotage":    {"score":0,"max":15,"count":0},
    }
    all_alertes, all_opps = [], []
    for r in results:
        for p, d in r.get("analyses",{}).items():
            if p in piliers:
                piliers[p]["score"] += d.get("score",0)
                piliers[p]["count"] += 1
        for a in r.get("alertes",[]):
            if a["message"] not in [x["message"] for x in all_alertes]:
                all_alertes.append(a)
        for o in r.get("opportunites",[]):
            if o["message"] not in [x["message"] for x in all_opps]:
                all_opps.append(o)

    analyses = {}
    score_total = 0
    for p, d in piliers.items():
        avg = round(d["score"]/d["count"]) if d["count"] > 0 else round(d["max"]*0.6)
        analyses[p] = {"score":avg,"max":d["max"],"alertes":[],"opportunites":[]}
        score_total += avg

    score_total = min(100, max(0, score_total))
    interp = (
        "Performance solide." if score_total >= 80 else
        "Performance correcte, axes d'amélioration identifiés." if score_total >= 65 else
        "Niveau moyen, plusieurs leviers à actionner." if score_total >= 50 else
        "Tensions importantes, actions prioritaires recommandées." if score_total >= 35 else
        "Niveau critique, failles structurelles détectées."
    )
    return {
        "score_total":    score_total,
        "nb_fichiers":    len(results),
        "analyses":       analyses,
        "alertes":        all_alertes,
        "opportunites":   all_opps,
        "interpretation": f"Analyse de {len(results)} fichier(s). {interp}",
    }
