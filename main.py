import os
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
from kord_charts import (generate_radar_chart, generate_bar_chart,
                         generate_gauge_chart, generate_dormance_chart,
                         generate_score_breakdown_chart, generate_cash_impact_chart)
from kord_html import generate_prereport_html

load_dotenv()
SUPABASE_URL         = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")
ANTHROPIC_API_KEY    = os.getenv("ANTHROPIC_API_KEY")
sb: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

app = FastAPI(title="KORD API", version="2.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

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

@app.get("/")
def root(): return {"status": "KORD API en ligne", "version": "2.0.0"}

@app.get("/health")
def health(): return {"status": "ok"}

@app.post("/chat")
async def chat(req: ChatRequest):
    import anthropic
    system = ("Tu es l'assistant virtuel KORD, plateforme d'audit supply chain. Réponds en français, concis, max 3 phrases."
              if req.context == "commercial"
              else "Tu es l'assistant KORD dans l'espace client. Guide sur la plateforme et supply chain. Max 4 phrases.")
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        msgs = [{"role": m.role, "content": m.content} for m in req.history[-8:]]
        msgs.append({"role": "user", "content": req.message})
        r = client.messages.create(model="claude-sonnet-4-6", max_tokens=400, system=system, messages=msgs)
        return {"response": r.content[0].text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/webhook/supabase")
async def webhook(payload: dict):
    try:
        record  = payload.get("record", {})
        user_id = record.get("user_id")
        if not user_id:
            return {"status": "ignored"}
        print(f"Webhook reçu — user: {user_id[:8]}")
        await asyncio.sleep(8)

        # Anti-doublon
        recent = sb.table("reports").select("id,created_at").eq("user_id", user_id).order("created_at", desc=True).limit(1).execute()
        if recent.data:
            diff = (datetime.utcnow() - datetime.fromisoformat(recent.data[0]["created_at"].replace("Z",""))).total_seconds()
            if diff < 120:
                print(f"Rapport récent trouvé ({int(diff)}s) — analyse ignorée pour éviter doublon")
                return {"status": "skipped"}

        subs = sb.table("submissions").select("id").eq("user_id", user_id).eq("status", "en_analyse").execute()
        ids  = [r["id"] for r in (subs.data or [])]
        if not ids:
            return {"status": "ignored", "reason": "no_files"}

        print(f"Lancement analyse globale : {len(ids)} fichiers")
        now = datetime.now()
        req = AnalyzeGlobalRequest(
            user_id=user_id,
            trimestre=f"T{(now.month-1)//3+1} {now.year}",
            submission_ids=ids
        )
        result = await analyze_global(req)
        return {"status": "ok", "result": result}
    except Exception as e:
        print(f"Erreur webhook : {e}")
        return {"status": "error", "detail": str(e)}

@app.post("/analyze-global")
async def analyze_global(req: AnalyzeGlobalRequest):
    print(f"Analyse globale : {len(req.submission_ids)} fichiers (user: {req.user_id[:8]})")

    # 1. Récupérer et analyser les fichiers
    submissions = []
    for sid in req.submission_ids:
        r = sb.table("submissions").select("*").eq("id", sid).execute()
        if r.data:
            submissions.append(r.data[0])
            sb.table("submissions").update({"status": "en_analyse"}).eq("id", sid).execute()

    if not submissions:
        raise HTTPException(status_code=400, detail="Aucun fichier trouvé")

    all_results = []
    for sub in submissions:
        try:
            file_bytes = sb.storage.from_("client-files").download(sub["file_path"])
            if file_bytes:
                result = run_audit(file_bytes, sub["file_name"])
                if "error" not in result:
                    all_results.append(result)
                    print(f"Analysé : {sub['file_name']} — score {result.get('score_total',0)}")
        except Exception as e:
            print(f"Erreur {sub['file_name']} : {e}")

    if not all_results:
        raise HTTPException(status_code=400, detail="Aucun fichier analysable")

    # 2. Consolider
    consolidated = _consolidate(all_results)
    score = consolidated["score_total"]
    print(f"Score consolidé : {score}/100 sur {len(all_results)} fichiers")

    # 3. IA
    try:
        reco = generate_recommendations_global(consolidated, all_results)
    except Exception as e:
        print(f"Erreur IA : {e}")
        reco = {"resume_executif": {"paragraphe_1": f"Score KORD : {score}/100.", "paragraphe_2": "", "paragraphe_3": ""},
                "message_dirigeant": "Analyse terminée. Session de restitution à planifier.",
                "priorites": [], "points_vigilance": [], "prochaine_etape": "Planifier la restitution.",
                "benchmark": "", "anomalies": [], "croisements_cles": [], "opportunites_cachees": [],
                "questions_restitution": [], "analyse_piliers": {}}

    # 4. Profil client
    profile = sb.table("profiles").select("*").eq("id", req.user_id).execute()
    client_name  = profile.data[0].get("contact_name","Client") if profile.data else "Client"
    company_name = profile.data[0].get("company_name","") if profile.data else ""

    # 5. Graphiques
    radar_png = bar_png = gauge_png = dormance_png = breakdown_png = cash_png = None
    try:
        radar_png     = generate_radar_chart(consolidated.get("analyses",{}))
        bar_png       = generate_bar_chart(consolidated.get("analyses",{}))
        gauge_png     = generate_gauge_chart(score)
        dormance_png  = generate_dormance_chart(consolidated)
        breakdown_png = generate_score_breakdown_chart(consolidated)
        cash_png      = generate_cash_impact_chart(consolidated)
        print("Graphiques générés")
    except Exception as e:
        print(f"Erreur graphiques : {e}")

    # 6. HTML
    html_url = None
    try:
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        html = generate_prereport_html(
            consolidated=consolidated, recommendations=reco, all_results=all_results,
            client_name=client_name, company_name=company_name, trimestre=req.trimestre,
            gauge_png=gauge_png, bar_png=bar_png, radar_png=radar_png,
            dormance_png=dormance_png, breakdown_png=breakdown_png, cash_png=cash_png,
        )
        html_path = f"rapports/{req.user_id}/{ts}_prereport.html"
        sb.storage.from_("client-files").upload(html_path, html.encode('utf-8'),
                                                 {"content-type": "text/html; charset=utf-8"})
        signed = sb.storage.from_("client-files").create_signed_url(html_path, 60*60*24*30)
        html_url = (signed.get("signedURL") or signed.get("signedUrl") or
                   (signed.get("data") or {}).get("signedURL"))
        print("HTML généré")
    except Exception as e:
        print(f"Erreur HTML : {e}")

    # 7. Sauvegarder rapport
    try:
        sb.table("reports").insert({
            "submission_id":  submissions[0]["id"],
            "user_id":        req.user_id,
            "score_kord":     score,
            "summary":        (reco.get("resume_executif") or {}).get("paragraphe_1","") if isinstance(reco.get("resume_executif"), dict) else str(reco.get("resume_executif","")),
            "report_url":     None,
            "prereport_html": html_url,
        }).execute()
        print(f"Rapport sauvegardé — score {score}/100")
    except Exception as e:
        print(f"Erreur sauvegarde : {e}")

    return {"status": "ok", "score": score, "html": html_url is not None}


def _consolidate(results):
    piliers = {
        "stock_cash":          {"score":0,"max":30,"count":0},
        "transport_service":   {"score":0,"max":20,"count":0},
        "achats_fournisseurs": {"score":0,"max":20,"count":0},
        "marges_retours":      {"score":0,"max":15,"count":0},
        "donnees_pilotage":    {"score":0,"max":15,"count":0},
    }
    all_alertes, all_opps = [], []
    for r in results:
        if not isinstance(r, dict): continue
        for p, d in r.get("analyses",{}).items():
            if p in piliers and isinstance(d, dict):
                piliers[p]["score"] += d.get("score",0)
                piliers[p]["count"] += 1
        for a in r.get("alertes",[]):
            if isinstance(a, dict) and a.get("message","") not in [x.get("message","") for x in all_alertes]:
                all_alertes.append(a)
        for o in r.get("opportunites",[]):
            if isinstance(o, dict) and o.get("message","") not in [x.get("message","") for x in all_opps]:
                all_opps.append(o)

    analyses, score_total = {}, 0
    for p, d in piliers.items():
        avg = round(d["score"]/d["count"]) if d["count"] > 0 else round(d["max"]*0.6)
        analyses[p] = {"score":avg,"max":d["max"]}
        score_total += avg

    score_total = min(100, max(0, score_total))
    return {"score_total":score_total,"nb_fichiers":len(results),
            "analyses":analyses,"alertes":all_alertes,"opportunites":all_opps,
            "interpretation":f"Analyse de {len(results)} fichier(s). Score : {score_total}/100."}
