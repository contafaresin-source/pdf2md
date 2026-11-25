"""PDF -> Markdown: app web local (Flask). Execute: python app.py"""
import atexit
import json
import logging
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import uuid
import webbrowser
from logging.handlers import RotatingFileHandler
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory

import engines
import postprocess

BASE = Path(__file__).parent.resolve()
LOG_DIR, TEMP, DEFAULT_OUT = BASE / "logs", BASE / "temp", BASE / "saida"
for d in (LOG_DIR, TEMP, DEFAULT_OUT):
    d.mkdir(exist_ok=True)

HOST, PORT = "127.0.0.1", 5057


# ------------------------------------------------------------------ log
def setup_logging():
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s")
    lg = logging.getLogger("pdf2md")
    lg.setLevel(logging.INFO)
    lg.propagate = False
    app_h = RotatingFileHandler(LOG_DIR / "app.log", maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    err_h = RotatingFileHandler(LOG_DIR / "errors.log", maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    err_h.setLevel(logging.ERROR)
    con_h = logging.StreamHandler()
    for h in (app_h, err_h, con_h):
        h.setFormatter(fmt)
        lg.addHandler(h)
    return lg


log = setup_logging()
app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024**3

JOBS: dict = {}
LOCK = threading.Lock()
Q: "queue.Queue[str]" = queue.Queue()
STATE = {"out_root": str(DEFAULT_OUT)}


# ------------------------------------------------------------------ helpers
def safe_name(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", Path(name).name).strip(" .") or "arquivo.pdf"
    ext = Path(name).suffix.lower()
    stem = Path(name).stem[:100].strip(" .") or "arquivo"
    return stem + (ext if ext in engines.ACCEPTED_EXT else ".pdf")


def unique_outdir(root: Path, stem: str) -> Path:
    used = {j["outdir"] for j in JOBS.values()}
    cand, n = root / stem, 2
    while cand.exists() or str(cand) in used:
        cand = root / f"{stem}_{n}"
        n += 1
    return cand


def normalize_opts(o: dict) -> dict:
    try:
        timeout = int(o.get("timeout", 900))
    except (TypeError, ValueError):
        timeout = 900
    return {
        "ocr": bool(o.get("ocr", False)),
        "no_images": bool(o.get("no_images", False)),
        "strip_refs": bool(o.get("strip_refs", False)),
        "clean": bool(o.get("clean", True)),
        "headings": bool(o.get("headings", False)),
        "keep_docx": bool(o.get("keep_docx", False)),
        "fast": bool(o.get("fast", False)),
        "persistent": bool(o.get("persistent", False)),
        "timeout": max(30, min(timeout, 7200)),
    }


def public(job: dict, detail=False) -> dict:
    keys = ("id", "name", "engine", "status", "message", "outdir", "md", "chars", "tokens", "created")
    d = {k: job.get(k) for k in keys}
    if job.get("started"):
        d["elapsed"] = round((job.get("finished") or time.time()) - job["started"], 1)
    else:
        d["elapsed"] = None
    d["has_detail"] = bool(job.get("detail"))
    if detail:
        d["detail"] = job.get("detail", "")
    return d


def open_path(p):
    p = str(p)
    if engines.IS_WIN:
        os.startfile(p)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", p])
    else:
        subprocess.Popen(["xdg-open", p])


# ------------------------------------------------------------------ worker
def fail(job, message, detail=""):
    log.error("FALHA | arquivo=%s | motor=%s | %s\n%s", job["name"], job["engine"], message, detail)
    job.update(status="error", message=message, detail=(detail or "")[-8000:])


def run_job(job):
    job.update(status="running", started=time.time(), finished=None, message="Convertendo…", detail="")
    pdf, outdir = Path(job["pdf"]), Path(job["outdir"])
    log.info("INÍCIO | arquivo=%s | motor=%s", job["name"], job["engine"])
    try:
        if not pdf.exists():
            raise engines.EngineError("Arquivo temporário não encontrado (arraste o PDF novamente)")
        if outdir.exists():
            shutil.rmtree(outdir, ignore_errors=True)
        outdir.mkdir(parents=True, exist_ok=True)
        md_path = engines.ENGINES[job["engine"]](pdf, outdir, job["opts"])
        text = md_path.read_text(encoding="utf-8", errors="replace")
        text = postprocess.process(text, job["opts"])
        if len(text.strip()) < 50:
            raise engines.EngineError(
                "Resultado vazio ou quase vazio (PDF escaneado? tente ativar OCR ou outro motor)", text)
        md_path.write_text(text, encoding="utf-8")
        job.update(status="done", message="Concluído", md=str(md_path),
                   chars=len(text), tokens=len(text) // 4)
        shutil.rmtree(pdf.parent, ignore_errors=True)
        log.info("OK | arquivo=%s | motor=%s | ~%s tokens", job["name"], job["engine"], job["tokens"])
    except engines.EngineError as e:
        fail(job, str(e), e.detail)
    except Exception as e:  # noqa: BLE001
        fail(job, f"{type(e).__name__}: {e}", traceback.format_exc())
    finally:
        job["finished"] = time.time()


def worker():
    while True:
        jid = Q.get()
        job = JOBS.get(jid)
        if job and job["status"] == "queued":
            run_job(job)


# ------------------------------------------------------------------ rotas
@app.get("/")
def index():
    return send_from_directory(BASE / "static", "index.html")


@app.get("/api/engines")
def api_engines():
    return jsonify(engines.availability())


@app.get("/static/<path:name>")
def static_files(name):
    return send_from_directory(BASE / "static", name)


@app.get("/favicon.ico")
def favicon():
    return send_from_directory(BASE / "static", "icon.svg", mimetype="image/svg+xml")


def _mask(key):
    return (key[:4] + "…" + key[-4:]) if len(key) > 10 else ("•" * len(key))


@app.get("/api/config")
def api_config():
    key = engines.datalab_key()
    return jsonify(out_root=STATE["out_root"], datalab_key=_mask(key) if key else "",
                   datalab_from_env=bool(os.environ.get("DATALAB_API_KEY")))


@app.post("/api/settings")
def api_settings():
    data = request.get_json(silent=True) or {}
    if "datalab_api_key" in data:
        key = str(data["datalab_api_key"] or "").strip()
        engines.set_setting("datalab_api_key", key)
        log.info("Chave da API Datalab %s", "salva" if key else "removida")
    return jsonify(ok=True, engines=engines.availability())


@app.get("/api/workers")
def api_workers():
    return jsonify({k: w.status() for k, w in engines.WORKERS.items()})


@app.post("/api/workers/stop")
def api_workers_stop():
    engines.stop_workers()
    log.info("Workers persistentes encerrados (memória liberada)")
    return jsonify(ok=True)


@app.post("/api/upload")
def api_upload():
    engine = request.form.get("engine", "docling")
    if engine not in engines.ENGINES or engine == "office":
        return jsonify(error="Motor inválido"), 400
    try:
        opts = normalize_opts(json.loads(request.form.get("opts") or "{}"))
    except json.JSONDecodeError:
        opts = normalize_opts({})
    out_root = Path(request.form.get("out") or DEFAULT_OUT).expanduser()
    try:
        out_root.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        log.error("Pasta de saída inválida: %s (%s)", out_root, e)
        return jsonify(error=f"Pasta de saída inválida: {e}"), 400
    STATE["out_root"] = str(out_root)

    created, rejected = [], []
    for f in request.files.getlist("files"):
        name = f.filename or "sem_nome"
        ext = Path(name).suffix.lower()
        if ext not in engines.ACCEPTED_EXT:
            rejected.append(name)
            log.warning("Arquivo ignorado (formato não suportado): %s", name)
            continue
        # documentos Word sempre vão pelo conversor de documentos (Pandoc)
        job_engine = "office" if ext in engines.OFFICE_EXT else engine
        jid = uuid.uuid4().hex[:8]
        try:
            updir = TEMP / jid
            updir.mkdir(parents=True)
            pdf = updir / safe_name(name)
            f.save(pdf)
        except Exception as e:  # noqa: BLE001
            log.error("Falha ao salvar upload %s: %s\n%s", name, e, traceback.format_exc())
            rejected.append(name)
            continue
        with LOCK:
            JOBS[jid] = {
                "id": jid, "name": name, "engine": job_engine, "status": "queued", "message": "Na fila",
                "pdf": str(pdf), "outdir": str(unique_outdir(out_root, pdf.stem)), "opts": opts,
                "created": time.time(), "started": None, "finished": None,
                "md": None, "chars": None, "tokens": None, "detail": "",
            }
        Q.put(jid)
        created.append(jid)
        log.info("ENFILEIRADO | arquivo=%s | motor=%s", name, job_engine)
    return jsonify(created=created, rejected=rejected)


@app.get("/api/jobs")
def api_jobs():
    with LOCK:
        jobs = sorted(JOBS.values(), key=lambda j: j["created"])
        return jsonify([public(j) for j in jobs])


@app.get("/api/jobs/<jid>")
def api_job(jid):
    job = JOBS.get(jid)
    return (jsonify(public(job, detail=True)) if job else (jsonify(error="não encontrado"), 404))


@app.post("/api/jobs/<jid>/retry")
def api_retry(jid):
    job = JOBS.get(jid)
    if not job:
        return jsonify(error="não encontrado"), 404
    if job["status"] != "error":
        return jsonify(error="Só é possível repetir jobs com erro"), 409
    if not Path(job["pdf"]).exists():
        return jsonify(error="Arquivo temporário removido — arraste o PDF novamente"), 410
    engine = (request.get_json(silent=True) or {}).get("engine", job["engine"])
    if engine not in engines.ENGINES:
        return jsonify(error="Motor inválido"), 400
    if (Path(job["pdf"]).suffix.lower() in engines.OFFICE_EXT) != (engine == "office"):
        return jsonify(error="Motor incompatível com o tipo de arquivo"), 400
    job.update(engine=engine, status="queued", message="Na fila", detail="", started=None, finished=None)
    Q.put(jid)
    log.info("REPETIR | arquivo=%s | motor=%s", job["name"], engine)
    return jsonify(ok=True)


@app.delete("/api/jobs/<jid>")
def api_delete(jid):
    job = JOBS.get(jid)
    if not job:
        return jsonify(error="não encontrado"), 404
    if job["status"] == "running":
        return jsonify(error="Job em execução"), 409
    with LOCK:
        JOBS.pop(jid, None)
    shutil.rmtree(Path(job["pdf"]).parent, ignore_errors=True)
    return jsonify(ok=True)


@app.post("/api/clear")
def api_clear():
    with LOCK:
        ids = [i for i, j in JOBS.items() if j["status"] == "done"]
        for i in ids:
            JOBS.pop(i)
    return jsonify(removed=len(ids))


@app.get("/api/jobs/<jid>/download")
def api_download(jid):
    job = JOBS.get(jid)
    if not job or not job.get("md") or not Path(job["md"]).exists():
        return jsonify(error="Arquivo não disponível"), 404
    return send_file(job["md"], as_attachment=True, download_name=Path(job["md"]).name)


@app.get("/api/jobs/<jid>/zip")
def api_zip(jid):
    job = JOBS.get(jid)
    if not job or not Path(job["outdir"]).exists():
        return jsonify(error="Pasta não disponível"), 404
    base = TEMP / f"zip_{jid}"
    archive = shutil.make_archive(str(base), "zip", job["outdir"])
    return send_file(archive, as_attachment=True, download_name=Path(job["outdir"]).name + ".zip")


@app.post("/api/open")
def api_open():
    data = request.get_json(silent=True) or {}
    target = data.get("target")
    try:
        if target == "job":
            job = JOBS.get(data.get("id"))
            path = job["outdir"] if job else None
        elif target == "logs":
            path = LOG_DIR
        else:
            path = STATE["out_root"]
        if not path or not Path(path).exists():
            return jsonify(error="Pasta não existe"), 404
        open_path(path)
        return jsonify(ok=True)
    except Exception as e:  # noqa: BLE001
        log.error("Falha ao abrir pasta: %s", e)
        return jsonify(error=str(e)), 500


@app.get("/api/log")
def api_log():
    f = LOG_DIR / "errors.log"
    if not f.exists():
        return "Nenhum erro registrado.", 200, {"Content-Type": "text/plain; charset=utf-8"}
    data = f.read_bytes()[-150_000:].decode("utf-8", errors="replace")
    return data or "Nenhum erro registrado.", 200, {"Content-Type": "text/plain; charset=utf-8"}


@app.errorhandler(Exception)
def on_error(e):
    from werkzeug.exceptions import HTTPException

    if isinstance(e, HTTPException):
        return jsonify(error=e.description), e.code
    log.error("Erro não tratado na API: %s\n%s", e, traceback.format_exc())
    return jsonify(error=f"Erro interno: {e}"), 500


if __name__ == "__main__":
    atexit.register(engines.stop_workers)
    threading.Thread(target=worker, daemon=True).start()
    if "--no-browser" not in sys.argv:
        threading.Timer(1.2, lambda: webbrowser.open(f"http://{HOST}:{PORT}")).start()
    log.info("Servidor em http://%s:%s", HOST, PORT)
    app.run(host=HOST, port=PORT, threaded=True, debug=False)
