"""Motores de conversão -> Markdown: Docling, Marker (local ou API), Microsoft Word (+ Pandoc)
e documentos Word (.doc/.docx/.rtf/.odt) via Pandoc."""
import base64
import json
import os
import queue
import re
import shutil
import subprocess
import threading
import time
import urllib.error
import urllib.request
import uuid
from collections import Counter, deque
from pathlib import Path

BASE = Path(__file__).parent.resolve()
IS_WIN = os.name == "nt"
NO_WINDOW = subprocess.CREATE_NO_WINDOW if IS_WIN else 0

PDF_EXT = {".pdf"}
OFFICE_EXT = {".doc", ".docx", ".docm", ".rtf", ".odt"}
ACCEPTED_EXT = PDF_EXT | OFFICE_EXT

# ---------------------------------------------------------------- configuração (config.json)
CONFIG_FILE = BASE / "config.json"
_cfg_lock = threading.Lock()


def get_setting(key, default=None):
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8")).get(key, default)
    except (OSError, ValueError):
        return default


def set_setting(key, value):
    with _cfg_lock:
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            data = {}
        if value in (None, ""):
            data.pop(key, None)
        else:
            data[key] = value
        CONFIG_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def datalab_key():
    return os.environ.get("DATALAB_API_KEY") or get_setting("datalab_api_key", "")


class EngineError(Exception):
    """Erro de conversão. `detail` guarda a saída técnica (vai para o log)."""

    def __init__(self, message, detail=""):
        super().__init__(message)
        self.detail = detail


# ---------------------------------------------------------------- disponibilidade
def _venv_exe(venv, name):
    sub = "Scripts" if IS_WIN else "bin"
    exe = BASE / venv / sub / (name + (".exe" if IS_WIN else ""))
    return str(exe) if exe.exists() else shutil.which(name)


ENTRY = {
    "docling": "from docling.cli.main import app; app()",
    "marker": "from marker.scripts.convert_single import convert_single_cli; convert_single_cli()",
}
CLI_NAME = {"docling": "docling", "marker": "marker_single"}


def portable_python():
    """Python embutido da versão portátil (pasta runtime/), se existir."""
    rt = BASE / "runtime"
    if not rt.exists():
        return None
    found = sorted(rt.glob("*/python.exe" if IS_WIN else "*/bin/python3"))
    return str(found[0]) if found else None


def _portable_env(lib):
    """Redireciona caches/modelos para dentro da pasta do app (pendrive)."""
    cache, home = BASE / "cache", BASE / "home"
    for d in (cache / "hf", cache / "torch", cache / "datalab", cache / "xdg",
              home / "AppData" / "Local", home / "AppData" / "Roaming"):
        d.mkdir(parents=True, exist_ok=True)
    env = {
        "HF_HOME": str(cache / "hf"), "TORCH_HOME": str(cache / "torch"),
        "MODEL_CACHE_DIR": str(cache / "datalab"), "XDG_CACHE_HOME": str(cache / "xdg"),
        "USERPROFILE": str(home), "HOME": str(home),
        "LOCALAPPDATA": str(home / "AppData" / "Local"), "APPDATA": str(home / "AppData" / "Roaming"),
        "PYTHONPATH": str(BASE / "libs" / lib), "PYTHONNOUSERSITE": "1", "PYTHONDONTWRITEBYTECODE": "1",
    }
    if not os.environ.get("PDF2MD_ONLINE"):
        env["HF_HUB_OFFLINE"] = "1"  # modelos já estão no pendrive; evita esperas sem internet
    return env


def resolve(engine):
    """Retorna (prefixo_do_comando, env_extra) para docling/marker, ou (None, None)."""
    py = portable_python()
    if py and (BASE / "libs" / engine).exists():
        return [py, "-X", "utf8", "-c", ENTRY[engine]], _portable_env(engine)
    exe = _venv_exe(f".venv-{engine}", CLI_NAME[engine])
    return ([exe], None) if exe else (None, None)


def word_available():
    if not IS_WIN:
        return False, "Disponível apenas no Windows"
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "pywin32 não instalado (rode setup.bat)"
    try:
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
    except OSError:
        return False, "Microsoft Word não encontrado"
    try:
        import pypandoc  # noqa: F401
    except ImportError:
        return False, "pypandoc_binary não instalado (rode setup.bat)"
    return True, ""


def resolve_python(engine):
    """Python do ambiente do motor (para o worker persistente): (python, env_extra) ou (None, None)."""
    py = portable_python()
    if py and (BASE / "libs" / engine).exists():
        return py, _portable_env(engine)
    sub, exe = ("Scripts", "python.exe") if IS_WIN else ("bin", "python")
    p = BASE / f".venv-{engine}" / sub / exe
    return (str(p), None) if p.exists() and resolve(engine)[0] else (None, None)


def pandoc_available():
    try:
        import pypandoc  # noqa: F401
        return True
    except ImportError:
        return False


def availability():
    d, m = resolve("docling")[0], resolve("marker")[0]
    w_ok, w_why = word_available()
    k = bool(datalab_key())
    return {
        "docling": {"available": bool(d), "reason": "" if d else "Não instalado — rode: setup.bat -Docling"},
        "marker": {"available": bool(m), "reason": "" if m else "Não instalado — rode: setup.bat -Marker"},
        "marker_api": {"available": k, "reason": "" if k else "Informe a chave da API Datalab abaixo"},
        "word": {"available": w_ok, "reason": w_why},
        # documentos Word: .docx/.odt só precisam do Pandoc; .doc/.rtf precisam do Word instalado
        "office": {"available": pandoc_available(), "word": w_ok,
                   "reason": "" if pandoc_available() else "pypandoc_binary não instalado (rode setup.bat)"},
    }


# ---------------------------------------------------------------- utilidades
_ANSI = re.compile(r"\x1b\[[0-9;?]*[A-Za-z]")


def _clean_output(raw: bytes) -> str:
    text = raw.decode("utf-8", errors="replace").replace("\r", "\n")
    text = _ANSI.sub("", text)
    return re.sub(r"\n{2,}", "\n", text)


def _kill_tree(p):
    try:
        if IS_WIN:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(p.pid)],
                           capture_output=True, creationflags=NO_WINDOW)
        else:
            p.kill()
    except Exception:
        pass


def run_cmd(cmd, timeout, cwd=None, what="Processo", env_extra=None):
    env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", **(env_extra or {}))
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                         cwd=cwd, env=env, creationflags=NO_WINDOW)
    try:
        out, _ = p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill_tree(p)
        out, _ = p.communicate()
        raise EngineError(f"Tempo limite excedido ({timeout}s)", _clean_output(out)[-6000:])
    text = _clean_output(out)
    if p.returncode != 0:
        last = next((l for l in reversed(text.splitlines()) if l.strip()), "sem saída")
        raise EngineError(f"{what} falhou (código {p.returncode}): {last[:300]}",
                          f"Comando: {' '.join(map(str, cmd))}\n\n{text[-6000:]}")
    return text


def _find_md(outdir: Path, stem: str) -> Path:
    mds = list(outdir.rglob("*.md"))
    if not mds:
        raise EngineError("O motor terminou sem gerar nenhum .md", "")
    exact = [m for m in mds if m.stem == stem]
    return exact[0] if exact else max(mds, key=lambda m: m.stat().st_size)


# ---------------------------------------------------------------- worker persistente
WORKER_SCRIPT = BASE / "engine_worker.py"
WORKER_PREFIX = "@@PDF2MD@@"
WORKER_IDLE_SECONDS = 20 * 60  # libera a memória após 20 min sem uso


class PersistentWorker:
    """Processo do motor que fica vivo entre conversões (modelos carregados uma só vez)."""

    def __init__(self, engine):
        self.engine = engine
        self.proc = None
        self.lines = None
        self.log = deque(maxlen=400)
        self.lock = threading.Lock()
        self.started = None
        self.last_used = None
        self.count = 0

    def alive(self):
        return self.proc is not None and self.proc.poll() is None

    def status(self):
        return {"running": self.alive(), "started": self.started if self.alive() else None,
                "last_used": self.last_used, "count": self.count if self.alive() else 0}

    def _pump(self, stream, is_proto):
        for raw in iter(stream.readline, b""):
            line = raw.decode("utf-8", errors="replace").rstrip("\r\n")
            if is_proto and line.startswith(WORKER_PREFIX):
                self.lines.put(line[len(WORKER_PREFIX):])
            elif line.strip():
                self.log.append(_ANSI.sub("", line))
        if is_proto:
            self.lines.put(None)  # processo terminou

    def _start(self):
        py, env_extra = resolve_python(self.engine)
        if not py:
            raise EngineError(f"{LABELS[self.engine]} não está instalado.")
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", **(env_extra or {}))
        self.lines, self.log = queue.Queue(), deque(maxlen=400)
        self.proc = subprocess.Popen([py, "-X", "utf8", str(WORKER_SCRIPT), self.engine],
                                     stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                     env=env, creationflags=NO_WINDOW)
        threading.Thread(target=self._pump, args=(self.proc.stdout, True), daemon=True).start()
        threading.Thread(target=self._pump, args=(self.proc.stderr, False), daemon=True).start()
        self.started, self.count = time.time(), 0

    def _read(self, deadline):
        try:
            line = self.lines.get(timeout=max(1, deadline - time.time()))
        except queue.Empty:
            self.stop()
            raise EngineError("Tempo limite excedido (worker reiniciado)", "\n".join(self.log)[-6000:])
        if line is None:
            raise EngineError(f"O processo do {LABELS[self.engine]} encerrou inesperadamente",
                              "\n".join(self.log)[-6000:])
        return json.loads(line)

    def convert(self, pdf: Path, outdir: Path, opts: dict, timeout: int) -> Path:
        with self.lock:
            deadline = time.time() + timeout
            if not self.alive():
                self._start()
                self._read(deadline)  # mensagem "ready"
            self.log.clear()
            req = {"pdf": str(pdf), "outdir": str(outdir), "opts": opts}
            try:
                self.proc.stdin.write((json.dumps(req, ensure_ascii=False) + "\n").encode("utf-8"))
                self.proc.stdin.flush()
            except OSError as e:
                self.stop()
                raise EngineError(f"Falha ao falar com o worker: {e}", "\n".join(self.log)[-6000:])
            resp = self._read(deadline)
            self.last_used = time.time()
            self.count += 1
            if not resp.get("ok"):
                raise EngineError(f"{LABELS[self.engine]} falhou: {resp.get('error', '')[:300]}",
                                  (resp.get("trace") or "") + "\n\n" + "\n".join(self.log)[-4000:])
            return Path(resp["md"])

    def stop(self):
        p, self.proc = self.proc, None
        if p is not None and p.poll() is None:
            try:
                p.stdin.close()
            except OSError:
                pass
            _kill_tree(p)

    def stop_if_idle(self):
        if self.alive() and self.last_used and time.time() - self.last_used > WORKER_IDLE_SECONDS:
            if self.lock.acquire(blocking=False):
                try:
                    self.stop()
                finally:
                    self.lock.release()


WORKERS = {"docling": PersistentWorker("docling"), "marker": PersistentWorker("marker")}


def stop_workers():
    for w in WORKERS.values():
        w.stop()


def _idle_watch():
    while True:
        time.sleep(60)
        for w in WORKERS.values():
            w.stop_if_idle()


threading.Thread(target=_idle_watch, daemon=True).start()


# ---------------------------------------------------------------- Docling
def convert_docling(pdf: Path, outdir: Path, opts: dict) -> Path:
    if opts.get("persistent"):
        return WORKERS["docling"].convert(pdf, outdir, opts, opts["timeout"])
    prefix, env = resolve("docling")
    if not prefix:
        raise EngineError("Docling não está instalado. Rode: setup.bat -Docling")
    cmd = prefix + [str(pdf), "--from", "pdf", "--to", "md", "--output", str(outdir),
                    "--image-export-mode", "placeholder" if opts.get("no_images") else "referenced",
                    "--ocr" if opts.get("ocr") else "--no-ocr"]
    if opts.get("fast"):
        cmd += ["--table-mode", "fast"]
    run_cmd(cmd, opts["timeout"], what="Docling", env_extra=env)
    return _find_md(outdir, pdf.stem)


# ---------------------------------------------------------------- Marker
def convert_marker(pdf: Path, outdir: Path, opts: dict) -> Path:
    if opts.get("persistent"):
        WORKERS["marker"].convert(pdf, outdir, opts, opts["timeout"])
    else:
        prefix, env = resolve("marker")
        if not prefix:
            raise EngineError("Marker não está instalado. Rode: setup.bat -Marker")
        cmd = prefix + [str(pdf), "--output_dir", str(outdir), "--output_format", "markdown"]
        if opts.get("ocr"):
            cmd.append("--force_ocr")
        elif opts.get("fast"):
            cmd += ["--mode", "no_ocr"]  # só a camada de texto do PDF
        if opts.get("no_images"):
            cmd.append("--disable_image_extraction")
        run_cmd(cmd, opts["timeout"], what="Marker", env_extra=env)
    # Marker cria outdir/<stem>/<stem>.md -> sobe um nível para ficar mais limpo
    sub = outdir / pdf.stem
    if sub.is_dir():
        for item in list(sub.iterdir()):
            dest = outdir / item.name
            if not dest.exists():
                shutil.move(str(item), str(dest))
        shutil.rmtree(sub, ignore_errors=True)
    for meta in outdir.glob("*_meta.json"):
        meta.unlink(missing_ok=True)
    return _find_md(outdir, pdf.stem)


# ---------------------------------------------------------------- Word + Pandoc
def _promote_headings(doc):
    """Heurística (experimental): aplica estilos Heading 1-4 por tamanho/negrito da fonte."""
    paras = doc.Paragraphs
    n = paras.Count
    rows, weight = [], Counter()
    for i in range(1, n + 1):
        try:
            r = paras.Item(i).Range
            text = r.Text.strip()
            if not text:
                continue
            size = float(r.Font.Size)
            if size > 1000:  # fonte mista
                continue
            bold = r.Font.Bold == -1 or r.Font.Bold is True
            weight[round(size, 1)] += len(text)
            rows.append((i, text, round(size, 1), bold))
        except Exception:
            continue
    if not weight:
        return
    body = weight.most_common(1)[0][0]
    big, cands = set(), []
    for i, text, size, bold in rows:
        if len(text) > 120 or text.endswith((".", ";", ",")) or text.isdigit():
            continue
        if size >= body + 1.5:
            big.add(size)
            cands.append((i, size, True))
        elif bold and size >= body - 0.5 and len(text) <= 80:
            cands.append((i, size, False))
    levels = {s: k + 1 for k, s in enumerate(sorted(big, reverse=True))}
    bold_level = min(len(levels) + 1, 4)
    for i, size, is_big in cands:
        try:
            p = paras.Item(i)
            if p.Range.Information(12):  # dentro de tabela
                continue
            level = min(levels.get(size, bold_level), 4)
            p.Range.Style = doc.Styles.Item(-1 - level)  # wdStyleHeading1 = -2 ...
        except Exception:
            continue


def _word_to_docx(src: Path, docx: Path, headings=False):
    """Abre `src` (PDF, .doc, .rtf...) no Word e salva como .docx."""
    ok, why = word_available()
    if not ok:
        raise EngineError(f"Word indisponível: {why}")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = doc = None
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
            doc = word.Documents.Open(str(src), False, True, False)
            if headings:
                _promote_headings(doc)
            doc.SaveAs2(str(docx), 16)  # wdFormatXMLDocument
        except EngineError:
            raise
        except Exception as e:
            raise EngineError(f"Falha ao usar o Word: {e}", repr(e))
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _pandoc_to_md(src: Path, outdir: Path, md: Path, fmt: str, opts: dict):
    import pypandoc

    cmd = [pypandoc.get_pandoc_path(), str(src), "-f", fmt, "-t", "gfm", "--wrap=none", "-o", md.name]
    if not opts.get("no_images"):
        cmd.insert(-2, "--extract-media=assets")
    run_cmd(cmd, opts["timeout"], cwd=str(outdir), what="Pandoc")


def convert_word(pdf: Path, outdir: Path, opts: dict) -> Path:
    docx = outdir / f"{pdf.stem}.docx"
    _word_to_docx(pdf, docx, opts.get("headings"))
    md = outdir / f"{pdf.stem}.md"
    _pandoc_to_md(docx, outdir, md, "docx", opts)
    if not opts.get("keep_docx"):
        docx.unlink(missing_ok=True)
    return md


# ---------------------------------------------------------------- Documentos Word (.doc/.docx/...)
def convert_office(src: Path, outdir: Path, opts: dict) -> Path:
    """.docx/.odt vão direto para o Pandoc (preservam títulos, listas e tabelas);
    .doc/.docm/.rtf passam antes pelo Word para virar .docx."""
    if not pandoc_available():
        raise EngineError("pypandoc_binary não instalado (rode setup.bat)")
    ext = src.suffix.lower()
    md = outdir / f"{src.stem}.md"
    if ext in (".docx", ".odt"):
        _pandoc_to_md(src, outdir, md, "docx" if ext == ".docx" else "odt", opts)
        return md
    if ext == ".rtf" and not word_available()[0]:
        _pandoc_to_md(src, outdir, md, "rtf", opts)  # Pandoc 3 lê RTF (com menos fidelidade)
        return md
    if not word_available()[0]:
        raise EngineError(f"Arquivos {ext} precisam do Microsoft Word instalado "
                          "(salve como .docx no Word ou LibreOffice e tente de novo)")
    docx = outdir / f"{src.stem}.docx"
    _word_to_docx(src, docx)
    _pandoc_to_md(docx, outdir, md, "docx", opts)
    if not opts.get("keep_docx"):
        docx.unlink(missing_ok=True)
    return md


# ---------------------------------------------------------------- Marker via API (Datalab)
def datalab_url():
    return (os.environ.get("DATALAB_API_URL") or get_setting("datalab_api_url")
            or "https://www.datalab.to/api/v1/marker")


def _http_json(req, timeout=120):
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")[:2000]
        msg = {401: "chave da API inválida", 403: "acesso negado (chave/créditos)",
               413: "arquivo grande demais para a API", 429: "limite de requisições atingido"}.get(e.code, "")
        raise EngineError(f"API Datalab respondeu HTTP {e.code}{' — ' + msg if msg else ''}", body)
    except urllib.error.URLError as e:
        raise EngineError(f"Sem conexão com a API Datalab: {e.reason}", repr(e))


def convert_marker_api(pdf: Path, outdir: Path, opts: dict) -> Path:
    key = datalab_key()
    if not key:
        raise EngineError("Chave da API Datalab não configurada")
    fields = {"output_format": "markdown", "paginate": "false",
              "force_ocr": "true" if opts.get("ocr") else "false",
              "disable_image_extraction": "true" if opts.get("no_images") else "false",
              "mode": "fast" if opts.get("fast") else "balanced"}
    boundary = uuid.uuid4().hex
    body = bytearray()
    for k, v in fields.items():
        body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n").encode()
    ascii_name = re.sub(r"[^A-Za-z0-9._-]", "_", pdf.name)
    body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{ascii_name}\"\r\n"
             f"Content-Type: application/pdf\r\n\r\n").encode()
    body += pdf.read_bytes() + f"\r\n--{boundary}--\r\n".encode()
    headers = {"X-Api-Key": key, "User-Agent": "pdf2md-local"}
    req = urllib.request.Request(datalab_url(), data=bytes(body), method="POST",
                                 headers={**headers, "Content-Type": f"multipart/form-data; boundary={boundary}"})
    first = _http_json(req)
    check = first.get("request_check_url")
    if not first.get("success", True) or not check:
        raise EngineError(f"API recusou o arquivo: {first.get('error') or 'sem request_check_url'}",
                          json.dumps(first, ensure_ascii=False)[:4000])

    deadline = time.time() + opts["timeout"]
    while True:
        time.sleep(2)
        data = _http_json(urllib.request.Request(check, headers=headers))
        if data.get("status") == "complete":
            break
        if time.time() > deadline:
            raise EngineError(f"Tempo limite excedido ({opts['timeout']}s) aguardando a API", check)
    if not data.get("success"):
        raise EngineError(f"API Datalab falhou: {data.get('error') or 'erro desconhecido'}",
                          json.dumps({k: v for k, v in data.items() if k != "images"}, ensure_ascii=False)[:4000])

    md = outdir / f"{pdf.stem}.md"
    md.write_text(data.get("markdown") or "", encoding="utf-8")
    for name, b64 in (data.get("images") or {}).items():
        safe = Path(name).name
        if safe:
            (outdir / safe).write_bytes(base64.b64decode(b64))
    return md


ENGINES = {"docling": convert_docling, "marker": convert_marker, "marker_api": convert_marker_api,
           "word": convert_word, "office": convert_office}
LABELS = {"docling": "Docling", "marker": "Marker", "marker_api": "Marker API",
          "word": "Word", "office": "Documento Word"}
