"""Worker persistente: mantém os modelos do Docling/Marker carregados entre conversões.

Roda dentro do ambiente do motor (.venv-docling / .venv-marker ou runtime portátil):
    python engine_worker.py docling|marker

Protocolo: uma requisição JSON por linha no stdin -> {"pdf", "outdir", "opts"}.
Resposta em uma linha do stdout, prefixada por PREFIX: {"ok": true, "md": "..."}
ou {"ok": false, "error": "...", "trace": "..."}. Todo o resto (logs, barras de
progresso das bibliotecas) é desviado para o stderr.
"""
import json
import os
import sys
import time
import traceback
from pathlib import Path

PREFIX = "@@PDF2MD@@"

# Canal exclusivo do protocolo; o fd 1 passa a apontar para o stderr para que
# prints/logs de bibliotecas (inclusive de código C) não corrompam as respostas.
_proto = os.fdopen(os.dup(1), "w", encoding="utf-8", buffering=1)
os.dup2(2, 1)
sys.stdout = sys.stderr


def reply(**data):
    _proto.write(PREFIX + json.dumps(data, ensure_ascii=False) + "\n")
    _proto.flush()


# ---------------------------------------------------------------- Docling
_docling_cache = {}


def _docling_converter(ocr, fast, images):
    key = (ocr, fast, images)
    if key in _docling_cache:
        return _docling_cache[key]
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import (PdfPipelineOptions, TableFormerMode,
                                                    TableStructureOptions)
    from docling.document_converter import DocumentConverter, PdfFormatOption

    po = PdfPipelineOptions(
        do_ocr=ocr,
        do_table_structure=True,
        table_structure_options=TableStructureOptions(
            mode=TableFormerMode.FAST if fast else TableFormerMode.ACCURATE),
    )
    if images:  # igual ao CLI com --image-export-mode referenced
        po.generate_page_images = True
        po.generate_picture_images = True
        po.images_scale = 2
    conv = DocumentConverter(format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=po)})
    _docling_cache[key] = conv
    return conv


def convert_docling(pdf: Path, outdir: Path, opts: dict) -> Path:
    from docling.datamodel.base_models import ConversionStatus
    from docling_core.types.doc import ImageRefMode

    images = not opts.get("no_images")
    conv = _docling_converter(bool(opts.get("ocr")), bool(opts.get("fast")), images)
    res = conv.convert(str(pdf), raises_on_error=True)
    if res.status not in (ConversionStatus.SUCCESS, ConversionStatus.PARTIAL_SUCCESS):
        raise RuntimeError(f"Docling: status {res.status}")
    md = outdir / f"{pdf.stem}.md"
    res.document.save_as_markdown(
        filename=md, image_mode=ImageRefMode.REFERENCED if images else ImageRefMode.PLACEHOLDER)
    return md


# ---------------------------------------------------------------- Marker
_marker_models = {}


def convert_marker(pdf: Path, outdir: Path, opts: dict) -> Path:
    from marker.config.parser import ConfigParser
    from marker.models import create_model_dict
    from marker.output import save_output
    from marker.settings import settings

    kwargs = {"output_format": "markdown", "output_dir": str(outdir)}
    if opts.get("ocr"):
        kwargs["force_ocr"] = True
    elif opts.get("fast"):
        kwargs["mode"] = "no_ocr"  # usa só a camada de texto do PDF
    if opts.get("no_images"):
        kwargs["disable_image_extraction"] = True

    # Mesma escolha de modelos do marker_single (marker/scripts/convert_single.py)
    mode = kwargs.get("mode") or ("balanced" if settings.TORCH_DEVICE_MODEL == "cuda" else "fast")
    load_indra = mode not in ("vllm", "surya", "accurate")
    no_server = mode in ("no_ocr", "fast", "indra")
    key = (load_indra, no_server)
    if key not in _marker_models:
        _marker_models[key] = create_model_dict(load_indra=load_indra, no_server=no_server)

    cp = ConfigParser(kwargs)
    converter = cp.get_converter_cls()(
        config=cp.generate_config_dict(),
        artifact_dict=_marker_models[key],
        processor_list=cp.get_processors(),
        renderer=cp.get_renderer(),
        llm_service=cp.get_llm_service(),
    )
    rendered = converter(str(pdf))
    save_output(rendered, str(outdir), pdf.stem)
    return outdir / f"{pdf.stem}.md"


def main():
    engine = sys.argv[1]
    fn = {"docling": convert_docling, "marker": convert_marker}[engine]
    reply(ok=True, ready=True, engine=engine)
    for line in sys.stdin:
        if not line.strip():
            continue
        try:
            req = json.loads(line)
            t0 = time.time()
            outdir = Path(req["outdir"])
            outdir.mkdir(parents=True, exist_ok=True)
            md = fn(Path(req["pdf"]), outdir, req.get("opts") or {})
            print(f"[worker] {engine}: {Path(req['pdf']).name} em {time.time() - t0:.1f}s", file=sys.stderr)
            reply(ok=True, md=str(md))
        except Exception as e:  # noqa: BLE001
            reply(ok=False, error=f"{type(e).__name__}: {e}", trace=traceback.format_exc())


if __name__ == "__main__":
    main()
