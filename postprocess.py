"""Pós-processamento do Markdown para reduzir tokens e ruído."""
import re
from collections import Counter

DATA_IMG = re.compile(r"!\[[^\]]*\]\(data:image[^)]*\)")
ANY_IMG = re.compile(r"!\[[^\]]*\]\([^)]*\)")
REF_HEAD = re.compile(
    r"^(#{1,6})\s*(?:\d+[.)]?\s*)?(?:\*\*)?"
    r"(references?|referências(?: bibliográficas)?|referencias|bibliography|bibliografia|literature cited|works cited)"
    r"(?:\*\*)?\s*:?\s*$",
    re.I,
)
PAGE_NUM = re.compile(r"\s*(page\s*)?\d{1,4}(\s*(of|/|de)\s*\d{1,4})?\s*", re.I)


def strip_images(text):
    text = ANY_IMG.sub("", text)
    return text.replace("<!-- image -->", "")


def strip_references(text):
    lines = text.splitlines()
    start = level = None
    for i, line in enumerate(lines):
        m = REF_HEAD.match(line.strip())
        if m:
            start, level = i, len(m.group(1))
    if start is None:
        return text
    end = len(lines)
    for j in range(start + 1, len(lines)):
        m = re.match(r"^(#{1,6})\s", lines[j])
        if m and len(m.group(1)) <= level:
            end = j
            break
    return "\n".join(lines[:start] + lines[end:])


def clean_noise(text):
    lines = [l.rstrip() for l in text.splitlines()]

    def candidate(l):
        s = l.strip()
        return (15 <= len(s) <= 100 and not s.startswith(("|", "#", "-", "*", ">", "!", "`", "<"))
                and not re.match(r"^\d+[.)]", s))

    counts = Counter(l.strip() for l in lines if candidate(l))
    repeated = {k for k, v in counts.items() if v >= 4}  # cabeçalho/rodapé repetido
    out = [l for l in lines if l.strip() not in repeated and not PAGE_NUM.fullmatch(l)]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(out)).strip() + "\n"


def process(text, opts):
    text = DATA_IMG.sub("", text)  # imagens base64 só gastam tokens
    if opts.get("no_images"):
        text = strip_images(text)
    if opts.get("strip_refs"):
        text = strip_references(text)
    if opts.get("clean", True):
        text = clean_noise(text)
    return text
