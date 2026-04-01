from __future__ import annotations

import re
from io import BytesIO

try:
    from pypdf import PdfReader
except ModuleNotFoundError:  # pragma: no cover
    PdfReader = None


AMOUNT_PATTERNS = [
    r"价税合计[（(]?大写[）)]?.{0,20}小写[:：]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
    r"金额合计[:：]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
    r"合计[:：]?\s*￥?\s*([0-9]+(?:\.[0-9]{1,2})?)",
    r"价税合计[:：]?\s*￥?\s*([0-9]+(?:\.[0-9]{1,2})?)",
]


def extract_invoice_amount(pdf_bytes: bytes) -> str:
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
    except Exception:
        return ""

    texts: list[str] = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            continue
    content = "\n".join(texts)
    if not content.strip():
        return ""

    for pattern in AMOUNT_PATTERNS:
        matched = re.search(pattern, content, re.IGNORECASE | re.DOTALL)
        if matched:
            return matched.group(1)

    candidates = re.findall(r"(?<!\d)(\d{1,8}\.\d{2})(?!\d)", content)
    if not candidates:
        return ""
    return max(candidates, key=lambda value: float(value))
