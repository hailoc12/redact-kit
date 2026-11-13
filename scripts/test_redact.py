#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_redact.py — Functional test cho redact_van_ban.py (benchmark deterministic E3/E6).

Cover cả happy-path + BYPASS VARIANTS (separator, header/footer, nested table, metadata,
signature/list names) — chống "báo xanh" sai (theo security critic E8.5).

Chạy:  python3 test_redact.py
Exit 0 = PASS.
"""
import sys, zipfile, re
from pathlib import Path
from docx import Document

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from redact_van_ban import Redactor  # noqa: E402

ASSET = HERE.parent / "samples" / "van-ban-mau-chua-redact.docx"
TMP = HERE / "_test_out"
TMP.mkdir(exist_ok=True)

failures = []
_n = {"p": 0}


def check(cond, msg):
    if cond:
        _n["p"] += 1
    else:
        failures.append(msg)


def text_of(docx_path):
    return "\n".join(p.text for p in Document(str(docx_path)).paragraphs)


def raw_xml(docx_path):
    """Gộp mọi word/*.xml + core.xml thành 1 chuỗi — để grep PII leak."""
    out = []
    with zipfile.ZipFile(str(docx_path)) as z:
        for n in z.namelist():
            if n.endswith(".xml") and (n.startswith("word/") or "core" in n or "app" in n):
                try:
                    out.append(z.read(n).decode("utf-8", "ignore"))
                except Exception:
                    pass
    return "\n".join(out)


# ── 0. sample tồn tại ──
check(ASSET.exists(), f"Thiếu sample {ASSET} — chạy tao_van_ban_mau.py trước")

if ASSET.exists():
    outs, raws = {}, {}
    for lvl in (1, 2, 3, 4):
        out = TMP / f"lv{lvl}.docx"
        r = Redactor(level=lvl)
        r.redact_docx(ASSET, out)
        outs[lvl] = text_of(out)
        raws[lvl] = raw_xml(out)

    # ── luật/điều sống sót mọi cấp (non-target) ──
    for needle in ["NĐ30/2020", "Điều 12", "Khoản 2", "Luật KTNN 2015"]:
        for lvl, t in outs.items():
            check(needle in t, f"[L{lvl}] non-target bị che nhầm: '{needle}'")

    # ── CẤP 1: số định danh tuyệt đối bị che (kể cả có separator — H2/H3) ──
    for needle in ["012345678901", "012345678", "0123456789",
                   "0123-456-789-01", "0123.456.789",          # separator variants
                   "9876543210", "C1234567", "4111 1111 1111 1111",
                   "098765432109"]:                            # nested-table CCCD
        check(needle not in raws[1], f"[L1] chưa che số định danh: '{needle}'")
    check("0912.345.678" in outs[1], "[L1] chớ che SĐT ở cấp 1")
    check("Nguyễn Văn A" in outs[1], "[L1] chớ che tên ở cấp 1")

    # ── CẤP 2: + SĐT/email/mã hồ sơ/ngày sinh (kể cả +84 space — H1) ──
    for needle in ["0912.345.678", "+84 912 345 678", "tranthib@example.gov.vn",
                   "HSKT-2025-0123", "QT/2024/0001", "15/03/1985"]:
        check(needle not in raws[2], f"[L2] chưa che: '{needle}'")

    # ── C2: header/footer bị che ở mọi cấp (PII liên lạc ở header là cấp 2+) ──
    check("0912.345.678" not in raws[2], "[L2] SĐT trong header chưa che")
    check("tranthib@example.gov.vn" not in raws[2], "[L2] email trong footer chưa che")

    # ── C1: metadata bị xoá (author/comments/title) ──
    cp = Document(str(TMP / "lv3.docx")).core_properties
    check(cp.author == "", f"[C1] metadata author chưa xoá: '{cp.author}'")
    check("0912.345.678" not in (cp.comments or ""), "[C1] PII còn trong comments")
    check("HSKT-2025-0123" not in (cp.title or ""), "[C1] PII còn trong title")

    # ── C3: nested table bị che ──
    check("098765432109" not in raws[1], "[C3] CCCD trong nested table chưa che")

    # ── CẤP 3: + tên + ĐVKT + số liệu; + tên trong chữ ký/danh sách (H4) ──
    check("Nguyễn Văn A" not in outs[3], "[L3] chưa che tên người")
    check("Công ty TNHH ABC" not in outs[3], "[L3] chưa che ĐVKT")
    check("1.250.000.000" not in outs[3], "[L3] chưa che số liệu tài chính")
    check("Phạm Hoàng D" not in outs[3], "[L3][H4] chưa che tên sau 'Đại diện:'")
    check("Trần Thị B" not in raws[3], "[L3][H4] chưa che tên trong block chữ ký/danh sách")
    check("[NGƯỜI_1]" in outs[3], "[L3] thiếu token [NGƯỜI_1]")
    check("[STK_1]" in outs[3] and "9876543210" not in outs[3], "[L3] STK sai nhãn/leak")
    check("12," not in outs[3], "[L3] che số liệu nửa vời")

    # ── CẤP 4: + ngày + ký hiệu VB (kể cả QĐ-TTg suffix — M1) ──
    check("12/08/2025" not in outs[4], "[L4] chưa che ngày")
    check("1234/QĐ-TTg" not in raws[4], "[L4][M1] chưa che ký hiệu VB (cả suffix)")
    check("-TTg" not in raws[4].replace("QĐ-TTg", ""), "[L4][M1] leak suffix '-TTg'")

    # ── token không lỗi (thiếu số thứ tự) ──
    for lvl, t in outs.items():
        check("[_]" not in t and "_]_" not in t, f"[L{lvl}] token lỗi (thiếu số thứ tự)")

print(f"PASS: {_n['p']}  FAIL: {len(failures)}")
for f in failures:
    print("  ✗", f)
sys.exit(1 if failures else 0)
