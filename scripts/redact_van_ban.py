#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
redact_van_ban.py — Che (redact) thông tin nhạy cảm khỏi văn bản .docx trước khi đưa vào AI.

Triết lý (redact trước khi đưa tài liệu vào AI):
  - DETERMINISTIC-FIRST: mọi che mất được làm bằng regex/pattern (đếm được, đo được),
    KHÔNG dùng AI/LLM để "tự nhận diện". Lý do: dữ liệu nhạy cảm không được
    phụ thuộc đoán mớ của AI — phải biết chắc chắn đã che gì.
  - 4 CẤP từ nhẹ → triệt để. Người dùng chỉ cần chọn 1 số.
  - ENTITY CONSISTENCY: cùng tên / cùng số → cùng token ([NGƯỜI_1], [ĐVKT_1]) để văn bản
    redact vẫn đọc được, không rối.
  - HITL: các category có độ tin cậy thấp (nhận diện tên riêng tiếng Việt bằng regex)
    bị gắn cờ "cần người duyệt" trong báo cáo. Code KHÔNG giả vờ che hết.

Cách chạy:
  python3 scripts/redact_van_ban.py <file_nhập.docx|.txt|.md> --level <1..4> --outdir <thư_mục_ra>

Output:
  <thư_mục_ra>/<tên>_redacted.docx        ← bản bảo mật, mang đi cho AI
  <thư_mục_ra>/<tên>_redaction_report.md  ← báo cáo: đã che gì, bao nhiêu, cờ HITL
  <thư_mục_ra>/<tên>_restore_map.json     ← bảng token→gốc (LUÔN sinh ở mọi cấp —
                                             dùng cho unredact_van_ban.py khôi phục;
                                             BẢO MẬT — xoá khi không còn cần)

Giới hạn (trung thực):
  - Regex nhận diện tên riêng tiếng Việt là XẤP XỈ (dựa vào danh xưng ông/bà/anh/chị/
    đồng chí + từ viết hoa). Luôn có cờ HITL → kiểm toán viên rà lại.
  - Chỉ hỗ trợ .docx (văn bản hành chính chính). .txt/.md đọc raw.
  - Không gửi dữ liệu đi đâu cả — chạy local 100%.
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    from docx import Document
except ImportError:
    sys.stderr.write(
        "Thiếu thư viện python-docx. Cài:  pip install python-docx\n"
    )
    sys.exit(2)


# ════════════════════════════════════════════════════════════════════
#  CẤU HÌNH 4 CẤP — "cấp nào che gì"
#  level = số nguyên dương; một pattern thuộc cấp N sẽ chạy khi --level >= N.
#  confidence: 1.0 = chắc chắn (số định dạng cố định) · 0.5-0.7 = xấp xỉ (tên/đơn vị)
#              → < 0.7 sẽ bị gắn cờ HITL trong báo cáo.
# ════════════════════════════════════════════════════════════════════

# Danh xưng tiếng Việt thường đi kèm tên riêng (để phát hiện tên — xấp xỉ)
_TITLES = r"(?:(?:Ông|Bà|Anh|Chị|Đồng chí|Cô|Giám đốc|Phó giám đốc|Trưởng phòng|Ký tên|Người lập|Đại diện|Thực hiện|Đệ trình|Lập bởi|Chịu trách nhiệm)\s*[:\-]?)"

# Đơn vị được kiểm toán / cơ quan — nhận theo từ khoá dẫn + cụm viết hoa
_ORG_KEYWORDS = r"(?:ĐVKT|đơn vị được kiểm toán|Khách hàng|Cơ quan|đơn vị|tại)"

# Các loại pattern. Thứ tự trong danh sách = thứ tự áp dụng (quan trọng cho overlap).
PATTERNS = [
    # ── CẤP 1 — TỐI THIỂU: số định danh tuyệt đối (confidence 1.0) ──
    # SO_TAI_KHOAN đặt TRƯỚC MST/CCCD: 10 số sau từ khoá "tài khoản" phải là STK, không phải MST.
    {
        "id": "SO_TAI_KHOAN",
        "level": 1, "confidence": 1.0,
        "token": "[STK]",
        "regex": r"(?i)(?:số\s*tài\s*khoản|STK|tk)\s*[:\-]?\s*\d{9,14}",
        "note": "Số tài khoản ngân hàng 9-14 số (cần từ khoá dẫn)",
    },
    {
        "id": "CCCD_CMND",
        "level": 1, "confidence": 1.0,
        "token": "[CCCD]",
        # Keyword-anchored (chấp nhận separator mọi nhóm) HOẶC bare liên tiếp đúng 9/12 số.
        # Tránh false-positive lên số tiền có dấu chấm (1.250.000.000) vì cần keyword hoặc liên tiếp.
        "regex": r"(?i)(?:CCCD|CMND|CMT|CMTND)\s*[:\-]?\s*\d(?:[\s.\-]?\d){8,11}"
                 r"|\b\d{9}\b|\b\d{12}\b",
        "digits": {9, 12},
        "note": "9 số (CMND) hoặc 12 số (CCCD) — kể cả 'CCCD 0123-456-789-01'",
    },
    {
        "id": "MST",
        "level": 1, "confidence": 1.0,
        "token": "[MST]",
        "regex": r"(?i)(?:MST|mã\s*số\s*thuế)\s*[:\-]?\s*\d(?:[\s.\-]?\d){9,12}"
                 r"|\b\d{10}(?:-\d{3})?\b",
        "digits": {10, 13},
        "note": "Mã số thuế 10/13 số — kể cả 'MST 0123-456-789', 'MST 0123.456.789'",
    },
    {
        "id": "SO_THE",
        "level": 1, "confidence": 1.0,
        "token": "[SỐ_THẺ]",
        "regex": r"(?:(?:\d[ -]?){13,18})",
        "note": "Số thẻ tín dụng 13-18 số (có thể nhóm)",
    },
    {
        "id": "HO_CHIEU",
        "level": 1, "confidence": 1.0,
        "token": "[HỘ_CHIẾU]",
        "regex": r"\b[A-Z]\d{7}\b",
        "note": "Số hộ chiếu (1 chữ + 7 số)",
    },

    # ── CẤP 2 — CƠ BẢN: + liên lạc + mã hồ sơ (confidence 1.0) ──
    {
        "id": "SDT",
        "level": 2, "confidence": 1.0,
        "token": "[SĐT]",
        "regex": r"(?:(?:\+84|0)[\s\-]?\d(?:[\s.\-]?\d){7,9})",
        "note": "SĐT Việt Nam (+84/0x...), kể cả '+84 912 345 678'",
    },
    {
        "id": "EMAIL",
        "level": 2, "confidence": 1.0,
        "token": "[EMAIL]",
        "regex": r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
        "note": "Địa chỉ email",
    },
    {
        "id": "MA_HO_SO_KT",
        "level": 2, "confidence": 1.0,
        "token": "[MÃ_HỒ_SƠ]",
        "regex": r"(?i)\b(?:HSKT|HSC?T?|hồ sơ KT|QT)\s*[\-/]?\s*\d{2,4}[\-/]\d{1,6}\b",
        "note": "Mã hồ sơ kiểm toán (HSKT-2025-0123, QT/2024/0001)",
    },
    {
        "id": "NGAY_SINH",
        "level": 2, "confidence": 0.9,
        "token": "[NGÀY_SINH]",
        "regex": r"(?i)(?:ngày\s*sinh|N/S|sinh\s*ngày)\s*[:\-]?\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",
        "note": "Ngày sinh (cần từ khoá dẫn — tránh che mọi ngày)",
    },

    # ── CẤP 3 — NÂNG CAO: + tên người + đơn vị + số liệu (xấp xỉ → HITL) ──
    {
        "id": "TEN_NGUOI",
        "level": 3, "confidence": 0.5,
        "token": "[NGƯỜI]",
        "regex": rf"{_TITLES}\s+[A-ZÀ-ỽĐ][A-Za-zÀ-ỹ\s]{{2,30}}",
        "note": "Tên sau danh xưng/chức danh/ký tên — XẤP XỈ, cờ HITL",
    },
    {
        "id": "TEN_NGUOI_LIST",
        "level": 3, "confidence": 0.4,
        "token": "[NGƯỜI]",
        # 2-4 từ viết hoa liên tiếp đứng sau ngắt câu/dòng/danh xưng ngầm — bắt tên trong
        # danh sách nhân sự, block chữ ký, "...ông A, bà B, anh C". Sai số cao → HITL.
        "regex": r"(?:^|\n|[.;,]\s+|và\s+|và\s)[A-ZÀ-ỽĐ][a-zà-ỹ]{1,15}\s+[A-ZÀ-ỽĐ][a-zà-ỹ]{1,15}"
                 r"(?:\s+[A-ZÀ-ỽĐ][a-zà-ỹ]{1,15}){0,2}",
        "note": "Chuỗi tên riêng viết hoa liên tiếp (danh sách/chữ ký) — XẤP XỈ CAO, cờ HITL",
    },
    {
        "id": "DON_VI_KT",
        "level": 3, "confidence": 0.6,
        "token": "[ĐVKT]",
        "regex": rf"(?:{_ORG_KEYWORDS})\s+"
                  r"[A-ZÀ-Ỹ0-9][A-Za-zÀ-ỹ0-9 .&\-]{2,50}",
        "note": "Tên đơn vị được kiểm toán/cơ quan sau từ khoá — XẤP XỈ, cờ HITL",
    },
    {
        "id": "SO_LIEU_TAI_CHINH",
        "level": 3, "confidence": 0.65,
        "token": "[SỐ_LIỆU_TÀI_CHÍNH]",
        "regex": r"(?i)\b\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{1,3})?\s*(?:đồng|đ|VND|VNĐ|ngàn|nghìn|triệu|tỷ|tỉ|USD)\b",
        "note": "Số kèm đơn vị tiền (1.250.000 đồng / 500 triệu / 12,5 tỷ) — có thể che nhầm số lượng; cờ HITL",
    },

    # ── CẤP 4 — TRIỆT ĐỂ: + ngày + ký hiệu văn bản + metadata ──
    {
        "id": "NGAY_THANG",
        "level": 4, "confidence": 0.6,
        "token": "[NGÀY]",
        "regex": r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
        "note": "Mọi ngày tháng dạng dd/mm/yyyy — có thể che ngày công văn hợp lệ; cờ HITL",
    },
    {
        "id": "KY_HIEU_VAN_BAN",
        "level": 4, "confidence": 0.8,
        "token": "[KÝ HIỆU VB]",
        "regex": r"\b\d{1,4}/[A-ZÀ-ỽĐ]{2,10}(?:-[A-Za-zÀ-ỹ0-9]{1,10})?\b",
        "note": "Ký hiệu văn bản (123/KTNN-PT, 45/QĐ-TTg) — che cả suffix cơ quan",
    },
    {
        "id": "SO_VAN_BAN",
        "level": 4, "confidence": 0.7,
        "token": "[SỐ VB]",
        "regex": r"(?i)(?:số|công văn số|thông báo số)\s*[:\-]?\s*\d{1,4}/[A-ZÀ-ỽĐ0-9\-]{2,20}/\d{2,4}",
        "note": "Số/ký hiệu công văn theo NĐ30 — cờ HITL (có thể cần giữ cho thể thức)",
    },
]

# Giữ nguyên (KHÔNG che dù khớp): luật, điều, nghị định — tránh phá văn bản pháp lý
NEVER_TOUCH = [
    re.compile(r"\bNĐ\s*\d{1,3}/\s*\d{4}\b", re.I),          # NĐ30/2020
    re.compile(r"\b(?:Điều|Khoản|Điểm)\s*\d+[A-Za-z]?\b", re.I),  # Điều 12
    re.compile(r"\b(?:Luật|Bộ luật)\s+[A-ZÀ-ỽĐ][A-Za-zÀ-ỹ ]+", ), # Luật KTNN
    re.compile(r"\b\d{4}/\s*\d{1,2}-\s*QH\d{0,3}\b"),         # symbol luật QH
]

LEVEL_NAMES = {
    1: "Tối thiểu",
    2: "Cơ bản (mặc định)",
    3: "Nâng cao",
    4: "Triệt để",
}


# ════════════════════════════════════════════════════════════════════
#  ENGINE
# ════════════════════════════════════════════════════════════════════

class Redactor:
    def __init__(self, level: int):
        self.level = level
        self.token_map: dict[str, str] = {}      # original → token
        self.counters: dict[str, int] = {}        # category → số lần
        self.hitl: list[dict] = []                # các thay đổi confidence < 0.7
        self.changelog: list[dict] = []           # mọi thay đổi (cho report)

    def _get_token(self, cat_id: str, token_base: str, original: str) -> str:
        """Entity consistency: cùng original → cùng token. Khác original cùng loại → số thứ tự mới."""
        if original in self.token_map:
            return self.token_map[original]
        n = self.counters.get(cat_id, 0) + 1
        self.counters[cat_id] = n
        # token_base dạng "[NGƯỜI]" → "[NGƯỜI_1]"
        if token_base.endswith("]"):
            token = token_base[:-1] + f"_{n}]"
        else:
            token = f"{token_base}_{n}"
        self.token_map[original] = token
        return token

    def _is_protected(self, text: str, start: int, end: int) -> bool:
        """True nếu đoạn khớp nằm trong vùng pháp lý cần giữ."""
        window = text[max(0, start - 40): min(len(text), end + 40)]
        return any(p.search(window) for p in NEVER_TOUCH)

    def redact_text(self, text: str) -> str:
        if not text:
            return text
        active = [p for p in PATTERNS if p["level"] <= self.level]
        # Áp dụng theo level (cấp thấp trước), mỗi category độc lập
        for pat in active:
            regex = re.compile(pat["regex"])
            allowed_digits = pat.get("digits")  # None hoặc set nếu cần kiểm số chữ số

            def _repl(m, _pat=pat, _ad=allowed_digits):
                original = m.group(0)
                if self._is_protected(text, m.start(), m.end()):
                    return original  # giữ nguyên (luật/điều)
                # Kiểm số chữ số nếu pattern yêu cầu (CCCD/MST có separator)
                if _ad is not None:
                    ndigits = sum(c.isdigit() for c in original)
                    if ndigits not in _ad:
                        return original  # không đúng độ dài → không che (tránh che bừa)
                token = self._get_token(_pat["id"], _pat["token"], original)
                entry = {
                    "category": _pat["id"],
                    "token": token,
                    "original_preview": (original[:30] + "…") if len(original) > 30 else original,
                    "confidence": _pat["confidence"],
                    "hitl": _pat["confidence"] < 0.7,
                }
                self.changelog.append(entry)
                if entry["hitl"]:
                    self.hitl.append(entry)
                return token

            text = regex.sub(_repl, text)
        return text

    # ── paragraph helper (gom runs → redact → ghi lại run đầu) ──
    def _redact_paragraph(self, para):
        runs = para.runs
        if not runs:
            return
        full = "".join(r.text for r in runs)
        new = self.redact_text(full)
        if new != full:
            runs[0].text = new
            for r in runs[1:]:
                r.text = ""

    def _redact_cell(self, cell):
        for para in cell.paragraphs:
            self._redact_paragraph(para)
        for nested in cell.tables:                  # C3: đệ quy bảng lồng
            for row in nested.rows:
                for c in row.cells:
                    self._redact_cell(c)

    # ── docx ──
    def redact_docx(self, in_path: Path, out_path: Path):
        # ── Pass 1: python-docx (paragraphs + bảng lồng + header/footer + core props) ──
        doc = Document(str(in_path))
        for para in doc.paragraphs:
            self._redact_paragraph(para)
        for table in doc.tables:                    # C3: đệ quy nested table
            for row in table.rows:
                for cell in row.cells:
                    self._redact_cell(cell)
        # C2: header/footer (văn bản HC hay để PII ở đây)
        for section in doc.sections:
            for hdr_ftr in (section.header, section.first_page_header, section.even_page_header,
                            section.footer, section.first_page_footer, section.even_page_footer):
                try:
                    for para in hdr_ftr.paragraphs:
                        self._redact_paragraph(para)
                    for tbl in hdr_ftr.tables:
                        for row in tbl.rows:
                            for cell in row.cells:
                                self._redact_cell(cell)
                except Exception:
                    pass
        # C1: xoá metadata (author/comments/title/...) — PII hay nấp trong docProps/core.xml
        cp = doc.core_properties
        for attr in ("author", "comments", "title", "subject", "keywords",
                     "category", "content_status", "last_modified_by"):
            try:
                setattr(cp, attr, "")
            except Exception:
                pass
        tmp = out_path.with_suffix(".tmp.docx")
        doc.save(str(tmp))

        # ── Pass 2: raw-XML defense-in-depth (C4 textbox/drawing + rà lại word/*.xml) ──
        # Che MỌI text node trong word/*.xml bằng cùng Redactor → bắt PII trong textbox,
        # SmartArt, footnotes... mà python-docx không iterate. Reuse token_map → nhất quán.
        self._redact_ooxml_xml(tmp, out_path)
        tmp.unlink(missing_ok=True)

    def _redact_ooxml_xml(self, src_docx: Path, dst_docx: Path):
        """Mở docx (zip), redact text giữa các thẻ XML trong mọi word/*.xml, ghi lại."""
        import zipfile, shutil
        tmp_zip = dst_docx.with_suffix(".tmp2.docx")
        with zipfile.ZipFile(src_docx, "r") as zin, \
             zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        xml = data.decode("utf-8")
                        # chỉ che text giữa '>' và '<' (không đụng attribute/thẻ)
                        xml = re.sub(r">([^<]*)<",
                                     lambda m: ">" + self.redact_text(m.group(1)) + "<",
                                     xml)
                        data = xml.encode("utf-8")
                    except Exception:
                        pass  # giữ nguyên nếu lỗi decode
                zout.writestr(item, data)
        shutil.move(str(tmp_zip), str(dst_docx))

    def redact_txt(self, in_path: Path, out_path: Path):
        text = in_path.read_text(encoding="utf-8")
        out_path.write_text(self.redact_text(text), encoding="utf-8")


# ════════════════════════════════════════════════════════════════════
#  REPORT + RESTORE MAP
# ════════════════════════════════════════════════════════════════════

def write_report(red: Redactor, in_path: Path, out_path: Path, level: int):
    total = len(red.changelog)
    n_hitl = len(red.hitl)
    by_cat = {}
    for e in red.changelog:
        by_cat[e["category"]] = by_cat.get(e["category"], 0) + 1

    lines = [
        f"# Báo cáo redaction — {in_path.name}",
        "",
        f"- Cấp redact: **{level} — {LEVEL_NAMES.get(level, '?')}**",
        f"- File vào: `{in_path.name}`",
        f"- File ra: `{out_path.name}`",
        f"- Tổng số lần che: **{total}**",
        f"- Cờ HITL (cần người duyệt): **{n_hitl}**" + ("  ⚠ Rà lại trước khi đưa AI!" if n_hitl else "  ✓ Không có"),
        "",
        "## Số lần che theo loại",
    ]
    if by_cat:
        lines += [
            "| Loại | Số lần | Độ tin cậy | HITL? |",
            "|------|-------:|-----------:|:-----:|",
        ]
        # map cat → confidence
        conf = {p["id"]: p["confidence"] for p in PATTERNS}
        for cat, n in sorted(by_cat.items(), key=lambda x: -x[1]):
            c = conf.get(cat, "?")
            lines.append(f"| {cat} | {n} | {c} | {'⚠ CÓ' if (isinstance(c,(int,float)) and c<0.7) else '—'} |")
    else:
        lines.append("_Không phát hiện thông tin nhạy cảm ở cấp này._")

    lines += ["", "## ⚠ Các thay đổi CẦN NGƯỜI DUYỆT (HITL)"]
    if red.hitl:
        lines += [
            "Regex che tên/đơn vị/số liệu là **xấp xỉ** — có thể che nhầm hoặc bỏ sót. "
            "Mở file redact, rà từng mục dưới đây:",
            "",
        ]
        seen = set()
        for e in red.hitl:
            key = (e["category"], e["original_preview"])
            if key in seen:
                continue
            seen.add(key)
            lines.append(f"- `{e['token']}` ← **{e['original_preview']}** ({e['category']})")
    else:
        lines.append("_Không có (cấp này chỉ che số định dạng cố định, tin cậy cao)._")

    lines += [
        "",
        "## Phạm vi đã che",
        "- Nội dung thân + **bảng (kể cả bảng lồng)** + **header/footer** + **textbox/drawing**.",
        "- **Metadata đã xoá** (tác giả, người sửa, tiêu đề, ghi chú — những trường dễ lộ PII).",
        "",
        "## Ghi chú",
        "- File này **không gửi dữ liệu đi đâu** — chạy local 100%.",
        "- Bản redact đã sẵn sàng đưa vào AI. Nếu có cờ HITL → rà rồi mới đưa.",
    ]
    return "\n".join(lines) + "\n"


def write_restore_map(red: Redactor, out_path: Path):
    """Ghi map token→gốc để unredact_van_ban.py khôi phục. BẢO MẬT — xoá sau khi xong."""
    data = {
        "warning": "BẢO MẬT — file này chứa ánh xạ token→giá trị gốc. "
                   "XOÁ ngay khi không còn cần khôi phục. Không bao giờ đưa file này vào AI.",
        "generated": datetime.now().isoformat(timespec="seconds"),
        "token_to_original": {tok: orig for orig, tok in red.token_map.items()},
    }
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        out_path.chmod(0o600)  # H6: giới hạn quyền đọc (chỉ chủ sở hữu)
    except OSError:
        pass


# ════════════════════════════════════════════════════════════════════
#  CLI
# ════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description="Redact thông tin nhạy cảm khỏi văn bản .docx/.txt/.md trước khi đưa vào AI.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Cấp: 1=Tối thiểu · 2=Cơ bản (mặc định) · 3=Nâng cao · 4=Triệt để",
    )
    ap.add_argument("input", help="File nhập (.docx hoặc .txt/.md)")
    ap.add_argument("--level", type=int, default=2, choices=[1, 2, 3, 4],
                    help="Cấp redact (1=Tối thiểu · 2=Cơ bản · 3=Nâng cao · 4=Triệt để)")
    ap.add_argument("--outdir", default="output", help="Thư mục ra (mặc định: output/)")
    args = ap.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        sys.stderr.write(f"Lỗi: không tìm thấy file nhập '{in_path}'\n")
        sys.exit(1)

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    red = Redactor(level=args.level)
    stem = in_path.stem
    out_docx = outdir / f"{stem}_redacted.docx"
    out_txt = outdir / f"{stem}_redacted.txt"
    report_md = outdir / f"{stem}_redaction_report.md"
    restore = outdir / f"{stem}_restore_map.json"

    ext = in_path.suffix.lower()
    if ext == ".docx":
        red.redact_docx(in_path, out_docx)
        deliverable = out_docx
    elif ext in (".txt", ".md"):
        red.redact_txt(in_path, out_txt)
        deliverable = out_txt
    else:
        sys.stderr.write(f"Lỗi: định dạng không hỗ trợ '{ext}'. Dùng .docx hoặc .txt/.md\n")
        sys.exit(1)

    report_md.write_text(write_report(red, in_path, deliverable, args.level), encoding="utf-8")
    # Luôn sinh restore_map ở MỌI cấp — unredact_van_ban.py cần nó để khôi phục
    write_restore_map(red, restore)
    # Ngăn vô tình commit/zip restore_map (chứa map ngược token→gốc)
    gi = outdir / ".gitignore"
    if not gi.exists():
        gi.write_text("# KHÔNG commit/đưa AI file restore_map (chứa PII gốc)\n"
                      "*_restore_map.json\n", encoding="utf-8")

    # Tóm tắt ra console (Antigravity hiển thị cho HV)
    print("━" * 56)
    print(f"  ✓ REDACTION xong — Cấp {args.level}: {LEVEL_NAMES[args.level]}")
    print("━" * 56)
    print(f"  File bảo mật : {deliverable}")
    print(f"  Báo cáo      : {report_md}")
    if args.level >= 4:
        print(f"  Restore map  : {restore}  (BẢO MẬT — xoá khi xong)")
    else:
        print(f"  Restore map  : {restore}  (dùng cho unredact; BẢO MẬT — xoá khi xong)")
    print(f"  Đã che       : {len(red.changelog)} mục")
    if red.hitl:
        print(f"  ⚠ Cờ HITL    : {len(red.hitl)} mục CẦN rà lại trước khi đưa AI")
    else:
        print("  ✓ Không có cờ HITL — an toàn đưa vào AI")
    print("━" * 56)


if __name__ == "__main__":
    main()
