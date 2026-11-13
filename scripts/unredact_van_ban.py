#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
unredact_van_ban.py — Khôi phục (unredact) văn bản đã che bằng redact_van_ban.py.

Nguyên lý:
  - redact_van_ban.py sinh <tên>_restore_map.json chứa ánh xạ token→giá trị gốc
    (vd "[NGƯỜI_1]" → "Ông Nguyễn Văn A"). Script này đọc map rồi thay MỌI token
    trong file redacted về giá trị gốc — đếm được, deterministic, chạy local 100%.
  - Thứ tự thay: token DÀI trước (tránh "[NGƯỜI_1]" ăn vào "[NGƯỜI_10]").
  - Phạm vi quét giống hệt redactor: thân + bảng (lồng) + header/footer +
    textbox/drawing (pass raw-XML) — không sót ổ PII-token.

Cách chạy:
  python3 scripts/unredact_van_ban.py <file_redacted.docx|.txt|.md> \
      [--restore-map <tên>_restore_map.json] [--outdir <thư_mục_ra>]

  Nếu bỏ --restore-map: tự tìm "<tên_gốc>_restore_map.json" cạnh file redacted
  (bỏ hậu tố "_redacted" rồi ghép "_restore_map.json") và trong ./output/.

Output:
  <thư_mục_ra>/<tên>_restored.docx   ← bản đã khôi phục về giá trị gốc
  <thư_mục_ra>/<tên>_unredact_report.md ← báo cáo: thay bao nhiêu, token nào còn sót

Lưu ý bảo mật:
  - File redacted là bản MANG ĐI AI (đã lộ ra ngoài). Khi khôi phục, bản _restored
    chứa lại PII gốc — giữ an toàn như văn bản gốc, KHÔNG đưa lại vào AI.
  - restore_map.json chứa PII gốc — xoá khi không còn cần khôi phục.
"""

import argparse
import json
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
#  LOAD RESTORE MAP
# ════════════════════════════════════════════════════════════════════

def find_restore_map(redacted_path: Path) -> Path | None:
    """Tự tìm restore map: <base>_restore_map.json cạnh file redacted hoặc trong ./output/."""
    stem = redacted_path.stem
    base = stem[:-len("_redacted")] if stem.endswith("_redacted") else stem
    candidates = [
        redacted_path.parent / f"{base}_restore_map.json",
        redacted_path.parent / f"{stem}_restore_map.json",
        Path("output") / f"{base}_restore_map.json",
        redacted_path.parent / "output" / f"{base}_restore_map.json",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def load_map(map_path: Path) -> dict[str, str]:
    data = json.loads(map_path.read_text(encoding="utf-8"))
    mapping = data.get("token_to_original")
    if not isinstance(mapping, dict) or not mapping:
        sys.stderr.write(f"Lỗi: '{map_path.name}' không có trường token_to_original hợp lệ.\n")
        sys.exit(1)
    # Token DÀI thay TRƯỚC — tránh "[NGƯỜI_1]" khớp phần đầu "[NGƯỜI_10]"
    return dict(sorted(mapping.items(), key=lambda kv: -len(kv[0])))


# ════════════════════════════════════════════════════════════════════
#  ENGINE — thay token về gốc, đếm mỗi lần thay
# ════════════════════════════════════════════════════════════════════

class Unredactor:
    def __init__(self, token_map: dict[str, str]):
        self.token_map = token_map
        self.counts: dict[str, int] = {}

    def unredact_text(self, text: str) -> str:
        if not text:
            return text
        for token, original in self.token_map.items():
            n = text.count(token)
            if n:
                text = text.replace(token, original)
                self.counts[token] = self.counts.get(token, 0) + n
        return text

    # ── paragraph helper (gom runs → unredact → ghi lại run đầu) ──
    def _unredact_paragraph(self, para):
        runs = para.runs
        if not runs:
            return
        full = "".join(r.text for r in runs)
        new = self.unredact_text(full)
        if new != full:
            runs[0].text = new
            for r in runs[1:]:
                r.text = ""

    def _unredact_cell(self, cell):
        for para in cell.paragraphs:
            self._unredact_paragraph(para)
        for nested in cell.tables:                  # bảng lồng
            for row in nested.rows:
                for c in row.cells:
                    self._unredact_cell(c)

    # ── docx ──
    def unredact_docx(self, in_path: Path, out_path: Path):
        # Pass 1: python-docx (paragraphs + bảng lồng + header/footer)
        doc = Document(str(in_path))
        for para in doc.paragraphs:
            self._unredact_paragraph(para)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    self._unredact_cell(cell)
        for section in doc.sections:
            for hdr_ftr in (section.header, section.first_page_header, section.even_page_header,
                            section.footer, section.first_page_footer, section.even_page_footer):
                try:
                    for para in hdr_ftr.paragraphs:
                        self._unredact_paragraph(para)
                    for tbl in hdr_ftr.tables:
                        for row in tbl.rows:
                            for cell in row.cells:
                                self._unredact_cell(cell)
                except Exception:
                    pass
        tmp = out_path.with_suffix(".tmp.docx")
        doc.save(str(tmp))

        # Pass 2: raw-XML (textbox/drawing/footnotes — mọi text node word/*.xml)
        self._unredact_ooxml_xml(tmp, out_path)
        tmp.unlink(missing_ok=True)

    def _unredact_ooxml_xml(self, src_docx: Path, dst_docx: Path):
        """Mở docx (zip), thay token trong text node của mọi word/*.xml, ghi lại."""
        import zipfile, shutil, re
        tmp_zip = dst_docx.with_suffix(".tmp2.docx")
        with zipfile.ZipFile(src_docx, "r") as zin, \
             zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        xml = data.decode("utf-8")
                        xml = re.sub(r">([^<]*)<",
                                     lambda m: ">" + self.unredact_text(m.group(1)) + "<",
                                     xml)
                        data = xml.encode("utf-8")
                    except Exception:
                        pass  # giữ nguyên nếu lỗi decode
                zout.writestr(item, data)
        shutil.move(str(tmp_zip), str(dst_docx))

    def unredact_txt(self, in_path: Path, out_path: Path):
        text = in_path.read_text(encoding="utf-8")
        out_path.write_text(self.unredact_text(text), encoding="utf-8")


# ════════════════════════════════════════════════════════════════════
#  REPORT
# ════════════════════════════════════════════════════════════════════

def write_report(un: Unredactor, in_path: Path, out_path: Path, map_path: Path,
                 leftover_tokens: list[str]) -> str:
    total = sum(un.counts.values())
    unused = [t for t in un.token_map if t not in un.counts]

    lines = [
        f"# Báo cáo unredact — {in_path.name}",
        "",
        f"- File vào: `{in_path.name}` (bản redacted)",
        f"- File ra: `{out_path.name}` (bản khôi phục)",
        f"- Restore map: `{map_path.name}`",
        f"- Số token trong map: **{len(un.token_map)}**",
        f"- Tổng số lần khôi phục: **{total}**",
        "",
        "## Chi tiết theo token",
    ]
    if un.counts:
        lines += ["| Token | Số lần khôi phục |", "|-------|-----------------:|"]
        for tok, n in sorted(un.counts.items(), key=lambda x: -x[1]):
            lines.append(f"| `{tok}` | {n} |")
    else:
        lines.append("_Không thay đổi nào — file vào có thể chưa từng được redact,_"
                     "_ hoặc map không khớp file này._")

    if unused:
        lines += ["", "## Token trong map nhưng không tìm thấy trong file",
                  "(Không sao nếu AI đã rút gọn/diễn giải — chỉ cần biết là có.)"]
        lines += [f"- `{t}`" for t in unused]

    if leftover_tokens:
        lines += ["", "## ⚠ Token VẪN CÒN trong file ra (map thiếu)",
                  "Các token này không có trong restore_map → không khôi phục được.",
                  "Rà thủ công các vị trí này:"]
        lines += [f"- `{t}`" for t in leftover_tokens]

    lines += [
        "",
        "## ⚠ Bảo mật sau khôi phục",
        "- Bản `_restored` chứa lại PII gốc — quản lý như văn bản gốc, KHÔNG đưa vào AI.",
        "- `restore_map.json` chứa PII gốc — xoá khi không còn cần khôi phục.",
    ]
    return "\n".join(lines) + "\n"


def find_leftover_tokens(text: str, token_map: dict[str, str]) -> list[str]:
    """Token dạng [CHỮ_IN_HOA...] còn sót trong text sau khi thay."""
    import re
    found = sorted(set(re.findall(r"\[[A-ZÀ-ỸĐ][A-ZÀ-ỸĐa-zà-ỹ0-9 _\-]*_\d+\]", text)))
    return [t for t in found if t not in token_map]


# ════════════════════════════════════════════════════════════════════
#  CLI
# ════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description="Khôi phục (unredact) văn bản đã che bằng redact_van_ban.py.",
    )
    ap.add_argument("input", help="File redacted (.docx hoặc .txt/.md) — *_redacted.*")
    ap.add_argument("--restore-map", default=None,
                    help="File restore_map.json (mặc định: tự tìm cạnh file vào / trong output/)")
    ap.add_argument("--outdir", default="output", help="Thư mục ra (mặc định: output/)")
    args = ap.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        sys.stderr.write(f"Lỗi: không tìm thấy file nhập '{in_path}'\n")
        sys.exit(1)

    # Resolve restore map
    if args.restore_map:
        map_path = Path(args.restore_map)
    else:
        map_path = find_restore_map(in_path)
    if map_path is None or not map_path.exists():
        sys.stderr.write(
            f"Lỗi: không tìm thấy restore map cho '{in_path.name}'.\n"
            f"  Đặt --restore-map <file>.json — file do redact_van_ban.py sinh ra\n"
            f"  (tên dạng <tên_gốc>_restore_map.json).\n"
        )
        sys.exit(1)

    token_map = load_map(map_path)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    un = Unredactor(token_map)
    stem = in_path.stem
    stem_base = stem[:-len("_redacted")] if stem.endswith("_redacted") else stem
    out_docx = outdir / f"{stem_base}_restored.docx"
    out_txt = outdir / f"{stem_base}_restored.txt"
    report_md = outdir / f"{stem_base}_unredact_report.md"

    ext = in_path.suffix.lower()
    if ext == ".docx":
        un.unredact_docx(in_path, out_docx)
        deliverable = out_docx
        # Rà token sót trên text đã trích (cảnh báo, không chặn)
        from docx import Document as _D
        d = _D(str(out_docx))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                for cell in row.cells:
                    parts.extend(p.text for p in cell.paragraphs)
        leftover = find_leftover_tokens("\n".join(parts), token_map)
    elif ext in (".txt", ".md"):
        un.unredact_txt(in_path, out_txt)
        deliverable = out_txt
        leftover = find_leftover_tokens(out_txt.read_text(encoding="utf-8"), token_map)
    else:
        sys.stderr.write(f"Lỗi: định dạng không hỗ trợ '{ext}'. Dùng .docx hoặc .txt/.md\n")
        sys.exit(1)

    report_md.write_text(write_report(un, in_path, deliverable, map_path, leftover),
                         encoding="utf-8")

    total = sum(un.counts.values())
    print("━" * 56)
    print("  ✓ UNREDACT xong")
    print("━" * 56)
    print(f"  File khôi phục: {deliverable}")
    print(f"  Restore map    : {map_path}")
    print(f"  Đã khôi phục   : {total} lần thay / {len(un.counts)} token")
    if leftover:
        print(f"  ⚠ Token sót    : {len(leftover)} token không có trong map — xem báo cáo")
    print("  ⚠ Bản restored chứa PII gốc — KHÔNG đưa lại vào AI.")
    print("  ⚠ Xoá restore_map.json khi không còn cần khôi phục.")
    print("━" * 56)


if __name__ == "__main__":
    main()
