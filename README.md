# Redact–Unredact Toolkit

> Che (redact) thông tin nhạy cảm khỏi tài liệu trước khi đưa vào AI — và khôi phục (unredact) khi AI xử lý xong. Chạy local 100%, deterministic bằng regex, không dùng AI để redact.

**Bối cảnh:** quy định cấm cán bộ, công chức đưa dữ liệu mật, tài liệu nhạy cảm lên các công cụ AI. Bộ kit này giúp bạn tuân thủ đúng quy định mà vẫn tận dụng được AI: che dữ liệu NHƯNG VIỆC CHE KHÔNG PHẢI DO AI ĐOÁN — mọi lần che đều bằng regex chạy trên chính máy bạn, script không có bất kỳ kết nối mạng nào.

## Cấu trúc

```
redact-kit/
├── HDSD.docx                    ← Hướng dẫn sử dụng đầy đủ (branding Lộc Đặng)
├── README.md                    ← file này
├── scripts/
│   ├── redact_van_ban.py        ← Che PII: 4 cấp, entity consistency, cờ HITL
│   ├── unredact_van_ban.py      ← Khôi phục: đọc restore_map, thay token về gốc
│   └── test_redact.py           ← Bộ test deterministic (55 checks)
├── samples/
│   └── van-ban-mau-chua-redact.docx   ← Văn bản mẫu chứa PII GIẢ để thử
└── output/                      ← Kết quả sinh ra ở đây
```

## Cài đặt

Yêu cầu: Python 3.10+, một dependency duy nhất:

```bash
pip3 install python-docx
```

Sau khi cài, mọi thứ chạy offline — không cần internet.

## 2 chế độ chạy

### Chế độ A — AI điều khiển (tiện, kém an toàn hơn)

Để một AI assistant (Claude Code, Copilot…) đọc HDSD và tự gõ lệnh script giúp bạn: chọn cấp redact, chạy, mở báo cáo. Tiện cho người không quen dòng lệnh — nhưng lưu ý AI nhìn thấy nội dung file **sau khi đã redact** và chạy trong phiên có ghi log. Chỉ dùng khi bạn đã chắc chắn bản redact sạch (đã rà cờ HITL).

### Chế độ B — Chạy trực tiếp script (an toàn tuyệt đối, khuyên dùng)

Tự gõ lệnh trong terminal. Dữ liệu gốc và restore_map không bao giờ đi qua bất kỳ AI hay tiến trình bên ngoài nào — chỉ bạn và script. Đây là chế độ chuẩn cho tài liệu mật.

## Dùng nhanh

**1. Redact trước khi đưa AI** (khuyên dùng cấp 3):

```bash
python3 scripts/redact_van_ban.py <file.docx|.txt|.md> --level 3 --outdir output
```

Sinh ra: `<tên>_redacted.docx` (dem đi cho AI) + `_redaction_report.md` (đã che gì) + `_restore_map.json` (**BẢO MẬT**).

**2. Unredact sau khi AI xử lý xong:**

```bash
python3 scripts/unredact_van_ban.py output/<tên>_redacted.docx --outdir output
```

Tự tìm restore map → sinh `<tên>_restored.docx` + báo cáo unredact.

## 4 cấp redact

| Cấp | Che gì |
|-----|--------|
| 1 — Tối thiểu | CCCD/CMND, MST, STK, số thẻ, hộ chiếu |
| 2 — Cơ bản (mặc định) | + SĐT, email, mã hồ sơ KT, ngày sinh |
| 3 — Nâng cao (khuyên dùng) | + tên người, đơn vị, số liệu tài chính |
| 4 — Triệt để | + ngày tháng, ký hiệu VB, xoá metadata |

## Quy tắc bảo mật

1. Bản `_restored` chứa lại PII gốc — KHÔNG đưa lại vào AI.
2. Xoá `restore_map.json` khi không còn cần khôi phục.
3. Không bao giờ đưa `restore_map.json` vào AI.

## Kiểm định

Round-trip redact → unredact validated trên văn bản mẫu: 30 che ↔ 30 khôi phục, bản restored == bản gốc, 0 token sót. Bộ test: `cd scripts && python3 test_redact.py` (PASS 55/55).

## Thử nhanh

```bash
python3 scripts/redact_van_ban.py samples/van-ban-mau-chua-redact.docx --level 4 --outdir output
python3 scripts/unredact_van_ban.py output/van-ban-mau-chua-redact_redacted.docx --outdir output
```

## Đóng góp

Bộ kit hiện chạy bằng dòng lệnh — **mời cộng đồng viết GUI** (chọn file → chọn cấp → nút Redact/Unredact) để đồng nghiệp không quen terminal cũng dùng được. Fork, mở issue, gửi PR đều hoan nghênh. Định hướng đóng góp khác: hỗ trợ PDF/XLSX, thêm pattern PII mới (bằng regex, giữ nguyên triết lý deterministic-first).

## License

[MIT](LICENSE) © Lộc Đặng · locdang.com

---
*LỘC ĐẶNG — Công nghệ cho người Việt · locdang.com*
