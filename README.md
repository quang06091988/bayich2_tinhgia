# bayich2_tinhgia — Tính Giá

Máy tính giá bán theo % lãi mong muốn, kèm bảng giá bán đề xuất cho mọi món trong tab **Retail** (sheet bayich2).
Chạm một món để dùng đúng % lãi của món đó và so với Giá Làm Tròn đang bán. Trang chỉ tính thử, **không bao giờ ghi vào Sheet**.

Mở từ trang khác với giá có sẵn: `?gia=216000&sl=20&ten=…` (vd link "Tính giá bán →" ở Đối chiếu toa).

## Dữ liệu

Đọc tab Retail qua Apps Script riêng, **chỉ đọc** — tab Retail không cần xuất bản lên web (giấu giá nhập, % lãi).
Máy chủ tìm cột theo tiêu đề (`Mặt Hàng`, `Giá Nhập Sỉ`, `Số Lượng`; có thì thêm `Đơn Vị Lẻ`, `% Lợi Nhuận`, `Giá Làm Tròn`)
và trả kèm **Bước làm tròn** trong tab `CauHinh` — cùng một chỗ khai với Retail, Đối chiếu toa, Đồng bộ giá.
Trang lưu bản lần gần nhất trong máy: mở là thấy ngay, bản mới tải ở nền.

## Cấu trúc

```
index.html        trang web (Vercel: bayich2-tinhgia.vercel.app)
og.png            ảnh xem trước khi chia sẻ
.clasp.json       Script ID của backend, rootDir = appsscript
appsscript/       backend — Apps Script STANDALONE "bayich2_tinhgia" (không deploy lên Vercel)
  Code.js         doGet?viec=retail — không có lệnh ghi
  appsscript.json
```

- Script: https://script.google.com/d/1QAWF2U_uT-mrMYSUjPaIK8HXfspStxiH-ctvY_WbH2QVSsA5_m7Bc8fO/edit
- Deployment (`API` trong index.html): `AKfycbz2p7bl4Z7pjtor-ClpphmANDtMIpgp6DE024NGWTKBLVijTdbRnT04vpt2i65HkAls`
- Cập nhật backend: `clasp push` → `clasp deploy -i <deploymentId>` (link /exec giữ nguyên).
- Lần đầu: mở script → chọn hàm `caiDat` → Run → cấp quyền đọc Sheet (hàm chỉ đọc thử).
