/**
 * Apps Script STANDALONE "bayich2_tinhgia" — backend CHỈ ĐỌC của trang Tính Giá (bayich2-tinhgia.vercel.app).
 *
 *  ĐỌC : doGet?viec=retail → các cột Retail trang cần + "Bước làm tròn" (tab CauHinh). KHÔNG có lệnh ghi.
 *
 * Thay cho link CSV công khai của tab Retail → tab Retail không cần xuất bản lên web nữa (giấu giá nhập, % lãi).
 * Cột tìm theo CHỮ TIÊU ĐỀ, không theo vị trí: chèn / đổi thứ tự cột Retail không làm tính sai.
 *
 * CÀI / CẤP QUYỀN: chọn hàm caiDat → Run (chỉ đọc thử, không ghi gì).
 * CẬP NHẬT CODE: clasp push → clasp deploy -i <deploymentId> (link /exec giữ nguyên).
 */

var ID_BAYICH2 = '1Wd4Zvq2xiIiEzou_dvE2YtOk-bJJhAD7se0yREYe9c8';   // sheet có tab Retail + CauHinh

var TAB_RETAIL = 'Retail';
var TAB_CAU_HINH = 'CauHinh';
var TD_TRUONG = 'Trường', TD_GIA_TRI = 'Giá trị';
var TRUONG_LAM_TRON = 'Bước làm tròn';

var COT_BAT_BUOC = { ten: 'Mặt Hàng', giaNhapSi: 'Giá Nhập Sỉ', soLuong: 'Số Lượng' };
var COT_PHU = { donViLe: 'Đơn Vị Lẻ', loiNhuan: '% Lợi Nhuận', giaLamTron: 'Giá Làm Tròn' };   // có thì dùng

/* ══════════════════ ĐỌC ══════════════════ */
function doGet(e) {
  try {
    var viec = (e && e.parameter && e.parameter.viec) || 'ping';
    if (viec === 'retail') return traLoi(docRetailTinhGia());
    return traLoi({ ok: true, ten: 'bayich2_tinhgia', thoiGian: new Date().toISOString() });
  } catch (err) {
    return traLoi({ ok: false, loi: String(err) });
  }
}

/* { ok, retail:[{ ten, donViLe, soLuong, loiNhuan (phân số, vd 0.1), giaNhapSi, giaLamTron }], buocLamTron, thoiGian } */
function docRetailTinhGia() {
  var ss = SpreadsheetApp.openById(ID_BAYICH2);
  var sh = ss.getSheetByName(TAB_RETAIL);
  if (!sh) return { ok: false, loi: 'Không tìm thấy tab ' + TAB_RETAIL };
  var hang = sh.getDataRange().getValues();
  if (!hang.length) return { ok: false, loi: 'Tab ' + TAB_RETAIL + ' đang trống' };

  var cot = {}, thieu = [];
  Object.keys(COT_BAT_BUOC).forEach(function (k) {
    cot[k] = timCot(hang[0], COT_BAT_BUOC[k]);
    if (cot[k] < 0) thieu.push(COT_BAT_BUOC[k]);
  });
  if (thieu.length) return { ok: false, loi: 'Tab ' + TAB_RETAIL + ' thiếu cột: ' + thieu.join(', ') + ' — kiểm tra lại tên cột ở dòng tiêu đề' };
  Object.keys(COT_PHU).forEach(function (k) { cot[k] = timCot(hang[0], COT_PHU[k]); });

  var ds = [];
  for (var i = 1; i < hang.length; i++) {
    var d = hang[i], ten = String(d[cot.ten] || '').trim();
    if (!ten) continue;
    var gia = soHoa(d[cot.giaNhapSi]);
    if (gia == null) continue;
    var ln = cot.loiNhuan < 0 ? null : soThuc(d[cot.loiNhuan]);
    ds.push({
      ten: ten,
      donViLe: cot.donViLe < 0 ? '' : String(d[cot.donViLe] || '').trim(),
      soLuong: soThuc(d[cot.soLuong]) || 1,
      loiNhuan: ln == null ? null : (ln > 1 ? ln / 100 : ln),
      giaNhapSi: gia,
      giaLamTron: cot.giaLamTron < 0 ? null : soHoa(d[cot.giaLamTron])
    });
  }
  var buoc = soThuc(layTheoTen(docCauHinhChung(ss), [TRUONG_LAM_TRON]));
  return { ok: true, retail: ds, buocLamTron: buoc > 0 ? buoc : null, thoiGian: new Date().toISOString() };
}

/* ══════════════════ CÀI / CẤP QUYỀN ══════════════════ */
function caiDat() {
  var kq = docRetailTinhGia();
  Logger.log(kq.ok ? ('Đọc thử: ' + kq.retail.length + ' mặt hàng Retail · bước làm tròn ' + kq.buocLamTron) : ('Lỗi: ' + kq.loi));
}

/* ══════════════════ phụ trợ (chép từ bayich2_doichieutoa — chỉ phần đọc) ══════════════════ */
/* { chuanHoa(tên trường): giá trị } của tab CauHinh — chưa có tab thì rỗng */
function docCauHinhChung(ss) {
  var sh = ss.getSheetByName(TAB_CAU_HINH), gt = {};
  if (!sh) return gt;
  var hang = sh.getDataRange().getValues(), td = hang[0] || [];
  var cTr = timCot(td, TD_TRUONG), cGt = timCot(td, TD_GIA_TRI);
  if (cTr < 0 || cGt < 0) return gt;
  for (var i = 1; i < hang.length; i++) { var t = chuanHoa(hang[i][cTr]); if (t && !(t in gt)) gt[t] = hang[i][cGt]; }
  return gt;
}

function layTheoTen(gt, ds) {
  for (var i = 0; i < ds.length; i++) { var k = chuanHoa(ds[i]); if (k in gt) return gt[k]; }
  return undefined;
}

/* Số từ ô Sheet: số giữ nguyên, chữ kiểu "17.000" / "0,5" / "10%" thì bóc ra. Trống → null */
function soThuc(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function soHoa(v) {
  var n = soThuc(v);
  return n == null ? null : Math.round(n);
}

function chuanHoa(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function timCot(tieuDe, ten) {
  var can = chuanHoa(ten);
  for (var i = 0; i < tieuDe.length; i++) if (chuanHoa(tieuDe[i]) === can) return i;
  return -1;
}

function traLoi(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
