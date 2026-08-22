/**
 * =========================================================================
 * DU LỊCH VUI - GOOGLE APPS SCRIPT BACKEND ENGINE (FULL PRODUCTION 3.0)
 * =========================================================================
 * Script tự động quản lý: ChuyenDi, ThanhVien, ChiTieu, ThuQuy.
 * Tự động tạo bảng, tự động sửa lỗi thiếu cột, lưu ảnh Google Drive.
 */

const DRIVE_FOLDER_ID = '1KHNbwt8ULtHU4cPJZSBXgb7XFs3crZxY'; // Thư mục Du Lịch Vui

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initAllSheets(ss);

    const action = e.parameter.action;
    const tripId = e.parameter.tripId;

    if (action === 'getTrips') {
      const tripSheet = ss.getSheetByName('ChuyenDi');
      const data = tripSheet.getDataRange().getValues();
      const trips = [];
      if (data.length > 1) {
        for (let i = 1; i < data.length; i++) {
          trips.push({
            MaChuyenDi: String(data[i][0] || ''),
            TenChuyenDi: String(data[i][1] || ''),
            DiaDiem: String(data[i][2] || ''),
            NgayBatDau: String(data[i][3] || ''),
            NgayKetThuc: String(data[i][4] || ''),
            NganSachDuKien: Number(data[i][5] || 0),
            TienTe: String(data[i][6] || 'VND'),
            TrangThai: String(data[i][7] || 'Đang đi'),
            ThoiGianTao: String(data[i][8] || '')
          });
        }
      }
      return jsonResponse({ trips: trips });
    }

    // Lấy chi tiết 1 chuyến đi
    if (tripId) {
      const tripSheet = ss.getSheetByName('ChuyenDi');
      const memSheet = ss.getSheetByName('ThanhVien');
      const expSheet = ss.getSheetByName('ChiTieu');
      const fundSheet = ss.getSheetByName('ThuQuy');

      let trip = null;
      const tripData = tripSheet.getDataRange().getValues();
      for (let i = 1; i < tripData.length; i++) {
        if (String(tripData[i][0]) === String(tripId)) {
          trip = {
            MaChuyenDi: String(tripData[i][0]),
            TenChuyenDi: String(tripData[i][1] || ''),
            DiaDiem: String(tripData[i][2] || ''),
            NgayBatDau: String(tripData[i][3] || ''),
            NgayKetThuc: String(tripData[i][4] || ''),
            NganSachDuKien: Number(tripData[i][5] || 0),
            TienTe: String(tripData[i][6] || 'VND'),
            TrangThai: String(tripData[i][7] || 'Đang đi'),
            ThoiGianTao: String(tripData[i][8] || '')
          };
          break;
        }
      }

      const members = [];
      const memData = memSheet.getDataRange().getValues();
      for (let i = 1; i < memData.length; i++) {
        if (String(memData[i][1]) === String(tripId) || !memData[i][1]) {
          members.push({
            MaThanhVien: String(memData[i][0] || ''),
            MaChuyenDi: String(memData[i][1] || tripId),
            TenThanhVien: String(memData[i][2] || ''),
            VaiTro: String(memData[i][3] || 'Thành viên'),
            SoTaiKhoan: String(memData[i][4] || ''),
            TenNganHang: String(memData[i][5] || 'MB'),
            TienQuyDaDong: Number(memData[i][6] || 0),
            ThoiGianThamGia: String(memData[i][7] || '')
          });
        }
      }

      const funds = [];
      const fundData = fundSheet.getDataRange().getValues();
      for (let i = 1; i < fundData.length; i++) {
        if (String(fundData[i][1]) === String(tripId) || !fundData[i][1]) {
          funds.push({
            MaGiaoDich: String(fundData[i][0] || ''),
            MaChuyenDi: String(fundData[i][1] || tripId),
            NguoiNop: String(fundData[i][2] || ''),
            payer_name: String(fundData[i][2] || ''),
            SoTien: Number(fundData[i][3] || 0),
            amount: Number(fundData[i][3] || 0),
            NguoiDuocBao: String(fundData[i][4] || ''),
            beneficiaries: String(fundData[i][4] || ''),
            ThoiGian: String(fundData[i][5] || ''),
            created_at: String(fundData[i][5] || '')
          });
        }
      }

      const expenses = [];
      const expData = expSheet.getDataRange().getValues();
      for (let i = 1; i < expData.length; i++) {
        if (String(expData[i][1]) === String(tripId) || !expData[i][1]) {
          expenses.push({
            MaGiaoDich: String(expData[i][0] || ''),
            MaChuyenDi: String(expData[i][1] || tripId),
            MoTa: String(expData[i][2] || ''),
            SoTien: Number(expData[i][3] || 0),
            DanhMuc: String(expData[i][4] || 'Ăn uống'),
            NguonTien: String(expData[i][5] || 'Quỹ chung'),
            NguoiTra: String(expData[i][6] || ''),
            CachChia: String(expData[i][7] || 'Chia đều'),
            NguoiThamGia: String(expData[i][8] || ''),
            LinkHoaDon: String(expData[i][9] || ''),
            NguoiBao: String(expData[i][10] || ''),
            NguoiDuocBao: String(expData[i][11] || ''),
            ThoiGianTao: String(expData[i][12] || '')
          });
        }
      }

      return jsonResponse({ trip: trip, members: members, funds: funds, expenses: expenses });
    }

    return jsonResponse({ status: 'ok', message: 'Du Lịch Vui API is ready' });
  } catch (err) {
    return jsonResponse({ error: true, message: err.toString() });
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initAllSheets(ss);

    let data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }

    const action = data.action;

    // 1. TẠO CHUYẾN ĐI
    if (action === 'createTrip') {
      const tripSheet = ss.getSheetByName('ChuyenDi');
      const memSheet = ss.getSheetByName('ThanhVien');
      const t = data.trip;
      tripSheet.appendRow([
        t.MaChuyenDi, t.TenChuyenDi, t.DiaDiem, t.NgayBatDau, t.NgayKetThuc,
        t.NganSachDuKien, t.TienTe || 'VND', t.TrangThai || 'Đang đi', t.ThoiGianTao || getNowString()
      ]);

      if (data.members && data.members.length) {
        data.members.forEach(m => {
          memSheet.appendRow([
            m.MaThanhVien || ('MEM_' + Date.now()),
            t.MaChuyenDi,
            m.TenThanhVien,
            m.VaiTro || 'Thành viên',
            m.SoTaiKhoan || '',
            m.TenNganHang || 'MB',
            m.TienQuyDaDong || 0,
            m.ThoiGianThamGia || getNowString()
          ]);
        });
      }
      return jsonResponse({ success: true, message: 'Đã tạo chuyến đi' });
    }

    // 2. THÊM THÀNH VIÊN
    if (action === 'addMember') {
      const memSheet = ss.getSheetByName('ThanhVien');
      const m = data.member;
      memSheet.appendRow([
        m.MaThanhVien || ('MEM_' + Date.now()),
        m.MaChuyenDi,
        m.TenThanhVien,
        m.VaiTro || 'Thành viên',
        m.SoTaiKhoan || '',
        m.TenNganHang || 'MB',
        m.TienQuyDaDong || 0,
        m.ThoiGianThamGia || getNowString()
      ]);
      return jsonResponse({ success: true });
    }

    // 3. XÓA THÀNH VIÊN
    if (action === 'deleteMember') {
      const memSheet = ss.getSheetByName('ThanhVien');
      const rows = memSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.MaThanhVien)) {
          memSheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ success: true });
    }

    // 4. THU QUỸ & BAO QUỸ
    if (action === 'addFund') {
      const fundSheet = ss.getSheetByName('ThuQuy');
      const fundId = 'FUND_' + Date.now();
      const tripId = data.tripId || data.MaChuyenDi || '';
      const payer = data.NguoiNop || data.payer || '';
      const amt = Number(data.SoTien || data.amount || 0);
      const bens = data.NguoiDuocBao || data.beneficiaries || payer;
      const time = data.ThoiGian || data.created_at || getNowString();

      fundSheet.appendRow([
        fundId, tripId, payer, amt, bens, time
      ]);
      return jsonResponse({ success: true, fundId: fundId });
    }

    // 5. THÊM CHI TIÊU & LƯU ẢNH DRIVE
    if (action === 'addExpense') {
      const expSheet = ss.getSheetByName('ChiTieu');
      const exp = data.expense;
      let driveUrl = exp.LinkHoaDon || '';

      // Xử lý upload ảnh hóa đơn lên Google Drive
      if (data.imageBase64 && data.imageBase64.startsWith('data:image/')) {
        try {
          const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
          const parts = data.imageBase64.split(',');
          const contentType = parts[0].split(':')[1].split(';')[0];
          const decoded = Utilities.base64Decode(parts[1]);
          const fileName = 'Bill_' + Date.now() + '.jpg';
          const blob = Utilities.newBlob(decoded, contentType, fileName);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          driveUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600';
        } catch (imgErr) {
          driveUrl = exp.LinkHoaDon || '';
        }
      }

      expSheet.appendRow([
        exp.MaGiaoDich || ('EXP_' + Date.now()),
        exp.MaChuyenDi,
        exp.MoTa,
        Number(exp.SoTien || 0),
        exp.DanhMuc || 'Ăn uống',
        exp.NguonTien || 'Quỹ chung',
        exp.NguoiTra || '',
        exp.CachChia || 'Chia đều',
        exp.NguoiThamGia || '',
        driveUrl,
        exp.NguoiBao || '',
        exp.NguoiDuocBao || '',
        exp.ThoiGianTao || getNowString()
      ]);

      return jsonResponse({ success: true, LinkHoaDon: driveUrl });
    }

    // 6. KHÓA SỔ
    if (action === 'lockTrip') {
      const tripSheet = ss.getSheetByName('ChuyenDi');
      const rows = tripSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.MaChuyenDi)) {
          tripSheet.getRange(i + 1, 8).setValue('Đã quyết toán');
          break;
        }
      }
      return jsonResponse({ success: true });
    }

    return jsonResponse({ error: true, message: 'Action not supported: ' + action });
  } catch (err) {
    return jsonResponse({ error: true, message: err.toString() });
  }
}

// Hàm khởi tạo và tự động tạo/sửa các Sheet nếu chưa có
function initAllSheets(ss) {
  getOrCreateSheet(ss, 'ChuyenDi', ['MaChuyenDi', 'TenChuyenDi', 'DiaDiem', 'NgayBatDau', 'NgayKetThuc', 'NganSachDuKien', 'TienTe', 'TrangThai', 'ThoiGianTao']);
  getOrCreateSheet(ss, 'ThanhVien', ['MaThanhVien', 'MaChuyenDi', 'TenThanhVien', 'VaiTro', 'SoTaiKhoan', 'TenNganHang', 'TienQuyDaDong', 'ThoiGianThamGia']);
  getOrCreateSheet(ss, 'ChiTieu', ['MaGiaoDich', 'MaChuyenDi', 'MoTa', 'SoTien', 'DanhMuc', 'NguonTien', 'NguoiTra', 'CachChia', 'NguoiThamGia', 'LinkHoaDon', 'NguoiBao', 'NguoiDuocBao', 'ThoiGianTao']);
  getOrCreateSheet(ss, 'ThuQuy', ['MaGiaoDich', 'MaChuyenDi', 'NguoiNop', 'SoTien', 'NguoiDuocBao', 'ThoiGian']);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getNowString() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())} ${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
}
