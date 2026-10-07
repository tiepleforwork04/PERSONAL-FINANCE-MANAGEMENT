/***********************************************************************
 * QUẢN LÝ TÀI CHÍNH CÁ NHÂN – Google Sheets + Apps Script
 * - Khoản nợ: kỳ tới hạn tiếp theo, số ngày còn lại, ngày hết nợ dự kiến,
 *   trạng thái (Đang trả / Sắp tới hạn / Quá hạn / Đã tất toán)
 * - Nhắc qua email mỗi ngày (trước hạn N ngày, quá hạn thì nhắc hằng ngày)
 * - Thu chi + ngân sách theo danh mục, cảnh báo khi sắp/vượt ngân sách
 * - Báo cáo tổng kết tháng gửi vào ngày 1
 * - Nhận giao dịch ngân hàng tự động qua webhook (SePay, Casso…),
 *   tự phân loại danh mục, tự ghi nhận trả nợ khi nội dung CK chứa mã nợ
 *
 * Cách dùng: dán toàn bộ file vào Tiện ích mở rộng → Apps Script,
 * lưu, tải lại bảng tính, chọn menu "💰 Tài chính" → "1. Thiết lập ban đầu".
 * Nhớ đặt múi giờ project = Asia/Ho_Chi_Minh (Project Settings, hoặc dùng
 * file appsscript.json đi kèm).
 ***********************************************************************/

const TZ = 'Asia/Ho_Chi_Minh';
const SH = {
  DASH: 'Tổng quan',
  DEBTS: 'Khoản nợ',
  PAYMENTS: 'Lịch sử trả nợ',
  TX: 'Thu chi',
  CATS: 'Danh mục',
  SETTINGS: 'Cài đặt',
};
const EXAMPLE_DUE_DAY = 28;
const PAY_TOLERANCE = 0.01; // trả thiếu dưới 1% của 1 kỳ (làm tròn, phí lẻ) vẫn tính là đã trả đủ kỳ đó
// Cột A..Q của sheet "Khoản nợ"
const DEBT_HEADERS = ['Mã', 'Tên khoản nợ', 'Chủ nợ', 'Loại', 'Tổng phải trả', 'Trả mỗi kỳ',
  'Ngày tới hạn (hàng tháng)', 'Ngày bắt đầu', 'Lãi suất %/năm', 'Đã trả', 'Còn lại',
  'Kỳ tới hạn tiếp theo', 'Còn (ngày)', 'Ngày hết nợ dự kiến', 'Trạng thái',
  'Từ khoá nhận diện', 'Ghi chú'];

/* ============================ MENU ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('💰 Tài chính')
    .addItem('1. Thiết lập ban đầu', 'setup')
    .addItem('2. Bật nhắc tự động hằng ngày', 'installTriggers')
    .addSeparator()
    .addItem('Cập nhật trạng thái nợ', 'updateDebts')
    .addItem('Gửi email thử ngay', 'sendTestEmail')
    .addItem('Xem link webhook ngân hàng', 'showWebhookUrl')
    .addItem('Thử nhận 1 giao dịch ngân hàng mẫu', 'testBankTx')
    .addSeparator()
    .addItem('Tắt nhắc tự động', 'removeTriggers')
    .addToUi();
}

// Tự cập nhật khi sửa khoản nợ hoặc thêm lịch sử trả nợ
function onEdit(e) {
  const name = e.range.getSheet().getName();
  if (name === SH.PAYMENTS || name === SH.TX || name === SH.CATS || (name === SH.DEBTS && e.range.getColumn() <= 9)) {
    try { updateDebts(); } catch (err) { /* bỏ qua */ }
  }
}

/* ======================= THIẾT LẬP BAN ĐẦU ======================= */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  const list = arr => SpreadsheetApp.newDataValidation().requireValueInList(arr, true).setAllowInvalid(true).build();
  const inRange = rg => SpreadsheetApp.newDataValidation().requireValueInRange(rg, true).setAllowInvalid(true).build();
  const textRule = (rg, text, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setRanges([rg]).build();

  // ---- Cài đặt
  const st = sheet_(SH.SETTINGS);
  if (st.getLastRow() === 0) {
    st.getRange('B2:B7').setNumberFormat('@');
    st.getRange(1, 1, 7, 3).setValues([
      ['Thiết lập', 'Giá trị', 'Giải thích'],
      ['Email nhận thông báo', Session.getActiveUser().getEmail(), 'Nhiều email thì cách nhau dấu phẩy'],
      ['Nhắc trước (ngày)', '7,3,1,0', 'Gửi nhắc khi còn đúng số ngày này là tới hạn. Quá hạn thì nhắc mỗi ngày'],
      ['Giờ gửi email', '8', '0–23 (giờ VN). Đổi xong chạy lại "Bật nhắc tự động"'],
      ['Cảnh báo ngân sách (%)', '80', 'Báo khi chi một danh mục chạm % này của ngân sách tháng (và khi vượt 100%)'],
      ['Gửi báo cáo tháng', 'Có', 'Có/Không – tổng kết tháng trước, gửi vào ngày 1'],
      ['Webhook token', Utilities.getUuid().replace(/-/g, ''), 'Mã bí mật cho webhook ngân hàng. KHÔNG chia sẻ'],
    ]);
    header_(st, 3);
    st.setColumnWidth(1, 200).setColumnWidth(2, 280).setColumnWidth(3, 480);
  }

  // ---- Danh mục
  const cat = sheet_(SH.CATS);
  if (cat.getLastRow() === 0) {
    const cats = [
      ['Danh mục', 'Loại', 'Ngân sách tháng', 'Từ khoá tự phân loại (cách nhau dấu phẩy)'],
      ['Ăn uống', 'Chi', 4000000, 'grabfood, grab food, shopeefood, baemin, highlands, phuc long, cafe, ca phe, an trua, an toi'],
      ['Đi lại', 'Chi', 1500000, 'grab, xanh sm, gojek, be group, petrolimex, xang, vetc, gui xe'],
      ['Nhà ở & hoá đơn', 'Chi', 6000000, 'tien nha, tien dien, evn, tien nuoc, internet, fpt telecom, viettel, vnpt'],
      ['Mua sắm', 'Chi', 2000000, 'shopee, lazada, tiki, tiktok shop, winmart, coopmart, bach hoa xanh'],
      ['Sức khoẻ', 'Chi', 1000000, 'pharmacity, long chau, nha thuoc, benh vien, phong kham'],
      ['Giải trí', 'Chi', 1000000, 'cgv, lotte cinema, netflix, spotify, youtube'],
      ['Giáo dục', 'Chi', '', 'hoc phi, khoa hoc'],
      ['Trả nợ', 'Chi', '', 'tra no, thanh toan the, tra gop'],
      ['Khác', 'Chi', '', ''],
      ['Lương', 'Thu', '', 'luong, salary'],
      ['Thu khác', 'Thu', '', ''],
    ];
    cat.getRange(1, 1, cats.length, 4).setValues(cats);
    cat.getRange('E1:F1').setValues([['Đã chi/thu tháng này', '% ngân sách']]); // E, F do script tự tính
    header_(cat, 6);
    cat.getRange('B2:B').setDataValidation(list(['Chi', 'Thu']));
    cat.getRange('C2:C').setNumberFormat('#,##0');
    cat.getRange('E2:E').setNumberFormat('#,##0');
    cat.getRange('F2:F').setNumberFormat('0%');
    const f = cat.getRange('F2:F');
    cat.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setBackground('#f8d7da').setFontColor('#842029').setRanges([f]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8).setBackground('#fff3cd').setRanges([f]).build(),
    ]);
    cat.setColumnWidth(1, 160).setColumnWidth(4, 520).setColumnWidth(5, 160);
  }

  // ---- Khoản nợ
  const debt = sheet_(SH.DEBTS);
  if (debt.getLastRow() === 0) {
    debt.getRange(1, 1, 1, DEBT_HEADERS.length).setValues([DEBT_HEADERS]);
    const today = startOfDay_(new Date());
    const start = new Date(today.getFullYear(), today.getMonth() - 1, 1); // đầu tháng trước
    debt.getRange(2, 1, 1, 9).setValues([['NO01', 'Vay mua xe (VÍ DỤ – xoá dòng này)', 'Ngân hàng ABC',
      'Vay trả góp', 60000000, 5500000, EXAMPLE_DUE_DAY, start, 9.5]]);
    debt.getRange('P2').setValue('tra gop xe, no01');
    // Cột J (Đã trả) và K (Còn lại) do script tự tính trong updateDebts()
    header_(debt, DEBT_HEADERS.length);
    debt.getRange('D2:D').setDataValidation(list(['Vay trả góp', 'Thẻ tín dụng', 'Vay cá nhân', 'Mua trả góp', 'Định kỳ (không có tổng)', 'Khác']));
    debt.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 31).setAllowInvalid(false)
      .setHelpText('Ngày trong tháng (1–31). Tháng ngắn hơn sẽ lấy ngày cuối tháng.').build());
    ['E2:E', 'F2:F', 'J2:J', 'K2:K'].forEach(a => debt.getRange(a).setNumberFormat('#,##0'));
    ['H2:H', 'L2:L', 'N2:N'].forEach(a => debt.getRange(a).setNumberFormat('dd/mm/yyyy'));
    debt.getRange('L1:O').setBackground('#eef3f8'); // cột do script tự tính
    header_(debt, DEBT_HEADERS.length);
    const o = debt.getRange('O2:O');
    debt.setConditionalFormatRules([
      textRule(o, 'Quá hạn', '#f8d7da', '#842029'),
      textRule(o, 'Sắp tới hạn', '#fff3cd', '#664d03'),
      textRule(o, 'Đã tất toán', '#d1e7dd', '#0f5132'),
    ]);
    debt.setColumnWidths(1, DEBT_HEADERS.length, 125);
    debt.setColumnWidth(2, 240).setColumnWidth(16, 200).setColumnWidth(17, 220);
  }

  // ---- Lịch sử trả nợ
  const pay = sheet_(SH.PAYMENTS);
  if (pay.getLastRow() === 0) {
    pay.getRange(1, 1, 1, 6).setValues([['Ngày', 'Mã khoản nợ', 'Số tiền', 'Nguồn', 'Mã giao dịch', 'Ghi chú']]);
    const d0 = debt.getRange('H2').getValue();
    if (d0 instanceof Date) pay.appendRow([nthDueDate_(startOfDay_(d0), EXAMPLE_DUE_DAY, 1), 'NO01', 5500000, 'Thủ công', '', 'Ví dụ – xoá dòng này']);
    header_(pay, 6);
    pay.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
    pay.getRange('C2:C').setNumberFormat('#,##0');
    pay.getRange('E2:E').setNumberFormat('@');
    pay.getRange('B2:B').setDataValidation(inRange(debt.getRange('A2:A')));
    pay.getRange('D2:D').setDataValidation(list(['Thủ công', 'Ngân hàng (tự động)']));
    pay.setColumnWidth(6, 320);
  }

  // ---- Thu chi
  const tx = sheet_(SH.TX);
  if (tx.getLastRow() === 0) {
    tx.getRange(1, 1, 1, 8).setValues([['Ngày', 'Loại', 'Danh mục', 'Số tiền', 'Mô tả', 'Tài khoản', 'Nguồn', 'Mã giao dịch']]);
    header_(tx, 8);
    tx.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
    tx.getRange('D2:D').setNumberFormat('#,##0');
    tx.getRange('H2:H').setNumberFormat('@');
    tx.getRange('B2:B').setDataValidation(list(['Chi', 'Thu']));
    tx.getRange('C2:C').setDataValidation(inRange(cat.getRange('A2:A')));
    tx.setColumnWidth(5, 340);
  }

  // ---- Tổng quan
  const dash = sheet_(SH.DASH);
  if (dash.getLastRow() === 0) {
    dash.getRange('A1').setValue('TỔNG QUAN TÀI CHÍNH').setFontSize(16).setFontWeight('bold');
    dash.getRange(3, 1, 7, 1).setValues([['Tổng nợ còn lại'], ['Phải trả nợ mỗi tháng'], ['Khoản nợ quá hạn'],
      ['Khoản nợ sắp tới hạn'], ['Thu tháng này'], ['Chi tháng này'], ['Còn lại tháng này']]); // số liệu do script tự tính
    dash.getRange('A3:A9').setFontWeight('bold');
    dash.getRange('B3:B9').setNumberFormat('#,##0');
    dash.getRange('B5:B6').setNumberFormat('0');
    dash.getRange('A11').setValue('Lịch trả nợ sắp tới').setFontWeight('bold').setFontSize(13);
    dash.getRange('C13:C').setNumberFormat('#,##0');
    dash.getRange('D13:D').setNumberFormat('dd/mm/yyyy');
    dash.setColumnWidth(1, 200).setColumnWidth(2, 240).setColumnWidths(3, 4, 140);
  }

  // Sắp xếp sheet, xoá sheet trống mặc định
  [SH.DASH, SH.DEBTS, SH.PAYMENTS, SH.TX, SH.CATS, SH.SETTINGS].forEach((n, i) => {
    ss.setActiveSheet(ss.getSheetByName(n));
    ss.moveActiveSheet(i + 1);
  });
  ss.getSheets().forEach(s => {
    if (!Object.values(SH).includes(s.getName()) && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  updateDebts();
  ss.setActiveSheet(ss.getSheetByName(SH.DASH));
  ss.toast('Thiết lập xong! Kiểm tra email ở sheet "Cài đặt", rồi chọn "2. Bật nhắc tự động hằng ngày".', '💰 Tài chính', 8);
}

/* ======================= TÍNH TOÁN KHOẢN NỢ ======================= */

/**
 * Logic: số kỳ đã trả = floor(tổng đã trả / tiền mỗi kỳ).
 * Kỳ tới hạn tiếp theo = kỳ thứ (số kỳ đã trả + 1) tính từ ngày bắt đầu.
 * Nếu ngày đó đã qua → Quá hạn. Trả một phần thì kỳ này chỉ còn phải trả phần thiếu.
 */
function updateDebts() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.DEBTS);
  if (!sh || sh.getLastRow() < 2) { refreshReports_([]); return []; }
  const remind = parseDays_(getSettings_()['Nhắc trước (ngày)']);
  const maxRemind = remind.length ? Math.max(...remind) : 3;
  const paidMap = paymentsById_();
  const today = startOfDay_(new Date());
  const n = sh.getLastRow() - 1;
  const rows = sh.getRange(2, 1, n, DEBT_HEADERS.length).getValues();
  const out = [], debts = [];

  rows.forEach(r => {
    const id = String(r[0]).trim();
    if (!id) { out.push(['', '', '', '', '', '']); return; }
    const total = Number(r[4]) || 0;
    const inst = Number(r[5]) || 0;
    const dueDay = Math.min(Number(r[6]) || 0, 31);
    const start = r[7] instanceof Date ? startOfDay_(r[7]) : today;
    const paid = paidMap[id] || 0;
    const remain = total > 0 ? Math.max(total - paid, 0) : null; // null = khoản định kỳ, không có tổng
    const d = { id, name: r[1], creditor: r[2], inst, paid, remain,
      nextDue: null, daysLeft: null, amountDue: 0, payoff: null, overdue: 0, status: '' };

    if (remain !== null && (remain === 0 || (inst > 0 && remain <= inst * PAY_TOLERANCE))) d.status = 'Đã tất toán';
    else if (!inst || !dueDay) d.status = 'Thiếu thông tin';
    else {
      const paidPeriods = Math.floor(paid / inst + PAY_TOLERANCE);
      d.nextDue = nthDueDate_(start, dueDay, paidPeriods + 1);
      d.daysLeft = Math.round((d.nextDue - today) / 86400000);
      d.amountDue = inst - Math.max(paid - paidPeriods * inst, 0);
      if (remain !== null) {
        d.amountDue = Math.min(d.amountDue, remain);
        d.payoff = nthDueDate_(start, dueDay, paidPeriods + Math.ceil(remain / inst - PAY_TOLERANCE));
      }
      let k = paidPeriods + 1;
      while (d.overdue < 240 && nthDueDate_(start, dueDay, k) < today) { d.overdue++; k++; }
      d.status = d.daysLeft < 0 ? 'Quá hạn' : d.daysLeft <= maxRemind ? 'Sắp tới hạn' : 'Đang trả';
    }
    out.push([d.paid, d.remain === null ? '' : d.remain, d.nextDue || '', d.daysLeft === null ? '' : d.daysLeft, d.payoff || '', d.status]);
    debts.push(d);
  });

  sh.getRange(1, 10, 1, 2).setValues([['Đã trả', 'Còn lại']]); // thay công thức cũ (nếu có) bằng chữ
  sh.getRange(2, 10, n, 6).setValues(out);
  refreshReports_(debts);
  return debts;
}

// Điền số liệu cho sheet "Tổng quan" và "Danh mục" bằng script (không dùng công thức,
// vì công thức dễ lỗi #ERROR! khi bảng tính đặt ngôn ngữ/vùng không phải US)
function refreshReports_(debts) {
  const ss = SpreadsheetApp.getActive();
  const now = new Date();
  const t = sumTx_(new Date(now.getFullYear(), now.getMonth(), 1), new Date(now.getFullYear(), now.getMonth() + 1, 1));

  const dash = ss.getSheetByName(SH.DASH);
  if (dash) {
    const open = debts.filter(d => d.status !== 'Đã tất toán');
    dash.getRange('B3:B9').setValues([
      [debts.reduce((sum, d) => sum + (d.remain || 0), 0)],
      [open.reduce((sum, d) => sum + d.inst, 0)],
      [debts.filter(d => d.status === 'Quá hạn').length],
      [debts.filter(d => d.status === 'Sắp tới hạn').length],
      [t.thu], [t.chi], [t.thu - t.chi],
    ]);
    dash.getRange('A12:F60').clearContent();
    const upcoming = open.filter(d => d.nextDue).sort((a, b) => a.nextDue - b.nextDue)
      .map(d => [d.id, d.name, d.inst, d.nextDue, d.daysLeft, d.status]);
    dash.getRange(12, 1, 1, 6).setValues([['Mã', 'Tên khoản nợ', 'Trả mỗi kỳ', 'Kỳ tới hạn', 'Còn (ngày)', 'Trạng thái']])
      .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
    if (upcoming.length) dash.getRange(13, 1, upcoming.length, 6).setValues(upcoming);
  }

  const cat = ss.getSheetByName(SH.CATS);
  if (cat && cat.getLastRow() > 1) {
    const n = cat.getLastRow() - 1;
    const rows = cat.getRange(2, 1, n, 3).getValues();
    cat.getRange(1, 5, 1, 2).setValues([['Đã chi/thu tháng này', '% ngân sách']]);
    cat.getRange(2, 5, n, 2).setValues(rows.map(([name, , budget]) => {
      if (!name) return ['', ''];
      const used = t.all[name] || 0;
      return [used, Number(budget) > 0 ? used / Number(budget) : ''];
    }));
  }
}

// Ngày tới hạn thứ n (n = 1 là kỳ đầu tiên kể từ ngày bắt đầu)
function nthDueDate_(start, dueDay, n) {
  const y = start.getFullYear();
  let m = start.getMonth();
  if (clampDate_(y, m, dueDay) < start) m += 1;
  return clampDate_(y, m + n - 1, dueDay);
}

// Ngày dueDay của tháng m; tháng ngắn hơn thì lấy ngày cuối tháng
function clampDate_(y, m, d) {
  const last = new Date(y, m + 1, 0).getDate();
  return new Date(y, m, Math.min(d, last));
}

function paymentsById_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.PAYMENTS);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 2, sh.getLastRow() - 1, 2).getValues().forEach(([id, amt]) => {
    id = String(id).trim();
    if (id) map[id] = (map[id] || 0) + (Number(amt) || 0);
  });
  return map;
}

/* ======================= EMAIL HẰNG NGÀY ======================= */

function dailyJob() {
  const s = getSettings_();
  const to = s['Email nhận thông báo'];
  if (!to) return;
  const remind = parseDays_(s['Nhắc trước (ngày)']);
  const debts = updateDebts();
  const alerts = debts.filter(d => d.daysLeft !== null && (d.daysLeft < 0 || remind.includes(d.daysLeft)));
  const budget = budgetAlerts_(Number(s['Cảnh báo ngân sách (%)']) || 80);

  const sections = [], subject = [];
  if (alerts.length) {
    sections.push(debtSection_(alerts));
    const late = alerts.filter(d => d.daysLeft < 0).length;
    subject.push(late ? `${late} khoản QUÁ HẠN` : `${alerts.length} khoản nợ sắp tới hạn`);
  }
  if (budget.list.length) {
    sections.push(budgetSection_(budget.list));
    subject.push('cảnh báo ngân sách');
  }
  const now = new Date();
  if (Utilities.formatDate(now, TZ, 'd') === '1' && /^c/i.test(s['Gửi báo cáo tháng'] || '')) {
    const from = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const toDate = new Date(now.getFullYear(), now.getMonth(), 1);
    sections.push(summarySection_(`📊 Tổng kết tháng ${Utilities.formatDate(from, TZ, 'MM/yyyy')}`, from, toDate, debts));
    subject.push('báo cáo tháng');
  }
  if (!sections.length) return;

  MailApp.sendEmail({
    to,
    subject: '💰 [Tài chính] ' + subject.join(' · '),
    htmlBody: wrapEmail_(sections.join('')),
  });
  // Chỉ đánh dấu "đã cảnh báo" sau khi email gửi thành công
  const props = PropertiesService.getScriptProperties();
  budget.keys.forEach(k => props.setProperty(k, '1'));
}

function sendTestEmail() {
  const s = getSettings_();
  const ui = SpreadsheetApp.getUi();
  if (!s['Email nhận thông báo']) return ui.alert('Chưa có email trong sheet "Cài đặt".');
  const debts = updateDebts();
  const active = debts.filter(d => d.nextDue);
  const now = new Date();
  const html = (active.length ? debtSection_(active, '📅 Tất cả khoản nợ đang trả') : '<p>Không có khoản nợ đang trả.</p>')
    + summarySection_('📊 Tháng này (tính đến hôm nay)', new Date(now.getFullYear(), now.getMonth(), 1),
      new Date(now.getFullYear(), now.getMonth() + 1, 1), debts);
  MailApp.sendEmail({ to: s['Email nhận thông báo'], subject: '💰 [Tài chính] Email thử', htmlBody: wrapEmail_(html) });
  SpreadsheetApp.getActive().toast('Đã gửi email thử tới ' + s['Email nhận thông báo']);
}

// Mỗi mức cảnh báo (ngưỡng % và 100%) chỉ gửi 1 lần/tháng/danh mục.
// Trả về { list, keys }: dailyJob ghi `keys` vào Script Properties sau khi gửi email xong.
function budgetAlerts_(threshold) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.CATS);
  const res = { list: [], keys: [] };
  if (!sh || sh.getLastRow() < 2) return res;
  const now = new Date();
  const spent = sumTx_(new Date(now.getFullYear(), now.getMonth(), 1), new Date(now.getFullYear(), now.getMonth() + 1, 1)).byCat;
  const props = PropertiesService.getScriptProperties();
  const month = Utilities.formatDate(now, TZ, 'yyyyMM');

  // Dọn key của các tháng cũ
  const all = props.getProperties();
  Object.keys(all).forEach(k => {
    const m = /^bud_(\d{6})_/.exec(k);
    if (m && m[1] < month) props.deleteProperty(k);
  });

  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(([name, type, budget]) => {
    budget = Number(budget) || 0;
    if (!name || type !== 'Chi' || !budget) return;
    const used = spent[name] || 0;
    const pct = used / budget * 100;
    const level = pct >= 100 ? 100 : pct >= threshold ? threshold : 0;
    if (!level) return;
    const key = `bud_${month}_${name}_${level}`;
    if (all[key]) return;
    res.keys.push(key);
    res.list.push({ name, budget, used, pct });
  });
  return res;
}

// Tổng thu/chi trong khoảng [from, to)
function sumTx_(from, to) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.TX);
  const res = { thu: 0, chi: 0, byCat: {}, all: {} };
  if (!sh || sh.getLastRow() < 2) return res;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(([date, type, cat, amt]) => {
    if (!(date instanceof Date) || date < from || date >= to) return;
    amt = Number(amt) || 0;
    if (cat) res.all[cat] = (res.all[cat] || 0) + amt;
    if (type === 'Thu') res.thu += amt;
    else if (type === 'Chi') {
      res.chi += amt;
      const c = cat || 'Chưa phân loại';
      res.byCat[c] = (res.byCat[c] || 0) + amt;
    }
  });
  return res;
}

function sumPaymentsInRange_(from, to) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.PAYMENTS);
  if (!sh || sh.getLastRow() < 2) return 0;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(([d]) => d instanceof Date && d >= from && d < to)
    .reduce((s, r) => s + (Number(r[2]) || 0), 0);
}

/* ======================= NỘI DUNG EMAIL ======================= */

function debtSection_(list, title) {
  const rows = list.slice().sort((a, b) => a.daysLeft - b.daysLeft).map(d => {
    const late = d.daysLeft < 0;
    const when = late ? `Quá hạn ${-d.daysLeft} ngày` + (d.overdue > 1 ? ` (${d.overdue} kỳ)` : '')
      : d.daysLeft === 0 ? 'Hôm nay' : `Còn ${d.daysLeft} ngày`;
    const color = late ? '#c0392b' : d.daysLeft <= 1 ? '#d35400' : '#333';
    return `<tr style="border-bottom:1px solid #eee">`
      + `<td>${esc_(d.name || d.id)}<br><small style="color:#888">${esc_(d.creditor)} · ${esc_(d.id)}</small></td>`
      + `<td>${fmtDate_(d.nextDue)}</td>`
      + `<td style="color:${color};font-weight:bold">${when}</td>`
      + `<td style="text-align:right">${money_(d.amountDue)}</td>`
      + `<td style="text-align:right">${d.remain === null ? '—' : money_(d.remain)}<br><small style="color:#888">${d.payoff ? 'hết nợ ' + fmtDate_(d.payoff) : ''}</small></td></tr>`;
  }).join('');
  return section_(title || '📅 Khoản nợ cần thanh toán', table_(['Khoản nợ', 'Hạn', 'Tình trạng', 'Kỳ này', 'Còn nợ'], rows));
}

function budgetSection_(list) {
  const rows = list.map(b => `<tr style="border-bottom:1px solid #eee"><td>${esc_(b.name)}</td>`
    + `<td style="text-align:right">${money_(b.used)}</td><td style="text-align:right">${money_(b.budget)}</td>`
    + `<td style="text-align:right;font-weight:bold;color:${b.pct >= 100 ? '#c0392b' : '#d35400'}">${Math.round(b.pct)}%</td></tr>`).join('');
  return section_('⚠️ Cảnh báo ngân sách tháng này', table_(['Danh mục', 'Đã chi', 'Ngân sách', 'Tỷ lệ'], rows));
}

function summarySection_(title, from, to, debts) {
  const t = sumTx_(from, to);
  const paidDebt = sumPaymentsInRange_(from, to);
  const totalRemain = debts.reduce((s, d) => s + (d.remain || 0), 0);
  const cats = Object.entries(t.byCat).sort((a, b) => b[1] - a[1]).slice(0, 8)
    .map(([c, v]) => `<tr style="border-bottom:1px solid #eee"><td>${esc_(c)}</td><td style="text-align:right">${money_(v)}</td>`
      + `<td style="text-align:right">${t.chi ? Math.round(v / t.chi * 100) : 0}%</td></tr>`).join('');
  return section_(title,
    `<p style="line-height:1.7">Thu: <b>${money_(t.thu)}</b> · Chi: <b>${money_(t.chi)}</b> · Chênh lệch: <b>${money_(t.thu - t.chi)}</b><br>`
    + `Đã trả nợ: <b>${money_(paidDebt)}</b> · Tổng nợ còn lại: <b>${money_(totalRemain)}</b></p>`
    + (cats ? table_(['Danh mục chi', 'Số tiền', 'Tỷ lệ'], cats) : ''));
}

function section_(title, body) {
  return `<h3 style="margin:24px 0 8px;color:#1f4e78">${title}</h3>${body}`;
}

function table_(headers, rowsHtml) {
  return `<table cellpadding="8" style="border-collapse:collapse;width:100%;font-size:14px">`
    + `<tr style="background:#1f4e78;color:#fff">${headers.map(h => `<th align="left">${h}</th>`).join('')}</tr>`
    + rowsHtml + '</table>';
}

function wrapEmail_(inner) {
  const ss = SpreadsheetApp.getActive();
  return `<div style="font-family:Arial,sans-serif;max-width:680px;color:#333">${inner}`
    + `<p style="margin-top:28px;font-size:12px;color:#888">Mở bảng tính: <a href="${ss.getUrl()}">${esc_(ss.getName())}</a></p></div>`;
}

/* ======================= TRIGGER ======================= */

function installTriggers() {
  removeTriggers(true);
  let hour = parseInt(getSettings_()['Giờ gửi email'], 10);
  if (isNaN(hour) || hour < 0 || hour > 23) hour = 8;
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(hour).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert(`Đã bật nhắc tự động: mỗi ngày trong khoảng ${hour}h–${hour + 1}h sẽ kiểm tra và gửi email nếu có khoản cần nhắc.`);
}

function removeTriggers(silent) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyJob')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (silent !== true) SpreadsheetApp.getActive().toast('Đã tắt nhắc tự động.');
}

/* ======================= WEBHOOK NGÂN HÀNG ======================= */
// Triển khai thành Web App rồi dán link (kèm ?token=...) vào SePay / Casso.
// Lưu ý: Web App trả về qua redirect 302; nếu dịch vụ báo lỗi, dữ liệu thường vẫn đã được ghi
// (giao dịch có mã được chống ghi trùng khi gửi lại).

function doGet() {
  return ContentService.createTextOutput('OK – webhook tài chính đang hoạt động');
}

function doPost(e) {
  try {
    const token = (e.parameter && e.parameter.token) || '';
    if (!token || token !== String(getSettings_()['Webhook token'])) return json_({ success: false, error: 'unauthorized' });
    const body = JSON.parse(e.postData.contents);
    // Hỗ trợ: 1 object (SePay), {data: [...]} hoặc {data: {...}} (Casso), hoặc mảng
    const list = Array.isArray(body) ? body
      : Array.isArray(body.data) ? body.data
      : body.data && typeof body.data === 'object' ? [body.data] : [body];
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let added = 0;
    try {
      list.forEach(t => { if (recordBankTx_(normalizeTx_(t))) added++; });
    } finally {
      lock.releaseLock();
    }
    return json_({ success: true, added });
  } catch (err) {
    return json_({ success: false, error: String(err) });
  }
}

// Chuẩn hoá dữ liệu từ nhiều dịch vụ về 1 dạng chung
function normalizeTx_(t) {
  let amount, type;
  if (t.transferType) { // SePay
    amount = Number(t.transferAmount) || 0;
    type = String(t.transferType).toLowerCase() === 'in' ? 'Thu' : 'Chi';
  } else { // Casso & dạng chung: số âm là tiền ra
    const a = Number(t.amount) || 0;
    amount = Math.abs(a);
    type = a >= 0 ? 'Thu' : 'Chi';
  }
  return {
    amount, type,
    desc: String(t.content || t.description || ''),
    date: parseDate_(t.transactionDate || t.transactionDateTime || t.when),
    ref: String(t.id || t.tid || t.referenceCode || t.reference || ''),
    account: String(t.accountNumber || t.subAccId || t.bank_sub_acc_id || ''),
    bank: String(t.gateway || t.bankName || ''),
  };
}

function recordBankTx_(tx) {
  if (!tx.amount) return false;
  const ss = SpreadsheetApp.getActive();
  const txSh = ss.getSheetByName(SH.TX);
  // Chống ghi trùng khi dịch vụ gửi lại
  if (tx.ref && txSh.getRange('H:H').createTextFinder(tx.ref).matchEntireCell(true).findNext()) return false;

  const text = norm_(tx.desc);
  const debtId = tx.type === 'Chi' ? matchDebt_(text) : null;
  const category = debtId ? 'Trả nợ' : matchCategory_(text, tx.type);

  // safe_: nội dung từ bên ngoài không được hiểu thành công thức
  txSh.appendRow([tx.date, tx.type, category, tx.amount, safe_(tx.desc), safe_(tx.account),
    safe_('Ngân hàng' + (tx.bank ? ' ' + tx.bank : '')), tx.ref]);
  if (debtId) {
    ss.getSheetByName(SH.PAYMENTS).appendRow([tx.date, debtId, tx.amount, 'Ngân hàng (tự động)', tx.ref, safe_(tx.desc)]);
  }
  updateDebts(); // cập nhật nợ + tổng quan + danh mục
  return true;
}

// Khớp khoản nợ: nội dung CK chứa mã nợ (vd NO01) hoặc 1 từ khoá ở cột "Từ khoá nhận diện".
// Mã nợ còn được khớp khi ngân hàng bỏ khoảng trắng (vd "TRAGOPNO01").
function matchDebt_(text) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.DEBTS);
  if (!sh || sh.getLastRow() < 2) return null;
  const compact = text.replace(/ /g, '');
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 16).getValues();
  for (const r of rows) {
    const id = String(r[0]).trim();
    if (!id || r[14] === 'Đã tất toán') continue;
    const kws = [id].concat(String(r[15] || '').split(','));
    if (kws.some(k => hasKw_(text, k))) return id;
    const idc = norm_(id).replace(/ /g, '');
    if (idc.length >= 3 && compact.includes(idc)) return id;
  }
  return null;
}

function matchCategory_(text, type) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.CATS);
  if (sh && sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (const [name, t, , kw] of rows) {
      if (!name || t !== type) continue;
      if (String(kw || '').split(',').some(k => hasKw_(text, k))) return name;
    }
  }
  return type === 'Thu' ? 'Thu khác' : 'Khác';
}

function testBankTx() {
  const ok = recordBankTx_(normalizeTx_({
    id: 'TEST' + Date.now(), gateway: 'Vietcombank', accountNumber: '0123456789',
    transactionDate: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    content: 'Thanh toan GRABFOOD don hang 123', transferType: 'out', transferAmount: 85000,
  }));
  SpreadsheetApp.getActive().toast(ok ? 'Đã thêm 1 giao dịch mẫu vào sheet "Thu chi" (danh mục Ăn uống).' : 'Không thêm được.');
}

function showWebhookUrl() {
  const ui = SpreadsheetApp.getUi();
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (err) { /* chưa triển khai */ }
  if (!url) return ui.alert('Chưa triển khai Web App.\n\nVào Apps Script → Triển khai → Triển khai mới → Ứng dụng web\n(Thực thi với tư cách: Tôi · Ai có quyền truy cập: Bất kỳ ai), rồi chạy lại mục này.');
  ui.alert('Link webhook', `${url}?token=${getSettings_()['Webhook token']}\n\nDán link này vào phần cấu hình Webhook của SePay / Casso. Không chia sẻ cho người khác.`, ui.ButtonSet.OK);
}

/* ======================= TIỆN ÍCH ======================= */

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);
}

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.SETTINGS);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .forEach(([k, v]) => { if (k) out[String(k).trim()] = String(v).trim(); });
  return out;
}

function parseDays_(s) {
  return String(s || '').split(/[,;\s]+/).filter(x => x !== '').map(Number).filter(n => !isNaN(n) && n >= 0);
}

function parseDate_(v) {
  if (!v) return new Date();
  if (typeof v === 'number') return new Date(v > 1e12 ? v : v * 1000);
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  const d = new Date(v);
  return isNaN(d) ? new Date() : d;
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// Bỏ dấu, chữ thường, chỉ giữ chữ + số → để so khớp nội dung chuyển khoản
function norm_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[đĐ]/g, 'd')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function hasKw_(normText, kw) {
  const k = norm_(kw);
  return !!k && (' ' + normText + ' ').includes(' ' + k + ' ');
}

// Chặn formula injection: chuỗi bắt đầu bằng = + - @ sẽ được thêm dấu ' phía trước
function safe_(s) {
  s = String(s == null ? '' : s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function money_(n) {
  const v = Math.round(Math.abs(Number(n) || 0)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (n < 0 ? '-' : '') + v + ' đ';
}

function fmtDate_(d) {
  return d instanceof Date ? Utilities.formatDate(d, TZ, 'dd/MM/yyyy') : '';
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
