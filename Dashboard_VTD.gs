/**
 * Dashboard_VTD.gs — tổng hợp số liệu cho trang /dashboard.html
 *
 * Thêm file này vào project Apps Script VTD (cùng project với Code_chinh_VTD.gs).
 * Cài đặt:
 *   1. Thêm 1 dòng vào map trong vtdApp_apiDispatch_:   dashboardV2: vtdApp_dashboardV2,
 *      Quyền xem: admin luôn xem được; nhân viên cần được bật ở Admin → Quyền xem Dashboard (màn dashview).
 *   2. Chạy tay dashBuild() một lần (cấp quyền), rồi chạy dashInstallTrigger().
 *   3. Deploy lại Web App (Manage deployments → Edit → New version) để giữ nguyên URL.
 *
 * Trang dashboard chỉ đọc 2 sheet ẩn _DASH_FACT và _DASH_SNAPSHOT, không quét
 * dữ liệu gốc mỗi lần mở, nên không làm chậm app nhập liệu.
 *
 * _DASH_FACT: mỗi dòng = [Ngày, Nguồn, Kênh, Loại, User, Giờ, Số dòng, Số lượng, Chỉ số]
 *   Chỉ số: op (thao tác nhập trên app) · order (lên đơn) · handover (bàn giao CT về shop)
 *           cond (tình trạng hoàn, Kênh = Chứng từ/Sản phẩm) · product (sản phẩm hoàn, Kênh = tên SP)
 *           pnclass (PN: Kênh = Phân loại, Loại = Tình trạng) · pnnote (PN: Kênh = Ghi chú, Loại = Phân loại)
 *
 * Tốc độ: mỗi sheet nguồn chỉ đọc đúng các cột cần (dash_read_), không đọc cả bảng.
 */

// Đổi mỗi lần sửa file này — dashboard hiện số phiên bản để biết máy chủ đã chạy bản mới chưa
const DASH_BUILD_VERSION = "2026-10-16e";

const DASH_CONFIG = {
  days: 200,               // số ngày lịch sử cho thao tác / lên đơn / bàn giao / tình trạng
  productDays: 100,        // sản phẩm nhiều dòng hơn nên giữ ngắn hơn
  timeBudgetMs: 280000,    // quá thời gian này thì bỏ qua các bước còn lại, vẫn ghi số đã có (giới hạn Google: 6 phút)
  factSheet: "_DASH_FACT",
  snapshotSheet: "_DASH_SNAPSHOT",
  timezone: "Asia/Ho_Chi_Minh",
  recentLimit: 30,
  handoverAliases: ["ngay ban giao ct", "ngay ban giao"],
  vtd: {
    // RAW_IN_M3 (app ghi trực tiếp) + RAW_IN_FULL (lịch sử), bỏ trùng theo ID
    historySheet: "RAW_IN_FULL",
    liveSheet: "RAW_IN_M3",
    // Danh sách đơn lên + chứng từ chưa nhận: "Full das chung tu VTD" / TongHop (VTD_APP.fullCtSpreadsheetId)
    fullCtSheet: "TongHop",
    // Nếu TongHop có cột trạng thái thì dùng cột đó; không có thì đối chiếu với đơn Chứng từ đã nhập trên app
    statusAliases: ["xac thuc hoa don", "xac thuc chung tu", "tinh trang chung tu", "trang thai chung tu"],
    openStatuses: ["chua nhan chung tu", "chua nhan"],
    derivedDays: 60        // cách đối chiếu: chỉ xét đơn lên trong 60 ngày gần nhất
  },
  pn: {
    spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
    sheetNames: ["Chứng từ_full", "Chứng từ_FF"],   // dùng sheet đầu tiên tồn tại
    productSheet: "Hoàn sản phẩm",
    statusSheet: "Chứng từ_FF",                 // tab chính: tháng hiện tại + tháng cũ chưa nhận
    openStatuses: ["chua nhan chung tu"]        // giá trị cột Xác thực hóa đơn được tính là "chưa nhận"
  },
  th: {
    // File TH Truemilk (cùng file với script FFPL tồn kho). Để trống ID = bỏ qua TH.
    spreadsheetId: "1TVck35ibO3zJ6lkKpHNyXckgLyjvUxEoLbMa-6kt-Pk",
    sheetName: "Hàng hoàn",
    headerRow: 3,          // tiêu đề tab Hàng hoàn nằm ở dòng 3 (giống FFPL.headerRows)
    dateAliases: ["ngay thang", "ngay hoan", "ngay hoan tra", "ngay", "date"],
    timeAliases: ["thoi gian", "gio", "time", "timestamp"],
    userAliases: ["user thao tac", "user", "nguoi nhap", "nguoi thao tac", "email"],
    qtyAliases: ["so luong", "sl", "qty"],
    maAliases: ["ma don", "ma don ghtk"],
    kenhAliases: ["kho", "kenh", "slog"],
    nameAliases: ["ten san pham", "ten sp", "san pham"],
    statusAliases: ["tinh trang", "tinh trang hang", "tinh trang ff"],
    ordersSheet: "DS đơn all",   // đơn lên của TH
    ordersHeaderRow: 1
  }
};

/* ============================== API ============================== */

function vtdApp_dashboardV2(params) {
  params = params || {};
  // Chỉ tin email lấy từ sessionToken — không nhận email gửi kèm từ client.
  const token = String(params.sessionToken || "").trim();
  const email = token ? vtdApp_sessionEmail_(token) : "";
  if (!email) return vtdApp_fail_("Phiên đăng nhập đã hết. Mở app, đăng nhập lại rồi tải lại dashboard.");
  if (!dash_canView_(email)) return vtdApp_fail_("Tài khoản chưa được cấp quyền xem dashboard.");

  const ss = vtdApp_ss_();
  const factSh = ss.getSheetByName(DASH_CONFIG.factSheet);
  const snapSh = ss.getSheetByName(DASH_CONFIG.snapshotSheet);
  if (!factSh || !snapSh) return vtdApp_fail_("Chưa có dữ liệu tổng hợp. Admin cần chạy dashBuild() một lần.");

  const last = factSh.getLastRow();
  const width = Math.min(9, Math.max(8, factSh.getLastColumn()));
  const facts = last > 1 ? factSh.getRange(2, 1, last - 1, width).getValues().map(r => [
    r[0] instanceof Date ? Utilities.formatDate(r[0], DASH_CONFIG.timezone, "yyyy-MM-dd") : String(r[0]),
    String(r[1]), String(r[2]), String(r[3]), String(r[4]), Number(r[5]), Number(r[6]) || 0, Number(r[7]) || 0,
    String(r[8] || "op")
  ]) : [];
  let snapshot = {};
  try { snapshot = JSON.parse(String(snapSh.getRange(2, 2).getValue() || "{}")); } catch (err) { snapshot = {}; }
  return vtdApp_ok_({
    updatedAt: String(snapSh.getRange(2, 1).getDisplayValue() || ""),
    today: Utilities.formatDate(new Date(), DASH_CONFIG.timezone, "yyyy-MM-dd"),
    apiVersion: DASH_BUILD_VERSION,
    viewer: email,
    facts: facts,
    snapshot: snapshot
  });
}

// Quyền riêng: màn "dashview" (admin bật ở Admin → Quyền xem Dashboard).
// Không dùng action "dashboard" vì action đó là Trang chủ, nhân viên nào cũng có.
function dash_canView_(email) {
  const config = vtdApp_permissionConfig_();
  if ((config.adminEmails || []).indexOf(email) >= 0) return true;
  const user = config.users && config.users[email];
  if (!user || user.active === false || user.allowed === false) return false;
  if (String(user.role || "").toLowerCase() === "admin") return true;
  return (user.screens || []).indexOf(DASH_SCREEN_ID) >= 0;
}
const DASH_SCREEN_ID = "dashview";

/* ============================ Trigger ============================ */

function dashInstallTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "dashBuild")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("dashBuild").timeBased().everyMinutes(10).create();
}

function dashBuild() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    console.log("dashBuild: đang có một lượt dashBuild khác chạy (thường là trigger 10 phút). Chờ lượt đó xong rồi chạy lại.");
    return;
  }
  const started = Date.now();
  try {
    const tz = DASH_CONFIG.timezone;
    const now = new Date();
    const keyBack = days => Utilities.formatDate(new Date(now.getTime() - days * 86400000), tz, "yyyy-MM-dd");
    const fromKey = keyBack(DASH_CONFIG.days);
    const productFromKey = keyBack(DASH_CONFIG.productDays);
    const todayKey = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    const facts = {};          // key -> [date, src, kenh, loai, user, hour, count, qty, metric]
    const recent = [];
    const freshness = {};
    const errors = [];
    const timings = {};

    const put = (metric, src, dateKey, kenh, loai, user, hour, qty) => {
      if (!dateKey || dateKey < (metric === "product" ? productFromKey : fromKey) || dateKey > todayKey) return;
      const k = metric + "|" + dateKey + "|" + src + "|" + kenh + "|" + loai + "|" + user + "|" + hour;
      const f = facts[k] || (facts[k] = [dateKey, src, kenh, loai, user, hour, 0, 0, metric]);
      f[6] += 1;
      f[7] += qty;
    };
    // Thao tác nhập trên app
    const add = (src, dateKey, hour, kenh, loai, user, qty, recentRow) => {
      if (!dateKey) return;
      if (!freshness[src] || dateKey > freshness[src]) freshness[src] = dateKey;
      if (dateKey < fromKey) return;
      put("op", src, dateKey, kenh, loai, user, hour, qty);
      if (recentRow && dateKey >= keyBack(3)) recent.push(recentRow);
    };
    // Chỉ số khác: lên đơn, bàn giao, tình trạng, sản phẩm
    const metric = (m, src, dateKey, kenh, loai, qty) => put(m, src, dateKey, String(kenh || "Khác").trim() || "Khác", String(loai || "").trim(), "", -1, Number(qty) || 0);

    // Mỗi bước có đo thời gian; hết ngân sách thời gian thì bỏ qua bước còn lại
    const step = (name, fn) => {
      if (Date.now() - started > DASH_CONFIG.timeBudgetMs) { errors.push(name + ": bỏ qua vì hết thời gian chạy"); return null; }
      const t = Date.now();
      try { return fn(); }
      catch (err) { errors.push(name + ": " + (err && err.message || err)); return null; }
      finally { timings[name] = Math.round((Date.now() - t) / 100) / 10; }
    };

    const vtdReceived = step("VTD thao tác", () => dash_collectVtd_(add, metric));
    const sync = step("VTD đồng bộ", () => dash_vtdSync_()) || {};
    const vtdBacklog = vtdReceived ? step("VTD lên đơn/chưa nhận", () => dash_collectVtdBacklog_(now, vtdReceived, metric)) : null;
    step("VTD sản phẩm", () => dash_collectVtdProducts_(metric));
    const pnBacklog = step("Phạm Nguyên", () => dash_collectPn_(add, now, metric));
    step("TH hàng hoàn", () => dash_collectTh_(add, metric));
    step("TH lên đơn", () => dash_collectThOrders_(metric));
    step("Lỗi hàng đợi", () => {
      if (typeof vtdApp_queueErrorSheet_ === "function") sync.queueErrors = Math.max(0, vtdApp_queueErrorSheet_().getLastRow() - 1);
    });

    recent.sort((a, b) => String(b.at).localeCompare(String(a.at)));
    const snapshot = {
      freshness: freshness,
      sync: sync,
      pnBacklog: pnBacklog,
      backlogs: {VTD: vtdBacklog, PN: pnBacklog},
      recent: recent.slice(0, DASH_CONFIG.recentLimit),
      errors: errors,
      thConfigured: !!(DASH_CONFIG.th.spreadsheetId && DASH_CONFIG.th.sheetName),
      metrics: 2,          // phiên bản dữ liệu: có cột Chỉ số
      buildVersion: DASH_BUILD_VERSION,
      timings: timings,
      counts: {}
    };

    const tWrite = Date.now();
    const ss = vtdApp_ss_();
    const factSh = dash_sheet_(ss, DASH_CONFIG.factSheet, ["Ngày", "Nguồn", "Kênh", "Loại", "User", "Giờ", "Số dòng", "Số lượng", "Chỉ số"]);
    const rows = Object.keys(facts).map(k => facts[k]).sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
    // Cột Ngày để dạng chữ, không cho Sheets tự đổi "2026-10-16" thành kiểu Ngày.
    factSh.getRange("A:A").setNumberFormat("@");
    // Ghi dữ liệu mới trước, xóa phần đuôi cũ sau — không có lúc sheet trống.
    if (rows.length) {
      if (factSh.getMaxRows() < rows.length + 1) factSh.insertRowsAfter(factSh.getMaxRows(), rows.length + 1 - factSh.getMaxRows());
      factSh.getRange(2, 1, rows.length, 9).setValues(rows);
    }
    const extra = factSh.getLastRow() - 1 - rows.length;
    if (extra > 0) factSh.getRange(rows.length + 2, 1, extra, 9).clearContent();
    timings["Ghi sheet"] = Math.round((Date.now() - tWrite) / 100) / 10;

    const sum = snapshot.counts;
    rows.forEach(r => { const k = r[8] + ":" + r[1]; sum[k] = (sum[k] || 0) + r[6]; });

    const snapSh = dash_sheet_(ss, DASH_CONFIG.snapshotSheet, ["Cập nhật lúc", "Snapshot JSON"]);
    snapSh.getRange(2, 1, 1, 2).setValues([[
      Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss"),
      JSON.stringify(snapshot)
    ]]);

    // Tóm tắt trong Execution log để tự kiểm tra sau mỗi lần chạy tay
    console.log("dashBuild " + DASH_BUILD_VERSION + ": xong trong " + Math.round((Date.now() - started) / 1000) + " giây; " + rows.length + " dòng tổng hợp");
    console.log("Thời gian từng bước (giây): " + JSON.stringify(timings));
    console.log("Số liệu đọc được: " + JSON.stringify(sum));
    console.log("Thao tác mới nhất: " + JSON.stringify(freshness) + (errors.length ? " | LỖI: " + errors.join(" | ") : ""));
    if (vtdBacklog) console.log("VTD chưa nhận CT: " + vtdBacklog.total + " (" + vtdBacklog.method + "); cột TongHop: " + vtdBacklog.headers.join(" | "));
    if (pnBacklog) console.log("PN chưa nhận CT: " + pnBacklog.total + "; " + JSON.stringify(pnBacklog.byStatus));
  } finally {
    lock.releaseLock();
  }
}

/* ======================= Đọc sheet theo cột ======================= */

// Chỉ đọc các cột cần dùng. spec = {khóa: [alias[], vị trí dự phòng?]}.
// Trả về {header (đủ cột), c: {khóa: vị trí trong dòng rút gọn hoặc -1}, rows: dòng rút gọn (không có tiêu đề)}.
function dash_read_(sh, spec, headerRow) {
  headerRow = Math.max(1, Number(headerRow) || 1);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const out = {header: [], c: {}, rows: []};
  if (lastCol < 1 || lastRow < headerRow) { Object.keys(spec).forEach(k => out.c[k] = -1); return out; }
  out.header = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const col = vtdApp_headerMap_(out.header);
  const idx = {}, need = [];
  Object.keys(spec).forEach(k => {
    let i = dash_col_(col, spec[k][0], spec[k].length > 1 ? spec[k][1] : -1);
    if (i >= lastCol) i = -1;
    idx[k] = i;
    if (i >= 0 && need.indexOf(i) < 0) need.push(i);
  });
  need.sort((a, b) => a - b);
  const pos = {};
  need.forEach((i, j) => pos[i] = j);
  Object.keys(spec).forEach(k => out.c[k] = idx[k] >= 0 ? pos[idx[k]] : -1);
  const n = lastRow - headerRow;
  if (n <= 0 || !need.length) return out;
  // Gom các cột gần nhau thành 1 lần đọc
  const runs = [];
  need.forEach(i => {
    const last = runs[runs.length - 1];
    if (last && i - last.end <= 3) last.end = i; else runs.push({start: i, end: i});
  });
  const data = {};
  runs.forEach(run => {
    const vals = sh.getRange(headerRow + 1, run.start + 1, n, run.end - run.start + 1).getValues();
    for (let j = run.start; j <= run.end; j++) if (pos[j] != null) data[j] = vals.map(r => r[j - run.start]);
  });
  for (let r = 0; r < n; r++) {
    const row = new Array(need.length);
    for (let j = 0; j < need.length; j++) row[j] = data[need[j]][r];
    out.rows.push(row);
  }
  return out;
}

/* ============================ Nguồn VTD ============================ */

// Đọc cả RAW_IN_M3 (dữ liệu mới, app ghi trực tiếp) và RAW_IN_FULL (lịch sử); bỏ dòng trùng theo ID
function dash_collectVtd_(add, metric) {
  const ss = vtdApp_ss_();
  const received = {};
  const seenIds = {};
  [DASH_CONFIG.vtd.liveSheet, DASH_CONFIG.vtd.historySheet].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) dash_collectVtdSheet_(sh, add, metric, received, seenIds);
  });
  return received;
}

function dash_collectVtdSheet_(sh, add, metric, received, seenIds) {
  // Cùng alias + vị trí dự phòng với vtdApp_compactRecord_
  const d = dash_read_(sh, {
    id: [["id", "raw_id"], 0],
    date: [["ngay hoan tra", "ngay nhap"], 1],
    time: [["thoi gian nhap hoan", "thoi gian", "timestamp"], 2],
    ma: [["ma don", "ma don ghn", "ma don ghtk", "madon"], 3],
    slog: [["slog"], 4],
    loai: [["loai hoan", "loai st"], 6],
    po: [["ma po chung tu hoan", "ma po"], 7],
    tinhTrang: [["tinh trang chung tu", "tinh trang"], 8],
    user: [["user_email", "user email", "email", "nguoi nhap", "user thao tac", "created by"], 11],
    od: [["od"]],
    hd: [["so hoa don", "invoice"]]
  });
  const c = d.c;
  const get = (row, i) => i >= 0 ? row[i] : "";
  d.rows.forEach(row => {
    if (row.join("") === "") return;
    const uid = String(get(row, c.id) || "").trim() || [get(row, c.ma), get(row, c.date), get(row, c.time), get(row, c.loai)].join("|");
    if (seenIds[uid]) return;
    seenIds[uid] = true;
    // Khóa các đơn Chứng từ đã nhập trên app → dùng để biết đơn nào ở TongHop đã nhận
    const isDoc = vtdApp_norm_(get(row, c.loai)).indexOf("chung tu") >= 0;
    if (isDoc) [c.ma, c.po, c.od, c.hd].forEach(i => { if (i >= 0) vtdApp_lookupKeys_(row[i]).forEach(k => received[k] = true); });
    const dateKey = dash_key_(get(row, c.date)) || dash_key_(get(row, c.time));
    if (!dateKey) return;
    const user = dash_user_(get(row, c.user));
    const kenh = String(get(row, c.slog) || "Khác").trim() || "Khác";
    const loai = String(get(row, c.loai) || "Khác").trim() || "Khác";
    add("VTD", dateKey, dash_hour_(get(row, c.time)), kenh, loai, user, 1, {
      at: dateKey + " " + dash_hhmm_(get(row, c.time)), src: "VTD", ma: String(get(row, c.ma) || ""), kenh: kenh, loai: loai, user: user
    });
    // Tình trạng chứng từ (Bình thường / lỗi...) — chỉ áp dụng cho loại Chứng từ
    if (isDoc && c.tinhTrang >= 0) metric("cond", "VTD", dateKey, "Chứng từ", String(row[c.tinhTrang] || "").trim() || "(trống)", 1);
  });
}

// List_SKU_hoàn: sản phẩm hoàn VTD (Tên sản phẩm hoàn về, Số lượng, Tình trạng, Ngày)
function dash_collectVtdProducts_(metric) {
  const sh = vtdApp_ss_().getSheetByName(VTD_APP.skuSheet);
  if (!sh || sh.getLastRow() < 2) return;
  const d = dash_read_(sh, {
    name: [["ten san pham hoan ve", "ten san pham hoan", "ten san pham", "sku"]],
    qty: [["so luong", "qty"]],
    status: [["tinh trang"]],
    date: [["ngay", "date"]]
  });
  const c = d.c;
  if (c.name < 0 || c.date < 0) throw new Error("List_SKU_hoàn thiếu cột Tên sản phẩm hoặc Ngày");
  d.rows.forEach(row => {
    const dateKey = dash_key_(row[c.date]);
    const name = String(row[c.name] || "").trim();
    if (!dateKey || !name) return;
    const status = String(c.status >= 0 ? row[c.status] : "").trim() || "(trống)";
    const qty = Number(c.qty >= 0 ? row[c.qty] : 0) || 0;
    metric("product", "VTD", dateKey, name, status, qty);
    metric("cond", "VTD", dateKey, "Sản phẩm", status, qty);
  });
}

function dash_collectVtdBacklog_(now, received, metric) {
  const cfg = DASH_CONFIG.vtd;
  const sh = SpreadsheetApp.openById(VTD_APP.fullCtSpreadsheetId).getSheetByName(cfg.fullCtSheet);
  if (!sh) throw new Error("Không thấy tab " + cfg.fullCtSheet + " trong Full das chung tu VTD");
  const backlog = {total: 0, over7: 0, buckets: [0, 0, 0, 0], byStore: {}, byStatus: {}, oldest: [],
    statusSheet: "Full das chứng từ VTD / " + cfg.fullCtSheet, method: "", headers: [], hasHandover: false};
  if (sh.getLastRow() < 2) return backlog;
  const d = dash_read_(sh, {
    ma: [["ma don", "ma don ghtk"]],
    od: [["od"]],
    po: [["ma po", "po"]],
    hd: [["so hoa don", "invoice"]],
    khach: [["khach hang", "ten khach hang", "customer"]],
    loaiSt: [["loai st", "loai sieu thi"]],
    lenDon: [["ngay len don", "date"]],
    banGiao: [DASH_CONFIG.handoverAliases],
    status: [cfg.statusAliases]
  });
  const c = d.c;
  backlog.headers = d.header.map(h => String(h || "").trim()).filter(Boolean);
  const byStatusCol = c.status >= 0;
  backlog.method = byStatusCol ? "theo cột trạng thái = Chưa nhận" : "đơn lên trong " + cfg.derivedDays + " ngày chưa có bản Chứng từ trên app";
  backlog.hasHandover = c.banGiao >= 0;
  const todayMs = dash_dayMs_(now);
  const seen = {};
  const get = (row, i) => i >= 0 ? row[i] : "";
  d.rows.forEach(row => {
    const ma = String(get(row, c.ma) || get(row, c.od) || get(row, c.po) || "").trim();
    if (!ma) return;
    const uniq = vtdApp_norm_(get(row, c.od) || ma);
    if (seen[uniq]) return;
    seen[uniq] = true;
    const khach = String(get(row, c.khach) || "").trim();
    const store = String(get(row, c.loaiSt) || "").trim() || vtdApp_storeType_(khach);
    const lenKey = dash_key_(get(row, c.lenDon));
    metric("order", "VTD", lenKey, store, "Lên đơn", 1);
    if (c.banGiao >= 0) metric("handover", "VTD", dash_key_(get(row, c.banGiao)), store, "Bàn giao", 1);
    const days = lenKey ? Math.max(0, Math.round((todayMs - dash_dayMs_(lenKey)) / 86400000)) : -1;

    let open, statusLabel;
    if (byStatusCol) {
      statusLabel = String(get(row, c.status) || "").trim() || "(trống)";
      open = cfg.openStatuses.indexOf(vtdApp_norm_(get(row, c.status))) >= 0;
    } else {
      if (days < 0 || days > cfg.derivedDays) return;
      const keys = [];
      [c.ma, c.po, c.od, c.hd].forEach(i => { if (i >= 0) Array.prototype.push.apply(keys, vtdApp_lookupKeys_(row[i])); });
      open = !keys.some(k => received[k]);
      statusLabel = open ? "Chưa nhận chứng từ" : "Đã nhận (có trên app)";
    }
    backlog.byStatus[statusLabel] = (backlog.byStatus[statusLabel] || 0) + 1;
    if (!open) return;
    backlog.total++;
    backlog.byStore[store] = (backlog.byStore[store] || 0) + 1;
    if (days < 0) return;
    if (days > 7) backlog.over7++;
    backlog.buckets[days <= 3 ? 0 : days <= 7 ? 1 : days <= 15 ? 2 : 3]++;
    backlog.oldest.push({ma: ma, khach: khach, kenh: store, ngay: lenKey, days: days, status: statusLabel});
  });
  backlog.oldest.sort((a, b) => b.days - a.days);
  backlog.oldest = backlog.oldest.slice(0, 12);
  return backlog;
}

function dash_vtdSync_() {
  const sh = vtdApp_ss_().getSheetByName(DASH_CONFIG.vtd.liveSheet);
  const out = {pending: 0, errHangHoan: 0, errChungTu: 0};
  if (!sh || sh.getLastRow() < 2) return out;
  const d = dash_read_(sh, {s: [["sync_status"]], p: [["pj2_status"]]});
  const isErr = v => v.indexOf("ERROR") >= 0 || v.indexOf("IMG") >= 0;
  const isPending = v => v === "YES" || v === "RUNNING" || v === "PENDING";
  d.rows.forEach(row => {
    const s = String(d.c.s >= 0 ? row[d.c.s] : "").toUpperCase().trim();
    const p = String(d.c.p >= 0 ? row[d.c.p] : "").toUpperCase().trim();
    if (isErr(s)) out.errChungTu++;
    if (isErr(p)) out.errHangHoan++;
    if (isPending(s) || isPending(p)) out.pending++;
  });
  return out;
}

/* ========================= Nguồn Phạm Nguyên ========================= */

function dash_collectPn_(add, now, metric) {
  const cfg = DASH_CONFIG.pn;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  let sh = null;
  for (let i = 0; i < cfg.sheetNames.length && !sh; i++) sh = ss.getSheetByName(cfg.sheetNames[i]);
  const backlog = {total: 0, over7: 0, buckets: [0, 0, 0, 0], byStore: {}, byStatus: {}, oldest: [], statusSheet: "",
    method: "theo cột Xác thực hóa đơn = Chưa nhận chứng từ", hasHandover: false};
  const spec = {
    ma: [["ma don ghtk", "ma don"]],
    od: [["so don hang", "od"]],
    khach: [["khach hang"]],
    store: [["loai sieu thi"]],
    // Trạng thái chứng từ nằm ở cột "Xác thực hóa đơn": Chưa nhận / Đã nhận chứng từ / Shop hủy OD
    status: [["xac thuc hoa don"]],
    time: [["thoi gian"]],
    user: [["user thao tac"]],
    lenDon: [["ngay len don"]],
    banGiao: [DASH_CONFIG.handoverAliases]
  };
  const get = (row, i) => i >= 0 ? row[i] : "";

  // 1) Tab full (mọi tháng): thao tác chứng từ, lên đơn, bàn giao CT
  if (sh && sh.getLastRow() > 1) {
    const d = dash_read_(sh, spec), c = d.c;
    backlog.hasHandover = c.banGiao >= 0;
    const seen = {};
    d.rows.forEach(row => {
      const ma = String(get(row, c.ma) || get(row, c.od) || "");
      if (!ma) return;
      const store = String(get(row, c.store) || "Khác").trim() || "Khác";
      const uniq = vtdApp_norm_(get(row, c.od) || ma);
      if (!seen[uniq]) {
        seen[uniq] = true;
        metric("order", "PN", dash_key_(get(row, c.lenDon)), store, "Lên đơn", 1);
        if (c.banGiao >= 0) metric("handover", "PN", dash_key_(get(row, c.banGiao)), store, "Bàn giao", 1);
      }
      if (!get(row, c.user)) return;
      const actKey = dash_key_(get(row, c.time));
      if (!actKey) return;
      const user = dash_user_(get(row, c.user));
      add("PN", actKey, dash_hour_(get(row, c.time)), store, "Chứng từ", user, 1, {
        at: actKey + " " + dash_hhmm_(get(row, c.time)), src: "PN", ma: ma, kenh: store, loai: "Chứng từ", user: user
      });
    });
  }

  // 2) Tồn chứng từ chưa nhận: đọc tab chính (khớp với bảng "Xác thực hóa đơn" trên sheet tổng)
  const stSh = ss.getSheetByName(cfg.statusSheet) || sh;
  if (stSh && stSh.getLastRow() > 1) {
    backlog.statusSheet = stSh.getName();
    const d = dash_read_(stSh, spec), c = d.c;
    if (c.status < 0) throw new Error("Tab " + stSh.getName() + " không có cột Xác thực hóa đơn");
    const todayMs = dash_dayMs_(now);
    d.rows.forEach(row => {
      const ma = String(get(row, c.ma) || get(row, c.od) || "");
      if (!ma) return;
      const status = String(get(row, c.status) || "").trim();
      const statusLabel = status || "(trống)";
      backlog.byStatus[statusLabel] = (backlog.byStatus[statusLabel] || 0) + 1;
      if (cfg.openStatuses.indexOf(vtdApp_norm_(status)) < 0) return;
      const store = String(get(row, c.store) || "Khác").trim() || "Khác";
      const lenKey = dash_key_(get(row, c.lenDon));
      const days = lenKey ? Math.max(0, Math.round((todayMs - dash_dayMs_(lenKey)) / 86400000)) : -1;
      backlog.total++;
      backlog.byStore[store] = (backlog.byStore[store] || 0) + 1;
      if (days < 0) return;            // thiếu Ngày lên đơn: vẫn đếm vào tổng, không xếp tuổi
      if (days > 7) backlog.over7++;
      backlog.buckets[days <= 3 ? 0 : days <= 7 ? 1 : days <= 15 ? 2 : 3]++;
      backlog.oldest.push({ma: ma, khach: String(get(row, c.khach) || ""), kenh: store, ngay: lenKey, days: days, status: status});
    });
    backlog.oldest.sort((a, b) => b.days - a.days);
    backlog.oldest = backlog.oldest.slice(0, 12);
  }

  // 3) Hoàn sản phẩm: mỗi dòng = 1 sản phẩm; qty = số lượng; Tình trạng FF
  const prod = ss.getSheetByName(cfg.productSheet);
  if (prod && prod.getLastRow() > 1) {
    const d = dash_read_(prod, {
      date: [["ngay hoan tra", "ngay"]],
      time: [["thoi gian", "gio"]],
      user: [["user thao tac", "user"]],
      qty: [["so luong"]],
      name: [["ten san pham", "ten vat tu"]],
      status: [["tinh trang ff", "tinh trang"]],
      cls: [["phan loai"]],
      note: [["ghi chu"]]
    }), c = d.c;
    d.rows.forEach(row => {
      const dateKey = c.date >= 0 ? dash_key_(row[c.date]) : "";
      if (!dateKey) return;
      const qty = Number(get(row, c.qty)) || 0;
      add("PN", dateKey, dash_hour_(get(row, c.time)), "Hàng hoàn PN", "Sản phẩm", dash_user_(get(row, c.user)), qty, null);
      const status = String(get(row, c.status) || "").trim() || "(trống)";
      const name = String(get(row, c.name) || "").trim();
      if (name) metric("product", "PN", dateKey, name, status, qty);
      metric("cond", "PN", dateKey, "Sản phẩm", status, qty);
      // Phân loại (Hoàn trả / Hàng thu hồi) và Ghi chú (nhiều lựa chọn, nối bằng "; ")
      const cls = String(get(row, c.cls) || "").trim() || "(chưa phân loại)";
      metric("pnclass", "PN", dateKey, cls, status, qty);
      dash_noteItems_(get(row, c.note)).forEach(note => metric("pnnote", "PN", dateKey, note, cls, qty));
    });
  }
  return backlog;
}

/* ============================ Nguồn TH ============================ */

function dash_collectTh_(add, metric) {
  const cfg = DASH_CONFIG.th;
  if (!cfg.spreadsheetId || !cfg.sheetName) return;
  const sh = SpreadsheetApp.openById(cfg.spreadsheetId).getSheetByName(cfg.sheetName);
  if (!sh) throw new Error("Không thấy tab " + cfg.sheetName);
  const d = dash_read_(sh, {
    date: [cfg.dateAliases], time: [cfg.timeAliases], user: [cfg.userAliases], qty: [cfg.qtyAliases],
    ma: [cfg.maAliases], kenh: [cfg.kenhAliases], name: [cfg.nameAliases], status: [cfg.statusAliases]
  }, cfg.headerRow), c = d.c;
  if (c.date < 0 && c.time < 0) throw new Error("Tab " + cfg.sheetName + " thiếu cột Ngày tháng/Thời gian ở dòng " + cfg.headerRow);
  const get = (row, i) => i >= 0 ? row[i] : "";
  d.rows.forEach(row => {
    const dateKey = dash_key_(get(row, c.date)) || dash_key_(get(row, c.time));
    if (!dateKey) return;
    const user = dash_user_(get(row, c.user));
    const kenh = String(get(row, c.kenh) || "").trim() || "TH Truemilk";
    const qty = Number(get(row, c.qty)) || 0;
    add("TH", dateKey, dash_hour_(get(row, c.time)), kenh, "Hàng hoàn", user, qty, {
      at: dateKey + " " + dash_hhmm_(get(row, c.time)), src: "TH", ma: String(get(row, c.ma) || ""), kenh: kenh, loai: "Hàng hoàn", user: user
    });
    const name = String(get(row, c.name) || "").trim();
    const status = c.status >= 0 ? (String(row[c.status] || "").trim() || "(trống)") : "";
    if (name) metric("product", "TH", dateKey, name, status || "(không ghi tình trạng)", qty);
    if (status) metric("cond", "TH", dateKey, "Sản phẩm", status, qty);
  });
}

// DS đơn all: mỗi dòng 1 SKU của đơn → gộp theo mã đơn, ngày = Thời gian tạo
function dash_collectThOrders_(metric) {
  const cfg = DASH_CONFIG.th;
  if (!cfg.spreadsheetId || !cfg.ordersSheet) return;
  const sh = SpreadsheetApp.openById(cfg.spreadsheetId).getSheetByName(cfg.ordersSheet);
  if (!sh) throw new Error("Không thấy tab " + cfg.ordersSheet);
  const d = dash_read_(sh, {
    date: [["thoi gian tao", "ngay tao", "ngay len don"]],
    ghtk: [["ma don ghtk"]], od: [["ma od"]], gam: [["ma don gam"]], dt: [["ma don doi tac"]]
  }, cfg.ordersHeaderRow), c = d.c;
  if (c.date < 0) throw new Error("Tab " + cfg.ordersSheet + " thiếu cột Thời gian tạo");
  const ids = [c.ghtk, c.od, c.gam, c.dt].filter(i => i >= 0);
  const seen = {};
  d.rows.forEach((row, r) => {
    const id = ids.map(i => String(row[i] || "").trim()).filter(Boolean)[0] || ("row" + r);
    if (seen[id]) return;
    seen[id] = true;
    metric("order", "TH", dash_key_(row[c.date]), "TH Truemilk", "Lên đơn", 1);
  });
}

/* ============================ Helpers ============================ */

function dash_sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.hideSheet();
  }
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  return sh;
}

function dash_col_(col, aliases, fallback) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = col[vtdApp_norm_(aliases[i])];
    if (idx != null) return idx;
  }
  return fallback;
}

// Ngày dạng yyyy-MM-dd. Kiểu Date đọc thẳng (nhanh hơn Utilities.formatDate khi chạy hàng chục nghìn dòng);
// Apps Script trả Date theo múi giờ của project (Asia/Ho_Chi_Minh).
function dash_key_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "";
    return value.getFullYear() + "-" + ("0" + (value.getMonth() + 1)).slice(-2) + "-" + ("0" + value.getDate()).slice(-2);
  }
  if (value === "" || value == null) return "";
  return vtdApp_dateKey_(value);
}

function dash_hour_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getHours();
  const m = String(value == null ? "" : value).match(/(\d{1,2}):(\d{2})/);
  return m ? Math.min(23, Number(m[1])) : -1;
}

function dash_hhmm_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return ("0" + value.getHours()).slice(-2) + ":" + ("0" + value.getMinutes()).slice(-2);
  const m = String(value == null ? "" : value).match(/(\d{1,2}):(\d{2})/);
  return m ? ("0" + m[1]).slice(-2) + ":" + m[2] : "";
}

// "Bình thường; Hàng cận date; Khác: móp góc" → ["Bình thường", "Hàng cận date", "Khác"]
function dash_noteItems_(value) {
  const text = String(value == null ? "" : value).trim();
  if (!text) return ["(không ghi chú)"];
  const seen = {}, out = [];
  text.split(/\s*;\s*/).forEach(part => {
    let item = part.trim();
    if (!item) return;
    if (/^kh[áa]c\s*:/i.test(item)) item = "Khác";
    if (item.length > 40) item = item.slice(0, 40) + "…";
    if (!seen[item]) { seen[item] = true; out.push(item); }
  });
  return out.length ? out : ["(không ghi chú)"];
}

function dash_user_(value) {
  const text = String(value == null ? "" : value).trim().toLowerCase();
  return text || "(không rõ)";
}

function dash_dayMs_(value) {
  const key = value instanceof Date ? Utilities.formatDate(value, DASH_CONFIG.timezone, "yyyy-MM-dd") : String(value);
  const p = key.split("-").map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2]);
}
