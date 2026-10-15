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
 */

// Đổi mỗi lần sửa file này — dashboard hiện số phiên bản để biết máy chủ đã chạy bản mới chưa
const DASH_BUILD_VERSION = "2026-10-16c";

const DASH_CONFIG = {
  days: 200,               // số ngày lịch sử cho thao tác / lên đơn / bàn giao / tình trạng
  productDays: 100,        // sản phẩm nhiều dòng hơn nên giữ ngắn hơn
  factSheet: "_DASH_FACT",
  snapshotSheet: "_DASH_SNAPSHOT",
  timezone: "Asia/Ho_Chi_Minh",
  recentLimit: 30,
  handoverAliases: ["ngay ban giao ct", "ngay ban giao"],
  vtd: {
    // RAW_IN_FULL chứa toàn bộ lịch sử; RAW_IN_M3 dùng cho trạng thái đồng bộ
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
  if (!lock.tryLock(5000)) return;
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

    const put = (metric, src, dateKey, kenh, loai, user, hour, qty) => {
      if (!dateKey || dateKey < (metric === "product" ? productFromKey : fromKey) || dateKey > todayKey) return;
      const k = [metric, dateKey, src, kenh, loai, user, hour].join("|");
      const f = facts[k] || (facts[k] = [dateKey, src, kenh, loai, user, hour, 0, 0, metric]);
      f[6] += 1;
      f[7] += qty;
    };
    // Thao tác nhập trên app (giữ chữ ký cũ)
    const add = (src, dateKey, hour, kenh, loai, user, qty, recentRow) => {
      if (!dateKey) return;
      if (!freshness[src] || dateKey > freshness[src]) freshness[src] = dateKey;
      if (dateKey < fromKey) return;
      put("op", src, dateKey, kenh, loai, user, hour, qty);
      if (recentRow) recent.push(recentRow);
    };
    // Chỉ số khác: lên đơn, bàn giao, tình trạng, sản phẩm
    const metric = (m, src, dateKey, kenh, loai, qty) => put(m, src, dateKey, String(kenh || "Khác").trim() || "Khác", String(loai || "").trim(), "", -1, Number(qty) || 0);

    let sync = {};
    let vtdReceived = null, vtdBacklog = null;
    try { vtdReceived = dash_collectVtd_(add, metric); } catch (err) { errors.push("VTD: " + err); }
    try { dash_collectVtdProducts_(metric); } catch (err) { errors.push("Sản phẩm VTD: " + err); }
    try { sync = dash_vtdSync_(); } catch (err) { errors.push("Đồng bộ VTD: " + err); }
    try { if (vtdReceived) vtdBacklog = dash_collectVtdBacklog_(now, vtdReceived, metric); } catch (err) { errors.push("VTD lên đơn/chưa nhận CT: " + err); }
    let pnBacklog = null;
    try { pnBacklog = dash_collectPn_(add, now, metric); } catch (err) { errors.push("Phạm Nguyên: " + err); }
    try { dash_collectTh_(add, metric); } catch (err) { errors.push("TH: " + err); }
    try { dash_collectThOrders_(metric); } catch (err) { errors.push("TH lên đơn: " + err); }
    try {
      if (typeof vtdApp_queueErrorSheet_ === "function") {
        const q = vtdApp_queueErrorSheet_();
        sync.queueErrors = Math.max(0, q.getLastRow() - 1);
      }
    } catch (err) {}

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
      counts: {}
    };

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

    // Tóm tắt trong Execution log để tự kiểm tra sau mỗi lần chạy tay
    const sum = snapshot.counts;
    rows.forEach(r => { const k = r[8] + ":" + r[1]; sum[k] = (sum[k] || 0) + r[6]; });
    console.log("dashBuild: " + rows.length + " dòng tổng hợp; " + JSON.stringify(sum)
      + "; thao tác mới nhất " + JSON.stringify(freshness) + (errors.length ? "; LỖI: " + errors.join(" | ") : ""));
    if (vtdBacklog) console.log("VTD chưa nhận CT: " + vtdBacklog.total + " (" + vtdBacklog.method + "); cột TongHop: " + vtdBacklog.headers.join(" | "));
    if (pnBacklog) console.log("PN chưa nhận CT: " + pnBacklog.total + "; " + JSON.stringify(pnBacklog.byStatus));

    const snapSh = dash_sheet_(ss, DASH_CONFIG.snapshotSheet, ["Cập nhật lúc", "Snapshot JSON"]);
    snapSh.getRange(2, 1, 1, 2).setValues([[
      Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss"),
      JSON.stringify(snapshot)
    ]]);
  } finally {
    lock.releaseLock();
  }
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
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const col = vtdApp_headerMap_(values[0]);
  // Cùng alias + vị trí dự phòng với vtdApp_compactRecord_
  const c = {
    date: dash_col_(col, ["ngay hoan tra", "ngay nhap"], 1),
    time: dash_col_(col, ["thoi gian nhap hoan", "thoi gian", "timestamp"], 2),
    ma: dash_col_(col, ["ma don", "ma don ghn", "ma don ghtk", "madon"], 3),
    slog: dash_col_(col, ["slog"], 4),
    loai: dash_col_(col, ["loai hoan", "loai st"], 6),
    po: dash_col_(col, ["ma po chung tu hoan", "ma po"], 7),
    tinhTrang: dash_col_(col, ["tinh trang chung tu", "tinh trang"], 8),
    od: dash_col_(col, ["od"], -1),
    hd: dash_col_(col, ["so hoa don", "invoice"], -1),
    user: dash_col_(col, ["user_email", "user email", "email", "nguoi nhap", "user thao tac", "created by"], 11),
    id: dash_col_(col, ["id", "raw_id"], 0)
  };
  // Khóa các đơn Chứng từ đã nhập trên app → dùng để biết đơn nào ở TongHop đã nhận
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.join("") === "") continue;
    const uid = String(row[c.id] || "").trim() || [row[c.ma], row[c.date], row[c.time], row[c.loai]].join("|");
    if (seenIds[uid]) continue;
    seenIds[uid] = true;
    const isDoc = vtdApp_norm_(row[c.loai]).indexOf("chung tu") >= 0;
    if (isDoc) {
      [c.ma, c.po, c.od, c.hd].forEach(i => { if (i >= 0) vtdApp_lookupKeys_(row[i]).forEach(k => received[k] = true); });
    }
    const dateKey = vtdApp_dateKey_(row[c.date]) || vtdApp_dateKey_(row[c.time]);
    if (!dateKey) continue;
    const hour = dash_hour_(row[c.time]);
    const user = dash_user_(row[c.user]);
    const kenh = String(row[c.slog] || "Khác").trim() || "Khác";
    const loai = String(row[c.loai] || "Khác").trim() || "Khác";
    add("VTD", dateKey, hour, kenh, loai, user, 1, {
      at: dateKey + " " + dash_hhmm_(row[c.time]), src: "VTD", ma: String(row[c.ma] || ""), kenh: kenh, loai: loai, user: user
    });
    // Tình trạng chứng từ (Bình thường / lỗi...) — chỉ áp dụng cho loại Chứng từ
    if (isDoc && c.tinhTrang >= 0) metric("cond", "VTD", dateKey, "Chứng từ", String(row[c.tinhTrang] || "").trim() || "(trống)", 1);
  }
}

// List_SKU_hoàn: sản phẩm hoàn VTD (Tên sản phẩm hoàn về, Số lượng, Tình trạng, Ngày)
function dash_collectVtdProducts_(metric) {
  const sh = vtdApp_ss_().getSheetByName(VTD_APP.skuSheet);
  if (!sh || sh.getLastRow() < 2) return;
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const col = vtdApp_headerMap_(values[0]);
  const c = {
    name: dash_col_(col, ["ten san pham hoan ve", "ten san pham hoan", "ten san pham", "sku"], -1),
    qty: dash_col_(col, ["so luong", "qty"], -1),
    status: dash_col_(col, ["tinh trang"], -1),
    date: dash_col_(col, ["ngay", "date"], -1)
  };
  if (c.name < 0 || c.date < 0) throw new Error("List_SKU_hoàn thiếu cột Tên sản phẩm hoặc Ngày");
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateKey = vtdApp_dateKey_(row[c.date]);
    const name = String(row[c.name] || "").trim();
    if (!dateKey || !name) continue;
    const status = String(c.status >= 0 ? row[c.status] : "").trim() || "(trống)";
    const qty = Number(c.qty >= 0 ? row[c.qty] : 0) || 0;
    metric("product", "VTD", dateKey, name, status, qty);
    metric("cond", "VTD", dateKey, "Sản phẩm", status, qty);
  }
}

function dash_collectVtdBacklog_(now, received, metric) {
  const cfg = DASH_CONFIG.vtd;
  const sh = SpreadsheetApp.openById(VTD_APP.fullCtSpreadsheetId).getSheetByName(cfg.fullCtSheet);
  if (!sh) throw new Error("Không thấy tab " + cfg.fullCtSheet + " trong Full das chung tu VTD");
  const backlog = {total: 0, over7: 0, buckets: [0, 0, 0, 0], byStore: {}, byStatus: {}, oldest: [],
    statusSheet: "Full das chứng từ VTD / " + cfg.fullCtSheet, method: "", headers: []};
  if (sh.getLastRow() < 2) return backlog;
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  backlog.headers = values[0].map(h => String(h || "").trim()).filter(Boolean);
  const col = vtdApp_headerMap_(values[0]);
  const c = {
    ma: dash_col_(col, ["ma don", "ma don ghtk"], -1),
    od: dash_col_(col, ["od"], -1),
    po: dash_col_(col, ["ma po", "po"], -1),
    hd: dash_col_(col, ["so hoa don", "invoice"], -1),
    khach: dash_col_(col, ["khach hang", "ten khach hang", "customer"], -1),
    loaiSt: dash_col_(col, ["loai st", "loai sieu thi"], -1),
    lenDon: dash_col_(col, ["ngay len don", "date"], -1),
    banGiao: dash_col_(col, DASH_CONFIG.handoverAliases, -1),
    status: dash_col_(col, cfg.statusAliases, -1)
  };
  const byStatusCol = c.status >= 0;
  backlog.method = byStatusCol
    ? "theo cột " + values[0][c.status] + " = Chưa nhận"
    : "đơn lên trong " + cfg.derivedDays + " ngày chưa có bản Chứng từ trên app";
  backlog.hasHandover = c.banGiao >= 0;
  const todayMs = dash_dayMs_(now);
  const seen = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const get = i => i >= 0 ? row[i] : "";
    const ma = String(get(c.ma) || get(c.od) || get(c.po) || "").trim();
    if (!ma) continue;
    const uniq = vtdApp_norm_(get(c.od) || ma);
    if (seen[uniq]) continue;
    seen[uniq] = true;
    const khach = String(get(c.khach) || "").trim();
    const store = String(get(c.loaiSt) || "").trim() || vtdApp_storeType_(khach);
    const lenKey = vtdApp_dateKey_(get(c.lenDon));
    metric("order", "VTD", lenKey, store, "Lên đơn", 1);
    if (c.banGiao >= 0) metric("handover", "VTD", vtdApp_dateKey_(get(c.banGiao)), store, "Bàn giao", 1);
    const days = lenKey ? Math.max(0, Math.round((todayMs - dash_dayMs_(lenKey)) / 86400000)) : -1;

    let open, statusLabel;
    if (byStatusCol) {
      statusLabel = String(get(c.status) || "").trim() || "(trống)";
      open = cfg.openStatuses.indexOf(vtdApp_norm_(get(c.status))) >= 0;
    } else {
      if (days < 0 || days > cfg.derivedDays) continue;
      const keys = [];
      [c.ma, c.po, c.od, c.hd].forEach(i => { if (i >= 0) Array.prototype.push.apply(keys, vtdApp_lookupKeys_(get(i))); });
      open = !keys.some(k => received[k]);
      statusLabel = open ? "Chưa nhận chứng từ" : "Đã nhận (có trên app)";
    }
    backlog.byStatus[statusLabel] = (backlog.byStatus[statusLabel] || 0) + 1;
    if (!open) continue;

    backlog.total++;
    backlog.byStore[store] = (backlog.byStore[store] || 0) + 1;
    if (days < 0) continue;
    if (days > 7) backlog.over7++;
    backlog.buckets[days <= 3 ? 0 : days <= 7 ? 1 : days <= 15 ? 2 : 3]++;
    backlog.oldest.push({ma: ma, khach: khach, kenh: store, ngay: lenKey, days: days, status: statusLabel});
  }
  backlog.oldest.sort((a, b) => b.days - a.days);
  backlog.oldest = backlog.oldest.slice(0, 12);
  return backlog;
}

function dash_vtdSync_() {
  const sh = vtdApp_ss_().getSheetByName(DASH_CONFIG.vtd.liveSheet);
  const out = {pending: 0, errHangHoan: 0, errChungTu: 0};
  if (!sh || sh.getLastRow() < 2) return out;
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const col = vtdApp_headerMap_(values[0]);
  const sCol = col["sync_status"], pCol = col["pj2_status"];
  for (let r = 1; r < values.length; r++) {
    const s = String(sCol != null ? values[r][sCol] : "").toUpperCase().trim();
    const p = String(pCol != null ? values[r][pCol] : "").toUpperCase().trim();
    const isErr = v => v.indexOf("ERROR") >= 0 || v.indexOf("IMG") >= 0;
    const isPending = v => v === "YES" || v === "RUNNING" || v === "PENDING";
    if (isErr(s)) out.errChungTu++;
    if (isErr(p)) out.errHangHoan++;
    if (isPending(s) || isPending(p)) out.pending++;
  }
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
  const pnCols = col => ({
    ma: dash_col_(col, ["ma don ghtk", "ma don"], -1),
    od: dash_col_(col, ["so don hang", "od"], -1),
    khach: dash_col_(col, ["khach hang"], -1),
    store: dash_col_(col, ["loai sieu thi"], -1),
    // Trạng thái chứng từ nằm ở cột "Xác thực hóa đơn": Chưa nhận / Đã nhận chứng từ / Shop hủy OD
    status: dash_col_(col, ["xac thuc hoa don"], -1),
    time: dash_col_(col, ["thoi gian"], -1),
    user: dash_col_(col, ["user thao tac"], -1),
    lenDon: dash_col_(col, ["ngay len don"], -1),
    banGiao: dash_col_(col, DASH_CONFIG.handoverAliases, -1)
  });

  // 1) Tab full (mọi tháng): thao tác chứng từ, lên đơn, bàn giao CT
  if (sh && sh.getLastRow() > 1) {
    const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const c = pnCols(vtdApp_headerMap_(values[0]));
    backlog.hasHandover = c.banGiao >= 0;
    const seen = {};
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const get = i => i >= 0 ? row[i] : "";
      const ma = String(get(c.ma) || get(c.od) || "");
      if (!ma) continue;
      const store = String(get(c.store) || "Khác").trim() || "Khác";
      const uniq = vtdApp_norm_(get(c.od) || ma);
      if (!seen[uniq]) {
        seen[uniq] = true;
        metric("order", "PN", vtdApp_dateKey_(get(c.lenDon)), store, "Lên đơn", 1);
        if (c.banGiao >= 0) metric("handover", "PN", vtdApp_dateKey_(get(c.banGiao)), store, "Bàn giao", 1);
      }
      if (!get(c.user)) continue;
      const actKey = vtdApp_dateKey_(get(c.time));
      if (!actKey) continue;
      const user = dash_user_(get(c.user));
      add("PN", actKey, dash_hour_(get(c.time)), store, "Chứng từ", user, 1, {
        at: actKey + " " + dash_hhmm_(get(c.time)), src: "PN", ma: ma, kenh: store, loai: "Chứng từ", user: user
      });
    }
  }

  // 2) Tồn chứng từ chưa nhận: đọc tab chính (khớp với bảng "Xác thực hóa đơn" trên sheet tổng)
  const stSh = ss.getSheetByName(cfg.statusSheet) || sh;
  if (stSh && stSh.getLastRow() > 1) {
    backlog.statusSheet = stSh.getName();
    const values = stSh.getRange(1, 1, stSh.getLastRow(), stSh.getLastColumn()).getValues();
    const c = pnCols(vtdApp_headerMap_(values[0]));
    if (c.status < 0) throw new Error("Tab " + stSh.getName() + " không có cột Xác thực hóa đơn");
    const todayMs = dash_dayMs_(now);
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const get = i => i >= 0 ? row[i] : "";
      const ma = String(get(c.ma) || get(c.od) || "");
      if (!ma) continue;
      const status = String(get(c.status) || "").trim();
      const statusLabel = status || "(trống)";
      backlog.byStatus[statusLabel] = (backlog.byStatus[statusLabel] || 0) + 1;
      if (cfg.openStatuses.indexOf(vtdApp_norm_(status)) < 0) continue;

      const store = String(get(c.store) || "Khác").trim() || "Khác";
      const lenKey = vtdApp_dateKey_(get(c.lenDon));
      const days = lenKey ? Math.max(0, Math.round((todayMs - dash_dayMs_(lenKey)) / 86400000)) : -1;
      backlog.total++;
      backlog.byStore[store] = (backlog.byStore[store] || 0) + 1;
      if (days < 0) continue;            // thiếu Ngày lên đơn: vẫn đếm vào tổng, không xếp tuổi
      if (days > 7) backlog.over7++;
      backlog.buckets[days <= 3 ? 0 : days <= 7 ? 1 : days <= 15 ? 2 : 3]++;
      backlog.oldest.push({ma: ma, khach: String(get(c.khach) || ""), kenh: store, ngay: lenKey, days: days, status: status});
    }
    backlog.oldest.sort((a, b) => b.days - a.days);
    backlog.oldest = backlog.oldest.slice(0, 12);
  }

  // 3) Hoàn sản phẩm: mỗi dòng = 1 sản phẩm; qty = số lượng; Tình trạng FF
  const prod = ss.getSheetByName(cfg.productSheet);
  if (prod && prod.getLastRow() > 1) {
    const values = prod.getRange(1, 1, prod.getLastRow(), prod.getLastColumn()).getValues();
    const col = vtdApp_headerMap_(values[0]);
    const c = {
      date: dash_col_(col, ["ngay hoan tra", "ngay"], -1),
      time: dash_col_(col, ["thoi gian", "gio"], -1),
      user: dash_col_(col, ["user thao tac", "user"], -1),
      qty: dash_col_(col, ["so luong"], -1),
      name: dash_col_(col, ["ten san pham", "ten vat tu"], -1),
      status: dash_col_(col, ["tinh trang ff", "tinh trang"], -1)
    };
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const dateKey = c.date >= 0 ? vtdApp_dateKey_(row[c.date]) : "";
      if (!dateKey) continue;
      const qty = Number(c.qty >= 0 ? row[c.qty] : 0) || 0;
      add("PN", dateKey, dash_hour_(c.time >= 0 ? row[c.time] : ""), "Hàng hoàn PN", "Sản phẩm",
        dash_user_(c.user >= 0 ? row[c.user] : ""), qty, null);
      const status = String(c.status >= 0 ? row[c.status] : "").trim() || "(trống)";
      const name = String(c.name >= 0 ? row[c.name] : "").trim();
      if (name) metric("product", "PN", dateKey, name, status, qty);
      metric("cond", "PN", dateKey, "Sản phẩm", status, qty);
    }
  }
  return backlog;
}

/* ============================ Nguồn TH ============================ */

function dash_collectTh_(add, metric) {
  const cfg = DASH_CONFIG.th;
  if (!cfg.spreadsheetId || !cfg.sheetName) return;
  const sh = SpreadsheetApp.openById(cfg.spreadsheetId).getSheetByName(cfg.sheetName);
  if (!sh) throw new Error("Không thấy tab " + cfg.sheetName);
  const headerRow = Math.max(1, Number(cfg.headerRow) || 1);
  if (sh.getLastRow() <= headerRow) return;
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const col = vtdApp_headerMap_(values[headerRow - 1]);
  const c = {
    date: dash_col_(col, cfg.dateAliases, -1),
    time: dash_col_(col, cfg.timeAliases, -1),
    user: dash_col_(col, cfg.userAliases, -1),
    qty: dash_col_(col, cfg.qtyAliases, -1),
    ma: dash_col_(col, cfg.maAliases, -1),
    kenh: dash_col_(col, cfg.kenhAliases, -1),
    name: dash_col_(col, cfg.nameAliases, -1),
    status: dash_col_(col, cfg.statusAliases, -1)
  };
  if (c.date < 0 && c.time < 0) throw new Error("Tab " + cfg.sheetName + " thiếu cột Ngày tháng/Thời gian ở dòng " + headerRow);
  for (let r = headerRow; r < values.length; r++) {
    const row = values[r];
    const dateKey = vtdApp_dateKey_(c.date >= 0 ? row[c.date] : "") || vtdApp_dateKey_(c.time >= 0 ? row[c.time] : "");
    if (!dateKey) continue;
    const user = dash_user_(c.user >= 0 ? row[c.user] : "");
    const kenh = String(c.kenh >= 0 ? row[c.kenh] : "").trim() || "TH Truemilk";
    const qty = Number(c.qty >= 0 ? row[c.qty] : 0) || 0;
    add("TH", dateKey, dash_hour_(c.time >= 0 ? row[c.time] : ""), kenh, "Hàng hoàn", user, qty, {
      at: dateKey + " " + dash_hhmm_(c.time >= 0 ? row[c.time] : ""), src: "TH", ma: String(c.ma >= 0 ? row[c.ma] || "" : ""), kenh: kenh, loai: "Hàng hoàn", user: user
    });
    const name = String(c.name >= 0 ? row[c.name] : "").trim();
    const status = c.status >= 0 ? (String(row[c.status] || "").trim() || "(trống)") : "";
    if (name) metric("product", "TH", dateKey, name, status || "(không ghi tình trạng)", qty);
    if (status) metric("cond", "TH", dateKey, "Sản phẩm", status, qty);
  }
}

// DS đơn all: mỗi dòng 1 SKU của đơn → gộp theo mã đơn, ngày = Thời gian tạo
function dash_collectThOrders_(metric) {
  const cfg = DASH_CONFIG.th;
  if (!cfg.spreadsheetId || !cfg.ordersSheet) return;
  const sh = SpreadsheetApp.openById(cfg.spreadsheetId).getSheetByName(cfg.ordersSheet);
  if (!sh) throw new Error("Không thấy tab " + cfg.ordersSheet);
  const headerRow = Math.max(1, Number(cfg.ordersHeaderRow) || 1);
  if (sh.getLastRow() <= headerRow) return;
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const col = vtdApp_headerMap_(values[headerRow - 1]);
  const c = {
    date: dash_col_(col, ["thoi gian tao", "ngay tao", "ngay len don"], -1),
    ids: [dash_col_(col, ["ma don ghtk"], -1), dash_col_(col, ["ma od"], -1), dash_col_(col, ["ma don gam"], -1), dash_col_(col, ["ma don doi tac"], -1)].filter(i => i >= 0)
  };
  if (c.date < 0) throw new Error("Tab " + cfg.ordersSheet + " thiếu cột Thời gian tạo");
  const seen = {};
  for (let r = headerRow; r < values.length; r++) {
    const row = values[r];
    const id = c.ids.map(i => String(row[i] || "").trim()).filter(Boolean)[0] || ("row" + r);
    if (seen[id]) continue;
    seen[id] = true;
    metric("order", "TH", vtdApp_dateKey_(row[c.date]), "TH Truemilk", "Lên đơn", 1);
  }
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

function dash_hour_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Number(Utilities.formatDate(value, DASH_CONFIG.timezone, "H"));
  }
  const m = String(value == null ? "" : value).match(/(\d{1,2}):(\d{2})/);
  return m ? Math.min(23, Number(m[1])) : -1;
}

function dash_hhmm_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, DASH_CONFIG.timezone, "HH:mm");
  const m = String(value == null ? "" : value).match(/(\d{1,2}):(\d{2})/);
  return m ? ("0" + m[1]).slice(-2) + ":" + m[2] : "";
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
