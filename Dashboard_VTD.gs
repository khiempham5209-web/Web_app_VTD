/**
 * Dashboard_VTD.gs — tổng hợp số liệu cho trang /dashboard.html
 *
 * Thêm file này vào project Apps Script VTD (cùng project với Code_chinh_VTD.gs).
 * Cài đặt:
 *   1. Điền DASH_CONFIG.th (sheet TH Truemilk) nếu muốn có số TH.
 *   2. Thêm 1 dòng vào map trong vtdApp_apiDispatch_:   dashboardV2: vtdApp_dashboardV2,
 *      Quyền xem: admin luôn xem được; nhân viên cần được tick màn "Dashboard" (id dashview).
 *   3. Chạy tay dashBuild() một lần (cấp quyền), rồi chạy dashInstallTrigger().
 *   4. Deploy lại Web App (Manage deployments → Edit → New version) để giữ nguyên URL.
 *
 * Trang dashboard chỉ đọc 2 sheet ẩn _DASH_FACT và _DASH_SNAPSHOT, không quét
 * RAW_IN_FULL mỗi lần mở, nên không làm chậm app nhập liệu.
 */

const DASH_CONFIG = {
  days: 130,               // số ngày lịch sử (đủ cho 90 ngày + so sánh cùng kỳ tháng trước)
  factSheet: "_DASH_FACT",
  snapshotSheet: "_DASH_SNAPSHOT",
  timezone: "Asia/Ho_Chi_Minh",
  recentLimit: 30,
  vtd: {
    // RAW_IN_FULL chứa toàn bộ lịch sử; RAW_IN_M3 dùng cho trạng thái đồng bộ
    historySheet: "RAW_IN_FULL",
    liveSheet: "RAW_IN_M3"
  },
  pn: {
    spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
    sheetNames: ["Chứng từ_full", "Chứng từ_FF"],   // dùng sheet đầu tiên tồn tại
    productSheet: "Hoàn sản phẩm",
    doneStatuses: ["da nhan chung tu", "shop huy od"]
  },
  th: {
    // Tab "Hàng hoàn" của file TH Truemilk (cùng file với script FFPL tồn kho). Để trống ID = bỏ qua TH.
    spreadsheetId: "1TVck35ibO3zJ6lkKpHNyXckgLyjvUxEoLbMa-6kt-Pk",
    sheetName: "Hàng hoàn",
    headerRow: 3,          // tiêu đề nằm ở dòng 3 (giống FFPL.headerRows)
    dateAliases: ["ngay thang", "ngay hoan", "ngay hoan tra", "ngay", "date"],
    timeAliases: ["thoi gian", "gio", "time", "timestamp"],
    userAliases: ["user thao tac", "user", "nguoi nhap", "nguoi thao tac", "email"],
    qtyAliases: ["so luong", "sl", "qty"],
    maAliases: ["ma don", "ma don ghtk"],
    kenhAliases: ["kho", "kenh", "slog"]
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
  const facts = last > 1 ? factSh.getRange(2, 1, last - 1, 8).getValues().map(r => [
    r[0] instanceof Date ? Utilities.formatDate(r[0], DASH_CONFIG.timezone, "yyyy-MM-dd") : String(r[0]), String(r[1]), String(r[2]), String(r[3]), String(r[4]), Number(r[5]) || 0, Number(r[6]) || 0, Number(r[7]) || 0
  ]) : [];
  let snapshot = {};
  try { snapshot = JSON.parse(String(snapSh.getRange(2, 2).getValue() || "{}")); } catch (err) { snapshot = {}; }
  return vtdApp_ok_({
    updatedAt: String(snapSh.getRange(2, 1).getDisplayValue() || ""),
    today: Utilities.formatDate(new Date(), DASH_CONFIG.timezone, "yyyy-MM-dd"),
    viewer: email,
    facts: facts,
    snapshot: snapshot
  });
}

// Quyền riêng: màn "dashview" (admin tick trong Admin → Người dùng → Màn được sử dụng).
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
    const fromKey = Utilities.formatDate(new Date(now.getTime() - DASH_CONFIG.days * 86400000), tz, "yyyy-MM-dd");
    const facts = {};          // key -> [date, src, kenh, loai, user, hour, count, qty]
    const recent = [];
    const freshness = {};
    const errors = [];

    const add = (src, dateKey, hour, kenh, loai, user, qty, recentRow) => {
      if (!dateKey) return;
      if (!freshness[src] || dateKey > freshness[src]) freshness[src] = dateKey;
      if (dateKey < fromKey) return;
      const k = [dateKey, src, kenh, loai, user, hour].join("|");
      const f = facts[k] || (facts[k] = [dateKey, src, kenh, loai, user, hour, 0, 0]);
      f[6] += 1;
      f[7] += qty;
      if (recentRow) recent.push(recentRow);
    };

    let sync = {};
    try { dash_collectVtd_(add); } catch (err) { errors.push("VTD: " + err); }
    try { sync = dash_vtdSync_(); } catch (err) { errors.push("Đồng bộ VTD: " + err); }
    let pnBacklog = null;
    try { pnBacklog = dash_collectPn_(add, now); } catch (err) { errors.push("Phạm Nguyên: " + err); }
    try { dash_collectTh_(add); } catch (err) { errors.push("TH: " + err); }
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
      recent: recent.slice(0, DASH_CONFIG.recentLimit),
      errors: errors,
      thConfigured: !!(DASH_CONFIG.th.spreadsheetId && DASH_CONFIG.th.sheetName)
    };

    const ss = vtdApp_ss_();
    const factSh = dash_sheet_(ss, DASH_CONFIG.factSheet, ["Ngày", "Nguồn", "Kênh", "Loại", "User", "Giờ", "Số thao tác", "Số lượng"]);
    const rows = Object.keys(facts).map(k => facts[k]).sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
    // Cột Ngày để dạng chữ, không cho Sheets tự đổi "2026-10-16" thành kiểu Ngày.
    factSh.getRange("A:A").setNumberFormat("@");
    // Ghi dữ liệu mới trước, xóa phần đuôi cũ sau — không có lúc sheet trống.
    if (rows.length) factSh.getRange(2, 1, rows.length, 8).setValues(rows);
    const extra = factSh.getLastRow() - 1 - rows.length;
    if (extra > 0) factSh.getRange(rows.length + 2, 1, extra, 8).clearContent();

    // Tóm tắt trong Execution log để tự kiểm tra sau mỗi lần chạy tay
    const bySrc = {};
    rows.forEach(r => bySrc[r[1]] = (bySrc[r[1]] || 0) + r[6]);
    console.log("dashBuild: " + rows.length + " dòng tổng hợp; thao tác theo nguồn " + JSON.stringify(bySrc)
      + "; dữ liệu mới nhất " + JSON.stringify(freshness) + (errors.length ? "; LỖI: " + errors.join(" | ") : ""));

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

function dash_collectVtd_(add) {
  const ss = vtdApp_ss_();
  const sh = ss.getSheetByName(DASH_CONFIG.vtd.historySheet) || ss.getSheetByName(DASH_CONFIG.vtd.liveSheet);
  if (!sh || sh.getLastRow() < 2) return;
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const col = vtdApp_headerMap_(values[0]);
  // Cùng alias + vị trí dự phòng với vtdApp_compactRecord_
  const c = {
    date: dash_col_(col, ["ngay hoan tra", "ngay nhap"], 1),
    time: dash_col_(col, ["thoi gian nhap hoan", "thoi gian", "timestamp"], 2),
    ma: dash_col_(col, ["ma don", "ma don ghn", "ma don ghtk", "madon"], 3),
    slog: dash_col_(col, ["slog"], 4),
    loai: dash_col_(col, ["loai hoan", "loai st"], 6),
    user: dash_col_(col, ["user_email", "user email", "email", "nguoi nhap", "user thao tac", "created by"], 11)
  };
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateKey = vtdApp_dateKey_(row[c.date]) || vtdApp_dateKey_(row[c.time]);
    if (!dateKey) continue;
    const hour = dash_hour_(row[c.time]);
    const user = dash_user_(row[c.user]);
    const kenh = String(row[c.slog] || "Khác").trim() || "Khác";
    const loai = String(row[c.loai] || "Khác").trim() || "Khác";
    add("VTD", dateKey, hour, kenh, loai, user, 1, {
      at: dateKey + " " + dash_hhmm_(row[c.time]), src: "VTD", ma: String(row[c.ma] || ""), kenh: kenh, loai: loai, user: user
    });
  }
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

function dash_collectPn_(add, now) {
  const cfg = DASH_CONFIG.pn;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  let sh = null;
  for (let i = 0; i < cfg.sheetNames.length && !sh; i++) sh = ss.getSheetByName(cfg.sheetNames[i]);
  const backlog = {total: 0, over7: 0, buckets: [0, 0, 0, 0], byStore: {}, byStatus: {}, openByStatus: {}, oldest: []};
  if (sh && sh.getLastRow() > 1) {
    const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const col = vtdApp_headerMap_(values[0]);
    const c = {
      ma: dash_col_(col, ["ma don ghtk", "ma don"], -1),
      od: dash_col_(col, ["so don hang", "od"], -1),
      khach: dash_col_(col, ["khach hang"], -1),
      store: dash_col_(col, ["loai sieu thi"], -1),
      status: dash_col_(col, ["trang thai"], -1),
      time: dash_col_(col, ["thoi gian"], -1),
      user: dash_col_(col, ["user thao tac"], -1),
      lenDon: dash_col_(col, ["ngay len don"], -1)
    };
    const todayMs = dash_dayMs_(now);
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const get = i => i >= 0 ? row[i] : "";
      const store = String(get(c.store) || "Khác").trim() || "Khác";
      const status = String(get(c.status) || "").trim();
      const statusKey = vtdApp_norm_(status);
      const ma = String(get(c.ma) || get(c.od) || "");
      if (!ma) continue;
      backlog.byStatus[status || "Chưa có trạng thái"] = (backlog.byStatus[status || "Chưa có trạng thái"] || 0) + 1;

      // Thao tác chứng từ (có user + thời gian thao tác)
      const user = dash_user_(get(c.user));
      const actKey = vtdApp_dateKey_(get(c.time));
      if (actKey && get(c.user)) {
        add("PN", actKey, dash_hour_(get(c.time)), store, "Chứng từ", user, 1, {
          at: actKey + " " + dash_hhmm_(get(c.time)), src: "PN", ma: ma, kenh: store, loai: "Chứng từ", user: user
        });
      }

      // Tồn chứng từ chưa nhận
      if (cfg.doneStatuses.indexOf(statusKey) >= 0) continue;
      const lenKey = vtdApp_dateKey_(get(c.lenDon));
      if (!lenKey) continue;
      const days = Math.max(0, Math.round((todayMs - dash_dayMs_(lenKey)) / 86400000));
      backlog.total++;
      const openKey = status || "(trống)";
      backlog.openByStatus[openKey] = (backlog.openByStatus[openKey] || 0) + 1;
      if (days > 7) backlog.over7++;
      backlog.buckets[days <= 3 ? 0 : days <= 7 ? 1 : days <= 15 ? 2 : 3]++;
      backlog.byStore[store] = (backlog.byStore[store] || 0) + 1;
      backlog.oldest.push({ma: ma, khach: String(get(c.khach) || ""), kenh: store, ngay: lenKey, days: days, status: status});
    }
    backlog.oldest.sort((a, b) => b.days - a.days);
    backlog.oldest = backlog.oldest.slice(0, 12);
  }

  // Hoàn sản phẩm: mỗi dòng = 1 sản phẩm; qty = số lượng
  const prod = ss.getSheetByName(cfg.productSheet);
  if (prod && prod.getLastRow() > 1) {
    const values = prod.getRange(1, 1, prod.getLastRow(), prod.getLastColumn()).getValues();
    const col = vtdApp_headerMap_(values[0]);
    const c = {
      date: dash_col_(col, ["ngay hoan tra", "ngay"], -1),
      time: dash_col_(col, ["thoi gian", "gio"], -1),
      user: dash_col_(col, ["user thao tac", "user"], -1),
      qty: dash_col_(col, ["so luong"], -1)
    };
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const dateKey = c.date >= 0 ? vtdApp_dateKey_(row[c.date]) : "";
      if (!dateKey) continue;
      add("PN", dateKey, dash_hour_(c.time >= 0 ? row[c.time] : ""), "Hàng hoàn PN", "Sản phẩm",
        dash_user_(c.user >= 0 ? row[c.user] : ""), Number(c.qty >= 0 ? row[c.qty] : 0) || 0, null);
    }
  }
  return backlog;
}

/* ============================ Nguồn TH ============================ */

function dash_collectTh_(add) {
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
    kenh: dash_col_(col, cfg.kenhAliases, -1)
  };
  if (c.date < 0 && c.time < 0) throw new Error("Tab " + cfg.sheetName + " thiếu cột Ngày tháng/Thời gian ở dòng " + headerRow);
  for (let r = headerRow; r < values.length; r++) {
    const row = values[r];
    const dateKey = vtdApp_dateKey_(c.date >= 0 ? row[c.date] : "") || vtdApp_dateKey_(c.time >= 0 ? row[c.time] : "");
    if (!dateKey) continue;
    const user = dash_user_(c.user >= 0 ? row[c.user] : "");
    const kenh = String(c.kenh >= 0 ? row[c.kenh] : "").trim() || "TH Truemilk";
    add("TH", dateKey, dash_hour_(c.time >= 0 ? row[c.time] : ""), kenh, "Hàng hoàn", user,
      Number(c.qty >= 0 ? row[c.qty] : 0) || 0, {
        at: dateKey + " " + dash_hhmm_(c.time >= 0 ? row[c.time] : ""), src: "TH", ma: String(c.ma >= 0 ? row[c.ma] || "" : ""), kenh: kenh, loai: "Hàng hoàn", user: user
      });
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
