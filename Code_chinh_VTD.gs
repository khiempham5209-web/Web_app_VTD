/************************************************
VTD INTERNAL WEB APP

Paste this file into the same Apps Script project that already contains
the PJ1, PJ2, dispatcher, and full data layer code.
************************************************/

const VTD_APP = {
  spreadsheetId: "",
  rawSheet: "RAW_IN_M3",
  rawFullSheet: "RAW_IN_FULL",
  imageSheet: "Ảnh_chứng_từ",
  imageSheetAlt: "Anh_chứng_từ",
  imageFullSheet: "Full_ảnh_chứng_từ",
  skuSheet: "List_SKU_hoàn",
  logSheet: "SYNC_LOG",
  permissionSheet: "VTD_USERS",
  systemConfigSheet: "_VTD_APP_SYSTEM_CONFIG",
  forgotPassSheet: "_VTD_FORGOT_PASS_REQUESTS",
  loginReportSheet: "_VTD_LOGIN_REPORT",
  storeInfoSheets: ["Search", "Slog_search", "Bảng tổng hợp 1", "SLOG"],
  imageRootFolderNames: ["Ảnh_chứng_từ", "Ảnh_chứng_từ_Images", "Anh_chứng_từ", "CHUNG_TU_HE_THONG"],
  maxSearchRows: 5000,
  dashboardRows: 300,
  defaultLimit: 80,
  requireGoogleLogin: false,
  requirePinLogin: true,
  adminBootstrapPin: "5209",
  googleClientId: ""
};

const VTD_PERMISSIONS = {
  adminEmails: [
    "khiempham5209@gmail.com",
    "fulfillment.wms.3pl@gmail.com"
  ],
  users: {
    "ntlamlogistics@gmail.com": {
      role: "user",
      screens: ["home", "search", "input", "store"],
      actions: ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"]
    },
    "letuyetthanh88@gmail.com": {
      role: "user",
      screens: ["home", "search", "input", "store"],
      actions: ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"]
    },
    "lananhdo940@gmail.com": {
      role: "user",
      screens: ["home", "search", "input", "store"],
      actions: ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"]
    },
    "tranhong1997121811@gmail.com": {
      role: "user",
      screens: ["home", "search", "input", "store"],
      actions: ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"]
    }
  },
  allScreens: ["home", "search", "input", "returnth", "store", "settings", "ops", "admin"],
  allActions: ["dashboard", "search", "getRecord", "saveRaw", "saveReturnTh", "updateRecord", "uploadImage", "queueRecord", "storeInfo", "changePass", "changePin", "runCommand", "adminConfig", "manageConfig", "retrySync", "activityPing", "loginReport"]
};

const VTD_STATUS_LABELS = {
  DONE: "Đã đồng bộ",
  YES: "Đang đồng bộ",
  RUNNING: "Đang đồng bộ",
  ROUTE_PJ2: "Đã đồng bộ",
  ERROR: "Lỗi",
  IMG_ERROR: "Lỗi",
  PENDING: "Đang đồng bộ",
  BLANK: "Lỗi"
};

function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Chứng từ VTD")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function vtdApp_publicConfig() {
  return vtdApp_ok_({
    appVersion: "vtd-2026-06-05-cache-01",
    googleClientId: VTD_APP.googleClientId,
    requireGoogleLogin: !!VTD_APP.requireGoogleLogin,
    requirePinLogin: !!VTD_APP.requirePinLogin,
    authMode: VTD_APP.googleClientId ? "google_identity" : "pin"
  });
}

function vtdApp_login(params) {
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  const pin = String(params.pin || "").trim();
  if (!email || !pin) return vtdApp_fail_("Nhập email và Pass.");

  const auth = vtdApp_authByPin_(email, pin);
  if (!auth.allowed) return vtdApp_fail_("Bạn không có quyền truy cập.");

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("VTD_SESSION_" + token, JSON.stringify({
    email: auth.email,
    createdAt: Date.now()
  }), 21600);

  return vtdApp_ok_({
    token: token,
    user: auth.email,
    auth: auth
  });
}

function vtdApp_logout(params) {
  params = params || {};
  if (params.sessionToken) CacheService.getScriptCache().remove("VTD_SESSION_" + params.sessionToken);
  return vtdApp_ok_({message: "Đã đăng xuất."});
}

function vtdApp_changePin(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_("Bạn không có quyền truy cập.");
  params = params || {};
  const oldPin = String(params.oldPin || "").trim();
  const newPin = String(params.newPin || "").trim();
  if (!newPin || newPin.length < 4) return vtdApp_fail_("Pass mới phải có ít nhất 4 ký tự.");
  if (!vtdApp_authByPin_(auth.email, oldPin).allowed) return vtdApp_fail_("Pass cũ không đúng.");

  vtdApp_updateUserPass_(auth.email, newPin, auth.email);
  vtdApp_clearUserAuthCache_();
  return vtdApp_ok_({message: "Đã đổi Pass."});
}

function vtdApp_forgotPass(params) {
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  if (!email || email.indexOf("@") < 0) return vtdApp_fail_("Nhập email cần lấy lại Pass.");

  const config = vtdApp_permissionConfig_();
  const mainAdmin = "khiempham5209@gmail.com";
  const adminEmails = config.adminEmails.map(e => String(e).toLowerCase().trim());
  let role = "";
  let pass = "";
  let note = "";

  if (adminEmails.indexOf(email) >= 0) {
    role = "admin";
    pass = config.pins[email] || VTD_APP.adminBootstrapPin;
    note = "Tài khoản admin.";
  } else if (config.users[email]) {
    role = config.users[email].role || "staff";
    pass = config.pins[email] || "";
    note = pass ? "Tài khoản đã được cấp quyền." : "Tài khoản có quyền nhưng chưa có Pass.";
  } else {
    note = "Email này chưa có trong danh sách được cấp quyền.";
  }

  const sh = vtdApp_forgotPassSheet_();
  const requestId = "FP-" + Utilities.formatDate(new Date(), "Asia/Saigon", "yyyyMMdd-HHmmss") + "-" + Utilities.getUuid().slice(0, 8);
  sh.appendRow([requestId, email, "PENDING", new Date(), "", "", "", role || "", note]);

  const subject = "[VTD] Yêu cầu quên Pass - " + email;
  const body = [
    "Có yêu cầu quên Pass từ web Chứng từ VTD.",
    "",
    "Email: " + email,
    "Vai trò: " + (role || "chưa cấp quyền"),
    "Pass hiện tại: " + (pass || "(chưa có / cần reset trong Admin)"),
    "Ghi chú: " + note,
    "",
    "Nếu cần đổi Pass, vào màn Admin của web để sửa hoặc reset lại."
  ].join("\n");

  try {
    MailApp.sendEmail(mainAdmin, subject, body);
  } catch (err) {
    return vtdApp_fail_("Chưa gửi được mail báo admin: " + err.message);
  }

  return vtdApp_ok_({message: "Đã gửi yêu cầu về mail admin. Vui lòng chờ admin kiểm tra/reset Pass."});
}

function vtdApp_forgotPassRequests(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  const sh = vtdApp_forgotPassSheet_();
  const values = sh.getDataRange().getDisplayValues();
  if (values.length <= 1) return vtdApp_ok_({rows: [], pending: 0});
  const col = vtdApp_headerMap_(values[0]);
  const rows = values.slice(1).map((row, i) => ({
    rowNumber: i + 2,
    requestId: row[col.requestid] || row[col["request id"]] || "",
    email: row[col.email] || "",
    status: row[col.status] || "",
    requestedAt: row[col.requestedat] || row[col["requested at"]] || "",
    resolvedAt: row[col.resolvedat] || row[col["resolved at"]] || "",
    resolvedBy: row[col.resolvedby] || row[col["resolved by"]] || "",
    newPass: row[col.newpass] || row[col["new pass"]] || "",
    role: row[col.role] || "",
    note: row[col.note] || ""
  })).filter(r => r.email).reverse();
  return vtdApp_ok_({
    rows: rows,
    pending: rows.filter(r => String(r.status || "").toUpperCase() === "PENDING").length
  });
}

function vtdApp_resetForgotPass(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  const requestId = String(params.requestId || "").trim();
  let newPass = String(params.newPass || "").trim();
  if (!email || email.indexOf("@") < 0) return vtdApp_fail_("Thiếu email cần reset Pass.");
  if (!newPass) newPass = String(Math.floor(100000 + Math.random() * 900000));
  if (newPass.length < 4) return vtdApp_fail_("Pass mới phải có ít nhất 4 ký tự.");

  const auth = vtdApp_auth_(params);
  const updatedBy = auth && auth.email || String(params.userEmail || "");
  vtdApp_updateUserPass_(email, newPass, updatedBy);
  vtdApp_clearUserAuthCache_();

  const sh = vtdApp_forgotPassSheet_();
  const values = sh.getDataRange().getValues();
  const col = vtdApp_headerMap_(values[0]);
  let updated = false;
  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][col.email] || "").toLowerCase().trim();
    const rowReq = String(values[i][col.requestid] || values[i][col["request id"]] || "").trim();
    const status = String(values[i][col.status] || "").toUpperCase();
    if (rowEmail === email && (!requestId || rowReq === requestId) && status !== "DONE") {
      vtdApp_setByAliases_(values[i], col, ["status"], "DONE");
      vtdApp_setByAliases_(values[i], col, ["resolved at"], new Date());
      vtdApp_setByAliases_(values[i], col, ["resolved by"], updatedBy);
      vtdApp_setByAliases_(values[i], col, ["new pass"], newPass);
      sh.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
      updated = true;
      if (requestId) break;
    }
  }

  try {
    MailApp.sendEmail(
      email,
      "[VTD] Pass mới của bạn",
      [
        "Admin đã reset Pass tài khoản Chứng từ VTD.",
        "",
        "Email: " + email,
        "Pass mới: " + newPass,
        "",
        "Vui lòng đăng nhập lại và đổi Pass nếu cần."
      ].join("\n")
    );
  } catch (err) {
    return vtdApp_fail_("Đã reset Pass nhưng chưa gửi được mail cho user: " + err.message);
  }

  return vtdApp_ok_({message: "Đã reset Pass và gửi mail cho user.", email: email, newPass: newPass, updatedRequest: updated});
}

function vtdApp_forgotPassSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.forgotPassSheet);
  const header = ["RequestId", "Email", "Status", "RequestedAt", "ResolvedAt", "ResolvedBy", "NewPass", "Role", "Note"];
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.forgotPassSheet);
    sh.appendRow(header);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow(header);
  }
  return sh;
}

function vtdApp_init(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const passwordPolicy = vtdApp_markPasswordFirstLogin_(auth.email, String((params || {}).appVersion || "").trim());
  return vtdApp_ok_({
    user: auth.email,
    auth: Object.assign({}, auth, {passwordPolicy: passwordPolicy}),
    passwordPolicy: passwordPolicy,
    config: {
      rawSheet: VTD_APP.rawSheet,
      rawFullSheet: VTD_APP.rawFullSheet,
      imageSheet: VTD_APP.imageSheet,
      skuSheet: VTD_APP.skuSheet,
      system: vtdApp_publicSystemConfig_()
    }
  });
}

function vtdApp_systemConfig(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  return vtdApp_ok_({config: vtdApp_publicSystemConfig_()});
}

function vtdApp_saveSystemConfig(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const input = params.config || {};
  const config = vtdApp_systemConfig_();
  const changed = {};
  if (Object.prototype.hasOwnProperty.call(input, "returnThApiUrl")) {
    const url = String(input.returnThApiUrl || "").trim();
    if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[-\w]+\/exec/i.test(url)) {
      return vtdApp_fail_("Link API Nhập TH không hợp lệ.");
    }
    config.returnThApiUrl = url;
    changed.returnThApiUrl = config.returnThApiUrl;
  }
  if (Object.prototype.hasOwnProperty.call(input, "docOpsApiUrl")) {
    const url = String(input.docOpsApiUrl || "").trim();
    if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[-\w]+\/exec/i.test(url)) {
      return vtdApp_fail_("Link API Thao tac CT khong hop le.");
    }
    config.docOpsApiUrl = url;
    changed.docOpsApiUrl = config.docOpsApiUrl;
  }
  if (Object.prototype.hasOwnProperty.call(input, "uiConfig")) {
    let uiConfig = input.uiConfig || {};
    if (typeof uiConfig === "string") {
      try {
        uiConfig = JSON.parse(uiConfig);
      } catch (err) {
        return vtdApp_fail_("Cấu hình UI không hợp lệ.");
      }
    }
    const cleanUiConfig = {
      views: Array.isArray(uiConfig.views) ? uiConfig.views : [],
      storeCharts: uiConfig.storeCharts && typeof uiConfig.storeCharts === "object" ? uiConfig.storeCharts : {},
      personal: uiConfig.personal && typeof uiConfig.personal === "object" ? uiConfig.personal : {}
    };
    if (uiConfig.update && typeof uiConfig.update === "object") {
      cleanUiConfig.update = vtdApp_cleanUpdateConfig_(uiConfig.update);
      config.update = JSON.stringify(cleanUiConfig.update);
      changed.update = config.update;
    }
    config.uiConfig = JSON.stringify(cleanUiConfig);
    changed.uiConfig = config.uiConfig;
  }
  if (Object.prototype.hasOwnProperty.call(input, "update")) {
    if (!input.update || typeof input.update !== "object") return vtdApp_fail_("Cấu hình cập nhật không hợp lệ.");
    config.update = JSON.stringify(vtdApp_cleanUpdateConfig_(input.update));
    changed.update = config.update;
  }
  if (Object.prototype.hasOwnProperty.call(input, "themeConfig")) {
    let themeConfig = input.themeConfig || {};
    if (typeof themeConfig === "string") {
      try {
        themeConfig = JSON.parse(themeConfig);
      } catch (err) {
        return vtdApp_fail_("Cau hinh theme khong hop le.");
      }
    }
    if (!themeConfig || typeof themeConfig !== "object") return vtdApp_fail_("Cau hinh theme khong hop le.");
    let previousThemeConfig = {};
    if (config.themeConfig) {
      try { previousThemeConfig = JSON.parse(config.themeConfig); } catch (err) {}
    }
    config.themeConfig = JSON.stringify(vtdApp_sanitizeThemeConfig_(themeConfig, previousThemeConfig));
    changed.themeConfig = config.themeConfig;
  }
  vtdApp_saveSystemConfig_(changed, params && params.userEmail);
  return vtdApp_ok_({message: "Đã lưu cấu hình hệ thống.", config: vtdApp_publicSystemConfig_(config)});
}

function vtdApp_publicSystemConfig_(config) {
  config = config || vtdApp_systemConfig_();
  let uiConfig = {};
  if (config.uiConfig) {
    try {
      uiConfig = JSON.parse(config.uiConfig);
    } catch (err) {
      uiConfig = {};
    }
  }
  let update = {};
  if (config.update) {
    try {
      update = JSON.parse(config.update);
    } catch (err) {
      update = {};
    }
  } else if (uiConfig.update && typeof uiConfig.update === "object") {
    update = uiConfig.update;
  }
  let themeConfig = {};
  if (config.themeConfig) {
    try {
      themeConfig = vtdApp_sanitizeThemeConfig_(JSON.parse(config.themeConfig), {});
    } catch (err) {
      themeConfig = {};
    }
  }
  const externalThemeImage = vtdApp_driveThemeUrl_(config.themeImageUrl || "");
  if (externalThemeImage.url) {
    themeConfig.loginAppearance = themeConfig.loginAppearance && typeof themeConfig.loginAppearance === "object" ? themeConfig.loginAppearance : {};
    themeConfig.loginAppearance.hasImage = true;
    themeConfig.loginAppearance.imageUrl = externalThemeImage.url;
    themeConfig.assets = themeConfig.assets && typeof themeConfig.assets === "object" ? themeConfig.assets : {};
    const currentLoginAsset = themeConfig.assets.login && typeof themeConfig.assets.login === "object" ? themeConfig.assets.login : {};
    themeConfig.assets.login = {
      version: String(currentLoginAsset.version || themeConfig.loginAppearance.imageVersion || ("drive_" + externalThemeImage.fileId)),
      url: externalThemeImage.url,
      fileId: externalThemeImage.fileId
    };
  }
  return {
    returnThApiUrl: String(config.returnThApiUrl || ""),
    docOpsApiUrl: String(config.docOpsApiUrl || ""),
    uiConfig: uiConfig,
    update: update,
    themeConfig: themeConfig
  };
}

function vtdApp_driveThemeUrl_(value) {
  const raw = String(value || "").trim();
  if (!raw) return {url: "", fileId: ""};
  let fileId = "";
  const pathMatch = raw.match(/drive\.google\.com\/file\/d\/([\w-]+)/i);
  const idMatch = raw.match(/[?&]id=([\w-]+)/i);
  if (pathMatch) fileId = pathMatch[1];
  else if (idMatch) fileId = idMatch[1];
  else if (/^[\w-]{20,}$/.test(raw)) fileId = raw;
  if (fileId) {
    return {
      url: "https://drive.usercontent.google.com/download?id=" + encodeURIComponent(fileId) + "&export=download",
      fileId: fileId
    };
  }
  if (/^https:\/\//i.test(raw)) return {url: raw, fileId: ""};
  return {url: "", fileId: ""};
}

function vtdApp_sanitizeThemeAsset_(asset, previousAsset, enabled) {
  asset = asset && typeof asset === "object" ? asset : {};
  previousAsset = previousAsset && typeof previousAsset === "object" ? previousAsset : {};
  const requested = vtdApp_driveThemeUrl_(asset.url || asset.imageUrl || asset.driveUrl || asset.fileId || "");
  const previous = vtdApp_driveThemeUrl_(previousAsset.url || previousAsset.imageUrl || previousAsset.driveUrl || previousAsset.fileId || "");
  const selected = requested.url ? requested : (enabled !== false ? previous : {url: "", fileId: ""});
  const out = {version: String(asset.version || previousAsset.version || ("asset_" + Date.now()))};
  if (selected.url) {
    out.url = selected.url;
    if (selected.fileId) out.fileId = selected.fileId;
  }
  return out;
}

function vtdApp_sanitizeThemeConfig_(themeConfig, previousThemeConfig) {
  themeConfig = themeConfig && typeof themeConfig === "object" ? themeConfig : {};
  previousThemeConfig = previousThemeConfig && typeof previousThemeConfig === "object" ? previousThemeConfig : {};
  let clean = {};
  try { clean = JSON.parse(JSON.stringify(themeConfig)); } catch (err) { clean = {}; }
  clean.assets = clean.assets && typeof clean.assets === "object" ? clean.assets : {};
  clean.assets.app = clean.assets.app && typeof clean.assets.app === "object" ? clean.assets.app : {};
  clean.assets.app.screens = clean.assets.app.screens && typeof clean.assets.app.screens === "object" ? clean.assets.app.screens : {};

  const previousAssets = previousThemeConfig.assets || {};
  const previousAppAssets = previousAssets.app || {};
  const previousScreens = previousAppAssets.screens || {};
  const loginEnabled = !clean.loginAppearance || clean.loginAppearance.hasImage !== false;
  clean.assets.login = vtdApp_sanitizeThemeAsset_(clean.assets.login, previousAssets.login, loginEnabled);

  const globalAppearance = clean.appAppearance && clean.appAppearance.global;
  clean.assets.app.global = vtdApp_sanitizeThemeAsset_(clean.assets.app.global, previousAppAssets.global, !globalAppearance || globalAppearance.hasImage !== false);

  const screenAppearances = clean.appAppearance && clean.appAppearance.screens || {};
  const targets = {};
  Object.keys(clean.assets.app.screens).forEach(key => targets[key] = true);
  Object.keys(previousScreens).forEach(key => targets[key] = true);
  Object.keys(screenAppearances).forEach(key => targets[key] = true);
  const cleanScreens = {};
  Object.keys(targets).forEach(key => {
    const appearance = screenAppearances[key];
    cleanScreens[key] = vtdApp_sanitizeThemeAsset_(clean.assets.app.screens[key], previousScreens[key], !appearance || appearance.hasImage !== false);
  });
  clean.assets.app.screens = cleanScreens;
  return clean;
}

function vtdApp_publicTheme(params) {
  params = params || {};
  const themeConfig = vtdApp_publicSystemConfig_().themeConfig || {};
  const knownVersion = String(params.knownVersion || "");
  const version = String(themeConfig.version || "");
  if (knownVersion && version && knownVersion === version) return vtdApp_ok_({unchanged: true, version: version});
  return vtdApp_ok_({themeConfig: themeConfig, version: version});
}

function vtdApp_cleanUpdateConfig_(update) {
  update = update || {};
  return {
    latestVersion: String(update.latestVersion || "").trim(),
    apkUrl: String(update.apkUrl || "").trim(),
    folderUrl: String(update.folderUrl || "").trim(),
    note: String(update.note || "").trim()
  };
}

function vtdApp_systemConfig_() {
  const cached = CacheService.getScriptCache().get("VTD_SYSTEM_CONFIG");
  if (cached) {
    try { return JSON.parse(cached); } catch (err) {}
  }
  const sh = vtdApp_systemConfigSheet_();
  const config = {returnThApiUrl: "", docOpsApiUrl: "", uiConfig: "", update: "", themeConfig: "", themeImageUrl: ""};
  if (sh.getLastRow() > 1) {
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    values.forEach(row => {
      const key = String(row[0] || "").trim();
      if (key) {
        config[key] = String(row[1] || "").trim();
        if (key === "themeConfig" && vtdApp_driveThemeUrl_(row[3]).url) config.themeImageUrl = String(row[3] || "").trim();
      }
    });
  }
  vtdApp_putSmallCache_("VTD_SYSTEM_CONFIG", config, 120);
  return config;
}

function vtdApp_putSmallCache_(key, value, seconds) {
  try {
    const text = typeof value === "string" ? value : JSON.stringify(value);
    if (text.length > 90000) return;
    CacheService.getScriptCache().put(key, text, seconds || 120);
  } catch (err) {}
}

function vtdApp_systemConfigSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.systemConfigSheet);
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.systemConfigSheet);
    sh.appendRow(["Key", "Value", "UpdatedAt", "UpdatedBy"]);
  }
  return sh;
}

function vtdApp_saveSystemConfig_(config, updatedBy) {
  const sh = vtdApp_systemConfigSheet_();
  const current = {};
  if (sh.getLastRow() > 1) {
    const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    keys.forEach((row, i) => {
      const key = String(row[0] || "").trim();
      if (key) current[key] = {row: i + 2, extra: String(row[3] || "").trim()};
    });
  }
  Object.keys(config || {}).forEach(key => {
    const currentRow = current[key];
    let extra = String(updatedBy || "");
    if (key === "themeConfig") {
      let parsed = {};
      try { parsed = JSON.parse(String(config[key] || "{}")); } catch (err) {}
      const asset = parsed.assets && parsed.assets.login || {};
      const appearance = parsed.loginAppearance || {};
      const configuredLink = String(asset.url || asset.imageUrl || asset.driveUrl || appearance.imageUrl || "").trim();
      extra = configuredLink || (currentRow && vtdApp_driveThemeUrl_(currentRow.extra).url ? currentRow.extra : "");
    }
    const values = [key, String(config[key] || ""), new Date(), extra];
    if (currentRow) sh.getRange(currentRow.row, 1, 1, values.length).setValues([values]);
    else sh.appendRow(values);
  });
  CacheService.getScriptCache().remove("VTD_SYSTEM_CONFIG");
}

function vtdApp_dashboard(params) {
  const denied = vtdApp_requireAction_("dashboard", params);
  if (denied) return denied;
  const cacheKey = "VTD_DASHBOARD_V2_" + vtdApp_dateKey_(params && params.date ? params.date : new Date());
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (cached) return JSON.parse(cached);
  const ss = vtdApp_ss_();
  const raw = ss.getSheetByName(VTD_APP.rawSheet);
  const todayKey = vtdApp_dateKey_(params && params.date ? params.date : new Date());
  const rows = vtdApp_readSheetObjects_(raw, VTD_APP.dashboardRows);
  const seen = {};
  const bySlog = {};
  const byStatus = {};
  const recent = [];
  let todayTotal = 0;
  let queued = 0;
  let error = 0;

  rows.forEach(item => {
    const id = vtdApp_pick_(item.row, ["id"]) || "";
    const maDon = vtdApp_pick_(item.row, ["ma don", "ma don ghn", "ma don ghtk"]) || "";
    const dedupeKey = id || maDon || item.source + ":" + item.rowNumber;
    if (seen[dedupeKey]) return;
    seen[dedupeKey] = true;

    const ngay = vtdApp_pick_(item.row, ["ngay hoan tra", "ngay nhap", "timestamp", "thoi gian nhap hoan"]);
    const rowDateKey = vtdApp_dateKey_(ngay);
    const syncStatus = String(vtdApp_pick_(item.row, ["sync_status"]) || "").toUpperCase();
    const pj2Status = String(vtdApp_pick_(item.row, ["pj2_status"]) || "").toUpperCase();
    const retrySync = String(vtdApp_pick_(item.row, ["retry_sync"]) || "").toUpperCase();
    const state = vtdApp_syncState_(syncStatus, pj2Status, retrySync);
    const slog = String(vtdApp_pick_(item.row, ["slog"]) || "Khác").trim() || "Khác";

    if (rowDateKey === todayKey) {
      todayTotal++;
      bySlog[slog] = (bySlog[slog] || 0) + 1;
      byStatus[state] = (byStatus[state] || 0) + 1;
      recent.push(vtdApp_compactRecord_(item));
    }

    if (state === "PENDING") queued++;
    if (state === "ERROR" || state === "BLANK") error++;
  });

  recent.sort((a, b) => String(b.rowNumber).localeCompare(String(a.rowNumber)));

  const result = {
    todayKey: todayKey,
    totalKnown: Object.keys(seen).length,
    todayTotal: todayTotal,
    queued: queued,
    error: error,
    bySlog: vtdApp_toPairs_(bySlog),
    byStatus: vtdApp_toPairs_(byStatus),
    recent: recent.slice(0, 30)
  };
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 60);
  return result;
}

function vtdApp_storeInfo(params) {
  const denied = vtdApp_requireAction_("storeInfo", params);
  if (denied) return denied;
  params = params || {};
  const ss = vtdApp_ss_();
  const sourceNames = VTD_APP.storeInfoSheets || [];
  const rows = [];
  let source = "";

  for (let i = 0; i < sourceNames.length; i++) {
    const sheet = ss.getSheetByName(sourceNames[i]);
    if (!sheet) continue;
    const items = vtdApp_readSheetObjects_(sheet, 1500);
    if (!items.length) continue;
    source = sheet.getName();
    items.forEach(item => {
      const row = item.row || {};
      const dateVal = vtdApp_pick_(row, ["ngay len don", "ngày lên đơn", "ngay hoan tra", "ngày hoàn trả", "date", "ngay"]);
      rows.push({
        khachHang: vtdApp_pick_(row, ["khach hang", "khách hàng", "ten khach hang", "tên khách hàng", "customer"]),
        maDon: vtdApp_pick_(row, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]),
        ngayLenDon: dateVal,
        od: vtdApp_pick_(row, ["od", "o d"]),
        raw: item.row,
        rowNumber: item.rowNumber
      });
    });
    if (rows.length) break;
  }

  const byDateMap = {};
  rows.forEach(r => {
    const key = vtdApp_dateKey_(r.ngayLenDon) || String(r.ngayLenDon || "Không rõ");
    byDateMap[key] = (byDateMap[key] || 0) + 1;
  });
  const byDate = Object.keys(byDateMap).sort().map(key => ({key: key, value: byDateMap[key]}));
  return vtdApp_ok_({
    source: source || "",
    byDate: byDate,
    rows: rows.slice(-200).reverse()
  });
}

function vtdApp_search(params) {
  const denied = vtdApp_requireAction_("search", params);
  if (denied) return denied;
  params = params || {};
  const query = vtdApp_norm_(params.query || "");
  const filters = {
    slog: vtdApp_norm_(params.slog || ""),
    loaiHoan: vtdApp_norm_(params.loaiHoan || ""),
    fromDate: params.fromDate ? vtdApp_dateKey_(params.fromDate) : "",
    toDate: params.toDate ? vtdApp_dateKey_(params.toDate) : ""
  };
  const limit = Math.min(Number(params.limit) || VTD_APP.defaultLimit, 300);
  const source = params.source || "all";
  const ss = vtdApp_ss_();
  const sheets = [];

  if (source === "hot" || source === "all") sheets.push(ss.getSheetByName(VTD_APP.rawSheet));
  if (source === "full" || source === "all") sheets.push(ss.getSheetByName(VTD_APP.rawFullSheet));

  const results = [];
  const seen = {};

  sheets.forEach(sheet => {
    vtdApp_readSheetObjects_(sheet, VTD_APP.maxSearchRows).forEach(item => {
      if (results.length >= limit) return;
      const compact = vtdApp_compactRecord_(item);
      const key = compact.id || compact.maDon || compact.source + ":" + compact.rowNumber;
      if (seen[key]) return;
      if (!vtdApp_recordPassesFilters_(item, compact, query, filters)) return;
      seen[key] = true;
      results.push(compact);
    });
  });

  return vtdApp_ok_({results: results, count: results.length});
}

function vtdApp_getRecord(params) {
  const denied = vtdApp_requireAction_("getRecord", params);
  if (denied) return denied;
  params = params || {};
  const key = String(params.id || params.maDon || params.query || "").trim();
  if (!key) return vtdApp_fail_("Thiếu ID hoặc mã đơn.");

  const ss = vtdApp_ss_();
  const sheets = [ss.getSheetByName(VTD_APP.rawSheet), ss.getSheetByName(VTD_APP.rawFullSheet)];
  for (let s = 0; s < sheets.length; s++) {
    const found = vtdApp_findRecordInSheet_(sheets[s], key);
    if (found) {
      found.images = vtdApp_findImages_(found.compact.id, found.compact.maDon);
      found.sku = vtdApp_findSku_(found.compact.id);
      found.fields = vtdApp_recordFields_(found);
      return vtdApp_ok_(found);
    }
  }
  return vtdApp_fail_("Không tìm thấy chứng từ.");
}

function vtdApp_saveRaw(params) {
  const denied = vtdApp_requireAction_("saveRaw", params);
  if (denied) return denied;
  params = params || {};
  const ss = vtdApp_ss_();
  const sheet = ss.getSheetByName(VTD_APP.rawSheet);
  if (!sheet) return vtdApp_fail_("Không thấy sheet " + VTD_APP.rawSheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill("");
  const col = vtdApp_headerMap_(headers);
  const payload = params.record || params;
  const id = payload.id || payload.ID || vtdApp_makeId_();

  vtdApp_setByAliases_(row, col, ["id"], id);
  vtdApp_setByAliases_(row, col, ["ma don", "ma don ghn", "ma don ghtk"], payload.maDon || payload["Mã đơn"]);
  vtdApp_setByAliases_(row, col, ["slog"], payload.slog);
  vtdApp_setByAliases_(row, col, ["loai hoan"], payload.loaiHoan);
  vtdApp_setByAliases_(row, col, ["ngay hoan tra"], payload.ngayHoanTra || new Date());
  vtdApp_setByAliases_(row, col, ["thoi gian nhap hoan", "timestamp"], new Date());
  vtdApp_setByAliases_(row, col, ["ma po chung tu hoan", "ma po"], payload.po);
  vtdApp_setByAliases_(row, col, ["tinh trang chung tu"], payload.tinhTrang);
  vtdApp_setByAliases_(row, col, ["thong tin chung tu"], payload.thongTin);
  const auth = vtdApp_auth_(params);
  vtdApp_setByAliases_(row, col, ["user_email", "nguoi nhap", "email"], auth.email || vtdApp_getUser_());

  sheet.appendRow(row);
  const rowNumber = sheet.getLastRow();
  vtdApp_queueRow_(sheet, rowNumber);

  return vtdApp_ok_({
    id: id,
    rowNumber: rowNumber,
    message: "Đã lưu và đưa vào queue xử lý."
  });
}

function vtdApp_updateRecord(params) {
  const denied = vtdApp_requireAction_("updateRecord", params);
  if (denied) return denied;
  params = params || {};
  const key = String(params.id || params.rawId || params.maDon || "").trim();
  if (!key) return vtdApp_fail_("Thiếu ID hoặc mã đơn.");

  const ss = vtdApp_ss_();
  const sheet = ss.getSheetByName(VTD_APP.rawSheet);
  const found = vtdApp_findRecordForUpdate_(sheet, key, params);
  if (!found) return vtdApp_fail_("Chỉ cập nhật được dòng còn trong " + VTD_APP.rawSheet + ".");

  const headers = found.headers;
  const col = vtdApp_headerMap_(headers);
  const row = found.values.slice();
  const auth = vtdApp_auth_(params);
  const maDon = vtdApp_cleanMaDonParam_(params) || found.compact.maDon;
  const xacThuc = String(params.xacThucChungTu || params.xacThucHoaDon || params.xacThuc || "").trim();
  const rawId = String(found.compact.id || key).trim();

  vtdApp_setByAliases_(row, col, ["ma don", "ma don ghn", "ma don ghtk"], maDon);
  vtdApp_setByAliases_(row, col, ["slog"], params.slog);
  vtdApp_setByAliases_(row, col, ["loai hoan"], params.loaiHoan);
  vtdApp_setByAliases_(row, col, ["xac thuc hoa don", "xac thuc chung tu"], xacThuc);
  vtdApp_setByAliases_(row, col, ["ma po chung tu hoan", "ma po"], params.po);
  vtdApp_setByAliases_(row, col, ["tinh trang chung tu"], params.tinhTrang);
  vtdApp_setByAliases_(row, col, ["thong tin chung tu", "ghi chu", "ghi chú"], params.thongTin || params.ghiChu);
  vtdApp_setByAliases_(row, col, ["need_sync", "need sync"], "Yes");
  const routePJ2 = typeof isRouteToPJ2_ === "function" ? isRouteToPJ2_(params.slog, params.loaiHoan) : false;
  const hasImageChange = (Array.isArray(params.images) && params.images.length) || (Array.isArray(params.removeImages) && params.removeImages.length);
  const hasSkuChange = Array.isArray(params.skuItems);
  if (routePJ2) {
    vtdApp_setByAliases_(row, col, ["sync_status"], "ROUTE_PJ2");
    vtdApp_setByAliases_(row, col, ["retry_sync"], "");
    vtdApp_setByAliases_(row, col, ["pj2_status"], "");
  } else {
    vtdApp_setByAliases_(row, col, ["sync_status"], "");
    vtdApp_setByAliases_(row, col, ["retry_sync"], "");
    if (hasImageChange || hasSkuChange) vtdApp_setByAliases_(row, col, ["pj2_status"], "");
  }

  sheet.getRange(found.rowNumber, 1, 1, headers.length).setValues([row]);
  const skuCount = vtdApp_replaceSkuForRaw_(ss, rawId, maDon, params.skuItems, auth);
  const imageUpdate = vtdApp_updateImagesForRaw_(ss, rawId, maDon, params);
  vtdApp_queueRow_(sheet, found.rowNumber, {manualUpdate: true, forcePJ2: hasImageChange || hasSkuChange});

  return vtdApp_ok_({
    id: rawId,
    maDon: maDon,
    rowNumber: found.rowNumber,
    skuCount: skuCount,
    imageUpdate: imageUpdate,
    message: "Đã cập nhật và đưa vào queue xử lý lại."
  });
}

function vtdApp_findRecordForUpdate_(sheet, key, params) {
  if (!sheet) return null;
  const rowNumber = Number(params && params.rowNumber);
  if (rowNumber > 1 && rowNumber <= sheet.getLastRow()) {
    const foundByRow = vtdApp_getRecordByRow_(sheet, rowNumber);
    if (foundByRow) {
      const normKey = vtdApp_norm_(key);
      const compact = foundByRow.compact || {};
      if (
        !normKey ||
        vtdApp_norm_(compact.id) === normKey ||
        vtdApp_norm_(compact.maDon) === normKey ||
        vtdApp_norm_(params && params.id) === vtdApp_norm_(compact.id)
      ) {
        return foundByRow;
      }
    }
  }
  const directKeys = [
    params && params.id,
    params && params.rawId,
    params && params.RAW_ID,
    params && params.ID,
    key
  ].filter(Boolean);
  for (let i = 0; i < directKeys.length; i++) {
    const found = vtdApp_findRecordInSheet_(sheet, String(directKeys[i]).trim());
    if (found) return found;
  }
  return vtdApp_findRecordInSheet_(sheet, key);
}

function vtdApp_replaceSkuForRaw_(ss, rawId, maDon, items, auth) {
  items = Array.isArray(items) ? items.map(vtdApp_normalizeSkuItem_) : [];
  const sheet = ss.getSheetByName(VTD_APP.skuSheet);
  if (sheet && sheet.getLastRow() > 1 && rawId) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col = vtdApp_headerMap_(headers);
    const rawCol = vtdApp_firstCol_(col, ["raw_id", "raw id"]);
    if (rawCol != null) {
      const values = sheet.getRange(2, rawCol + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (let i = values.length - 1; i >= 0; i--) {
        if (String(values[i][0] || "").trim() === String(rawId).trim()) {
          sheet.deleteRow(i + 2);
        }
      }
    }
  }
  if (!items.length) return 0;
  vtdApp_syncNewMasterSku_(ss, items);
  return vtdApp_saveSkuRows_(ss, rawId, maDon, items, auth);
}

function vtdApp_updateImagesForRaw_(ss, rawId, maDon, params) {
  params = params || {};
  const removed = Array.isArray(params.removeImages) ? params.removeImages : [];
  const files = Array.isArray(params.images) ? params.images : [];
  const removedCount = vtdApp_removeImageRowsForRaw_(ss, rawId, maDon, removed);
  const uploaded = [];
  const errors = [];
  files.forEach((file, i) => {
    file = file || {};
    const res = vtdApp_uploadImage(Object.assign({}, params, file, {
      rawId: rawId,
      maDon: maDon,
      fileName: file.fileName || file.name || params.fileName,
      mimeType: file.mimeType || file.type || params.mimeType,
      base64: file.base64 || file.data || params.base64,
      forceFullSync: false
    }));
    if (res && res.ok) uploaded.push(res);
    else errors.push("Ảnh " + (i + 1) + ": " + (res && res.message || "Upload lỗi"));
  });
  return {removed: removedCount, uploaded: uploaded.length, errors: errors};
}

function vtdApp_removeImageRowsForRaw_(ss, rawId, maDon, removed) {
  if (!Array.isArray(removed) || !removed.length) return 0;
  const removeKeys = {};
  removed.forEach(item => {
    if (!item) return;
    ["rowNumber", "name", "url", "rawPath", "fileUrl", "folderUrl"].forEach(key => {
      if (item[key] !== undefined && item[key] !== "") removeKeys[String(item[key]).trim()] = true;
    });
  });
  let count = 0;
  [VTD_APP.imageSheet, VTD_APP.imageSheetAlt, VTD_APP.imageFullSheet].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const col = vtdApp_headerMap_(values[0]);
    for (let r = values.length - 1; r >= 1; r--) {
      const row = values[r];
      const rowRawId = String(vtdApp_pickFromArray_(row, col, ["raw_id", "raw id", "id raw"]) || "").trim();
      const rowMaDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn"]) || "").trim();
      if (rawId && rowRawId !== String(rawId).trim() && maDon && rowMaDon !== String(maDon).trim()) continue;
      const candidates = [
        String(r + 1),
        String(vtdApp_pickFromArray_(row, col, ["anh", "ảnh", "ten file", "file name"]) || "").trim(),
        String(vtdApp_pickFromArray_(row, col, ["link anh", "link anh/video", "file url"]) || "").trim(),
        String(vtdApp_pickFromArray_(row, col, ["folder url"]) || "").trim()
      ];
      if (candidates.some(key => key && removeKeys[key])) {
        sheet.deleteRow(r + 1);
        count++;
      }
    }
  });
  return count;
}

function vtdApp_queueRecord(params) {
  const denied = vtdApp_requireAction_("queueRecord", params);
  if (denied) return denied;
  params = params || {};
  const key = String(params.id || params.maDon || "").trim();
  if (!key) return vtdApp_fail_("Thiếu ID hoặc mã đơn.");

  const ss = vtdApp_ss_();
  const sheet = ss.getSheetByName(VTD_APP.rawSheet);
  const found = vtdApp_findRecordInSheet_(sheet, key);
  if (!found) return vtdApp_fail_("Không thấy dòng còn trong " + VTD_APP.rawSheet + ".");

  vtdApp_queueRow_(sheet, found.rowNumber);
  return vtdApp_ok_({message: "Đã đưa dòng " + found.rowNumber + " vào queue.", rowNumber: found.rowNumber});
}

function vtdApp_uploadImage(params) {
  const denied = vtdApp_requireAction_("uploadImage", params);
  if (denied) return denied;
  params = params || {};
  const rawId = String(params.rawId || params.id || "").trim();
  const maDon = String(params.maDon || "").trim();
  const fileName = String(params.fileName || "chung-tu.jpg").trim();
  const mimeType = String(params.mimeType || "image/jpeg").trim();
  const base64 = String(params.base64 || params.data || "").replace(/^data:[^,]+,/, "");
  const imageType = String(params.imageType || params.type || "").toLowerCase().trim();

  if (!rawId && !maDon) return vtdApp_fail_("Thiếu RAW_ID hoặc mã đơn.");
  if (!base64) return vtdApp_fail_("Thiếu dữ liệu ảnh.");

  const isProductImage = imageType === "product" || imageType === "sku" || imageType === "sanpham" || imageType === "san pham";
  const root = isProductImage ? vtdApp_getSkuImageRootFolder_() : vtdApp_getImageRootFolder_();
  const folderName = maDon || rawId;
  const needsProductFolder = Number(params.productImageCount || 0) > 1 || Number(params.skuCount || 0) > 1;
  const folder = isProductImage && !needsProductFolder ? root : vtdApp_getOrCreateSubFolder_(root, folderName);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, vtdApp_safeFileName_(fileName, rawId, maDon));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  vtdApp_shareFolderOnce_(folder);

  const ss = vtdApp_ss_();
  let skuUpdate = null;
  if (isProductImage) {
    skuUpdate = vtdApp_updateSkuImage_(ss, {
      rawId: rawId,
      maDon: maDon,
      skuIndex: params.skuIndex,
      skuName: params.skuName,
      imagePath: isProductImage && !needsProductFolder ? root.getName() + "/" + file.getName() : root.getName() + "/" + folder.getName() + "/" + file.getName(),
      imageUrl: file.getUrl(),
      user: (vtdApp_auth_(params).email || vtdApp_getUser_())
    });
  } else {
    const imageSheet = vtdApp_getOrCreateImageSheet_(ss);
    const append = vtdApp_buildImageRow_(imageSheet, {
      rawId: rawId,
      maDon: maDon,
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      folderUrl: folder.getUrl(),
      user: (vtdApp_auth_(params).email || vtdApp_getUser_())
    });
    imageSheet.appendRow(append);
  }

  if (!isProductImage && rawId && (params.forceFullSync === true || params.forceFullSync === "true") && typeof fullSyncImagesByRawId === "function") {
    try { fullSyncImagesByRawId(rawId); } catch (err) {}
  }

  return vtdApp_ok_({
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    folderUrl: folder.getUrl(),
    imageType: isProductImage ? "product" : "document",
    skuUpdate: skuUpdate,
    message: "Đã upload ảnh."
  });
}

function vtdApp_runCommand(params) {
  const denied = vtdApp_requireAction_("runCommand", params);
  if (denied) return denied;
  params = params || {};
  const action = params.action;
  const startRow = Number(params.startRow) || 2;
  const endRow = Number(params.endRow) || startRow;

  if (action === "processPJ1" && typeof processRetryQueue === "function") {
    return vtdApp_ok_({result: processRetryQueue(Number(params.limit) || 15)});
  }
  if (action === "processPJ2" && typeof processPJ2Queue === "function") {
    return vtdApp_ok_({result: processPJ2Queue(Number(params.limit) || 15)});
  }
  if (action === "dispatchBackfill" && typeof dispatchBackfillPendingRows === "function") {
    return vtdApp_ok_({result: dispatchBackfillPendingRows(startRow, endRow)});
  }
  if (action === "fullSync" && typeof fullSyncInputToHistory === "function") {
    return vtdApp_ok_({result: fullSyncInputToHistory()});
  }

  return vtdApp_fail_("Lệnh không hợp lệ hoặc function cũ chưa có trong project.");
}

function vtdApp_adminConfig(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  const sh = vtdApp_ensureUserSheet_();
  const config = vtdApp_permissionConfig_();
  const userMap = {};
  Object.keys(config.users).forEach(email => {
    userMap[email] = {
      email: email,
      role: config.users[email].role || "user",
      screens: config.users[email].screens || [],
      actions: config.users[email].actions || [],
      pin: config.pins[email] || "",
      active: config.users[email].active !== false
    };
  });
  try {
    const values = sh.getDataRange().getValues();
    if (values.length > 1) {
      const col = vtdApp_headerMap_(values[0]);
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const email = String(row[col.email] || "").toLowerCase().trim();
        if (!email) continue;
        const activeText = String(row[col.active] || "TRUE").toUpperCase().trim();
        const role = String(row[col.role] || "user").toLowerCase().trim();
        const screens = vtdApp_splitCsv_(row[col.screens]);
        const actions = vtdApp_splitCsv_(row[col.actions]);
        userMap[email] = {
          email: email,
          role: role === "admin" ? "admin" : "staff",
          screens: screens.length ? screens : ["home", "search", "input", "store", "settings"],
          actions: actions.length ? actions : ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"],
          pin: String(vtdApp_pickFromArray_(row, col, ["pass", "pin"]) || "").trim(),
          active: !(activeText === "FALSE" || activeText === "NO" || activeText === "0")
        };
      }
    }
  } catch (err) {}
  const users = Object.keys(userMap).sort().map(email => userMap[email]);
  return vtdApp_ok_({
    adminEmails: config.adminEmails.map(email => ({email: email, pin: config.pins[email] || ""})),
    users: users,
    allScreens: VTD_PERMISSIONS.allScreens.slice(),
    allActions: VTD_PERMISSIONS.allActions.slice()
  });
}

function vtdApp_savePermission(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  if (!email || email.indexOf("@") < 0) return vtdApp_fail_("Email không hợp lệ.");

  const role = params.role === "admin" ? "admin" : "staff";
  const pin = String(params.pin || "").trim();
  const screens = Array.isArray(params.screens) ? params.screens : [];
  const actions = Array.isArray(params.actions) ? params.actions : [];
  if (!screens.length || !actions.length) return vtdApp_fail_("Phải chọn ít nhất 1 màn và 1 thao tác.");

  const sh = vtdApp_ensureUserSheet_();
  const values = sh.getDataRange().getValues();
  const header = values[0];
  const col = vtdApp_headerMap_(header);
  const emailCol = vtdApp_firstCol_(col, ["email"]);
  const passCol = vtdApp_firstCol_(col, ["pass", "pin"]);
  let targetRow = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][emailCol] || "").toLowerCase().trim() === email) {
      targetRow = i + 1;
      break;
    }
  }

  const row = targetRow ? values[targetRow - 1].slice(0, header.length) : new Array(header.length).fill("");
  vtdApp_setByAliases_(row, col, ["email"], email);
  vtdApp_setByAliases_(row, col, ["role"], role);
  vtdApp_setByAliases_(row, col, ["screens"], screens.join(","));
  vtdApp_setByAliases_(row, col, ["actions"], actions.join(","));
  const currentPin = targetRow && passCol != null ? String(values[targetRow - 1][passCol] || "").trim() : "";
  vtdApp_setByAliases_(row, col, ["pass", "pin"], pin || currentPin);
  const active = params.active === false || String(params.active || "").toUpperCase() === "FALSE" ? "FALSE" : "TRUE";
  vtdApp_setByAliases_(row, col, ["active"], active);
  vtdApp_setByAliases_(row, col, ["updated_by", "updated by"], vtdApp_auth_(params).email);
  vtdApp_setByAliases_(row, col, ["updated_at", "updated at"], new Date());

  if (targetRow) {
    sh.getRange(targetRow, 1, 1, header.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  vtdApp_clearUserAuthCache_();
  return vtdApp_ok_({message: "Đã lưu quyền cho " + email});
}

function vtdApp_deletePermission(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  if (!email) return vtdApp_fail_("Thiếu email.");

  const sh = vtdApp_ensureUserSheet_();
  const values = sh.getDataRange().getValues();
  const col = vtdApp_headerMap_(values[0]);
  const emailCol = vtdApp_firstCol_(col, ["email"]);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][emailCol] || "").toLowerCase().trim() === email) {
      sh.deleteRow(i + 1);
      vtdApp_clearUserAuthCache_();
      return vtdApp_ok_({message: "Đã xóa quyền " + email});
    }
  }
  return vtdApp_ok_({message: "Email không có trong danh sách quyền."});
}

function vtdApp_upsertPermissionPin_(email, pin, updatedBy) {
  vtdApp_updateUserPass_(email, pin, updatedBy);
}

function vtdApp_queueRow_(sheet, rowNumber, options) {
  options = options || {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const slog = vtdApp_pickFromArray_(values, col, ["slog"]);
  const loai = vtdApp_pickFromArray_(values, col, ["loai hoan"]);
  const routePJ2 = typeof isRouteToPJ2_ === "function" ? isRouteToPJ2_(slog, loai) : false;

  if (options.manualUpdate) {
    if (routePJ2 || options.forcePJ2) vtdApp_setCellByAliases_(sheet, rowNumber, col, ["pj2_status"], "Yes");
    if (routePJ2) {
      vtdApp_setCellByAliases_(sheet, rowNumber, col, ["sync_status"], "ROUTE_PJ2");
      vtdApp_setCellByAliases_(sheet, rowNumber, col, ["retry_sync"], "");
    } else {
      vtdApp_setCellByAliases_(sheet, rowNumber, col, ["retry_sync"], "Yes");
    }
  } else if (routePJ2) {
    vtdApp_setCellByAliases_(sheet, rowNumber, col, ["pj2_status"], "Yes");
    vtdApp_setCellByAliases_(sheet, rowNumber, col, ["sync_status"], "ROUTE_PJ2");
    vtdApp_setCellByAliases_(sheet, rowNumber, col, ["retry_sync"], "");
  } else {
    vtdApp_setCellByAliases_(sheet, rowNumber, col, ["pj2_status"], "Yes");
    vtdApp_setCellByAliases_(sheet, rowNumber, col, ["retry_sync"], "Yes");
  }

  if (typeof fullSyncRawRowByNumber === "function") {
    try { fullSyncRawRowByNumber(rowNumber); } catch (err) {}
  }
}

function vtdApp_ss_() {
  if (VTD_APP.spreadsheetId) {
    return SpreadsheetApp.openById(VTD_APP.spreadsheetId);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) {
    throw new Error("Khong lay duoc Google Sheet. Hay dien VTD_APP.spreadsheetId bang ID file Google Sheet goc.");
  }
  return ss;
}

function vtdApp_readSheetObjects_(sheet, maxRows) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return [];

  const fromRow = Math.max(2, lastRow - (Number(maxRows) || VTD_APP.maxSearchRows) + 1);
  const numRows = lastRow - fromRow + 1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const range = sheet.getRange(fromRow, 1, numRows, lastCol);
  const data = range.getValues();
  const displayData = range.getDisplayValues();
  const out = [];

  data.forEach((values, index) => {
    if (values.join("") === "") return;
    const row = {};
    const displayRow = {};
    headers.forEach((header, i) => {
      const key = vtdApp_norm_(header);
      row[key] = values[i];
      displayRow[key] = displayData[index][i];
    });
    out.push({
      source: sheet.getName(),
      rowNumber: fromRow + index,
      headers: headers,
      values: values,
      displayValues: displayData[index],
      row: row,
      displayRow: displayRow
    });
  });

  out.reverse();
  return out;
}

function vtdApp_readTodayRawM3Objects_(sheet, todayKey) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 2) return [];

  const dateCol = 2; // RAW_IN_M3: column B = Ngày hoàn trả.
  const lastDateKey = vtdApp_dateKey_(sheet.getRange(lastRow, dateCol).getDisplayValue());
  if (lastDateKey && lastDateKey < todayKey) return [];
  const chunkSize = 400;
  let endRow = lastRow;
  let firstTodayRow = 0;
  let lastTodayRow = 0;
  let foundToday = false;
  let stopScan = false;

  while (endRow >= 2 && !stopScan) {
    const startRow = Math.max(2, endRow - chunkSize + 1);
    const numRows = endRow - startRow + 1;
    const dates = sheet.getRange(startRow, dateCol, numRows, 1).getDisplayValues();

    for (let i = numRows - 1; i >= 0; i--) {
      const rowNumber = startRow + i;
      const key = vtdApp_dateKey_(dates[i][0]);
      if (key === todayKey) {
        if (!lastTodayRow) lastTodayRow = rowNumber;
        firstTodayRow = rowNumber;
        foundToday = true;
        continue;
      }
      if (foundToday && key && key !== todayKey) {
        stopScan = true;
        break;
      }
    }

    endRow = startRow - 1;
  }

  if (!firstTodayRow || !lastTodayRow) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const numRows = lastTodayRow - firstTodayRow + 1;
  const range = sheet.getRange(firstTodayRow, 1, numRows, lastCol);
  const data = range.getValues();
  const displayData = range.getDisplayValues();
  const out = [];

  data.forEach((values, index) => {
    if (values.join("") === "") return;
    const row = {};
    const displayRow = {};
    headers.forEach((header, i) => {
      const key = vtdApp_norm_(header);
      row[key] = values[i];
      displayRow[key] = displayData[index][i];
    });
    out.push({
      source: sheet.getName(),
      rowNumber: firstTodayRow + index,
      headers: headers,
      values: values,
      displayValues: displayData[index],
      row: row,
      displayRow: displayRow
    });
  });

  out.reverse();
  return out;
}

function vtdApp_findRecordInSheet_(sheet, key) {
  if (!sheet) return null;
  const normKey = vtdApp_norm_(key);
  const items = vtdApp_readSheetObjects_(sheet, VTD_APP.maxSearchRows);
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const compact = vtdApp_compactRecord_(item);
    if (
      vtdApp_norm_(compact.id) === normKey ||
      vtdApp_norm_(compact.maDon) === normKey ||
      vtdApp_norm_(compact.po) === normKey
    ) {
      return {
        source: item.source,
        rowNumber: item.rowNumber,
        headers: item.headers,
        values: item.values,
        row: item.row,
        compact: compact
      };
    }
  }
  return null;
}

function vtdApp_recordPassesFilters_(item, compact, query, filters) {
  const rowText = vtdApp_norm_(JSON.stringify(compact) + " " + JSON.stringify(item.row));
  if (query && rowText.indexOf(query) < 0) return false;
  if (filters.slog && vtdApp_norm_(compact.slog).indexOf(filters.slog) < 0) return false;
  if (filters.loaiHoan && vtdApp_norm_(compact.loaiHoan).indexOf(filters.loaiHoan) < 0) return false;
  const dateKey = vtdApp_dateKey_(vtdApp_pick_(item.row, ["ngay hoan tra", "ngày hoàn trả", "ngay nhap", "timestamp", "thoi gian nhap hoan", "thời gian nhập hoàn"]));
  if (filters.fromDate && dateKey && dateKey < filters.fromDate) return false;
  if (filters.toDate && dateKey && dateKey > filters.toDate) return false;
  return true;
}

function vtdApp_itemValue_(item, aliases, fallbackIndex, useDisplay) {
  const row = useDisplay ? (item.displayRow || item.row || {}) : (item.row || {});
  const picked = vtdApp_pick_(row, aliases);
  if (picked !== "") return picked;
  const values = useDisplay ? (item.displayValues || item.values || []) : (item.values || []);
  if (fallbackIndex != null && values[fallbackIndex] !== undefined) return values[fallbackIndex];
  return "";
}

function vtdApp_timeDisplay24_(value, fallback) {
  const fallbackText = String(fallback == null ? "" : fallback).trim();
  const fallbackMatch = fallbackText.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (fallbackMatch && fallbackMatch[3] !== "00") {
    return [fallbackMatch[1], fallbackMatch[2], fallbackMatch[3]].map((x, i) => i ? String(x).padStart(2, "0") : String(Number(x)).padStart(2, "0")).join(":");
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, "Asia/Ho_Chi_Minh", "HH:mm:ss");
  }
  if (typeof value === "number" && isFinite(value)) {
    const totalSeconds = Math.round(((value % 1) + 1) % 1 * 24 * 60 * 60);
    const hh = Math.floor(totalSeconds / 3600);
    const mm = Math.floor((totalSeconds % 3600) / 60);
    const ss = totalSeconds % 60;
    return [hh, mm, ss].map(n => String(n).padStart(2, "0")).join(":");
  }
  const text = String(value == null ? "" : value).trim();
  if (text) {
    const m = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m) return [m[1], m[2], m[3]].filter(x => x !== undefined && x !== "").map((x, i) => i ? String(x).padStart(2, "0") : String(Number(x)).padStart(2, "0")).join(":");
    return text;
  }
  return fallbackText;
}

function vtdApp_compactRecord_(item) {
  const row = item.row;
  const displayRow = item.displayRow || row;
  const syncStatus = vtdApp_pick_(row, ["sync_status"]);
  const pj2Status = vtdApp_pick_(row, ["pj2_status"]);
  const retrySync = vtdApp_pick_(row, ["retry_sync"]);
  const ngayHoanTra = vtdApp_itemValue_(item, ["ngay hoan tra", "ngày hoàn trả", "ngay nhap"], 1, true);
  const thoiGianNhapCellText = String((item.displayValues && item.displayValues[2]) || "").trim();
  const thoiGianNhapDisplay = thoiGianNhapCellText || vtdApp_itemValue_(item, ["thoi gian nhap hoan", "thời gian nhập hoàn", "thoi gian", "timestamp"], 2, true);
  const thoiGianNhapRaw = vtdApp_itemValue_(item, ["thoi gian nhap hoan", "thời gian nhập hoàn", "thoi gian", "timestamp"], 2, false);
  const thoiGianNhap = vtdApp_timeDisplay24_(thoiGianNhapDisplay, thoiGianNhapRaw);
  return {
    source: item.source,
    rowNumber: item.rowNumber,
    id: vtdApp_itemValue_(item, ["id"], 0, false),
    maDon: vtdApp_itemValue_(item, ["ma don", "mã đơn", "ma don ghn", "ma don ghtk", "mã đơn ghtk", "madon"], 3, false),
    slog: vtdApp_itemValue_(item, ["slog"], 4, false),
    loaiHoan: vtdApp_itemValue_(item, ["loai hoan", "loại hoàn", "loai st", "loại st"], 6, false),
    ngayHoanTra: ngayHoanTra || vtdApp_toDisplay_(vtdApp_itemValue_(item, ["ngay hoan tra", "ngày hoàn trả", "ngay nhap"], 1, false)), 
    thoiGianNhap: thoiGianNhap || vtdApp_toDisplay_(vtdApp_itemValue_(item, ["thoi gian nhap hoan", "thời gian nhập hoàn", "thoi gian", "timestamp"], 2, false)),
    thoiGianNhapText: thoiGianNhapCellText || thoiGianNhap,
    thoiGianNhapSource: thoiGianNhapCellText ? "displayValues[2]" : "fallback",
    userEmail: vtdApp_itemValue_(item, ["user_email", "user email", "email", "nguoi nhap", "người nhập", "nguoi thao tac", "người thao tác", "user thao tac", "user thao tác", "created by", "createdby"], 11, false),
    po: vtdApp_itemValue_(item, ["ma po chung tu hoan", "mã po chứng từ hoàn", "ma po"], 7, false),
    maQr: vtdApp_itemValue_(item, ["ma qr", "mã qr", "qr"], 9, false),
    soHoaDon: vtdApp_pick_(row, ["so hoa don", "so hoa don", "invoice"]),
    od: vtdApp_pick_(row, ["od"]),
    khachHang: vtdApp_pick_(row, ["khach hang", "customer"]),
    serial: vtdApp_pick_(row, ["ma serial", "serial", "ma qr sua bot"]),
    xacThucHoaDon: vtdApp_pick_(row, ["xac thuc hoa don", "xac thuc chung tu"]),
    tinhTrang: vtdApp_itemValue_(item, ["tinh trang chung tu", "tình trạng chứng từ", "tinh trang"], 8, false),
    thongTin: vtdApp_itemValue_(item, ["thong tin chung tu", "thông tin chứng từ", "ghi chu", "ghi chú"], 10, false),
    syncStatus: syncStatus,
    pj2Status: pj2Status,
    retrySync: retrySync,
    syncState: vtdApp_syncState_(syncStatus, pj2Status, retrySync),
    syncLabel: vtdApp_syncLabel_(syncStatus, pj2Status, retrySync)
  };
}

function vtdApp_recordFields_(found) {
  const fields = [];
  found.headers.forEach((header, i) => {
    if (!header) return;
    fields.push({
      label: header,
      value: vtdApp_toDisplay_(found.values[i])
    });
  });
  return fields;
}

function vtdApp_syncState_(syncStatus, pj2Status, retrySync) {
  const values = [syncStatus, pj2Status, retrySync].map(v => String(v || "").trim().toUpperCase()).filter(Boolean);
  if (!values.length) return "BLANK";
  if (values.some(v => v.indexOf("ERROR") >= 0 || v.indexOf("IMG") >= 0)) return "ERROR";
  if (values.some(v => v === "DONE" || v === "ROUTE_PJ2")) return "DONE";
  if (values.some(v => v === "YES" || v === "RUNNING")) return "PENDING";
  return "DONE";
}

function vtdApp_syncLabel_(syncStatus, pj2Status, retrySync) {
  const state = vtdApp_syncState_(syncStatus, pj2Status, retrySync);
  return VTD_STATUS_LABELS[state] || state;
}

function vtdApp_cleanStatusLabel_(value) {
  const text = String(value == null ? "" : value).trim();
  if (!text) return "";
  const upper = text.toUpperCase();
  if (upper === "DONE" || upper === "ROUTE_PJ2") return "Đã đồng bộ";
  if (upper === "YES" || upper === "RUNNING" || upper === "PENDING") return "Đang đồng bộ";
  if (upper.indexOf("ERROR") >= 0 || upper.indexOf("IMG") >= 0) return "Lỗi";
  if (text.indexOf("\u00c4") >= 0 || text.indexOf("\u00c3") >= 0 || text.indexOf("\u00e1") >= 0 || text.indexOf("\u00c6") >= 0) {
    if (upper.indexOf("RUNNING") >= 0 || upper.indexOf("PENDING") >= 0 || upper.indexOf("YES") >= 0) return "Đang đồng bộ";
    if (upper.indexOf("ERROR") >= 0 || upper.indexOf("IMG") >= 0 || text.indexOf("L") === 0) return "Lỗi";
    return "Đã đồng bộ";
  }
  return text;
}

function vtdApp_findImages_(rawId, maDon) {
  const ss = vtdApp_ss_();
  const sheets = [
    ss.getSheetByName(VTD_APP.imageSheet),
    ss.getSheetByName(VTD_APP.imageSheetAlt),
    ss.getSheetByName(VTD_APP.imageFullSheet)
  ];
  const out = [];
  const seen = {};
  sheets.forEach(sheet => {
    vtdApp_readSheetObjects_(sheet, 3000).forEach(item => {
      const id = vtdApp_pick_(item.row, ["raw_id", "raw id", "id raw"]);
      const order = vtdApp_pick_(item.row, ["ma don"]);
      if (rawId && String(id).trim() !== String(rawId).trim() && maDon && String(order).trim() !== String(maDon).trim()) return;
      if (!rawId && maDon && String(order).trim() !== String(maDon).trim()) return;
      const url = vtdApp_pick_(item.row, ["link anh", "link anh/video", "anh", "folder url", "file url"]);
      const key = url || item.source + ":" + item.rowNumber;
      if (seen[key]) return;
      seen[key] = true;
      out.push({
        source: item.source,
        rowNumber: item.rowNumber,
        rawId: id,
        maDon: order,
        name: vtdApp_pick_(item.row, ["anh", "ten file", "file name"]),
        url: vtdApp_resolveImageUrl_(url),
        rawPath: url
      });
    });
  });
  return out.slice(0, 80);
}

function vtdApp_findSku_(rawId) {
  if (!rawId) return [];
  const ss = vtdApp_ss_();
  const sheet = ss.getSheetByName(VTD_APP.skuSheet);
  return vtdApp_readSheetObjects_(sheet, 3000)
    .filter(item => String(vtdApp_pick_(item.row, ["raw_id", "raw id"])).trim() === String(rawId).trim())
    .slice(0, 100)
    .map(item => ({
      source: item.source,
      rowNumber: item.rowNumber,
      ten: vtdApp_pick_(item.row, ["ten san pham hoan ve", "ten san pham"]),
      soLuong: vtdApp_pick_(item.row, ["so luong"]),
      tinhTrang: vtdApp_pick_(item.row, ["tinh trang"]),
      loHsd: vtdApp_pick_(item.row, ["ma lo/hsd", "ma lo hsd"]),
      serial: vtdApp_pick_(item.row, ["ma serial", "ma qr sua bot"]),
      anh: vtdApp_pick_(item.row, ["anh san pham", "anh"]),
      anhUrl: vtdApp_resolveImageUrl_(vtdApp_pick_(item.row, ["anh san pham", "anh"]))
    }));
}

function vtdApp_resolveImageUrl_(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  const name = raw.split("/").pop();
  if (!name) return raw;
  const cacheKey = "VTD_IMG_URL_" + Utilities.base64EncodeWebSafe(name).slice(0, 80);
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (cached) return cached;
  try {
    const files = DriveApp.getFilesByName(name);
    if (files.hasNext()) {
      const f = files.next();
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (errShare) {}
      const url = f.getUrl();
      CacheService.getScriptCache().put(cacheKey, url, 21600);
      return url;
    }
  } catch (err) {}
  return raw;
}

function vtdApp_getOrCreateImageSheet_(ss) {
  let sheet = ss.getSheetByName(VTD_APP.imageSheet) || ss.getSheetByName(VTD_APP.imageSheetAlt);
  if (!sheet) {
    sheet = ss.insertSheet(VTD_APP.imageSheet);
    sheet.appendRow(["ID", "RAW_ID", "Mã đơn", "Ảnh", "Link ảnh", "Folder URL", "Người nhập", "Thời gian"]);
  }
  return sheet;
}

function vtdApp_buildImageRow_(sheet, data) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const row = new Array(headers.length).fill("");
  vtdApp_setByAliases_(row, col, ["id"], vtdApp_makeId_());
  vtdApp_setByAliases_(row, col, ["raw_id", "raw id"], data.rawId);
  vtdApp_setByAliases_(row, col, ["ma don"], data.maDon);
  vtdApp_setByAliases_(row, col, ["anh", "ten file", "file name"], data.fileName);
  vtdApp_setByAliases_(row, col, ["link anh", "link anh/video", "file url"], data.fileUrl);
  vtdApp_setByAliases_(row, col, ["folder url"], data.folderUrl);
  vtdApp_setByAliases_(row, col, ["nguoi nhap", "email", "user_email", "user email"], data.user);
  const now = new Date();
  vtdApp_setByAliases_(row, col, ["ngay", "date"], Utilities.formatDate(now, "Asia/Saigon", "dd/MM/yyyy"));
  vtdApp_setByAliases_(row, col, ["thoi gian", "timestamp", "time"], Utilities.formatDate(now, "Asia/Saigon", "HH:mm:ss"));
  return row;
}

function vtdApp_cleanMaDonParam_(params) {
  params = params || {};
  const values = [params.maDon, params.ma_don, params.orderCode, params.ma, params["Mã đơn"], params["ma don"]];
  let fallback = "";
  for (let i = 0; i < values.length; i++) {
    const value = String(values[i] == null ? "" : values[i]).trim();
    if (!value) continue;
    if (!fallback) fallback = value;
    if (value !== "0") return value;
  }
  return fallback;
}

function vtdApp_getImageRootFolder_() {
  for (let i = 0; i < VTD_APP.imageRootFolderNames.length; i++) {
    const folders = DriveApp.getFoldersByName(VTD_APP.imageRootFolderNames[i]);
    if (folders.hasNext()) return folders.next();
  }
  return DriveApp.createFolder(VTD_APP.imageRootFolderNames[0]);
}

function vtdApp_getSkuImageRootFolder_() {
  const names = ["List_SKU_hoàn_Images", "List_SKU_hoan_Images", "List_SKU_hoàn"];
  for (let i = 0; i < names.length; i++) {
    const folders = DriveApp.getFoldersByName(names[i]);
    if (folders.hasNext()) return folders.next();
  }
  if (VTD_APP.driveSkuFolderId) {
    try { return DriveApp.getFolderById(VTD_APP.driveSkuFolderId); } catch (err) {}
  }
  return DriveApp.createFolder(names[0]);
}

function vtdApp_updateSkuImage_(ss, data) {
  data = data || {};
  const sheet = ss.getSheetByName(VTD_APP.skuSheet);
  if (!sheet || sheet.getLastRow() <= 1) return {updated: false, message: "Không thấy dòng SKU để cập nhật ảnh."};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const rawId = String(data.rawId || "").trim();
  const maDon = String(data.maDon || "").trim();
  const skuIndex = Math.max(0, Number(data.skuIndex) || 0);
  const skuNameNorm = vtdApp_norm_(data.skuName || "");
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const candidates = [];
  values.forEach((row, i) => {
    const rowRawId = String(vtdApp_pickFromArray_(row, col, ["raw_id", "raw id"]) || "").trim();
    const rowMaDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || "").trim();
    if ((rawId && rowRawId === rawId) || (maDon && rowMaDon === maDon)) {
      const rowSkuName = vtdApp_norm_(vtdApp_pickFromArray_(row, col, ["ten san pham hoan", "tên sản phẩm hoàn về", "ten san pham", "tên sản phẩm", "sku"]));
      candidates.push({rowNumber: i + 2, skuName: rowSkuName});
    }
  });
  if (!candidates.length) return {updated: false, message: "Không thấy RAW_ID trong List_SKU_hoàn."};
  let target = candidates[Math.min(skuIndex, candidates.length - 1)] || null;
  if (!target && skuNameNorm) target = candidates.find(item => item.skuName === skuNameNorm) || null;
  const currentImage = String(vtdApp_pickFromArray_(sheet.getRange(target.rowNumber, 1, 1, sheet.getLastColumn()).getDisplayValues()[0], col, ["anh san pham", "Ảnh sản phẩm", "ảnh sản phẩm", "anh", "ảnh", "link anh", "link ảnh"]) || "").trim();
  const imagePath = String(data.imagePath || data.imageUrl || "").trim();
  const nextImage = currentImage && currentImage.indexOf(imagePath) < 0 ? currentImage + "; " + imagePath : (currentImage || imagePath);
  const now = new Date();
  const ngayText = Utilities.formatDate(now, "Asia/Saigon", "dd/MM/yyyy");
  const gioText = Utilities.formatDate(now, "Asia/Saigon", "HH:mm:ss");
  const updatedImage = vtdApp_setCellByAliases_(sheet, target.rowNumber, col, ["anh san pham","Ảnh sản phẩm", "ảnh sản phẩm", "anh", "ảnh", "link anh", "link ảnh"], nextImage);
  vtdApp_setCellByAliases_(sheet, target.rowNumber, col, ["user_email", "email", "nguoi nhap", "người nhập"], data.user || "");
  vtdApp_setCellByAliases_(sheet, target.rowNumber, col, ["ngay", "ngày", "date"], ngayText);
  vtdApp_setCellByAliases_(sheet, target.rowNumber, col, ["gio", "giờ", "thoi gian", "thời gian", "timestamp", "time"], gioText);
  return {updated: updatedImage, rowNumber: target.rowNumber, imagePath: imagePath};
}

function vtdApp_getOrCreateSubFolder_(root, name) {
  const folders = root.getFoldersByName(String(name || "NO_CODE").trim());
  return folders.hasNext() ? folders.next() : root.createFolder(String(name || "NO_CODE").trim());
}

function vtdApp_shareFolderOnce_(folder) {
  if (!folder) return;
  try {
    const cache = CacheService.getScriptCache();
    const key = "VTD_SHARED_FOLDER_" + folder.getId();
    if (cache.get(key)) return;
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    cache.put(key, "1", 21600);
  } catch (err) {}
}

function vtdApp_headerMap_(headers) {
  const col = {};
  headers.forEach((header, i) => col[vtdApp_norm_(header)] = i);
  return col;
}

function vtdApp_pick_(row, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = vtdApp_norm_(aliases[i]);
    if (row[key] !== undefined && row[key] !== "") return row[key];
  }
  return "";
}

function vtdApp_pickFromArray_(row, col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = col[vtdApp_norm_(aliases[i])];
    if (idx != null) return row[idx];
  }
  return "";
}

function vtdApp_setByAliases_(row, col, aliases, value) {
  if (value === undefined) return false;
  for (let i = 0; i < aliases.length; i++) {
    const idx = col[vtdApp_norm_(aliases[i])];
    if (idx != null) {
      row[idx] = value;
      return true;
    }
  }
  return false;
}

function vtdApp_setCellByAliases_(sheet, rowNumber, col, aliases, value) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = col[vtdApp_norm_(aliases[i])];
    if (idx != null) {
      sheet.getRange(rowNumber, idx + 1).setValue(value);
      return true;
    }
  }
  return false;
}

function vtdApp_norm_(value) {
  return String(value == null ? "" : value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[đĐ]/g, "d")
    .replace(/\u00C4\u0091/g, "d")
    .replace(/[^\w/]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function vtdApp_dateKey_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "";
    return Utilities.formatDate(value, Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
  }
  const text = String(value == null ? "" : value).trim();
  if (!text) return "";
  let m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+.*)?$/);
  if (m) return m[3] + "-" + String(Number(m[2])).padStart(2, "0") + "-" + String(Number(m[1])).padStart(2, "0");
  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$/);
  if (m) return m[1] + "-" + String(Number(m[2])).padStart(2, "0") + "-" + String(Number(m[3])).padStart(2, "0");
  m = text.match(/^(\d{1,2})-(\d{1,2})-(\d{4})(?:\s+.*)?$/);
  if (m) return m[3] + "-" + String(Number(m[2])).padStart(2, "0") + "-" + String(Number(m[1])).padStart(2, "0");
  const d = new Date(text);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
}

function vtdApp_toDisplay_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm");
  }
  return value || "";
}

function vtdApp_toPairs_(obj) {
  return Object.keys(obj)
    .sort()
    .map(key => ({key: key, value: obj[key]}));
}

function vtdApp_makeId_() {
  return "WEB-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh", "yyyyMMdd-HHmmss") + "-" + Utilities.getUuid().slice(0, 8);
}

function vtdApp_safeFileName_(fileName, rawId, maDon) {
  const clean = String(fileName || "image.jpg").replace(/[\\/:*?"<>|]+/g, "_");
  const prefix = String(maDon || rawId || "CT").replace(/[\\/:*?"<>|]+/g, "_");
  return prefix + "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh", "yyyyMMdd_HHmmss") + "_" + clean;
}

function vtdApp_getUser_() {
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (err) {
    return "";
  }
}


function vtdApp_loginReportSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.loginReportSheet || "_VTD_LOGIN_REPORT");
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.loginReportSheet || "_VTD_LOGIN_REPORT");
    sh.appendRow(["Ngày", "User", "Số lần đăng nhập", "Mốc thời gian đăng nhập", "Thời gian hoạt động app (giây)", "Device", "UpdatedAt"]);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow(["Ngày", "User", "Số lần đăng nhập", "Mốc thời gian đăng nhập", "Thời gian hoạt động app (giây)", "Device", "UpdatedAt"]);
  }
  return sh;
}

function vtdApp_findLoginReportRow_(sh, dayKey, email) {
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 0;
  const values = sh.getRange(2, 1, lastRow - 1, Math.max(2, sh.getLastColumn())).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === dayKey && String(values[i][1] || "").toLowerCase().trim() === email) return i + 2;
  }
  return 0;
}

function vtdApp_recordLogin_(email, deviceId) {
  email = String(email || "").toLowerCase().trim();
  if (!email) return;
  const now = new Date();
  const dayKey = vtdApp_vnDayKey_(now);
  const timeText = Utilities.formatDate(now, "Asia/Saigon", "HH:mm:ss");
  const updatedAt = Utilities.formatDate(now, "Asia/Saigon", "yyyy-MM-dd HH:mm:ss");
  const sh = vtdApp_loginReportSheet_();
  const rowNo = vtdApp_findLoginReportRow_(sh, dayKey, email);
  if (!rowNo) {
    sh.appendRow([dayKey, email, 1, timeText, 0, deviceId || "", updatedAt]);
    return;
  }
  const row = sh.getRange(rowNo, 1, 1, 7).getValues()[0];
  const count = Number(row[2] || 0) + 1;
  const times = String(row[3] || "").trim();
  sh.getRange(rowNo, 3, 1, 5).setValues([[count, times ? times + ", " + timeText : timeText, Number(row[4] || 0), deviceId || row[5] || "", updatedAt]]);
}

function vtdApp_activityPing(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const now = new Date();
  const dayKey = vtdApp_vnDayKey_(now);
  const updatedAt = Utilities.formatDate(now, "Asia/Saigon", "yyyy-MM-dd HH:mm:ss");
  const email = String(auth.email || "").toLowerCase().trim();
  const sh = vtdApp_loginReportSheet_();
  let rowNo = vtdApp_findLoginReportRow_(sh, dayKey, email);
  const incomingSeconds = Math.max(0, Number(params.activeSeconds || 0));
  if (!rowNo) {
    sh.appendRow([dayKey, email, 0, "", incomingSeconds, String(params.deviceId || ""), updatedAt]);
    rowNo = sh.getLastRow();
  } else {
    const row = sh.getRange(rowNo, 1, 1, 7).getValues()[0];
    const seconds = Math.max(0, Number(row[4] || 0)) + incomingSeconds;
    sh.getRange(rowNo, 5, 1, 3).setValues([[seconds, String(params.deviceId || row[5] || ""), updatedAt]]);
  }
  return vtdApp_ok_({activeSeconds: incomingSeconds});
}

function vtdApp_loginReport(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  if (String(auth.email || "").toLowerCase().trim() !== "khiempham5209@gmail.com") return vtdApp_fail_("Bạn không có quyền xem báo cáo đăng nhập.");
  const dayKey = String(params.date || vtdApp_vnDayKey_(new Date())).trim();
  const sh = vtdApp_loginReportSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return vtdApp_ok_({rows: []});
  const values = sh.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
  const rows = values.filter(r => String(r[0] || "").trim() === dayKey).map(r => ({
    date: r[0],
    user: r[1],
    loginCount: Number(r[2] || 0),
    loginTimes: r[3],
    activeSeconds: Number(r[4] || 0),
    device: r[5],
    updatedAt: r[6]
  }));
  return vtdApp_ok_({rows: rows});
}

function vtdApp_auth_(params) {
  params = params || {};
  const sessionInfo = params.sessionToken ? vtdApp_sessionEmail_(params.sessionToken) : null;
  const tokenInfo = params.authToken ? vtdApp_verifyGoogleToken_(params.authToken) : null;
  if (VTD_APP.requireGoogleLogin && VTD_APP.googleClientId && !tokenInfo) {
    return {
      allowed: false,
      email: "",
      role: "blocked",
      screens: [],
      actions: [],
      message: "Vui lòng đăng nhập Google để truy cập."
    };
  }
  if (VTD_APP.requireGoogleLogin && !VTD_APP.googleClientId) {
    return {
      allowed: false,
      email: "",
      role: "blocked",
      screens: [],
      actions: [],
      message: "Chưa cấu hình Google Client ID cho màn đăng nhập."
    };
  }
  if (VTD_APP.requirePinLogin && !sessionInfo && !tokenInfo) {
    return {
      allowed: false,
      email: "",
      role: "blocked",
      screens: [],
      actions: [],
      message: "Vui lòng đăng nhập để truy cập."
    };
  }
  const email = String((sessionInfo && sessionInfo.email) || (tokenInfo && tokenInfo.email) || vtdApp_getUser_() || "").toLowerCase().trim();
  const config = vtdApp_permissionConfig_();
  const adminEmails = config.adminEmails.map(e => String(e).toLowerCase().trim());
  if (adminEmails.indexOf(email) >= 0) {
    return {
      allowed: true,
      email: email,
      role: "admin",
      screens: VTD_PERMISSIONS.allScreens.slice(),
      actions: VTD_PERMISSIONS.allActions.slice()
    };
  }

  const user = config.users[email];
  if (user) {
    return {
      allowed: true,
      email: email,
      role: user.role || "user",
      screens: (user.screens || []).slice(),
      actions: (user.actions || []).slice()
    };
  }

  return {
    allowed: false,
    email: email,
    role: "blocked",
    screens: [],
    actions: [],
    message: "Bạn không có quyền truy cập."
  };
}

function vtdApp_authByPin_(email, pin) {
  const config = vtdApp_permissionConfig_();
  const adminEmails = config.adminEmails.map(e => String(e).toLowerCase().trim());
  if (adminEmails.indexOf(email) >= 0) {
    const storedPin = config.pins[email] || VTD_APP.adminBootstrapPin;
    if (pin !== storedPin) return {allowed: false};
    return {
      allowed: true,
      email: email,
      role: "admin",
      screens: VTD_PERMISSIONS.allScreens.slice(),
      actions: VTD_PERMISSIONS.allActions.slice()
    };
  }

  const user = config.users[email];
  if (!user || pin !== config.pins[email]) return {allowed: false};
  return {
    allowed: true,
    email: email,
    role: user.role || "user",
    screens: (user.screens || []).slice(),
    actions: (user.actions || []).slice()
  };
}

function vtdApp_sessionEmail_(token) {
  try {
    const raw = CacheService.getScriptCache().get("VTD_SESSION_" + token);
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    return null;
  }
}

function vtdApp_permissionConfig_() {
  const cached = CacheService.getScriptCache().get("VTD_PERMISSION_CONFIG");
  if (cached) return JSON.parse(cached);

  const config = {
    adminEmails: VTD_PERMISSIONS.adminEmails.map(e => String(e).toLowerCase().trim()),
    users: {},
    pins: {}
  };

  Object.keys(VTD_PERMISSIONS.users).forEach(email => {
    const key = String(email).toLowerCase().trim();
    config.users[key] = {
      role: VTD_PERMISSIONS.users[email].role || "user",
      screens: (VTD_PERMISSIONS.users[email].screens || []).slice(),
      actions: (VTD_PERMISSIONS.users[email].actions || []).slice()
    };
  });

  try {
    const sh = vtdApp_ensurePermissionSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length > 1) {
      const col = vtdApp_headerMap_(values[0]);
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const active = String(row[col.active] || "TRUE").toUpperCase();
        if (active === "FALSE" || active === "NO" || active === "0") continue;
        const email = String(row[col.email] || "").toLowerCase().trim();
        if (!email) continue;
        const role = String(row[col.role] || "user").toLowerCase().trim();
        const screens = vtdApp_splitCsv_(row[col.screens]);
        const actions = vtdApp_splitCsv_(row[col.actions]);
        const pin = col.pin == null ? "" : String(row[col.pin] || "").trim();
        if (pin) config.pins[email] = pin;
        if (role === "admin" && config.adminEmails.indexOf(email) < 0) {
          config.adminEmails.push(email);
        } else {
          config.users[email] = {
            role: role || "user",
            screens: screens.length ? screens : ["home", "search", "input", "store"],
            actions: actions.length ? actions : ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"]
          };
        }
      }
    }
  } catch (err) {}

  CacheService.getScriptCache().put("VTD_PERMISSION_CONFIG", JSON.stringify(config), 120);
  return config;
}

function vtdApp_ensureUserSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.userSheet);
  const header = ["Email", "Role", "Pass", "Active", "Screens", "Actions", "Updated_By", "Updated_At"];
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.userSheet);
    sh.appendRow(header);
    Object.keys(VTD_PERMISSIONS.users).sort().forEach(email => {
      const user = VTD_PERMISSIONS.users[email];
      sh.appendRow([
        String(email).toLowerCase().trim(),
        user.role || "staff",
        "",
        "TRUE",
        (user.screens || []).join(","),
        (user.actions || []).join(","),
        "seed",
        new Date()
      ]);
    });
  }
  if (sh.getLastRow() === 0) sh.appendRow(header);
  vtdApp_ensurePermissionHeader_(sh, header);
  return sh;
}

function vtdApp_ensurePermissionSheet_() {
  return vtdApp_ensureUserSheet_();
}

function vtdApp_ensurePermissionHeader_(sh, requiredHeader) {
  const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const map = vtdApp_headerMap_(current);
  requiredHeader.forEach(name => {
    if (map[vtdApp_norm_(name)] == null) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(name);
    }
  });
}

function vtdApp_splitCsv_(value) {
  return String(value || "")
    .split(",")
    .map(x => String(x).trim())
    .filter(Boolean);
}

function vtdApp_requireAction_(action, params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const ownerEmails = (VTD_PERMISSIONS.adminEmails || []).map(e => String(e || "").toLowerCase().trim());
  const authEmail = String(auth.email || "").toLowerCase().trim();
  if (ownerEmails.indexOf(authEmail) >= 0) return null;
  const actionAliases = {
    adminConfig: ["adminConfig", "manageConfig"],
    manageConfig: ["manageConfig", "adminConfig"]
  };
  const allowedActions = actionAliases[action] || [action];
  if (!allowedActions.some(a => auth.actions.indexOf(a) >= 0)) {
    return vtdApp_fail_("Bạn không có quyền thực hiện thao tác này: " + action);
  }
  return null;
}

function vtdApp_verifyGoogleToken_(idToken) {
  if (!VTD_APP.googleClientId) return null;
  try {
    const url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken);
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    if (response.getResponseCode() !== 200) return null;
    const data = JSON.parse(response.getContentText());
    if (data.aud !== VTD_APP.googleClientId) return null;
    if (String(data.email_verified) !== "true") return null;
    return {email: data.email};
  } catch (err) {
    return null;
  }
}

function vtdApp_findImages_(rawId, maDon) {
  const ss = vtdApp_ss_();
  const sheetNames = [VTD_APP.imageSheet, VTD_APP.imageSheetAlt, VTD_APP.imageFullSheet];
  const out = [];
  const seen = {};
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const lastCol = sheet.getLastColumn();
    const values = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    const header = values[0];
    const col = vtdApp_headerMap_(header);
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowRawId = vtdApp_pickFromArray_(row, col, ["raw_id", "raw id", "id raw"]);
      const rowMaDon = vtdApp_pickFromArray_(row, col, ["ma don"]);
      const matchRawId = rawId && String(rowRawId || "").trim() === String(rawId).trim();
      const matchMaDon = maDon && String(rowMaDon || "").trim() === String(maDon).trim();
      if (!matchRawId && !matchMaDon) continue;
      const path = vtdApp_pickFromArray_(row, col, ["link anh", "link anh/video", "file url", "folder url", "anh"]);
      const key = String(path || name + ":" + (i + 1));
      if (seen[key]) continue;
      seen[key] = true;
      out.push({
        source: name,
        rowNumber: i + 1,
        rawId: rowRawId,
        maDon: rowMaDon,
        name: String(path || "").split("/").pop(),
        rawPath: path,
        url: vtdApp_resolveImageUrl_(path)
      });
      if (out.length >= 80) return;
    }
  });
  return out;
}

function vtdApp_getRecord(params) {
  const denied = vtdApp_requireAction_("getRecord", params);
  if (denied) return denied;
  params = params || {};
  const ss = vtdApp_ss_();

  if (params.source && params.rowNumber) {
    const found = vtdApp_getRecordByRow_(ss.getSheetByName(String(params.source)), Number(params.rowNumber));
    if (found) return vtdApp_ok_(vtdApp_attachRecordDetails_(found));
  }

  const key = String(params.id || params.maDon || params.query || "").trim();
  if (!key) return vtdApp_fail_("Thiếu ID hoặc mã đơn.");
  const sheets = [ss.getSheetByName(VTD_APP.rawSheet), ss.getSheetByName(VTD_APP.rawFullSheet)];
  for (let s = 0; s < sheets.length; s++) {
    const found = vtdApp_findRecordInSheet_(sheets[s], key);
    if (found) return vtdApp_ok_(vtdApp_attachRecordDetails_(found));
  }
  return vtdApp_fail_("Không tìm thấy chứng từ.");
}

function vtdApp_getRecordByRow_(sheet, rowNumber) {
  if (!sheet || rowNumber <= 1 || rowNumber > sheet.getLastRow()) return null;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const row = {};
  headers.forEach((header, i) => row[vtdApp_norm_(header)] = values[i]);
  const item = {
    source: sheet.getName(),
    rowNumber: rowNumber,
    headers: headers,
    values: values,
    row: row
  };
  return {
    source: item.source,
    rowNumber: item.rowNumber,
    headers: item.headers,
    values: item.values,
    row: item.row,
    compact: vtdApp_compactRecord_(item)
  };
}

function vtdApp_attachRecordDetails_(found) {
  found.images = vtdApp_findImages_(found.compact.id, found.compact.maDon);
  vtdApp_enrichCompactsFromFullCtByOrders_([found.compact]);
  vtdApp_attachSkuItemsToCompacts_(vtdApp_ss_(), [found.compact]);
  found.sku = found.compact.skuItems || vtdApp_findSku_(found.compact.id);
  found.fields = vtdApp_recordFields_(found);
  return found;
}

function vtdApp_saveRaw(params) {
  const denied = vtdApp_requireAction_("saveRaw", params);
  if (denied) return denied;
  params = params || {};
  const ss = vtdApp_ss_();
  const sheet = ss.getSheetByName(VTD_APP.rawSheet);
  if (!sheet) return vtdApp_fail_("Không thấy sheet " + VTD_APP.rawSheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill("");
  const col = vtdApp_headerMap_(headers);
  const id = params.id || params.ID || vtdApp_makeId_();
  const auth = vtdApp_auth_(params);

  vtdApp_setByAliases_(row, col, ["id"], id);
  const maDon = vtdApp_cleanMaDonParam_(params);
  const xacThuc = String(params.xacThucHoaDon || params.xacThucChungTu || params.xacThuc || "").trim();
  vtdApp_setByAliases_(row, col, ["ma don", "ma don ghn", "ma don ghtk"], maDon);
  vtdApp_setByAliases_(row, col, ["slog"], params.slog);
  vtdApp_setByAliases_(row, col, ["xac thuc hoa don", "xac thuc chung tu"], xacThuc);
  vtdApp_setByAliases_(row, col, ["loai hoan"], params.loaiHoan);
  vtdApp_setByAliases_(row, col, ["ngay hoan tra"], params.ngayHoanTra || new Date());
  vtdApp_setByAliases_(row, col, ["thoi gian nhap hoan", "timestamp"], new Date());
  vtdApp_setByAliases_(row, col, ["ma po chung tu hoan", "ma po"], params.po);
  vtdApp_setByAliases_(row, col, ["tinh trang chung tu"], params.tinhTrang);
  vtdApp_setByAliases_(row, col, ["thong tin chung tu"], params.thongTin);
  vtdApp_setByAliases_(row, col, ["user_email", "nguoi nhap", "email"], auth.email || vtdApp_getUser_());

  sheet.appendRow(row);
  const rowNumber = sheet.getLastRow();
  const skuCount = vtdApp_saveSkuRows_(ss, id, maDon, params.skuItems || [], auth);
  vtdApp_queueRow_(sheet, rowNumber);
  return vtdApp_ok_({id: id, maDon: maDon, rowNumber: rowNumber, skuCount: skuCount, message: "Đã lưu chứng từ" + (skuCount ? " và " + skuCount + " SKU." : ".")});
}

function vtdApp_skuNameFrom_(item) {
  item = item || {};
  return String(
    item.ten ||
    item.name ||
    item.productName ||
    item.tenSanPham ||
    item.tenSanPhamHoanVe ||
    item.skuName ||
    item.sku ||
    item.product ||
    item.itemName ||
    item["Tên sản phẩm hoàn về"] ||
    item["Ten san pham hoan ve"] ||
    item["Tên sản phẩm"] ||
    item["Ten san pham"] ||
    item["ten san pham"] ||
    item["ten_san_pham"] ||
    item["ten_san_pham_hoan_ve"] ||
    ""
  ).trim();
}

function vtdApp_skuCodeFrom_(item) {
  item = item || {};
  return String(item.ma || item.code || item.material || item.Material || item["Material / Mã SP"] || item["Mã vật tư"] || "").trim();
}

function vtdApp_normalizeSkuItem_(item) {
  item = item || {};
  const ten = vtdApp_skuNameFrom_(item);
  const ma = vtdApp_skuCodeFrom_(item);
  return Object.assign({}, item, {
    ten: ten,
    name: item.name || ten,
    productName: item.productName || ten,
    tenSanPham: item.tenSanPham || ten,
    skuName: item.skuName || ten,
    ma: item.ma || ma,
    code: item.code || ma
  });
}

function vtdApp_saveSkuRows_(ss, rawId, maDon, items, auth) {
  if (!Array.isArray(items) || !items.length) return 0;
  let sheet = ss.getSheetByName(VTD_APP.skuSheet);
  if (!sheet) {
    sheet = ss.insertSheet(VTD_APP.skuSheet);
    sheet.appendRow(["ID", "RAW_ID", "Tên sản phẩm hoàn về", "Type", "Số lượng", "Tình trạng", "Mã lô/HSD", "Mã serial", "Ảnh sản phẩm", "Ghi chú", "User_email", "Ngày", "Giờ"]);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  let count = 0;
  items.forEach(item => {
    item = vtdApp_normalizeSkuItem_(item);
    const skuName = vtdApp_skuNameFrom_(item);
    if (!skuName) return;
    const now = new Date();
    const ngayText = Utilities.formatDate(now, "Asia/Saigon", "dd/MM/yyyy");
    const gioText = Utilities.formatDate(now, "Asia/Saigon", "HH:mm:ss");
    const row = new Array(headers.length).fill("");
    vtdApp_setByAliases_(row, col, ["id"], item.id || vtdApp_makeId_());
    vtdApp_setByAliases_(row, col, ["raw_id", "raw id"], rawId);
    vtdApp_setByAliases_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"], maDon);
    vtdApp_setByAliases_(row, col, ["ten san pham hoan", "tên sản phẩm hoàn về", "ten san pham", "tên sản phẩm", "sku"], skuName);
    vtdApp_setByAliases_(row, col, ["type", "loai", "loại"], item.type);
    vtdApp_setByAliases_(row, col, ["so luong", "số lượng", "qty"], item.soLuong || item.qty);
    vtdApp_setByAliases_(row, col, ["tinh trang", "tình trạng"], item.tinhTrang || item.condition);
    vtdApp_setByAliases_(row, col, ["ma lo hsd", "mã lô/hsd", "ma lo", "mã lô", "hsd"], item.maLoHsd || item.lot);
    vtdApp_setByAliases_(row, col, ["ma serial", "mã serial", "serial"], item.serial || item.maSerial);
    vtdApp_setByAliases_(row, col, ["anh san pham","Ảnh sản phẩm", "ảnh sản phẩm", "anh", "ảnh"], item.anhSanPham || item.imagePath || item.imageUrl);
    vtdApp_setByAliases_(row, col, ["ghi chu", "ghi chú", "note"], item.ghiChu || item.note);
    vtdApp_setByAliases_(row, col, ["user_email", "email", "nguoi nhap", "người nhập"], (auth && auth.email) || vtdApp_getUser_());
    vtdApp_setByAliases_(row, col, ["ngay", "ngày", "date"], ngayText);
    vtdApp_setByAliases_(row, col, ["gio", "giờ", "thoi gian", "thời gian", "timestamp", "time"], gioText);
    sheet.appendRow(row);
    count++;
  });
  return count;
}

function vtdApp_syncNewMasterSku_(ss, items) {
  if (!Array.isArray(items) || !items.length) return 0;
  let sheet = ss.getSheetByName("MASTER SKU");
  if (!sheet) {
    sheet = ss.insertSheet("MASTER SKU");
    sheet.appendRow(["Mã vật tư", "Tên sản phẩm", "Type"]);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues() : [];
  const seen = {};
  values.forEach(row => {
    const ma = String(vtdApp_pickFromArray_(row, col, ["ma vat tu", "mã vật tư", "ma sku", "sku", "id"]) || "").trim();
    const ten = String(vtdApp_pickFromArray_(row, col, ["ten san pham", "tên sản phẩm", "ten sku", "san pham"]) || "").trim();
    if (ma || ten) seen[vtdApp_norm_(ma + "|" + ten)] = true;
  });
  let count = 0;
  items.forEach(item => {
    item = vtdApp_normalizeSkuItem_(item);
    if (!item.isNew) return;
    const ma = vtdApp_skuCodeFrom_(item);
    const ten = vtdApp_skuNameFrom_(item);
    if (!ten) return;
    const key = vtdApp_norm_(ma + "|" + ten);
    if (seen[key]) return;
    const row = new Array(headers.length).fill("");
    vtdApp_setByAliases_(row, col, ["ma vat tu", "mã vật tư", "ma sku", "sku", "id"], ma);
    vtdApp_setByAliases_(row, col, ["ten san pham", "tên sản phẩm", "ten sku", "san pham"], ten);
    vtdApp_setByAliases_(row, col, ["loai", "loại", "type"], item.type || "");
    if (!row.some(v => v !== "")) {
      row[0] = ma;
      row[1] = ten;
      row[2] = item.type || "";
    }
    sheet.appendRow(row);
    seen[key] = true;
    count++;
  });
  return count;
}

function vtdApp_uploadImages(params) {
  const denied = vtdApp_requireAction_("uploadImage", params);
  if (denied) return denied;
  params = params || {};
  const files = Array.isArray(params.files) ? params.files : [];
  if (!files.length) return vtdApp_fail_("Chưa chọn ảnh.");
  const results = [];
  const errors = [];
  for (let i = 0; i < files.length; i++) {
    const file = files[i] || {};
    const one = Object.assign({}, params, file, {
      fileName: file.fileName || file.name || params.fileName,
      mimeType: file.mimeType || file.type || params.mimeType,
      base64: file.base64 || file.data || params.base64
    });
    const res = vtdApp_uploadImage(one);
    if (res && res.ok) results.push(res);
    else errors.push((file.fileName || file.name || ("ảnh " + (i + 1))) + ": " + (res && res.message || "Upload lỗi"));
  }
  if (errors.length) return vtdApp_fail_("Upload ảnh lỗi: " + errors.join("; "));
  if (results.length !== files.length) return vtdApp_fail_("Upload ảnh thiếu: " + results.length + "/" + files.length);
  return vtdApp_ok_({count: results.length, files: results, message: "Đã upload " + results.length + " ảnh."});
}

function vtdApp_ok_(data) {
  data = data || {};
  data.ok = true;
  return data;
}

function vtdApp_fail_(message) {
  return {ok: false, message: message};
}

/************************************************
VTD LIVE BRIDGE - 2026-05-27

Keep this block at the end of Code.gs. Apps Script keeps the last function
definition when names are duplicated, so these functions override the older
prototype functions above without touching PJ1/PJ2/dispatcher logic.
************************************************/

VTD_APP.spreadsheetId = "1rrg7ZHuBiJtX91y7C_HDwQKDDPmwJ5GuMxg1UT31uqQ";
VTD_APP.fullCtSpreadsheetId = "1S9XbCRRQutwBoAhrKJ8awdy-tj54cEWRswy8mXzWY_o";
VTD_APP.fileBSpreadsheetId = "1Mv_nkb0dx7t82kRusrBGGqD37UYxS6PqE24I7D0Oxfk";
VTD_APP.fileCSpreadsheetId = "1Z3sIdjjPFwZtwkGtOUiAJ5y2TEcyBEWB1Omi2T4H3YA";
VTD_APP.driveCtFolderId = "1XUiULXcfW38zom_0-HnZxPM-ikEA-rug";
VTD_APP.driveSkuFolderId = "1aj6g1IekBawP_xFpy0BByUOFz-iuH2cM";
VTD_APP.userSheet = "VTD_USERS";
VTD_APP.indexPrefix = "VTD_SEARCH_INDEX_";
VTD_APP.rowMapSheet = "VTD_ROW_MAP";
VTD_APP.recordCacheSheet = "VTD_RECORD_CACHE";
VTD_APP.searchTokenSheet = "VTD_SEARCH_TOKEN";
VTD_APP.lookupCacheStatusSheet = "VTD_LOOKUP_CACHE_STATUS";
VTD_APP.queueErrorSheet = "VTD_QUEUE_ERRORS";
VTD_APP.queueRetrySheet = "VTD_QUEUE_RETRY";
VTD_APP.dynamicConfigSheet = "VTD_DYNAMIC_CONFIG";
VTD_APP.dashboardCacheSheet = "VTD_DASHBOARD_CACHE";
VTD_APP.apiLogSheet = "VTD_API_LOG";
VTD_APP.uploadQueueSheet = "VTD_UPLOAD_QUEUE";
VTD_APP.useRecordCache = false;
VTD_APP.dashboardRows = 5000;
VTD_APP.maxSearchRows = 20000;
VTD_APP.defaultLimit = 1000;

function vtdApp_publicConfig() {
  return vtdApp_ok_({
    appVersion: "1.5",
    googleClientId: "",
    requireGoogleLogin: false,
    requirePinLogin: true,
    authMode: "pin",
    sheetId: VTD_APP.spreadsheetId
  });
}

function vtdApp_login(params) {
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  const pass = String(params.pin || params.pass || "").trim();
  const appVersion = String(params.appVersion || "").trim();
  if (!email || !pass) return vtdApp_fail_("Nhap email va Pass.");

  const config = vtdApp_permissionConfig_();
  const currentPass = String(config.pins[email] || "").trim();
  const auth = config.adminEmails.indexOf(email) >= 0 ? {
    allowed: true,
    email: email,
    role: "admin",
    screens: VTD_PERMISSIONS.allScreens.slice(),
    actions: VTD_PERMISSIONS.allActions.slice()
  } : config.users[email];

  if (!auth || !auth.allowed && auth.active === false) return vtdApp_fail_("Bạn không có quyền truy cập.");
  if (!currentPass || currentPass !== pass) return vtdApp_fail_("Email hoặc Pass không đúng.");
  const passwordPolicy = vtdApp_markPasswordFirstLogin_(email, appVersion);
  try { vtdApp_recordLogin_(email, String(params.deviceId || "")); } catch (err) { vtdApp_log_("WARN", "loginReport", email, "", "", "ERR", String(err), ""); }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("VTD_SESSION_" + token, JSON.stringify({
    email: email,
    createdAt: Date.now(),
    dayKey: vtdApp_vnDayKey_(new Date())
  }), vtdApp_secondsToMidnight_());

  return vtdApp_ok_({
    token: token,
    user: email,
    auth: Object.assign(vtdApp_authShape_(email, auth), {passwordPolicy: passwordPolicy}),
    passwordPolicy: passwordPolicy
  });
}


function vtdApp_loginReportSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.loginReportSheet || "_VTD_LOGIN_REPORT");
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.loginReportSheet || "_VTD_LOGIN_REPORT");
    sh.appendRow(["Ngày", "User", "Số lần đăng nhập", "Mốc thời gian đăng nhập", "Thời gian hoạt động app (giây)", "Device", "UpdatedAt"]);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow(["Ngày", "User", "Số lần đăng nhập", "Mốc thời gian đăng nhập", "Thời gian hoạt động app (giây)", "Device", "UpdatedAt"]);
  }
  return sh;
}

function vtdApp_findLoginReportRow_(sh, dayKey, email) {
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 0;
  const values = sh.getRange(2, 1, lastRow - 1, Math.max(2, sh.getLastColumn())).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === dayKey && String(values[i][1] || "").toLowerCase().trim() === email) return i + 2;
  }
  return 0;
}

function vtdApp_recordLogin_(email, deviceId) {
  email = String(email || "").toLowerCase().trim();
  if (!email) return;
  const now = new Date();
  const dayKey = vtdApp_vnDayKey_(now);
  const timeText = Utilities.formatDate(now, "Asia/Saigon", "HH:mm:ss");
  const updatedAt = Utilities.formatDate(now, "Asia/Saigon", "yyyy-MM-dd HH:mm:ss");
  const sh = vtdApp_loginReportSheet_();
  const rowNo = vtdApp_findLoginReportRow_(sh, dayKey, email);
  if (!rowNo) {
    sh.appendRow([dayKey, email, 1, timeText, 0, deviceId || "", updatedAt]);
    return;
  }
  const row = sh.getRange(rowNo, 1, 1, 7).getValues()[0];
  const count = Number(row[2] || 0) + 1;
  const times = String(row[3] || "").trim();
  sh.getRange(rowNo, 3, 1, 5).setValues([[count, times ? times + ", " + timeText : timeText, Number(row[4] || 0), deviceId || row[5] || "", updatedAt]]);
}

function vtdApp_activityPing(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const now = new Date();
  const dayKey = vtdApp_vnDayKey_(now);
  const updatedAt = Utilities.formatDate(now, "Asia/Saigon", "yyyy-MM-dd HH:mm:ss");
  const email = String(auth.email || "").toLowerCase().trim();
  const sh = vtdApp_loginReportSheet_();
  let rowNo = vtdApp_findLoginReportRow_(sh, dayKey, email);
  const incomingSeconds = Math.max(0, Number(params.activeSeconds || 0));
  if (!rowNo) {
    sh.appendRow([dayKey, email, 0, "", incomingSeconds, String(params.deviceId || ""), updatedAt]);
    rowNo = sh.getLastRow();
  } else {
    const row = sh.getRange(rowNo, 1, 1, 7).getValues()[0];
    const seconds = Math.max(0, Number(row[4] || 0)) + incomingSeconds;
    sh.getRange(rowNo, 5, 1, 3).setValues([[seconds, String(params.deviceId || row[5] || ""), updatedAt]]);
  }
  return vtdApp_ok_({activeSeconds: incomingSeconds});
}

function vtdApp_loginReport(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  if (String(auth.email || "").toLowerCase().trim() !== "khiempham5209@gmail.com") return vtdApp_fail_("Bạn không có quyền xem báo cáo đăng nhập.");
  const dayKey = String(params.date || vtdApp_vnDayKey_(new Date())).trim();
  const sh = vtdApp_loginReportSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return vtdApp_ok_({rows: []});
  const values = sh.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
  const rows = values.filter(r => String(r[0] || "").trim() === dayKey).map(r => ({
    date: r[0],
    user: r[1],
    loginCount: Number(r[2] || 0),
    loginTimes: r[3],
    activeSeconds: Number(r[4] || 0),
    device: r[5],
    updatedAt: r[6]
  }));
  return vtdApp_ok_({rows: rows});
}

function vtdApp_auth_(params) {
  params = params || {};
  const token = String(params.sessionToken || "").trim();
  const emailFromToken = token ? vtdApp_sessionEmail_(token) : "";
  const email = String(emailFromToken || params.authEmail || params.currentEmail || params.email || "").toLowerCase().trim();
  if (!email) return {allowed: false, message: "Vui lòng đăng nhập để truy cập."};

  const config = vtdApp_permissionConfig_();
  if (config.adminEmails.indexOf(email) >= 0) {
    return vtdApp_authShape_(email, {
      role: "admin",
      screens: VTD_PERMISSIONS.allScreens.slice(),
      actions: VTD_PERMISSIONS.allActions.slice()
    });
  }
  if (config.users[email]) return vtdApp_authShape_(email, config.users[email]);
  return {allowed: false, message: "Bạn không có quyền truy cập."};
}

function vtdApp_sessionEmail_(token) {
  try {
    const raw = CacheService.getScriptCache().get("VTD_SESSION_" + token);
    if (!raw) return "";
    const data = JSON.parse(raw);
    if (data.dayKey && data.dayKey !== vtdApp_vnDayKey_(new Date())) return "";
    return String(data.email || "").toLowerCase().trim();
  } catch (err) {
    return "";
  }
}

function vtdApp_permissionConfig_() {
  const cached = CacheService.getScriptCache().get("VTD_USERS_CONFIG_LIVE");
  if (cached) return JSON.parse(cached);

  const config = {
    adminEmails: [],
    users: {},
    pins: {}
  };

  VTD_PERMISSIONS.adminEmails.forEach(email => {
    const key = String(email).toLowerCase().trim();
    if (key && config.adminEmails.indexOf(key) < 0) config.adminEmails.push(key);
    if (key && !config.pins[key]) config.pins[key] = VTD_APP.adminBootstrapPin;
  });

  const ss = vtdApp_ss_();
  const sh = ss.getSheetByName(VTD_APP.userSheet);
  if (sh && sh.getLastRow() > 1) {
    const values = sh.getDataRange().getValues();
    const col = vtdApp_headerMap_(values[0]);
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const email = String(vtdApp_pickFromArray_(row, col, ["email"]) || "").toLowerCase().trim();
      if (!email) continue;
      const active = String(vtdApp_pickFromArray_(row, col, ["active"]) || "TRUE").toUpperCase();
      if (active === "FALSE" || active === "NO" || active === "0") continue;
      const role = String(vtdApp_pickFromArray_(row, col, ["role"]) || "staff").toLowerCase().trim();
      const pass = String(vtdApp_pickFromArray_(row, col, ["pass", "pin"]) || "").trim();
      const screens = vtdApp_splitCsv_(vtdApp_pickFromArray_(row, col, ["screens"]));
      const actions = vtdApp_splitCsv_(vtdApp_pickFromArray_(row, col, ["actions"]));
      if (pass) config.pins[email] = pass;
      config.users[email] = {
        allowed: true,
        role: role === "admin" ? "admin" : "staff",
        screens: screens.length ? screens : (role === "admin" ? [] : ["home", "search", "input", "store"]),
        actions: actions.length ? actions : (role === "admin" ? [] : ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"])
      };
    }
  }

  CacheService.getScriptCache().put("VTD_USERS_CONFIG_LIVE", JSON.stringify(config), 120);
  return config;
}

function vtdApp_changePin(params) {
  params = params || {};
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_("Bạn không có quyền truy cập.");
  const oldPass = String(params.oldPin || params.oldPass || params.currentPass || "").trim();
  const newPass = String(params.newPin || params.newPass || params.pass || "").trim();
  const ruleError = vtdApp_validateStrongPass_(newPass, oldPass);
  if (ruleError) return vtdApp_fail_(ruleError);

  const config = vtdApp_permissionConfig_();
  const email = String(auth.email || "").toLowerCase().trim();
  const currentPass = String(config.pins[email] || "").trim();
  if (!currentPass || currentPass !== oldPass) return vtdApp_fail_("Pass cũ không đúng.");

  vtdApp_updateUserPass_(email, newPass, email);
  vtdApp_markPasswordChanged_(email, String(params.appVersion || "").trim());
  vtdApp_clearUserAuthCache_();
  return vtdApp_ok_({message: "Đã đổi Pass.", email: email, pass: newPass});
}

function vtdApp_changePass(params) {
  return vtdApp_changePin(params);
}

function vtdApp_validateStrongPass_(pass, oldPass) {
  pass = String(pass || "");
  oldPass = String(oldPass || "");
  if (pass.length < 8) return "Pass mới phải có ít nhất 8 ký tự.";
  if (oldPass && pass === oldPass) return "Pass mới phải khác Pass mặc định/Pass cũ.";
  if (!/[A-Z]/.test(pass)) return "Pass mới cần có ít nhất 1 chữ hoa.";
  if (!/[a-z]/.test(pass)) return "Pass mới cần có ít nhất 1 chữ thường.";
  if (!/[0-9]/.test(pass)) return "Pass mới cần có ít nhất 1 số.";
  return "";
}

function vtdApp_userSheetSchema_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.userSheet);
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.userSheet);
    sh.appendRow(["Email", "Role", "Pass", "Active", "Screens", "Actions", "Updated_By", "Updated_At"]);
  }
  let headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  let col = vtdApp_headerMap_(headers);
  const aliases = [
    ["first_login_app_version", "First_Login_App_Version"],
    ["first_login_at", "First_Login_At"],
    ["must_change_pass", "Must_Change_Pass"],
    ["pass_changed_at", "Pass_Changed_At"],
    ["pass_change_count", "Pass_Change_Count"],
    ["pass_changed_app_version", "Pass_Changed_App_Version"]
  ];
  aliases.forEach(pair => {
    if (vtdApp_firstCol_(col, [pair[0], pair[1]]) == null) {
      headers.push(pair[1]);
      sh.getRange(1, headers.length).setValue(pair[1]);
      col = vtdApp_headerMap_(headers);
    }
  });
  return {sheet: sh, headers: headers, col: col};
}

function vtdApp_findUserRow_(sh, col, email) {
  const emailCol = vtdApp_firstCol_(col, ["email"]);
  if (emailCol == null || sh.getLastRow() < 2) return 0;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  email = String(email || "").toLowerCase().trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][emailCol] || "").toLowerCase().trim() === email) return i + 2;
  }
  return 0;
}

function vtdApp_markPasswordFirstLogin_(email, appVersion) {
  appVersion = String(appVersion || "").trim() || "unknown";
  const schema = vtdApp_userSheetSchema_();
  const sh = schema.sheet;
  const col = schema.col;
  const row = vtdApp_findUserRow_(sh, col, email);
  if (!row) return {firstLoginAt: "", ageDays: 0, mustChange: false, deadlineDays: 3};
  const firstVersionCol = vtdApp_firstCol_(col, ["first_login_app_version", "First_Login_App_Version"]);
  const firstAtCol = vtdApp_firstCol_(col, ["first_login_at", "First_Login_At"]);
  const mustCol = vtdApp_firstCol_(col, ["must_change_pass", "Must_Change_Pass"]);
  const changedAtCol = vtdApp_firstCol_(col, ["pass_changed_at", "Pass_Changed_At"]);
  const countCol = vtdApp_firstCol_(col, ["pass_change_count", "Pass_Change_Count"]);
  const changedVersionCol = vtdApp_firstCol_(col, ["pass_changed_app_version", "Pass_Changed_App_Version"]);
  const currentFirstVersion = String(sh.getRange(row, firstVersionCol + 1).getValue() || "").trim();
  const changedVersion = changedVersionCol == null ? "" : String(sh.getRange(row, changedVersionCol + 1).getValue() || "").trim();
  const changedAtValue = changedAtCol == null ? "" : sh.getRange(row, changedAtCol + 1).getValue();
  const changeCount = countCol == null ? 0 : Number(sh.getRange(row, countCol + 1).getValue() || 0);
  const hasChangedPassword = changeCount > 0 || !!changedAtValue;
  let firstAt = sh.getRange(row, firstAtCol + 1).getValue();
  if (!firstAt) {
    firstAt = new Date();
    sh.getRange(row, firstVersionCol + 1).setValue(appVersion);
    sh.getRange(row, firstAtCol + 1).setValue(firstAt);
  } else if (!currentFirstVersion) {
    sh.getRange(row, firstVersionCol + 1).setValue(appVersion);
  }
  const ageDays = Math.max(0, Math.floor((Date.now() - new Date(firstAt).getTime()) / 86400000));
  const mustChange = !hasChangedPassword;
  if (mustCol != null) sh.getRange(row, mustCol + 1).setValue(mustChange ? "TRUE" : "FALSE");
  return {
    firstLoginAt: new Date(firstAt).toISOString(),
    ageDays: ageDays,
    mustChange: mustChange,
    forceChange: mustChange && ageDays >= 3,
    deadlineDays: Math.max(0, 3 - ageDays),
    appVersion: appVersion,
    changedAt: changedAtValue && !isNaN(new Date(changedAtValue).getTime()) ? new Date(changedAtValue).toISOString() : "",
    changeCount: changeCount,
    changedAppVersion: changedVersion
  };
}

function vtdApp_markPasswordChanged_(email, appVersion) {
  const schema = vtdApp_userSheetSchema_();
  const sh = schema.sheet;
  const col = schema.col;
  const row = vtdApp_findUserRow_(sh, col, email);
  if (!row) return;
  const now = new Date();
  const mustCol = vtdApp_firstCol_(col, ["must_change_pass", "Must_Change_Pass"]);
  const changedAtCol = vtdApp_firstCol_(col, ["pass_changed_at", "Pass_Changed_At"]);
  const countCol = vtdApp_firstCol_(col, ["pass_change_count", "Pass_Change_Count"]);
  const changedVersionCol = vtdApp_firstCol_(col, ["pass_changed_app_version", "Pass_Changed_App_Version"]);
  if (mustCol != null) sh.getRange(row, mustCol + 1).setValue("FALSE");
  if (changedAtCol != null) sh.getRange(row, changedAtCol + 1).setValue(now);
  if (changedVersionCol != null) sh.getRange(row, changedVersionCol + 1).setValue(appVersion || "");
  if (countCol != null) {
    const current = Number(sh.getRange(row, countCol + 1).getValue() || 0);
    sh.getRange(row, countCol + 1).setValue(current + 1);
  }
}

function vtdApp_updateUserPass_(email, pass, updatedBy) {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.userSheet);
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.userSheet);
    sh.appendRow(["Email", "Role", "Pass", "Active", "Screens", "Actions", "Updated_By", "Updated_At"]);
  }
  let values = sh.getDataRange().getValues();
  let headers = values[0] || [];
  let col = vtdApp_headerMap_(headers);
  let passCol = vtdApp_firstCol_(col, ["pass", "pin"]);
  if (passCol == null) {
    passCol = headers.length;
    sh.getRange(1, passCol + 1).setValue("Pass");
    headers.push("Pass");
    col = vtdApp_headerMap_(headers);
  }
  let pinCol = vtdApp_firstCol_(col, ["pin"]);
  let emailCol = vtdApp_firstCol_(col, ["email"]);
  if (emailCol == null) {
    emailCol = 0;
    sh.getRange(1, 1).setValue("Email");
    values = sh.getDataRange().getValues();
    headers = values[0] || [];
    col = vtdApp_headerMap_(headers);
  }
  const targetRows = [];
  values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][emailCol] || "").toLowerCase().trim() === email) {
      targetRows.push(i + 1);
    }
  }
  if (!targetRows.length) {
    const row = new Array(sh.getLastColumn()).fill("");
    row[emailCol] = email;
    row[passCol] = pass;
    if (pinCol != null) row[pinCol] = pass;
    vtdApp_setByAliases_(row, col, ["role"], "staff");
    vtdApp_setByAliases_(row, col, ["active"], "TRUE");
    vtdApp_setByAliases_(row, col, ["screens"], "home,search,input,store");
    vtdApp_setByAliases_(row, col, ["actions"], "dashboard,search,getRecord,saveRaw,uploadImage,queueRecord,storeInfo");
    vtdApp_setByAliases_(row, col, ["updated_by", "updated by"], updatedBy || email);
    vtdApp_setByAliases_(row, col, ["updated_at", "updated at"], new Date());
    sh.appendRow(row);
  } else {
    const updatedByCol = vtdApp_firstCol_(col, ["updated_by", "updated by"]);
    const updatedAtCol = vtdApp_firstCol_(col, ["updated_at", "updated at"]);
    targetRows.forEach(targetRow => {
      sh.getRange(targetRow, passCol + 1).setValue(pass);
      if (pinCol != null && pinCol !== passCol) sh.getRange(targetRow, pinCol + 1).setValue(pass);
      if (updatedByCol != null) sh.getRange(targetRow, updatedByCol + 1).setValue(updatedBy || email);
      if (updatedAtCol != null) sh.getRange(targetRow, updatedAtCol + 1).setValue(new Date());
    });
  }
  SpreadsheetApp.flush();
  vtdApp_clearUserAuthCache_();
}

function vtdApp_clearUserAuthCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove("VTD_USERS_CONFIG_LIVE");
  cache.remove("VTD_PERMISSION_CONFIG");
}

function vtdApp_firstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = col[vtdApp_norm_(aliases[i])];
    if (idx != null) return idx;
  }
  return null;
}

function vtdApp_dashboard(params) {
  const denied = vtdApp_requireAction_("dashboard", params);
  if (denied) return denied;

  const ss = vtdApp_ss_();
  const raw = ss.getSheetByName(VTD_APP.rawSheet);
  const today = vtdApp_dateKey_(new Date());
  const rows = vtdApp_readTodayRawM3Objects_(raw, today);
  const bySlog = {};
  const byStatus = {};
  const byUser = {};
  const syncErrorsByTarget = {
    hangHoan: 0,
    chungTu: 0
  };
  const recent = [];
  const seen = {};
  rows.forEach(item => {
    const compact = vtdApp_compactRecord_(item);
    const dedupeKey = compact.id || compact.maDon || item.source + ":" + item.rowNumber;
    if (seen[dedupeKey]) return;
    seen[dedupeKey] = true;
    const dateKey = vtdApp_dateKey_(
      (item.displayValues && item.displayValues[1]) ||
      vtdApp_pick_(item.displayRow, ["ngay hoan tra", "ngày hoàn trả"]) ||
      vtdApp_pick_(item.row, ["ngay hoan tra", "ngày hoàn trả"])
    );
    if (dateKey !== today) return;
    const slog = String(compact.slog || "Khac").trim() || "Khac";
    bySlog[slog] = (bySlog[slog] || 0) + 1;
    const syncStatusDirect = vtdApp_pick_(item.row, ["sync_status", "sync status", "syncstatus"]);
    const pj2StatusDirect = vtdApp_pick_(item.row, ["pj2_status", "pj2 status", "pj2status"]);
    const retrySyncDirect = vtdApp_pick_(item.row, ["retry_sync", "retry sync", "retrysync"]);
    const label = vtdApp_syncLabel_(syncStatusDirect, pj2StatusDirect, retrySyncDirect);
    byStatus[label] = (byStatus[label] || 0) + 1;
    if (vtdApp_isTargetNotSynced_(pj2StatusDirect)) syncErrorsByTarget.hangHoan = (syncErrorsByTarget.hangHoan || 0) + 1;
    if (vtdApp_isTargetNotSynced_(syncStatusDirect)) syncErrorsByTarget.chungTu = (syncErrorsByTarget.chungTu || 0) + 1;
    const user = String(compact.userEmail || compact.user || "Chua co user").trim() || "Chua co user";
    byUser[user] = (byUser[user] || 0) + 1;
    recent.push(compact);
  });
  vtdApp_enrichCompactsFromFullCtByOrders_(recent);
  vtdApp_attachSkuItemsToCompacts_(ss, recent.filter(compact => {
    const loai = vtdApp_norm_(compact && compact.loaiHoan);
    return loai === "san pham" || loai === "thu hoi";
  }));
  const statusRows = vtdApp_toKeyRows_(byStatus);
  const slogRows = vtdApp_toKeyRows_(bySlog);
  const userRows = vtdApp_toKeyRows_(byUser);
  return vtdApp_ok_({
    totalKnown: raw ? Math.max(0, raw.getLastRow() - 1) : 0,
    todayTotal: recent.length,
    queued: Number(byStatus["Dang dong bo"] || byStatus["Đang đồng bộ"] || 0),
    error: Number(syncErrorsByTarget.hangHoan || 0) + Number(syncErrorsByTarget.chungTu || 0),
    syncErrorsByTarget: [
      {key: "Hàng hoàn", value: Number(syncErrorsByTarget.hangHoan || 0), target: "FFHN | Hoan VITA"},
      {key: "Chứng từ", value: Number(syncErrorsByTarget.chungTu || 0), target: "FFHN | Chung tu VTD"}
    ],
    bySlog: slogRows,
    byUser: userRows,
    byStatus: statusRows,
    recent: recent.reverse()
  });
}

function vtdApp_search(params) {
  const denied = vtdApp_requireAction_("search", params);
  if (denied) return denied;
  params = params || {};
  const query = vtdApp_norm_(params.query || "");
  const filters = {
    slog: vtdApp_norm_(params.slog || ""),
    loaiHoan: vtdApp_norm_(params.loaiHoan || ""),
    fromDate: params.fromDate ? vtdApp_dateKey_(params.fromDate) : "",
    toDate: params.toDate ? vtdApp_dateKey_(params.toDate) : ""
  };
  const limit = Math.min(Number(params.limit) || VTD_APP.defaultLimit, 2000);
  const cached = VTD_APP.useRecordCache ? vtdApp_searchRecordCache_(query, filters, limit + 1) : null;
  if (cached !== null) {
    vtdApp_attachSkuItemsToCompacts_(vtdApp_ss_(), cached);
    return vtdApp_ok_({
      results: cached.slice(0, limit),
      count: Math.min(cached.length, limit),
      limit: limit,
      hasMore: cached.length > limit,
      source: "recordCache"
    });
  }
  const results = [];
  const seen = {};
  let hasMore = false;
  const addResult = compact => {
    const key = compact.id || compact.maDon || compact.source + ":" + compact.rowNumber;
    if (seen[key]) return;
    if (results.length >= limit) {
      hasMore = true;
      return;
    }
    seen[key] = true;
    results.push(compact);
  };

  const ss = vtdApp_ss_();
  const sheets = [ss.getSheetByName(VTD_APP.rawSheet), ss.getSheetByName(VTD_APP.rawFullSheet)];
  sheets.forEach(sheet => {
    const readLimit = Math.max(VTD_APP.maxSearchRows, limit);
    vtdApp_readSheetObjects_(sheet, readLimit).forEach(item => {
      const compact = vtdApp_compactRecord_(item);
      if (!vtdApp_recordPassesFilters_(item, compact, query, filters)) return;
      addResult(compact);
    });
  });
  vtdApp_enrichCompactsFromFullCtByOrders_(results);
  vtdApp_attachSkuItemsToCompacts_(ss, results);
  return vtdApp_ok_({results: results, count: results.length, limit: limit, hasMore: hasMore, source: "sheet"});
}

function vtdApp_saveRaw(params) {
  const denied = vtdApp_requireAction_("saveRaw", params);
  if (denied) return denied;
  params = params || {};
  const auth = vtdApp_auth_(params);
  const ss = vtdApp_ss_();
  const sheet = ss.getSheetByName(VTD_APP.rawSheet);
  if (!sheet) return vtdApp_fail_("Khong thay sheet " + VTD_APP.rawSheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill("");
  const col = vtdApp_headerMap_(headers);
  const now = new Date();
  const id = params.id || params.ID || vtdApp_makeId_();
  const maDon = vtdApp_cleanMaDonParam_(params);
  const slog = String(params.slog || params.SLOG || "").trim();
  const loaiHoan = String(params.loaiHoan || "").trim();
  const xacThuc = String(params.xacThucHoaDon || params.xacThucChungTu || params.xacThuc || "").trim();
  const orderInfo = params.orderInfo || params.lookup || {};
  const warnings = [];
  const ngayText = Utilities.formatDate(now, "Asia/Saigon", "dd/MM/yyyy");
  const gioText = Utilities.formatDate(now, "Asia/Saigon", "HH:mm:ss");

  if (!maDon) warnings.push("Thiếu mã đơn GHTK");
  if (!slog) warnings.push("Thiếu SLOG");
  if (!loaiHoan) warnings.push("Thiếu loại hoàn");
  if (loaiHoan === "Chứng từ" && !xacThuc) warnings.push("Thiếu xác thực chứng từ");

  vtdApp_setByAliases_(row, col, ["id"], id);
  vtdApp_setByAliases_(row, col, ["ma don", "ma don ghn", "ma don ghtk"], maDon);
  vtdApp_setByAliases_(row, col, ["slog"], slog);
  vtdApp_setByAliases_(row, col, ["xac thuc hoa don", "xac thuc chung tu"], xacThuc);
  vtdApp_setByAliases_(row, col, ["loai hoan"], loaiHoan);
  vtdApp_setByAliases_(row, col, ["ngay hoan tra"], ngayText);
  vtdApp_setByAliases_(row, col, ["thoi gian nhap hoan", "timestamp"], gioText);
  vtdApp_setByAliases_(row, col, ["ma po chung tu hoan", "ma po"], params.po);
  vtdApp_setByAliases_(row, col, ["ma qr", "qr"], params.maQr || params.qr);
  vtdApp_setByAliases_(row, col, ["tinh trang chung tu"], params.tinhTrang || "Binh thuong");
  vtdApp_setByAliases_(row, col, ["thong tin chung tu"], params.thongTin);
  vtdApp_setByAliases_(row, col, ["user_email", "nguoi nhap", "email"], auth.email || "");
  vtdApp_setByAliases_(row, col, ["khach hang", "khách hàng", "customer"], orderInfo.khachHang || orderInfo.customer);
  vtdApp_setByAliases_(row, col, ["od"], orderInfo.od);
  vtdApp_setByAliases_(row, col, ["so hoa don", "số hóa đơn", "invoice"], orderInfo.soHoaDon || orderInfo.invoice);
  vtdApp_setByAliases_(row, col, ["loai st", "loại st", "loai sieu thi", "loại siêu thị"], orderInfo.loaiSt || orderInfo.storeType);
  vtdApp_setByAliases_(row, col, ["thang", "tháng", "month"], orderInfo.month);

  sheet.appendRow(row);
  const rowNumber = sheet.getLastRow();
  const skuItems = Array.isArray(params.skuItems) ? params.skuItems.map(vtdApp_normalizeSkuItem_) : [];
  const masterSkuCount = vtdApp_syncNewMasterSku_(ss, skuItems);
  const skuCount = vtdApp_saveSkuRows_(ss, id, maDon, skuItems, auth);
  vtdApp_queueRow_(sheet, rowNumber);
  if (VTD_APP.useRecordCache) vtdApp_appendSearchCacheForRawRow_(sheet, rowNumber);
  vtdApp_log_(warnings.length ? "WARN" : "INFO", "saveRaw", auth.email, id, maDon, "OK", "saved row " + rowNumber, warnings.join("; "));

  return vtdApp_ok_({
    id: id,
    maDon: maDon,
    rowNumber: rowNumber,
    skuCount: skuCount,
    masterSkuCount: masterSkuCount,
    warnings: warnings,
    message: "Da luu chung tu" + (skuCount ? " va " + skuCount + " SKU." : ".")
  });
}

function vtdApp_storeInfo(params) {
  const denied = vtdApp_requireAction_("storeInfo", params);
  if (denied) return denied;
  params = params || {};
  const ss = SpreadsheetApp.openById(VTD_APP.fullCtSpreadsheetId);
  const byStore = {};
  ss.getSheets().forEach(sheet => {
    if (sheet.isSheetHidden()) return;
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return;
    const readRows = Math.min(1500, lastRow - 1);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const col = vtdApp_headerMap_(headers);
    const values = sheet.getRange(Math.max(2, lastRow - readRows + 1), 1, readRows, lastCol).getValues();
    values.forEach(row => {
      const customer = String(vtdApp_pickFromArray_(row, col, ["khach hang", "khách hàng", "customer", "ten khach hang"]) || "").trim();
      const store = vtdApp_storeType_(customer || name);
      byStore[store] = (byStore[store] || 0) + 1;
    });
  });
  vtdApp_enrichCompactsFromFullCtByOrders_(results);
  vtdApp_attachSkuItemsToCompacts_(ss, results);
  return vtdApp_ok_({
    source: "Full das chung tu VTD",
    byDate: vtdApp_toKeyRows_(byStore).sort((a, b) => b.value - a.value),
    rows: []
  });
}

function vtdApp_listProducts(params) {
  const denied = vtdApp_requireAction_("saveRaw", params);
  if (denied) return denied;
  params = params || {};
  const query = vtdApp_norm_(params.query || "");
  const limit = Math.min(Number(params.limit) || 5000, 5000);
  const ss = vtdApp_ss_();
  const sheetNames = ["MASTER SKU"];
  const out = [];
  const seen = {};
  sheetNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() <= 1 || out.length >= limit) return;
    const lastCol = sh.getLastColumn();
    const readRows = Math.min(5000, sh.getLastRow() - 1);
    const values = sh.getRange(1, 1, readRows + 1, lastCol).getValues();
    const col = vtdApp_headerMap_(values[0]);
    for (let i = 1; i < values.length; i++) {
      if (out.length >= limit) break;
      const row = values[i];
      const ma = String(vtdApp_pickFromArray_(row, col, ["ma vat tu", "mã vật tư", "ma sku", "sku", "id", "material", "ma sp", "mã sp"]) || "").trim();
      const ten = String(vtdApp_pickFromArray_(row, col, ["ten san pham", "tên sản phẩm", "ten sku", "san pham", "tên sản phẩm hoàn về", "ten san pham hoan ve", "product name", "name"]) || "").trim();
      const loai = String(vtdApp_pickFromArray_(row, col, ["loai", "loại", "type"]) || "").trim();
      if (!ten && !ma) continue;
      const key = vtdApp_norm_(ma + "|" + ten);
      if (seen[key]) continue;
      const text = vtdApp_norm_([ma, ten, loai].join(" "));
      if (query && text.indexOf(query) < 0) continue;
      seen[key] = true;
      const displayName = ten || ma;
      out.push({
        ma: ma,
        code: ma,
        material: ma,
        ten: displayName,
        name: displayName,
        productName: displayName,
        tenSanPham: displayName,
        "Tên sản phẩm hoàn về": displayName,
        loai: loai,
        type: loai,
        source: name
      });
    }
  });
  const count = query ? out.length : Object.keys(seen).length;
  if (!query && Number(params.knownCount || 0) >= count && count > 0) {
    return vtdApp_ok_({products: [], count: count, unchanged: true});
  }
  return vtdApp_ok_({products: out, count: count});
}

function vtdApp_lookupOrder(params) {
  const denied = vtdApp_requireAction_("saveRaw", params);
  if (denied) return denied;
  params = params || {};
  const mode = String(params.mode || "").trim();
  if (mode === "cache") {
    const months = Array.isArray(params.months) ? params.months : (params.month ? [params.month] : []);
    const limit = Math.min(Number(params.limit) || 20000, 40000);
    const orders = vtdApp_lookupOrdersByMonths_(months, limit);
    return vtdApp_ok_({
      source: "Full das chung tu VTD / TongHop",
      mode: "cache",
      months: months,
      count: orders.length,
      orders: orders
    });
  }

  const query = String(params.query || params.maDon || "").trim();
  if (!query) return vtdApp_fail_("Thiếu mã đơn cần tra.");
  const order = vtdApp_findFullCtOrderExact_(query);
  if (!order) return vtdApp_ok_({found: false, order: null, message: "Không tìm thấy mã đơn trong Full chứng từ."});
  return vtdApp_ok_({found: true, order: order});
}

function vtdApp_lookupOrderManifest(params) {
  const denied = vtdApp_requireAction_("saveRaw", params);
  if (denied) return denied;
  params = params || {};
  const days = Math.max(1, Math.min(Number(params.days) || 14, 60));
  const sh = vtdApp_fullCtTongHopSheet_();
  if (!sh || sh.getLastRow() <= 1) {
    return vtdApp_ok_({
      source: "Full das chung tu VTD / TongHop",
      lastRow: 1,
      refreshStartRow: 2,
      cutoffDateKey: "",
      months: [],
      monthTotals: [],
      pageSize: 2000,
      retentionDays: 400
    });
  }

  const lastRow = sh.getLastRow();
  const headers = sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), 16)).getDisplayValues()[0];
  const col = vtdApp_headerMap_(headers);
  const dateCol = vtdApp_firstCol_(col, ["ngay len don", "ngày lên đơn", "date"]);
  const monthCol = vtdApp_firstCol_(col, ["thang", "tháng", "month"]);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const cutoffDateKey = Utilities.formatDate(cutoff, "Asia/Saigon", "yyyy-MM-dd");
  let refreshStartRow = lastRow + 1;
  const months = {};

  const numRows = lastRow - 1;
  if (numRows > 0) {
    const dateValues = dateCol == null ? [] : sh.getRange(2, dateCol + 1, numRows, 1).getDisplayValues();
    const monthValues = monthCol == null ? [] : sh.getRange(2, monthCol + 1, numRows, 1).getDisplayValues();
    for (let i = 0; i < numRows; i++) {
      const dateRaw = dateValues[i] ? dateValues[i][0] : "";
      const monthRaw = monthValues[i] ? monthValues[i][0] : "";
      const month = vtdApp_monthDisplay_(monthRaw) || vtdApp_monthDisplay_(dateRaw);
      if (month) months[month] = (months[month] || 0) + 1;
      const key = vtdApp_dateKey_(dateRaw);
      if (key && key >= cutoffDateKey && refreshStartRow === lastRow + 1) refreshStartRow = i + 2;
    }
  }
  if (refreshStartRow > lastRow) refreshStartRow = Math.max(2, lastRow);

  return vtdApp_ok_({
    source: "Full das chung tu VTD / TongHop",
    lastRow: lastRow,
    refreshStartRow: refreshStartRow,
    cutoffDateKey: cutoffDateKey,
    months: Object.keys(months).sort(vtdApp_monthSort_),
    monthTotals: vtdApp_toKeyRows_(months).sort((a, b) => vtdApp_monthSort_(a.key, b.key)),
    pageSize: 2000,
    retentionDays: 400,
    generatedAt: Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss")
  });
}

function vtdApp_lookupOrderPage(params) {
  const denied = vtdApp_requireAction_("saveRaw", params);
  if (denied) return denied;
  params = params || {};
  const sh = vtdApp_fullCtTongHopSheet_();
  if (!sh || sh.getLastRow() <= 1) return vtdApp_ok_({orders: [], count: 0, done: true, nextRow: 2, lastRow: 1});

  const lastRow = sh.getLastRow();
  const pageSize = Math.max(1, Math.min(Number(params.pageSize) || 2000, 2000));
  const startRow = Math.max(2, Number(params.startRow || params.cursor || 2));
  const endRow = Math.min(lastRow, Number(params.endRow || lastRow));
  if (startRow > endRow) return vtdApp_ok_({orders: [], count: 0, done: true, nextRow: startRow, lastRow: lastRow});

  const take = Math.min(pageSize, endRow - startRow + 1);
  const lastCol = Math.min(sh.getLastColumn(), 16);
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const col = vtdApp_headerMap_(headers);
  const values = sh.getRange(startRow, 1, take, lastCol).getDisplayValues();
  const orders = [];
  const seen = {};
  for (let i = 0; i < values.length; i++) {
    const rec = vtdApp_orderLookupRecordFromRow_(values[i], col, sh.getName(), startRow + i);
    if (!rec.maDon) continue;
    const key = vtdApp_norm_([rec.maDon, rec.dateKey || rec.ngayLenDon].join(" "));
    if (seen[key]) continue;
    seen[key] = true;
    orders.push(rec);
  }
  const nextRow = startRow + take;
  return vtdApp_ok_({
    source: "Full das chung tu VTD / TongHop",
    startRow: startRow,
    nextRow: nextRow,
    endRow: endRow,
    lastRow: lastRow,
    done: nextRow > endRow,
    count: orders.length,
    orders: orders
  });
}

function vtdApp_cacheDataSet_(type) {
  const key = String(type || "").trim();
  if (key === "rawFull") {
    return {
      type: "rawFull",
      label: "RAW_IN_FULL",
      sheetId: VTD_APP.spreadsheetId,
      sheetName: VTD_APP.rawFullSheet,
      capRows: 20000
    };
  }
  if (key === "skuReturn") {
    return {
      type: "skuReturn",
      label: "List SKU hoàn",
      sheetId: VTD_APP.spreadsheetId,
      sheetName: VTD_APP.skuSheet,
      capRows: 5000
    };
  }
  return null;
}

function vtdApp_cacheDataManifest(params) {
  const denied = vtdApp_requireAction_("search", params);
  if (denied) return denied;
  params = params || {};
  const dataSet = vtdApp_cacheDataSet_(params.type);
  if (!dataSet) return vtdApp_fail_("Cache type không hợp lệ.");
  const ss = SpreadsheetApp.openById(dataSet.sheetId);
  const sh = ss.getSheetByName(dataSet.sheetName);
  const lastRow = sh ? sh.getLastRow() : 1;
  const totalRows = Math.max(0, lastRow - 1);
  const capRows = Number(dataSet.capRows || 0);
  const startRow = totalRows > capRows ? Math.max(2, lastRow - capRows + 1) : 2;
  return vtdApp_ok_({
    type: dataSet.type,
    label: dataSet.label,
    sheetName: dataSet.sheetName,
    lastRow: lastRow,
    startRow: startRow,
    endRow: lastRow,
    total: Math.max(0, lastRow - startRow + 1),
    capRows: capRows,
    pageSize: 2000,
    generatedAt: Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss")
  });
}

function vtdApp_cacheDataPage(params) {
  const denied = vtdApp_requireAction_("search", params);
  if (denied) return denied;
  params = params || {};
  const dataSet = vtdApp_cacheDataSet_(params.type);
  if (!dataSet) return vtdApp_fail_("Cache type không hợp lệ.");
  const ss = SpreadsheetApp.openById(dataSet.sheetId);
  const sh = ss.getSheetByName(dataSet.sheetName);
  if (!sh || sh.getLastRow() <= 1) {
    return vtdApp_ok_({type: dataSet.type, label: dataSet.label, records: [], count: 0, done: true, nextRow: 2, lastRow: 1});
  }
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const pageSize = Math.max(1, Math.min(Number(params.pageSize) || 2000, 2000));
  const startRow = Math.max(2, Number(params.startRow || params.cursor || 2));
  const endRow = Math.min(lastRow, Number(params.endRow || lastRow));
  if (startRow > endRow) {
    return vtdApp_ok_({type: dataSet.type, label: dataSet.label, records: [], count: 0, done: true, nextRow: startRow, endRow: endRow, lastRow: lastRow});
  }
  const take = Math.min(pageSize, endRow - startRow + 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const col = vtdApp_headerMap_(headers);
  const displayValues = sh.getRange(startRow, 1, take, lastCol).getDisplayValues();
  const records = [];
  displayValues.forEach((row, offset) => {
    if (row.join("") === "") return;
    const rowNumber = startRow + offset;
    if (dataSet.type === "rawFull") {
      records.push(vtdApp_rawFullCacheRecord_(row, col, headers, dataSet.sheetName, rowNumber));
    } else if (dataSet.type === "skuReturn") {
      records.push(vtdApp_skuReturnCacheRecord_(row, col, headers, dataSet.sheetName, rowNumber));
    }
  });
  const nextRow = startRow + take;
  return vtdApp_ok_({
    type: dataSet.type,
    label: dataSet.label,
    sheetName: dataSet.sheetName,
    startRow: startRow,
    nextRow: nextRow,
    endRow: endRow,
    lastRow: lastRow,
    done: nextRow > endRow,
    count: records.length,
    records: records
  });
}

function vtdApp_rawFullCacheRecord_(row, col, headers, sourceSheet, rowNumber) {
  const rowObj = {};
  const displayRow = {};
  headers.forEach((header, i) => {
    const key = vtdApp_norm_(header);
    rowObj[key] = row[i];
    displayRow[key] = row[i];
  });
  const item = {
    source: sourceSheet,
    rowNumber: rowNumber,
    headers: headers,
    values: row,
    displayValues: row,
    row: rowObj,
    displayRow: displayRow
  };
  const c = vtdApp_compactRecord_(item);
  const ngay = c.ngayHoanTra || vtdApp_pickFromArray_(row, col, ["ngay hoan tra", "ngày hoàn trả", "ngay nhap"]);
  const gio = c.thoiGianNhap || vtdApp_pickFromArray_(row, col, ["thoi gian nhap hoan", "thời gian nhập hoàn", "thoi gian"]);
  const searchText = [
    c.id, c.maDon, c.slog, c.loaiHoan, c.xacThucHoaDon, c.po, c.maQr, c.soHoaDon,
    c.od, c.khachHang, c.userEmail, c.tinhTrang, c.thongTin, ngay, gio
  ].join(" ");
  return {
    sourceSheet: sourceSheet,
    sourceRow: rowNumber,
    rawId: c.id,
    id: c.id,
    maDon: c.maDon,
    slog: c.slog,
    loaiHoan: c.loaiHoan,
    xacThucHoaDon: c.xacThucHoaDon,
    po: c.po,
    maQr: c.maQr,
    soHoaDon: c.soHoaDon,
    od: c.od,
    khachHang: c.khachHang,
    userEmail: c.userEmail,
    ngayHoanTra: ngay,
    thoiGianNhap: gio,
    dateKey: vtdApp_dateKey_(ngay),
    tinhTrang: c.tinhTrang,
    thongTin: c.thongTin,
    syncStatus: c.syncStatus,
    pj2Status: c.pj2Status,
    retrySync: c.retrySync,
    syncLabel: c.syncLabel,
    searchText: searchText
  };
}

function vtdApp_skuReturnCacheRecord_(row, col, headers, sourceSheet, rowNumber) {
  const rawId = String(vtdApp_pickFromArray_(row, col, ["raw_id", "raw id"]) || "").trim();
  const maDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || "").trim();
  const ten = String(vtdApp_pickFromArray_(row, col, ["ten san pham hoan ve", "tên sản phẩm hoàn về", "ten san pham", "tên sản phẩm"]) || "").trim();
  const type = String(vtdApp_pickFromArray_(row, col, ["type", "loai", "loại"]) || "").trim();
  const soLuong = String(vtdApp_pickFromArray_(row, col, ["so luong", "số lượng", "sl"]) || "").trim();
  const tinhTrang = String(vtdApp_pickFromArray_(row, col, ["tinh trang", "tình trạng"]) || "").trim();
  const lot = String(vtdApp_pickFromArray_(row, col, ["ma lo/hsd", "mã lô/hsd", "ma lo hsd", "mã lô hsd"]) || "").trim();
  const serial = String(vtdApp_pickFromArray_(row, col, ["ma serial", "mã serial", "serial"]) || "").trim();
  const image = String(vtdApp_pickFromArray_(row, col, ["anh san pham", "Ảnh sản phẩm", "ảnh sản phẩm", "anh", "ảnh"]) || "").trim();
  const note = String(vtdApp_pickFromArray_(row, col, ["ghi chu", "ghi chú", "note"]) || "").trim();
  const userEmail = String(vtdApp_pickFromArray_(row, col, ["user_email", "user email", "email"]) || "").trim();
  const ngay = String(vtdApp_pickFromArray_(row, col, ["ngay", "ngày", "ngay hoan tra", "ngày hoàn trả"]) || "").trim();
  const gio = String(vtdApp_pickFromArray_(row, col, ["gio", "giờ", "thoi gian", "thời gian"]) || "").trim();
  return {
    sourceSheet: sourceSheet,
    sourceRow: rowNumber,
    rawId: rawId,
    id: String(vtdApp_pickFromArray_(row, col, ["id"]) || "").trim(),
    maDon: maDon,
    ten: ten,
    name: ten,
    type: type,
    soLuong: soLuong,
    qty: soLuong,
    tinhTrang: tinhTrang,
    maLoHsd: lot,
    lot: lot,
    serial: serial,
    image: image,
    images: image ? [image] : [],
    ghiChu: note,
    note: note,
    userEmail: userEmail,
    ngay: ngay,
    gio: gio,
    dateKey: vtdApp_dateKey_(ngay),
    searchText: [rawId, maDon, ten, type, soLuong, tinhTrang, lot, serial, note, userEmail, image].join(" ")
  };
}

function vtdApp_lookupCacheStatus(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  const sh = vtdApp_lookupCacheStatusSheet_();
  const values = sh.getDataRange().getDisplayValues();
  if (values.length <= 1) return vtdApp_ok_({rows: []});
  const col = vtdApp_headerMap_(values[0]);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    rows.push({
      email: vtdApp_pickFromArray_(row, col, ["email"]),
      deviceId: vtdApp_pickFromArray_(row, col, ["deviceid", "device id"]),
      cacheType: vtdApp_pickFromArray_(row, col, ["cachetype", "cache type"]) || "customerLookup",
      label: vtdApp_pickFromArray_(row, col, ["label", "cache label"]) || "",
      status: vtdApp_pickFromArray_(row, col, ["status"]),
      downloaded: vtdApp_pickFromArray_(row, col, ["downloaded"]),
      total: vtdApp_pickFromArray_(row, col, ["total"]),
      startRow: vtdApp_pickFromArray_(row, col, ["startrow"]),
      endRow: vtdApp_pickFromArray_(row, col, ["endrow"]),
      nextRow: vtdApp_pickFromArray_(row, col, ["nextrow"]),
      lastError: vtdApp_pickFromArray_(row, col, ["lasterror"]),
      updatedAt: vtdApp_pickFromArray_(row, col, ["updatedat"]),
      appVersion: vtdApp_pickFromArray_(row, col, ["appversion"])
    });
  }
  rows.sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
  return vtdApp_ok_({rows: rows.slice(0, 100)});
}

function vtdApp_deleteLookupCacheStatus(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  const deviceId = String(params.deviceId || "").trim();
  const cacheType = String(params.cacheType || "").trim();
  if (!email && !deviceId) return vtdApp_fail_("Thiếu thông tin báo cáo cần xóa.");
  const sh = vtdApp_lookupCacheStatusSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return vtdApp_ok_({message: "Không có báo cáo cache."});
  const col = vtdApp_headerMap_(values[0]);
  for (let i = values.length - 1; i >= 1; i--) {
    const rowEmail = String(values[i][col.email] || "").toLowerCase().trim();
    const rowDevice = String(values[i][col.deviceid] || "").trim();
    const rowCacheType = String(values[i][col.cachetype] || "customerLookup").trim();
    if ((!email || rowEmail === email) && (!deviceId || rowDevice === deviceId) && (!cacheType || rowCacheType === cacheType)) {
      sh.deleteRow(i + 1);
    }
  }
  return vtdApp_ok_({message: "Đã xóa báo cáo cache.", email: email, deviceId: deviceId});
}

function vtdApp_queueErrors(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  const sh = vtdApp_queueErrorSheet_();
  const values = sh.getDataRange().getDisplayValues();
  if (values.length <= 1) return vtdApp_ok_({rows: []});
  const col = vtdApp_headerMap_(values[0]);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    rows.push({
      updatedAt: vtdApp_pickFromArray_(row, col, ["updatedat"]),
      email: vtdApp_pickFromArray_(row, col, ["email"]),
      deviceId: vtdApp_pickFromArray_(row, col, ["deviceid"]),
      localId: vtdApp_pickFromArray_(row, col, ["localid"]),
      maDon: vtdApp_pickFromArray_(row, col, ["madon"]),
      slog: vtdApp_pickFromArray_(row, col, ["slog"]),
      loaiHoan: vtdApp_pickFromArray_(row, col, ["loaihoan"]),
      status: vtdApp_pickFromArray_(row, col, ["status"]),
      step: vtdApp_pickFromArray_(row, col, ["step"]),
      attempts: vtdApp_pickFromArray_(row, col, ["attempts"]),
      imageCount: vtdApp_pickFromArray_(row, col, ["imagecount"]),
      imageIndex: vtdApp_pickFromArray_(row, col, ["imageindex"]),
      lastError: vtdApp_pickFromArray_(row, col, ["lasterror"]),
      createdAt: vtdApp_pickFromArray_(row, col, ["createdat"]),
      appVersion: vtdApp_pickFromArray_(row, col, ["appversion"])
    });
  }
  rows.sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
  return vtdApp_ok_({rows: rows.slice(0, 300)});
}

function vtdApp_reportQueueErrors(params) {
  params = params || {};
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const email = String((auth && auth.email) || params.email || "").toLowerCase().trim();
  const deviceId = String(params.deviceId || "unknown").trim();
  const appVersion = String(params.appVersion || "").trim();
  const entries = Array.isArray(params.entries) ? params.entries : [];
  const sh = vtdApp_queueErrorSheet_();
  const values = sh.getDataRange().getValues();
  const col = vtdApp_headerMap_(values[0]);

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][col.email] || "").toLowerCase().trim() === email &&
        String(values[i][col.deviceid] || "").trim() === deviceId) {
      sh.deleteRow(i + 1);
    }
  }
  if (!entries.length) return vtdApp_ok_({message: "Không có lỗi queue.", count: 0});

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = vtdApp_headerMap_(headers);
  const now = Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss");
  const rows = entries.map(entry => {
    entry = entry || {};
    const row = new Array(headers.length).fill("");
    vtdApp_setByAliases_(row, map, ["UpdatedAt"], now);
    vtdApp_setByAliases_(row, map, ["Email"], email);
    vtdApp_setByAliases_(row, map, ["DeviceId"], deviceId);
    vtdApp_setByAliases_(row, map, ["LocalId"], entry.localId || "");
    vtdApp_setByAliases_(row, map, ["MaDon"], entry.maDon || "");
    vtdApp_setByAliases_(row, map, ["Slog"], entry.slog || "");
    vtdApp_setByAliases_(row, map, ["LoaiHoan"], entry.loaiHoan || "");
    vtdApp_setByAliases_(row, map, ["Status"], entry.status || "");
    vtdApp_setByAliases_(row, map, ["Step"], entry.step || "");
    vtdApp_setByAliases_(row, map, ["Attempts"], entry.attempts || 0);
    vtdApp_setByAliases_(row, map, ["ImageCount"], entry.imageCount || 0);
    vtdApp_setByAliases_(row, map, ["ImageIndex"], entry.imageIndex || 0);
    vtdApp_setByAliases_(row, map, ["LastError"], entry.lastError || "");
    vtdApp_setByAliases_(row, map, ["CreatedAt"], entry.createdAt || "");
    vtdApp_setByAliases_(row, map, ["AppVersion"], appVersion || entry.appVersion || "");
    return row;
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  return vtdApp_ok_({message: "Đã cập nhật lỗi queue.", count: rows.length});
}

function vtdApp_deleteQueueError(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const email = String(params.email || "").toLowerCase().trim();
  const deviceId = String(params.deviceId || "").trim();
  const localId = String(params.localId || "").trim();
  const sh = vtdApp_queueErrorSheet_();
  const values = sh.getDataRange().getValues();
  const col = vtdApp_headerMap_(values[0]);
  for (let i = values.length - 1; i >= 1; i--) {
    const okEmail = !email || String(values[i][col.email] || "").toLowerCase().trim() === email;
    const okDevice = !deviceId || String(values[i][col.deviceid] || "").trim() === deviceId;
    const okLocal = !localId || String(values[i][col.localid] || "").trim() === localId;
    if (okEmail && okDevice && okLocal) sh.deleteRow(i + 1);
  }
  return vtdApp_ok_({message: "Đã xóa lỗi queue."});
}

function vtdApp_requestQueueRetry(params) {
  const denied = vtdApp_requireAction_("retrySync", params);
  if (denied) return denied;
  params = params || {};
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);

  const email = String(params.email || params.targetEmail || "").toLowerCase().trim();
  const deviceId = String(params.deviceId || "").trim();
  const localId = String(params.localId || "").trim();
  const maDon = String(params.maDon || "").trim();
  const step = String(params.step || "").trim();
  if (!email && !deviceId && !localId && !maDon) {
    return vtdApp_fail_("Thiếu thông tin đơn/máy cần retry.");
  }

  const sh = vtdApp_queueRetrySheet_();
  const values = sh.getDataRange().getValues();
  const col = vtdApp_headerMap_(values[0]);
  const now = Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sameOpen = String(vtdApp_pickFromArray_(row, col, ["status"]) || "").toUpperCase() === "PENDING";
    const sameEmail = String(vtdApp_pickFromArray_(row, col, ["targetemail", "email"]) || "").toLowerCase().trim() === email;
    const sameDevice = String(vtdApp_pickFromArray_(row, col, ["deviceid"]) || "").trim() === deviceId;
    const sameLocal = String(vtdApp_pickFromArray_(row, col, ["localid"]) || "").trim() === localId;
    const sameMaDon = String(vtdApp_pickFromArray_(row, col, ["madon"]) || "").trim() === maDon;
    if (sameOpen && sameEmail && sameDevice && (localId ? sameLocal : sameMaDon)) {
      return vtdApp_ok_({message: "Đơn này đã có lệnh retry đang chờ máy user nhận.", id: vtdApp_pickFromArray_(row, col, ["id"]) || ""});
    }
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = vtdApp_headerMap_(headers);
  const id = "QR-" + Utilities.formatDate(new Date(), "Asia/Saigon", "yyyyMMdd-HHmmss") + "-" + Utilities.getUuid().slice(0, 8);
  const out = new Array(headers.length).fill("");
  vtdApp_setByAliases_(out, map, ["ID"], id);
  vtdApp_setByAliases_(out, map, ["CreatedAt"], now);
  vtdApp_setByAliases_(out, map, ["UpdatedAt"], now);
  vtdApp_setByAliases_(out, map, ["Status"], "PENDING");
  vtdApp_setByAliases_(out, map, ["TargetEmail"], email);
  vtdApp_setByAliases_(out, map, ["DeviceId"], deviceId);
  vtdApp_setByAliases_(out, map, ["LocalId"], localId);
  vtdApp_setByAliases_(out, map, ["MaDon"], maDon);
  vtdApp_setByAliases_(out, map, ["Step"], step);
  vtdApp_setByAliases_(out, map, ["RequestedBy"], auth.email || "");
  sh.appendRow(out);
  return vtdApp_ok_({message: "Đã gửi lệnh retry về máy user.", id: id});
}

function vtdApp_queueRetryRequests(params) {
  params = params || {};
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const email = String((auth && auth.email) || params.email || "").toLowerCase().trim();
  const deviceId = String(params.deviceId || "").trim();
  const sh = vtdApp_queueRetrySheet_();
  const values = sh.getDataRange().getDisplayValues();
  if (values.length <= 1) return vtdApp_ok_({rows: []});
  const col = vtdApp_headerMap_(values[0]);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = String(vtdApp_pickFromArray_(row, col, ["status"]) || "").toUpperCase().trim();
    const targetEmail = String(vtdApp_pickFromArray_(row, col, ["targetemail", "email"]) || "").toLowerCase().trim();
    const targetDevice = String(vtdApp_pickFromArray_(row, col, ["deviceid"]) || "").trim();
    if (status !== "PENDING") continue;
    if (targetEmail && targetEmail !== email) continue;
    if (targetDevice && deviceId && targetDevice !== deviceId) continue;
    if (targetDevice && !deviceId) continue;
    rows.push({
      id: vtdApp_pickFromArray_(row, col, ["id"]),
      email: targetEmail,
      deviceId: targetDevice,
      localId: vtdApp_pickFromArray_(row, col, ["localid"]),
      maDon: vtdApp_pickFromArray_(row, col, ["madon"]),
      step: vtdApp_pickFromArray_(row, col, ["step"]),
      requestedBy: vtdApp_pickFromArray_(row, col, ["requestedby"]),
      createdAt: vtdApp_pickFromArray_(row, col, ["createdat"])
    });
  }
  return vtdApp_ok_({rows: rows.slice(0, 50)});
}

function vtdApp_ackQueueRetryRequests(params) {
  params = params || {};
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const ids = Array.isArray(params.ids) ? params.ids.map(x => String(x || "").trim()).filter(Boolean) : [];
  if (!ids.length) return vtdApp_ok_({count: 0});

  const sh = vtdApp_queueRetrySheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return vtdApp_ok_({count: 0});
  const col = vtdApp_headerMap_(values[0]);
  const now = Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss");
  let count = 0;
  for (let i = 1; i < values.length; i++) {
    const id = String(vtdApp_pickFromArray_(values[i], col, ["id"]) || "").trim();
    if (ids.indexOf(id) < 0) continue;
    if (col.status != null) sh.getRange(i + 1, col.status + 1).setValue("ACK");
    if (col.updatedat != null) sh.getRange(i + 1, col.updatedat + 1).setValue(now);
    if (col.ackby != null) sh.getRange(i + 1, col.ackby + 1).setValue(auth.email || "");
    count++;
  }
  return vtdApp_ok_({count: count});
}

function vtdApp_queueErrorSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.queueErrorSheet);
  const headers = ["UpdatedAt", "Email", "DeviceId", "LocalId", "MaDon", "Slog", "LoaiHoan", "Status", "Step", "Attempts", "ImageCount", "ImageIndex", "LastError", "CreatedAt", "AppVersion"];
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.queueErrorSheet);
    sh.appendRow(headers);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow(headers);
  } else {
    const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const existingMap = vtdApp_headerMap_(existing);
    headers.forEach(header => {
      if (existingMap[vtdApp_norm_(header)] == null) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
      }
    });
  }
  return sh;
}

function vtdApp_queueRetrySheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.queueRetrySheet);
  const headers = ["ID", "CreatedAt", "UpdatedAt", "Status", "TargetEmail", "DeviceId", "LocalId", "MaDon", "Step", "RequestedBy", "AckBy"];
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.queueRetrySheet);
    sh.appendRow(headers);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow(headers);
  } else {
    const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    const existingMap = vtdApp_headerMap_(existing);
    headers.forEach(header => {
      if (existingMap[vtdApp_norm_(header)] == null) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
      }
    });
  }
  return sh;
}

function vtdApp_reportLookupCacheStatus(params) {
  params = params || {};
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const email = String((auth && auth.email) || params.email || "").toLowerCase().trim();
  const deviceId = String(params.deviceId || "unknown").trim();
  const cacheType = String(params.cacheType || "customerLookup").trim();
  const label = String(params.label || params.cacheLabel || "").trim();
  const sh = vtdApp_lookupCacheStatusSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const values = sh.getDataRange().getValues();
  let targetRow = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col.email] || "").toLowerCase().trim() === email &&
        String(values[i][col.deviceid] || "").trim() === deviceId &&
        String(values[i][col.cachetype] || "customerLookup").trim() === cacheType) {
      targetRow = i + 1;
      break;
    }
  }
  const row = new Array(headers.length).fill("");
  vtdApp_setByAliases_(row, col, ["Email"], email);
  vtdApp_setByAliases_(row, col, ["DeviceId"], deviceId);
  vtdApp_setByAliases_(row, col, ["CacheType"], cacheType);
  vtdApp_setByAliases_(row, col, ["Label"], label);
  vtdApp_setByAliases_(row, col, ["Status"], params.status || "");
  vtdApp_setByAliases_(row, col, ["Downloaded"], params.downloaded || 0);
  vtdApp_setByAliases_(row, col, ["Total"], params.total || 0);
  vtdApp_setByAliases_(row, col, ["StartRow"], params.startRow || "");
  vtdApp_setByAliases_(row, col, ["EndRow"], params.endRow || "");
  vtdApp_setByAliases_(row, col, ["NextRow"], params.nextRow || "");
  vtdApp_setByAliases_(row, col, ["LastError"], params.lastError || "");
  vtdApp_setByAliases_(row, col, ["UpdatedAt"], Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss"));
  vtdApp_setByAliases_(row, col, ["AppVersion"], params.appVersion || "");
  if (targetRow) sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return vtdApp_ok_({message: "Đã cập nhật trạng thái cache.", email: email, deviceId: deviceId});
}

function vtdApp_lookupCacheStatusSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.lookupCacheStatusSheet);
  const headers = ["Email", "DeviceId", "CacheType", "Label", "Status", "Downloaded", "Total", "StartRow", "EndRow", "NextRow", "LastError", "UpdatedAt", "AppVersion"];
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.lookupCacheStatusSheet);
    sh.appendRow(headers);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow(headers);
  } else {
    const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const existingMap = vtdApp_headerMap_(existing);
    headers.forEach(header => {
      if (existingMap[vtdApp_norm_(header)] == null) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
      }
    });
  }
  return sh;
}

function vtdApp_fullCtTongHopSheet_() {
  const ss = SpreadsheetApp.openById(VTD_APP.fullCtSpreadsheetId);
  return ss.getSheetByName("TongHop");
}

function vtdApp_findFullCtOrderExact_(query) {
  const sh = vtdApp_fullCtTongHopSheet_();
  if (!sh || sh.getLastRow() <= 1) return null;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const col = vtdApp_headerMap_(headers);
  const maDonCol = vtdApp_firstCol_(col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]);
  if (maDonCol == null) return null;

  const lastRow = sh.getLastRow();
  const finder = sh.getRange(2, maDonCol + 1, lastRow - 1, 1)
    .createTextFinder(String(query).trim())
    .matchEntireCell(true)
    .matchCase(false);
  const cell = finder.findNext();
  if (cell) {
    const rowNumber = cell.getRow();
    const row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    return vtdApp_orderLookupRecordFromRow_(row, col, sh.getName(), rowNumber);
  }

  const normQuery = vtdApp_norm_(query);
  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    const maDon = String(vtdApp_pickFromArray_(values[i], col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || "").trim();
    if (vtdApp_norm_(maDon) === normQuery) {
      return vtdApp_orderLookupRecordFromRow_(values[i], col, sh.getName(), i + 2);
    }
  }
  return null;
}

function vtdApp_lookupOrdersByMonths_(months, limit) {
  const monthSet = {};
  (months || []).forEach(month => {
    const normalized = vtdApp_monthDisplay_(month);
    if (normalized) monthSet[normalized] = true;
  });
  const sh = vtdApp_fullCtTongHopSheet_();
  if (!sh || sh.getLastRow() <= 1) return [];
  const lastCol = Math.min(sh.getLastColumn(), 16);
  const values = sh.getRange(1, 1, sh.getLastRow(), lastCol).getDisplayValues();
  const col = vtdApp_headerMap_(values[0]);
  const out = [];
  const seen = {};
  for (let i = 1; i < values.length && out.length < limit; i++) {
    const rec = vtdApp_orderLookupRecordFromRow_(values[i], col, sh.getName(), i + 1);
    if (!rec.maDon) continue;
    if (Object.keys(monthSet).length && !monthSet[rec.month]) continue;
    const key = vtdApp_norm_(rec.maDon);
    if (seen[key]) continue;
    seen[key] = true;
    out.push(rec);
  }
  return out;
}

function vtdApp_orderLookupRecordFromRow_(row, col, sourceSheet, rowNumber) {
  const ngayLenDon = String(vtdApp_pickFromArray_(row, col, ["ngay len don", "ngày lên đơn", "date"]) || "").trim();
  const maDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || "").trim();
  const customer = String(vtdApp_pickFromArray_(row, col, ["khach hang", "khách hàng", "customer", "ten khach hang", "tên khách hàng"]) || "").trim();
  const od = String(vtdApp_pickFromArray_(row, col, ["od"]) || "").trim();
  const po = String(vtdApp_pickFromArray_(row, col, ["ma po", "mã po", "po"]) || "").trim();
  const invoice = String(vtdApp_pickFromArray_(row, col, ["so hoa don", "số hóa đơn", "invoice"]) || "").trim();
  const storeType = String(vtdApp_pickFromArray_(row, col, ["loai sieu thi", "loại siêu thị", "loai st", "loại st"]) || "").trim();
  const month = vtdApp_monthDisplay_(vtdApp_pickFromArray_(row, col, ["thang", "tháng", "month"])) ||
    vtdApp_monthDisplay_(vtdApp_pickFromArray_(row, col, ["ngay len don", "ngày lên đơn", "date"]));
  return {
    maDon: maDon,
    ngayLenDon: ngayLenDon,
    dateKey: vtdApp_dateKey_(ngayLenDon),
    customer: customer,
    khachHang: customer,
    od: od,
    po: po,
    invoice: invoice,
    soHoaDon: invoice,
    storeType: storeType || vtdApp_storeType_(customer),
    loaiSt: storeType || vtdApp_storeType_(customer),
    month: month,
    sourceSheet: sourceSheet || "",
    rowNumber: rowNumber || ""
  };
}

function vtdApp_rebuildSearchIndex(params) {
  const denied = params && params.internalRun ? null : vtdApp_requireAction_("runCommand", params);
  if (denied) return denied;
  params = params || {};
  const limit = Math.min(Number(params.limit) || 20000, 50000);
  const clear = params.clear !== false;
  const ss = vtdApp_ss_();
  const index = vtdApp_ensureIndexSheet_(ss, VTD_APP.indexPrefix + "ALL");
  if (clear) {
    index.clearContents();
    index.appendRow(["RAW_ID", "SourceSheet", "SourceRow", "CreatedDate", "CreatedTime", "MaDon", "Slog", "LoaiHoan", "PO", "QR", "SoHoaDon", "OD", "KhachHang", "UserEmail", "SyncLabel", "PJ2Label", "ImageFolderUrl", "SearchText"]);
  }
  const sheets = [ss.getSheetByName(VTD_APP.rawSheet), ss.getSheetByName(VTD_APP.rawFullSheet)];
  const output = [];
  let count = 0;
  sheets.forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return;
    const start = Math.max(2, lastRow - limit + 1);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const values = sheet.getRange(start, 1, lastRow - start + 1, lastCol).getValues();
    values.forEach((rowValues, offset) => {
      if (rowValues.join("") === "") return;
      const rowObj = {};
      headers.forEach((header, i) => rowObj[vtdApp_norm_(header)] = rowValues[i]);
      const rowNumber = start + offset;
      const item = {source: sheet.getName(), rowNumber: rowNumber, headers: headers, values: rowValues, row: rowObj};
      const c = vtdApp_compactRecord_(item);
      const date = vtdApp_toDisplay_(vtdApp_pick_(rowObj, ["ngay hoan tra", "ngày hoàn trả", "ngay nhap"])) || vtdApp_dateKey_(new Date());
      const time = vtdApp_toDisplay_(vtdApp_pick_(rowObj, ["thoi gian nhap hoan", "thời gian nhập hoàn", "thoi gian", "timestamp"]));
      const user = vtdApp_pick_(rowObj, ["user_email", "user email", "email", "nguoi nhap", "người nhập", "nguoi thao tac", "người thao tác", "user thao tac", "user thao tác", "created by", "createdby"]);
      const soHd = vtdApp_pick_(rowObj, ["so hoa don", "số hóa đơn", "invoice"]);
      const od = vtdApp_pick_(rowObj, ["od"]);
      const kh = vtdApp_pick_(rowObj, ["khach hang", "khách hàng", "customer"]);
      const searchText = [c.id, c.maDon, c.slog, c.loaiHoan, c.po, c.maQr, soHd, od, kh, user, c.tinhTrang, c.thongTin].join(" ");
      output.push([c.id, sheet.getName(), rowNumber, date, time, c.maDon, c.slog, c.loaiHoan, c.po, c.maQr, soHd, od, kh, user, c.syncLabel, c.pj2Status, "", searchText]);
      count++;
    });
    while (output.length >= 1000) {
      const chunk = output.splice(0, 1000);
      index.getRange(index.getLastRow() + 1, 1, chunk.length, chunk[0].length).setValues(chunk);
    }
  });
  if (output.length) {
    index.getRange(index.getLastRow() + 1, 1, output.length, output[0].length).setValues(output);
  }
  return vtdApp_ok_({message: "Da rebuild index " + count + " dong.", count: count, indexSheet: index.getName()});
}

function vtdApp_rebuildRowMap(params) {
  const denied = params && params.internalRun ? null : vtdApp_requireAction_("runCommand", params);
  if (denied) return denied;
  params = params || {};
  const limit = Math.min(Number(params.limit) || 20000, 50000);
  const ss = vtdApp_ss_();
  vtdApp_deleteOldSearchIndexSheets_(ss);
  const map = vtdApp_ensureRowMapSheet_(ss);
  map.clearContents();
  map.appendRow(["MapKey", "KeyType", "SourceSheet", "SourceRow", "DateKey", "Slog", "LoaiHoan", "MaDon", "UserEmail"]);
  const output = [];
  let sourceRows = 0;
  [ss.getSheetByName(VTD_APP.rawFullSheet), ss.getSheetByName(VTD_APP.rawSheet)].forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return;
    const start = Math.max(2, lastRow - limit + 1);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const values = sheet.getRange(start, 1, lastRow - start + 1, lastCol).getValues();
    values.forEach((rowValues, offset) => {
      if (rowValues.join("") === "") return;
      const rowObj = {};
      headers.forEach((header, i) => rowObj[vtdApp_norm_(header)] = rowValues[i]);
      const rowNumber = start + offset;
      const item = {source: sheet.getName(), rowNumber: rowNumber, headers: headers, values: rowValues, row: rowObj};
      const c = vtdApp_compactRecord_(item);
      const dateKey = vtdApp_dateKey_(vtdApp_pick_(rowObj, ["ngay hoan tra", "ngày hoàn trả", "ngay nhap", "timestamp", "thoi gian nhap hoan", "thời gian nhập hoàn"]));
      const base = [sheet.getName(), rowNumber, dateKey, c.slog || "", c.loaiHoan || "", c.maDon || "", c.userEmail || ""];
      const add = (type, value) => {
        const keyValue = type === "DATE" ? String(value || "") : vtdApp_norm_(value);
        if (!keyValue) return;
        output.push([type + "|" + keyValue, type].concat(base));
      };
      add("DATE", dateKey);
      add("MADON", c.maDon);
      add("PO", c.po);
      add("SOHD", c.soHoaDon);
      add("OD", c.od);
      add("QR", c.maQr);
      add("SERIAL", c.serial);
      add("USER", c.userEmail);
      add("SLOG", c.slog);
      add("LOAI", c.loaiHoan);
      sourceRows++;
      while (output.length >= 1000) {
        const chunk = output.splice(0, 1000);
        map.getRange(map.getLastRow() + 1, 1, chunk.length, chunk[0].length).setValues(chunk);
      }
    });
  });
  if (output.length) map.getRange(map.getLastRow() + 1, 1, output.length, output[0].length).setValues(output);
  return vtdApp_ok_({message: "Da rebuild row map " + sourceRows + " dong nguon.", sourceRows: sourceRows, mapSheet: map.getName(), mapRows: Math.max(0, map.getLastRow() - 1)});
}

function vtdApp_runCommand(params) {
  const denied = vtdApp_requireAction_("runCommand", params);
  if (denied) return denied;
  params = params || {};
  const action = String(params.action || "").trim();
  if (action === "rebuildIndex" || action === "rebuildRowMap" || action === "rebuildSearchCache") return vtdApp_rebuildSearchCache(params);
  if (action === "processPJ1" && typeof processRetryQueue === "function") {
    return vtdApp_ok_({result: processRetryQueue(Number(params.limit) || 15)});
  }
  if (action === "processPJ2" && typeof processPJ2Queue === "function") {
    return vtdApp_ok_({result: processPJ2Queue(Number(params.limit) || 15)});
  }
  if (action === "fullSync" && typeof fullSyncInputToHistory === "function") {
    return vtdApp_ok_({result: fullSyncInputToHistory()});
  }
  if (action === "dispatchBackfill") {
    if (typeof dispatchBackfillPendingRows === "function") {
      return vtdApp_ok_({result: dispatchBackfillPendingRows(Number(params.startRow) || 2, Number(params.endRow) || Number(params.startRow) || 2)});
    }
    const ss = vtdApp_ss_();
    const sheet = ss.getSheetByName(VTD_APP.rawSheet);
    const start = Math.max(2, Number(params.startRow) || 2);
    const end = Math.min(sheet.getLastRow(), Number(params.endRow) || start);
    for (let r = start; r <= end; r++) vtdApp_queueRow_(sheet, r);
    return vtdApp_ok_({message: "Da backfill " + (end - start + 1) + " dong."});
  }
  return vtdApp_ok_({message: "Command da nhan: " + action});
}

function vtdApp_rebuildSearchCache(params) {
  const denied = params && params.internalRun ? null : vtdApp_requireAction_("runCommand", params);
  if (denied) return denied;
  params = params || {};
  const ss = vtdApp_ss_();
  const recordSheet = vtdApp_ensureRecordCacheSheet_(ss);
  const tokenSheet = vtdApp_ensureSearchTokenSheet_(ss);
  const reset = params.reset !== false && params.continueBuild !== true;
  if (reset) {
    recordSheet.clearContents();
    tokenSheet.clearContents();
    recordSheet.appendRow(vtdApp_recordCacheHeaders_());
    tokenSheet.appendRow(vtdApp_searchTokenHeaders_());
    PropertiesService.getScriptProperties().setProperty("VTD_SEARCH_CACHE_BUILD", JSON.stringify({
      phase: 0,
      nextRow: 2,
      batchSize: Math.min(Number(params.batchSize) || 700, 1500),
      processed: 0,
      startedAt: new Date().toISOString()
    }));
  }
  return vtdApp_rebuildSearchCacheBatch_(ss, recordSheet, tokenSheet);
}

function vtdApp_rebuildSearchCacheBatch_(ss, recordSheet, tokenSheet) {
  const props = PropertiesService.getScriptProperties();
  const rawState = props.getProperty("VTD_SEARCH_CACHE_BUILD");
  const state = rawState ? JSON.parse(rawState) : {phase: 0, nextRow: 2, batchSize: 700, processed: 0};
  const sheetNames = [VTD_APP.rawFullSheet, VTD_APP.rawSheet];
  const fullCtLookup = vtdApp_fullCtLookupMap_();
  const recordRows = [];
  const tokenRows = [];
  let processedThisRun = 0;
  const maxRowsThisRun = Math.min(Number(state.batchSize) || 700, 1500);

  while (state.phase < sheetNames.length && processedThisRun < maxRowsThisRun) {
    const sheet = ss.getSheetByName(sheetNames[state.phase]);
    if (!sheet || sheet.getLastRow() <= 1) {
      state.phase++;
      state.nextRow = 2;
      continue;
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (state.nextRow > lastRow) {
      state.phase++;
      state.nextRow = 2;
      continue;
    }
    const take = Math.min(maxRowsThisRun - processedThisRun, lastRow - state.nextRow + 1);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const range = sheet.getRange(state.nextRow, 1, take, lastCol);
    const values = range.getValues();
    const displayValues = range.getDisplayValues();
    const items = [];
    const rawIds = {};
    values.forEach((rowValues, offset) => {
      if (rowValues.join("") === "") return;
      const rowObj = {};
      const displayRow = {};
      headers.forEach((header, i) => {
        const key = vtdApp_norm_(header);
        rowObj[key] = rowValues[i];
        displayRow[key] = displayValues[offset][i];
      });
      const rowNumber = state.nextRow + offset;
      const item = {source: sheet.getName(), rowNumber: rowNumber, headers: headers, values: rowValues, displayValues: displayValues[offset], row: rowObj, displayRow: displayRow};
      const c = vtdApp_enrichCompactFromFullCt_(vtdApp_compactRecord_(item), fullCtLookup);
      c.id = String(c.id || c.maDon || sheet.getName() + "_" + rowNumber).trim();
      rawIds[c.id] = true;
      if (c.maDon) rawIds[String(c.maDon).trim()] = true;
      items.push({compact: c, rowNumber: rowNumber});
    });
    const skuByRaw = vtdApp_readSkuByRawIds_(ss, rawIds);
    items.forEach(item => {
      const skuRows = []
        .concat(skuByRaw[item.compact.id] || [])
        .concat(skuByRaw[String(item.compact.maDon || "").trim()] || []);
      const cacheRow = vtdApp_buildRecordCacheRow_(item.compact, vtdApp_uniqueSkuRows_(skuRows));
      recordRows.push(cacheRow);
      Array.prototype.push.apply(tokenRows, vtdApp_buildTokenRowsFromCacheRow_(cacheRow));
    });
    state.nextRow += take;
    processedThisRun += take;
    state.processed = Number(state.processed || 0) + take;
    if (recordRows.length >= 500) vtdApp_flushRows_(recordSheet, recordRows);
    if (tokenRows.length >= 1500) vtdApp_flushRows_(tokenSheet, tokenRows);
  }
  vtdApp_flushRows_(recordSheet, recordRows);
  vtdApp_flushRows_(tokenSheet, tokenRows);

  const done = state.phase >= sheetNames.length;
  if (done) props.deleteProperty("VTD_SEARCH_CACHE_BUILD");
  else props.setProperty("VTD_SEARCH_CACHE_BUILD", JSON.stringify(state));

  return vtdApp_ok_({
    done: done,
    message: done ? "DONE - Da build xong search cache." : "Dang build search cache, chay tiep continueRebuildIndexNow().",
    processedThisRun: processedThisRun,
    processedTotal: state.processed || 0,
    nextSheet: done ? "" : sheetNames[state.phase],
    nextRow: done ? "" : state.nextRow,
    recordRows: Math.max(0, recordSheet.getLastRow() - 1),
    tokenRows: Math.max(0, tokenSheet.getLastRow() - 1)
  });
}

function vtdApp_searchRecordCache_(query, filters, limit) {
  const ss = vtdApp_ss_();
  const recordSheet = ss.getSheetByName(VTD_APP.recordCacheSheet);
  const tokenSheet = ss.getSheetByName(VTD_APP.searchTokenSheet);
  if (!recordSheet || !tokenSheet || recordSheet.getLastRow() <= 1 || tokenSheet.getLastRow() <= 1) return null;
  const queryTokens = vtdApp_searchInputTokens_(query);
  const sets = [];
  if (queryTokens.length) {
    let querySet = {};
    queryTokens.forEach(token => {
      vtdApp_findTokenRows_(tokenSheet, token).forEach(row => querySet[String(row[2] || "")] = true);
    });
    sets.push(querySet);
  }
  const dates = vtdApp_dateKeysBetween_(filters.fromDate, filters.toDate);
  if (dates.length) {
    const dateSet = {};
    dates.forEach(d => vtdApp_findTokenRows_(tokenSheet, d, "DATE").forEach(row => dateSet[String(row[2] || "")] = true));
    sets.push(dateSet);
  }
  if (filters.slog) {
    const slogSet = {};
    vtdApp_findTokenRows_(tokenSheet, filters.slog, "SLOG").forEach(row => slogSet[String(row[2] || "")] = true);
    sets.push(slogSet);
  }
  if (filters.loaiHoan) {
    const loaiSet = {};
    vtdApp_findTokenRows_(tokenSheet, filters.loaiHoan, "LOAI").forEach(row => loaiSet[String(row[2] || "")] = true);
    sets.push(loaiSet);
  }

  let ids = [];
  if (sets.length) {
    let merged = sets[0] || {};
    for (let i = 1; i < sets.length; i++) merged = vtdApp_intersectIdSets_(merged, sets[i]);
    ids = Object.keys(merged);
  } else {
    ids = null;
  }

  let records = ids ? vtdApp_recordsByIds_(recordSheet, ids, limit) : vtdApp_latestCachedRecords_(recordSheet, limit);
  if (!records.length && queryTokens.length) records = vtdApp_searchRecordCacheText_(recordSheet, queryTokens[0], filters, limit);
  return records;
}

function vtdApp_recordCacheHeaders_() {
  return ["RAW_ID", "SourceSheet", "SourceRow", "DateKey", "Month", "NgayHoanTra", "ThoiGianNhap", "MaDon", "Slog", "LoaiHoan", "UserEmail", "PO", "SoHoaDon", "OD", "KhachHang", "QR", "TinhTrang", "ThongTin", "SyncLabel", "PJ2Status", "SkuNames", "Serials", "Lots", "SkuJson", "SearchText"];
}

function vtdApp_searchTokenHeaders_() {
  return ["Token", "TokenType", "RAW_ID", "DateKey", "Month", "Slog", "LoaiHoan"];
}

function vtdApp_buildRecordCacheRow_(c, skus) {
  skus = skus || [];
  const dateKey = vtdApp_dateKey_(c.ngayHoanTra || c.thoiGianNhap);
  const month = dateKey ? dateKey.slice(0, 7) : "";
  const syncLabel = vtdApp_cleanStatusLabel_(c.syncLabel || vtdApp_syncLabel_(c.syncStatus, c.pj2Status, c.retrySync));
  const skuNames = vtdApp_uniqueJoin_(skus.map(x => x.ten));
  const serials = vtdApp_uniqueJoin_(skus.map(x => x.serial));
  const lots = vtdApp_uniqueJoin_(skus.map(x => x.loHsd));
  const skuJson = JSON.stringify(skus.map(x => ({
    ten: x.ten || "",
    name: x.ten || "",
    type: x.type || "",
    soLuong: x.soLuong || "",
    qty: x.soLuong || "",
    tinhTrang: x.tinhTrang || "",
    status: x.tinhTrang || "",
    serial: x.serial || "",
    loHsd: x.loHsd || "",
    lot: x.loHsd || "",
    ghiChu: x.ghiChu || "",
    note: x.ghiChu || ""
  })));
  const searchText = [c.id, c.maDon, c.slog, c.loaiHoan, c.userEmail, c.po, c.soHoaDon, c.od, c.khachHang, c.maQr, c.tinhTrang, c.thongTin, skuNames, serials, lots].join(" ");
  return [c.id, c.source, c.rowNumber, dateKey, month, c.ngayHoanTra, c.thoiGianNhap, c.maDon, c.slog, c.loaiHoan, c.userEmail, c.po, c.soHoaDon, c.od, c.khachHang, c.maQr, c.tinhTrang, c.thongTin, syncLabel, c.pj2Status, skuNames, serials, lots, skuJson, vtdApp_norm_(searchText)];
}

function vtdApp_uniqueSkuRows_(skus) {
  const seen = {};
  const out = [];
  (skus || []).forEach(sku => {
    const key = vtdApp_norm_([sku.ten, sku.type, sku.soLuong, sku.tinhTrang, sku.loHsd, sku.serial, sku.ghiChu].join("|"));
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(sku);
  });
  return out;
}

function vtdApp_buildTokenRowsFromCacheRow_(row) {
  const rawId = row[0];
  const dateKey = row[3];
  const month = row[4];
  const slog = row[8];
  const loai = row[9];
  const base = [rawId, dateKey, month, slog, loai];
  const rows = [];
  const add = (type, value) => {
    vtdApp_tokenValues_(value).forEach(token => rows.push([token, type].concat(base)));
  };
  add("DATE", dateKey);
  add("MONTH", month);
  add("MADON", row[7]);
  add("SLOG", slog);
  add("LOAI", loai);
  add("USER", row[10]);
  add("PO", row[11]);
  add("SOHD", row[12]);
  add("OD", row[13]);
  add("KHACH", row[14]);
  add("QR", row[15]);
  add("SKU", row[20]);
  add("SERIAL", row[21]);
  add("LOT", row[22]);
  return rows;
}

function vtdApp_searchInputTokens_(query) {
  const raw = String(query || "").trim();
  const out = {};
  const add = value => vtdApp_tokenValues_(value).forEach(token => out[token] = true);
  add(raw);
  const m = raw.match(/[?&]qr=([^&#]+)/i);
  if (m) add(decodeURIComponent(m[1]));
  return Object.keys(out);
}

function vtdApp_tokenValues_(value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return [];
  const out = {};
  raw.split(/[;,|\n]+/).forEach(part => {
    const token = vtdApp_norm_(part);
    if (token) out[token] = true;
  });
  const m = raw.match(/[?&]qr=([^&#]+)/i);
  if (m) {
    const qr = vtdApp_norm_(decodeURIComponent(m[1]));
    if (qr) out[qr] = true;
  }
  return Object.keys(out);
}

function vtdApp_findTokenRows_(tokenSheet, key, type) {
  if (!key) return [];
  key = vtdApp_norm_(key);
  const lastRow = tokenSheet.getLastRow();
  const lastCol = tokenSheet.getLastColumn();
  if (lastRow <= 1) return [];
  const cells = tokenSheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(key).matchEntireCell(true).findAll();
  const rowNums = cells.map(c => c.getRow()).sort((a, b) => a - b);
  const rows = vtdApp_getRowsByNumbers_(tokenSheet, rowNums, lastCol);
  if (!type) return rows;
  return rows.filter(row => String(row[1] || "") === type);
}

function vtdApp_intersectIdSets_(a, b) {
  const out = {};
  Object.keys(a || {}).forEach(id => { if (b && b[id]) out[id] = true; });
  return out;
}

function vtdApp_recordsByIds_(recordSheet, ids, limit) {
  const idSet = {};
  ids.forEach(id => { if (id) idSet[String(id)] = true; });
  const values = recordSheet.getDataRange().getValues();
  const out = [];
  for (let i = values.length - 1; i >= 1; i--) {
    if (out.length >= limit) break;
    if (!idSet[String(values[i][0])]) continue;
    out.push(vtdApp_recordCacheRowToCompact_(values[i]));
  }
  return out;
}

function vtdApp_latestCachedRecords_(recordSheet, limit) {
  const values = recordSheet.getDataRange().getValues();
  const out = [];
  for (let i = values.length - 1; i >= 1 && out.length < limit; i--) out.push(vtdApp_recordCacheRowToCompact_(values[i]));
  return out;
}

function vtdApp_searchRecordCacheText_(recordSheet, queryToken, filters, limit) {
  if (!queryToken) return [];
  const lastRow = recordSheet.getLastRow();
  const lastCol = recordSheet.getLastColumn();
  if (lastRow <= 1) return [];
  const searchCol = 25;
  const cells = recordSheet.getRange(2, searchCol, lastRow - 1, 1).createTextFinder(queryToken).matchCase(false).findAll();
  const rowNums = cells.map(c => c.getRow()).sort((a, b) => b - a);
  const out = [];
  rowNums.forEach(rowNum => {
    if (out.length >= limit) return;
    const row = recordSheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    if (!vtdApp_recordCachePassesFilters_(row, filters)) return;
    out.push(vtdApp_recordCacheRowToCompact_(row));
  });
  return out;
}

function vtdApp_recordCachePassesFilters_(row, filters) {
  const dateKey = row[3];
  if (filters.fromDate && dateKey && dateKey < filters.fromDate) return false;
  if (filters.toDate && dateKey && dateKey > filters.toDate) return false;
  if (filters.slog && vtdApp_norm_(row[8]) !== filters.slog) return false;
  if (filters.loaiHoan && vtdApp_norm_(row[9]) !== filters.loaiHoan) return false;
  return true;
}

function vtdApp_recordCacheRowToCompact_(row) {
  let skuItems = [];
  try {
    skuItems = JSON.parse(String(row[23] || "[]"));
  } catch (err) {
    skuItems = [];
  }
  if (!Array.isArray(skuItems) || !skuItems.length) {
    const names = String(row[20] || "").split(";").map(x => x.trim()).filter(Boolean);
    const serials = String(row[21] || "").split(";").map(x => x.trim()).filter(Boolean);
    const lots = String(row[22] || "").split(";").map(x => x.trim()).filter(Boolean);
    skuItems = names.map((name, i) => ({name: name, ten: name, serial: serials[i] || "", loHsd: lots[i] || ""}));
  }
  return {
    id: row[0],
    source: row[1],
    rowNumber: row[2],
    ngayHoanTra: row[5],
    thoiGianNhap: row[6],
    maDon: row[7],
    slog: row[8],
    loaiHoan: row[9],
    userEmail: row[10],
    po: row[11],
    soHoaDon: row[12],
    od: row[13],
    khachHang: row[14],
    maQr: row[15],
    tinhTrang: row[16],
    thongTin: row[17],
    syncLabel: vtdApp_cleanStatusLabel_(row[18]) || "Đã đồng bộ",
    pj2Status: row[19],
    skuItems: skuItems
  };
}

function vtdApp_readSkuByRaw_(ss, limit) {
  const sheet = ss.getSheetByName(VTD_APP.skuSheet);
  const out = {};
  if (!sheet || sheet.getLastRow() <= 1) return out;
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const start = Math.max(2, lastRow - limit + 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const values = sheet.getRange(start, 1, lastRow - start + 1, lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rawId = String(vtdApp_pickFromArray_(row, col, ["raw_id", "raw id"]) || "").trim();
    if (!rawId) continue;
    out[rawId] = out[rawId] || [];
    out[rawId].push({
      ten: vtdApp_pickFromArray_(row, col, ["ten san pham hoan ve", "ten san pham hoan", "ten san pham", "sku"]),
      type: vtdApp_pickFromArray_(row, col, ["type", "loai"]),
      soLuong: vtdApp_pickFromArray_(row, col, ["so luong"]),
      tinhTrang: vtdApp_pickFromArray_(row, col, ["tinh trang"]),
      loHsd: vtdApp_pickFromArray_(row, col, ["ma lo/hsd", "ma lo hsd", "ma lo", "hsd"]),
      serial: vtdApp_pickFromArray_(row, col, ["ma serial", "serial", "ma qr sua bot", "ma qr"]),
      ghiChu: vtdApp_pickFromArray_(row, col, ["ghi chu", "note"])
    });
  }
  return out;
}

function vtdApp_readSkuByRawIds_(ss, rawIds) {
  const out = {};
  rawIds = rawIds || {};
  const wanted = Object.keys(rawIds);
  if (!wanted.length) return out;
  const sheet = ss.getSheetByName(VTD_APP.skuSheet);
  if (!sheet || sheet.getLastRow() <= 1) return out;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rawId = String(vtdApp_pickFromArray_(row, col, ["raw_id", "raw id"]) || "").trim();
    const maDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || "").trim();
    const keys = [rawId, maDon].filter(key => key && rawIds[key]);
    if (!keys.length) continue;
    const sku = {
      ten: vtdApp_pickFromArray_(row, col, ["ten san pham hoan ve", "tên sản phẩm hoàn về", "ten san pham hoan", "ten san pham", "sku"]) || row[2],
      type: vtdApp_pickFromArray_(row, col, ["type", "loai"]) || row[3],
      soLuong: vtdApp_pickFromArray_(row, col, ["so luong", "số lượng", "qty"]) || row[4],
      tinhTrang: vtdApp_pickFromArray_(row, col, ["tinh trang", "tình trạng"]) || row[5],
      loHsd: vtdApp_pickFromArray_(row, col, ["ma lo/hsd", "mã lô/hsd", "ma lo hsd", "ma lo", "mã lô", "hsd"]) || row[6],
      serial: vtdApp_pickFromArray_(row, col, ["ma serial", "mã serial", "serial", "ma qr sua bot", "ma qr", "mã qr"]) || row[7],
      ghiChu: vtdApp_pickFromArray_(row, col, ["ghi chu", "ghi chú", "note"]) || row[9]
    };
    keys.forEach(key => {
      out[key] = out[key] || [];
      out[key].push(sku);
    });
  }
  return out;
}

function vtdApp_attachSkuItemsToCompacts_(ss, records) {
  records = Array.isArray(records) ? records : [];
  const keys = {};
  const orderKeys = {};
  records.forEach(record => {
    if (record && record.id) keys[String(record.id).trim()] = true;
    if (record && record.maDon) {
      keys[String(record.maDon).trim()] = true;
      orderKeys[String(record.maDon).trim()] = true;
    }
  });
  const rawIdsByOrder = vtdApp_rawIdsByOrder_(ss, orderKeys);
  Object.keys(rawIdsByOrder).forEach(maDon => {
    rawIdsByOrder[maDon].forEach(id => { if (id) keys[id] = true; });
  });
  const skuByKey = vtdApp_readSkuByRawIds_(ss, keys);
  records.forEach(record => {
    if (!record) return;
    const siblingRawIds = rawIdsByOrder[String(record.maDon || "").trim()] || [];
    const skuRows = []
      .concat(skuByKey[String(record.id || "").trim()] || [])
      .concat(skuByKey[String(record.maDon || "").trim()] || [])
      .concat.apply([], siblingRawIds.map(id => skuByKey[id] || []));
    const nextRows = vtdApp_uniqueSkuRows_(skuRows);
    if (nextRows.length) record.skuItems = nextRows;
  });
  return records;
}

function vtdApp_rawIdsByOrder_(ss, orderKeys) {
  orderKeys = orderKeys || {};
  const wanted = {};
  Object.keys(orderKeys).forEach(key => {
    const norm = vtdApp_norm_(key);
    if (norm) wanted[norm] = key;
  });
  const out = {};
  if (!Object.keys(wanted).length) return out;
  [VTD_APP.rawSheet, VTD_APP.rawFullSheet].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const col = vtdApp_headerMap_(headers);
    const readRows = Math.min(Math.max(VTD_APP.maxSearchRows || 5000, 20000), sheet.getLastRow() - 1);
    const start = Math.max(2, sheet.getLastRow() - readRows + 1);
    const values = sheet.getRange(start, 1, sheet.getLastRow() - start + 1, lastCol).getDisplayValues();
    values.forEach(row => {
      const maDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || row[3] || "").trim();
      const canonical = wanted[vtdApp_norm_(maDon)];
      if (!canonical) return;
      const rawId = String(vtdApp_pickFromArray_(row, col, ["id", "raw_id", "raw id"]) || row[0] || "").trim();
      if (!rawId) return;
      out[canonical] = out[canonical] || [];
      if (out[canonical].indexOf(rawId) < 0) out[canonical].push(rawId);
    });
  });
  return out;
}

function vtdApp_fullCtLookupMap_() {
  const map = {};
  try {
    const ss = SpreadsheetApp.openById(VTD_APP.fullCtSpreadsheetId);
    const sheets = ss.getSheets().filter(sh => /^T\d{1,2}\/\d{2}$/i.test(sh.getName()));
    const tongHop = ss.getSheetByName("TongHop");
    if (tongHop) sheets.push(tongHop);
    sheets.forEach(sh => vtdApp_addFullCtLookupSheet_(map, sh));
  } catch (err) {}
  return map;
}

function vtdApp_addFullCtLookupSheet_(map, sh) {
  if (!sh || sh.getLastRow() <= 1) return;
  const lastCol = Math.min(16, sh.getLastColumn());
  const values = sh.getRange(1, 1, sh.getLastRow(), lastCol).getDisplayValues();
  if (!values.length) return;
  const col = vtdApp_headerMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rec = {
      od: String(vtdApp_pickFromArray_(row, col, ["od"]) || "").trim(),
      khachHang: String(vtdApp_pickFromArray_(row, col, ["khach hang", "khách hàng", "customer", "ten khach hang", "tên khách hàng"]) || "").trim(),
      po: String(vtdApp_pickFromArray_(row, col, ["ma po", "mã po", "po"]) || "").trim(),
      soHoaDon: String(vtdApp_pickFromArray_(row, col, ["so hoa don", "số hóa đơn", "invoice"]) || "").trim(),
      maDon: String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || "").trim(),
      sourceSheet: sh.getName()
    };
    // Ma don is the primary source of truth. PO/OD/invoice are fallback search keys.
    ["maDon", "po", "od", "soHoaDon"].forEach(field => {
      vtdApp_lookupKeys_(rec[field]).forEach(key => {
        if (key && !map[key]) map[key] = rec;
      });
    });
  }
}

function vtdApp_enrichCompactFromFullCt_(compact, lookup) {
  lookup = lookup || {};
  if (!compact) return compact;
  const keys = [];
  [compact.maDon, compact.po, compact.od, compact.soHoaDon].forEach(value => {
    Array.prototype.push.apply(keys, vtdApp_lookupKeys_(value));
  });
  let rec = null;
  for (let i = 0; i < keys.length; i++) {
    if (lookup[keys[i]]) {
      rec = lookup[keys[i]];
      break;
    }
  }
  if (!rec) return compact;
  if (!compact.khachHang && rec.khachHang) compact.khachHang = rec.khachHang;
  if (!compact.od && rec.od) compact.od = rec.od;
  if (!compact.soHoaDon && rec.soHoaDon) compact.soHoaDon = rec.soHoaDon;
  if (!compact.po && rec.po) compact.po = rec.po;
  return compact;
}

function vtdApp_enrichCompactsFromFullCtByOrders_(compacts) {
  compacts = Array.isArray(compacts) ? compacts : [];
  const wanted = {};
  compacts.forEach(compact => {
    const key = vtdApp_norm_(compact && compact.maDon);
    if (key) wanted[key] = true;
  });
  const wantedKeys = Object.keys(wanted);
  if (!wantedKeys.length) return compacts;
  const found = {};
  try {
    const ss = SpreadsheetApp.openById(VTD_APP.fullCtSpreadsheetId);
    const sheets = ss.getSheets().filter(sh => /^T\d{1,2}\/\d{2}$/i.test(sh.getName()) || sh.getName() === "TongHop");
    for (let s = 0; s < sheets.length && Object.keys(found).length < wantedKeys.length; s++) {
      const sh = sheets[s];
      if (!sh || sh.getLastRow() <= 1) continue;
      const lastCol = Math.min(16, sh.getLastColumn());
      const values = sh.getRange(1, 1, sh.getLastRow(), lastCol).getDisplayValues();
      if (!values.length) continue;
      const col = vtdApp_headerMap_(values[0]);
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const maDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn", "ma don ghtk", "mã đơn ghtk"]) || "").trim();
        const key = vtdApp_norm_(maDon);
        if (!wanted[key] || found[key]) continue;
        found[key] = {
          od: String(vtdApp_pickFromArray_(row, col, ["od"]) || "").trim(),
          khachHang: String(vtdApp_pickFromArray_(row, col, ["khach hang", "khách hàng", "customer", "ten khach hang", "tên khách hàng"]) || "").trim(),
          poFullCt: String(vtdApp_pickFromArray_(row, col, ["ma po", "mã po", "po"]) || "").trim(),
          soHoaDon: String(vtdApp_pickFromArray_(row, col, ["so hoa don", "số hóa đơn", "invoice"]) || "").trim()
        };
      }
    }
  } catch (err) {}
  compacts.forEach(compact => {
    const rec = found[vtdApp_norm_(compact && compact.maDon)];
    if (!rec) return;
    if (rec.khachHang) compact.khachHang = rec.khachHang;
    if (rec.od) compact.od = rec.od;
    if (rec.soHoaDon) compact.soHoaDon = rec.soHoaDon;
    if (!compact.po && rec.poFullCt) compact.po = rec.poFullCt;
  });
  return compacts;
}

function vtdApp_syncTargetKey_(compact) {
  const loai = vtdApp_norm_(compact && compact.loaiHoan);
  if (loai.indexOf("chung tu") >= 0) return "chungTu";
  return "hangHoan";
}

function vtdApp_isNotSynced_(compact) {
  if (!compact) return false;
  const syncStatus = String(compact.syncStatus || "").trim().toUpperCase();
  const pj2Status = String(compact.pj2Status || "").trim().toUpperCase();
  const retrySync = String(compact.retrySync || "").trim().toUpperCase();
  const target = vtdApp_syncTargetKey_(compact);
  const targetStatus = target === "chungTu" ? syncStatus : pj2Status;
  const status = String(targetStatus || retrySync || "").trim().toUpperCase();
  if (!status) return true;
  if (status === "DONE" || status === "ROUTE_PJ2") return false;
  return status === "YES" || status === "RUNNING" || status === "PENDING" || status.indexOf("ERROR") >= 0 || status.indexOf("IMG") >= 0;
}

function vtdApp_isTargetNotSynced_(value) {
  const status = String(value || "").trim().toUpperCase();
  if (!status) return false;
  if (status === "DONE" || status === "ROUTE_PJ2") return false;
  return status === "YES" || status === "RUNNING" || status === "PENDING" || status.indexOf("ERROR") >= 0 || status.indexOf("IMG") >= 0;
}

function vtdApp_lookupKeys_(value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return [];
  const out = {};
  const add = v => {
    const key = vtdApp_norm_(v);
    if (key) out[key] = true;
  };
  add(raw);
  raw.split(/[\n;,|]+/).forEach(add);
  raw.replace(/\bPO\b/ig, " ").split(/[\n;,|]+/).forEach(add);
  const nums = raw.match(/\d{5,}/g) || [];
  nums.forEach(add);
  return Object.keys(out);
}

function vtdApp_appendSearchCacheForRawRow_(sheet, rowNumber) {
  if (!sheet || rowNumber <= 1) return;
  const ss = vtdApp_ss_();
  const recordSheet = vtdApp_ensureRecordCacheSheet_(ss);
  const tokenSheet = vtdApp_ensureSearchTokenSheet_(ss);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const range = sheet.getRange(rowNumber, 1, 1, lastCol);
  const values = range.getValues()[0];
  const displayValues = range.getDisplayValues()[0];
  const rowObj = {};
  const displayRow = {};
  headers.forEach((header, i) => {
    const key = vtdApp_norm_(header);
    rowObj[key] = values[i];
    displayRow[key] = displayValues[i];
  });
  const item = {source: sheet.getName(), rowNumber: rowNumber, headers: headers, values: values, displayValues: displayValues, row: rowObj, displayRow: displayRow};
  const c = vtdApp_enrichCompactFromFullCt_(vtdApp_compactRecord_(item), vtdApp_fullCtLookupMap_());
  const skus = vtdApp_findSku_(c.id);
  const cacheRow = vtdApp_buildRecordCacheRow_(c, skus);
  recordSheet.appendRow(cacheRow);
  const tokens = vtdApp_buildTokenRowsFromCacheRow_(cacheRow);
  if (tokens.length) tokenSheet.getRange(tokenSheet.getLastRow() + 1, 1, tokens.length, tokens[0].length).setValues(tokens);
}

function vtdApp_ensureRecordCacheSheet_(ss) {
  let sh = vtdApp_getSheetByNameRetry_(ss, VTD_APP.recordCacheSheet);
  if (!sh) {
    sh = vtdApp_insertSheetRetry_(ss, VTD_APP.recordCacheSheet);
    sh.appendRow(vtdApp_recordCacheHeaders_());
  }
  return sh;
}

function vtdApp_ensureSearchTokenSheet_(ss) {
  let sh = vtdApp_getSheetByNameRetry_(ss, VTD_APP.searchTokenSheet);
  if (!sh) {
    sh = vtdApp_insertSheetRetry_(ss, VTD_APP.searchTokenSheet);
    sh.appendRow(vtdApp_searchTokenHeaders_());
  }
  return sh;
}

function vtdApp_getSheetByNameRetry_(ss, name) {
  let lastErr = null;
  for (let i = 0; i < 4; i++) {
    try {
      return ss.getSheetByName(name);
    } catch (err) {
      lastErr = err;
      Utilities.sleep(600 + i * 700);
    }
  }
  throw lastErr;
}

function vtdApp_insertSheetRetry_(ss, name) {
  let lastErr = null;
  for (let i = 0; i < 4; i++) {
    try {
      return ss.insertSheet(name);
    } catch (err) {
      lastErr = err;
      Utilities.sleep(700 + i * 800);
      const existing = ss.getSheetByName(name);
      if (existing) return existing;
    }
  }
  throw lastErr;
}

function vtdApp_deleteOldSearchCacheSheets_(ss) {
  [VTD_APP.recordCacheSheet, VTD_APP.searchTokenSheet, VTD_APP.rowMapSheet].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  vtdApp_deleteOldSearchIndexSheets_(ss);
}

function vtdApp_flushRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  rows.splice(0, rows.length);
}

function vtdApp_uniqueJoin_(values) {
  const seen = {};
  const out = [];
  (values || []).forEach(value => {
    const text = String(value || "").trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    out.push(text);
  });
  return out.join("; ");
}

function vtdApp_searchRowMap_(query, filters, limit) {
  const ss = vtdApp_ss_();
  const map = ss.getSheetByName(VTD_APP.rowMapSheet);
  if (!map || map.getLastRow() <= 1) return null;
  let hits = [];
  if (query) {
    ["MADON", "PO", "SOHD", "OD", "QR", "SERIAL", "USER"].forEach(type => {
      hits = hits.concat(vtdApp_findRowMapRows_(map, type + "|" + query));
    });
  }
  const dates = vtdApp_dateKeysBetween_(filters.fromDate, filters.toDate);
  if (dates.length) {
    let dateHits = [];
    dates.forEach(d => dateHits = dateHits.concat(vtdApp_findRowMapRows_(map, "DATE|" + d)));
    hits = hits.length ? vtdApp_intersectRowMapHits_(hits, dateHits) : dateHits;
  }
  if (!query && !dates.length) return null;
  const seenRefs = {};
  const refs = [];
  hits.forEach(row => {
    const ref = {
      source: row[2],
      rowNumber: Number(row[3]),
      dateKey: row[4],
      slog: row[5],
      loaiHoan: row[6]
    };
    if (!ref.source || !ref.rowNumber) return;
    if (filters.slog && vtdApp_norm_(ref.slog) !== filters.slog) return;
    if (filters.loaiHoan && vtdApp_norm_(ref.loaiHoan) !== filters.loaiHoan) return;
    const key = ref.source + ":" + ref.rowNumber;
    if (seenRefs[key]) return;
    seenRefs[key] = true;
    refs.push(ref);
  });
  refs.sort((a, b) => b.rowNumber - a.rowNumber);
  return vtdApp_fetchCompactsByRefs_(ss, refs.slice(0, limit), query, filters);
}

function vtdApp_findRowMapRows_(map, key) {
  if (!key) return [];
  const lastRow = map.getLastRow();
  const lastCol = map.getLastColumn();
  if (lastRow <= 1) return [];
  const cells = map.getRange(2, 1, lastRow - 1, 1).createTextFinder(key).matchEntireCell(true).findAll();
  const rowNums = cells.map(c => c.getRow()).sort((a, b) => a - b);
  return vtdApp_getRowsByNumbers_(map, rowNums, lastCol);
}

function vtdApp_getRowsByNumbers_(sheet, rowNums, lastCol) {
  const out = [];
  let i = 0;
  while (i < rowNums.length) {
    let start = rowNums[i];
    let end = start;
    i++;
    while (i < rowNums.length && rowNums[i] === end + 1) {
      end = rowNums[i];
      i++;
    }
    out.push.apply(out, sheet.getRange(start, 1, end - start + 1, lastCol).getValues());
  }
  return out;
}

function vtdApp_intersectRowMapHits_(a, b) {
  const set = {};
  b.forEach(row => set[row[2] + ":" + row[3]] = true);
  return a.filter(row => set[row[2] + ":" + row[3]]);
}

function vtdApp_dateKeysBetween_(fromDate, toDate) {
  if (!fromDate && !toDate) return [];
  const startKey = fromDate || toDate;
  const endKey = toDate || fromDate;
  const start = new Date(startKey + "T00:00:00");
  const end = new Date(endKey + "T00:00:00");
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || start > end) return [];
  const out = [];
  const d = new Date(start);
  while (d <= end && out.length <= 62) {
    out.push(Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh", "yyyy-MM-dd"));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function vtdApp_fetchCompactsByRefs_(ss, refs, query, filters) {
  const out = [];
  const grouped = {};
  refs.forEach(ref => {
    grouped[ref.source] = grouped[ref.source] || [];
    grouped[ref.source].push(ref.rowNumber);
  });
  Object.keys(grouped).forEach(source => {
    const sheet = ss.getSheetByName(source);
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    vtdApp_getRowsByNumbers_(sheet, grouped[source].sort((a, b) => a - b), lastCol).forEach((values, idx) => {
      const rowObj = {};
      headers.forEach((header, i) => rowObj[vtdApp_norm_(header)] = values[i]);
      const rowNumber = grouped[source].sort((a, b) => a - b)[idx];
      const item = {source: source, rowNumber: rowNumber, headers: headers, values: values, row: rowObj};
      const compact = vtdApp_compactRecord_(item);
      if (!vtdApp_recordPassesFilters_(item, compact, query, filters)) return;
      out.push(compact);
    });
  });
  return out;
}

function vtdApp_searchIndex_(query, filters, limit) {
  const ss = vtdApp_ss_();
  const monthNames = ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => String(name).indexOf(VTD_APP.indexPrefix) === 0)
    .sort()
    .reverse();
  if (!monthNames.length) monthNames.push(vtdApp_indexSheetName_(new Date()));
  const out = [];
  const seen = {};
  monthNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() <= 1) return;
    const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const col = vtdApp_headerMap_(values[0]);
    for (let i = values.length - 1; i >= 1; i--) {
      if (out.length >= limit) break;
      const row = values[i];
      const searchText = vtdApp_norm_(vtdApp_pickFromArray_(row, col, ["searchtext"]));
      const slog = vtdApp_pickFromArray_(row, col, ["slog"]);
      const loai = vtdApp_pickFromArray_(row, col, ["loaihoan"]);
      const date = vtdApp_dateKey_(vtdApp_pickFromArray_(row, col, ["createddate"]));
      if (query && searchText.indexOf(query) < 0) continue;
      if (filters.slog && vtdApp_norm_(slog) !== filters.slog) continue;
      if (filters.loaiHoan && vtdApp_norm_(loai) !== filters.loaiHoan) continue;
      if (filters.fromDate && date && date < filters.fromDate) continue;
      if (filters.toDate && date && date > filters.toDate) continue;
      const source = vtdApp_pickFromArray_(row, col, ["sourcesheet"]);
      const rowNumber = vtdApp_pickFromArray_(row, col, ["sourcerow"]);
      const rawId = vtdApp_pickFromArray_(row, col, ["raw_id"]);
      const key = rawId || source + ":" + rowNumber;
      if (seen[key]) continue;
      seen[key] = true;
      out.push({
        source: source,
        rowNumber: rowNumber,
        id: rawId,
        maDon: vtdApp_pickFromArray_(row, col, ["madon"]),
        slog: slog,
        loaiHoan: loai,
        ngayHoanTra: vtdApp_pickFromArray_(row, col, ["createddate"]),
        thoiGianNhap: vtdApp_pickFromArray_(row, col, ["createdtime"]),
        po: vtdApp_pickFromArray_(row, col, ["po"]),
        maQr: vtdApp_pickFromArray_(row, col, ["qr"]),
        soHoaDon: vtdApp_pickFromArray_(row, col, ["sohoadon"]),
        od: vtdApp_pickFromArray_(row, col, ["od"]),
        khachHang: vtdApp_pickFromArray_(row, col, ["khachhang"]),
        userEmail: vtdApp_pickFromArray_(row, col, ["useremail", "user_email", "user email"]),
        syncLabel: vtdApp_pickFromArray_(row, col, ["synclabel"]),
        syncState: vtdApp_pickFromArray_(row, col, ["synclabel"])
      });
    }
  });
  return out;
}

function vtdApp_appendIndexForRawRow_(sheet, rowNumber) {
  if (!sheet || rowNumber <= 1) return;
  const ss = vtdApp_ss_();
  const index = vtdApp_ensureIndexSheet_(ss, vtdApp_indexSheetName_(new Date()));
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const col = vtdApp_headerMap_(headers);
  const rowObj = {};
  headers.forEach((header, i) => rowObj[vtdApp_norm_(header)] = values[i]);
  const item = {source: sheet.getName(), rowNumber: rowNumber, headers: headers, values: values, row: rowObj};
  const c = vtdApp_compactRecord_(item);
  const date = vtdApp_toDisplay_(vtdApp_pick_(rowObj, ["ngay hoan tra", "ngay nhap"])) || vtdApp_dateKey_(new Date());
  const time = vtdApp_toDisplay_(vtdApp_pick_(rowObj, ["thoi gian nhap hoan", "thoi gian", "timestamp"]));
  const user = vtdApp_pick_(rowObj, ["user_email", "user email", "email", "nguoi nhap", "người nhập", "nguoi thao tac", "người thao tác", "user thao tac", "user thao tác", "created by", "createdby"]);
  const soHd = vtdApp_pick_(rowObj, ["so hoa don", "so hoa don", "invoice"]);
  const od = vtdApp_pick_(rowObj, ["od"]);
  const kh = vtdApp_pick_(rowObj, ["khach hang", "customer"]);
  const searchText = [c.id, c.maDon, c.slog, c.loaiHoan, c.po, c.maQr, soHd, od, kh, user, c.tinhTrang, c.thongTin].join(" ");
  index.appendRow([c.id, sheet.getName(), rowNumber, date, time, c.maDon, c.slog, c.loaiHoan, c.po, c.maQr, soHd, od, kh, user, c.syncLabel, c.pj2Status, "", searchText]);
}

function vtdApp_appendRowMapForRawRow_(sheet, rowNumber) {
  if (!sheet || rowNumber <= 1) return;
  const ss = vtdApp_ss_();
  const map = vtdApp_ensureRowMapSheet_(ss);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const rowObj = {};
  headers.forEach((header, i) => rowObj[vtdApp_norm_(header)] = values[i]);
  const item = {source: sheet.getName(), rowNumber: rowNumber, headers: headers, values: values, row: rowObj};
  const c = vtdApp_compactRecord_(item);
  const dateKey = vtdApp_dateKey_(vtdApp_pick_(rowObj, ["ngay hoan tra", "ngày hoàn trả", "ngay nhap", "timestamp", "thoi gian nhap hoan", "thời gian nhập hoàn"]));
  const rows = [];
  const base = [sheet.getName(), rowNumber, dateKey, c.slog || "", c.loaiHoan || "", c.maDon || "", c.userEmail || ""];
  const add = (type, value) => {
    const keyValue = type === "DATE" ? String(value || "") : vtdApp_norm_(value);
    if (!keyValue) return;
    rows.push([type + "|" + keyValue, type].concat(base));
  };
  add("DATE", dateKey);
  add("MADON", c.maDon);
  add("PO", c.po);
  add("SOHD", c.soHoaDon);
  add("OD", c.od);
  add("QR", c.maQr);
  add("SERIAL", c.serial);
  add("USER", c.userEmail);
  add("SLOG", c.slog);
  add("LOAI", c.loaiHoan);
  if (rows.length) map.getRange(map.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function vtdApp_ensureIndexSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(["RAW_ID", "SourceSheet", "SourceRow", "CreatedDate", "CreatedTime", "MaDon", "Slog", "LoaiHoan", "PO", "QR", "SoHoaDon", "OD", "KhachHang", "UserEmail", "SyncLabel", "PJ2Label", "ImageFolderUrl", "SearchText"]);
  }
  return sh;
}

function vtdApp_ensureRowMapSheet_(ss) {
  let sh = ss.getSheetByName(VTD_APP.rowMapSheet);
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.rowMapSheet);
    sh.appendRow(["MapKey", "KeyType", "SourceSheet", "SourceRow", "DateKey", "Slog", "LoaiHoan", "MaDon", "UserEmail"]);
  }
  return sh;
}

function vtdApp_deleteOldSearchIndexSheets_(ss) {
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (String(name).indexOf(VTD_APP.indexPrefix) === 0) ss.deleteSheet(sheet);
  });
}

function vtdApp_dashboardCache_() {
  try {
    const ss = vtdApp_ss_();
    const sh = ss.getSheetByName(VTD_APP.dashboardCacheSheet);
    if (!sh || sh.getLastRow() <= 1) return null;
    const values = sh.getDataRange().getValues();
    const col = vtdApp_headerMap_(values[0]);
    const updatedAt = vtdApp_pickFromArray_(values[1], col, ["updatedat"]);
    if (!updatedAt) return null;
    const age = (Date.now() - new Date(updatedAt).getTime()) / 1000;
    if (!isFinite(age) || age > 180) return null;
    const result = {bySlog: [], byStatus: [], recent: []};
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const group = String(vtdApp_pickFromArray_(row, col, ["group"]) || "").trim();
      const label = String(vtdApp_pickFromArray_(row, col, ["label"]) || "").trim();
      const value = Number(vtdApp_pickFromArray_(row, col, ["value"]) || 0);
      if (!label) continue;
      if (group === "stat") result[label] = value;
      if (group === "slog") result.bySlog.push({key: label, value: value});
      if (group === "status") result.byStatus.push({key: label, value: value});
    }
    if (!result.totalKnown && !result.todayTotal && !result.bySlog.length && !result.byStatus.length) return null;
    return result;
  } catch (err) {
    return null;
  }
}

function vtdApp_indexSheetName_(date) {
  return VTD_APP.indexPrefix + Utilities.formatDate(date || new Date(), "Asia/Saigon", "yyyy_MM");
}

function vtdApp_authShape_(email, auth) {
  auth = auth || {};
  return {
    allowed: true,
    email: email,
    role: auth.role === "admin" ? "admin" : "staff",
    screens: auth.role === "admin" ? VTD_PERMISSIONS.allScreens.slice() : (auth.screens || ["home", "search", "input", "store", "settings"]),
    actions: auth.role === "admin" ? VTD_PERMISSIONS.allActions.slice() : (auth.actions || ["dashboard", "search", "getRecord", "saveRaw", "uploadImage", "queueRecord", "storeInfo"])
  };
}

function vtdApp_vnDayKey_(date) {
  return Utilities.formatDate(date || new Date(), "Asia/Saigon", "yyyy-MM-dd");
}

function vtdApp_secondsToMidnight_() {
  const now = new Date();
  const midnight = new Date(now);
  midnight.setHours(24, 0, 0, 0);
  return Math.max(60, Math.floor((midnight.getTime() - now.getTime()) / 1000));
}

function vtdApp_toKeyRows_(map) {
  return Object.keys(map).map(key => ({key: key, value: map[key]}));
}

function vtdApp_storeType_(text) {
  const x = vtdApp_norm_(text);
  if (x.indexOf("aeon") >= 0) return "Aeon";
  if (x.indexOf("winmart") >= 0 || x.indexOf("win mart") >= 0 || x.indexOf("st-wm") >= 0) return "Winmart";
  if (x.indexOf("lotte") >= 0) return "Lotte";
  if (x.indexOf("bach hoa") >= 0 || x.indexOf("bhx") >= 0) return "BHX";
  if (x.indexOf("nha thuoc") >= 0 || x.indexOf("long chau") >= 0) return "Nha thuoc";
  return "Khac";
}

function vtdApp_log_(level, action, email, rawId, maDon, result, message, meta) {
  try {
    const ss = vtdApp_ss_();
    const sh = ss.getSheetByName(VTD_APP.apiLogSheet);
    if (!sh) return;
    sh.appendRow([new Date(), level, action, email, rawId, maDon, result, message, meta || ""]);
  } catch (err) {}
}

/************************************************
VTD APK/API BRIDGE - 2026-05-28

This final block is intentionally placed at the end so Apps Script uses these
latest doGet/doPost/storeInfo definitions.
************************************************/

function doGet(e) {
  try {
    e = e || {};
    if (e.parameter && e.parameter.action) {
      return vtdApp_json_(vtdApp_apiDispatch_(String(e.parameter.action), e.parameter || {}));
    }
    return HtmlService
      .createTemplateFromFile("Index")
      .evaluate()
      .setTitle("Chung tu VTD")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return vtdApp_json_(vtdApp_fail_(String(err && err.stack || err)));
  }
}

function doPost(e) {
  try {
    let body = {};
    try {
      body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    } catch (err) {
      body = {};
    }
    const action = String(body.action || (e && e.parameter && e.parameter.action) || "").trim();
    const params = body.params || body || {};
    return vtdApp_json_(vtdApp_apiDispatch_(action, params));
  } catch (err) {
    return vtdApp_json_(vtdApp_fail_(String(err && err.stack || err)));
  }
}

function vtdApp_json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function vtdApp_apiDispatch_(action, params) {
  try {
    params = params || {};
    const map = {
      publicConfig: vtdApp_publicConfig,
      login: vtdApp_login,
      logout: vtdApp_logout,
      changePin: vtdApp_changePin,
      changePass: vtdApp_changePin,
      forgotPass: vtdApp_forgotPass,
      forgotPassRequests: vtdApp_forgotPassRequests,
      resetForgotPass: vtdApp_resetForgotPass,
      init: vtdApp_init,
      systemConfig: vtdApp_systemConfig,
      saveSystemConfig: vtdApp_saveSystemConfig,
      publicTheme: vtdApp_publicTheme,
      dashboard: vtdApp_dashboard,
      search: vtdApp_search,
      getRecord: vtdApp_getRecord,
      saveRaw: vtdApp_saveRaw,
      updateRecord: vtdApp_updateRecord,
      uploadImage: vtdApp_uploadImage,
      uploadImages: vtdApp_uploadImages,
      queueRecord: vtdApp_queueRecord,
      runCommand: vtdApp_runCommand,
      storeInfo: vtdApp_storeInfo,
      listProducts: vtdApp_listProducts,
      listSku: vtdApp_listProducts,
      lookupOrder: vtdApp_lookupOrder,
      lookupOrderManifest: vtdApp_lookupOrderManifest,
      lookupOrderPage: vtdApp_lookupOrderPage,
      cacheDataManifest: vtdApp_cacheDataManifest,
      cacheDataPage: vtdApp_cacheDataPage,
      reportLookupCacheStatus: vtdApp_reportLookupCacheStatus,
      lookupCacheStatus: vtdApp_lookupCacheStatus,
      deleteLookupCacheStatus: vtdApp_deleteLookupCacheStatus,
      reportQueueErrors: vtdApp_reportQueueErrors,
      queueErrors: vtdApp_queueErrors,
      deleteQueueError: vtdApp_deleteQueueError,
      requestQueueRetry: vtdApp_requestQueueRetry,
      queueRetryRequests: vtdApp_queueRetryRequests,
      ackQueueRetryRequests: vtdApp_ackQueueRetryRequests,
      adminConfig: vtdApp_adminConfig,
      savePermission: vtdApp_savePermission,
      deletePermission: vtdApp_deletePermission,
      dynamicConfig: vtdApp_dynamicConfig,
      saveDynamicConfig: vtdApp_saveDynamicConfig,
      dynamicListSheets: vtdApp_dynamicListSheets,
      dynamicReadColumns: vtdApp_dynamicReadColumns,
      dynamicReadRows: vtdApp_dynamicReadRows,
      dynamicSaveRow: vtdApp_dynamicSaveRow,
      activityPing: vtdApp_activityPing,
      loginReport: vtdApp_loginReport
    };
    if (!action || !map[action]) return vtdApp_fail_("Action khong hop le: " + action);
    return map[action](params);
  } catch (err) {
    return vtdApp_fail_(String(err && err.stack || err));
  }
}

function vtdApp_dynamicConfig(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  const config = vtdApp_dynamicConfig_();
  return vtdApp_ok_({config: vtdApp_dynamicConfigForAuth_(config, auth)});
}

function vtdApp_saveDynamicConfig(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const config = params.config || {};
  const sh = vtdApp_dynamicConfigSheet_();
  sh.clearContents();
  sh.appendRow(["Key", "Json", "UpdatedAt", "UpdatedBy"]);
  sh.appendRow([
    "active",
    JSON.stringify(config),
    Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss"),
    vtdApp_auth_(params).email || ""
  ]);
  CacheService.getScriptCache().remove("VTD_DYNAMIC_CONFIG");
  return vtdApp_ok_({message: "Đã lưu cấu hình mở rộng.", config: config});
}

function vtdApp_dynamicListSheets(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const ss = vtdApp_openDynamicSpreadsheet_(params.spreadsheetId || params.url);
  const sheets = ss.getSheets().map(sh => ({
    name: sh.getName(),
    rows: sh.getLastRow(),
    columns: sh.getLastColumn(),
    hidden: sh.isSheetHidden()
  }));
  return vtdApp_ok_({spreadsheetId: ss.getId(), name: ss.getName(), sheets: sheets});
}

function vtdApp_dynamicReadColumns(params) {
  const denied = vtdApp_requireAction_("adminConfig", params);
  if (denied) return denied;
  params = params || {};
  const ss = vtdApp_openDynamicSpreadsheet_(params.spreadsheetId || params.url);
  const sheetName = String(params.sheetName || "").trim();
  const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sh) return vtdApp_fail_("Không thấy tab sheet cần đọc.");
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return vtdApp_ok_({spreadsheetId: ss.getId(), sheetName: sh.getName(), columns: [], sample: []});
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const sampleRows = Math.max(0, Math.min(sh.getLastRow() - 1, 3));
  const sample = sampleRows ? sh.getRange(2, 1, sampleRows, lastCol).getDisplayValues() : [];
  const columns = headers.map((name, index) => vtdApp_dynamicColumnFromHeader_(name || ("Column " + (index + 1)), index, sample.map(r => r[index])));
  return vtdApp_ok_({spreadsheetId: ss.getId(), sheetName: sh.getName(), columns: columns, sample: sample});
}

function vtdApp_dynamicReadRows(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  params = params || {};
  const config = vtdApp_dynamicConfigForAuth_(vtdApp_dynamicConfig_(), auth);
  const view = vtdApp_dynamicFindView_(config, params.viewId);
  if (!view) return vtdApp_fail_("Không thấy view mở rộng hoặc chưa được phân quyền.");
  const table = vtdApp_dynamicFindTable_(config, view.tableId);
  if (!table) return vtdApp_fail_("View chưa gắn table.");
  const ss = vtdApp_openDynamicSpreadsheet_(table.spreadsheetId || table.url);
  const sh = ss.getSheetByName(table.sheetName);
  if (!sh || sh.getLastRow() <= 1) return vtdApp_ok_({rows: [], columns: table.columns || [], count: 0});
  const lastCol = sh.getLastColumn();
  const maxRows = Math.min(Number(params.limit) || 200, 1000);
  const readRows = Math.min(sh.getLastRow() - 1, maxRows);
  const start = Math.max(2, sh.getLastRow() - readRows + 1);
  const values = sh.getRange(start, 1, readRows, lastCol).getDisplayValues();
  const columns = table.columns || [];
  const query = vtdApp_norm_(params.query || "");
  const out = [];
  values.forEach((row, offset) => {
    const obj = {__rowNumber: start + offset};
    columns.forEach((col, index) => obj[col.key || col.name] = row[index] || "");
    const searchText = columns.filter(c => c.search !== false).map(c => obj[c.key || c.name]).join(" ");
    if (query && vtdApp_norm_(searchText).indexOf(query) < 0) return;
    out.push(obj);
  });
  out.reverse();
  return vtdApp_ok_({rows: out, columns: columns, count: out.length, view: view, table: {id: table.id, name: table.name, sheetName: table.sheetName}});
}

function vtdApp_dynamicSaveRow(params) {
  const auth = vtdApp_auth_(params);
  if (!auth.allowed) return vtdApp_fail_(auth.message);
  params = params || {};
  const config = vtdApp_dynamicConfigForAuth_(vtdApp_dynamicConfig_(), auth);
  const view = vtdApp_dynamicFindView_(config, params.viewId);
  if (!view || String(view.type || "").toLowerCase() !== "form") return vtdApp_fail_("Không thấy form mở rộng hoặc chưa được phân quyền.");
  const table = vtdApp_dynamicFindTable_(config, view.tableId);
  if (!table) return vtdApp_fail_("Form chưa gắn table.");
  const ss = vtdApp_openDynamicSpreadsheet_(table.spreadsheetId || table.url);
  const sh = ss.getSheetByName(table.sheetName);
  if (!sh) return vtdApp_fail_("Không thấy tab sheet cần ghi.");
  const columns = table.columns || [];
  const data = params.data || {};
  const row = new Array(sh.getLastColumn()).fill("");
  columns.forEach((col, index) => {
    const key = col.key || col.name;
    if (col.editable === false) return;
    row[index] = data[key] != null ? data[key] : (col.initialValue || "");
  });
  sh.appendRow(row);
  return vtdApp_ok_({message: "Đã lưu dữ liệu mở rộng.", rowNumber: sh.getLastRow()});
}

function vtdApp_dynamicConfig_() {
  const cached = CacheService.getScriptCache().get("VTD_DYNAMIC_CONFIG");
  if (cached) return JSON.parse(cached);
  const sh = vtdApp_dynamicConfigSheet_();
  const values = sh.getDataRange().getValues();
  let config = {tables: [], views: [], version: 1};
  if (values.length > 1 && values[1][1]) {
    try { config = JSON.parse(values[1][1]); } catch (err) {}
  }
  if (!Array.isArray(config.tables)) config.tables = [];
  if (!Array.isArray(config.views)) config.views = [];
  CacheService.getScriptCache().put("VTD_DYNAMIC_CONFIG", JSON.stringify(config), 120);
  return config;
}

function vtdApp_dynamicConfigSheet_() {
  const ss = vtdApp_ss_();
  let sh = ss.getSheetByName(VTD_APP.dynamicConfigSheet);
  if (!sh) {
    sh = ss.insertSheet(VTD_APP.dynamicConfigSheet);
    sh.appendRow(["Key", "Json", "UpdatedAt", "UpdatedBy"]);
    sh.appendRow(["active", JSON.stringify({tables: [], views: [], version: 1}), Utilities.formatDate(new Date(), "Asia/Saigon", "yyyy-MM-dd HH:mm:ss"), "system"]);
  }
  return sh;
}

function vtdApp_dynamicConfigForAuth_(config, auth) {
  config = config || {tables: [], views: []};
  const email = String(auth && auth.email || "").toLowerCase().trim();
  const isAdmin = String(auth && auth.role || "").toLowerCase() === "admin";
  if (isAdmin) return config;
  const allowed = {};
  (auth.screens || []).forEach(id => allowed[String(id)] = true);
  const views = (config.views || []).filter(v => v && v.enabled !== false && (allowed["dyn_" + v.id] || allowed[v.id]));
  const tableIds = {};
  views.forEach(v => tableIds[v.tableId] = true);
  return Object.assign({}, config, {
    views: views,
    tables: (config.tables || []).filter(t => tableIds[t.id])
  });
}

function vtdApp_dynamicFindView_(config, id) {
  id = String(id || "").replace(/^dyn_/, "");
  return (config.views || []).find(v => String(v.id) === id && v.enabled !== false);
}

function vtdApp_dynamicFindTable_(config, id) {
  return (config.tables || []).find(t => String(t.id) === String(id));
}

function vtdApp_openDynamicSpreadsheet_(value) {
  const raw = String(value || "").trim();
  const idMatch = raw.match(/[-\w]{25,}/);
  const id = idMatch ? idMatch[0] : raw;
  if (!id) return vtdApp_ss_();
  return SpreadsheetApp.openById(id);
}

function vtdApp_dynamicColumnFromHeader_(name, index, sample) {
  const norm = vtdApp_norm_(name);
  let type = "Text";
  if (/date|ngay|ngày/.test(norm)) type = "Date";
  else if (/time|gio|giờ/.test(norm)) type = "Time";
  else if (/so luong|qty|number|amount|sl/.test(norm)) type = "Number";
  else if (/anh|image|photo|link/.test(norm)) type = "Image/URL";
  else if (/status|active|done|yes|no|tinh trang|trang thai/.test(norm)) type = "Enum";
  const key = norm.replace(/[^\w]+/g, "_") || ("col_" + (index + 1));
  return {
    name: String(name || "").trim(),
    key: key,
    type: type,
    keyColumn: index === 0 || /(^id$|ma don|mã đơn|raw id)/.test(norm),
    label: /ten|tên|ma don|mã đơn|name/.test(norm),
    formula: "",
    show: true,
    editable: true,
    required: index === 0,
    initialValue: "",
    displayName: String(name || "").trim(),
    description: "",
    search: true,
    scan: /qr|barcode|ma don|mã đơn|serial/.test(norm),
    nfc: false,
    pii: /email|phone|sdt|điện thoại|dia chi|địa chỉ/.test(norm),
    sample: (sample || []).filter(Boolean).slice(0, 3)
  };
}

function vtdApp_storeInfo(params) {
  const denied = vtdApp_requireAction_("storeInfo", params);
  if (denied) return denied;
  params = params || {};

  const monthFilter = vtdApp_monthDisplay_(params.month || params.thang || "");
  const typeFilter = vtdApp_norm_(params.loaiSt || params.type || "");
  const keyword = vtdApp_norm_(params.query || "");
  const cacheKey = "VTD_STOREINFO_TONGHOP_V6_" + monthFilter + "_" + typeFilter + "_" + keyword;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (cached) return vtdApp_ok_(JSON.parse(cached));

  const ss = SpreadsheetApp.openById(VTD_APP.fullCtSpreadsheetId);
  const sh = ss.getSheetByName("TongHop");
  if (!sh || sh.getLastRow() <= 1) return vtdApp_ok_({source: "TongHop", byStore: [], total: 0, months: []});

  const values = sh.getRange(1, 1, sh.getLastRow(), Math.min(8, sh.getLastColumn())).getValues();
  const col = vtdApp_headerMap_(values[0]);
  const byStore = {};
  const months = {};
  const monthTotals = {};
  const storeTypes = {};
  const dailyStores = {};
  const seenOd = {};
  const seenMonth = {};
  let total = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const od = String(vtdApp_pickFromArray_(row, col, ["od"]) || "").trim();
    const customer = String(vtdApp_pickFromArray_(row, col, ["khach hang", "khách hàng"]) || "").trim();
    const po = String(vtdApp_pickFromArray_(row, col, ["ma po", "mã po"]) || "").trim();
    const invoice = String(vtdApp_pickFromArray_(row, col, ["so hoa don", "số hóa đơn"]) || "").trim();
    const maDon = String(vtdApp_pickFromArray_(row, col, ["ma don", "mã đơn"]) || "").trim();
    const loaiStRaw = String(vtdApp_pickFromArray_(row, col, ["loai st", "loại st"]) || "").trim();
    const dateRaw = vtdApp_pickFromArray_(row, col, ["ngay len don", "ngày lên đơn", "date"]);
    const month = vtdApp_monthDisplay_(vtdApp_pickFromArray_(row, col, ["thang", "tháng", "month"])) || vtdApp_monthDisplay_(dateRaw);
    const storeType = loaiStRaw || vtdApp_storeType_(customer);
    const text = vtdApp_norm_([od, customer, po, invoice, maDon, storeType, month].join(" "));
    const uniqKey = od || maDon || String(i);

    if (month) months[month] = true;
    if (storeType) storeTypes[storeType] = true;
    const monthKey = (month || "Khac") + "::" + uniqKey;
    if (!seenMonth[monthKey]) {
      seenMonth[monthKey] = true;
      monthTotals[month || "Khac"] = (monthTotals[month || "Khac"] || 0) + 1;
    }
    if (monthFilter && month !== monthFilter) continue;
    if (typeFilter && vtdApp_norm_(storeType) !== typeFilter) continue;
    if (keyword && text.indexOf(keyword) < 0) continue;
    if (seenOd[uniqKey]) continue;
    seenOd[uniqKey] = true;
    byStore[storeType || "Khac"] = (byStore[storeType || "Khac"] || 0) + 1;
    const dateKey = vtdApp_dateKey_(dateRaw);
    if (dateKey) {
      if (!dailyStores[dateKey]) dailyStores[dateKey] = {};
      dailyStores[dateKey][storeType || "Khac"] = (dailyStores[dateKey][storeType || "Khac"] || 0) + 1;
    }
    total++;
  }

  const result = {
    source: "Full das chung tu VTD / TongHop",
    total: total,
    months: Object.keys(months).sort(vtdApp_monthSort_),
    monthTotals: vtdApp_toKeyRows_(monthTotals).sort((a, b) => vtdApp_monthSort_(a.key, b.key)),
    storeTypes: Object.keys(storeTypes).sort(),
    byStore: vtdApp_toKeyRows_(byStore).sort((a, b) => b.value - a.value),
    dailyStores: Object.keys(dailyStores).sort().map(date => ({
      date: date,
      total: Object.keys(dailyStores[date]).reduce((sum, key) => sum + Number(dailyStores[date][key] || 0), 0),
      stores: vtdApp_toKeyRows_(dailyStores[date]).sort((a, b) => b.value - a.value)
    }))
  };
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 300);
  return vtdApp_ok_(result);
}

function vtdApp_monthDisplay_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/yyyy");
  }
  const text = String(value || "").trim();
  if (!text) return "";
  let m = text.match(/^(\d{1,2})\/(\d{4})$/);
  if (m) return String(Number(m[1])).padStart(2, "0") + "/" + m[2];
  m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+.*)?$/);
  if (m) return String(Number(m[2])).padStart(2, "0") + "/" + m[3];
  m = text.match(/^T?(\d{1,2})\/(\d{2,4})$/i);
  if (m) return String(Number(m[1])).padStart(2, "0") + "/" + (String(m[2]).length === 2 ? "20" + m[2] : m[2]);
  m = text.match(/^(\d{4})-(\d{1,2})/);
  if (m) return String(Number(m[2])).padStart(2, "0") + "/" + m[1];
  return text;
}

function vtdApp_monthSort_(a, b) {
  const aa = String(a || "").split("/").map(Number);
  const bb = String(b || "").split("/").map(Number);
  const av = (aa[1] || 0) * 100 + (aa[0] || 0);
  const bv = (bb[1] || 0) * 100 + (bb[0] || 0);
  return av - bv;
}

function rebuildIndexNow() {
  const exists = PropertiesService.getScriptProperties().getProperty("VTD_SEARCH_CACHE_BUILD");
  if (exists) {
    return vtdApp_rebuildSearchCache({
      internalRun: true,
      continueBuild: true,
      reset: false,
      batchSize: 700
    });
  }
  return vtdApp_rebuildSearchCache({
    internalRun: true,
    clear: true,
    batchSize: 700
  });
}

function resetRebuildIndexNow() {
  PropertiesService.getScriptProperties().deleteProperty("VTD_SEARCH_CACHE_BUILD");
  return vtdApp_rebuildSearchCache({
    internalRun: true,
    clear: true,
    batchSize: 700
  });
}

function continueRebuildIndexNow() {
  return vtdApp_rebuildSearchCache({
    internalRun: true,
    continueBuild: true,
    reset: false,
    batchSize: 700
  });
}

function repairCacheFontNow() {
  const ss = vtdApp_ss_();
  const sh = ss.getSheetByName(VTD_APP.recordCacheSheet);
  if (!sh || sh.getLastRow() <= 1) return vtdApp_ok_({message: "Không có cache để sửa.", fixed: 0});
  const range = sh.getRange(2, 19, sh.getLastRow() - 1, 1);
  const values = range.getValues();
  let fixed = 0;
  const next = values.map(row => {
    const oldValue = row[0];
    const newValue = vtdApp_cleanStatusLabel_(oldValue) || oldValue;
    if (String(newValue) !== String(oldValue)) fixed++;
    return [newValue];
  });
  if (fixed) range.setValues(next);
  return vtdApp_ok_({message: "Đã sửa font SyncLabel trong cache.", fixed: fixed});
}
