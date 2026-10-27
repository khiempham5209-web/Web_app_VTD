/*************************************************
OCR PHIEU "TONG BANH CUA GIAO NHAN" -> TAB "TT Nhập"
OCR bang Google Drive (mien phi, khong can API key)

Ghi 5 cot doc tu anh: Ordered item | Quy cách/thùng | Tổng nhập | Thùng nguyên | Cái lẻ
Tu dien them: Ngày = hom nay, Barcode + Tên bánh = tra tu DS SKU.
Ma moi chua co trong DS SKU -> them vao cuoi DS SKU (Ma + Quy cach), khong sua dong da co.

2 cach dua anh vao:
  Cach 3. Menu "📷 OCR > Nhập từ ảnh" trong Google Sheet -> sidebar xem truoc, sua, roi ghi
  Cach 2. Google Form co cau hoi "Tải tệp lên" -> nop form la tu OCR va ghi luon

CAI DAT (1 lan):
  1. Google Sheet > Tien ich mo rong > Apps Script > them file moi, paste file nay.
  2. (Neu dung Form) dien ID form vao OCR_NHAP.formId.
  3. Chay ham ocrNhap_setup() 1 lan, cap quyen.
*************************************************/

const OCR_NHAP = {
  spreadsheetId: '1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco',
  sheetName: 'TT Nhập',
  headerRow: 2,
  headers: {
    item: 'Ordered item',
    quyCach: 'Quy cách/thùng',
    tong: 'Tổng nhập',
    thung: 'Thùng nguyên',
    le: 'Cái lẻ'
  },
  skuSheet: 'DS SKU',
  skuHeaderRow: 1,
  ocrLanguage: 'vi',
  formId: '',
  logSheet: 'OCR_LOG',
  notifyEmail: true,
  timezone: 'Asia/Saigon'
};

const OCR_NHAP_FIELDS = ['item', 'quyCach', 'tong', 'thung', 'le'];

/* ================= SETUP ================= */

function ocrNhap_setup() {
  const handlers = ['ocrNhap_addMenu', 'ocrNhap_onFormSubmit'];
  ScriptApp.getProjectTriggers()
    .filter(t => handlers.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Dung trigger onOpen cai dat de khong dung ham onOpen() cua project khac
  ScriptApp.newTrigger('ocrNhap_addMenu').forSpreadsheet(OCR_NHAP.spreadsheetId).onOpen().create();
  if (OCR_NHAP.formId) {
    ScriptApp.newTrigger('ocrNhap_onFormSubmit').forForm(OCR_NHAP.formId).onFormSubmit().create();
  }
  // Goi thu cac dich vu de xin quyen ngay luc setup
  DriveApp.getRootFolder();
  if (OCR_NHAP.notifyEmail) MailApp.getRemainingDailyQuota();
  Logger.log('OK. Triggers: ' + ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()).join(', '));
}

function ocrNhap_addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('📷 OCR')
    .addItem('Nhập từ ảnh', 'ocrNhap_openSidebar')
    .addToUi();
}

/* ================= CACH 3: SIDEBAR ================= */

function ocrNhap_openSidebar() {
  const html = HtmlService.createHtmlOutput(ocrNhap_sidebarHtml_()).setTitle('Nhập từ ảnh');
  SpreadsheetApp.getUi().showSidebar(html);
}

/*
  Goi tu sidebar. mode:
  - "cells": sidebar da cat tung O so (Ma, QC, HD, KHD, Tong, Thung, Le) cua moi dong,
             xep thanh anh "NHAN: gia tri"; o "-" / trong da ve san so 0.  (chinh xac nhat)
  - "rows" : khong tach duoc cot -> moi anh la 1 dong bang.
  - con lai: doc ca anh.
*/
function ocrNhap_preview(strips, wholeImage, mode) {
  let result;
  if (strips && strips.length >= 3) {
    const blobs = strips.map((b, i) => Utilities.newBlob(Utilities.base64Decode(b), 'image/png', 'row' + i + '.png'));
    const texts = ocrNhap_ocrMany_(blobs);
    const qcMap = ocrNhap_qcHistory_();
    if (mode === 'cells') {
      result = ocrNhap_check_(ocrNhap_parseCells_(texts, qcMap));
      result.rawText = '[Chế độ: đọc theo từng ô]\n' +
        texts.map(t => t.replace(/\s*\n\s*/g, ' | ').trim()).filter(Boolean).join('\n');
    } else {
      const lines = texts.map(t => t.replace(/\s+/g, ' ').trim()).filter(Boolean);
      result = ocrNhap_check_(ocrNhap_parseLines_(lines, qcMap));
      result.rawText = '[Chế độ: đọc theo dòng - không tách được cột]\n' + lines.join('\n');
    }
  } else {
    const blob = Utilities.newBlob(Utilities.base64Decode(wholeImage), 'image/png', 'ocr.png');
    result = ocrNhap_extract_(blob);
  }
  result.duplicate = ocrNhap_isDuplicate_(result.rows);
  return result;
}

// Goi tu sidebar: ghi cac dong da duoc nguoi dung xac nhan/sua.
// ngay = "yyyy-mm-dd" tu o chon ngay (de trong = hom nay).
function ocrNhap_commit(rows, ngay) {
  const m = String(ngay || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const dateText = m ? m[3] + '/' + m[2] + '/' + m[1] : '';
  const checked = ocrNhap_check_({items: rows.map(r => ({
    ordered_item: r.item, quy_cach: r.quyCach, tong: r.tong, thung: r.thung, le: r.le
  }))});
  const written = ocrNhap_write_(checked.rows, dateText);
  ocrNhap_markDone_(checked.rows);
  ocrNhap_log_('Sidebar', '', checked, 'OK', 'Ghi từ dòng ' + written.startRow + ', ngày ' + written.date);
  return written;
}

/* ================= CACH 2: GOOGLE FORM ================= */

function ocrNhap_onFormSubmit(e) {
  const fileIds = [];
  e.response.getItemResponses().forEach(ir => {
    if (ir.getItem().getType() === FormApp.ItemType.FILE_UPLOAD) {
      [].concat(ir.getResponse() || []).forEach(id => fileIds.push(id));
    }
  });

  const messages = [];
  fileIds.forEach(id => {
    let link = '';
    try {
      const file = DriveApp.getFileById(id);
      link = file.getUrl();
      const result = ocrNhap_extract_(file.getBlob());

      if (ocrNhap_isDuplicate_(result.rows)) {
        ocrNhap_log_('Form', link, result, 'BỎ QUA', 'Ảnh này đã được ghi trước đó');
        messages.push('⏭ Bỏ qua (trùng ảnh đã ghi): ' + link);
        return;
      }
      // Form khong co buoc xem truoc -> chi ghi khi moi dong va Grand Total deu khop
      if (!result.ok) {
        ocrNhap_log_('Form', link, result, 'KHÔNG GHI', result.warnings.join(' | '));
        messages.push('⚠️ Không ghi vì đọc chưa khớp. Dùng menu 📷 OCR > Nhập từ ảnh để nhập.\n' +
          result.warnings.join('\n') + '\n' + link);
        return;
      }
      const written = ocrNhap_write_(result.rows);
      ocrNhap_markDone_(result.rows);
      const note = 'Ghi ' + written.count + ' dòng từ dòng ' + written.startRow;
      ocrNhap_log_('Form', link, result, 'OK', note);
      messages.push('✅ ' + note + '\nTổng: ' + result.sum +
        (result.grandTotal != null ? ' / Grand Total ảnh: ' + result.grandTotal : ''));
    } catch (err) {
      ocrNhap_log_('Form', link, null, 'LỖI', String(err && err.message || err));
      messages.push('❌ Lỗi: ' + (err && err.message || err) + (link ? '\n' + link : ''));
    }
  });

  if (OCR_NHAP.notifyEmail && messages.length) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '[OCR TT Nhập] ' + (messages.every(m => m.indexOf('✅') === 0) ? 'Đã ghi' : 'Cần kiểm tra'),
      messages.join('\n\n') + '\n\nSheet: https://docs.google.com/spreadsheets/d/' + OCR_NHAP.spreadsheetId);
  }
}

/* ================= OCR (GOOGLE DRIVE) ================= */

function ocrNhap_extract_(blob) {
  const text = ocrNhap_ocrMany_([blob])[0];
  const result = ocrNhap_check_(ocrNhap_parseText_(text, ocrNhap_qcHistory_()));
  result.rawText = text;
  return result;
}

/*
  Drive OCR nhieu anh cung luc (UrlFetchApp.fetchAll):
  upload anh -> Google Docs (OCR) -> export text -> bo file tam vao thung rac.
*/
function ocrNhap_ocrMany_(blobs) {
  const auth = {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()};
  const api = 'https://www.googleapis.com/drive/v3/files';
  const boundary = 'ocrnhap' + Date.now();

  const uploadReq = (blob, i) => {
    const meta = JSON.stringify({name: 'ocr_tmp_' + Date.now() + '_' + i, mimeType: 'application/vnd.google-apps.document'});
    const head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
      '\r\n--' + boundary + '\r\nContent-Type: ' + blob.getContentType() + '\r\n\r\n').getBytes();
    const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
    return {
      url: 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id&ocrLanguage=' + OCR_NHAP.ocrLanguage,
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      payload: head.concat(blob.getBytes()).concat(tail),
      headers: auth,
      muteHttpExceptions: true
    };
  };

  // Chay song song, dong nao loi (vd. Drive gioi han toc do) thi thu lai tung cai
  const runAll = reqs => {
    const res = UrlFetchApp.fetchAll(reqs);
    return res.map((r, i) => {
      for (let k = 0; k < 3 && r.getResponseCode() >= 300; k++) {
        Utilities.sleep(1000 * (k + 1));
        r = UrlFetchApp.fetch(reqs[i].url, reqs[i]);
      }
      if (r.getResponseCode() >= 300) throw new Error('Drive OCR lỗi ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
      return r;
    });
  };

  const ids = runAll(blobs.map(uploadReq)).map(r => JSON.parse(r.getContentText()).id);
  try {
    return runAll(ids.map(id => ({url: api + '/' + id + '/export?mimeType=text/plain', headers: auth, muteHttpExceptions: true})))
      .map(r => r.getContentText().replace(/^﻿/, ''));
  } finally {
    UrlFetchApp.fetchAll(ids.map(id => ({
      url: api + '/' + id, method: 'patch', contentType: 'application/json',
      payload: JSON.stringify({trashed: true}), headers: auth, muteHttpExceptions: true
    })));
  }
}

/* ---------- Che do "cells": moi anh = 1 dong bang, moi dong chu = "NHAN: gia tri" ---------- */

// DATE, HAN = 2 cot "Date moi nhat co/khong co HD" (co luc chua so luong, cong vao Tong)
const OCR_NHAP_LABELS = ['MA', 'QC', 'DATE', 'HAN', 'HD', 'KHD', 'TONG', 'THUNG', 'LE'];

function ocrNhap_parseCells_(texts, qcMap) {
  qcMap = qcMap || {};
  const items = [];
  const skipped = [];
  let grandTotal = null;
  texts.forEach(text => {
    const f = ocrNhap_cellFields_(text);
    const n = k => ocrNhap_cellNumber_(f[k]);
    // Ma: 8 chu so trong o MA; khong thay thi tim trong ca text cua dong
    let code = String(f.MA || '').replace(/\s/g, '').match(/\d{8}/);
    if (!code) code = String(text || '').match(/(^|\D)(\d{8})(?!\d)/);
    code = code ? code[code.length - 1] : '';
    // Ma bi doc nham 1-2 ky tu (vd. "5506O332", "55O6l331")
    if (!code) {
      String(f.MA || '').split(/\s+/).forEach(tok => {
        if (!code && (tok.match(/\d/g) || []).length >= 6) {
          const fixed = tok.replace(/[Oo]/g, '0').replace(/[lI|!]/g, '1').replace(/\D/g, '');
          if (fixed.length === 8) code = fixed;
        }
      });
    }
    if (!code) {
      if (/grand|total/i.test(f.MA || text || '') && n('TONG') != null) grandTotal = n('TONG');
      else if (/\d{4,}/.test(f.MA || '')) skipped.push(String(f.MA).trim());
      return;
    }
    let tong = n('TONG'), thung = n('THUNG'), le = n('LE');
    const parts = ['DATE', 'HAN', 'HD', 'KHD'].map(n);
    // QC: uu tien so tren anh neu khop Thung x QC + Le = Tong; khong thi dung QC da luu
    const imgQc = n('QC'), histQc = qcMap[code];
    const fitsQc = q => q > 0 && thung != null && le != null && tong && thung * q + le === tong;
    const qc = fitsQc(imgQc) ? imgQc : fitsQc(histQc) ? histQc : (histQc || imgQc || 0);

    // Doi chieu 3 nguon cua Tong: o Tong, tong 4 cot so luong truoc no, Thung x QC + Le.
    // 2 nguon khop thi tin; o nao doc hong thi suy tu 2 nguon con lai.
    const byParts = qc && thung != null && le != null ? thung * qc + le : null;
    const bySum = parts.some(v => v == null) ? null : parts.reduce((a, b) => a + b, 0);
    if (tong && tong === byParts) {
      // khop
    } else if (bySum && bySum === byParts) {
      tong = bySum;
    } else if (qc && tong && tong === bySum) {
      thung = Math.floor(tong / qc);
      le = tong % qc;
    } else if (!tong && (bySum || byParts)) {
      tong = bySum || byParts;
      if (qc) { thung = Math.floor(tong / qc); le = tong % qc; }
    }
    items.push({ordered_item: code, quy_cach: qc, tong: tong || 0, thung: thung || 0, le: le || 0,
      newQc: !histQc, histQc: histQc});
  });
  return {items: items, grand_total: grandTotal, skipped: skipped};
}

// Tach "NHAN: gia tri". Nhan doc sai 1 ky tu van nhan ra; neu mat nhan thi theo thu tu dong.
function ocrNhap_cellFields_(text) {
  const out = {};
  let pending = null, next = 0;
  String(text || '').split(/\r?\n/).map(s => s.replace(/^[\s_\-–—|.·~=]+/, '').trim()).filter(Boolean).forEach(line => {
    const m = line.match(/^([A-Za-zÀ-ỹĐđ]{1,6})\s*[:;.,]?\s*(.*)$/);
    const label = m ? ocrNhap_closestLabel_(m[1]) : null;
    if (label) {
      pending = null;
      if (m[2].trim()) {
        out[label] = m[2].trim();
        next = OCR_NHAP_LABELS.indexOf(label) + 1;
      } else {
        pending = label;
      }
      return;
    }
    const target = pending || OCR_NHAP_LABELS[next];
    if (target && out[target] == null) {
      // Nhan doc sai (vd. "OC: 30") -> bo phan nhan, chi lay gia tri sau dau ":"
      out[target] = m && /[:;]/.test(line) ? m[2].trim() : line;
      next = OCR_NHAP_LABELS.indexOf(target) + 1;
    }
    pending = null;
  });
  return out;
}

function ocrNhap_closestLabel_(word) {
  const w = ocrNhap_norm_(word).toUpperCase().replace(/[^A-Z]/g, '');
  if (!w) return null;
  if (OCR_NHAP_LABELS.indexOf(w) >= 0) return w;
  if (w.length < 3) return null;
  const dist = (a, b) => {
    const d = [];
    for (let i = 0; i <= a.length; i++) d.push([i]);
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    return d[a.length][b.length];
  };
  const hits = OCR_NHAP_LABELS.filter(l => l.length >= 3 && dist(w, l) <= 1);
  return hits.length === 1 ? hits[0] : null;
}

// So dau tien trong o; o chi co dau gach -> 0; khong doc duoc -> null
function ocrNhap_cellNumber_(value) {
  if (value == null) return null;
  const v = String(value).trim();
  if (/^[-–—_.·~]*$/.test(v)) return 0;
  const m = ocrNhap_digitFix_(v).replace(/(\d)[.,](\d{3})/g, '$1$2').match(/\d+/);
  return m ? Number(m[0]) : null;
}

// O chi chua so -> doi cac ky tu OCR hay nham ve chu so (vd. "l7", "I7", "IT" -> 17)
function ocrNhap_digitFix_(value) {
  const map = {l: '1', I: '1', i: '1', '|': '1', '!': '1', j: '1', J: '1', O: '0', o: '0', D: '0', Q: '0',
    T: '7', S: '5', s: '5', B: '8', Z: '2', z: '2', G: '6', b: '6', g: '9', q: '9', A: '4'};
  return String(value == null ? '' : value).replace(/[lIij|!JOoDQTSsBZzGbgqA]/g, ch => map[ch]);
}

/* ---------- Che do "rows": moi dong text = 1 dong cua bang ---------- */

// Moi dong text = 1 dong cua bang (da cat san) -> ma + day so cuoi dong
function ocrNhap_parseLines_(lines, qcMap) {
  qcMap = qcMap || {};
  const items = [];
  let grandTotal = null;
  lines.forEach(line => {
    const clean = line.replace(/[–—−]/g, '-');
    if (/grand\s*t[o0]tal/i.test(clean)) {
      const nums = ocrNhap_tokens_(clean).filter(t => t !== null);
      if (nums.length) grandTotal = Math.max.apply(null, nums);
      return;
    }
    const m = clean.match(/(^|\D)(\d{8})(?!\d)/);
    if (!m) return;
    const r = ocrNhap_solveRow_(ocrNhap_tailNumbers_(clean.slice(m.index + m[0].length)), qcMap[m[2]]);
    items.push({ordered_item: m[2], quy_cach: r.qc, tong: r.tong, thung: r.thung, le: r.le, newQc: !qcMap[m[2]], histQc: qcMap[m[2]]});
  });
  return {items: items, grand_total: grandTotal};
}

// Chay tay de xem Drive OCR ra text gi (dan ID file anh tren Drive vao)
function ocrNhap_debug(fileId) {
  const result = ocrNhap_extract_(DriveApp.getFileById(fileId).getBlob());
  Logger.log('--- TEXT OCR ---\n' + result.rawText);
  Logger.log('--- KET QUA ---\n' + JSON.stringify(result.rows, null, 1));
  Logger.log('Tong: ' + result.sum + ' | Grand Total: ' + result.grandTotal + ' | ' + result.warnings.join(' | '));
}

/*
  Tach text theo ma Ordered item (8 chu so). Moi doan = 1 dong san pham.
  Cuoi moi dong co day so: QC | Date HĐ | Date không HĐ | HĐ | Không HĐ | Tổng | Thùng | Lẻ
  ("-" = 0). Ten banh co the chua so (90g, 250...) nen chi lay day so o cuoi
  va dung phep kiem tra Thùng×QC+Lẻ=Tổng, HĐ+Không HĐ=Tổng de chon dung cot.
*/
function ocrNhap_parseText_(text, qcMap) {
  qcMap = qcMap || {};
  const clean = String(text || '').replace(/[–—−]/g, '-').replace(/\u00a0/g, ' ');
  const gt = clean.match(/grand\s*t[o0]tal/i);
  const body = gt ? clean.slice(0, gt.index) : clean;

  const hits = [];
  const codeRe = /(^|\D)(\d{8})(?!\d)/g;
  let m;
  while ((m = codeRe.exec(body))) {
    hits.push({code: m[2], start: m.index + m[1].length, end: m.index + m[0].length});
  }

  const items = hits.map((h, i) => {
    const seg = body.slice(h.end, i + 1 < hits.length ? hits[i + 1].start : body.length);
    const r = ocrNhap_solveRow_(ocrNhap_tailNumbers_(seg), qcMap[h.code]);
    return {ordered_item: h.code, quy_cach: r.qc, tong: r.tong, thung: r.thung, le: r.le, newQc: !qcMap[h.code], histQc: qcMap[h.code]};
  });

  // Grand Total = so lon nhat trong dong Grand Total (cot Tổng luon lon nhat)
  let grandTotal = null;
  if (gt) {
    const nums = ocrNhap_tokens_(clean.slice(gt.index + gt[0].length, gt.index + gt[0].length + 200))
      .filter(t => t !== null);
    if (nums.length) grandTotal = Math.max.apply(null, nums);
  }
  return {items: items, grand_total: grandTotal};
}

// Tach token: so -> number, "-" -> 0, chu -> null, ky tu ke bang -> bo qua
function ocrNhap_tokens_(seg) {
  return seg.split(/\s+/).map(tok => {
    const t = tok.replace(/^[|\[\]!:;.,]+|[|\[\]!:;.,]+$/g, '');
    if (!t) return undefined;
    if (/^[-·•.~_]+$/.test(t)) return 0;
    if (/^[Oo]$/.test(t)) return 0;
    if (/^[\dOo]{1,3}([.,]\d{3})*$|^[\dOo]{1,6}$/.test(t) && /\d/.test(t)) {
      return Number(t.replace(/[Oo]/g, '0').replace(/[.,]/g, ''));
    }
    // Chu rac ngan do OCR (vd. "сл", "I") -> bo qua, khong lam dut day so
    if (t.length <= 2 && !/^[A-Za-zÀ-ỹ]*\d/.test(t) && !/^(x|MB|MT|g)$/i.test(t)) return undefined;
    return null;
  }).filter(t => t !== undefined);
}

// Day so lien tiep o cuoi doan (sau chu cuoi cung cua ten banh)
function ocrNhap_tailNumbers_(seg) {
  const tokens = ocrNhap_tokens_(seg);
  let i = tokens.length;
  while (i > 0 && tokens[i - 1] !== null) i--;
  return tokens.slice(i);
}

/*
  t = day so cuoi dong: QC | Date HĐ | Date không HĐ | HĐ | Không HĐ | Tổng | Thùng | Lẻ
  OCR hay: mat dau "-", doc "-" thanh 1, doc "20|" thanh 201, them "-" thua o cuoi.
  => Chon so nao la Tổng: khi biet QC thi Thùng = Tổng/QC, Lẻ = Tổng%QC phai hien ngay sau Tổng
     (Thùng/Lẻ = 0 co the bi mat), va HĐ + Không HĐ ngay truoc Tổng = Tổng.
  histQc = quy cach cua ma nay trong lich su TT Nhập (tin hon so OCR doc duoc).
*/
function ocrNhap_solveRow_(t, histQc) {
  // Tim Thùng roi Lẻ theo thu tu trong day so sau Tổng (cho phep so thua/lap xen giua,
  // so 0 = dau "-" co the bi mat). Tra ve null neu thieu so khac 0.
  const matchFrom = (after, need, from) => {
    if (!need.length) return 0;
    let best = null;
    const pos = after.indexOf(need[0], from);
    if (pos >= 0) {
      const rest = matchFrom(after, need.slice(1), pos + 1);
      if (rest !== null) best = rest + 1;
    }
    if (need[0] === 0) {
      const rest = matchFrom(after, need.slice(1), from);
      if (rest !== null && (best === null || rest > best)) best = rest;
    }
    return best;
  };
  const matchAfter = (after, need) => {
    const hit = matchFrom(after, need, 0);
    return hit === null ? null : {hit: hit, extra: after.length - hit};
  };

  let best = null;
  for (let j = t.length - 1; j >= 0; j--) {
    const tong = t[j];
    const after = t.slice(j + 1);
    if (after.length > 4) break;
    if (!(tong > 0)) continue;
    // QC: thu ca so tren anh lan QC da luu (DS SKU / lich su). QC da luu chi duoc cong diem,
    // khong ep buoc -> neu DS SKU luu sai (vd. 12) thi so tren anh (48) van thang.
    const qcs = [];
    if (histQc) qcs.push({qc: histQc, idx: -1});
    t.slice(0, j).forEach((v, i) => {
      if (v > 0 && v <= 200) qcs.push({qc: v, idx: i});
      if (v >= 100 && v % 10 === 1) qcs.push({qc: (v - 1) / 10, idx: i}); // "20|" doc thanh 201
    });

    qcs.forEach(c => {
      const thung = Math.floor(tong / c.qc);
      const le = tong % c.qc;
      const m = matchAfter(after, [thung, le]);
      if (!m || !m.hit) return;
      let score = 2 * m.hit - m.extra;

      const mid = t.slice(c.idx + 1, j);
      if (mid.length >= 2 && mid[mid.length - 2] + mid[mid.length - 1] === tong) score += 3;
      else if (mid.length && mid[mid.length - 1] === tong) score += 2;
      if (c.qc > 1) score += 1;
      if (histQc && c.qc === histQc) score += 2;
      if (!best || score > best.score) best = {score: score, qc: c.qc, tong: tong, thung: thung, le: le};
    });
  }
  if (best) return best;

  // Khong giai duoc: lay tam theo vi tri, dong nay se bi to vang
  const s = t.length >= 8 ? t.slice(-8) : [];
  return {qc: histQc || s[0] || 0, tong: s[5] || 0, thung: s[6] || 0, le: s[7] || 0};
}

// Quy cach da biet cua tung ma: uu tien DS SKU (chuan), thieu thi lay dong moi nhat trong TT Nhập
function ocrNhap_qcHistory_() {
  const sh = SpreadsheetApp.openById(OCR_NHAP.spreadsheetId).getSheetByName(OCR_NHAP.sheetName);
  const map = {};
  if (sh) {
    const col = ocrNhap_columns_(sh);
    const first = OCR_NHAP.headerRow + 1;
    const n = sh.getLastRow() - first + 1;
    if (n > 0) {
      const items = sh.getRange(first, col.item, n, 1).getValues();
      const qcs = sh.getRange(first, col.quyCach, n, 1).getValues();
      items.forEach((r, i) => {
        const code = String(r[0]).replace(/\D/g, '');
        const qc = Number(qcs[i][0]);
        if (code && qc > 0) map[code] = qc;
      });
    }
  }
  const sku = ocrNhap_skuMap_();
  Object.keys(sku).forEach(code => { if (sku[code].qc > 0) map[code] = sku[code].qc; });
  return map;
}

// Doc DS SKU (CHI DOC, khong bao gio ghi): ma -> {barcode, name, qc}
function ocrNhap_skuMap_() {
  const sh = SpreadsheetApp.openById(OCR_NHAP.spreadsheetId).getSheetByName(OCR_NHAP.skuSheet);
  const map = {};
  if (!sh || sh.getLastRow() <= OCR_NHAP.skuHeaderRow) return map;
  const values = sh.getRange(OCR_NHAP.skuHeaderRow, 1, sh.getLastRow() - OCR_NHAP.skuHeaderRow + 1, sh.getLastColumn()).getDisplayValues();
  const headers = values.shift().map(ocrNhap_norm_);
  const find = names => {
    for (let i = 0; i < names.length; i++) {
      const idx = headers.indexOf(ocrNhap_norm_(names[i]));
      if (idx >= 0) return idx;
    }
    return -1;
  };
  const cCode = find(['Mã Vật Tư', 'Ordered item']);
  const cBarcode = find(['Barcode']);
  const cName = find(['Tên Vật Tư', 'Tên bánh']);
  const cQc = find(['Quy cách', 'Quy cách/thùng']);
  if (cCode < 0) return map;
  values.forEach(r => {
    const code = String(r[cCode]).replace(/\D/g, '');
    if (!code) return;
    map[code] = {
      barcode: cBarcode >= 0 ? String(r[cBarcode]).trim() : '',
      name: cName >= 0 ? String(r[cName]).trim() : '',
      qc: cQc >= 0 ? Number(String(r[cQc]).replace(/[^\d]/g, '')) || 0 : 0
    };
  });
  return map;
}

// Them ma CHUA CO vao cuoi DS SKU (STT, Ma, Quy cach). Khong sua dong nao da co.
function ocrNhap_addNewSku_(rows) {
  const seen = {};
  rows = rows.filter(r => r.item && !seen[r.item] && (seen[r.item] = true));
  if (!rows.length) return [];
  const sh = SpreadsheetApp.openById(OCR_NHAP.spreadsheetId).getSheetByName(OCR_NHAP.skuSheet);
  if (!sh) return [];
  const lastCol = Math.max(sh.getLastColumn(), 5);
  const headers = sh.getRange(OCR_NHAP.skuHeaderRow, 1, 1, lastCol).getValues()[0].map(ocrNhap_norm_);
  const at = names => {
    for (let i = 0; i < names.length; i++) {
      const idx = headers.indexOf(ocrNhap_norm_(names[i]));
      if (idx >= 0) return idx;
    }
    return -1;
  };
  const cCode = at(['Mã Vật Tư', 'Ordered item']);
  const cQc = at(['Quy cách', 'Quy cách/thùng']);
  const cStt = at(['STT']);
  if (cCode < 0) return [];

  const first = OCR_NHAP.skuHeaderRow + 1;
  const lastRow = sh.getLastRow();
  let stt = lastRow >= first
    ? sh.getRange(first, cCode + 1, lastRow - first + 1, 1).getDisplayValues().filter(r => String(r[0]).trim()).length
    : 0;
  const out = rows.map(r => {
    const line = new Array(lastCol).fill('');
    line[cCode] = r.item;
    if (cQc >= 0) line[cQc] = r.quyCach;
    if (cStt >= 0) line[cStt] = ++stt;
    return line;
  });
  const start = Math.max(lastRow + 1, first);
  sh.getRange(start, 1, out.length, lastCol).setValues(out);
  sh.getRange(start, 1, out.length, 5)
    .setFontFamily('Times New Roman').setFontSize(12).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  return rows.map(r => r.item);
}

function ocrNhap_norm_(s) {
  return String(s == null ? '' : s).trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ');
}

/* ================= KIEM TRA + GHI ================= */

// Chuan hoa + kiem tra: thung*quyCach + le = tong, va tong cong = Grand Total
function ocrNhap_check_(raw) {
  const num = v => {
    const n = Number(String(v == null ? '' : v).replace(/[^\d-]/g, ''));
    return isFinite(n) ? n : 0;
  };
  const rows = (raw.items || []).map(it => {
    const r = {
      item: String(it.ordered_item || '').replace(/\D/g, ''),
      quyCach: num(it.quy_cach),
      tong: num(it.tong),
      thung: num(it.thung),
      le: num(it.le),
      warn: ''
    };
    if (!r.item) r.warn = 'Thiếu mã';
    else if (!r.tong || !r.quyCach) r.warn = 'Không đọc được số';
    else if (r.thung * r.quyCach + r.le !== r.tong) {
      r.warn = r.thung + '×' + r.quyCach + '+' + r.le + '≠' + r.tong;
    } else if (it.newQc) r.warn = 'Mã mới, kiểm tra Quy cách';
    else if (it.histQc && it.histQc !== r.quyCach) r.warn = 'QC khác DS SKU: ảnh ' + r.quyCach + ', đã lưu ' + it.histQc;
    return r;
  }).filter(r => r.item || r.tong);

  const sum = rows.reduce((s, r) => s + r.tong, 0);
  const grandTotal = raw.grand_total == null ? null : num(raw.grand_total);
  // "Ma moi" chi la nhac nho, khong tinh la loi
  const warnings = rows.filter(r => r.warn && !/^(Mã mới|QC khác)/.test(r.warn)).map(r => r.item + ': ' + r.warn);
  if (grandTotal != null && grandTotal !== sum) warnings.push('Tổng ' + sum + ' ≠ Grand Total ' + grandTotal);
  if (!rows.length) warnings.push('Không đọc được dòng nào');
  if (raw.skipped && raw.skipped.length) {
    warnings.push(raw.skipped.length + ' dòng không đọc được mã (' + raw.skipped.join(', ') + ')');
  }

  return {rows: rows, sum: sum, grandTotal: grandTotal, warnings: warnings, ok: warnings.length === 0};
}

function ocrNhap_write_(rows, dateText) {
  if (!rows.length) throw new Error('Không có dòng nào để ghi');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(OCR_NHAP.spreadsheetId).getSheetByName(OCR_NHAP.sheetName);
    if (!sh) throw new Error('Không thấy tab ' + OCR_NHAP.sheetName);
    const col = ocrNhap_columns_(sh);
    const startRow = ocrNhap_nextRow_(sh, col.item);
    const n = rows.length;
    const lastCol = sh.getLastColumn();

    // Lay dinh dang cua dong ngay tren (vd. so 0 hien thanh "-")
    if (startRow > OCR_NHAP.headerRow + 1) {
      sh.getRange(startRow - 1, 1, 1, lastCol)
        .copyTo(sh.getRange(startRow, 1, n, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sh.getRange(startRow, col.item, n, 1).setBackground(null);
    }

    OCR_NHAP_FIELDS.forEach(f => {
      sh.getRange(startRow, col[f], n, 1).setValues(rows.map(r => [r[f]]));
    });

    // Tu dien Ngay / Barcode / Ten banh (script ghi thi onEdit khong chay, nen phai dien o day).
    // Barcode + Ten banh lay tu DS SKU, KHONG ghi gi vao DS SKU.
    const extra = ocrNhap_extraColumns_(sh);
    const sku = ocrNhap_skuMap_();
    const ngay = dateText || Utilities.formatDate(new Date(), OCR_NHAP.timezone, 'dd/MM/yyyy');
    if (extra.date) sh.getRange(startRow, extra.date, n, 1).setValues(rows.map(() => [ngay]));
    if (extra.barcode) sh.getRange(startRow, extra.barcode, n, 1).setValues(rows.map(r => [(sku[r.item] || {}).barcode || '']));
    if (extra.name) sh.getRange(startRow, extra.name, n, 1).setValues(rows.map(r => [(sku[r.item] || {}).name || '']));

    // Ma moi -> them sang DS SKU (Ma + Quy cach). Barcode/Ten ban bo sung o DS SKU,
    // code sync se tu dua ve TT Nhập.
    const newCodes = ocrNhap_addNewSku_(rows.filter(r => !sku[r.item]));

    // To vang + ghi chu: dong nghi doc sai, hoac ma moi
    rows.forEach((r, i) => {
      const notes = [];
      if (r.warn) notes.push('OCR: ' + r.warn);
      if (newCodes.indexOf(r.item) >= 0) notes.push('Mã mới, đã thêm vào DS SKU. Cần bổ sung Barcode/Tên ở DS SKU');
      if (notes.length) sh.getRange(startRow + i, col.item).setBackground('#fff2cc').setNote(notes.join('\n'));
    });
    SpreadsheetApp.flush();
    return {startRow: startRow, count: n, missingSku: newCodes, date: ngay};
  } finally {
    lock.releaseLock();
  }
}

function ocrNhap_columns_(sh) {
  const headers = sh.getRange(OCR_NHAP.headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(ocrNhap_norm_);
  const col = {};
  OCR_NHAP_FIELDS.forEach(f => {
    const idx = headers.indexOf(ocrNhap_norm_(OCR_NHAP.headers[f]));
    if (idx < 0) throw new Error('Không thấy cột "' + OCR_NHAP.headers[f] + '" ở dòng ' + OCR_NHAP.headerRow);
    col[f] = idx + 1;
  });
  return col;
}

// Cot Ngay / Barcode / Ten banh trong TT Nhập (khong co thi bo qua)
function ocrNhap_extraColumns_(sh) {
  const headers = sh.getRange(OCR_NHAP.headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(ocrNhap_norm_);
  const at = name => headers.indexOf(ocrNhap_norm_(name)) + 1;
  return {date: at('Ngày'), barcode: at('Barcode'), name: at('Tên bánh')};
}

// Dong trong dau tien sau dong cuoi co Ordered item (khong dung getLastRow vi cot dropdown keo dai)
function ocrNhap_nextRow_(sh, itemCol) {
  const first = OCR_NHAP.headerRow + 1;
  const last = sh.getLastRow();
  if (last < first) return first;
  const values = sh.getRange(first, itemCol, last - first + 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== '') return first + i + 1;
  }
  return first;
}

/* ================= CHONG GHI TRUNG ================= */

function ocrNhap_fingerprint_(rows) {
  const text = rows.map(r => r.item + ':' + r.tong).join('|');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text);
  return 'OCR_FP_' + bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function ocrNhap_isDuplicate_(rows) {
  return !!PropertiesService.getScriptProperties().getProperty(ocrNhap_fingerprint_(rows));
}

function ocrNhap_markDone_(rows) {
  PropertiesService.getScriptProperties().setProperty(ocrNhap_fingerprint_(rows), new Date().toISOString());
}

/* ================= LOG ================= */

function ocrNhap_log_(source, link, result, status, message) {
  try {
    const ss = SpreadsheetApp.openById(OCR_NHAP.spreadsheetId);
    let sh = ss.getSheetByName(OCR_NHAP.logSheet);
    if (!sh) {
      sh = ss.insertSheet(OCR_NHAP.logSheet);
      sh.appendRow(['Thời gian', 'Nguồn', 'Ảnh', 'Số dòng', 'Tổng', 'Grand Total ảnh', 'Trạng thái', 'Ghi chú', 'Text OCR']);
    }
    sh.appendRow([
      Utilities.formatDate(new Date(), OCR_NHAP.timezone, 'dd/MM/yyyy HH:mm:ss'),
      source, link,
      result ? result.rows.length : '',
      result ? result.sum : '',
      result && result.grandTotal != null ? result.grandTotal : '',
      status, message,
      result && result.rawText ? result.rawText.slice(0, 40000) : ''
    ]);
  } catch (e) {}
}

/* ================= SIDEBAR HTML ================= */

const OCR_NHAP_CSS = `
body{font:13px Arial,sans-serif;margin:10px;color:#222}
#drop{border:2px dashed #999;border-radius:8px;padding:16px;text-align:center;cursor:pointer;background:#fafafa}
#drop img{max-width:100%;margin-top:8px}
button{padding:8px 12px;margin:8px 4px 0 0;border:0;border-radius:4px;background:#1a73e8;color:#fff;cursor:pointer}
button:disabled{background:#aaa}
table{border-collapse:collapse;width:100%;margin-top:10px;font-size:12px}
th,td{border:1px solid #ddd;padding:2px}th{background:#f1f3f4}
td input{width:100%;box-sizing:border-box;border:0;text-align:right;font:inherit;background:transparent}
td input.code{text-align:left}tr.bad td{background:#fff2cc}tr.new td{background:#e8f0fe}
.x{color:#c00;cursor:pointer;text-align:center}
#msg{margin-top:8px;white-space:pre-line}.ok{color:#188038}.err{color:#c5221f}
details{margin-top:10px}pre{white-space:pre-wrap;font-size:11px;background:#f5f5f5;padding:6px;max-height:240px;overflow:auto}
#cells img{max-width:100%;border:1px solid #ddd;margin:2px 0}
.date{margin-top:10px}.date input{padding:4px;font:inherit}.date small{color:#666}
`;

// Chay tren trinh duyet (sidebar). Khong dung dau backtick trong doan nay.
const OCR_NHAP_CLIENT_JS = String.raw`
var img = null, rows = [], grand = null;
var F = ["item", "quyCach", "tong", "thung", "le"], H = ["Mã", "QC", "Tổng", "Thùng", "Lẻ"];
// Cot can doc: Ma = cot dau tien; cac cot so dem tu PHAI sang (cot Ten banh co the sinh duong doc gia)
var FIELDS = [["MA", 0], ["QC", -8], ["DATE", -7], ["HAN", -6], ["HD", -5], ["KHD", -4], ["TONG", -3], ["THUNG", -2], ["LE", -1]];

function $(id) { return document.getElementById(id); }
function msg(t, c) { $("msg").className = c || ""; $("msg").textContent = t; }

$("drop").onclick = function() { $("file").click(); };
$("file").onchange = function() { if (this.files[0]) load(this.files[0]); };
document.addEventListener("paste", function(e) {
  var it = e.clipboardData.items;
  for (var i = 0; i < it.length; i++) {
    if (it[i].type.indexOf("image") === 0) { load(it[i].getAsFile()); return; }
  }
});

function load(file) {
  var r = new FileReader();
  r.onload = function() {
    var im = new Image();
    im.onload = function() {
      // Khong phong to (lam nhoe duong ke); chi thu nho anh qua lon
      var s = Math.min(1, 2400 / Math.max(im.width, im.height)), c = document.createElement("canvas");
      c.width = Math.round(im.width * s); c.height = Math.round(im.height * s);
      var g = c.getContext("2d");
      g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height);
      g.imageSmoothingQuality = "high"; g.drawImage(im, 0, 0, c.width, c.height);
      prepare(c);
    };
    im.src = r.result;
  };
  r.readAsDataURL(file);
}

function prepare(c) {
  var url = c.toDataURL("image/png"), g = gray(c);
  var bands = findBands(g), cols = bands.length >= 3 ? findCols(g, bands) : [];
  var mode = "whole", strips = [];
  if (cols.length >= 9) {
    mode = "cells";
    strips = bands.map(function(b) { return cellsImage(c, g, b, cols); });
  } else if (bands.length >= 3) {
    mode = "rows";
    strips = bands.map(function(b) { return cropRow(c, b[0], b[1]); });
  }
  img = {whole: url.split(",")[1], strips: strips, mode: mode};
  $("preview").innerHTML = "<img src=\"" + url + "\">";
  $("cells").innerHTML = mode === "cells" ? strips.map(function(s) { return "<img src=\"data:image/png;base64," + s + "\">"; }).join("") : "";
  $("btnRead").disabled = false; $("btnSave").disabled = true;
  $("table").innerHTML = ""; $("rawBox").hidden = true;
  msg(mode === "cells" ? "Tách được " + bands.length + " dòng × " + cols.length + " cột, đọc theo từng ô."
    : mode === "rows" ? "Tách được " + bands.length + " dòng (không thấy đủ cột), đọc theo dòng."
    : "Không thấy đường kẻ bảng, sẽ đọc cả ảnh.");
}

function gray(c) {
  var w = c.width, h = c.height, d = c.getContext("2d").getImageData(0, 0, w, h).data, L = new Float32Array(w * h);
  for (var i = 0; i < w * h; i++) L[i] = .299 * d[i * 4] + .587 * d[i * 4 + 1] + .114 * d[i * 4 + 2];
  return {L: L, w: w, h: h};
}

// Pixel thuoc duong ke: toi hon pixel cach 2-5px o hai phia (chiu duoc duong mảnh/nhạt khi anh nho
// va duong day khi anh lon)
function isLine(g, x, y, dx, dy) {
  var v = g.L[y * g.w + x];
  for (var k = 2; k <= 5; k++) {
    var x1 = x - dx * k, y1 = y - dy * k, x2 = x + dx * k, y2 = y + dy * k;
    if (x1 < 0 || y1 < 0 || x2 >= g.w || y2 >= g.h) continue;
    if (v < g.L[y1 * g.w + x1] - 12 && v < g.L[y2 * g.w + x2] - 12) return true;
  }
  return false;
}

function runs(flags, minLen) {
  var out = [], start = 0;
  for (var i = 0; i <= flags.length; i++) {
    if (i === flags.length || flags[i]) { if (i - start >= minLen) out.push([start, i]); start = i + 1; }
  }
  return out;
}

function findBands(g) {
  var f = [];
  for (var y = 0; y < g.h; y++) {
    var n = 0;
    for (var x = 0; x < g.w; x++) if (isLine(g, x, y, 0, 1)) n++;
    f.push(n > g.w * .5);
  }
  var bands = runs(f, 6);
  // 2 dong dinh nhau (duong ke qua mo) -> dai cao gap doi binh thuong -> tach doi
  var hs = bands.map(function(b) { return b[1] - b[0]; }).sort(function(a, b) { return a - b; });
  var med = hs[Math.floor(hs.length / 2)] || 0, out = [];
  bands.forEach(function(b) {
    var h = b[1] - b[0];
    if (med && h > med * 1.7 && h < med * 2.3) {
      var mid = Math.round((b[0] + b[1]) / 2);
      out.push([b[0], mid]); out.push([mid, b[1]]);
    } else out.push(b);
  });
  return out.slice(0, 80);
}

// Duong ke doc, chi xet vung cac dong du lieu (bo dong tieu de cao)
function findCols(g, bands) {
  var hs = bands.map(function(b) { return b[1] - b[0]; }).sort(function(a, b) { return a - b; });
  var med = hs[Math.floor(hs.length / 2)];
  var data = bands.filter(function(b) { return b[1] - b[0] <= med * 1.5; });
  if (data.length < 3) return [];
  var y0 = data[0][0], y1 = data[data.length - 1][1], counts = [];
  for (var x = 0; x < g.w; x++) {
    var n = 0;
    for (var y = y0; y < y1; y++) if (isLine(g, x, y, 1, 0)) n++;
    counts.push(n / (y1 - y0));
  }
  // Duong ke that chay suot chieu cao (> 80%); net chu so can phai chi ~50% -> khong bi nham
  var tries = [.8, .65], cols = [];
  for (var t = 0; t < tries.length && cols.length < 9; t++) {
    cols = runs(counts.map(function(v) { return v > tries[t]; }), 6);
  }
  return cols;
}

// Khung bao phan muc (chu) trong o. Bo qua hang/cot gan nhu toan muc = duong ke bang
// (anh dan tu Zalo co the lon/nho hon, duong ke lot vao o).
function inkBox(g, x0, y0, x1, y1) {
  var w = x1 - x0, h = y1 - y0, rowN = [], colN = [], x, y;
  for (y = 0; y < h; y++) rowN.push(0);
  for (x = 0; x < w; x++) colN.push(0);
  for (y = y0; y < y1; y++) for (x = x0; x < x1; x++) {
    if (g.L[y * g.w + x] < 150) { rowN[y - y0]++; colN[x - x0]++; }
  }
  var a = 1e9, b = 1e9, c = -1, d = -1, n = 0;
  for (y = y0; y < y1; y++) {
    // Chi bo hang/cot SAT MEP o (duong ke bang lot vao); so "1" nam ben trong -> giu nguyen
    if (rowN[y - y0] > w * .85 && (y - y0 < 3 || y1 - y <= 3)) continue;
    for (x = x0; x < x1; x++) {
      if (colN[x - x0] > h * .85 && (x - x0 < 3 || x1 - x <= 3)) continue;
      if (g.L[y * g.w + x] < 150) { n++; if (x < a) a = x; if (x > c) c = x; if (y < b) b = y; if (y > d) d = y; }
    }
  }
  return n < 3 ? null : {x0: a, y0: b, x1: c + 1, y1: d + 1};
}

// 1 dong bang -> 1 anh: moi o 1 dong "NHAN: gia tri". O trong / dau "-" -> tu ve so 0.
function cellsImage(c, g, band, cols) {
  var h = band[1] - band[0], s = Math.max(1, Math.min(4, 60 / h)), pad = 3;
  var parts = FIELDS.map(function(f) {
    var col = f[1] >= 0 ? cols[f[1]] : cols[cols.length + f[1]];
    var box = inkBox(g, col[0] + pad, band[0] + pad, col[1] - pad, band[1] - pad);
    var zero = !box || (f[0] !== "MA" && box.y1 - box.y0 < h * 0.3);
    return {label: f[0], box: zero ? null : box};
  });
  var font = Math.max(24, Math.round(h * s * 0.55)), labelW = Math.round(font * 4.2);
  var lineH = Math.round(h * s) + 24, maxW = font * 2;
  parts.forEach(function(p) { if (p.box) maxW = Math.max(maxW, (p.box.x1 - p.box.x0 + 6) * s); });
  var o = document.createElement("canvas");
  o.width = Math.round(labelW + maxW + 60); o.height = parts.length * lineH + 40;
  var ctx = o.getContext("2d");
  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, o.width, o.height);
  ctx.fillStyle = "#000"; ctx.font = "bold " + font + "px Arial"; ctx.textBaseline = "middle";
  ctx.imageSmoothingQuality = "high";
  parts.forEach(function(p, i) {
    var mid = 20 + i * lineH + lineH / 2;
    ctx.textAlign = "right"; ctx.fillText(p.label + ":", labelW - font * 0.4, mid);
    ctx.textAlign = "left";
    if (!p.box) { ctx.fillText("0", labelW, mid); return; }
    var bw = p.box.x1 - p.box.x0 + 6, bh = p.box.y1 - p.box.y0 + 6;
    ctx.drawImage(c, p.box.x0 - 3, p.box.y0 - 3, bw, bh, labelW, mid - bh * s / 2, bw * s, bh * s);
  });
  return o.toDataURL("image/png").split(",")[1];
}

// Du phong khi khong tach duoc cot: cat ca dong
function cropRow(c, y0, y1) {
  var hh = y1 - y0, s = Math.max(1, Math.min(4, 60 / hh)), pad = 20, o = document.createElement("canvas");
  o.width = Math.round(c.width * s); o.height = Math.round(hh * s) + pad * 2;
  var g = o.getContext("2d");
  g.fillStyle = "#fff"; g.fillRect(0, 0, o.width, o.height); g.imageSmoothingQuality = "high";
  g.drawImage(c, 0, y0, c.width, hh, 0, pad, o.width, Math.round(hh * s));
  return o.toDataURL("image/png").split(",")[1];
}

$("btnRead").onclick = function() {
  $("btnRead").disabled = true; msg("Đang đọc ảnh (khoảng 10–30 giây)...");
  google.script.run.withSuccessHandler(function(res) {
    rows = res.rows; grand = res.grandTotal; render();
    var skip = (res.warnings || []).filter(function(w) { return /không đọc được mã/.test(w); });
    if (skip.length) msg($("msg").textContent + "\n❌ " + skip.join("; ") + " — xem text OCR thô.", "err");
    $("raw").textContent = res.rawText || ""; $("rawBox").hidden = false;
    $("btnRead").disabled = false; $("btnSave").disabled = !rows.length;
    if (res.duplicate) msg("⚠️ Ảnh này có vẻ đã được ghi rồi. Kiểm tra trước khi ghi lại.", "err");
  }).withFailureHandler(function(e) {
    msg("❌ " + e.message, "err"); $("btnRead").disabled = false;
  }).ocrNhap_preview(img.strips, img.whole, img.mode);
};

function check() {
  var sum = 0, bad = 0;
  rows.forEach(function(r) {
    F.slice(1).forEach(function(f) { r[f] = Number(r[f]) || 0; });
    r.bad = !r.item || !r.tong || !r.quyCach || r.thung * r.quyCach + r.le !== r.tong;
    if (r.bad) bad++;
    sum += r.tong;
  });
  return {sum: sum, bad: bad};
}

function render() {
  var c = check(), h = "<table><tr>";
  H.forEach(function(t) { h += "<th>" + t + "</th>"; });
  h += "<th></th></tr>";
  rows.forEach(function(r, i) {
    var cls = r.bad ? "bad" : /^(Mã mới|QC khác)/.test(r.warn || "") ? "new" : "";
    h += "<tr title=\"" + (r.warn || "") + "\" class=\"" + cls + "\">";
    F.forEach(function(f) {
      h += "<td><input class=\"" + (f === "item" ? "code" : "") + "\" data-i=\"" + i + "\" data-f=\"" + f + "\" value=\"" + r[f] + "\"></td>";
    });
    h += "<td class=\"x\" data-del=\"" + i + "\">✕</td></tr>";
  });
  $("table").innerHTML = h + "</table>";
  var t = rows.length + " dòng | Tổng: " + c.sum + (grand != null ? " | Grand Total ảnh: " + grand : "");
  var nw = rows.filter(function(r) { return /^(Mã mới|QC khác)/.test(r.warn || ""); }).length;
  var ok = rows.length && !c.bad && (grand == null || grand === c.sum);
  msg(t + (nw ? "\nℹ️ " + nw + " dòng tô xanh: mã mới hoặc QC khác DS SKU (rê chuột để xem)." : "") +
    (ok ? "\n✅ Khớp" : "\n⚠️ " + (c.bad ? c.bad + " dòng lệch (tô vàng). " : "") +
      (grand != null && grand !== c.sum ? "Tổng lệch Grand Total." : "")), ok ? "ok" : "err");
}

$("table").addEventListener("change", function(e) {
  var t = e.target;
  if (t.dataset.f) { rows[t.dataset.i][t.dataset.f] = t.value.trim(); render(); }
});
$("table").addEventListener("click", function(e) {
  var d = e.target.dataset.del;
  if (d != null) { rows.splice(+d, 1); render(); }
});

$("btnSave").onclick = function() {
  var c = check();
  if ((c.bad || (grand != null && grand !== c.sum)) && !confirm("Vẫn còn dòng lệch. Vẫn ghi?")) return;
  var d = $("ngay").value;
  if (d && !confirm("Ghi với ngày " + d.split("-").reverse().join("/") + "?")) return;
  $("btnSave").disabled = true; msg("Đang ghi...");
  google.script.run.withSuccessHandler(function(w) {
    msg("✅ Đã ghi " + w.count + " dòng, từ dòng " + w.startRow + ", ngày " + w.date +
      (w.missingSku && w.missingSku.length ? "\n⚠️ Mã mới đã thêm vào DS SKU, cần bổ sung Barcode/Tên ở DS SKU: " + w.missingSku.join(", ") : ""), "ok");
    rows = []; $("table").innerHTML = "";
  }).withFailureHandler(function(e) {
    msg("❌ " + e.message, "err"); $("btnSave").disabled = false;
  }).ocrNhap_commit(rows.map(function(r) {
    return {item: String(r.item), quyCach: r.quyCach, tong: r.tong, thung: r.thung, le: r.le};
  }), $("ngay").value);
};
`;

function ocrNhap_sidebarHtml_() {
  return '<!DOCTYPE html><html><head><base target="_top"><style>' + OCR_NHAP_CSS + '</style></head><body>' +
    '<div id="drop">Bấm để chọn ảnh<br>hoặc <b>Ctrl+V</b> dán ảnh (copy từ Zalo)' +
    '<input type="file" id="file" accept="image/*" hidden><div id="preview"></div></div>' +
    '<div class="date">Ngày nhập: <input type="date" id="ngay"> <small>(để trống = hôm nay)</small></div>' +
    '<button id="btnRead" disabled>Đọc ảnh</button>' +
    '<button id="btnSave" disabled>Ghi vào TT Nhập</button>' +
    '<div id="msg"></div><div id="table"></div>' +
    '<details id="rawBox" hidden><summary>Xem text OCR thô</summary><pre id="raw"></pre>' +
    '<div id="cells"></div></details>' +
    '<script>' + OCR_NHAP_CLIENT_JS + '</script></body></html>';
}
