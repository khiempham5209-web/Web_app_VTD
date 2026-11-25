/* Audit of PN saves. The business request journal remains the source for retries. */
const PN_APP_LOG_HEADERS_ = [
  'Mã đơn GHTK', 'Số đơn hàng', 'Nguồn ghi', 'Người thao tác', 'Status', 'Detail',
  'Vị trí ghi', 'Cập nhật cuối', 'Update 1', 'Update 2', 'Update 3', 'Update 4',
  'Update 5', 'Update 6', 'Update 7', 'Update 8', 'Update 9', 'Log'
];
let pnAppLogContext_ = null;

function pnAppLogSafe_(fn) {
  try { return fn(); }
  catch (err) { console.error('PN_APP_LOG: ' + String(err.message || err)); }
}

function pnAppLogCode_(value) {
  const text = String(value || '').replace(/https?:\/\/\S+/gi, '[URL]')
    .replace(/(?:Bearer|sessionToken)\s*[:=]?\s*\S+/gi, '[REDACTED]')
    .replace(/\s+/g, ' ').trim();
  return text.slice(0, 350);
}

function pnAppLogOrderKey_(value) {
  const key = pnText_(value);
  return /^[a-z0-9][a-z0-9._-]{0,79}$/i.test(key) ? key.toUpperCase() : '';
}

function pnAppLogBegin_(params) {
  const rawType = params.returnType || params.loaiHoan || 'Chung tu';
  const types = {
    'chung tu': 'Chứng từ', docs: 'Chứng từ', 'san pham': 'Sản phẩm', product: 'Sản phẩm',
    'san pham + chung tu': 'Sản phẩm + Chứng từ', 'product-docs': 'Sản phẩm + Chứng từ'
  };
  const items = Array.isArray(params.items) ? params.items : [];
  const operationId = pnAppLogCode_(params.clientId) || Utilities.getUuid();
  pnAppLogContext_ = {
    key: pnAppLogOrderKey_(params.maDonGhtk), orderNo: pnText_(params.orderNo),
    source: types[norm_(rawType)] || 'UNKNOWN',
    user: pnAppLogCode_(params.userEmail || params.user),
    operationId, requestId: Utilities.getUuid(), startedAt: new Date().toISOString(),
    input: {
      skuCount: items.length,
      docPhotoCount: Array.isArray(params.files) ? params.files.length : 0,
      productPhotoCount: items.reduce((n, item) => {
        const files = item && (item.files || item.images);
        return n + (Array.isArray(files) ? files.length : 0);
      }, 0)
    },
    events: [], stage: 'API', writeCount: 0
  };
  pnAppLogEvent_('API_RECEIVED', 4, {status: 'PROCESSING'});
  pnAppLogFlush_();
}

function pnAppLogEvent_(state, update, data) {
  const log = pnAppLogContext_;
  if (!log) return;
  log.events.push(Object.assign({
    state, at: new Date().toISOString(), origin: 'API'
  }, update ? {update} : {}, data || {}));
}

function pnAppLogResolve_(found) {
  const log = pnAppLogContext_;
  if (!log || log.resolved) return;
  const record = found.record;
  log.resolved = true;
  log.key = log.key || pnAppLogOrderKey_(record.maDonGhtk);
  log.orderNo = pnText_(record.orderNo) || log.orderNo;
  pnAppLogEvent_('ORDER_RESOLVED', 5, {row: found.rowNumber, recordKey: record.syncKey || ''});
}

function pnAppLogWrite_(state, sheet, rows, ids) {
  const log = pnAppLogContext_;
  if (!log || !rows.length) return;
  log.writeCount += rows.length;
  pnAppLogEvent_(state, 0, {write: {tab: sheet.getName(), gid: sheet.getSheetId(), rows, ids}});
}

function pnAppLogProductRows_(sheet, col, ids, complete) {
  if (!pnAppLogContext_) return;
  pnAppLogSafe_(() => {
    const last = sheet.getLastRow();
    if (last < 2 || !ids.length) return;
    const escaped = ids.map(id => id.replace(/[.*+?^$${}()|[\]\\]/g, '\\$&'));
    const matches = sheet.getRange(2, col, last - 1, 1)
      .createTextFinder('^(?:' + escaped.join('|') + ')$')
      .useRegularExpression(true).matchCase(true).findAll();
    const rows = matches.map(cell => cell.getRow()).sort((a, b) => a - b);
    pnAppLogEvent_('PRODUCT_ROWS_CHECKED', 0, {expected: ids.length, confirmed: rows.length});
    pnAppLogWrite_(complete ? 'PRODUCTS_WRITTEN' : 'PRODUCT_ROWS_CONFIRMED', sheet, rows, ids);
  });
}

function pnAppLogFinish_(result, error) {
  const log = pnAppLogContext_;
  if (!log) return;
  const message = pnAppLogCode_(error && (error.message || error) || result && result.message);
  if (result && result.ok) {
    pnAppLogEvent_(log.cached ? 'CACHED_RESULT' : 'DONE', 8, {status: 'DONE'});
    return;
  }
  const busy = (result && result.code === 'PN_BUSY') || /PN_BUSY/.test(message);
  const notFound = /Không tìm thấy đơn|Khong tim thay dong/i.test(message);
  if (notFound) pnAppLogEvent_('ORDER_NOT_FOUND', 5);
  if (error && log.stage === 'UPLOAD') pnAppLogEvent_('UPLOAD_FAILED', 6);
  if (error && /^WRITE_/.test(log.stage)) {
    pnAppLogEvent_(log.writeCount ? 'PARTIAL_WRITE' : 'WRITE_UNCONFIRMED', 7);
  }
  const status = busy ? 'RETRY_PENDING' : error && !notFound ? 'FAILED' : 'REJECTED';
  pnAppLogEvent_(status, 8, {
    status, code: busy ? 'PN_BUSY' : notFound ? 'ORDER_NOT_FOUND' :
      result && result.code || (error ? log.stage + '_FAILED' : 'SAVE_REJECTED'),
    stage: log.stage, message
  });
}

function pnAppLogRanges_(rows) {
  const ranges = [];
  [...new Set(rows)].sort((a, b) => a - b).forEach(row => {
    const last = ranges[ranges.length - 1];
    if (last && row === last[1] + 1) last[1] = row;
    else ranges.push([row, row]);
  });
  return ranges.map(([a, b]) => a === b ? 'A' + a : 'A' + a + ':A' + b).join(', ');
}

function pnAppLogSummary_(data) {
  const operation = data.operations[data.operations.length - 1];
  const attempt = operation.attempts[operation.attempts.length - 1];
  const updates = Array(9).fill('');
  const writes = new Map();
  let status = 'PROCESSING', detail = '', updatedAt = '';
  operation.attempts.forEach(entry => entry.events.forEach(event => {
    if (event.write) {
      const w = event.write;
      const rows = writes.get(w.tab) || [];
      writes.set(w.tab, rows.concat(w.rows));
    }
  }));
  attempt.events.forEach(event => {
    updatedAt = event.at;
    if (event.update) updates[event.update - 1] = Utilities.formatDate(new Date(event.at), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss') + ' ' + event.state;
    if (event.status) status = event.status;
    if (event.code) {
      detail = [event.code, event.stage && 'stage=' + event.stage, event.message && 'message=' + event.message].filter(Boolean).join(' · ');
    }
  });
  if (status === 'FAILED' && writes.size) detail += ' · write=PARTIAL_OR_UNCONFIRMED';
  if (status === 'DONE') {
    const mirror = attempt.events.filter(event => /^MIRROR_/.test(event.state)).pop();
    if (mirror && /PENDING|FAILED/.test(mirror.state)) detail = mirror.state + ' · target=Chứng từ_FF';
  }
  const position = [...writes].map(([tab, rows]) => tab + '!' + pnAppLogRanges_(rows)).join(' · ');
  return [data.maDonGhtk, operation.orderNo || data.soDonHang || '', operation.source, operation.user,
    status, detail, position, new Date(updatedAt), ...updates, JSON.stringify(data)];
}

function pnAppLogSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (name === 'PN_LOG_APP') {
    sheet.getRange('A:B').setNumberFormat('@');
    sheet.getRange('H:H').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange('F:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.getRange('R:R').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  return sheet;
}

function pnAppLogFlush_() {
  const log = pnAppLogContext_;
  if (!log || !log.key || !log.events.length) return;
  pnAppLogSafe_(() => {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(500)) {
      console.warn('PN_APP_LOG: LOCK_BUSY ' + log.requestId);
      return;
    }
    try {
      const ss = pnSS_();
      const sheet = pnAppLogSheet_(ss, 'PN_LOG_APP', PN_APP_LOG_HEADERS_);
      const last = sheet.getLastRow();
      const match = last < 2 ? null : sheet.getRange(2, 1, last - 1, 1)
        .createTextFinder(log.key).matchCase(true).matchEntireCell(true).findNext();
      const row = match ? match.getRow() : last + 1;
      const data = match ? JSON.parse(sheet.getRange(row, 18).getValue()) :
        {schemaVersion: 1, maDonGhtk: log.key, operations: []};
      if (log.orderNo) data.soDonHang = log.orderNo;
      let operation = data.operations.find(x => x.clientId === log.operationId);
      if (!operation) {
        operation = {clientId: log.operationId, orderNo: log.orderNo, source: log.source,
          user: log.user, input: log.input, attempts: []};
        const localTime = /^LOCAL-(\d{13})-/.exec(log.operationId);
        const created = localTime && Number(localTime[1]);
        operation.startedAtMs = created && created <= Date.now() + 300000 ? created : Date.parse(log.startedAt);
        data.operations.push(operation);
        data.operations.sort((a, b) => a.startedAtMs - b.startedAtMs);
      }
      if (log.orderNo) operation.orderNo = log.orderNo;
      let attempt = operation.attempts.find(x => x.requestId === log.requestId);
      if (!attempt) {
        attempt = {requestId: log.requestId, startedAt: log.startedAt, events: []};
        operation.attempts.push(attempt);
      }
      attempt.events = log.events.slice();
      let values = pnAppLogSummary_(data);
      if (values[17].length > 42000) {
        const archive = pnAppLogSheet_(ss, 'PN_LOG_HISTORY', ['Mã đơn GHTK', 'ArchivedAt', 'Log']);
        const old = match ? sheet.getRange(row, 18).getValue() : '';
        if (old) archive.appendRow([log.key, new Date(), old]);
        data.archivedIn = 'PN_LOG_HISTORY';
        while (data.operations.length > 1 && JSON.stringify(data).length > 42000) data.operations.shift();
        const attempts = data.operations[0].attempts;
        while (attempts.length > 1 && JSON.stringify(data).length > 42000) attempts.shift();
        values = pnAppLogSummary_(data);
      }
      const end = sheet.getMaxRows();
      if (row > end) sheet.insertRowsAfter(end, 100);
      // Prevent user-supplied text being treated as a Sheets formula.
      sheet.getRange(row, 1, 1, values.length).setValues([values.map(v =>
        typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v)]);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  });
}

function pnAppLogStage_(stage) {
  const log = pnAppLogContext_;
  if (!log || log.stage === stage) return;
  log.stage = stage;
  pnAppLogEvent_(stage + '_STARTED', 0);
  // Checkpoint the long upload once. Later events are batched at request completion.
  if (stage === 'UPLOAD') pnAppLogFlush_();
}

function pnAppLogCached_(result) {
  if (!pnAppLogContext_ || !result) return;
  pnAppLogContext_.cached = true;
  if (result.record) pnAppLogResolve_({record: result.record, rowNumber: result.rowNumber});
}

function pnAppLogMirror_(fullRow, result) {
  if (!pnAppLogContext_ || !pnAppLogContext_.events.some(event => event.state === 'DOC_WRITTEN')) return;
  if (result && result.queued) {
    pnAppLogEvent_('MIRROR_PENDING', 0);
    return;
  }
  const pair = PN_FULL.pairs[2];
  const source = pnV2Layout_(pnSheet_(pair.full), true);
  const target = pnV2Layout_(pnSheet_(pair.main), false);
  if (!fullRow || source.id == null || target.id == null || target.sh.getLastRow() < 2) return;
  const id = pnText_(source.sh.getRange(fullRow, source.id + 1).getValue());
  if (!id) return;
  const cells = target.sh.getRange(2, target.id + 1, target.sh.getLastRow() - 1, 1)
    .createTextFinder(id).matchEntireCell(true).matchCase(true).findAll();
  if (cells.length === 1) pnAppLogWrite_('MIRROR_CONFIRMED', target.sh, [cells[0].getRow()], [id]);
}
