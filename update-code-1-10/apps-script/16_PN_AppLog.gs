/* Audit of PN saves. The business request journal remains the source for retries. */
const PN_APP_LOG_HEADERS_ = [
  'Mã đơn GHTK', 'Số đơn hàng', 'Nguồn ghi', 'Người thao tác', 'Status', 'Detail',
  'Vị trí ghi', 'Cập nhật cuối', 'Update 1', 'Update 2', 'Update 3', 'Update 4',
  'Update 5', 'Update 6', 'Update 7', 'Update 8', 'Update 9', 'Log', 'Log JSON'
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

function pnAppLogTime_(at) {
  return Utilities.formatDate(new Date(at), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
}

function pnAppLogSummary_(data) {
  const operation = data.operations[data.operations.length - 1];
  const attempt = operation && operation.attempts[operation.attempts.length - 1];
  const updates = Array.from({length: 9}, () => []);
  const writes = new Map();
  const timeline = [];
  let status = '', detail = '', updatedAt = '';
  const add = (events, label) => events.forEach(event => timeline.push({event, label}));
  data.operations.forEach((op, i) => {
    const label = 'APP_SAVE #' + (i + 1);
    add(op.appEvents || [], label);
    op.attempts.forEach((entry, j) => add(entry.events, label + '.' + (j + 1)));
  });
  add(data.sheetEvents || [], 'SHEET');
  timeline.sort((a, b) => a.event.at.localeCompare(b.event.at));
  const lines = timeline.map(({event, label}) => {
    updatedAt = event.at;
    const time = pnAppLogTime_(event.at);
    if (event.update >= 1 && event.update <= 9) updates[event.update - 1].push(time + ' ' + event.state);
    const flow = /^MIRROR_/.test(event.state) ? 'SYNC_FF' :
      /^PRODUCT_SYNC_|^VHFF_SCHEDULED$/.test(event.state) ? 'SYNC_PRODUCT_VHFF' :
      /^SYNC_SV_/.test(event.state) ? 'SYNC_SV_VHFF' : label;
    const w = event.write;
    const context = [w && w.tab + '!' + pnAppLogRanges_(w.rows),
      event.editor && 'editor=' + event.editor,
      event.sourceTab && 'source=' + event.sourceTab + (event.sourceRow ? '!A' + event.sourceRow : ''),
      event.sourceCell && 'cell=' + event.sourceCell,
      event.recordKey && 'record=' + event.recordKey, event.target && 'target=' + event.target,
      event.files != null && 'photos=' + event.files, event.code, event.message].filter(Boolean);
    return [time, flow, event.state, ...context].join(' | ');
  });
  if (data.archivedIn) lines.unshift('Lịch sử trước đó: ' + data.archivedIn);
  if (attempt) {
    operation.attempts.forEach(entry => entry.events.forEach(event => {
      if (event.write) {
        const w = event.write;
        writes.set(w.tab, (writes.get(w.tab) || []).concat(w.rows));
      }
    }));
  }
  if (operation) {
    (operation.appEvents || []).forEach(event => {
      if (event.status) status = event.status;
      if (event.code && event.code !== 'OK') detail = event.code + ' · origin=APP';
    });
  }
  if (attempt) {
    status = 'PROCESSING';
    detail = '';
    attempt.events.forEach(event => {
      if (event.status) status = event.status;
      if (event.code) detail = [event.code, event.stage && 'stage=' + event.stage,
        event.message && 'message=' + event.message].filter(Boolean).join(' · ');
    });
    if (status === 'FAILED') {
      const confirmed = operation.attempts.some(entry => entry.events.some(event =>
        ['SHEET_WRITTEN', 'SHEET_WRITE_REUSED'].includes(event.state)));
      if (confirmed || writes.size) detail += confirmed ? ' · write=CONFIRMED' : ' · write=PARTIAL_OR_UNCONFIRMED';
    }
    const warnings = attempt.events.filter(event => /^(MIRROR|PRODUCT_SYNC)_(PENDING|FAILED)$/.test(event.state));
    if (warnings.length) detail = [detail, ...warnings.map(event =>
      event.state + (event.message ? ' · message=' + event.message : ''))].filter(Boolean).join(' · ');
  }
  const position = [...writes].map(([tab, rows]) => tab + '!' + pnAppLogRanges_(rows)).join(' · ');
  return [data.maDonGhtk, (operation && operation.orderNo) || data.soDonHang || '',
    operation ? operation.source : '', operation ? operation.user : '', status, detail, position,
    updatedAt ? new Date(updatedAt) : '', ...updates.map(history => history.join('\n')),
    lines.join('\n'), JSON.stringify(data)];
}

function pnAppLogSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (name === 'PN_LOG_APP') {
      sheet.getRange('A:B').setNumberFormat('@');
      sheet.getRange('H:H').setNumberFormat('dd/MM/yyyy HH:mm:ss');
      sheet.getRange('F:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
  }
  if (name === 'PN_LOG_APP' && sheet.getRange(1, 19).getValue() !== 'Log JSON') {
    sheet.getRange(1, 19).setValue('Log JSON').setFontWeight('bold');
    sheet.getRange('I:R').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
    sheet.hideColumns(19);
  }
  return sheet;
}

function pnAppLogReadRow_(sheet, row) {
  const [visible, raw] = sheet.getRange(row, 18, 1, 2).getValues()[0];
  return JSON.parse(raw || visible);
}

function pnAppLogWriteRow_(sheet, row, data) {
  let values = pnAppLogSummary_(data);
  if (values.slice(8).some(value => value.length > 42000)) {
    const archive = pnAppLogSheet_(pnSS_(), 'PN_LOG_HISTORY', ['Mã đơn GHTK', 'ArchivedAt', 'Log']);
    const overflow = {schemaVersion: 1, maDonGhtk: data.maDonGhtk, operations: [], sheetEvents: []};
    while (values.slice(8).some(value => value.length > 42000)) {
      if (data.operations.length > 1) overflow.operations.push(data.operations.shift());
      else if ((data.sheetEvents || []).length > 1) overflow.sheetEvents.push(data.sheetEvents.shift());
      else if (data.operations[0] && data.operations[0].attempts.length > 1) {
        const op = data.operations[0];
        const archived = overflow.operations.find(x => x.clientId === op.clientId) ||
          (overflow.operations[overflow.operations.push(Object.assign({}, op, {attempts: [], appEvents: []})) - 1]);
        archived.attempts.push(op.attempts.shift());
      } else {
        // A single attempt can exceed one cell with large SKU batches or device history.
        const op = data.operations[0];
        const entry = op && op.attempts[0];
        const source = entry && entry.events.length > 1 ? entry.events : op && op.appEvents;
        if (!source || source.length <= 1) throw new Error('PN log history exceeds cell capacity');
        const event = source.shift();
        archive.appendRow([data.maDonGhtk, new Date(), JSON.stringify({
          clientId: op.clientId, requestId: entry && entry.requestId, event
        })]);
      }
      data.archivedIn = 'PN_LOG_HISTORY';
      values = pnAppLogSummary_(data);
    }
    if (overflow.operations.length || overflow.sheetEvents.length) {
      const chunks = overflow.operations.flatMap(op => op.attempts.map(attempt =>
        ({clientId: op.clientId, orderNo: op.orderNo, source: op.source, user: op.user, input: op.input, attempt})));
      overflow.operations.forEach(op => (op.appEvents || []).forEach(event => chunks.push({clientId: op.clientId, event})));
      overflow.sheetEvents.forEach(event => chunks.push({event}));
      chunks.forEach(chunk => archive.appendRow([data.maDonGhtk, new Date(), JSON.stringify(chunk)]));
    }
  }
  const end = sheet.getMaxRows();
  if (row > end) sheet.insertRowsAfter(end, 100);
  sheet.getRange(row, 1, 1, values.length).setValues([values.map(v =>
    typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v)]);
}

// Chạy tay từng lô để dựng Update/Log cho dữ liệu từ phiên bản cũ.
function pnRebuildAppLogHistory() {
  const sheet = pnSS_().getSheetByName('PN_LOG_APP');
  if (!sheet || sheet.getLastRow() < 2) return {updated: 0, complete: true};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return {updated: 0, busy: true};
  try {
    pnAppLogSheet_(pnSS_(), 'PN_LOG_APP', PN_APP_LOG_HEADERS_);
    const props = PropertiesService.getScriptProperties(), key = 'PN_APP_LOG_REBUILD_ROW';
    let row = Number(props.getProperty(key)) || 2;
    const end = sheet.getLastRow();
    let updated = 0;
    for (; row <= end && updated < 100; row++) {
      const [visible, raw] = sheet.getRange(row, 18, 1, 2).getValues()[0];
      if (!visible || raw) continue;
      pnAppLogWriteRow_(sheet, row, JSON.parse(visible));
      updated++;
    }
    SpreadsheetApp.flush();
    if (row > end) props.deleteProperty(key);
    else props.setProperty(key, String(row));
    return {updated, complete: row > end};
  } finally { lock.releaseLock(); }
}

function pnAppLogSheetEvents_(source, events) {
  const previous = pnAppLogContext_;
  pnAppLogContext_ = {key: pnAppLogOrderKey_(source.maDonGhtk), orderNo: source.orderNo,
    requestId: Utilities.getUuid(), events, sheetOnly: true};
  try { pnAppLogFlush_(); }
  finally { pnAppLogContext_ = previous; }
}

function pnAppLogFlush_() {
  const log = pnAppLogContext_;
  if (!log || !log.key || !log.events.length) return;
  return pnAppLogSafe_(() => {
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
      const data = match ? pnAppLogReadRow_(sheet, row) :
        {schemaVersion: 1, maDonGhtk: log.key, operations: []};
      if (log.orderNo) data.soDonHang = log.orderNo;
      if (log.sheetOnly) {
        data.sheetEvents = (data.sheetEvents || []).concat(log.events);
        pnAppLogWriteRow_(sheet, row, data);
        SpreadsheetApp.flush();
        return true;
      }
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
      pnAppLogWriteRow_(sheet, row, data);
      SpreadsheetApp.flush();
      return true;
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
  if (!pnAppLogContext_) return;
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
