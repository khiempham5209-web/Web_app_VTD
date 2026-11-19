/*************************************************
DON DONG TRONG O CUOI TAB — moi ngay luc 05h sang, file rieng.

Nhan vien bam "Them 1000 hang" -> Google tu sao dropdown/checkbox xuong cac dong moi.
Moi sang: giu lai DUNG 5 dong trong ngay sau dong cuoi co du lieu, xoa het phan thua (ke ca dropdown, checkbox chua tich).
Dong co bat ky gia tri nao (o bat ky cot nao) deu khong bi xoa; chi dong o CUOI tab moi bi dung toi.

Cai dat: chay ham pnCaiTriggerDonDong() 1 lan, cap quyen.
Chay thu ngay: chay ham pnDonDongTrong(), xem ket qua o Nhat ky thuc thi.
*************************************************/

const PN_DON_DONG = {
  tabs: ['Chứng từ_FF'],
  blankRows: 5
};

function pnCaiTriggerDonDong() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['pnDonDongTrong', 'pnDon5h'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  // Google chay trong khung 05h-06h, khong dung phut 05:00. pnDon5h = don thang cu + don dong trong.
  ScriptApp.newTrigger('pnDon5h').timeBased().atHour(5).everyDays(1).inTimezone('Asia/Ho_Chi_Minh').create();
  return 'Da cai: pnDon5h moi ngay luc 05h-06h.';
}

function pnDonDongTrong() {
  const result = pnWithLock_(() => {
    pnRequireReady_();
    return PN_DON_DONG.tabs.map(pnDonDongTab_);
  }, 120000);
  console.log('Don dong trong: ' + JSON.stringify(result));
  return result;
}

/* 5h sáng mỗi ngày (cài bằng pnInstallTriggers): việc có XÓA DÒNG chạy lúc không ai làm, để không lệch dòng của ai.
 * Bật cờ PN_CLEANING: trong lúc dọn, lần sửa/app nào tới thì được lưu lại và tự chạy bù ngay sau khi dọn xong.
 * 1. Đồng bộ đủ + dọn tháng cũ khỏi tab chính (dòng đã khớp 100% với full).  2. Dọn nhật ký API cũ.
 * 3. Giữ 5 dòng trống cuối Chứng từ_FF. */
function pnDon5h() {
  const props = PropertiesService.getScriptProperties(), out = {};
  props.setProperty('PN_CLEANING', String(Date.now()));
  try {
    pnWithLock_(() => {
      pnRequireReady_();
      pnV2Start_ = Date.now();
      try {
        if (props.getProperty('PN_FULL_CLEANUP') === 'enabled') {
          const sync = pnV2SyncAll_({fullSweep: true, areaSweep: true});
          out.cleanup = sync.partial ? {ok: false, reason: 'Đồng bộ chưa xong trong thời gian cho phép; hoãn dọn.'} : pnV2Cleanup_();
        }
        try { out.journal = pnJournalPrune_(); } catch (err) { out.journal = String(err && err.message || err); }
        out.blank = PN_DON_DONG.tabs.map(pnDonDongTab_);
      } finally { pnV2Start_ = 0; }
    }, 120000);
  } finally {
    props.deleteProperty('PN_CLEANING');
    console.log('Don 5h: ' + JSON.stringify(out));
  }
  // Lần sửa đến trong lúc dọn: chạy bù ngay.
  if (Object.keys(props.getProperties()).some(k => k.indexOf('PN_EDITQ_') === 0) || props.getProperty('PN_JOB_NEED')) pnScheduleJob_([]);
  return out;
}

function pnDonDongTab_(name) {
  const sh = pnSheet_(name), maxRows = sh.getMaxRows(), maxCols = sh.getMaxColumns();
  const values = sh.getRange(1, 1, maxRows, maxCols).getValues();
  // Dong cuoi co gia tri o bat ky cot nao (checkbox chua tich = FALSE khong tinh).
  let lastData = 1;
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i].some(v => v !== '' && v !== false && v != null)) { lastData = i + 1; break; }
  }
  const keep = lastData + PN_DON_DONG.blankRows;
  let removed = 0, added = 0;
  if (maxRows > keep) { removed = maxRows - keep; sh.deleteRows(keep + 1, removed); }
  else if (maxRows < keep) { added = keep - maxRows; sh.insertRowsAfter(maxRows, added); }
  SpreadsheetApp.flush();
  return {tab: name, lastData, maxRowsBefore: maxRows, removed, added};
}
