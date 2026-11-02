/*************************************************
DON DONG TRONG O CUOI TAB — moi ngay luc 00h, file rieng.

Nhan vien bam "Them 1000 hang" -> Google tu sao dropdown/checkbox xuong cac dong moi.
Moi dem: giu lai DUNG 5 dong trong ngay sau dong cuoi co du lieu, xoa het phan thua (ke ca dropdown, checkbox chua tich).
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
    .filter(t => t.getHandlerFunction() === 'pnDonDongTrong')
    .forEach(t => ScriptApp.deleteTrigger(t));
  // Google chay trong khung 00h-01h, khong dung phut 00:00.
  ScriptApp.newTrigger('pnDonDongTrong').timeBased().atHour(0).everyDays(1).inTimezone('Asia/Ho_Chi_Minh').create();
  return 'Da cai: pnDonDongTrong moi ngay luc 00h-01h.';
}

function pnDonDongTrong() {
  const result = pnWithLock_(() => {
    pnRequireReady_();
    return PN_DON_DONG.tabs.map(pnDonDongTab_);
  }, 120000);
  console.log('Don dong trong: ' + JSON.stringify(result));
  return result;
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
