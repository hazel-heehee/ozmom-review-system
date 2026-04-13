// ===================================================
// OzMom 체험단 관리 시스템 - Apps Script 최종본
// ===================================================
// 설정값 (반드시 입력)
// ===================================================

// 대시보드용 시트 (읽기 전용)
const DASHBOARD_SHEET_ID = '1fhdHU7Y5a_EDr8fATJ3OMyh1WqWFPGpCHSUCm7fxIRI';
const DASHBOARD_SHEET_NAME = '★자사들 의 체험 후기';

// 운영용 시트 (접수/선정/OMG/후기 누적)
// 처음 배포 후 아래 함수 initOperationSheet() 를 한 번 실행하면 자동 생성됩니다
let OPERATION_SHEET_ID = '1mTzbHDCqe8sDrilkCwEEInvu_oXyKwqvrJb6BMO51Pw';

// ===================================================
// 초기 설정: 운영 시트 최초 1회 생성
// Apps Script 편집기에서 이 함수를 한 번만 실행하세요!
// ===================================================
function initOperationSheet() {
  const ss = SpreadsheetApp.create('OzMom 체험단 운영 시트');
  const sheetNames = ['접수목록', '선정결과', 'OMG리스트', '추기제출'];
  ss.getSheets()[0].setName('접수목록');
  sheetNames.slice(1).forEach(name => ss.insertSheet(name));

  const headers = {
    '접수목록': ['타임스탬프','이름','연락처','채널','채널명','팔로워','제품','신청사유','상태','선정일','완료일','초안일','승인일','비고'],
    '선정결과': ['선정일','이름','연락처','채널','채널명','팔로워','제품','상태','완료일','초안일','승인일','비고'],
    'OMG리스트': ['이름','연락처','사유','날짜','비고'],
    '추기제출': ['제출일','이름','연락처','채널','URL','캡처','상태','비고']
  };

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length)
      .setBackground('#4a90d9').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  });

  OPERATION_SHEET_ID = ss.getId();
  Logger.log('운영 시트 생성 완료! ID: ' + OPERATION_SHEET_ID);
  return OPERATION_SHEET_ID;
}

// ===================================================
// 웹앱 진입점
// ===================================================
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  let result;
  try {
    switch(params.action) {
      case 'getDashboard':  result = getDashboard(params);  break;
      case 'getApplicants': result = getApplicants(params); break;
      case 'getSelected':   result = getSelected(params);   break;
      case 'getReviews':    result = getReviews(params);    break;
      case 'getOmgList':    result = getOmgList();          break;
      default: result = { success: false, error: '알 수 없는 요청' };
    }
  } catch(err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  let result;
  try {
    switch(data.action) {
      case 'createForm':   result = createForm(data);   break;
      case 'saveSelected': result = saveSelected(data); break;
      case 'saveOMG':      result = saveOMG(data);      break;
      case 'deleteOMG':    result = deleteOMG(data);    break;
      case 'submitReview': result = submitReview(data); break;
      case 'updateStatus': result = updateStatus(data); break;
      default: result = { success: false, error: '알 수 없는 요청' };
    }
  } catch(err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================================================
// 1. 대시보드 통계
// ===================================================
function getDashboard(params) {
  params = params || {};
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('접수목록');
  if (!sheet) return { success: false, error: '접수목록 시트를 찾을 수 없습니다' };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, stats: getEmptyStats(), calendar: {}, rows: [] };

  function fmt(val) {
    if (!val) return '-';
    try {
      const d = new Date(val);
      if (isNaN(d.getTime())) return '-';
      return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
    } catch(e) { return '-'; }
  }

  let rows = data.slice(1).map((row, i) => ({
    rowIndex:    i + 2,
    date:        fmt(row[0]),
    name:        String(row[1] || '').trim(),
    phone:       String(row[2] || '').trim(),
    channel:     String(row[3] || '').trim(),
    channelName: String(row[4] || '').trim(),
    followers:   row[5] || 0,
    product:     String(row[6] || '').trim(),
    reason:      String(row[7] || '').trim(),
    status:      String(row[8] || '').trim(),
    approveDate: fmt(row[9]),
    doneDate:    fmt(row[10]),
    draftDate:   fmt(row[11]),
    approveDate2:fmt(row[12]),
    note:        String(row[13] || '').trim(),
    raw:         row[0],
  })).filter(r => r.name || r.status);

  // 기간 필터
  if (params.from) rows = rows.filter(r => r.raw && new Date(r.raw) >= new Date(params.from));
  if (params.to)   rows = rows.filter(r => r.raw && new Date(r.raw) <= new Date(params.to + 'T23:59:59'));

  // 달력 집계
  const calMap = {};
  rows.forEach(r => {
    [['deadline',r.date],['doneDate',r.doneDate],['draftDate',r.draftDate],['approveDate',r.approveDate2]]
      .forEach(function(pair) {
        var type = pair[0], dateStr = pair[1];
        if (!dateStr || dateStr === '-') return;
        if (!calMap[dateStr]) calMap[dateStr] = { deadline:0, doneDate:0, draftDate:0, approveDate:0 };
        calMap[dateStr][type]++;
      });
  });

  const stats = {
    total:    rows.length,
    selected: rows.filter(r => ['선정','완료','초안','승인'].includes(r.status)).length,
    done:     rows.filter(r => r.status === '완료').length,
    draft:    rows.filter(r => r.status === '초안').length,
    approved: rows.filter(r => r.status === '승인').length,
    pending:  rows.filter(r => r.status === '대기' || !r.status).length,
    rejected: rows.filter(r => r.status === '미선정').length,
    omg:      rows.filter(r => r.status === 'OMG').length,
  };

  return { success: true, stats: stats, calendar: calMap, rows: rows };
}

function getEmptyStats() {
  return { total:0, selected:0, done:0, draft:0, approved:0, pending:0, rejected:0, omg:0 };
}

// ===================================================
// 2. 접수 폼 저장
// ===================================================
function createForm(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('접수목록');
  if (!sheet) return { success: false, error: '접수목록 시트 없음' };

  const now = new Date();
  sheet.appendRow([
    now, data.name||'', data.phone||'', data.channel||'',
    data.channelName||'', data.followers||0, data.product||'',
    data.reason||'', '대기', '', '', '', '', ''
  ]);

  try {
    const dash = SpreadsheetApp.openById(DASHBOARD_SHEET_ID);
    const dashSheet = dash.getSheetByName(DASHBOARD_SHEET_NAME) || dash.getSheets()[0];
    dashSheet.appendRow([
      Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
      data.name||'', data.phone||'', data.channel||'',
      data.channelName||'', data.followers||0, data.product||'', data.reason||''
    ]);
  } catch(e) {}

  return { success: true, message: '신청이 완료되었습니다!' };
}

// ===================================================
// 3. 접수자 목록 조회
// ===================================================
function getApplicants(params) {
  params = params || {};
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('접수목록');
  if (!sheet) return { success: true, applicants: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, applicants: [] };

  function fmt(val) {
    if (!val) return '-';
    try {
      const d = new Date(val);
      return isNaN(d.getTime()) ? '-' : Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
    } catch(e) { return '-'; }
  }

  let applicants = data.slice(1).map((row, i) => ({
    rowIndex:    i + 2,
    date:        fmt(row[0]),
    name:        String(row[1]||'').trim(),
    phone:       String(row[2]||'').trim(),
    channel:     String(row[3]||'').trim(),
    channelName: String(row[4]||'').trim(),
    followers:   row[5]||0,
    product:     String(row[6]||'').trim(),
    reason:      String(row[7]||'').trim(),
    status:      String(row[8]||'').trim(),
    note:        String(row[13]||'').trim(),
  })).filter(r => r.name);

  if (params.status && params.status !== 'all') {
    applicants = applicants.filter(r => r.status === params.status);
  }

  return { success: true, applicants: applicants };
}

// ===================================================
// 4. 선정 저장 / 조회
// ===================================================
function saveSelected(data) {
  const ss = getOperationSheet();
  const appSheet = ss.getSheetByName('접수목록');
  if (appSheet && data.rowIndex) {
    appSheet.getRange(data.rowIndex, 9).setValue('선정');
    appSheet.getRange(data.rowIndex, 10).setValue(new Date());
  }

  const selSheet = ss.getSheetByName('선정결과');
  if (selSheet) {
    selSheet.appendRow([
      new Date(), data.name||'', data.phone||'', data.channel||'',
      data.channelName||'', data.followers||0, data.product||'',
      '선정', '', '', '', data.note||''
    ]);
  }
  return { success: true };
}

function getSelected(params) {
  params = params || {};
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('선정결과');
  if (!sheet) return { success: true, selected: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, selected: [] };

  const headers = data[0];
  const selected = data.slice(1).map((row, i) => {
    const obj = { rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  }).filter(r => r['이름'] || r['선정일']);

  return { success: true, selected: selected };
}

// ===================================================
// 5. 상태 업데이트
// ===================================================
function updateStatus(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('접수목록');
  if (!sheet || !data.rowIndex) return { success: false, error: '시트 또는 행 없음' };

  const statusMap = {
    '완료':   { col: 9, dateCol: 11 },
    '초안':   { col: 9, dateCol: 12 },
    '승인':   { col: 9, dateCol: 13 },
    '미선정': { col: 9, dateCol: null },
    'OMG':    { col: 9, dateCol: null },
  };

  const config = statusMap[data.status];
  if (!config) return { success: false, error: '알 수 없는 상태' };

  sheet.getRange(data.rowIndex, config.col).setValue(data.status);
  if (config.dateCol) sheet.getRange(data.rowIndex, config.dateCol).setValue(new Date());

  const selSheet = ss.getSheetByName('선정결과');
  if (selSheet && data.name) {
    const selData = selSheet.getDataRange().getValues();
    for (var i = 1; i < selData.length; i++) {
      if (String(selData[i][1]).trim() === String(data.name).trim()) {
        selSheet.getRange(i+1, 8).setValue(data.status);
        if (data.status === '완료') selSheet.getRange(i+1, 9).setValue(new Date());
        if (data.status === '초안') selSheet.getRange(i+1, 10).setValue(new Date());
        if (data.status === '승인') selSheet.getRange(i+1, 11).setValue(new Date());
        break;
      }
    }
  }
  return { success: true };
}

// ===================================================
// 6. OMG 리스트
// ===================================================
function getOmgList() {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('OMG리스트');
  if (!sheet) return { success: true, omgList: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, omgList: [] };

  const omgList = data.slice(1).map((row, i) => ({
    rowIndex: i+2, name: row[0], phone: row[1], reason: row[2], date: row[3], note: row[4]
  }));
  return { success: true, omgList: omgList };
}

function saveOMG(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('OMG리스트');
  if (!sheet) return { success: false, error: 'OMG리스트 시트 없음' };

  sheet.appendRow([data.name||'', data.phone||'', data.reason||'', new Date(), data.note||'']);

  if (data.rowIndex) {
    const appSheet = ss.getSheetByName('접수목록');
    if (appSheet) appSheet.getRange(data.rowIndex, 9).setValue('OMG');
  }
  return { success: true };
}

function deleteOMG(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('OMG리스트');
  if (!sheet || !data.rowIndex) return { success: false, error: '행 정보 없음' };
  sheet.deleteRow(data.rowIndex);
  return { success: true };
}

// ===================================================
// 7. 후기 제출
// ===================================================
function submitReview(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('추기제출');
  if (!sheet) return { success: false, error: '추기제출 시트 없음' };

  sheet.appendRow([
    new Date(), data.name||'', data.phone||'', data.channel||'',
    data.url||'', data.capture||'', '검토중', ''
  ]);
  return { success: true, message: '후기가 제출되었습니다!' };
}

function getReviews(params) {
  params = params || {};
  const ss = getOperationSheet();
  const sheets = ss.getSheets();
  const reviewSheet = sheets.filter(s => s.getName().includes('추기') || s.getName().includes('review'))[0]
    || ss.getSheetByName('추기제출');

  if (!reviewSheet) return { success: true, reviews: [] };

  const data = reviewSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, reviews: [] };

  const headers = data[0];
  const reviews = data.slice(1).map((row, i) => {
    const obj = { rowIndex: i+2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  }).filter(r => r['이름']);

  return { success: true, reviews: reviews };
}

// ===================================================
// 유틸: 운영 시트 가져오기
// ===================================================
function getOperationSheet() {
  if (OPERATION_SHEET_ID && OPERATION_SHEET_ID !== '') {
    try {
      return SpreadsheetApp.openById(OPERATION_SHEET_ID);
    } catch(e) {}
  }

  const files = DriveApp.getFilesByName('OzMom 체험단 운영 시트');
  if (files.hasNext()) {
    OPERATION_SHEET_ID = files.next().getId();
    return SpreadsheetApp.openById(OPERATION_SHEET_ID);
  }

  initOperationSheet();
  const files2 = DriveApp.getFilesByName('OzMom 체험단 운영 시트');
  if (files2.hasNext()) {
    OPERATION_SHEET_ID = files2.next().getId();
    return SpreadsheetApp.openById(OPERATION_SHEET_ID);
  }

  throw new Error('운영 시트를 찾거나 생성할 수 없습니다.');
}
