// ====================================================
// OzMom 체험단 관리 시스템 - Apps Script 최종본
// ====================================================
// 📌 설정값 (반드시 입력)
// ====================================================

// 대시보드용 시트 (읽기 전용)
const DASHBOARD_SHEET_ID = '1fhdHU7Y5a_EDr8fATJ3OMyhlWqWFPGpCHSUCm7fxIRI';
const DASHBOARD_SHEET_NAME = '★자사몰 외 채널 후기';

// 운영용 시트 (설문/선정/OMG/후기 누적)
// ⚠️ 처음 배포 후 아래 함수 initOperationSheet() 를 한 번 실행하면 자동 생성됩니다
let OPERATION_SHEET_ID = ''; // 자동 생성 후 여기에 ID가 채워집니다

// ====================================================
// 초기 설정: 운영 시트 최초 1회 생성
// Apps Script 편집기에서 이 함수를 한 번만 실행하세요!
// ====================================================
function initOperationSheet() {
  const ss = SpreadsheetApp.create('OzMom 체험단 운영 시트');
  
  // 시트 구성
  const sheetNames = ['설문폼목록', '선정결과', 'OMG리스트', '후기제출'];
  
  // 기본 시트 이름 변경
  ss.getSheets()[0].setName('설문폼목록');
  
  // 나머지 시트 생성
  sheetNames.slice(1).forEach(name => ss.insertSheet(name));
  
  // 헤더 설정
  ss.getSheetByName('설문폼목록')
    .appendRow(['폼ID', '폼제목', '폼URL', '편집URL', '신청마감일', '후기마감일', '생성일', '상태']);
  
  ss.getSheetByName('선정결과')
    .appendRow(['폼ID', '회차', '성함', '전화번호', 'SNS', '제품명', '제품링크', '선정일', '후기마감일', '문자발송', '후기상태']);
  
  ss.getSheetByName('OMG리스트')
    .appendRow(['성함', '전화번호', '사유', '등록일', '메모']);
  
  ss.getSheetByName('후기제출')
    .appendRow(['타임스탬프', '성함', '전화번호', '후기URL', '자사몰후기이미지', '몰후기이미지', '구매내역이미지', '페이백금액', '계좌번호', '예금주', '은행']);
  
  // ID 로그 출력
  Logger.log('✅ 운영 시트 생성 완료!');
  Logger.log('📋 운영 시트 ID: ' + ss.getId());
  Logger.log('📋 운영 시트 URL: ' + ss.getUrl());
  Logger.log('⚠️ 위 ID를 복사해서 OPERATION_SHEET_ID 변수에 넣어주세요!');
}

// ====================================================
// CORS 헤더
// ====================================================
function setCORSHeaders(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ====================================================
// 라우팅
// ====================================================
function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch(action) {
      case 'getDashboard': result = getDashboard(e.parameter); break;
      case 'getFormList':  result = getFormList(); break;
      case 'getApplicants': result = getApplicants(e.parameter.formId); break;
      case 'getSelected':  result = getSelected(); break;
      case 'getOmgList':   result = getOmgList(); break;
      case 'getReviews':   result = getReviews(); break;
      default: result = { success: false, error: '알 수 없는 요청' };
    }
  } catch(err) {
    result = { success: false, error: err.message };
  }
  return setCORSHeaders(
    ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON)
  );
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  let result;
  try {
    switch(data.action) {
      case 'createForm':   result = createForm(data); break;
      case 'saveSelected': result = saveSelected(data); break;
      case 'saveOMG':      result = saveOMG(data); break;
      case 'deleteOMG':    result = deleteOMG(data); break;
      default: result = { success: false, error: '알 수 없는 요청' };
    }
  } catch(err) {
    result = { success: false, error: err.message };
  }
  return setCORSHeaders(
    ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON)
  );
}

// ====================================================
// 1. 대시보드 데이터 읽기
// 열 구조: A=yyyy, B=m, C=헤이즐요청일자, D=진행상황,
//          E=제품명, F=채널, M=마감기한, N=완료일,
//          P=기안일, Q=결재예정일
// ====================================================
function getDashboard(params) {
  const ss = SpreadsheetApp.openById(DASHBOARD_SHEET_ID);
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return { success: false, error: '시트를 찾을 수 없습니다: ' + DASHBOARD_SHEET_NAME };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [], calendar: {} };

  // 헤더 행 찾기 (yyyy, m 같은 값이 있는 행)
  let headerRow = 0;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    if (String(row[0]).toLowerCase().includes('y') || String(row[2]).includes('일자') || String(row[3]).includes('진행')) {
      headerRow = i;
      break;
    }
  }

  // 열 인덱스 (0-based)
  const C=2, D=3, E=4, F=5, M=12, N=13, P=15, Q=16;

  let rows = data.slice(headerRow + 1).map(row => ({
    date:        formatDate(row[C]),
    status:      String(row[D] || '').trim(),
    product:     String(row[E] || '').trim(),
    channel:     String(row[F] || '').trim(),
    deadline:    formatDate(row[M]),
    doneDate:    formatDate(row[N]),
    draftDate:   formatDate(row[P]),
    approveDate: formatDate(row[Q]),
    raw:         row[C],
  })).filter(r => r.product || r.status);

  // 기간 필터 (C열 기준)
  if (params.from) rows = rows.filter(r => r.raw && new Date(r.raw) >= new Date(params.from));
  if (params.to)   rows = rows.filter(r => r.raw && new Date(r.raw) <= new Date(params.to + 'T23:59:59'));

  // 달력 집계
  const calMap = {};
  rows.forEach(r => {
    [['deadline',r.deadline],['doneDate',r.doneDate],['draftDate',r.draftDate],['approveDate',r.approveDate]]
      .forEach(([type, dateStr]) => {
        if (!dateStr || dateStr === '-') return;
        if (!calMap[dateStr]) calMap[dateStr] = { deadline:0, doneDate:0, draftDate:0, approveDate:0 };
        calMap[dateStr][type]++;
      });
  });

  return { success: true, data: rows, calendar: calMap };
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  if (!s || s === '0') return '';
  // "2026. 4. 1" 또는 "2026.4.1" 형식
  const m = s.match(/(\d{4})[.\s]+(\d{1,2})[.\s]+(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  return s.split('T')[0];
}

// ====================================================
// 2. 구글 폼 자동 생성
// ====================================================
function createForm(data) {
  const { title, guide, deadline, reviewDeadline, products } = data;

  const description =
    (guide || '') + '\n\n' +
    (deadline       ? '📅 신청 마감일: ' + deadline       + '\n' : '') +
    (reviewDeadline ? '⭐ 후기 제출 마감일: ' + reviewDeadline + '\n' : '');

  const form = FormApp.create(title || '스피드 체험단 신청 설문지');
  form.setDescription(description.trim());
  form.setCollectEmail(false);

  // 신청자 정보
  form.addSectionHeaderItem().setTitle('📝 신청자 정보');

  form.addTextItem().setTitle('성함').setRequired(true);

  const phoneItem = form.addTextItem();
  phoneItem.setTitle('전화번호').setHelpText('형식: 010-0000-0000').setRequired(true);
  phoneItem.setValidation(
    phoneItem.createValidation()
      .requireTextMatchesPattern('^\\d{3}-\\d{3,4}-\\d{4}$')
      .setHelpText('010-0000-0000 형식으로 입력해주세요')
      .build()
  );

  form.addTextItem()
    .setTitle('SNS URL')
    .setHelpText('블로그, 인스타그램 등 후기를 작성하실 SNS 주소')
    .setRequired(true);

  // 제품 선택
  form.addSectionHeaderItem()
    .setTitle('📦 신청 제품')
    .setHelpText('신청할 제품을 선택해주세요');

  if (products && products.length > 0) {
    products.forEach(p => {
      // 제품 이미지 삽입 시도
      if (p.img) {
        try {
          const blob = UrlFetchApp.fetch(p.img, {
            followRedirects: true,
            muteHttpExceptions: true,
            headers: { 'User-Agent': 'Mozilla/5.0' }
          }).getBlob();
          blob.setContentType('image/jpeg');
          form.addImageItem()
            .setTitle(p.name)
            .setImage(blob)
            .setAlignment(FormApp.Alignment.CENTER);
        } catch(e) {
          Logger.log('이미지 실패: ' + p.name);
        }
      }
      // 제품별 체크박스
      form.addCheckboxItem()
        .setTitle(p.name + (p.link ? '\n링크: ' + p.link : ''))
        .setChoiceValues(['신청합니다 ✅'])
        .setRequired(false);
    });
  }

  // 개인정보 동의
  form.addCheckboxItem()
    .setTitle('개인정보 수집 및 이용 동의 (필수)')
    .setRequired(true)
    .setChoiceValues(['동의합니다']);

  // 운영 시트에 폼 응답 연결
  const opSS = getOperationSheet();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, opSS.getId());

  // 폼 목록 시트에 기록
  const formSheet = opSS.getSheetByName('설문폼목록');
  formSheet.appendRow([
    form.getId(),
    title,
    form.getPublishedUrl(),
    form.getEditUrl(),
    deadline || '',
    reviewDeadline || '',
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'),
    '진행중'
  ]);

  return {
    success: true,
    formUrl: form.getPublishedUrl(),
    editUrl: form.getEditUrl(),
    formId:  form.getId(),
    sheetId: opSS.getId(),
    message: '구글 폼이 생성되었습니다!'
  };
}

// ====================================================
// 3. 설문 폼 목록 가져오기
// ====================================================
function getFormList() {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('설문폼목록');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, forms: [] };

  const forms = data.slice(1).map(row => ({
    formId:       row[0],
    title:        row[1],
    formUrl:      row[2],
    editUrl:      row[3],
    deadline:     row[4],
    reviewDeadline: row[5],
    createdAt:    row[6],
    status:       row[7],
  })).filter(r => r.formId);

  return { success: true, forms };
}

// ====================================================
// 4. 설문 응답(신청자) 불러오기
// ====================================================
function getApplicants(formId) {
  const ss = getOperationSheet();
  const sheets = ss.getSheets();

  // 폼 응답 시트 찾기 (이름에 "응답" 또는 "폼 응답" 포함)
  let responseSheet = null;
  if (formId) {
    // formId로 정확히 매칭되는 시트 찾기
    const formSheet = ss.getSheetByName('설문폼목록');
    const formData = formSheet.getDataRange().getValues();
    const formRow = formData.find(r => String(r[0]) === String(formId));
    if (formRow) {
      // 폼 제목으로 응답 시트 찾기
      const title = formRow[1];
      responseSheet = sheets.find(s => s.getName().includes(title.substring(0, 10)));
    }
  }
  // 못 찾으면 가장 최근 응답 시트
  if (!responseSheet) {
    responseSheet = sheets.find(s =>
      s.getName().includes('응답') || s.getName().includes('폼 응답')
    );
  }
  if (!responseSheet) return { success: true, applicants: [], message: '응답 시트가 없습니다' };

  const data = responseSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, applicants: [] };

  const headers = data[0];
  const applicants = data.slice(1).map((row, i) => {
    // 헤더 기반으로 매핑
    const obj = { rowIndex: i + 2, sheetName: responseSheet.getName() };
    const nameIdx   = findHeaderIdx(headers, ['성함', '이름', 'name']);
    const phoneIdx  = findHeaderIdx(headers, ['전화번호', 'phone', '연락처']);
    const snsIdx    = findHeaderIdx(headers, ['SNS', 'sns', 'URL', '링크']);
    const timeIdx   = findHeaderIdx(headers, ['타임스탬프', 'Timestamp', '제출']);

    obj.name    = row[nameIdx]  || '';
    obj.phone   = row[phoneIdx] || '';
    obj.sns     = row[snsIdx]   || '';
    obj.time    = formatDate(row[timeIdx]) || '';
    obj.product = []; // 제품은 아래에서 추출
    // 체크박스 응답에서 제품명 추출 (신청합니다 ✅ 포함된 열)
    headers.forEach((h, j) => {
      if (String(row[j]).includes('신청합니다')) {
        obj.product.push(String(h).replace('신청하시겠습니까?', '').trim());
      }
    });
    obj.product = obj.product.join(', ');
    return obj;
  }).filter(r => r.name);

  return { success: true, applicants };
}

function findHeaderIdx(headers, candidates) {
  for (const c of candidates) {
    const i = headers.findIndex(h => String(h).includes(c));
    if (i >= 0) return i;
  }
  return 0;
}

// ====================================================
// 5. 선정 결과 저장
// ====================================================
function saveSelected(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('선정결과');
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

  data.selected.forEach(person => {
    sheet.appendRow([
      data.formId || '',
      data.round  || today + ' 선정',
      person.name,
      person.phone,
      person.sns  || '',
      person.product || '',
      person.productLink || '',
      today,
      data.reviewDeadline || '',
      '미발송',
      '미제출'
    ]);
  });

  return { success: true, message: `${data.selected.length}명 선정 결과 저장 완료` };
}

function getSelected() {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('선정결과');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, selected: [] };
  const headers = data[0];
  const selected = data.slice(1).map((row, i) => {
    const obj = { rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
  return { success: true, selected };
}

// ====================================================
// 6. OMG 리스트
// ====================================================
function getOmgList() {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('OMG리스트');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, omgList: [] };

  const omgList = data.slice(1).map((row, i) => ({
    rowIndex: i + 2,
    name:   row[0],
    phone:  row[1],
    reason: row[2],
    date:   row[3],
    memo:   row[4],
  })).filter(r => r.name);

  return { success: true, omgList };
}

function saveOMG(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('OMG리스트');
  sheet.appendRow([
    data.name,
    data.phone  || '',
    data.reason || '',
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'),
    data.memo   || ''
  ]);
  return { success: true, message: data.name + '님 OMG 등록 완료' };
}

function deleteOMG(data) {
  const ss = getOperationSheet();
  const sheet = ss.getSheetByName('OMG리스트');
  if (data.rowIndex) sheet.deleteRow(data.rowIndex);
  return { success: true, message: 'OMG 삭제 완료' };
}

// ====================================================
// 7. 후기 데이터
// ====================================================
function getReviews() {
  const ss = getOperationSheet();
  // 후기 제출 폼 응답 시트 찾기
  const sheets = ss.getSheets();
  const reviewSheet = sheets.find(s =>
    s.getName().includes('후기') || s.getName().includes('review')
  ) || ss.getSheetByName('후기제출');

  const data = reviewSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, reviews: [] };

  const headers = data[0];
  const reviews = data.slice(1).map((row, i) => {
    const obj = { rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  }).filter(r => r['성함'] || r['이름']);

  return { success: true, reviews };
}

// ====================================================
// 유틸: 운영 시트 가져오기
// ====================================================
function getOperationSheet() {
  if (OPERATION_SHEET_ID && OPERATION_SHEET_ID !== '') {
    return SpreadsheetApp.openById(OPERATION_SHEET_ID);
  }
  // ID가 없으면 자동으로 찾거나 생성
  const files = DriveApp.getFilesByName('OzMom 체험단 운영 시트');
  if (files.hasNext()) {
    const file = files.next();
    OPERATION_SHEET_ID = file.getId();
    return SpreadsheetApp.openById(OPERATION_SHEET_ID);
  }
  // 없으면 새로 생성
  initOperationSheet();
  const files2 = DriveApp.getFilesByName('OzMom 체험단 운영 시트');
  if (files2.hasNext()) {
    OPERATION_SHEET_ID = files2.next().getId();
    return SpreadsheetApp.openById(OPERATION_SHEET_ID);
  }
  throw new Error('운영 시트를 찾을 수 없습니다. initOperationSheet()를 실행해주세요.');
}
