// ====================================================
// 체험단 관리 시스템 - Google Apps Script 백엔드
// 아래 코드를 Apps Script 편집기에 전체 붙여넣기 하세요
// ====================================================

// ▼ 여기에 본인 스프레드시트 ID를 넣어주세요
//   (구글 시트 주소에서 /d/XXXXX/edit 에서 XXXXX 부분)
const SHEET_ID = "여기에_스프레드시트_ID_넣기";

// ▼ 발송자 이름 (문자 발송 시 표시될 이름)
const SENDER_NAME = "체험단";


// ====================================================
// 요청을 받아주는 메인 함수 (수정하지 마세요)
// ====================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    if (action === "createForm")       result = createForm(data);
    else if (action === "getProducts") result = getProducts();
    else if (action === "saveProduct") result = saveProduct(data);
    else if (action === "getApplicants") result = getApplicants();
    else if (action === "getOmgList")  result = getOmgList();
    else if (action === "saveOmgList") result = saveOmgList(data);
    else if (action === "saveSelected") result = saveSelected(data);
    else if (action === "getReviews")  result = getReviews();
    else if (action === "updateReview") result = updateReview(data);
    else if (action === "getPaybacks") result = getPaybacks();
    else if (action === "savePayback") result = savePayback(data);
    else result = { success: false, message: "알 수 없는 요청입니다." };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: "체험단 관리 서버 정상 작동 중!" }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ====================================================
// 1. 구글 폼 자동 생성
// ====================================================
function createForm(data) {
  const products       = data.products || [];
  const deadline       = data.deadline || "";
  const reviewDeadline = data.reviewDeadline || "";

  const description =
    "안녕하세요! 스피드 체험단 신청 설문지입니다.\n" +
    "아래 제품을 확인 후 신청해 주세요.\n" +
    "선정 결과는 오즈키즈 카카오톡 알람톡으로 안내드립니다.\n\n" +
    (deadline       ? "📅 신청 마감: " + deadline       + "\n" : "") +
    (reviewDeadline ? "⭐ 리뷰 마감: " + reviewDeadline + " ← 반드시 지켜주세요!\n" : "");

  const form = FormApp.create("스피드 체험단 신청");
  form.setDescription(description.trim());
  form.setCollectEmail(false);

  // ── 1. 신청자 정보 (상단) ──
  form.addSectionHeaderItem().setTitle("📝 신청자 정보");

  form.addTextItem()
    .setTitle("성함")
    .setRequired(true);

  form.addTextItem()
    .setTitle("전화번호")
    .setHelpText("형식: 010-0000-0000")
    .setRequired(true);

  form.addTextItem()
    .setTitle("SNS 후기 작성하실 경우 업로드해주실 SNS 1개 써주세요")
    .setHelpText("URL을 입력해 주세요")
    .setRequired(false);

  form.addMultipleChoiceItem()
    .setTitle("자사몰 후기, 몰 후기, SNS 후기 3가지 후기를 모두 해주셔야 합니다. 가능하신가요?")
    .setChoiceValues(["예, 모두 작성 가능합니다", "아니오, 불가합니다"])
    .setRequired(true);

  // ── 2. 제품 선택 (하단) ──
  form.addSectionHeaderItem()
    .setTitle("📦 신청 제품 목록")
    .setHelpText("신청할 제품을 모두 선택해 주세요 (복수 선택 가능)");

  products.forEach(function(p) {
    // 이미지 삽입
    if (p.imageUrl) {
      try {
        var imgBlob = UrlFetchApp.fetch(p.imageUrl, {
          followRedirects: true,
          muteHttpExceptions: true,
          headers: { "User-Agent": "Mozilla/5.0" }
        }).getBlob();
        imgBlob.setContentType("image/jpeg");
        form.addImageItem()
          .setTitle(p.name)
          .setImage(imgBlob)
          .setAlignment(FormApp.Alignment.CENTER);
      } catch(e) {
        Logger.log("이미지 실패: " + p.name + " / " + e.message);
      }
    }
    // 이미지 바로 아래 체크박스
    form.addCheckboxItem()
      .setTitle(p.name + " 신청하시겠습니까?")
      .setChoiceValues(["신청합니다 ✅"])
      .setRequired(false);
  });

  // 스프레드시트 연결
  if (SHEET_ID && SHEET_ID !== "여기에_스프레드시트_ID_넣기") {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, SHEET_ID);
  }

  return {
    success: true,
    formUrl: form.getPublishedUrl(),
    editUrl: form.getEditUrl(),
    formId:  form.getId(),
    message: "구글 폼이 생성되었습니다!"
  };
}

// 쿠팡 링크에서 썸네일 이미지 URL 추출
function getCoupangThumbnail(url) {
  try {
    var response = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" }
    });
    var html = response.getContentText();

    // og:image 메타태그에서 이미지 추출
    var match = html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i);
    if (!match) {
      match = html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image["']/i);
    }
    if (match && match[1]) {
      return match[1];
    }
    return null;
  } catch(e) {
    return null;
  }
}


// ====================================================
// 2. 제품 목록 저장/불러오기
// ====================================================
function getProducts() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "제품목록");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, products: [] };

  const products = data.slice(1).map(row => ({
    id: row[0], name: row[1], link: row[2], imageUrl: row[3], status: row[4]
  }));
  return { success: true, products };
}

function saveProduct(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "제품목록");

  // 헤더 없으면 추가
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "제품명", "링크", "이미지URL", "상태", "등록일"]);
  }

  const id = new Date().getTime().toString();
  sheet.appendRow([
    id, data.name, data.link, data.imageUrl || "", "등록됨",
    new Date().toLocaleDateString("ko-KR")
  ]);
  return { success: true, id, message: "제품이 저장되었습니다." };
}


// ====================================================
// 3. 신청자 불러오기 (구글 폼 응답 시트에서)
// ====================================================
function getApplicants() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();

  // "폼 응답" 또는 "신청" 포함된 시트 찾기
  let responseSheet = sheets.find(s =>
    s.getName().includes("폼 응답") || s.getName().includes("신청") || s.getName().includes("응답")
  );

  if (!responseSheet) return { success: true, applicants: [], message: "폼 응답 시트가 없습니다." };

  const data = responseSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, applicants: [] };

  const headers = data[0];
  const applicants = data.slice(1).map((row, i) => {
    const obj = { rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });

  return { success: true, applicants };
}


// ====================================================
// 4. OMG 리스트 관리
// ====================================================
function getOmgList() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "OMG리스트");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, omgList: [] };

  const omgList = data.slice(1).map(row => ({
    name: row[0], phone: row[1], count: row[2], addedDate: row[3]
  }));
  return { success: true, omgList };
}

function saveOmgList(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "OMG리스트");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["이름", "전화번호", "마감미준수횟수", "등록일"]);
  }

  sheet.appendRow([
    data.name, data.phone || "", data.count || 1,
    new Date().toLocaleDateString("ko-KR")
  ]);
  return { success: true, message: `${data.name}님이 OMG 리스트에 추가되었습니다.` };
}


// ====================================================
// 5. 선정 결과 저장
// ====================================================
function saveSelected(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "선정결과");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["회차", "성함", "전화번호", "제품명", "선정일", "후기마감일", "문자발송", "후기상태"]);
  }

  const round = data.round || new Date().toLocaleDateString("ko-KR") + " 선정";
  data.selected.forEach(person => {
    sheet.appendRow([
      round, person.name, person.phone, person.product,
      new Date().toLocaleDateString("ko-KR"),
      data.reviewDeadline || "",
      "미발송", "미제출"
    ]);
  });

  return { success: true, message: `${data.selected.length}명의 선정 결과가 저장되었습니다.` };
}


// ====================================================
// 6. 후기 현황 조회 및 업데이트
// ====================================================
function getReviews() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "선정결과");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, reviews: [] };

  const headers = data[0];
  const reviews = data.slice(1).map((row, i) => {
    const obj = { rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
  return { success: true, reviews };
}

function updateReview(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("선정결과");
  if (!sheet) return { success: false, message: "선정결과 시트가 없습니다." };

  // data.rowIndex, data.column, data.value
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIdx = headers.indexOf(data.column) + 1;
  if (colIdx === 0) return { success: false, message: "컬럼을 찾을 수 없습니다." };

  sheet.getRange(data.rowIndex, colIdx).setValue(data.value);

  // 마감일 초과 체크 후 셀 빨강 표시
  if (data.isLate) {
    sheet.getRange(data.rowIndex, 1, 1, sheet.getLastColumn())
      .setBackground("#FFE0E0")
      .setFontColor("#CC0000");
  }

  return { success: true, message: "후기 상태가 업데이트되었습니다." };
}


// ====================================================
// 7. 페이백 저장
// ====================================================
function getPaybacks() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "페이백");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, paybacks: [] };

  const headers = data[0];
  const paybacks = data.slice(1).map((row, i) => {
    const obj = { rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
  return { success: true, paybacks };
}

function savePayback(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "페이백");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["회차", "성함", "전화번호", "제품명", "금액", "지급예정일", "등록일", "상태"]);
  }

  data.list.forEach(person => {
    sheet.appendRow([
      data.round, person.name, person.phone, person.product,
      data.amount, data.payDate,
      new Date().toLocaleDateString("ko-KR"), "처리중"
    ]);
  });

  return { success: true, message: `${data.list.length}명의 페이백 정보가 저장되었습니다.` };
}


// ====================================================
// 유틸: 시트가 없으면 만들어주는 함수
// ====================================================
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}
