// ============================================================
// 법인카드 영수증 관리 - Google Apps Script
// ============================================================
// 설정: 아래 두 값을 본인 것으로 교체하세요
const FOLDER_ID = '10e0PUtuw0HfG8sotx6IiueQbFHZeOBhE';
const SHEET_ID  = '17h2ZJWsSbLsI4SDaxC8szKYIj07xwb28MEeFSFkgEJQ';

// ============================================================
// POST: 영수증 등록 / 삭제
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'delete') return deleteRow(data.id);
    return insertRow(data);
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function insertRow(data) {
  // 이미지 → 드라이브 저장
  let imageUrl = '';
  if (data.image) {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const bytes  = Utilities.base64Decode(data.image);
    const blob   = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', `영수증_${data.id}.jpg`);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    imageUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  }

  // 시트에 행 추가
  const sheet = getSheet();
  sheet.appendRow([
    data.id,
    new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
    data.date,
    data.vendor,
    data.amount,
    data.lab,
    data.team,
    data.category,
    data.program     || '',
    data.description || '',
    data.note        || '',
    imageUrl
  ]);

  return respond({ success: true, imageUrl });
}

function deleteRow(id) {
  const sheet = getSheet();
  const vals  = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return respond({ success: true });
}

// ============================================================
// GET: 영수증 목록 조회
// ============================================================
function doGet(e) {
  const sheet = getSheet();
  const vals  = sheet.getDataRange().getValues();
  if (vals.length <= 1) return respond([]);

  const rows = vals.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id:          String(r[0]),
      created_at:  String(r[1]),
      date:        String(r[2]),
      vendor:      r[3],
      amount:      Number(r[4]),
      lab:         r[5],
      team:        r[6],
      category:    r[7],
      program:     r[8],
      description: r[9],
      note:        r[10],
      image_url:   r[11]
    }))
    .reverse();

  return respond(rows);
}

// ============================================================
// 초기 실행: 시트 헤더 생성 (최초 1회)
// ============================================================
function setup() {
  const sheet = getSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID','등록일시','날짜','가맹점','금액','실','담당팀','항목','프로그램','내용','비고','영수증URL']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#1e2235').setFontColor('#ffffff');
  }
}

// ============================================================
// 공통
// ============================================================
function getSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const name  = '영수증';
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
