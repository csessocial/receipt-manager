const FOLDER_ID = '10e0PUtuw0HfG8sotx6IiueQbFHZeOBhE';
const SHEET_ID  = '17h2ZJWsSbLsI4SDaxC8szKYIj07xwb28MEeFSFkgEJQ';

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
  let imageUrl = '';
  if (data.image) {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const bytes  = Utilities.base64Decode(data.image);
    const blob   = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', '영수증_' + data.id + '.jpg');
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    imageUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  }
  var sheet = getSheet();
  sheet.appendRow([
    data.id,
    new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
    data.date,
    data.vendor,
    data.amount,
    data.purpose || '',
    data.lab,
    data.team,
    data.category,
    data.program  || '',
    data.note     || '',
    data.consent  || '',
    imageUrl
  ]);
  return respond({ success: true, imageUrl: imageUrl });
}

function deleteRow(id) {
  var sheet = getSheet();
  var vals  = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return respond({ success: true });
}

function doGet(e) {
  var sheet = getSheet();
  var vals  = sheet.getDataRange().getValues();
  if (vals.length <= 1) return respond([]);
  var rows = [];
  for (var i = vals.length - 1; i >= 1; i--) {
    var r = vals[i];
    if (!r[0]) continue;
    rows.push({
      id: String(r[0]), created_at: String(r[1]), date: String(r[2]),
      vendor: r[3], amount: Number(r[4]), purpose: r[5],
      lab: r[6], team: r[7], category: r[8],
      program: r[9], note: r[10], consent: r[11], image_url: r[12]
    });
  }
  return respond(rows);
}

function setup() {
  var sheet = getSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID','등록일시','날짜','결제처','금액','목적','실','담당팀','예산항목','프로그램','비고','동의서','영수증URL']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#1e2235').setFontColor('#ffffff');
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(13, 300);
  }
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('영수증') || ss.insertSheet('영수증');
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
