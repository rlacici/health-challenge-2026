/**
 * 2026 청소년 구강건강 챌린지 — Google Apps Script
 *
 * [최초 1회] Apps Script 편집기에서 setupSpreadsheet 실행 → 시트 자동 생성
 * [배포] 배포 > 새 배포 > 웹 앱 (액세스: 모든 사용자)
 * [연결] 배포 URL을 index-challenge.html 의 CONFIG.appsScriptUrl 에 입력
 */

const SPREADSHEET_ID = "1h6WU1UPvdDNji57uPqdD3vGv3pjF9iJ4F8R8uT44n3g";

const DEFAULT_IMAGES = [
  ["01", "그냥 다 귀찮음", 2],
  ["02", "충치는 연락 없이 옴", 2],
  ["03", "칫솔은 아직 자는 중", 2],
  ["04", "웃다가 점심시간 끝남", 2],
  ["05", "나중에 할게요", 2],
  ["06", "점심시간 통장", 2],
  ["07", "오늘만 안 해야지", 2],
  ["08", "하고 싶은 거 하는 시간", 2],
  ["09", "오늘의 치아 날씨", 2],
  ["10", "아무 것도 안 해도 재밌음", 2],
  ["11", "오늘의 퀘스트", 2],
  ["12", "건강 점검 결과", 2],
  ["13", "놀거리는 많음", 2],
  ["14", "원래 하는 거라서", 2],
  ["15", "칫솔 못 찾겠다", 2],
  ["16", "아플 땐 양치 생각남", 2],
  ["17", "자연스럽게 다음 순서", 2],
  ["18", "하다 보니 함", 2],
  ["19", "10명 중 2명", 1],
  ["20", "인천 전국 16위", 1]
];

const DEFAULT_HEALTH_CENTER = "🦷 미추홀구 보건소와 함께해요!✨";

function doGet(e) {
  e = e || {};
  try {
    validateSpreadsheetId();
    const result = recordTagAndPickImage();
    return createResponse(result, e.parameter.callback);
  } catch (error) {
    return createResponse({
      error: true,
      message: String(error)
    }, e.parameter.callback);
  }
}

function createResponse(data, callback) {
  const json = JSON.stringify(data);

  if (callback) {
    const safeCallback = String(callback).replace(/[^a-zA-Z0-9_]/g, "");
    return ContentService
      .createTextOutput(safeCallback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function validateSpreadsheetId() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf("여기에") >= 0) {
    throw new Error("SPREADSHEET_ID를 실제 시트 ID로 바꿔 주세요. (Apps Script 맨 위)");
  }
}

function recordTagAndPickImage() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const settings = ss.getSheetByName("설정");
  const imagesSheet = ss.getSheetByName("이미지");
  const statsSheet = ss.getSheetByName("통계");

  if (!settings || !imagesSheet || !statsSheet) {
    throw new Error("시트가 없습니다. setupSpreadsheet()를 먼저 실행하세요.");
  }

  const images = getImageList(imagesSheet);
  if (!images.length) {
    throw new Error("이미지 시트에 데이터가 없습니다.");
  }

  const picked = pickWeighted(images);
  const newCount = incrementTotal(statsSheet);
  incrementImageCount(statsSheet, picked.id);

  return {
    totalCount: newCount,
    imageId: picked.id,
    healthCenterName: String(settings.getRange("B1").getValue() || DEFAULT_HEALTH_CENTER)
  };
}

function getImageList(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return rows
    .filter((row) => row[0])
    .map((row) => ({
      id: padId(row[0]),
      alt: String(row[1] || ""),
      weight: Number(row[2]) || 1
    }));
}

function pickWeighted(images) {
  const total = images.reduce((sum, img) => sum + img.weight, 0);
  let random = Math.random() * total;

  for (const image of images) {
    random -= image.weight;
    if (random <= 0) return image;
  }

  return images[images.length - 1];
}

function incrementTotal(statsSheet) {
  const cell = statsSheet.getRange("B1");
  const current = Number(cell.getValue()) || 0;
  const next = current + 1;
  cell.setValue(next);
  return next;
}

function incrementImageCount(statsSheet, imageId) {
  const lastRow = statsSheet.getLastRow();
  if (lastRow < 4) return;

  const ids = statsSheet.getRange(4, 1, lastRow - 3, 1).getValues();
  const targetId = padId(imageId);

  for (let i = 0; i < ids.length; i++) {
    if (padId(ids[i][0]) === targetId) {
      const row = i + 4;
      const countCell = statsSheet.getRange(row, 3);
      countCell.setValue((Number(countCell.getValue()) || 0) + 1);
      return;
    }
  }
}

function padId(value) {
  const text = String(value).trim();
  if (/^\d+$/.test(text) && text.length < 2) {
    return text.padStart(2, "0");
  }
  return text;
}

/**
 * 최초 1회 실행: 설정·이미지·통계 시트를 만들고 20장 데이터를 채웁니다.
 * Apps Script 편집기에서 이 함수 선택 → ▶ 실행
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let settings = ss.getSheetByName("설정");
  if (!settings) settings = ss.insertSheet("설정");
  settings.clear();
  settings.getRange("A1").setValue("보건소 이름");
  settings.getRange("B1").setValue(DEFAULT_HEALTH_CENTER);
  settings.setColumnWidth(1, 120);
  settings.setColumnWidth(2, 360);

  let images = ss.getSheetByName("이미지");
  if (!images) images = ss.insertSheet("이미지");
  images.clear();
  images.getRange(1, 1, 1, 3).setValues([["ID", "제목", "가중치"]]);
  images.getRange(2, 1, DEFAULT_IMAGES.length, 3).setValues(DEFAULT_IMAGES);
  images.setFrozenRows(1);
  images.setColumnWidth(1, 60);
  images.setColumnWidth(2, 220);
  images.setColumnWidth(3, 80);

  let stats = ss.getSheetByName("통계");
  if (!stats) stats = ss.insertSheet("통계");
  stats.clear();
  stats.getRange("A1").setValue("누적 태그 수");
  stats.getRange("B1").setValue(0);
  stats.getRange("A3:C3").setValues([["ID", "제목", "노출 횟수"]]);

  const statRows = DEFAULT_IMAGES.map((row) => [row[0], row[1], 0]);
  stats.getRange(4, 1, statRows.length, 3).setValues(statRows);
  stats.setFrozenRows(3);
  stats.setColumnWidth(1, 60);
  stats.setColumnWidth(2, 220);
  stats.setColumnWidth(3, 100);

  const defaultSheet = ss.getSheetByName("시트1");
  if (defaultSheet && ss.getSheets().length > 3) {
    ss.deleteSheet(defaultSheet);
  }
}

/**
 * 제목이 깨졌을 때 복구: 이미지 시트 제목 → 통계 시트 B열
 * 노출 횟수(C열)·누적 태그(B1)는 그대로 둡니다.
 * Apps Script 편집기에서 1회 실행
 */
function repairSpreadsheet() {
  validateSpreadsheetId();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const statsSheet = ss.getSheetByName("통계");

  if (!statsSheet) {
    throw new Error("통계 시트가 없습니다. setupSpreadsheet()를 먼저 실행하세요.");
  }

  const statsLastRow = statsSheet.getLastRow();
  if (statsLastRow < 4) return;

  const statsIds = statsSheet.getRange(4, 1, statsLastRow - 3, 1).getValues();

  for (let i = 0; i < statsIds.length; i++) {
    const statsId = padId(statsIds[i][0]);
    const match = DEFAULT_IMAGES.find((row) => padId(row[0]) === statsId);
    if (match) {
      statsSheet.getRange(i + 4, 2).setValue(String(match[1] || ""));
    }
  }
}

/**
 * Apps Script 편집기에서 1회 실행 — 시트 연결이 되는지 확인
 */
function testConnection() {
  validateSpreadsheetId();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("연결 성공: " + ss.getName());
}
