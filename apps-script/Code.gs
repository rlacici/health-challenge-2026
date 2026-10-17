/**
 * 2026 청소년 구강건강 챌린지 — Google Apps Script
 *
 * [최초 1회] Apps Script 편집기에서 setupSpreadsheet 실행 → 시트 자동 생성
 * [주의] 통계가 있는 뒤에는 setupSpreadsheet 를 다시 실행하지 마세요. 횟수가 지워집니다.
 * [집계] 태그는 수신 시트에 한 줄로 쌓고, 통계 칸은 기존 횟수+수신 고유번호로 계산합니다.
 *         같은 태그번호는 한 번만 셉니다. 숫자는 지우지 않습니다.
 * [1회 실행] 붙여 넣은 뒤 useSumFormulas 실행
 * [배포] 이미 배포한 뒤 코드만 바꿀 때는 「새 배포」 금지.
 *         배포 > 배포 관리 > 기존 웹 앱 연필 > 버전을 새 버전으로 > 배포
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

const DEFAULT_HEALTH_CENTER = "🦷 미추홀구 보건소와 함께해요!";

function doGet(e) {
  return handleTagRequest(e || {});
}

function doPost(e) {
  return handleTagRequest(e || {});
}

function handleTagRequest(e) {
  const callback = e.parameter && e.parameter.callback;
  try {
    validateSpreadsheetId();
    const result = recordIncomingTags(collectTags_(e));
    return createResponse(result, callback);
  } catch (error) {
    return createResponse({
      error: true,
      message: String(error),
      saved: []
    }, callback);
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

function collectTags_(e) {
  const params = (e && e.parameter) || {};
  let raw = [];
  if (params.batch) {
    try {
      const parsed = JSON.parse(params.batch);
      if (Array.isArray(parsed)) raw = parsed;
    } catch (error) {
      raw = [];
    }
  }
  if (e && e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      if (body && Array.isArray(body.tags)) raw = raw.concat(body.tags);
    } catch (error) {
      raw = raw;
    }
  }
  if (!raw.length && params.imageId) {
    raw = [{ id: params.tagId || "", imageId: params.imageId, source: params.source }];
  }
  const seen = {};
  const tags = [];
  raw.forEach(function (item) {
    if (!item || !item.imageId) return;
    const tag = {
      id: cleanTagId_(item.id),
      imageId: padId(item.imageId),
      source: normalizeSource(item.source)
    };
    if (seen[tag.id]) return;
    seen[tag.id] = true;
    tags.push(tag);
  });
  return tags;
}

function cleanTagId_(value) {
  const text = String(value || "").replace(/[^a-zA-Z0-9]/g, "").slice(0, 40);
  return text || Utilities.getUuid().replace(/-/g, "");
}

function recordIncomingTags(tags) {
  if (!tags.length) {
    return { ok: true, saved: [], countMode: "tag-log" };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    return { error: true, message: "busy", saved: [], countMode: "tag-log" };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const statsSheet = ss.getSheetByName("통계");
    if (!statsSheet) {
      throw new Error("통계 시트가 없습니다. setupSpreadsheet()를 먼저 실행하세요.");
    }
    ensureLoggedCounts_(ss, statsSheet);
    const log = ensureReceiveSheet_(ss);
    const rows = tags.map(function (tag) {
      return [tag.id, tag.imageId, tag.source, new Date()];
    });
    log.getRange(log.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    return {
      ok: true,
      saved: tags.map(function (tag) { return tag.id; }),
      countMode: "tag-log"
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureReceiveSheet_(ss) {
  let sheet = ss.getSheetByName("수신");
  if (!sheet) {
    sheet = ss.insertSheet("수신");
    sheet.getRange(1, 1, 1, 4).setValues([["태그번호", "이미지", "구분", "시각"]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureLoggedCounts_(ss, statsSheet) {
  const current = String(statsSheet.getRange("C4").getFormula() || "");
  if (current.indexOf("수신") >= 0) {
    ensureSumFormulas(statsSheet);
    return;
  }

  const lastRow = Math.max(statsSheet.getLastRow(), 4);
  const count = lastRow - 3;
  let base = ss.getSheetByName("기준");
  if (!base) base = ss.insertSheet("기준");
  base.getRange("A1").setValue("기존 횟수 보관. 이 시트를 지우면 숫자가 어긋납니다.");
  const snapshot = statsSheet.getRange(4, 1, count, 4).getValues();
  base.getRange(4, 1, count, 4).setValues(snapshot);

  const formulasC = [];
  const formulasD = [];
  for (let i = 0; i < count; i++) {
    const row = 4 + i;
    formulasC.push([
      '=N(기준!C' + row + ')+IF(A' + row + '="","",IFERROR(COUNTUNIQUE(FILTER(수신!$A$2:$A$20000,수신!$B$2:$B$20000=A' + row + ',수신!$C$2:$C$20000="nfc")),0))'
    ]);
    formulasD.push([
      '=N(기준!D' + row + ')+IF(A' + row + '="","",IFERROR(COUNTUNIQUE(FILTER(수신!$A$2:$A$20000,수신!$B$2:$B$20000=A' + row + ',수신!$C$2:$C$20000="web")),0))'
    ]);
  }
  statsSheet.getRange(4, 3, count, 1).setFormulas(formulasC);
  statsSheet.getRange(4, 4, count, 1).setFormulas(formulasD);
  ensureSumFormulas(statsSheet);
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

function normalizeSource(source) {
  return String(source || "").trim() === "nfc" ? "nfc" : "web";
}

/**
 * 전체·연도 합계를 시트 함수로 바꿉니다. 이미지별 2025/2026 숫자는 지우지 않습니다.
 * Apps Script 편집기에서 이 함수 선택 → ▶ 실행
 */
function useSumFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  const statsSheet = ss.getSheetByName("통계");
  if (!statsSheet) {
    throw new Error("통계 시트가 없습니다. setupSpreadsheet()를 먼저 실행하세요.");
  }
  ensureSumFormulas(statsSheet);
  ensureLoggedCounts_(ss, statsSheet);
  SpreadsheetApp.flush();
  ss.toast("기존 횟수는 기준 시트에 두고, 새 태그는 수신 시트에 쌓습니다.", "완료", 8);
}

function addSourceColumns() {
  useSumFormulas();
}

function ensureSumFormulas(statsSheet) {
  migrateAwayFromTotalColumn(statsSheet);

  statsSheet.getRange("A1").setValue("전체 태그 수");
  statsSheet.getRange("A3:D3").setValues([["ID", "제목", "2025년 태그", "2026년 태그"]]);
  statsSheet.getRange("E1:F1").clearContent();
  statsSheet.getRange("E3").clearContent();
  // 합계는 각 열 바로 위. C1=2025합, D1=2026합, B1=둘의 합. SUM(C:C) 는 순환 오류가 납니다.
  statsSheet.getRange("C1").setFormula("=SUM(C4:C100)");
  statsSheet.getRange("D1").setFormula("=SUM(D4:D100)");
  statsSheet.getRange("B1").setFormula("=N(C1)+N(D1)");
  statsSheet.setColumnWidth(3, 110);
  statsSheet.setColumnWidth(4, 130);
}

/**
 * 예전에 있던 이미지 「전체」열(C)을 빼고, 2025→C / 2026→D 로 옮깁니다.
 */
function migrateAwayFromTotalColumn(statsSheet) {
  const headerC = String(statsSheet.getRange("C3").getValue());
  const headerE = String(statsSheet.getRange("E3").getValue());
  const alreadyNew = headerC.indexOf("2025") >= 0 && String(statsSheet.getRange("D3").getValue()).indexOf("2026") >= 0;
  if (alreadyNew) return;

  const oldLayout = headerC.indexOf("전체") >= 0 || headerE.indexOf("2026") >= 0;
  if (!oldLayout) return;

  const lastRow = statsSheet.getLastRow();
  if (lastRow >= 4) {
    const count = lastRow - 3;
    const year2025 = statsSheet.getRange(4, 4, count, 1).getValues();
    const year2026 = statsSheet.getRange(4, 5, count, 1).getValues();
    statsSheet.getRange(4, 3, count, 1).setValues(year2025);
    statsSheet.getRange(4, 4, count, 1).setValues(year2026);
    statsSheet.getRange(4, 5, count, 1).clearContent();
  }
}

function incrementImageCount(statsSheet, imageId, source) {
  const lastRow = statsSheet.getLastRow();
  if (lastRow < 4) return;

  const ids = statsSheet.getRange(4, 1, lastRow - 3, 1).getValues();
  const targetId = padId(imageId);
  const sourceCol = source === "nfc" ? 3 : 4;

  for (let i = 0; i < ids.length; i++) {
    if (padId(ids[i][0]) === targetId) {
      const cell = statsSheet.getRange(i + 4, sourceCol);
      cell.setValue((Number(cell.getValue()) || 0) + 1);
      return;
    }
  }

  const row = lastRow + 1;
  statsSheet.getRange(row, 1).setValue(targetId);
  statsSheet.getRange(row, 2).setValue(targetId === "키링" ? "키링(2025 메시지)" : targetId);
  statsSheet.getRange(row, sourceCol).setValue(1);
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
  stats.getRange("A1").setValue("전체 태그 수");
  stats.getRange("C1").setValue("2025년 태그");
  stats.getRange("E1").setValue("2026년 태그");
  stats.getRange("A3:D3").setValues([["ID", "제목", "2025년 태그", "2026년 태그"]]);

  const statRows = DEFAULT_IMAGES.map((row) => [row[0], row[1], 0, 0]);
  stats.getRange(4, 1, statRows.length, 4).setValues(statRows);
  ensureSumFormulas(stats);
  stats.setFrozenRows(3);
  stats.setColumnWidth(1, 60);
  stats.setColumnWidth(2, 220);
  stats.setColumnWidth(3, 110);
  stats.setColumnWidth(4, 130);

  const defaultSheet = ss.getSheetByName("시트1");
  if (defaultSheet && ss.getSheets().length > 3) {
    ss.deleteSheet(defaultSheet);
  }
}

/**
 * 제목이 깨졌을 때 복구: 이미지 시트 제목 → 통계 시트 B열
 * 이미지별 2025/2026 횟수와 맨 위 합계는 그대로 둡니다.
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
