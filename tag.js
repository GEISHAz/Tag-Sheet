/***** 설정값 *****/
const DEFAULT_WEEK_HEIGHT = 20;  // 템플릿 기본 높이
const WEEK_WIDTH = 21;           // 한 주 열(A~U)
const DAY_BLOCK = 3;             // 하루 열 수
const NUM_DAYS = 6;              // 월~토

/***** 주차 메타 읽기 *****/
function getWeekHeights_(sheet) {
  const result = [];
  let row = 1;
  while (row <= sheet.getMaxRows()) {
    const val = sheet.getRange(row, 1).getValue();
    if (!val) break;

    // ★ 변경점: H=값 읽는 위치를 A열 두 번째 행에서 T열 첫 번째 행으로 변경 (1열 -> 20열)
    const meta = sheet.getRange(row, 20).getValue();
    const match = (meta || '').toString().match(/H=(\d+)/);
    const h = match ? parseInt(match[1], 10) : DEFAULT_WEEK_HEIGHT;

    result.push({ start: row, height: h });
    row += h;
  }
  return result;
}

/***** 날짜 읽기 *****/
function readMondayDate_(sheet, weekIndex) {
  const heights = getWeekHeights_(sheet);
  const row = heights[weekIndex]?.start || 1;
  const raw = sheet.getRange(row, 1).getValue();
  if (!raw) throw new Error("A" + row + " 셀에 날짜가 없습니다.");

  if (raw instanceof Date) return raw;

  const norm = raw.toString().replace(/[^\d]/g, '.').replace(/\.+/g, '.')
                      .replace(/^\./, '').replace(/\.$/, '');
  const parts = norm.split('.').map(s => s.trim()).filter(Boolean);
  const now = new Date();
  if (parts.length >= 3) return new Date(+parts[0], +parts[1]-1, +parts[2]);
  if (parts.length === 2) return new Date(now.getFullYear(), +parts[0]-1, +parts[1]);
  if (parts.length === 1) return new Date(now.getFullYear(), now.getMonth(), +parts[0]);
  throw new Error("날짜 파싱 실패: " + raw);
}

/***** 요일 헤더 작성 *****/
function setWeekHeader_(sheet, startRow, mondayDate) {
  const days = ['월요일','화요일','수요일','목요일','금요일','토요일'];
  for (let i = 0; i < NUM_DAYS; i++) {
    const col = 1 + i * DAY_BLOCK;
    const d = new Date(mondayDate.getTime());
    d.setDate(d.getDate() + i);
    sheet.getRange(startRow, col).setValue(d);
    sheet.getRange(startRow + 1, col).setValue(days[i]);
  }
  // ★ 변경점: H=값 저장 위치를 T열 첫 번째 행으로 변경
  sheet.getRange(startRow, 20).setValue('H=' + DEFAULT_WEEK_HEIGHT);
}

/***** 다음 주 생성 *****/
function createWeek() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setting = ss.getSheetByName("Setting");
  const sheet = ss.getSheetByName("태그일정");
  if (!setting || !sheet) { SpreadsheetApp.getUi().alert("Setting/태그일정 시트 확인"); return; }

  const templateRange = setting.getRange(1, 1, DEFAULT_WEEK_HEIGHT, WEEK_WIDTH);
  const weekInfos = getWeekHeights_(sheet);
  const weekCount = weekInfos.length;
  const newStartRow = weekInfos.reduce((acc, w) => acc + w.height, 1);

  const needLast = newStartRow + DEFAULT_WEEK_HEIGHT - 1;
  if (sheet.getMaxRows() < needLast) sheet.insertRowsAfter(sheet.getMaxRows(), needLast - sheet.getMaxRows());

  templateRange.copyTo(sheet.getRange(newStartRow, 1), { contentsOnly: false });

  let newMonday;
  if (weekCount === 0) {
    newMonday = readMondayDate_(setting, 0);
  } else {
    const baseMonday = readMondayDate_(sheet, 0);
    newMonday = new Date(baseMonday.getTime());
    newMonday.setDate(baseMonday.getDate() + 7 * weekCount);
  }
  setWeekHeader_(sheet, newStartRow, newMonday);
}

/***** 저번 주 삭제 (A~U만) *****/
function deleteWeek() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("태그일정");
  const archive = ss.getSheetByName("삭제된일정") || ss.insertSheet("삭제된일정");
  if (!sheet) { SpreadsheetApp.getUi().alert("태그일정 시트 없음"); return; }

  const weekInfos = getWeekHeights_(sheet);
  if (weekInfos.length === 0) {
    SpreadsheetApp.getUi().alert("삭제할 주가 없습니다.");
    return;
  }

  const first = weekInfos[0];
  const destRow = archive.getLastRow() + 1;

  sheet.getRange(first.start, 1, first.height, WEEK_WIDTH)
       .copyTo(archive.getRange(destRow, 1), { contentsOnly: false });

  // A~U 열만 아래 주차 내용으로 덮어쓰기 → 자연스러운 '위로 당김'
  const below = sheet.getLastRow() - (first.start + first.height - 1);
  if (below > 0) {
    sheet.getRange(first.start + first.height, 1, below, WEEK_WIDTH)
         .copyTo(sheet.getRange(first.start, 1), { contentsOnly: false });
    sheet.getRange(first.start + below, 1, first.height, WEEK_WIDTH).clear();
  } else {
    sheet.getRange(first.start, 1, first.height, WEEK_WIDTH).clear();
  }
}

/***** 현재 주 찾기 *****/
function findWeekByActiveCell_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== "태그일정") return null;
  const row = sheet.getActiveRange().getRow();
  const infos = getWeekHeights_(sheet);
  for (const w of infos) {
    if (row >= w.start && row < w.start + w.height) return w;
  }
  return null;
}


function addRowsToCurrentWeek() {
  const ui = SpreadsheetApp.getUi();
  const week = findWeekByActiveCell_();
  if (!week) { ui.alert("현재 주를 찾을 수 없습니다."); return; }
  const numStr = ui.prompt("추가할 단위를 입력 (1 입력 시 3줄 추가)", ui.ButtonSet.OK_CANCEL).getResponseText();
  const num = parseInt(numStr, 10);
  if (isNaN(num) || num < 1) return;

  const rowsToAdd = num * 3;
  const sheet = SpreadsheetApp.getActiveSheet();

  // 현재 주의 끝 바로 다음 행
  const start = week.start + week.height;

  // 1) 바닥에 여유공간 확보 (이건 시트 전체 행 추가 - 맨 아래에 추가하므로 태그영역 영향 없음)
  const originalMax = sheet.getMaxRows();
  if (originalMax < start) {
    // 드물게 start가 현재 MaxRows를 넘는 경우 (안전장치)
    sheet.insertRowsAfter(originalMax, start - originalMax);
  }
  sheet.insertRowsAfter(originalMax, rowsToAdd); // 항상 충분한 여유 확보

  // 2) A~U (WEEK_WIDTH) 블록을 아래로 이동 (T열은 이 블록 안에 포함되어 있음)
  const rowsBelow = originalMax - start + 1; // 원래 존재하던 아래 블록의 행수
  if (rowsBelow > 0) {
    sheet.getRange(start, 1, rowsBelow, WEEK_WIDTH)
         .moveTo(sheet.getRange(start + rowsToAdd, 1));
  }

  // 3) 새로 생긴 구간(A~U) 내용은 비움(내용만 비움)
  sheet.getRange(start, 1, rowsToAdd, WEEK_WIDTH).clearContent();

  // 4) 메타(H= 값) 갱신: 현재 주의 T열(20번 열)에 높이 반영
  sheet.getRange(week.start, 20).setValue('H=' + (week.height + rowsToAdd));

}
function removeRowsFromCurrentWeek() {
  const ui = SpreadsheetApp.getUi();
  const week = findWeekByActiveCell_();
  if (!week) {
    ui.alert("현재 주를 찾을 수 없습니다.");
    return;
  }

  const num = parseInt(
    ui.prompt("제거할 행 수를 입력", ui.ButtonSet.OK_CANCEL).getResponseText(),
    10
  );
  if (!num || num < 1 || num >= week.height - 2) {
    ui.alert("제거 수가 유효하지 않거나 너무 큽니다.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();

  /**
   * 1) 위쪽의 데이터를 아래로 밀지 않고
   *    지정 구간만 ‘깨끗이’ 지웁니다.
   *    A~U(21열 = U) 영역만 모두 삭제.
   */
  const START_COL = 1;     // A
  const COLS      = 21;    // A~U
  const startRow  = week.start + week.height - num;
  const clearRange = sheet.getRange(startRow, START_COL, num, COLS);
  clearRange.clear();      // 내용 + 배경색 + 테두리 + 조건부서식 모두 삭제

  /**
   * 2) H=값을 T열(20번째 열)로 기록
   */
  sheet.getRange(week.start, 20).setValue('H=' + (week.height - num));

}



/***** 메뉴 등록 *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("주차도구")
    .addItem("다음 주 생성", "createWeek")
    .addItem("저번 주 삭제", "deleteWeek")
    .addSeparator()
    .addItem("현재 주 행 추가", "addRowsToCurrentWeek")
    .addItem("현재 주 행 제거", "removeRowsFromCurrentWeek")
    .addToUi();
}