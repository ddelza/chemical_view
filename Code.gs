// =============================================
// 스프레드시트 ID 설정
// =============================================
const SS_BOOTH = '1avD0D72BhxP7aZfM0BO7Y4xc2B7UWrT9U5o5S6fG2tU';
const SS_INQUIRY = '1JJy80Ah9NaJN_BNk21Nyb7P3dtli5OFtR2AfMNgPkb4';

const STUDENT_INFO_SHEET = '학생 정보';
const RESPONSE_SHEET = '정리';

// K~P열 = 인덱스 10~15 (0-based)
const RESPONSE_COL_START = 10;
const RESPONSE_COL_END = 15;

// =============================================
// 웹앱 진입점 — HTML 서빙 + JSON API 겸용
// GitHub Pages 등 외부에서 ?action=xxx 로 호출 가능
// =============================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (!action) {
    // 기본: HTML 반환
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('성찰일지 뷰어')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // JSON API 모드
  let data;
  try {
    if (action === 'getFilterOptions') {
      data = getFilterOptions();
    } else if (action === 'getFilteredStudents') {
      data = getFilteredStudents(
        e.parameter.ban   || '',
        e.parameter.modum || '',
        e.parameter.name  || ''
      );
    } else if (action === 'getStudentReflections') {
      data = getStudentReflections(
        e.parameter.studentName || '',
        e.parameter.ban         || '',
        e.parameter.modum       || ''
      );
    } else if (action === 'debug') {
      data = debugSheetInfo();
    } else {
      data = { error: 'unknown action' };
    }
  } catch (err) {
    data = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================
// 디버그: 시트 이름과 헤더 확인용
// ?action=debug 로 호출
// =============================================
function debugSheetInfo() {
  const result = {};
  [
    { id: SS_BOOTH, label: '부스운영' },
    { id: SS_INQUIRY, label: '탐구과정' }
  ].forEach(({ id, label }) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sheets = ss.getSheets().map(s => s.getName());
      result[label] = { sheets };

      const infoSheet = ss.getSheetByName(STUDENT_INFO_SHEET);
      if (infoSheet) {
        result[label].infoHeader = infoSheet.getRange(1, 1, 1, infoSheet.getLastColumn()).getValues()[0];
        result[label].infoRow2   = infoSheet.getRange(2, 1, 1, infoSheet.getLastColumn()).getValues()[0];
      } else {
        result[label].infoSheetMissing = true;
      }

      const respSheet = ss.getSheetByName(RESPONSE_SHEET);
      if (respSheet) {
        result[label].respHeader = respSheet.getRange(1, 1, 1, respSheet.getLastColumn()).getValues()[0];
        result[label].respRow2   = respSheet.getRange(2, 1, 1, respSheet.getLastColumn()).getValues()[0];
      } else {
        result[label].respSheetMissing = true;
      }
    } catch (e) {
      result[label] = { error: e.toString() };
    }
  });
  return result;
}

// =============================================
// 학생 정보 시트에서 전체 학생 목록 로드
// =============================================
function getAllStudents() {
  const result = [];

  [
    { id: SS_BOOTH, label: '부스 운영' },
    { id: SS_INQUIRY, label: '탐구 과정' }
  ].forEach(({ id, label }) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sheet = findSheet(ss, [STUDENT_INFO_SHEET, '학생정보', '학생 명단', '명단']);
      if (!sheet) { Logger.log(`[${label}] 학생 정보 시트 없음`); return; }

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      const header = data[0].map(h => String(h).trim());
      Logger.log(`[${label}] 학생정보 헤더: ${JSON.stringify(header)}`);

      const col = {
        ban:   findColIndex(header, ['반', '학반', '학년반', '학년/반', '학년-반']),
        modum: findColIndex(header, ['모둠', '모둠번호', '그룹', '팀']),
        name:  findColIndex(header, ['이름', '학생이름', '성명', '학생 이름']),
        num:   findColIndex(header, ['번호', '학번', '출석번호', '학생번호']),
      };
      Logger.log(`[${label}] 컬럼 인덱스: ${JSON.stringify(col)}`);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const name = col.name >= 0 ? norm(row[col.name]) : '';
        if (!name) continue;

        result.push({
          source: label,
          sourceId: id,
          ban:   col.ban   >= 0 ? String(row[col.ban]).trim()   : '',
          modum: col.modum >= 0 ? String(row[col.modum]).trim() : '',
          num:   col.num   >= 0 ? String(row[col.num]).trim()   : '',
          name:  col.name  >= 0 ? String(row[col.name]).trim()  : '',
        });
      }
    } catch (e) {
      Logger.log(`[getAllStudents] ${label} 오류: ${e}`);
    }
  });

  Logger.log(`[getAllStudents] 총 ${result.length}명`);
  return result;
}

// =============================================
// 필터 조건으로 학생 목록 반환
// =============================================
function getFilteredStudents(ban, modum, name) {
  const all = getAllStudents();
  return all.filter(s => {
    if (ban   && s.ban   !== ban)        return false;
    if (modum && s.modum !== modum)      return false;
    if (name  && !s.name.includes(name)) return false;
    return true;
  });
}

// =============================================
// 필터 선택지 목록 반환
// =============================================
function getFilterOptions() {
  const all = getAllStudents();
  const bans   = [...new Set(all.map(s => s.ban).filter(Boolean))].sort(naturalSort);
  const modums = [...new Set(all.map(s => s.modum).filter(Boolean))].sort(naturalSort);
  return { bans, modums };
}

// =============================================
// 특정 학생의 성찰일지 응답 (K~P열) 반환
// =============================================
function getStudentReflections(studentName, ban, modum) {
  const reflections = [];

  [
    { id: SS_BOOTH, label: '부스 운영 및 과학 전시 최종 성찰일지' },
    { id: SS_INQUIRY, label: '탐구 과정 성찰일지' }
  ].forEach(({ id, label }) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sheet = findSheet(ss, [RESPONSE_SHEET, '정리', '설문지 응답 시트1', '설문지 응답 시트 1', '응답시트1', 'Form Responses 1']);
      if (!sheet) { Logger.log(`[${label}] 응답 시트 없음`); return; }

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      const header = data[0].map(h => String(h).trim());

      const nameCol  = findColIndex(header, ['이름', '학생이름', '성명', '학생 이름', '이름을 입력', '이름 입력']);
      const banCol   = findColIndex(header, ['반', '학반', '학년반', '학년/반', '반을 입력', '학반 입력']);
      const modumCol = findColIndex(header, ['모둠', '모둠번호', '그룹', '팀']);

      // K~P열 중 헤더가 실제로 있는 열만 사용
      const validCols = [];
      for (let c = RESPONSE_COL_START; c <= RESPONSE_COL_END; c++) {
        const h = c < header.length ? header[c].trim() : '';
        if (h) validCols.push({ colIndex: c, question: h });
      }
      if (validCols.length === 0) {
        for (let c = RESPONSE_COL_START; c <= RESPONSE_COL_END; c++) {
          validCols.push({ colIndex: c, question: `${c + 1}열` });
        }
      }

      const matchedRows = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rName  = nameCol  >= 0 ? String(row[nameCol]).trim()  : '';
        const rBan   = banCol   >= 0 ? String(row[banCol]).trim()   : '';
        const rModum = modumCol >= 0 ? String(row[modumCol]).trim() : '';

        const nameMatch  = norm(rName) === norm(studentName);
        const banMatch   = !ban   || norm(rBan)   === norm(ban);
        const modumMatch = !modum || norm(rModum) === norm(modum);

        if (nameMatch && banMatch && modumMatch) {
          const answers = validCols.map(({ colIndex: c }) =>
            c < row.length ? String(row[c] || '').trim() : ''
          );
          matchedRows.push({ timestamp: String(row[0] || ''), answers });
        }
      }

      if (matchedRows.length > 0) {
        reflections.push({
          source: label,
          questions: validCols.map(v => v.question),
          responses: matchedRows
        });
      }
    } catch (e) {
      Logger.log(`[getStudentReflections] ${label} 오류: ${e}`);
    }
  });

  return reflections;
}

// =============================================
// 유틸리티
// =============================================
// 공백·전각공백·제로폭 문자 제거 후 NFC 정규화
function norm(s) {
  return String(s).replace(/[\s ​　]/g, '').normalize('NFC');
}

function findSheet(ss, candidates) {
  for (const name of candidates) {
    const s = ss.getSheetByName(name);
    if (s) return s;
  }
  // 부분 매칭
  const all = ss.getSheets();
  for (const cand of candidates) {
    const found = all.find(s => s.getName().includes(cand.replace(/\s/g, '')));
    if (found) return found;
  }
  return null;
}

function findColIndex(header, candidates) {
  for (const cand of candidates) {
    const idx = header.findIndex(h => h.replace(/\s/g, '').includes(cand.replace(/\s/g, '')));
    if (idx >= 0) return idx;
  }
  return -1;
}

function naturalSort(a, b) {
  return a.localeCompare(b, 'ko', { numeric: true });
}
