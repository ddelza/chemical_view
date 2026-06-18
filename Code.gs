// =============================================
// 스프레드시트 ID 설정
// =============================================
const SS_BOOTH = '1avD0D72BhxP7aZfM0BO7Y4xc2B7UWrT9U5o5S6fG2tU';
const SS_INQUIRY = '1JJy80Ah9NaJN_BNk21Nyb7P3dtli5OFtR2AfMNgPkb4';
const SS_LINKS = '1ENuXr_ibmRNeUlAGjQLwh3Vs5zLP5hCCNy_hDEgFb-0'; // 계획서/보고서/캔바/상호작용 링크

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
    } else if (action === 'getStudentLinks') {
      data = getStudentLinks(
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
// 필터 조건으로 학생 목록 반환 (응답 수 포함)
// =============================================
function getFilteredStudents(ban, modum, name) {
  const all = getAllStudents();
  const filtered = all.filter(s => {
    if (ban   && s.ban   !== ban)        return false;
    if (modum && s.modum !== modum)      return false;
    if (name  && !s.name.includes(name)) return false;
    return true;
  });

  // 두 스프레드시트의 정리탭 응답 데이터를 미리 로드
  const countMap = buildResponseCountMap();

  return filtered.map(s => {
    const key = norm(s.name);
    const counts = countMap[key] || { booth: 0, inquiry: 0, boothTotal: 0, inquiryTotal: 0 };
    return { ...s, counts };
  });
}

// 이름 → {booth응답수, inquiry응답수} 맵 생성
function buildResponseCountMap() {
  const map = {};

  const sources = [
    { id: SS_BOOTH,   key: 'booth'   },
    { id: SS_INQUIRY, key: 'inquiry' }
  ];

  sources.forEach(({ id, key }) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sheet = findSheet(ss, [RESPONSE_SHEET, '정리', '설문지 응답 시트1', 'Form Responses 1']);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      const header = data[0].map(h => String(h).trim());
      const nameCol = findColIndex(header, ['이름', '학생이름', '성명', '학생 이름', '이름을 입력', '이름 입력']);

      // 유효 질문 열 수 계산
      const validCols = [];
      for (let c = RESPONSE_COL_START; c <= RESPONSE_COL_END; c++) {
        const h = c < header.length ? header[c].trim() : '';
        if (h) validCols.push(c);
      }
      const total = validCols.length || (RESPONSE_COL_END - RESPONSE_COL_START + 1);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rName = nameCol >= 0 ? norm(row[nameCol]) : '';
        if (!rName) continue;

        // 답변 있는 문항 수 (비어있지 않은 validCols)
        const answered = validCols.filter(c => c < row.length && String(row[c]).trim() !== '').length;

        if (!map[rName]) map[rName] = { booth: 0, inquiry: 0, boothTotal: total, inquiryTotal: total };
        map[rName][`${key}Total`] = total;
        // 여러 행 있으면 최댓값 사용
        if (answered > (map[rName][key] || 0)) map[rName][key] = answered;
      }
    } catch (e) {
      Logger.log(`[buildResponseCountMap] ${key} 오류: ${e}`);
    }
  });

  return map;
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
// 계획서·보고서·캔바·상호작용 링크 조회 (시트1, K~N열)
// =============================================
function getStudentLinks(studentName, ban, modum) {
  const LINK_LABELS = ['계획서', '보고서', '캔바', '1,2차 상호작용'];
  // K~N = 인덱스 10~13
  const LINK_COL_START = 10;
  const LINK_COL_END   = 13;

  try {
    const ss = SpreadsheetApp.openById(SS_LINKS);
    const sheet = findSheet(ss, ['시트1', 'Sheet1']);
    if (!sheet) return { error: '시트1 없음' };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { links: [] };

    const header = data[0].map(h => String(h).trim());
    const nameCol  = findColIndex(header, ['이름', '학생이름', '성명', '학생 이름']);
    const banCol   = findColIndex(header, ['반', '학반', '학년반', '학년/반']);
    const modumCol = findColIndex(header, ['모둠', '모둠번호', '그룹', '팀']);

    // 헤더에 실제 라벨이 있으면 우선 사용
    const labels = [];
    for (let c = LINK_COL_START; c <= LINK_COL_END; c++) {
      labels.push(c < header.length && header[c] ? header[c] : LINK_LABELS[c - LINK_COL_START] || `${c+1}열`);
    }

    // 모둠별 매칭 (반 + 모둠으로 행 찾기)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rBan   = banCol   >= 0 ? norm(row[banCol])   : '';
      const rModum = modumCol >= 0 ? norm(row[modumCol]) : '';

      const banMatch   = !ban   || rBan   === norm(ban);
      const modumMatch = !modum || rModum === norm(modum);
      if (!banMatch || !modumMatch) continue;

      const links = [];
      for (let c = LINK_COL_START; c <= LINK_COL_END; c++) {
        const val = c < row.length ? String(row[c] || '').trim() : '';
        links.push({ label: labels[c - LINK_COL_START], url: val });
      }
      return { links };
    }
    return { links: [] };
  } catch (e) {
    Logger.log(`[getStudentLinks] 오류: ${e}`);
    return { error: e.toString() };
  }
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
