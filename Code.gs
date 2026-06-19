// =============================================
// 스프레드시트 ID 설정
// =============================================
const SS_BOOTH = '1avD0D72BhxP7aZfM0BO7Y4xc2B7UWrT9U5o5S6fG2tU';
const SS_INQUIRY = '1JJy80Ah9NaJN_BNk21Nyb7P3dtli5OFtR2AfMNgPkb4';
const SS_LINKS = '1ENuXr_ibmRNeUlAGjQLwh3Vs5zLP5hCCNy_hDEgFb-0'; // 계획서/보고서/캔바/상호작용 링크
const SS_PEER     = '1WItERzz5PtYV-T3BGTE7Dfh2N951qWMbn3iKz65T5-w';  // 동료평가
const SS_PASSWORD = '1JcgoufQUypJR6ItEBGWR7xVE-e1vBqXqfYddkEgXlJg';  // 학생 비밀번호
const SS_GRADING  = SS_INQUIRY; // 채점탭이 탐구 스프레드시트에 있음

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

  const page = e && e.parameter && e.parameter.page;
  if (!action) {
    if (page === 'student') {
      return HtmlService.createHtmlOutputFromFile('student')
        .setTitle('내 학습 현황')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
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
    } else if (action === 'verifyStudent') {
      data = verifyStudent(
        e.parameter.ban      || '',
        e.parameter.num      || '',
        e.parameter.name     || '',
        e.parameter.password || ''
      );
    } else if (action === 'getMyData') {
      data = getMyData(
        e.parameter.ban   || '',
        e.parameter.num   || '',
        e.parameter.name  || ''
      );
    } else if (action === 'getPeerStats') {
      data = getPeerStats(
        e.parameter.studentName || '',
        e.parameter.ban         || '',
        e.parameter.modum       || ''
      );
    } else if (action === 'debugGrading') {
      data = debugGradingSheet();
    } else if (action === 'debugPassword') {
      data = debugPasswordSheet();
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

  const countMap = buildResponseCountMap();
  const peerMap  = buildPeerCountMap();

  return filtered.map(s => {
    const key = norm(s.name);
    const counts = countMap[key] || { booth: 0, inquiry: 0, boothTotal: 0, inquiryTotal: 0 };
    const peer   = peerMap[key]  || { given: 0, received: 0 };
    return { ...s, counts, peer };
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

// =============================================
// 동료평가 공통 — raw 시트 로드 + 중복 제거
// (작성자·대상 쌍이 같으면 글자수 가장 긴 행만 유지)
// =============================================
function loadDeduplicatedPeerRows() {
  const CONTENT_START = 5;
  const ss = SpreadsheetApp.openById(SS_PEER);
  const sheet = findSheet(ss, ['data', 'Data', '데이터']);
  if (!sheet) return { header: [], rows: [] };

  const raw    = sheet.getDataRange().getValues();
  const header = raw[0];

  // (작성자norm, 대상norm) → 가장 긴 행
  const best = {};
  for (let i = 1; i < raw.length; i++) {
    const row   = raw[i];
    const wName = norm(row[1] || '');
    const tName = norm(row[3] || '');
    if (!wName || !tName) continue;

    const key  = wName + '||' + tName;
    const text = header.slice(CONTENT_START)
      .map((_, idx) => String(row[CONTENT_START + idx] || '').trim())
      .join('');
    const len = text.length;

    if (!best[key] || len >= best[key].len) {
      best[key] = { row, len };
    }
  }

  return { header, rows: Object.values(best).map(b => b.row), CONTENT_START };
}

// =============================================
// 동료평가 — 칩용 횟수 맵 (중복 제거 적용)
// =============================================
function buildPeerCountMap() {
  const map = {};
  try {
    const { rows } = loadDeduplicatedPeerRows();
    rows.forEach(row => {
      const wName = norm(row[1] || '');
      const tName = norm(row[3] || '');
      if (wName) { if (!map[wName]) map[wName] = { given: 0, received: 0 }; map[wName].given++; }
      if (tName) { if (!map[tName]) map[tName] = { given: 0, received: 0 }; map[tName].received++; }
    });
  } catch (e) {
    Logger.log(`[buildPeerCountMap] 오류: ${e}`);
  }
  return map;
}

// =============================================
// 동료평가 — 상세 통계 (해준/받은 횟수·글자수·순위 + 내용)
// =============================================
function getPeerStats(studentName, ban, modum) {
  try {
    const { header, rows, CONTENT_START } = loadDeduplicatedPeerRows();
    if (!rows.length) return { given: null, received: null };

    const contentHeaders = [];
    for (let c = CONTENT_START; c < header.length; c++) {
      contentHeaders.push(String(header[c]).trim() || `${c + 1}열`);
    }

    const normTarget = norm(studentName);

    const givenCharMap    = {};
    const receivedCharMap = {};
    const givenItems      = [];
    const receivedItems   = [];

    rows.forEach(row => {
      const wName = norm(row[1] || '');
      const tName = norm(row[3] || '');

      const charCount = contentHeaders
        .map((_, idx) => String(row[CONTENT_START + idx] || '').trim())
        .join('').length;

      if (wName) {
        givenCharMap[wName] = (givenCharMap[wName] || 0) + charCount;
        if (wName === normTarget) {
          givenItems.push({
            targetName: String(row[3] || '').trim(),
            targetId:   String(row[4] || '').trim(),
            contents: contentHeaders.map((h, idx) => ({
              q: h, a: String(row[CONTENT_START + idx] || '').trim()
            })).filter(c => c.a)
          });
        }
      }
      if (tName) {
        receivedCharMap[tName] = (receivedCharMap[tName] || 0) + charCount;
        if (tName === normTarget) {
          receivedItems.push({
            writerName: String(row[1] || '').trim(),
            writerId:   String(row[2] || '').trim(),
            contents: contentHeaders.map((h, idx) => ({
              q: h, a: String(row[CONTENT_START + idx] || '').trim()
            })).filter(c => c.a)
          });
        }
      }
    });

    // 학급 구성원 집합 (반 기준)
    const classNames = getClassMemberNorms(ban);

    const schoolGivenTotal    = Object.keys(givenCharMap).length;
    const schoolReceivedTotal = Object.keys(receivedCharMap).length;

    return {
      given: {
        count:       givenItems.length,
        charSum:     givenCharMap[normTarget] || 0,
        schoolRank:  computeRank(givenCharMap, normTarget),
        schoolTotal: schoolGivenTotal,
        classRank:   computeRankFiltered(givenCharMap, normTarget, classNames),
        classTotal:  classNames.size,
        items:       givenItems
      },
      received: {
        count:       receivedItems.length,
        charSum:     receivedCharMap[normTarget] || 0,
        schoolRank:  computeRank(receivedCharMap, normTarget),
        schoolTotal: schoolReceivedTotal,
        classRank:   computeRankFiltered(receivedCharMap, normTarget, classNames),
        classTotal:  classNames.size,
        items:       receivedItems
      }
    };
  } catch (e) {
    Logger.log(`[getPeerStats] 오류: ${e}`);
    return { error: e.toString() };
  }
}

// 동일 반 학생 norm(이름) 집합 반환
function getClassMemberNorms(ban) {
  const names = new Set();
  if (!ban) return names;
  [SS_BOOTH, SS_INQUIRY].forEach(id => {
    try {
      const ss    = SpreadsheetApp.openById(id);
      const sheet = findSheet(ss, [STUDENT_INFO_SHEET]);
      if (!sheet) return;
      const data   = sheet.getDataRange().getValues();
      const header = data[0].map(h => String(h).trim());
      const banCol  = findColIndex(header, ['반', '학반', '학년반', '학년/반']);
      const nameCol = findColIndex(header, ['이름', '학생이름', '성명', '학생 이름']);
      for (let i = 1; i < data.length; i++) {
        const rBan  = banCol  >= 0 ? String(data[i][banCol]).trim()  : '';
        const rName = nameCol >= 0 ? String(data[i][nameCol]).trim() : '';
        if (rName && rBan === ban) names.add(norm(rName));
      }
    } catch (e) {}
  });
  return names;
}

// 전체 순위 (값 높을수록 1위)
function computeRank(charMap, targetName) {
  const myVal = charMap[targetName] || 0;
  return Object.values(charMap).filter(v => v > myVal).length + 1;
}

// 특정 이름 집합 내 순위
function computeRankFiltered(charMap, targetName, nameSet) {
  if (!nameSet.size) return null;
  const myVal = charMap[targetName] || 0;
  let rank = 1;
  nameSet.forEach(n => {
    if (n !== targetName && (charMap[n] || 0) > myVal) rank++;
  });
  return rank;
}

// =============================================
// 비밀번호 시트 디버그
// =============================================
function debugGradingSheet() {
  try {
    const ss    = SpreadsheetApp.openById(SS_GRADING);
    const sheet = findSheet(ss, ['채점', '점수', 'Score']);
    if (!sheet) return { error: '채점 시트 없음', sheets: ss.getSheets().map(s=>s.getName()) };
    const data = sheet.getDataRange().getValues();
    return {
      header:    data[0],
      row2:      data[1] || [],
      row3:      data[2] || [],
      totalRows: data.length
    };
  } catch(e) { return { error: e.toString() }; }
}

function debugPasswordSheet() {
  try {
    const ss    = SpreadsheetApp.openById(SS_PASSWORD);
    const sheets = ss.getSheets().map(s => s.getName());
    const sheet  = ss.getSheets()[0];
    const data   = sheet.getDataRange().getValues();
    return {
      sheets,
      header: data[0],
      row2:   data[1] || [],
      row3:   data[2] || [],
      totalRows: data.length
    };
  } catch(e) { return { error: e.toString() }; }
}

// =============================================
// 학생 포털 — 비밀번호 검증
// =============================================
function verifyStudent(ban, num, name, password) {
  try {
    const ss    = SpreadsheetApp.openById(SS_PASSWORD);
    const sheet = ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    const header = data[0].map(h => String(h).trim());

    // 학번(A=0), 성명(B=1), 비밀번호(C=2) 구조
    const idCol   = findColIndex(header, ['학번', '번호']);
    const nameCol = findColIndex(header, ['성명', '이름', '학생이름']);
    const pwCol   = findColIndex(header, ['비밀번호', '개인 비밀번호', '패스워드', 'password']);

    const inputBan = parseInt(ban,  10);
    const inputNum = parseInt(num,  10);

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const rId    = idCol   >= 0 ? parseInt(row[idCol],   10) : 0;
      const rName  = nameCol >= 0 ? norm(row[nameCol])         : '';
      const rPw    = pwCol   >= 0 ? String(row[pwCol]).trim()  : '';

      // 학번 파싱: 학년(1자리) + 반(1자리) + 번호(2자리) = 4자리
      // 예: 3101 → 학년=3, 반=1, 번호=01
      const rBan = Math.floor((rId % 1000) / 100);
      const rNum = rId % 100;

      const nameMatch = rName === norm(name);
      const banMatch  = rBan  === inputBan;
      const numMatch  = rNum  === inputNum;

      if (nameMatch && banMatch && numMatch) {
        if (rPw === password.trim()) return { ok: true };
        else return { ok: false, reason: '비밀번호가 일치하지 않습니다.' };
      }
    }
    return { ok: false, reason: '학생 정보를 찾을 수 없습니다.' };
  } catch (e) {
    return { ok: false, reason: '오류: ' + e.toString() };
  }
}

// =============================================
// 학생 포털 — 내 데이터 통합 조회
// (성찰일지 + 링크 + 동료평가 익명 + 점수)
// =============================================
function getMyData(ban, num, name) {
  // 학생 정보에서 모둠 찾기
  let modum = '';
  [SS_BOOTH, SS_INQUIRY].forEach(id => {
    if (modum) return;
    try {
      const ss    = SpreadsheetApp.openById(id);
      const sheet = findSheet(ss, [STUDENT_INFO_SHEET]);
      if (!sheet) return;
      const data   = sheet.getDataRange().getValues();
      const header = data[0].map(h => String(h).trim());
      const banCol   = findColIndex(header, ['반', '학반', '학년반']);
      const numCol   = findColIndex(header, ['번호', '학번', '출석번호']);
      const nameCol  = findColIndex(header, ['이름', '성명', '학생이름']);
      const modumCol = findColIndex(header, ['모둠', '모둠번호', '그룹']);
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (norm(row[nameCol]) === norm(name) &&
            (!ban || norm(row[banCol]) === norm(ban))) {
          modum = modumCol >= 0 ? String(row[modumCol]).trim() : '';
          break;
        }
      }
    } catch (e) {}
  });

  // 성찰일지
  const reflections = getStudentReflections(name, ban, modum);

  // 링크
  const linksData = getStudentLinks(name, ban, modum);

  // 동료평가 (받은 것 익명 처리)
  const peerRaw = getPeerStats(name, ban, modum);
  const peer = peerRaw;
  if (peer && peer.received && peer.received.items) {
    peer.received.items = peer.received.items.map(item => ({
      ...item,
      writerName: '익명',
      writerId:   ''
    }));
  }

  // 점수
  const scores = getStudentScore(ban, num, name);

  return { reflections, linksData, peer, scores, modum };
}

// =============================================
// 학생 포털 — 채점탭에서 점수 조회
// =============================================
function getStudentScore(ban, num, name) {
  try {
    const ss    = SpreadsheetApp.openById(SS_GRADING);
    const sheet = findSheet(ss, ['채점', '점수', 'Score', 'score']);
    if (!sheet) return null;

    const data   = sheet.getDataRange().getValues();
    const header = data[0].map(h => String(h).trim());

    // 헤더: ["","학번","반","모둠","주제","실험","이름","계획서(3점)","보고서(4점)","부스 준비(3점)","동료평가(3점)","탐구성찰(6점)","부스성찰(6점)","합"]
    const banCol  = findColIndex(header, ['반']);
    const nameCol = findColIndex(header, ['이름', '성명', '학생이름']);
    const idCol   = findColIndex(header, ['학번']);

    const scoreKeys = {
      plan:    findColIndex(header, ['계획서']),
      report:  findColIndex(header, ['보고서']),
      booth:   findColIndex(header, ['부스 준비', '부스준비', '캔바']),
      peer:    findColIndex(header, ['동료평가', '동료 평가']),
      inquiry: findColIndex(header, ['탐구성찰', '탐구 성찰']),
      booth2:  findColIndex(header, ['부스성찰', '부스 성찰']),
      total:   findColIndex(header, ['합계', '합', '총점', '최종']),
    };

    const inputBan = parseInt(ban, 10);
    const inputNum = parseInt(num, 10);

    for (let i = 1; i < data.length; i++) {
      const row   = data[i];
      const rName = nameCol >= 0 ? norm(row[nameCol]) : '';
      const rBan  = banCol  >= 0 ? parseInt(row[banCol], 10) : 0;
      const rId   = idCol   >= 0 ? parseInt(row[idCol],  10) : 0;
      const rNum  = rId % 100; // 학번 끝 2자리 = 번호

      const nameMatch = rName === norm(name);
      const banMatch  = !ban || rBan === inputBan;
      const numMatch  = !num || rNum === inputNum;

      if (nameMatch && banMatch && numMatch) {
        const get = col => {
          if (col < 0 || col >= row.length) return null;
          const v = row[col];
          if (v === '' || v === null || v === undefined) return null;
          const n = Number(v);
          return isNaN(n) ? null : n;
        };
        return {
          plan:    get(scoreKeys.plan),
          report:  get(scoreKeys.report),
          booth:   get(scoreKeys.booth),
          peer:    get(scoreKeys.peer),
          inquiry: get(scoreKeys.inquiry),
          booth2:  get(scoreKeys.booth2),
          total:   get(scoreKeys.total),
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log(`[getStudentScore] 오류: ${e}`);
    return null;
  }
}
