// =============================================
// 스프레드시트 ID 설정
// =============================================
const SS_BOOTH = '1avD0D72BhxP7aZfM0BO7Y4xc2B7UWrT9U5o5S6fG2tU'; // 부스 운영 및 과학 전시 최종 성찰일지
const SS_INQUIRY = '1JJy80Ah9NaJN_BNk21Nyb7P3dtli5OFtR2AfMNgPkb4'; // 탐구 과정 성찰일지

const STUDENT_INFO_SHEET = '학생 정보';
const RESPONSE_SHEET = '설문지 응답 시트1';

// K~P열 = 인덱스 10~15 (0-based)
const RESPONSE_COL_START = 10;
const RESPONSE_COL_END = 15;

// =============================================
// 웹앱 진입점
// =============================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('성찰일지 뷰어')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================
// 학생 정보 시트에서 전체 학생 목록 로드
// (반, 모둠, 이름 + 고유키 포함)
// 두 스프레드시트 모두에서 수집
// =============================================
function getAllStudents() {
  const result = [];

  [
    { id: SS_BOOTH, label: '부스 운영' },
    { id: SS_INQUIRY, label: '탐구 과정' }
  ].forEach(({ id, label }) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sheet = ss.getSheetByName(STUDENT_INFO_SHEET);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      const header = data[0].map(h => String(h).trim());

      // 컬럼 인덱스 자동 탐지 (유연하게 처리)
      const col = {
        ban:   findColIndex(header, ['반', '학반', '학년반', '학년/반']),
        modum: findColIndex(header, ['모둠', '모둠번호', '그룹']),
        name:  findColIndex(header, ['이름', '학생이름', '성명']),
        num:   findColIndex(header, ['번호', '학번', '출석번호']),
      };

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const name = col.name >= 0 ? String(row[col.name]).trim() : '';
        if (!name) continue;

        result.push({
          source: label,
          sourceId: id,
          ban:   col.ban   >= 0 ? String(row[col.ban]).trim()   : '',
          modum: col.modum >= 0 ? String(row[col.modum]).trim() : '',
          num:   col.num   >= 0 ? String(row[col.num]).trim()   : '',
          name:  name,
          rowIndex: i
        });
      }
    } catch (e) {
      Logger.log(`[getAllStudents] ${label} 오류: ${e}`);
    }
  });

  return result;
}

// =============================================
// 필터 조건으로 학생 목록 반환
// =============================================
function getFilteredStudents(ban, modum, name) {
  const all = getAllStudents();
  return all.filter(s => {
    if (ban   && s.ban   !== ban)   return false;
    if (modum && s.modum !== modum) return false;
    if (name  && !s.name.includes(name)) return false;
    return true;
  });
}

// =============================================
// 필터 선택지 목록 반환 (반, 모둠 드롭다운용)
// =============================================
function getFilterOptions() {
  const all = getAllStudents();
  const bans   = [...new Set(all.map(s => s.ban).filter(Boolean))].sort(naturalSort);
  const modums = [...new Set(all.map(s => s.modum).filter(Boolean))].sort(naturalSort);
  return { bans, modums };
}

// =============================================
// 특정 학생의 성찰일지 응답 (K~P열) 반환
// 두 스프레드시트 모두 검색
// =============================================
function getStudentReflections(studentName, ban, modum) {
  const reflections = [];

  [
    { id: SS_BOOTH, label: '부스 운영 및 과학 전시 최종 성찰일지' },
    { id: SS_INQUIRY, label: '탐구 과정 성찰일지' }
  ].forEach(({ id, label }) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sheet = ss.getSheetByName(RESPONSE_SHEET);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      const header = data[0].map(h => String(h).trim());

      // 응답 시트에서 이름 컬럼 탐지
      const nameCol = findColIndex(header, ['이름', '학생이름', '성명', '이름을 입력하세요', '이름 입력']);
      const banCol  = findColIndex(header, ['반', '학반', '학년반', '학년/반', '학반을 입력']);
      const modumCol = findColIndex(header, ['모둠', '모둠번호', '그룹']);

      // K~P열 헤더 (질문 제목)
      const questionHeaders = [];
      for (let c = RESPONSE_COL_START; c <= RESPONSE_COL_END; c++) {
        questionHeaders.push(header[c] || `K${c - RESPONSE_COL_START + 1}열`);
      }

      // 해당 학생 행 필터
      const matchedRows = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rName  = nameCol  >= 0 ? String(row[nameCol]).trim()  : '';
        const rBan   = banCol   >= 0 ? String(row[banCol]).trim()   : '';
        const rModum = modumCol >= 0 ? String(row[modumCol]).trim() : '';

        const nameMatch  = rName === studentName;
        const banMatch   = !ban   || rBan === ban;
        const modumMatch = !modum || rModum === modum;

        if (nameMatch && banMatch && modumMatch) {
          const answers = [];
          for (let c = RESPONSE_COL_START; c <= RESPONSE_COL_END; c++) {
            answers.push(String(row[c] || '').trim());
          }
          matchedRows.push({
            timestamp: String(row[0] || ''),
            answers: answers
          });
        }
      }

      if (matchedRows.length > 0) {
        reflections.push({
          source: label,
          questions: questionHeaders,
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
function findColIndex(header, candidates) {
  for (const cand of candidates) {
    const idx = header.findIndex(h => h.includes(cand));
    if (idx >= 0) return idx;
  }
  return -1;
}

function naturalSort(a, b) {
  return a.localeCompare(b, 'ko', { numeric: true });
}
