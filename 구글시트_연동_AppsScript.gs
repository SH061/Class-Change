// ===== 수업교체시스템 - 전체교사 시간표 연동 Apps Script (시안 v2) =====
//
// 시간표(전체시간표 시트)는 "읽기 전용" — 구글시트에서 고치면 앱에 그대로 반영되고,
// 앱(브라우저)에서는 시간표를 고칠 수 없습니다.
//
// 다만 아래 2가지는 "1학년 담임"으로 등록된 사람만 브라우저에서 직접 등록/삭제할 수 있고,
// 그 내용이 이 스프레드시트에 그대로 저장되어 다른 모든 사람의 화면에도 반영됩니다.
//   - 학사일정 (휴업일·휴일·단축수업·시험·행사·회의)
//   - 학교 및 학년 안내
// ([수업계] 탭의 학사일정 표는 기존처럼 관리자가 직접 그 화면에서 대량으로 수정하는 용도로
//   남겨뒀습니다 — 이건 그 브라우저에만 저장되는 별개의 로컬 편집 기능입니다.)
//
// ── 사용법 ──────────────────────────────────────────────
// 1) 지금 쓰시는 "전체교사 시간표"를 구글시트에 그대로 붙여넣습니다.
//    (성명, 요일, 교시가 있는 기존 격자 형식 그대로 — 형식을 바꿀 필요 없습니다)
// 2) 같은 스프레드시트에 "설정"이라는 이름의 시트(탭)를 하나 만들고, A1셀에 "담임명단",
//    B1셀에 수정 권한을 줄 이름을 쉼표로 나열해 적습니다. (예: 조시현,허인겸,김서경)
//    이 목록은 여기서만(구글시트에서 직접) 관리합니다 — 브라우저에서는 못 바꿉니다.
//    "학사일정", "안내" 시트는 없어도 됩니다 — 처음 등록할 때 자동으로 만들어집니다.
// 2-1) (선택, 권장) "교사"라는 이름의 시트를 하나 더 만들고, 1행은 헤더(이름/담당과목/교과),
//    2행부터 한 명씩 이름·담당과목·교과(자격증)를 적습니다. (예: 심서아, 화학1, 과학)
//    "동교과 보강 가능" 목록이 이 표의 3번째 칸(교과) 기준으로 정확하게 표시됩니다.
//    이 시트가 없으면 앱이 시간표에 적힌 과목명만 보고 동교과를 추측하므로 부정확할 수 있습니다.
//    이 표도 여기서만(구글시트에서 직접) 관리합니다 — 브라우저에서는 못 바꿉니다.
// 3) 상단 메뉴 [확장 프로그램] > [Apps Script] 를 엽니다.
// 4) 기본으로 생성된 코드를 모두 지우고, 이 파일의 내용을 그대로 붙여넣습니다.
// 5) 아래 SHEET_NAME 값을 실제 시간표가 있는 시트(하단 탭) 이름으로 바꿉니다.
// 6) 상단 [배포] > [새 배포] 클릭
//      - 유형 선택(톱니바퀴) > "웹 앱"
//      - 설명: 아무거나 (예: 시간표 연동)
//      - 실행 계정: 나
//      - 액세스 권한: "링크가 있는 모든 사용자"
//        (학교 밖에서 접근 못 하게 하려면 "Google Workspace 조직 내" 선택 — 단, 학교 계정이
//         Workspace 조직으로 관리되고 있어야 선택 가능합니다. 잘 모르면 우선 "모든 사용자"로 배포하세요.)
// 7) [배포] 클릭 → 나오는 "웹 앱 URL"을 복사합니다. (…/exec 로 끝나는 주소)
// 8) 수업교체시스템 앱의 [수업계] 탭 → "전체교사 시간표 연동(구글시트)" 칸에 이 URL을 붙여넣고 저장합니다.
//
// ── 이후 시간표/설정을 고칠 때 ────────────────────────────
// 이 시트에서 시간표·"설정" 탭의 담임명단을 고치기만 하면 됩니다. 배포를 다시 할 필요 없습니다.
// 앱에서 "수업계 > 저장 후 불러오기"를 누르거나, 앱을 새로고침하면 최신 내용이 반영됩니다.
// (같은 배포 URL을 계속 씁니다. 시트 구조/열 순서를 크게 바꾸지 않는 한 URL은 그대로입니다.)
//
// ── 시트가 여러 개(학년별 등)일 때 ──────────────────────
// URL 뒤에 ?sheet=시트이름 을 붙이면 그 시트를 읽습니다. 예:
//   https://script.google.com/macros/s/xxxx/exec?sheet=1학년시간표
// 아무것도 안 붙이면 아래 SHEET_NAME(기본 시트)을 읽습니다.
// ─────────────────────────────────────────────────────

const SHEET_NAME = '전체시간표'; // ← 실제 탭 이름으로 수정하세요
const ACADEMIC_SHEET = '학사일정';
const NOTICE_SHEET = '안내';
const CONFIG_SHEET = '설정';
const TEACHER_SHEET = '교사'; // 이름 | 담당과목 | 교과(자격증) — "동교과 보강" 판정에 쓰임

function doGet(e) {
  try {
    const wantSheet = (e && e.parameter && e.parameter.sheet) ? e.parameter.sheet : SHEET_NAME;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(wantSheet);
    if (!sheet) {
      return jsonOut({ error: '시트를 찾을 수 없습니다: ' + wantSheet });
    }
    // 화면에 보이는 그대로(서식이 적용된 표시값)를 그대로 내보냅니다.
    // 기존 CSV/엑셀 업로드와 완전히 같은 격자 구조이므로, 앱은 지금까지 쓰던 방식 그대로 해석합니다.
    const rows = sheet.getDataRange().getDisplayValues();
    return jsonOut({
      rows: rows,
      academic: readAcademic(ss),
      notices: readNotices(ss),
      homerooms: readHomerooms(ss),
      teachers: readTeachers(ss)
    });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// 브라우저(1학년 담임)가 학사일정/안내를 추가·삭제할 때만 호출됩니다.
// 그 외 어떤 것도 앱에서 이 스프레드시트로 써 넣지 않습니다(시간표·기타 설정은 여기서 안 건드림).
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editor = String(body.editor || '').trim();
    const homerooms = readHomerooms(ss);
    if (!homerooms.length) {
      return jsonOut({ error: '"설정" 시트에 담임명단이 등록되어 있지 않습니다. A1=담임명단, B1=이름들(쉼표구분)을 채워주세요.' });
    }
    if (!homerooms.includes(editor)) {
      return jsonOut({ error: '수정 권한이 없습니다: ' + editor });
    }
    if (body.action === 'addAcademic') {
      appendAcademic(ss, body.entry || {});
    } else if (body.action === 'updateAcademic') {
      updateAcademic(ss, body.id, body.entry || {});
    } else if (body.action === 'deleteAcademic') {
      deleteRowById(ss, ACADEMIC_SHEET, body.id);
    } else if (body.action === 'addNotice') {
      appendNotice(ss, body.entry || {}, editor);
    } else if (body.action === 'updateNotice') {
      updateNotice(ss, body.id, body.entry || {});
    } else if (body.action === 'deleteNotice') {
      deleteRowById(ss, NOTICE_SHEET, body.id);
    } else {
      return jsonOut({ error: '알 수 없는 action: ' + body.action });
    }
    return jsonOut({ ok: true, academic: readAcademic(ss), notices: readNotices(ss) });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function getOrCreateSheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
  }
  return sh;
}

// ── 학사일정: id | date | end | type | last | title | grades | noClass | body | periods ──
// periods: 이 일정 때문에 특정 교시만 수업이 안 되는 경우(예: 영어듣기평가 3교시) 그 교시 번호들(쉼표구분).
// 하루 전체가 아니라 그 교시만 막고 싶을 때 씁니다 — noClass(종일 휴업)와는 별개입니다.
function readAcademic(ss) {
  const sh = ss.getSheetByName(ACADEMIC_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[1]) continue; // 시작일이 비어있으면 건너뜀
    out.push({
      id: r[0], date: r[1], end: r[2] || r[1], type: r[3] || '휴업일',
      last: r[4] ? Number(r[4]) : undefined, title: r[5] || '',
      grades: r[6] ? String(r[6]).split(',').map(function (x) { return +x.trim(); }).filter(Boolean) : [],
      noClass: String(r[7]).toUpperCase() === 'TRUE', body: r[8] || '',
      periods: r[9] ? String(r[9]).split(',').map(function (x) { return +x.trim(); }).filter(Boolean) : []
    });
  }
  return out;
}

function academicRowValues(id, entry) {
  return [
    id, entry.date, entry.end || entry.date, entry.type || '휴업일', entry.last || '',
    entry.title || '', (entry.grades || []).join(','), entry.noClass ? 'TRUE' : 'FALSE',
    entry.body || '', (entry.periods || []).join(',')
  ];
}

function appendAcademic(ss, entry) {
  if (!entry.date) throw new Error('날짜가 없습니다.');
  const sh = getOrCreateSheet(ss, ACADEMIC_SHEET, ['id', 'date', 'end', 'type', 'last', 'title', 'grades', 'noClass', 'body', 'periods']);
  const id = 'a' + new Date().getTime();
  sh.appendRow(academicRowValues(id, entry));
}

function updateAcademic(ss, id, entry) {
  if (!id) throw new Error('수정할 항목의 id가 없습니다.');
  const sh = ss.getSheetByName(ACADEMIC_SHEET);
  if (!sh) throw new Error('학사일정 시트가 없습니다.');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.getRange(i + 1, 1, 1, 10).setValues([academicRowValues(id, entry)]);
      return;
    }
  }
  throw new Error('해당 id를 찾을 수 없습니다: ' + id);
}

// ── 안내: id | date | writer | title | body ──
function readNotices(ss) {
  const sh = ss.getSheetByName(NOTICE_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({ id: r[0], date: r[1], writer: r[2], title: r[3], body: r[4] || '' });
  }
  return out.reverse(); // 최신 글이 맨 위로
}

function appendNotice(ss, entry, editor) {
  if (!entry.title) throw new Error('제목이 없습니다.');
  const sh = getOrCreateSheet(ss, NOTICE_SHEET, ['id', 'date', 'writer', 'title', 'body']);
  const id = 'n' + new Date().getTime();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([id, today, editor, entry.title || '', entry.body || '']);
}

function updateNotice(ss, id, entry) {
  if (!id) throw new Error('수정할 항목의 id가 없습니다.');
  const sh = ss.getSheetByName(NOTICE_SHEET);
  if (!sh) throw new Error('안내 시트가 없습니다.');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.getRange(i + 1, 4, 1, 2).setValues([[entry.title || '', entry.body || '']]);
      return;
    }
  }
  throw new Error('해당 id를 찾을 수 없습니다: ' + id);
}

function deleteRowById(ss, sheetName, id) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      break;
    }
  }
}

// "설정" 시트에서 "담임명단" 행을 찾아 그 옆칸(쉼표로 구분된 이름들)을 읽습니다.
// 이 목록은 관리자가 구글시트에서 직접 관리합니다 — 브라우저에서는 바꿀 수 없습니다.
function readHomerooms(ss) {
  const sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === '담임명단') {
      return String(rows[i][1] || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
    }
  }
  return [];
}

// "교사" 시트: 이름 | 담당과목 | 교과(자격증) — 예: 심서아 | 화학1 | 과학
// 3번째 칸(교과)이 같은 사람끼리 "동교과 보강 가능"으로 판정됩니다. 시트가 없으면 빈 값을 보내고,
// 앱이 시간표에 나온 과목명만 보고 동교과를 추정합니다(부정확할 수 있음 — 이 시트를 채우는 걸 권장).
// 이 시트는 관리자가 구글시트에서 직접 관리합니다 — 브라우저에서는 바꿀 수 없습니다.
function readTeachers(ss) {
  const sh = ss.getSheetByName(TEACHER_SHEET);
  if (!sh) return '';
  const rows = sh.getDataRange().getDisplayValues();
  const lines = [];
  for (let i = 1; i < rows.length; i++) { // 1행은 헤더로 보고 건너뜀
    const name = String(rows[i][0] || '').trim();
    if (!name) continue;
    const subject = String(rows[i][1] || '').trim();
    const dept = String(rows[i][2] || '').trim();
    lines.push([name, subject, dept].join(','));
  }
  return lines.join('\n');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
