// ===== 수업교체시스템 - 전체교사 시간표 연동 Apps Script (시안 v2) =====
//
// 시간표(전체시간표 시트)는 "읽기 전용" — 구글시트에서 고치면 앱에 그대로 반영되고,
// 앱(브라우저)에서는 시간표를 고칠 수 없습니다.
//
// 다만 아래 항목들은 브라우저에서 직접 등록/삭제할 수 있고, 그 내용이 이 스프레드시트에
// 그대로 저장되어 다른 모든 사람의 화면에도 반영됩니다.
//   - 학사일정 (휴업일·휴일·단축수업·시험·행사·회의) — "1학년 담임"만 등록/수정/삭제 가능
//   - 학교 및 학년 안내 — "1학년 담임"만 등록/수정/삭제 가능
//   - 수업교체·보강 확정 기록 — 아무 교사나 자기 자신의 것을 등록/취소 가능. 이걸 통해
//     반별 시간표에 접속한 모든 사람이 확정된 수업교체·보강을 실시간으로 볼 수 있습니다.
//   - 진도표(반·과목별 진도/특이사항) — "설정" 시트의 "진도표접근허용" 명단에 있는 교사만
//     "내 시간표"에서 반 칸을 눌러 열람·수정 가능.
//     ("학사일정"/"안내"/"수업교체"/"진도표" 시트는 없어도 됩니다 — 처음 등록할 때 자동으로 만들어집니다.)
//
// ── 사용법 ──────────────────────────────────────────────
// 1) 지금 쓰시는 "전체교사 시간표"를 구글시트에 그대로 붙여넣습니다.
//    (성명, 요일, 교시가 있는 기존 격자 형식 그대로 — 형식을 바꿀 필요 없습니다)
// 2) 같은 스프레드시트에 "설정"이라는 이름의 시트(탭)를 하나 만들고, A1셀에 "담임명단",
//    B1셀에 수정 권한을 줄 이름을 쉼표로 나열해 적습니다. (예: 조시현,허인겸,김서경)
//    A2셀에 "표시주수", B2셀에 몇 주치를 보여줄지 숫자로 적습니다. (예: 10)
//    A3셀에 "수업교체대상", B3셀에 수업교체·보강 기록을 이 시트로 받을 교사만 쉼표로
//    나열해 적습니다. (예: 조시현,정병성,최광혁 — 관리 범위인 1학년 관련 교사만) 이 칸을
//    비워두면(행 자체가 없으면) 예전처럼 모든 교사의 기록을 다 받습니다.
//    A4셀에 "진도표접근허용", B4셀에 "내 시간표"에서 진도표를 열람·수정할 수 있는 교사를
//    쉼표로 나열해 적습니다. (예: 조시현) 이 칸을 비워두면 아무도 진도표 기능을 쓸 수
//    없습니다(기본이 "전부 차단"인 유일한 항목 — 명단을 넣기 전까진 기능 자체가 꺼져 있음).
//    이 값들은 여기서만(구글시트에서 직접) 관리합니다 — 브라우저에서는 못 바꿉니다.
//    "학사일정", "안내", "진도표" 시트는 없어도 됩니다 — 처음 등록할 때 자동으로 만들어집니다.
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
const SWAP_SHEET = '수업교체'; // 확정된 맞교체/보강 기록 — 반별 시간표에 모두에게 반영됨
const PROGRESS_SHEET = '진도표'; // 반·과목별 진도/특이사항 기록 — 열람·수정 모두 권한이 있는 교사만

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
      teachers: readTeachers(ss),
      weeks: readWeeks(ss),
      swaps: readSwaps(ss),
      // 진도표 열람·수정 권한 명단만 실어 보냄(실제 진도 내용은 반 칸을 눌렀을 때만 따로 요청함) —
      // 로그인 시 이 목록만 한 번 받아서 브라우저가 캐시해두고, 그 뒤로는 클릭할 때마다
      // 서버에 권한을 다시 묻지 않음.
      progressAccess: readProgressAccess(ss)
    });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// 학사일정/안내 추가·삭제는 "1학년 담임"만 — 그 외(수업교체·보강 기록)는 자기 자신의
// 시간표를 바꾸는 것이므로 아무 교사나 가능합니다. 시간표·기타 설정은 여기서 안 건드림.
const HOMEROOM_ONLY_ACTIONS = ['addAcademic', 'updateAcademic', 'deleteAcademic', 'addNotice', 'updateNotice', 'deleteNotice'];

// 진도표 열람·수정은 별도 권한 명단(설정 시트의 "진도표접근허용")으로 가둠 — 학사일정/안내와도,
// 수업교체와도 다른 별개의 명단.
const PROGRESS_ONLY_ACTIONS = ['getProgress', 'saveProgress'];

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const editor = String(body.editor || '').trim();
    if (HOMEROOM_ONLY_ACTIONS.indexOf(body.action) >= 0) {
      const homerooms = readHomerooms(ss);
      if (!homerooms.length) {
        return jsonOut({ error: '"설정" 시트에 담임명단이 등록되어 있지 않습니다. A1=담임명단, B1=이름들(쉼표구분)을 채워주세요.' });
      }
      if (!homerooms.includes(editor)) {
        return jsonOut({ error: '수정 권한이 없습니다: ' + editor });
      }
    }
    if (PROGRESS_ONLY_ACTIONS.indexOf(body.action) >= 0) {
      const allowed = readProgressAccess(ss);
      if (!allowed.length || allowed.indexOf(editor) < 0) {
        return jsonOut({ error: '진도표 접근 권한이 없습니다: ' + editor });
      }
      // 진도표는 다른 시트들(학사일정·안내·수업교체)을 매번 다시 안 읽고 이 요청만 처리해서
      // 응답을 가볍게 유지함(칸을 누를 때마다 호출되므로).
      if (body.action === 'getProgress') {
        return jsonOut({ ok: true, progress: readProgress(ss, body.cls, body.subject) });
      }
      if (body.action === 'saveProgress') {
        upsertProgress(ss, body.entry || {}, editor);
        return jsonOut({ ok: true, progress: readProgress(ss, body.entry.cls, body.entry.subject) });
      }
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
    } else if (body.action === 'addSwapBatch') {
      const scope = readSwapScope(ss);
      const entries = body.entries || [];
      const allowed = scope.length ? entries.filter(function (en) { return scope.indexOf(en.applicant) >= 0; }) : entries;
      appendSwapBatch(ss, allowed);
    } else if (body.action === 'deleteSwap') {
      deleteSwapGroup(ss, body.grp, body.id);
    } else {
      return jsonOut({ error: '알 수 없는 action: ' + body.action });
    }
    return jsonOut({ ok: true, academic: readAcademic(ss), notices: readNotices(ss), swaps: readSwaps(ss) });
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

// ── 안내: id | date | writer | title | body | category | dept | done ──
// category: 'school'(학교안내) | 'grade'(학년안내). dept(관련부서)는 학교안내에서만 씀.
// done: 'TRUE'면 완료사항으로 넘어가 초기 화면 요약 카드에서는 숨겨지고, 상세 모달의
// "완료사항" 탭에서만 계속 확인할 수 있음(삭제 아님).
const NOTICE_HEADER = ['id', 'date', 'writer', 'title', 'body', 'category', 'dept', 'done'];

function readNotices(ss) {
  const sh = ss.getSheetByName(NOTICE_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({
      id: r[0], date: r[1], writer: r[2], title: r[3], body: r[4] || '',
      category: r[5] || 'school', dept: r[6] || '', done: String(r[7]).toUpperCase() === 'TRUE'
    });
  }
  return out.reverse(); // 최신 글이 맨 위로
}

function noticeRowValues(id, entry) {
  return [
    id, entry.date, entry.writer, entry.title || '', entry.body || '',
    entry.category || 'school', entry.dept || '', entry.done ? 'TRUE' : 'FALSE'
  ];
}

function appendNotice(ss, entry, editor) {
  if (!entry.title) throw new Error('제목이 없습니다.');
  const sh = getOrCreateSheet(ss, NOTICE_SHEET, NOTICE_HEADER);
  const id = 'n' + new Date().getTime();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow(noticeRowValues(id, {
    date: today, writer: editor, title: entry.title, body: entry.body,
    category: entry.category, dept: entry.dept, done: false
  }));
}

// entry로 넘어온 항목만 바꾸고 나머지(작성일·작성자 등)는 기존 값을 그대로 유지합니다.
// (예: {done:true}만 보내면 완료 체크만 반영되고 제목·내용은 그대로.)
function updateNotice(ss, id, entry) {
  if (!id) throw new Error('수정할 항목의 id가 없습니다.');
  const sh = ss.getSheetByName(NOTICE_SHEET);
  if (!sh) throw new Error('안내 시트가 없습니다.');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const cur = values[i];
      const merged = {
        date: cur[1], writer: cur[2], title: cur[3], body: cur[4],
        category: cur[5] || 'school', dept: cur[6] || '', done: String(cur[7]).toUpperCase() === 'TRUE'
      };
      Object.keys(entry).forEach(function (k) { if (entry[k] !== undefined) merged[k] = entry[k]; });
      sh.getRange(i + 1, 1, 1, NOTICE_HEADER.length).setValues([noticeRowValues(id, merged)]);
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

// "설정" 시트에서 "수업교체대상" 행을 찾아 그 옆칸(쉼표로 구분된 이름들)을 읽습니다.
// 여기 적힌 교사만 수업교체/보강 기록이 이 스프레드시트에 쌓이고 브라우저로도 내려갑니다
// (관리 범위 밖 학년 교사까지 전부 쌓이면 시트가 커져서 앱스크립트가 느려지거나 오류가
// 날 수 있어서, 관리자가 직접 볼 필요가 있는 사람만 고를 수 있게 한 것). 이 칸이 비어있으면
// (행 자체가 없으면) 제한 없이 전부 동기화합니다 — 기존처럼 쓰던 학교는 영향 없음.
function readSwapScope(ss) {
  const sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === '수업교체대상') {
      return String(rows[i][1] || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
    }
  }
  return [];
}

// "설정" 시트에서 "진도표접근허용" 행을 찾아 그 옆칸(쉼표로 구분된 이름들)을 읽습니다.
// 여기 적힌 교사만 "내 시간표"에서 반 칸을 눌러 진도표를 열람·수정할 수 있습니다. 이 칸이
// 비어있으면(행 자체가 없으면) 아무도 못 씁니다 — 안내/학사일정과 달리 기본값이 "전부 허용"이
// 아니라 "전부 차단"인 이유는, 관리자가 명단을 넣기 전까지는 이 기능 자체가 없는 셈이어야
// (실수로 아무나 다 쓰다가 데이터가 커지는 일이 없도록) 하기 때문입니다.
function readProgressAccess(ss) {
  const sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === '진도표접근허용') {
      return String(rows[i][1] || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
    }
  }
  return [];
}

// "설정" 시트에서 "담임명단" 행을 찾아 그 옆칸(쉼표로 구분된 이름들)을 읽습니다.
// 이 목록은 관리자가 구글시트에서 직접 관리합니다 — 브라우저에서는 못 바꿉니다.
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

// "설정" 시트에서 "표시주수" 행을 찾아 그 옆칸(숫자)을 읽습니다. 앱의 "내 시간표/수업교체/반별 시간표"에
// 몇 주 앞까지 보여줄지 정하는 값입니다 — 이것도 관리자가 구글시트에서 직접 관리합니다.
function readWeeks(ss) {
  const sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === '표시주수') {
      const n = Number(rows[i][1]);
      return n > 0 ? n : null;
    }
  }
  return null;
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

// ── 수업교체: id | type | applicant | initiator | grp | primary | sameDept | reason | writeDate |
//    absent_date | absent_grade | absent_cls | absent_period | absent_subject |
//    cover_teacher | cover_subject | makeup_date | makeup_period ──
// 맞교체(swap) 1건은 당사자 두 명이 각자의 시점으로 레코드 2개(같은 grp)를, 3인 이상 연쇄이동은
// 3개를 함께 씁니다. id는 브라우저(각자 기기)가 만들어 그대로 저장 — 서버가 새로 만들지 않습니다.
// 아무 교사나 자기 자신의 교체·보강 기록을 쓸 수 있습니다(1학년 담임 전용 아님).
const SWAP_HEADER = ['id', 'type', 'applicant', 'initiator', 'grp', 'primary', 'sameDept', 'reason', 'writeDate',
  'absent_date', 'absent_grade', 'absent_cls', 'absent_period', 'absent_subject',
  'cover_teacher', 'cover_subject', 'makeup_date', 'makeup_period'];

function readSwaps(ss) {
  const sh = ss.getSheetByName(SWAP_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({
      id: r[0], type: r[1], applicant: r[2], initiator: r[3], grp: r[4] || undefined,
      primary: r[5] === '' ? undefined : String(r[5]).toUpperCase() === 'TRUE',
      sameDept: String(r[6]).toUpperCase() === 'TRUE', reason: r[7] || '', writeDate: r[8] || '',
      absent: { date: r[9], grade: r[10], cls: r[11], period: r[12] ? Number(r[12]) : '', subject: r[13] || '' },
      cover: { teacher: r[14] || '', subject: r[15] || '' },
      makeup: { date: r[16] || '', period: r[17] || '' }
    });
  }
  const scope = readSwapScope(ss);
  return scope.length ? out.filter(function (r) { return scope.indexOf(r.applicant) >= 0; }) : out;
}

function swapRowValues(entry) {
  const a = entry.absent || {}, c = entry.cover || {}, mk = entry.makeup || {};
  return [
    entry.id, entry.type, entry.applicant, entry.initiator || '', entry.grp || '',
    entry.primary === undefined ? '' : (entry.primary ? 'TRUE' : 'FALSE'),
    entry.sameDept ? 'TRUE' : 'FALSE', entry.reason || '', entry.writeDate || '',
    a.date || '', a.grade || '', a.cls || '', a.period || '', a.subject || '',
    c.teacher || '', c.subject || '', mk.date || '', mk.period || ''
  ];
}

function appendSwapBatch(ss, entries) {
  if (!entries.length) return;
  const sh = getOrCreateSheet(ss, SWAP_SHEET, SWAP_HEADER);
  const values = entries.map(function (entry) {
    if (!entry.id) throw new Error('레코드에 id가 없습니다.');
    return swapRowValues(entry);
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, SWAP_HEADER.length).setValues(values);
}

// grp가 있으면 같은 grp를 가진 모든 행을(맞교체 양쪽 당사자), 없으면 그 id 하나만 지웁니다.
function deleteSwapGroup(ss, grp, id) {
  const sh = ss.getSheetByName(SWAP_SHEET);
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = 1; i < values.length; i++) {
    const matches = grp ? String(values[i][4]) === String(grp) : String(values[i][0]) === String(id);
    if (matches) rowsToDelete.push(i + 1);
  }
  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowNum) { sh.deleteRow(rowNum); });
}

// ── 진도표: id | cls | subject | date | period | body | writer | updatedAt ──
// 반(cls, 예 "1-6") + 과목(subject, 예 "통합과학2") 조합별로 날짜·교시마다 한 줄씩.
// 같은 반·과목·날짜·교시에 다시 저장하면 새 줄을 만들지 않고 그 줄의 내용만 덮어씁니다.
const PROGRESS_HEADER = ['id', 'cls', 'subject', 'date', 'period', 'body', 'writer', 'updatedAt'];

function readProgress(ss, cls, subject) {
  const sh = ss.getSheetByName(PROGRESS_SHEET);
  if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    if (String(r[1]) !== String(cls) || String(r[2]) !== String(subject)) continue;
    out.push({ id: r[0], cls: r[1], subject: r[2], date: r[3], period: r[4] ? Number(r[4]) : '', body: r[5] || '', writer: r[6] || '' });
  }
  out.sort(function (a, b) { return (a.date + '_' + a.period).localeCompare(b.date + '_' + b.period); });
  return out;
}

function upsertProgress(ss, entry, editor) {
  if (!entry.cls || !entry.subject || !entry.date || !entry.period) throw new Error('반·과목·날짜·교시가 모두 있어야 합니다.');
  const sh = getOrCreateSheet(ss, PROGRESS_SHEET, PROGRESS_HEADER);
  const values = sh.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[1]) === String(entry.cls) && String(r[2]) === String(entry.subject) &&
      String(r[3]) === String(entry.date) && String(r[4]) === String(entry.period)) {
      sh.getRange(i + 1, 6, 1, 3).setValues([[entry.body || '', editor, now]]);
      return;
    }
  }
  const id = 'p' + new Date().getTime();
  sh.appendRow([id, entry.cls, entry.subject, entry.date, entry.period, entry.body || '', editor, now]);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
