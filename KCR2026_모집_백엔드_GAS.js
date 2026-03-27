// ============================================================
// KCR 2026 운영요원 모집 — Apps Script 백엔드
//
// [배포 방법]
// 1. 구글 스프레드시트 새로 만들기 (모집 전용)
// 2. 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
// 3. 저장 → 배포 → 새 배포 → 웹 앱 선택
// 4. 설정:
//    - 설명: KCR2026 운영요원 모집
//    - 다음 사용자로 실행: 나(본인)
//    - 액세스 권한: 모든 사용자
// 5. 배포 → 승인 → URL 복사
// 6. 복사한 URL을 recruit.html의 GAS_URL 변수에 붙여넣기
//
// [트리거 설정 — 최초 1회]
// Apps Script 에디터에서 setRecruitTriggers() 함수를 한 번 실행하세요.
// → onEdit 트리거가 등록되어 상태 변경 시 자동 이메일이 발송됩니다.
// ============================================================

const PM_EMAIL    = "info@kcr2026.com";
const SHEET_NAME  = "운영요원모집";
const HEADERS     = [
  "타임스탬프", "이름", "성별", "생년월일", "휴대전화", "이메일", "주소",
  "최종학력", "재학상태", "전공", "자기소개",
  "경험유무", "참여행사1", "참여행사2", "참여행사3",
  "영어능력", "기타언어",
  "희망업무", "4/13참여", "5/14참여", "5/15참여", "5/16참여",
  "상태", "상태변경일시", "메모",
  "최종배정파트",
  "출근_4/13", "퇴근_4/13",
  "출근_5/14", "퇴근_5/14",
  "출근_5/15", "퇴근_5/15",
  "출근_5/16", "퇴근_5/16"
];

// ── GET 요청: 연결 테스트 ─────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "KCR2026 운영요원 모집 API 정상 작동 중" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST 요청: 지원서 접수 ────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 중복 체크
    if (checkDuplicate_(data.email)) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "duplicate", message: "이미 접수된 이메일입니다" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    saveApplication_(data);
    sendConfirmationEmail_(data);
    sendNewApplicationAlert_(data);
    updateRecruitDashboard_();

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 시트에 저장 ───────────────────────────────────────────────
function saveApplication_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setBackground("#185FA5").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
    // 상태 열에 드롭다운 설정 (V열 = 22번째 열)
    setupStatusDropdown(sheet);
  }

  const row = [
    data.timestamp  || "",
    data.name       || "",
    data.gender     || "",
    data.birth      || "",
    data.phone      || "",
    data.email      || "",
    data.address    || "",
    data.education  || "",
    data.enrollment || "",
    data.major      || "",
    data.intro      || "",
    data.experience || "",
    data.exp1       || "",
    data.exp2       || "",
    data.exp3       || "",
    data.english    || "",
    data.other_lang || "",
    data.roles      || "",
    data.avail_13   || "",    // 4/13참여
    data.avail_14   || "",
    data.avail_15   || "",
    data.avail_16   || "",
    "접수완료",                // 초기 상태
    "",                       // 상태변경일시
    "",                       // 메모
    "",                       // 최종배정파트
    "", "",                   // 출근_4/13, 퇴근_4/13
    "", "",                   // 출근_5/14, 퇴근_5/14
    "", "",                   // 출근_5/15, 퇴근_5/15
    "", ""                    // 출근_5/16, 퇴근_5/16
  ];

  const lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);
  sheet.autoResizeColumns(1, HEADERS.length);
}

// ── 상태 열 드롭다운 설정 ─────────────────────────────────────
// 시트 생성 시 자동 호출됨. 기존 시트에 적용하려면 수동으로 실행하세요.
function setupStatusDropdown(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  }
  if (!sheet) return;

  const statusCol = HEADERS.indexOf("상태") + 1;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["접수완료", "서류검토중", "승인", "거절"], true)
    .setAllowInvalid(false)
    .build();

  // 2행부터 1000행까지 드롭다운 적용
  sheet.getRange(2, statusCol, 999, 1).setDataValidation(rule);
  Logger.log("상태 드롭다운 설정 완료");
}

// ── 이메일 중복 체크 ──────────────────────────────────────────
function checkDuplicate_(email) {
  if (!email) return false;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;

  const emailCol = HEADERS.indexOf("이메일") + 1;
  const emails   = sheet.getRange(2, emailCol, sheet.getLastRow() - 1, 1).getValues();
  return emails.some(r => r[0] && r[0].toString().toLowerCase() === email.toLowerCase());
}

// ── 접수 확인 이메일 (지원자에게) ─────────────────────────────
function sendConfirmationEmail_(data) {
  const subject = "[KCR2026] 운영요원 지원서가 접수되었습니다";
  const body =
    `${data.name || ""}님, 안녕하세요.\n\n` +
    `KCR 2026 운영요원 지원서가 정상적으로 접수되었습니다.\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `접수 일시  : ${data.timestamp || ""}\n` +
    `지원자     : ${data.name || ""}\n` +
    `희망 업무  : ${data.roles || ""}\n` +
    `참여 가능일: 5/14 ${data.avail_14||""} · 5/15 ${data.avail_15||""} · 5/16 ${data.avail_16||""}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `서류 검토 후 결과를 이메일로 안내드리겠습니다.\n` +
    `감사합니다.\n\n` +
    `KCR 2026 사무국 드림`;

  GmailApp.sendEmail(data.email, subject, body, { name: "KCR2026 운영요원 모집" });
}

// ── 신규 지원 알림 (PM에게) ───────────────────────────────────
function sendNewApplicationAlert_(data) {
  const subject = `[KCR2026 모집] 신규 지원 — ${data.name || ""} (${data.roles || ""})`;
  const body =
    `새로운 운영요원 지원서가 접수되었습니다.\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `이름      : ${data.name || ""}\n` +
    `연락처    : ${data.phone || ""}\n` +
    `이메일    : ${data.email || ""}\n` +
    `최종학력  : ${data.education || ""}\n` +
    `희망 업무 : ${data.roles || ""}\n` +
    `참여 가능 : 5/14 ${data.avail_14||""} · 5/15 ${data.avail_15||""} · 5/16 ${data.avail_16||""}\n` +
    `영어      : ${data.english || ""}\n` +
    `경험      : ${data.experience || ""}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `전체 지원 현황: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  GmailApp.sendEmail(PM_EMAIL, subject, body, { name: "KCR2026 운영요원 모집" });
}

// ── 대시보드 자동 갱신 ────────────────────────────────────────
function updateRecruitDashboard_() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const src  = ss.getSheetByName(SHEET_NAME);
  if (!src) return;

  const data = src.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const rows    = data.slice(1).filter(r => r[0]);
  const col     = name => headers.indexOf(name);

  let dash = ss.getSheetByName("대시보드");
  if (!dash) dash = ss.insertSheet("대시보드");
  dash.clearContents();
  dash.clearFormats();

  const BLUE  = "#185FA5";
  const LBLUE = "#E6F1FB";
  const GRAY  = "#F9F9F9";
  const GREEN = "#D1F2E0";
  const YELLOW = "#FFF3CD";

  // ── 제목 ──
  dash.getRange("A1").setValue("KCR 2026 운영요원 모집 대시보드")
    .setFontSize(15).setFontWeight("bold").setFontColor(BLUE);
  dash.getRange("A2")
    .setValue("최종 업데이트: " + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"))
    .setFontSize(10).setFontColor("#8E8E93");

  // ── 요약 통계 ──
  const total     = rows.length;
  const accepted  = rows.filter(r => r[col("상태")] === "접수완료").length;
  const reviewing = rows.filter(r => r[col("상태")] === "서류검토중").length;
  const approved  = rows.filter(r => r[col("상태")] === "승인").length;
  const rejected  = rows.filter(r => r[col("상태")] === "거절").length;

  dash.getRange("A4:E4")
    .setValues([["총 지원자", "접수완료", "서류검토중", "승인", "거절"]])
    .setBackground(BLUE).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  dash.getRange("A5:E5")
    .setValues([[total, accepted, reviewing, approved, rejected]])
    .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");

  // ── 희망업무별 현황 ──
  dash.getRange("A7").setValue("희망업무별 지원 현황")
    .setFontWeight("bold").setFontColor(BLUE);
  const roles = ["세션장", "등록 데스크", "프리뷰", "포스터"];
  dash.getRange(8, 1, 1, roles.length + 1)
    .setValues([["업무", ...roles]])
    .setBackground(LBLUE).setFontWeight("bold");

  const roleCounts = roles.map(role =>
    rows.filter(r => (r[col("희망업무")] || "").includes(role)).length
  );
  dash.getRange(9, 1, 1, roles.length + 1)
    .setValues([["지원자 수", ...roleCounts]])
    .setHorizontalAlignment("center");

  // ── 날짜별 참여가능 인원 ──
  dash.getRange("A11").setValue("날짜별 참여가능 인원")
    .setFontWeight("bold").setFontColor(BLUE);
  const dates = ["5/14", "5/15", "5/16"];
  const slots = ["전일", "오전", "오후", "불가"];
  dash.getRange(12, 1, 1, slots.length + 1)
    .setValues([["날짜", ...slots]])
    .setBackground(LBLUE).setFontWeight("bold");

  dates.forEach((d, di) => {
    const dateCol = col(d + "참여");
    const counts = slots.map(s =>
      rows.filter(r => r[dateCol] === s).length
    );
    dash.getRange(13 + di, 1, 1, slots.length + 1)
      .setValues([[d, ...counts]])
      .setBackground(di % 2 === 0 ? GRAY : "#FFFFFF")
      .setHorizontalAlignment("center");
  });

  // ── 최근 지원자 목록 (최신 10건) ──
  dash.getRange("A17").setValue(`최근 지원자 (${Math.min(10, total)}건)`)
    .setFontWeight("bold").setFontColor(BLUE);

  const listH = ["접수일시", "이름", "최종학력", "희망업무", "영어", "경험", "상태"];
  dash.getRange(18, 1, 1, listH.length)
    .setValues([listH])
    .setBackground(LBLUE).setFontWeight("bold");

  const recent = rows.slice(-10).reverse();
  if (recent.length > 0) {
    const listData = recent.map(r => [
      r[col("타임스탬프")], r[col("이름")], r[col("최종학력")],
      r[col("희망업무")], r[col("영어능력")], r[col("경험유무")], r[col("상태")]
    ]);
    dash.getRange(19, 1, listData.length, listH.length).setValues(listData);

    // 상태별 색상
    listData.forEach((row, ri) => {
      const status = row[6];
      if (status === "승인") dash.getRange(19 + ri, 7).setBackground(GREEN);
      else if (status === "거절") dash.getRange(19 + ri, 7).setBackground(YELLOW);
    });
  }

  // 열 폭 조정
  for (let c = 1; c <= 7; c++) dash.setColumnWidth(c, 120);
  dash.autoResizeColumns(1, listH.length);
}

// ── 상태 변경 감지 → 이메일 발송 ──────────────────────────────
function onStatusChange(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const range = e.range;
  const statusCol = HEADERS.indexOf("상태") + 1;

  // 상태 열 변경인지 확인
  if (range.getColumn() !== statusCol) return;

  const row      = range.getRow();
  if (row < 2) return; // 헤더행 무시

  const newStatus = range.getValue();
  const emailCol  = HEADERS.indexOf("이메일") + 1;
  const nameCol   = HEADERS.indexOf("이름") + 1;
  const tsCol     = HEADERS.indexOf("상태변경일시") + 1;

  const email = sheet.getRange(row, emailCol).getValue();
  const name  = sheet.getRange(row, nameCol).getValue();

  if (!email) return;

  // 상태변경일시 기록
  sheet.getRange(row, tsCol)
    .setValue(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"));

  // 상태별 이메일 발송
  sendStatusChangeEmail_(email, name, newStatus);
}

// ── 상태 변경 이메일 ──────────────────────────────────────────
function sendStatusChangeEmail_(email, name, status) {
  let subject = "";
  let body = "";

  switch (status) {
    case "서류검토중":
      subject = "[KCR2026] 운영요원 지원서 검토 안내";
      body =
        `${name || ""}님, 안녕하세요.\n\n` +
        `KCR 2026 운영요원에 지원해주셔서 진심으로 감사드립니다.\n\n` +
        `현재 제출하신 지원서를 검토 중에 있으며, 검토가 완료되는 대로 결과를 이메일로 안내드리겠습니다.\n\n` +
        `조금만 기다려 주시면 감사하겠습니다.\n\n` +
        `KCR 2026 사무국 드림`;
      break;
    case "승인":
      subject = "[KCR2026] 운영요원 선발 안내";
      body =
        `${name || ""}님, 안녕하세요.\n\n` +
        `KCR 2026 운영요원 모집에 지원해주셔서 감사드립니다.\n\n` +
        `서류 검토 결과, ${name || ""}님을 KCR 2026 운영요원으로 선발하게 되었습니다.\n` +
        `함께하게 되어 정말 기쁩니다!\n\n` +
        `오리엔테이션 일정 및 상세 안내는 추후 별도 이메일로 안내드리겠습니다.\n\n` +
        `KCR 2026 사무국 드림`;
      break;
    case "거절":
      subject = "[KCR2026] 운영요원 지원 결과 안내";
      body =
        `${name || ""}님, 안녕하세요.\n\n` +
        `KCR 2026 운영요원 모집에 소중한 관심을 가져주시고 지원해주셔서 진심으로 감사드립니다.\n\n` +
        `신중한 검토 끝에, 아쉽게도 이번에는 함께하지 못하게 되었습니다.\n` +
        `지원해주신 열정과 관심에 깊이 감사드리며, ${name || ""}님의 앞날을 진심으로 응원하겠습니다.\n\n` +
        `KCR 2026 사무국 드림`;
      break;
    default:
      return;
  }

  GmailApp.sendEmail(email, subject, body, { name: "KCR2026 운영요원 모집" });
}

// ── 트리거 등록 (최초 1회 실행) ────────────────────────────────
function setRecruitTriggers() {
  // 기존 트리거 제거
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // onEdit 트리거 등록 (상태 변경 감지)
  ScriptApp.newTrigger("onStatusChange")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log("트리거 등록 완료: onStatusChange (onEdit)");
}
