// ============================================================
// KCR 2026 만족도 설문 — Apps Script 백엔드
//
// [배포 방법]
// 1. 구글 스프레드시트 새로 만들기 (설문 전용)
// 2. 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
// 3. 저장 → 배포 → 새 배포 → 웹 앱 선택
// 4. 설정:
//    - 설명: KCR2026 만족도 설문
//    - 다음 사용자로 실행: 나(본인)
//    - 액세스 권한: 모든 사용자
// 5. 배포 → 승인 → URL 복사
// 6. 복사한 URL을 survey.html의 GAS_URL 변수에 붙여넣기
// ============================================================

const PM_EMAIL   = "info@kcr2026.com";
const SHEET_NAME = "만족도설문";
const HEADERS    = [
  "타임스탬프", "응답자구분", "언어",
  "소속직종", "참석횟수", "인지경로", "AI친숙도", "통역경험",
  "AI번역인지", "이용수단", "APP편리성",
  "용어정확성", "번역자연스러움", "실시간성", "가독성", "음성번역품질", "음성vs자막도움",
  "스크린번역만족", "효과적방식", "AI지속찬성", "AI개선의견",
  "프로그램구성", "연자전문성", "등록절차", "APP공지", "전시홀", "시설편의성",
  "식사품질", "다과서비스", "FB개선의견",
  "재참석의향", "전반적만족도", "추천의향NPS", "기타건의",
  "이름(추첨)", "연락처(추첨)", "경품유형"
];

// ── GET 요청: 연결 테스트 ─────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "KCR2026 만족도 설문 API 정상 작동 중" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST 요청: 설문 접수 ──────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveSurvey_(data);
    sendAdminAlert_(data);
    updateSurveyDashboard_();

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
function saveSurvey_(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setBackground("#185FA5").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const row = [
    data.timestamp    || "",
    data.respondent   || "",
    data.lang         || "",
    data.q1           || "",
    data.q2           || "",
    data.q3           || "",
    data.q4           || "",
    data.q5           || "",
    data.q6           || "",
    data.q7           || "",
    data.q8           || "",
    data.q9_1         || "",
    data.q9_2         || "",
    data.q9_3         || "",
    data.q9_4         || "",
    data.q9_5         || "",
    data.q10          || "",
    data.q11          || "",
    data.q12          || "",
    data.q13          || "",
    data.q14          || "",
    data.q15          || "",
    data.q16          || "",
    data.q17          || "",
    data.q18          || "",
    data.q19          || "",
    data.q20          || "",
    data.q21          || "",
    data.q22          || "",
    data.q23          || "",
    data.q24          || "",
    data.q25          || "",
    data.q26          || "",
    data.q27          || "",
    data.prize_name   || "",
    data.prize_phone  || "",
    data.prize_type   || ""
  ];

  const lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);
  sheet.autoResizeColumns(1, HEADERS.length);
}

// ── 관리자 알림 이메일 ────────────────────────────────────────
function sendAdminAlert_(data) {
  const type = data.respondent === 'domestic' ? '내국인' : '외국인';
  const subject = `[KCR2026 설문] 신규 응답 — ${type} (${data.q1 || ""})`;
  const body =
    `새로운 설문 응답이 접수되었습니다.\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
    `응답자 유형 : ${type}\n` +
    `소속        : ${data.q1 || ""}\n` +
    `참석 횟수   : ${data.q2 || ""}\n` +
    `AI 친숙도   : ${data.q4 || "-"}\n` +
    `AI 번역 인지: ${data.q6 || ""}\n` +
    `전반적 만족도: ${data.q25 || "-"}\n` +
    `추천 의향   : ${data.q26 || "-"}\n` +
    `재참석 의향 : ${data.q24 || ""}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `전체 응답 현황: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  GmailApp.sendEmail(PM_EMAIL, subject, body, { name: "KCR2026 만족도 설문" });
}

// ── 대시보드 자동 갱신 ────────────────────────────────────────
function updateSurveyDashboard_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SHEET_NAME);
  if (!src) return;

  const data    = src.getDataRange().getValues();
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

  // ── 제목 ──
  dash.getRange("A1").setValue("KCR 2026 만족도 설문 대시보드")
    .setFontSize(15).setFontWeight("bold").setFontColor(BLUE);
  dash.getRange("A2")
    .setValue("최종 업데이트: " + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"))
    .setFontSize(10).setFontColor("#8E8E93");

  // ── 응답 요약 ──
  const total    = rows.length;
  const domestic = rows.filter(r => r[col("응답자구분")] === "domestic").length;
  const foreign  = rows.filter(r => r[col("응답자구분")] === "foreign").length;

  dash.getRange("A4:C4")
    .setValues([["총 응답", "내국인", "외국인"]])
    .setBackground(BLUE).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  dash.getRange("A5:C5")
    .setValues([[total, domestic, foreign]])
    .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");

  // ── 소속별 분포 ──
  dash.getRange("A7").setValue("소속별 분포")
    .setFontWeight("bold").setFontColor(BLUE);
  const affiliations = ["대학교수", "전문의", "전임의", "전공의", "간호사", "연구원", "학생", "기업관계자", "기타"];
  dash.getRange(8, 1, 1, affiliations.length + 1)
    .setValues([["소속", ...affiliations]])
    .setBackground(LBLUE).setFontWeight("bold");
  const affCounts = affiliations.map(a =>
    rows.filter(r => r[col("소속직종")] === a).length
  );
  dash.getRange(9, 1, 1, affiliations.length + 1)
    .setValues([["응답 수", ...affCounts]])
    .setHorizontalAlignment("center");

  // ── Likert 평균 점수 ──
  dash.getRange("A11").setValue("항목별 평균 점수 (5점 만점)")
    .setFontWeight("bold").setFontColor(BLUE);

  const likertCols = [
    ["AI친숙도", "AI친숙도"],
    ["APP편리성", "APP편리성"],
    ["용어정확성", "용어정확성"],
    ["번역자연스러움", "번역자연스러움"],
    ["실시간성", "실시간성"],
    ["가독성", "가독성"],
    ["음성번역품질", "음성번역품질"],
    ["음성vs자막도움", "음성vs자막도움"],
    ["AI지속찬성", "AI지속찬성"],
    ["프로그램구성", "프로그램구성"],
    ["연자전문성", "연자전문성"],
    ["등록절차", "등록절차"],
    ["APP공지", "APP공지"],
    ["전시홀", "전시홀"],
    ["시설편의성", "시설편의성"],
    ["식사품질", "식사품질"],
    ["다과서비스", "다과서비스"],
    ["전반적만족도", "전반적만족도"],
    ["추천의향NPS", "추천의향NPS"]
  ];

  dash.getRange(12, 1, 1, 2)
    .setValues([["항목", "평균"]])
    .setBackground(LBLUE).setFontWeight("bold");

  likertCols.forEach((pair, i) => {
    const colIdx = col(pair[0]);
    const vals   = rows.map(r => parseFloat(r[colIdx])).filter(v => !isNaN(v) && v > 0);
    const avg    = vals.length > 0 ? (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2) : "-";
    dash.getRange(13 + i, 1, 1, 2)
      .setValues([[pair[1], avg]])
      .setBackground(i % 2 === 0 ? GRAY : "#FFFFFF");
  });

  const likertEnd = 13 + likertCols.length;

  // ── Q6 인지여부 분포 ──
  dash.getRange(likertEnd + 1, 1).setValue("AI 번역 인지 분포")
    .setFontWeight("bold").setFontColor(BLUE);
  const q6Opts = ["예", "아니오", "몰랐음"];
  dash.getRange(likertEnd + 2, 1, 1, q6Opts.length + 1)
    .setValues([["응답", ...q6Opts]])
    .setBackground(LBLUE).setFontWeight("bold");
  const q6Counts = q6Opts.map(o => rows.filter(r => r[col("AI번역인지")] === o).length);
  dash.getRange(likertEnd + 3, 1, 1, q6Opts.length + 1)
    .setValues([["건수", ...q6Counts]])
    .setHorizontalAlignment("center");

  // ── Q12 효과적 방식 분포 ──
  dash.getRange(likertEnd + 5, 1).setValue("효과적 방식 분포")
    .setFontWeight("bold").setFontColor(BLUE);
  const q12Opts = ["스크린자막", "전용APP", "음성번역", "병행사용"];
  dash.getRange(likertEnd + 6, 1, 1, q12Opts.length + 1)
    .setValues([["방식", ...q12Opts]])
    .setBackground(LBLUE).setFontWeight("bold");
  const q12Counts = q12Opts.map(o => rows.filter(r => r[col("효과적방식")] === o).length);
  dash.getRange(likertEnd + 7, 1, 1, q12Opts.length + 1)
    .setValues([["건수", ...q12Counts]])
    .setHorizontalAlignment("center");

  // ── Q24 재참석 의향 분포 ──
  dash.getRange(likertEnd + 9, 1).setValue("재참석 의향 분포")
    .setFontWeight("bold").setFontColor(BLUE);
  const q24Opts = ["예", "아니오", "미정"];
  dash.getRange(likertEnd + 10, 1, 1, q24Opts.length + 1)
    .setValues([["응답", ...q24Opts]])
    .setBackground(LBLUE).setFontWeight("bold");
  const q24Counts = q24Opts.map(o => rows.filter(r => r[col("재참석의향")] === o).length);
  dash.getRange(likertEnd + 11, 1, 1, q24Opts.length + 1)
    .setValues([["건수", ...q24Counts]])
    .setHorizontalAlignment("center");

  // ── 최근 응답 목록 (10건) ──
  const recentStart = likertEnd + 13;
  dash.getRange(recentStart, 1).setValue(`최근 응답 (${Math.min(10, total)}건)`)
    .setFontWeight("bold").setFontColor(BLUE);
  const listH = ["접수일시", "유형", "소속", "참석횟수", "전반적만족도", "추천의향", "재참석"];
  dash.getRange(recentStart + 1, 1, 1, listH.length)
    .setValues([listH])
    .setBackground(LBLUE).setFontWeight("bold");

  const recent = rows.slice(-10).reverse();
  if (recent.length > 0) {
    const listData = recent.map(r => [
      r[col("타임스탬프")],
      r[col("응답자구분")] === "domestic" ? "내국인" : "외국인",
      r[col("소속직종")],
      r[col("참석횟수")],
      r[col("전반적만족도")],
      r[col("추천의향NPS")],
      r[col("재참석의향")]
    ]);
    dash.getRange(recentStart + 2, 1, listData.length, listH.length).setValues(listData);
  }

  // ── 경품 응모자 목록 ──
  const prizeStart = recentStart + 14;
  const prizeRows = rows.filter(r => r[col("이름(추첨)")] && r[col("이름(추첨)")].toString().trim());
  dash.getRange(prizeStart, 1).setValue(`경품 응모자 (${prizeRows.length}명)`)
    .setFontWeight("bold").setFontColor(BLUE);
  const prizeH = ["이름", "연락처", "경품유형", "소속", "접수일시"];
  dash.getRange(prizeStart + 1, 1, 1, prizeH.length)
    .setValues([prizeH])
    .setBackground(LBLUE).setFontWeight("bold");

  if (prizeRows.length > 0) {
    const prizeData = prizeRows.map(r => [
      r[col("이름(추첨)")],
      r[col("연락처(추첨)")],
      r[col("경품유형")],
      r[col("소속직종")],
      r[col("타임스탬프")]
    ]);
    dash.getRange(prizeStart + 2, 1, prizeData.length, prizeH.length).setValues(prizeData);
  }

  // 열 폭 조정
  for (let c = 1; c <= 10; c++) dash.setColumnWidth(c, 120);
  dash.autoResizeColumns(1, 10);
}
