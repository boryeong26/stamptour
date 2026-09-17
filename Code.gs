/**
 * 가을 보령생태로드 스탬프 투어 - 응모 접수 백엔드 (v4)
 *
 * [이번 버전에서 추가된 것]
 * 1. 이메일 인증번호 발송(requestCode) / 확인(verifyCode)
 * 2. 현장 인증사진을 구글 드라이브에 자동 저장, 링크를 시트에 기록
 * 3. 휴대전화번호/이메일로 접수 여부 조회(checkStatus)
 * 4. 선착순 자동판정 + 중복 응모 방지는 기존과 동일하게 유지
 *
 * [설정 방법]
 * 1. https://sheet.new 로 새 시트 만들기
 * 2. 주소창 URL에서 /d/ 와 /edit 사이의 문자열이 시트 ID
 * 3. 아래 SPREADSHEET_ID = "" 안에 붙여넣기 (필수)
 * 4. 이 코드 전체를 기존 Apps Script 프로젝트에 덮어쓰기
 * 5. 재배포: 배포 → 배포 관리 → 연필 아이콘 → 버전: 새 버전 → 배포
 *    (URL은 그대로 유지되니 HTML은 다시 안 고쳐도 됨)
 *
 * [주의] 인증사진은 구글 드라이브에 저장돼. 처음 실행 시 드라이브 접근 권한도
 * 같이 승인해야 해서, 재배포 후 처음 한 번은 testSendMail 대신 아래
 * testFullFlow 함수를 실행해서 권한 승인 팝업을 띄워주는 게 좋아.
 */

const ADMIN_EMAIL = "whiteoks@koem.or.kr";
const SPREADSHEET_ID = ""; // 여기에 시트 ID 필수로 넣어야 함
const DRIVE_FOLDER_NAME = "보령스탬프투어_인증사진";
const FIRSTCOME_LIMIT = 200; // 선착순 지급 인원
const CODE_TTL_SECONDS = 300; // 인증번호 유효시간 5분

const HEADER = ["접수시각","응모번호","성명","휴대전화번호","이메일","방문인증장소","사진링크","제출시각(참가자기기기준)","선착순 지급여부","접수순번"];

function getSheet() {
  if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID가 설정되지 않았습니다.");
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
  }
  return sheet;
}

function getDriveFolder() {
  const existing = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function findEntryByQuery(query) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const phone = row[3];
    const email = row[4];
    if (phone === query || email === query) {
      return { entryNo: row[1], submittedAt: Utilities.formatDate(new Date(row[0]), "Asia/Seoul", "yyyy-MM-dd HH:mm") };
    }
  }
  return null;
}

function isDuplicate(phone, email) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const rows = sheet.getRange(2, 4, lastRow - 1, 2).getValues(); // 휴대전화번호, 이메일 컬럼
  return rows.some(r => r[0] === phone || r[1] === email);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  try {
    if (action === "requestCode") return handleRequestCode(data);
    if (action === "verifyCode") return handleVerifyCode(data);
    if (action === "checkStatus") return handleCheckStatus(data);
    if (action === "submitEntry") return handleSubmitEntry(data);
    return jsonOut({ ok:false, message:"알 수 없는 요청입니다." });
  } catch (err) {
    return jsonOut({ ok:false, message: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 1) 인증번호 발송 ---------- */
function handleRequestCode(data) {
  const name = (data.name || "").trim();
  const phone = (data.phone || "").trim();
  const email = (data.email || "").trim();

  if (!name || !phone || !email) {
    return jsonOut({ ok:false, message:"이름, 휴대전화번호, 이메일을 모두 입력해 주세요." });
  }

  if (isDuplicate(phone, email)) {
    return jsonOut({ ok:false, duplicate:true });
  }

  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put("otp_" + email, code, CODE_TTL_SECONDS);

  const subject = "[가을 보령생태로드] 이메일 인증번호";
  const body =
`안녕하세요, ${name}님.

가을 보령생태로드 스탬프 투어 참여를 위한 인증번호는 다음과 같습니다.

인증번호: ${code}

이 번호는 5분간 유효합니다.
본인이 요청하지 않았다면 이 메일을 무시해주세요.`;

  MailApp.sendEmail(email, subject, body);
  return jsonOut({ ok:true });
}

/* ---------- 2) 인증번호 확인 ---------- */
function handleVerifyCode(data) {
  const email = (data.email || "").trim();
  const code = (data.code || "").trim();
  const cached = CacheService.getScriptCache().get("otp_" + email);

  if (!cached) {
    return jsonOut({ ok:false, message:"인증번호가 만료됐어요. 다시 받아주세요." });
  }
  if (cached !== code) {
    return jsonOut({ ok:false, message:"인증번호가 올바르지 않아요." });
  }
  CacheService.getScriptCache().remove("otp_" + email);
  return jsonOut({ ok:true });
}

/* ---------- 3) 접수 여부 조회 ---------- */
function handleCheckStatus(data) {
  const query = (data.query || "").trim();
  if (!query) return jsonOut({ found:false });

  const found = findEntryByQuery(query);
  if (found) {
    return jsonOut({ found:true, entryNo: found.entryNo, submittedAt: found.submittedAt });
  }
  return jsonOut({ found:false });
}

/* ---------- 4) 최종 접수 (사진 저장 + 시트 기록 + 관리자 메일) ---------- */
function handleSubmitEntry(data) {
  const name = (data.name || "").trim();
  const phone = (data.phone || "").trim();
  const email = (data.email || "").trim();
  const entryNo = data.entryNo || "";
  const submittedAt = data.submittedAt || "";
  const visitedStations = data.visitedStations || [];
  const photos = data.photos || [];

  if (isDuplicate(phone, email)) {
    return jsonOut({ ok:true, duplicate:true }); // 이미 기록됨 - 참가자에겐 정상 완료로만 안내
  }

  // 사진을 드라이브에 저장하고 링크 수집
  let photoLinks = [];
  try {
    const folder = getDriveFolder();
    photos.forEach(p => {
      if (!p.base64 || p.skipped) {
        photoLinks.push(`${p.stationName}: (사진 없음/테스트)`);
        return;
      }
      const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mimeType || "image/jpeg",
        `${entryNo}_${p.stationId}.jpg`);
      const file = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (shareErr) {}
      photoLinks.push(`${p.stationName}: ${file.getUrl()}`);
    });
  } catch (driveErr) {
    photoLinks.push("(사진 저장 중 오류: " + driveErr.message + ")");
  }

  // 시트 기록 + 선착순 판정
  let firstComeText = "";
  let seqNo = "";
  try {
    const sheet = getSheet();
    const currentCount = sheet.getLastRow() - 1;
    seqNo = currentCount + 1;
    firstComeText = (currentCount < FIRSTCOME_LIMIT) ? `선착순 ${FIRSTCOME_LIMIT}명 해당` : "추첨 대상";

    sheet.appendRow([
      new Date(),
      entryNo,
      name,
      phone,
      email,
      visitedStations.join(", "),
      photoLinks.join(" / "),
      submittedAt,
      firstComeText,
      seqNo
    ]);
  } catch (sheetErr) {
    // 시트 기록 실패해도 메일은 계속 발송
  }

  // 담당자 메일 발송
  try {
    const subject = `[가을 보령생태로드] 스탬프투어 응모 - ${name} (${firstComeText || "접수"})`;
    const body =
`가을 보령생태로드 모바일 스탬프 투어 응모가 접수되었습니다.

응모번호 : ${entryNo}
성명     : ${name}
연락처   : ${phone}
이메일   : ${email}
제출시각 : ${submittedAt}
접수순번 : ${seqNo}번째 (${firstComeText})

[방문 인증 내역 및 사진]
${photoLinks.join("\n")}

※ 이 메일은 시스템에서 자동 발송되었습니다.`;

    MailApp.sendEmail(ADMIN_EMAIL, subject, body);
  } catch (mailErr) {
    return jsonOut({ ok:false, message: mailErr.message });
  }

  return jsonOut({ ok:true, duplicate:false, seqNo: seqNo, firstComeText: firstComeText });
}

/**
 * [테스트용] 편집기에서 이 함수를 직접 실행(▶)해서 인증번호 발송부터
 * 최종 접수(사진 없이)까지 전체 흐름과 권한 승인을 한 번에 확인할 수 있어.
 */
function testFullFlow() {
  const testEmail = Session.getActiveUser().getEmail() || ADMIN_EMAIL;

  const r1 = handleRequestCode({ name:"테스트유저", phone:"010-0000-" + Math.floor(Math.random()*10000), email: testEmail });
  Logger.log("requestCode 결과: " + r1.getContent());

  const cached = CacheService.getScriptCache().get("otp_" + testEmail);
  Logger.log("발급된 코드(테스트 확인용): " + cached);

  const r2 = handleVerifyCode({ email: testEmail, code: cached });
  Logger.log("verifyCode 결과: " + r2.getContent());

  const r3 = handleSubmitEntry({
    name:"테스트유저", phone:"010-0000-9999", email: testEmail,
    entryNo:"TEST" + Date.now(), submittedAt: new Date().toLocaleString('ko-KR'),
    visitedStations:["소황사구","무창포 닭벼슬섬","군헌어촌체험마을"],
    photos:[
      {stationId:"sohwang", stationName:"소황사구", skipped:true},
      {stationId:"dakbyeoseul", stationName:"무창포 닭벼슬섬", skipped:true},
      {stationId:"gunheon", stationName:"군헌어촌체험마을", skipped:true}
    ]
  });
  Logger.log("submitEntry 결과: " + r3.getContent());
}
