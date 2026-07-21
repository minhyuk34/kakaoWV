// ============================================================
//  카카오프렌즈 배부시스템 - Google Apps Script 백엔드
//  붙여넣을 곳: script.google.com → 새 프로젝트
// ============================================================

const SHEET_ID   = 'YOUR_SPREADSHEET_ID'; // ← 스프레드시트 ID
const SHEET_ACCT = '계정';
const SHEET_REQ  = '신청';
const SHEET_STK  = '재고';
const ADMIN_EMAIL = 'minhyuk_jang@worldvision.or.kr';

// ── 스프레드시트에서 열었을 때 수동 갱신 메뉴 추가 ──────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('배부현황')
    .addItem('지금 새로고침', 'generateReport')
    .addSeparator()
    .addItem('수취예정 리마인더 지금 발송', 'sendUpcomingPickupReminder')
    .addItem('수취예정 리마인더 매주 월요일 자동발송 등록', 'setupWeeklyPickupReminderTrigger')
    .addSeparator()
    .addItem('수취일정 캘린더 지금 동기화', 'syncPickupCalendar')
    .addToUi();
}

// ── 시트 초기화 (최초 1회 실행) ──────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 계정 시트 (이메일 컬럼 추가)
  let s = ss.getSheetByName(SHEET_ACCT) || ss.insertSheet(SHEET_ACCT);
  if (s.getLastRow() === 0)
    s.appendRow(['이름', '비밀번호해시', '권한', '이메일', '가입일']);

  // 관리자 계정 자동 생성
  const rows = s.getDataRange().getValues();
  const exists = rows.some(r => r[0] === '관리자');
  if (!exists) {
    const hash = sha256('kakao1234');
    s.appendRow(['관리자', hash, 'admin', '', new Date().toLocaleString('ko-KR')]);
  }

  // 신청 시트 (이메일 컬럼 추가)
  s = ss.getSheetByName(SHEET_REQ) || ss.insertSheet(SHEET_REQ);
  if (s.getLastRow() === 0)
    s.appendRow(['ID', '신청일시', '본부', '팀', '이름', '연락처', '이메일', '사용처',
                 '수취예정일', '사용예정일', '제품목록(JSON)', '총수량', '상태', '처리일시', '관리자메모']);

  // 재고 시트
  s = ss.getSheetByName(SHEET_STK) || ss.insertSheet(SHEET_STK);
  if (s.getLastRow() === 0)
    s.appendRow(['제품번호', '제품명', '재고수량']);

  return '초기화 완료';
}

// ── 라우터 ────────────────────────────────────────────────────
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  let result;
  try {
    const { action } = data;
    if      (action === 'register')      result = register(data);
    else if (action === 'login')         result = login(data);
    else if (action === 'submitRequest') result = submitRequest(data);
    else if (action === 'mergeIntoRequest') result = mergeIntoRequest(data);
    else if (action === 'getRequests')   result = getRequests(data);
    else if (action === 'updateRequest') result = updateRequest(data);
    else if (action === 'updateRequestSchedule') result = updateRequestSchedule(data);
    else if (action === 'getStock')      result = getStock();
    else if (action === 'updateStock')   result = updateStock(data);
    else if (action === 'syncStock')       result = syncStock();
    else if (action === 'approveItems')    result = approveItems(data);
    else if (action === 'updateItemQty')   result = updateItemQty(data);
    else if (action === 'distributeItems') result = distributeItems(data);
    else if (action === 'cancelRequest')   result = cancelRequest(data);
    else if (action === 'cancelItems')       result = cancelItems(data);
    else if (action === 'requestCancelItem') result = requestCancelItem(data);
    else if (action === 'confirmCancelItem') result = confirmCancelItem(data);
    else if (action === 'rejectCancelItem')  result = rejectCancelItem(data);
    else if (action === 'uploadReport')      result = uploadReport(data);
    else if (action === 'submitFeedback')    result = submitFeedback(data);
    else if (action === 'getFeedback')       result = getFeedback();
    else result = { ok: false, error: '알 수 없는 action' };

    // 신청/재고 상태를 바꾸는 액션 이후에는 배부현황 시트를 자동 갱신
    const REPORT_TRIGGER_ACTIONS = [
      'submitRequest', 'updateRequest', 'approveItems', 'distributeItems',
      'cancelRequest', 'cancelItems', 'confirmCancelItem', 'rejectCancelItem',
      'updateRequestSchedule', 'updateItemQty', 'mergeIntoRequest'
    ];
    if (action === 'generateReport') {
      // 프런트에서 수동 새로고침 요청 시 직접 호출 가능
      try { generateReport(); result = { ok: true }; }
      catch (reportErr) { result = { ok: false, error: reportErr.message }; }
    } else if (result && result.ok !== false && REPORT_TRIGGER_ACTIONS.includes(action)) {
      try { generateReport(); }
      catch (reportErr) {
        Logger.log('generateReport 실패: ' + reportErr.message);
        result._reportError = reportErr.message; // 화면에서 확인 가능하도록 원래 결과에 첨부
      }
      try { syncPickupCalendar(); }
      catch (calErr) {
        Logger.log('syncPickupCalendar 실패: ' + calErr.message);
        result._calendarError = calErr.message;
      }
    }
  } catch(err) {
    result = { ok: false, error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: '카카오 배부시스템 API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 회원가입 ─────────────────────────────────────────────────
function register({ name, email, hash }) {
  if (!name || !hash) return { ok: false, error: '이름과 비밀번호를 입력해주세요.' };
  const s = sheet(SHEET_ACCT);
  const rows = s.getDataRange().getValues();
  if (rows.some(r => r[0] === name))
    return { ok: false, error: `"${name}" 이름으로 이미 가입된 계정이 있습니다.` };
  s.appendRow([name, hash, 'team', email || '', new Date().toLocaleString('ko-KR')]);
  return { ok: true, role: 'team', name, email: email || '' };
}

// ── 로그인 ───────────────────────────────────────────────────
function login({ name, hash }) {
  if (!name || !hash) return { ok: false, error: '이름과 비밀번호를 입력해주세요.' };
  const rows = sheet(SHEET_ACCT).getDataRange().getValues().slice(1);
  const acc = rows.find(r => r[0] === name);
  if (!acc) return { ok: false, error: `"${name}" 계정이 없습니다. 회원가입을 먼저 해주세요.` };
  if (acc[1] !== hash) return { ok: false, error: '비밀번호가 올바르지 않습니다.' };
  return { ok: true, name: acc[0], role: acc[2], email: acc[3] || '' };
}

// ── 신청 제출 (신청 즉시 재고 차감) ──────────────────────────
function submitRequest({ dept, team, name, contact, email, reason, pickupDate, useDate, items }) {
  // 동시에 여러 명이 신청할 때 "재고 확인 → 차감" 사이에 끼어들어
  // 둘 다 통과해버리는 경쟁 상태(race condition)를 막기 위한 잠금
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return submitRequest_(arguments[0]);
  } finally {
    lock.releaseLock();
  }
}

function submitRequest_({ dept, team, name, contact, email, reason, pickupDate, useDate, items }) {
  // ── 재고 사전 검증 (서버에서 실시간 확인) ──────────────────
  const stockResult = getStock();
  const insufficient = [];
  items.forEach(item => {
    const num       = String(item.num).padStart(3, '0');
    const available = stockResult.stock[num]?.current ?? stockResult.stock[num] ?? 0;
    if (item.qty > available) {
      insufficient.push(`${item.name} (신청 ${item.qty}개 / 잔여 ${available}개)`);
    }
  });
  if (insufficient.length > 0) {
    return {
      ok: false,
      error: '재고 부족으로 신청할 수 없습니다:\n' + insufficient.join('\n')
    };
  }

  // ── 재고 충분 → 신청 저장 + 차감 ─────────────────────────
  const s = sheet(SHEET_REQ);
  const id = Date.now().toString();
  const totalQty = items.reduce((sum, i) => sum + i.qty, 0);
  const newRow = s.getLastRow() + 1;
  s.appendRow([
    id,
    new Date().toLocaleString('ko-KR'),
    dept, team, name, contact, email || '', reason,
    pickupDate || '', useDate || '',
    JSON.stringify(items),
    totalQty,
    'pending', '', ''
  ]);
  // 수취예정일/사용예정일 셀이 Date 타입으로 자동 변환되지 않도록 텍스트 서식 고정
  s.getRange(newRow, 9).setNumberFormat('@').setValue(formatDateOnly(pickupDate));
  s.getRange(newRow, 10).setNumberFormat('@').setValue(formatDateOnly(useDate));
  items.forEach(item => deductStock(item.num, item.qty, { reqId: id, name, reason: '신청 제출' }));
  return { ok: true, id };
}

// ── 신청 목록 조회 ────────────────────────────────────────────
function getRequests({ name, role }) {
  const rows = sheet(SHEET_REQ).getDataRange().getValues().slice(1);
  const reqs = [];

  rows.forEach(r => {
    if (!r[0]) return; // 빈 행 skip

    // 컬럼 자동 감지: 각 컬럼에서 JSON을 찾아서 위치 파악
    let itemsJson = '[]';
    let status    = '';
    let pickupDate = '';
    let useDate    = '';
    let totalQty   = 0;
    let updatedAt  = '';
    let adminNote  = '';
    let plannedDate = ''; // 배부일 (관리자가 실제 배분한 날짜)

    // r[8]이 JSON이면 구형(13열), r[10]이 JSON이면 신형(15열)
    const r8  = String(r[8]  || '').trim();
    const r10 = String(r[10] || '').trim();

    if (r8.startsWith('[') || r8.startsWith('{')) {
      // 구형
      itemsJson  = r8;
      totalQty   = r[9];
      status     = String(r[10] || '');
      updatedAt  = r[11];
      adminNote  = r[12];
    } else if (r10.startsWith('[') || r10.startsWith('{')) {
      // 신형
      pickupDate = r[8];
      useDate    = r[9];
      itemsJson  = r10;
      totalQty   = r[11];
      status     = String(r[12] || '');
      updatedAt  = r[13];
      adminNote  = r[14];
      plannedDate = r[15] || '';
    } else {
      // 어느 쪽도 JSON이 아닌 경우 — 시트에 수동으로 입력된 행일 가능성이 높음.
      // (제품 목록 JSON이 없어 관리자 화면에서 조용히 사라지는 것을 방지하기 위해 pending으로 노출)
      const colCount = r.filter(c => c !== '' && c !== null && c !== undefined).length;
      status = colCount >= 13 ? String(r[12] || '') : String(r[10] || '');
      if (!status) status = 'pending';
      Logger.log(`⚠️ 신청 행 형식 인식 실패 (제품 목록 JSON 없음) — ID:${r[0]}, 이름:${r[4]}. 시트에서 직접 입력된 행인지 확인 필요.`);
    }

    let items = [];
    try { items = JSON.parse(itemsJson || '[]'); } catch(e) { items = []; }

    reqs.push({
      id: r[0], createdAt: r[1], dept: r[2], team: r[3], name: r[4],
      contact: r[5], email: r[6], reason: r[7],
      pickupDate: formatDateOnly(pickupDate),
      useDate: formatDateOnly(useDate),
      items, totalQty, status, updatedAt, adminNote,
      plannedDate: formatDateOnly(plannedDate)
    });
  });

  const filtered = reqs.filter(r => role === 'admin' || r.name === name);
  return { ok: true, requests: filtered };
}

// ── 수취예정일 / 사용예정일 / 배부일 수정 ──
function updateRequestSchedule({ id, pickupDate, useDate, plannedDate }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    if (isOld) return { ok: false, error: '구형 신청 건은 수취예정일/사용예정일/배부일을 지원하지 않습니다.' };

    // 날짜 셀이 자동으로 Date 타입으로 바뀌지 않도록 텍스트 서식 고정 후 기록
    if (pickupDate !== undefined) {
      s.getRange(i + 1, 8 + 1).setNumberFormat('@').setValue(formatDateOnly(pickupDate)); // 열 8: 수취예정일
    }
    if (useDate !== undefined) {
      s.getRange(i + 1, 9 + 1).setNumberFormat('@').setValue(formatDateOnly(useDate));     // 열 9: 사용예정일
    }
    if (plannedDate !== undefined) {
      s.getRange(i + 1, 15 + 1).setNumberFormat('@').setValue(formatDateOnly(plannedDate)); // 열 15: 배부일
    }

    return { ok: true };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 신청 상태 변경 + 이메일 발송 ─────────────────────────────
function updateRequest({ id, status, adminNote }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    // 구형/신형 자동 감지
    const isOld = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol   = isOld ? 8  : 10;
    const statusCol  = isOld ? 10 : 12;
    const updatedCol = isOld ? 11 : 13;
    const noteCol    = isOld ? 12 : 14;

    s.getRange(i + 1, statusCol  + 1).setValue(status);
    s.getRange(i + 1, updatedCol + 1).setValue(new Date().toLocaleString('ko-KR'));
    if (adminNote !== undefined) s.getRange(i + 1, noteCol + 1).setValue(adminNote);

    const items = JSON.parse(rows[i][itemsCol] || '[]');

    // 반려 시 재고 복구
    if (status === 'rejected') {
      items.forEach(item => restoreStock(item.num, item.qty, { reqId: id, name: rows[i][4], reason: '신청 반려' }));
    }

    // 이메일 발송
    const recipientEmail = rows[i][6];
    const requesterName  = rows[i][4];
    const dept           = rows[i][2];
    const team           = rows[i][3];
    const reason         = rows[i][7];

    if (recipientEmail && (status === 'approved' || status === 'rejected' || status === 'distributed')) {
      const emailItems = items.filter(item => !item.cancelled); // 취소 항목 제외
      sendNotificationEmail(recipientEmail, requesterName, dept, team, reason, emailItems.length > 0 ? emailItems : items, status, adminNote || '');
    }

    return { ok: true };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 이메일 발송 ───────────────────────────────────────────────
function sendNotificationEmail(to, name, dept, team, reason, items, status, adminNote) {
  const statusText = {
    approved:    '✅ 승인되었습니다',
    rejected:    '❌ 반려되었습니다',
    distributed: '📦 배부 완료되었습니다'
  }[status] || status;

  const statusColor = {
    approved:    '#00B894',
    rejected:    '#E17055',
    distributed: '#0984E3'
  }[status] || '#888';

  const itemRows = items.map(i =>
    `<tr style="border-bottom:1px solid #eee">
      <td style="padding:8px 12px">${i.name}</td>
      <td style="padding:8px 12px;text-align:center;font-weight:600">${i.qty}개</td>
    </tr>`
  ).join('');

  const adminNoteHtml = adminNote
    ? `<div style="background:#FFF3CD;border-radius:8px;padding:12px 16px;margin-top:16px;font-size:14px">
        💬 <b>관리자 메모:</b> ${adminNote}
       </div>`
    : '';

  const html = `
  <div style="font-family:'Apple SD Gothic Neo',sans-serif;max-width:560px;margin:0 auto">
    <div style="background:#3C1E1E;padding:24px 32px;border-radius:12px 12px 0 0">
      <h2 style="color:#FEE500;margin:0;font-size:20px">🐾 카카오프렌즈 배부시스템</h2>
    </div>
    <div style="background:#fff;padding:28px 32px;border:1px solid #eee;border-top:none">
      <p style="font-size:16px;font-weight:700;color:#1A1A1A">${name}님의 신청이</p>
      <p style="font-size:24px;font-weight:800;color:${statusColor};margin:8px 0 20px">${statusText}</p>

      <div style="background:#FAFAFA;border-radius:8px;padding:12px 16px;margin-bottom:20px;font-size:14px">
        <div style="margin-bottom:6px"><span style="color:#888">소속</span> &nbsp;<b>${dept} ${team}</b></div>
        <div><span style="color:#888">사용처</span> &nbsp;<b>${reason}</b></div>
      </div>

      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <thead>
          <tr style="background:#3C1E1E">
            <th style="padding:10px 12px;text-align:left;color:#FEE500">제품명</th>
            <th style="padding:10px 12px;text-align:center;color:#FEE500">수량</th>
          </tr>
        </thead>
        <tbody>${itemRows}</tbody>
      </table>

      ${adminNoteHtml}

      <p style="font-size:12px;color:#aaa;margin-top:24px;border-top:1px solid #eee;padding-top:16px">
        본 메일은 카카오프렌즈 배부시스템에서 자동 발송되었습니다.
      </p>
    </div>
  </div>`;

  MailApp.sendEmail({
    to:      to,
    subject: `[카카오프렌즈 배부] ${name}님 신청이 ${statusText}`,
    htmlBody: html
  });
}

// ── 재고 조회 ─────────────────────────────────────────────────
// ── 재고 조회 ─────────────────────────────────────────────────
// 재고 시트 구조: A=제품번호, B=제품명, C=현재재고(잔여), D=원래재고
// D열이 비어있으면 C열을 원래재고로 취급
function getStock() {
  const rows = sheet(SHEET_STK).getDataRange().getValues().slice(1);
  const stock = {};
  rows.forEach(r => {
    if (!r[0]) return;
    const num = String(r[0]).padStart(3, '0');
    const current  = Number(r[2]) || 0;
    const original = Number(r[3]) || current; // D열 없으면 C열을 원래재고로
    stock[num] = { current, original };
  });
  return { ok: true, stock };
}

// ── 재고 수정 (관리자 직접 수정) ──────────────────────────────
function updateStock({ num, qty }) {
  const s = sheet(SHEET_STK);
  const rows = s.getDataRange().getValues();
  const paddedNum = String(num).padStart(3, '0');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).padStart(3, '0') === paddedNum) {
      s.getRange(i + 1, 3).setValue(qty);
      // 원래재고(D열)가 비어있으면 함께 설정
      if (!rows[i][3]) s.getRange(i + 1, 4).setValue(qty);
      return { ok: true };
    }
  }
  return { ok: false, error: '제품을 찾을 수 없습니다.' };
}

// ── 재고 재계산 (구글시트 행 삭제 후 동기화) ──────────────────
// 활성 신청(pending/approved/distributed) 기준으로 현재재고를 재계산
function syncStock() {
  const s = sheet(SHEET_STK);
  const stockRows = s.getDataRange().getValues();

  // 활성 신청의 제품별 사용량 합산 (구형/신형 자동 감지)
  const reqRows = sheet(SHEET_REQ).getDataRange().getValues().slice(1);
  const usedQty = {};
  reqRows.forEach(r => {
    if (!r[0]) return;
    // r[8]이 JSON이면 구형(13열), 아니면 신형(15열)
    const isOld    = String(r[8]).trim().startsWith('[') || String(r[8]).trim().startsWith('{');
    const itemsCol = isOld ? 8  : 10;
    const statusCol= isOld ? 10 : 12;
    const status   = r[statusCol];
    if (['pending', 'approved', 'distributed'].includes(status)) {
      try {
        const items = JSON.parse(r[itemsCol] || '[]');
        items.forEach(item => {
          if (item.cancelled) return; // 취소된 항목 제외
          const n = String(item.num).padStart(3, '0');
          usedQty[n] = (usedQty[n] || 0) + Number(item.qty);
        });
      } catch(e) {}
    }
  });

  // 잔여재고(C열) = 원래재고(D열) - 배분수량
  // 배분수량(E열) = 활성 신청 합산
  for (let i = 1; i < stockRows.length; i++) {
    if (!stockRows[i][0]) continue;
    const num = String(stockRows[i][0]).padStart(3, '0');

    // D열(원래재고) 없으면 C열 값을 원래재고로 저장
    const original = Number(stockRows[i][3]) || Number(stockRows[i][2]) || 0;
    if (!stockRows[i][3] && stockRows[i][2]) {
      s.getRange(i + 1, 4).setValue(original);
    }

    const used    = usedQty[num] || 0;
    const current = Math.max(0, original - used);

    s.getRange(i + 1, 3).setValue(current);  // C열: 잔여재고
    s.getRange(i + 1, 5).setValue(used);     // E열: 배분수량
  }
  return { ok: true, msg: '재고 재계산 완료' };
}

// ── 재고 차감 / 복구 ──────────────────────────────────────────
function deductStock(num, qty, meta)  { adjustStock(num, -qty, meta); }
function restoreStock(num, qty, meta) { adjustStock(num, +qty, meta); }

// ── 변경이력 기록 ────────────────────────────────────────────
// 재고가 바뀌는 모든 지점(adjustStock/adjustStockAllowNegative)에서 공통으로 호출되어
// "언제, 어떤 신청 때문에, 왜, 얼마나" 바뀌었는지 별도 시트에 남긴다.
function logStockChange(meta, num, delta, afterQty) {
  if (!meta) return; // 사유가 없는 호출(예: 테스트)은 기록하지 않음
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let s = ss.getSheetByName('변경이력');
    if (!s) {
      s = ss.insertSheet('변경이력');
      s.appendRow(['일시', '제품번호', '신청ID', '신청인', '변경사유', '증감', '변경후재고']);
      s.getRange(1, 1, 1, 7).setBackground('#3C1E1E').setFontColor('#FEE500').setFontWeight('bold');
      s.setFrozenRows(1);
      s.setColumnWidth(5, 220);
    }
    s.appendRow([
      new Date().toLocaleString('ko-KR'),
      String(num).padStart(3, '0'),
      meta.reqId || '',
      meta.name || '',
      meta.reason || '',
      (delta > 0 ? '+' : '') + delta,
      afterQty
    ]);
  } catch (e) {
    Logger.log('변경이력 기록 실패: ' + e.message);
  }
}

function adjustStock(num, delta, meta) {
  const s = sheet(SHEET_STK);
  const rows = s.getDataRange().getValues();
  const paddedNum = String(num).padStart(3, '0');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).padStart(3, '0') === paddedNum) {
      const cur = Number(rows[i][2]) || 0;
      const next = cur + delta;
      if (next < 0) {
        // 재고보다 많이 차감되는 경우 — 배부현황 리포트(JSON 원본 기준)와 재고 시트가
        // 어긋나는 원인이 되므로 0으로 조용히 밀어넣지 않고 반드시 로그를 남긴다.
        Logger.log(`⚠️ 재고 음수 발생 방지: 제품 ${paddedNum} 현재 ${cur} + 변화량 ${delta} = ${next} → 0으로 조정. syncStock() 실행 권장.`);
      }
      const finalVal = Math.max(0, next);
      s.getRange(i + 1, 3).setValue(finalVal);
      logStockChange(meta, paddedNum, delta, finalVal);
      return;
    }
  }
}

// 관리자가 신청 수량을 직접 수정할 때 사용 — 재고 부족 상태를 감추지 않고
// 음수(마이너스 표시)까지 그대로 반영해 실제 부족분을 화면에서 바로 알 수 있게 한다
function adjustStockAllowNegative(num, delta, meta) {
  const s = sheet(SHEET_STK);
  const rows = s.getDataRange().getValues();
  const paddedNum = String(num).padStart(3, '0');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).padStart(3, '0') === paddedNum) {
      const cur = Number(rows[i][2]) || 0;
      const next = cur + delta;
      s.getRange(i + 1, 3).setValue(next);
      logStockChange(meta, paddedNum, delta, next);
      return next;
    }
  }
  return null;
}

function currentStockOf(num) {
  const paddedNum = String(num).padStart(3, '0');
  const rows = sheet(SHEET_STK).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).padStart(3, '0') === paddedNum) return Number(rows[i][2]) || 0;
  }
  return null;
}

// 관리자가 수량을 줄이면서 "다음에 배부"를 선택했을 때, 같은 신청인 앞으로
// 작업일로부터 한 달 뒤 수취예정일/사용예정일을 가진 신청을 새로 만든다.
function createFollowUpRequest({ dept, team, name, contact, email, originalReqId }, item) {
  const s = sheet(SHEET_REQ);
  const id = Date.now().toString();
  const followUp = new Date();
  followUp.setMonth(followUp.getMonth() + 1);
  const dateStr = Utilities.formatDate(followUp, 'Asia/Seoul', 'yyyy-MM-dd');

  const items = [{ num: item.num, name: item.name, qty: item.qty, price: item.price || 0 }];
  const newRow = s.getLastRow() + 1;
  s.appendRow([
    id, new Date().toLocaleString('ko-KR'),
    dept || '', team || '', name || '', contact || '', email || '',
    `[자동생성] 수량조정 이월 (원신청 ${originalReqId})`,
    dateStr, dateStr,
    JSON.stringify(items), item.qty,
    'pending', '', ''
  ]);
  s.getRange(newRow, 9).setNumberFormat('@').setValue(dateStr);
  s.getRange(newRow, 10).setNumberFormat('@').setValue(dateStr);
  deductStock(item.num, item.qty, {
    reqId: id, name,
    reason: `수량조정 이월 신청 생성(원신청 ${originalReqId}): ${item.name} +${item.qty}`
  });
  return id;
}

// ── 신청 병합 ────────────────────────────────────────────────
// 수취예정일·사용예정일이 같은 기존 신청에 새 항목들을 합쳐 넣는다.
function mergeIntoRequest({ existingId, items }) {
  const stockResult = getStock();
  const insufficient = [];
  items.forEach(item => {
    const num = String(item.num).padStart(3, '0');
    const available = stockResult.stock[num]?.current ?? 0;
    if (item.qty > available) {
      insufficient.push(`${item.name} (신청 ${item.qty}개 / 잔여 ${available}개)`);
    }
  });
  if (insufficient.length > 0) {
    return { ok: false, error: '재고 부족으로 합칠 수 없습니다:\n' + insufficient.join('\n') };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet(SHEET_REQ);
    const rows = s.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(existingId)) continue;

      const isOld = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
      if (isOld) return { ok: false, error: '구형 신청 건과는 합칠 수 없습니다.' };

      const itemsCol = 10, totalQtyCol = 11, statusCol = 12, updatedCol = 13;
      const status = String(rows[i][statusCol] || '');
      if (status === 'cancelled' || status === 'rejected' || status === 'distributed') {
        return { ok: false, error: '이미 처리 완료되었거나 취소된 신청과는 합칠 수 없습니다.' };
      }

      let existingItems = [];
      try { existingItems = JSON.parse(rows[i][itemsCol] || '[]'); } catch (e) {}

      items.forEach(newItem => {
        const match = existingItems.find(it =>
          String(it.num).padStart(3, '0') === String(newItem.num).padStart(3, '0') && !it.cancelled
        );
        if (match) {
          match.qty = (Number(match.qty) || 0) + Number(newItem.qty);
        } else {
          existingItems.push({
            num: newItem.num, name: newItem.name, qty: newItem.qty,
            price: newItem.price || 0, detail: newItem.detail || ''
          });
        }
      });

      items.forEach(item => deductStock(item.num, item.qty, {
        reqId: existingId, name: rows[i][4], reason: `신청 합치기: ${item.name} +${item.qty}`
      }));

      const activeTotal = existingItems.filter(it => !it.cancelled).reduce((s2, it) => s2 + (Number(it.qty) || 0), 0);
      s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(existingItems));
      s.getRange(i + 1, totalQtyCol + 1).setValue(activeTotal);
      s.getRange(i + 1, updatedCol + 1).setValue(new Date().toLocaleString('ko-KR'));

      return { ok: true, id: existingId };
    }
    return { ok: false, error: '기존 신청을 찾을 수 없습니다.' };
  } finally {
    lock.releaseLock();
  }
}

// ── 관리자 - 신청 항목 수량 직접 수정 ─────────────────────────
// 수량이 줄면 그만큼 재고 복구, 늘면 그만큼 재고 추가 차감.
// 결과 재고가 0 이하가 되면 stockWarning:true 로 알려 프런트에서 경고 팝업을 띄우게 한다.
function updateItemQty({ id, idx, qty, adminName, deferReduced }) {
  qty = Number(qty);
  if (isNaN(qty) || qty < 0) return { ok: false, error: '올바른 수량을 입력하세요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet(SHEET_REQ);
    const rows = s.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(id)) continue;

      const isOld = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
      if (isOld) return { ok: false, error: '구형 신청 건은 수량 수정을 지원하지 않습니다.' };

      const itemsCol = 10, totalQtyCol = 11, updatedCol = 13;
      let items = [];
      try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch (e) {}

      if (!items[idx]) return { ok: false, error: '항목을 찾을 수 없습니다.' };
      if (items[idx].cancelled)   return { ok: false, error: '취소된 항목은 수량을 수정할 수 없습니다.' };
      if (items[idx].distributed) return { ok: false, error: '이미 배부된 항목은 수량을 수정할 수 없습니다.' };

      const oldQty = Number(items[idx].qty) || 0;
      const delta  = qty - oldQty; // 늘어난 만큼 재고 추가 차감(음수), 줄어든 만큼 재고 복구(양수)
      const productNum  = items[idx].num;
      const productName = items[idx].name;
      const productPrice= items[idx].price || 0;

      let deferredReqId = null;
      if (delta < 0 && deferReduced) {
        // 줄어든 만큼 재고 복구 후, 같은 신청인 앞으로 한 달 뒤 신청을 새로 만들어 이월
        const reducedAmt = -delta;
        adjustStockAllowNegative(productNum, reducedAmt, {
          reqId: id, name: rows[i][4],
          reason: `수량 감소분 이월(${adminName || '관리자'}): ${productName} -${reducedAmt} → 다음 신청으로`
        });
        deferredReqId = createFollowUpRequest({
          dept: rows[i][2], team: rows[i][3], name: rows[i][4],
          contact: rows[i][5], email: rows[i][6], originalReqId: id
        }, { num: productNum, name: productName, qty: reducedAmt, price: productPrice });
      } else {
        adjustStockAllowNegative(productNum, -delta, {
          reqId: id, name: rows[i][4],
          reason: `수량 수정(${adminName || '관리자'}): ${productName} ${oldQty}→${qty}`
        });
      }

      items[idx].qty = qty;
      const activeTotal = items.filter(it => !it.cancelled).reduce((s2, it) => s2 + (Number(it.qty) || 0), 0);

      s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items));
      s.getRange(i + 1, totalQtyCol + 1).setValue(activeTotal);
      s.getRange(i + 1, updatedCol + 1).setValue(new Date().toLocaleString('ko-KR'));

      const newStock = currentStockOf(productNum);
      return {
        ok: true,
        newStock,
        stockWarning: newStock !== null && newStock <= 0,
        productName,
        deferredReqId
      };
    }
    return { ok: false, error: '신청을 찾을 수 없습니다.' };
  } finally {
    lock.releaseLock();
  }
}

// ── 이메일 테스트 (Apps Script에서 직접 실행해서 권한 승인) ────
// ── 재고 진단 (Apps Script에서 직접 실행해서 확인) ──────────────
function diagnoseStock() {
  const s = sheet(SHEET_STK);
  const rows = s.getDataRange().getValues();

  Logger.log('=== 재고 시트 진단 ===');
  Logger.log('총 행 수: ' + rows.length);
  Logger.log('헤더: ' + JSON.stringify(rows[0]));

  if (rows.length <= 1) {
    Logger.log('❌ 재고 시트에 데이터가 없습니다!');
    return;
  }

  // 첫 5개 행 확인
  Logger.log('--- 데이터 샘플 (첫 5행) ---');
  for (let i = 1; i <= Math.min(5, rows.length - 1); i++) {
    Logger.log(`행${i}: A열="${rows[i][0]}" B열="${rows[i][1]}" C열(현재재고)=${rows[i][2]} D열(원래재고)=${rows[i][3]}`);
  }

  // 002번 제품 찾기 테스트
  const testNum = '002';
  const found = rows.find(r => String(r[0]).padStart(3,'0') === testNum);
  if (found) {
    Logger.log(`✅ 002번 제품 발견: ${found[1]}, 현재재고=${found[2]}`);
  } else {
    Logger.log(`❌ 002번 제품을 찾지 못함. A열 값들: ${rows.slice(1,6).map(r=>r[0]).join(', ')}`);
  }
}

// ── 재고 강제 초기화 (엑셀 데이터를 구글시트에 직접 입력) ────────
// 이 함수를 실행하면 재고 시트에 헤더+원래재고+현재재고가 설정됨
function initStockFromData() {
  const STOCK = {"001":3,"002":158,"003":1292,"004":8,"005":14,"006":917,"007":4,"008":79,"009":107,"011":1100,"012":800,"013":500,"014":400,"015":1300,"016":1200,"017":1100,"018":400,"019":1000,"020":1400,"021":500,"022":700,"023":66,"024":43,"025":700,"026":700,"027":800,"028":1200,"029":6300,"030":10300,"031":14600,"032":70,"033":135,"034":151,"035":500,"036":500,"037":900,"038":300,"039":1400,"040":1800,"041":1300,"042":2563,"043":1500,"044":900,"045":800,"046":1600,"047":500,"048":250,"049":500,"050":300,"051":1000,"052":400,"053":1200,"054":1400,"055":1100,"056":14,"057":470,"058":356,"059":203,"060":459,"061":111,"062":48,"063":200,"064":71,"065":117,"066":17,"067":90,"068":400,"069":900,"070":200,"071":200,"072":400,"073":800,"074":400,"075":400,"076":40,"077":400,"078":200,"079":500,"080":100,"081":400,"082":400,"083":100,"084":800,"085":100,"086":377,"087":100,"088":100,"089":100,"090":50,"091":800,"092":500,"093":100,"094":300,"095":200,"096":100,"097":800,"098":100,"099":80,"100":100,"101":300,"102":400,"103":100,"104":100,"105":300,"106":400,"107":500,"108":600,"109":200,"110":1400,"111":200,"112":500,"113":390,"114":2000,"115":300,"116":200,"117":400,"118":500,"119":300,"120":600,"121":200,"122":200,"123":1000,"124":800,"125":800,"126":100,"127":600,"128":300,"129":1300,"130":400,"131":100,"132":1200,"133":1380,"134":1450,"135":200,"136":500,"137":400,"138":200,"139":300,"140":600,"141":200,"142":8300,"143":168,"144":10000};

  const s = sheet(SHEET_STK);
  s.clearContents();

  // 헤더 (5열: 번호, 제품명, 잔여재고, 원래재고, 배분수량)
  s.appendRow(['제품번호', '제품명', '잔여재고', '원래재고', '배분수량']);

  // 데이터 입력
  const NAMES = {"001": "넷플릭스 오징어게임X프렌즈 플레이어 키링_춘식이", "002": "넷플릭스 오징어게임X프렌즈 영희키링(뉴코스튬)_춘식이", "003": "넷플릭스 오징어게임X프렌즈 투명스티커_춘식이", "004": "넷플릭스 오징어게임X카카오 프렌즈 영희 엽서_춘식이", "005": "넷플릭스 오징어게임X카카오 프렌즈 중형인형_춘식이", "006": "넷플릭스 오징어게임X카카오 프렌즈 LED 키링_춘식이", "007": "춘식이생일파티 실리콘참", "008": "춘식이 포토 형태 아크릴 키링 후드티 춘식이", "009": "쬬르디 미니파우치_뽀짝쬬", "010": "포슬인형 코스튬_명화시리즈_고흐", "011": "골골즈 스마트폰 스트랩_박밤이", "012": "골골즈  스마트폰 스트랩_김콩이", "013": "골골즈  안경 케이스_박밤이", "014": "골골즈 안경 케이스_김콩이", "015": "쬬르디 실리콘 러기지택_어쩔쬬&철벽쬬", "016": "쬬르디 실리콘 러기지택_할미쬬&말티쬬", "017": "포슬인형 코스튬_낚시", "018": "포슬인형 코스튬_테니스", "019": "힙스터 거북이 데코인형_라이언", "020": "힙스터 거북이 데코인형_춘식이", "021": "김다예X카카오프렌즈 코스터_춘식이", "022": "김다예X카카오프렌즈 카드(투게더)_춘식이", "023": "골골즈 차량용 목쿠션_박밤이", "024": "골골즈 차량용 목쿠션_김콩이", "025": "춘식버스 데코인형 귀족영애_춘식이", "026": "갱얼쥐 닥스훈트 인형_춘식이", "027": "갱얼쥐 푸들 인형_춘식이", "028": "고구마 요정 코듀로이 인형_춘식이", "029": "쬬르디 랜덤 펠트 키링인형_쬬르디", "030": "쬬르디 랜덤 북마크_쬬르디", "031": "쬬르디 랜덤 실리콘참_쬬르디", "032": "세이치즈 와펜 헤어밴드 세트 라이언", "033": "골골즈 얼굴쿠션_박밤이", "034": "골골즈 얼굴쿠션_김콩이", "035": "세비지 맥세이프 카드홀더 골골즈_프렌즈", "036": "러브 러브 스탠딩 자석판 인형_춘식이", "037": "골골즈 투명파우치_골골즈", "038": "골골즈 미니키링인형 스포티_김콩이", "039": "미니거울 스마트폰참키링 행운_라이언", "040": "미니거울 스마트폰참키링 하트_춘식이", "041": "페이스 체인지 피규어_춘식이", "042": "무빙피규어 춘데렐라_춘식이", "043": "쬬르디 미니피규어키링_탕후루쬬", "044": "김다예X프렌즈 카드커버 스티커 세트_라&춘", "045": "김다예X프렌즈 이너피스 부클스티커_춘식이", "046": "시골집 핫팩 목베개_춘식이", "047": "포슬인형 코스튬_춘식버스 귀족영애", "048": "오늘의요정 풍경 능이요정_춘식이", "049": "오늘의요정 인형 반반요정_춘식이", "050": "오늘의요정 인형 능이요정_춘식이", "051": "피규어꾸미기세트 꿀벌_춘식이", "052": "포슬인형 하우스_라이언 테이블&스툴SET", "053": "쬬르디 차량용 미니 데코 피규어 세트_뽀짝쬬", "054": "쬬르디 차량 송풍구 피규어_할미쬬", "055": "눈사람이 된 뽀글 인형_춘식이", "056": "별별춘식 메탈 키링 마법사", "057": "라이즈X프렌즈 틴케이스 스티커팩", "058": "라이즈X프렌즈 DIY 아크릴 키링_멍룡이 (앤톤)", "059": "라이즈X프렌즈 DIY 아크릴 키링_똘병 (소희)", "060": "라이즈X프렌즈 DIY 아크릴 키링_우락밤 (성찬)", "061": "라이즈X프렌즈 DIY 아크릴 키링_송용돌이 (은석)", "062": "라이즈X프렌즈 DIY 아크릴 키링_리즈코 (쇼타로)", "063": "포슬인형 코스튬_루돌프", "064": "해피 가드닝_인형_튜브", "065": "HBD 프린세스 피규어_춘식이", "066": "키즈 애착인형 무찡", "067": "오해피치데이 폰케이스_S22_라이언", "068": "유캔두잇 시간관리 타이머_춘식이", "069": "쬬르디 차량 데코인형_말티쬬", "070": "포슬인형 하우스_춘식이 테이블&스툴SET", "071": "러브 허그미 스탠딩 자석판 인형_라이언", "072": "메론빵 거북이 청소솔_라이언", "073": "골골즈 투명스티커 블루_골골즈", "074": "포슬인형 코스튬_가디건&비니", "075": "쬬르디 실리콘 동전지갑_뽀짝쬬", "076": "춘식버스 PVC 카드홀더_춘식이", "077": "쬬르디 미니파우치_어쩔쬬", "078": "포슬인형 코스튬_케이프&모자", "079": "쬬르디 인형_누구시쬬", "080": "포슬인형 하우스_죠르디 테이블&스툴SET", "081": "쬬르디 차량용 도어가드_말티쬬", "082": "갱얼쥐 말티즈 인형_춘식이", "083": "볼 발그레 포슬 키링인형_리틀무지", "084": "달콤하구마 향기나는 키링인형_춘식이", "085": "별별춘식 드레스업 인형 마스터", "086": "쬬르디 인형_뽀짝쬬", "087": "포슬인형 코스튬_후드&선글라스", "088": "포슬인형 코스튬_벚꽃놀이", "089": "춘식버스 데코인형 냥아치_춘식이", "090": "레이지선데이 스트링 파우치_라이언&튜브", "091": "포슬인형 코스튬-배쓰타임세트", "092": "쬬르디 인형_말티쬬", "093": "쬬르디 데코스티커 피곤_군인쬬", "094": "세비지 중형인형_춘식이", "095": "포슬인형 코스튬_니트&크로스백", "096": "쬬르디 미니파우치_철벽쬬", "097": "쬬르디 인형뽑기 아크릴키링", "098": "썸머홀리데이 투명우산_라이언", "099": "CHOONHOUSE 쿠션 거울_춘식이", "100": "프로포즈 팔찌인형_춘식이", "101": "김다예X프렌즈 투명스티커_라이언&춘식이", "102": "트윙클스타 스탠딩 자석판 인형_리틀튜브", "103": "별별춘식 PVC 반짝이 키링 마스터", "104": "Friendzoo 판다 바디 필로우_리틀어피치", "105": "포슬인형 코스튬-썸머 튜브", "106": "골골즈 미니 헤어핀 4P_박밤이", "107": "포근한 플란넬 담요_리틀어피치", "108": "포근한 플란넬 담요_리틀라이언", "109": "골골즈 미니키링인형 파자마_박밤이", "110": "쬬르디 미니피규어키링_어쩔쬬", "111": "오늘부터갓생 포토키링 아자아자_춘식이", "112": "쬬르디 미니파우치_말티쬬", "113": "쬬르디 미니피규어키링_긍정쬬", "114": "행운만땅 송풍구 피규어_춘식이", "115": "쬬르디 차량용 목쿠션_말티쬬", "116": "쬬르디 리버서블 키링인형 죠르디_뽀짝쬬", "117": "춘식버스 미니파우치키링 신입생_춘식이", "118": "굿바이윙윙 향기나는 인형_춘식이", "119": "골골즈 투명스티커 핑크_골골즈", "120": "롱 멀티 도어가드 4P", "121": "오늘부터갓생 포토키링 넌예뻐_라&어", "122": "오늘부터갓생 포토키링 유캔두잇_춘식이", "123": "쬬르디 미니피규어키링_복어쬬", "124": "쬬르디 미니피규어키링_철벽쬬", "125": "조이스틱 듀얼 주차번호판_춘식이", "126": "쬬르디 차량용 목쿠션_할미쬬", "127": "트윙클스타 스탠딩 자석판 인형_리틀라이언", "128": "쬬르디 안심 톡톡 주차번호판_말티쬬", "129": "춘식버스 무빙키링", "130": "쬬르디 미니피규어키링_뽀짝쬬", "131": "도도도춘식이 워터볼 횃불_춘식이", "132": "세비지 스트랩 키링(핸드폰)_춘식이", "133": "세비지 스트랩 키링(가방)_춘식이", "134": "세비지 스트랩 키링(카메라)_춘식이", "135": "김다예X카카오프렌즈 카드(쇼파)_라이언&춘식이", "136": "프렌즈in제주 쬬르디 마그넷_누구시쬬", "137": "프렌즈in제주 쬬르디 마그넷_할미쬬", "138": "프렌즈in제주 쬬르디 마그넷_복어쬬", "139": "춘식버스 퍼레이드 브릭피규어 망나니_춘식이", "140": "힙스터 거북이 후크_춘식이", "141": "쬬르디 미니피규어키링_말티쬬", "142": "프렌즈in제주 랜덤피규어_쬬르디", "143": "멀티 타이머_춘식이", "144": "냥냥특집 츈켄슈타인 퍼스널컬러 랜덤 액정 클리너_춘식이 (9종)"};
  const rows = [];
  Object.entries(STOCK).sort((a,b)=>Number(a[0])-Number(b[0])).forEach(([num, qty]) => {
    const paddedNum = num.padStart(3,'0');
    rows.push([num, NAMES[paddedNum] || '', qty, qty, 0]);
  });
  s.getRange(2, 1, rows.length, 5).setValues(rows);

  // 신청 데이터 기반으로 잔여재고 + 배분수량 재계산
  syncStock();

  Logger.log('재고 초기화 완료: ' + Object.keys(STOCK).length + '개 품목');
}

// ── 재고 차감 직접 테스트 ────────────────────────────────────
// 002번 제품에서 1개 차감 → 시트에서 직접 확인
function testDeductStock() {
  Logger.log('=== 재고 차감 테스트 시작 ===');

  // 차감 전
  const before = getStock().stock['002'];
  Logger.log('차감 전 002번: ' + JSON.stringify(before));

  // 1개 차감
  adjustStock('002', -1);

  // 차감 후
  const after = getStock().stock['002'];
  Logger.log('차감 후 002번: ' + JSON.stringify(after));

  if (JSON.stringify(before) === JSON.stringify(after)) {
    Logger.log('❌ 차감 실패 - 값이 변하지 않음');
  } else {
    Logger.log('✅ 차감 성공! 구글시트 재고탭 C열 확인');
    // 원복
    adjustStock('002', +1);
    Logger.log('(테스트 후 원복 완료)');
  }
}

// ── 전체 진단 ────────────────────────────────────────────────
// ── 항목별 승인 처리 ──────────────────────────────────────────
function approveItems({ id, indices }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld    = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol = isOld ? 8 : 10;
    const statusCol= isOld ? 10 : 12;

    let items = [];
    try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch(e) { items = []; }

    indices.forEach(idx => { if (items[idx] && !items[idx].cancelled) items[idx].approved = true; });

    // 취소된 항목은 제외하고, 나머지가 모두 승인되면 전체 승인
    const activeItems  = items.filter(item => !item.cancelled);
    const allApproved  = activeItems.length > 0 && activeItems.every(item => item.approved);
    const prevStatus   = String(rows[i][statusCol] || '');
    const newStatus    = allApproved ? 'approved' : prevStatus;

    s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items));
    s.getRange(i + 1, statusCol + 1).setValue(newStatus);
    s.getRange(i + 1, (isOld ? 11 : 13) + 1).setValue(new Date().toLocaleString('ko-KR'));

    // 전체 승인 완료 시 이메일 발송 (활성 항목만 포함)
    if (allApproved && prevStatus !== 'approved') {
      const recipientEmail = rows[i][6];
      if (recipientEmail) {
        sendNotificationEmail(recipientEmail, rows[i][4], rows[i][2], rows[i][3], rows[i][7], activeItems, 'approved', '');
      }
    }

    return { ok: true, status: newStatus };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 항목별 배부 처리 ──────────────────────────────────────────
function distributeItems({ id, distributedIndices, distributeDate, distributeMethod, dateMap, methodMap }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld    = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol = isOld ? 8 : 10;
    const statusCol= isOld ? 10 : 12;

    let items = [];
    try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch(e) { items = []; }

    // 선택되지 않은 항목 → 재고 복구
    items.forEach((item, idx) => {
      const wasDistributed = item.distributed || false;
      const willDistribute = distributedIndices.includes(idx);
      if (wasDistributed && !willDistribute) {
        // 이전에 배부됐는데 이번에 해제 → 복구
        restoreStock(item.num, item.qty, { reqId: id, name: rows[i][4], reason: '배부 해제(재배정)' });
      }
      item.distributed = willDistribute;
      if (willDistribute) {
        // 항목별 날짜/방법 우선, 없으면 전체 공통값 사용
        const d = (dateMap   && dateMap[idx])   || distributeDate   || '';
        const m = (methodMap && methodMap[idx]) || distributeMethod || '';
        if (d) item.distributeDate   = d;
        if (m) item.distributeMethod = m;
      }
    });

    // 전체 배부됐으면 status=distributed, 일부만이면 status=approved 유지
    const allDone = items.every(item => item.distributed);
    const newStatus = allDone ? 'distributed' : 'approved';

    s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items));
    s.getRange(i + 1, statusCol + 1).setValue(newStatus);
    s.getRange(i + 1, (isOld ? 11 : 13) + 1).setValue(new Date().toLocaleString('ko-KR'));

    // 이메일 발송 (전체 배부완료 시)
    if (allDone) {
      const recipientEmail = rows[i][6];
      if (recipientEmail) {
        sendNotificationEmail(recipientEmail, rows[i][4], rows[i][2], rows[i][3], rows[i][7], items, 'distributed', '');
      }
    }

    return { ok: true, status: newStatus };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 신청 취소 (재고 복구) ─────────────────────────────────────
function cancelRequest({ id }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld    = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol = isOld ? 8 : 10;
    const statusCol= isOld ? 10 : 12;

    const status = String(rows[i][statusCol] || '');
    if (status === 'distributed') return { ok: false, error: '이미 배부 완료된 신청은 취소할 수 없습니다.' };
    if (status === 'cancelled')   return { ok: false, error: '이미 취소된 신청입니다.' };

    // 재고 복구
    let items = [];
    try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch(e) {}
    items.forEach(item => restoreStock(item.num, item.qty, { reqId: id, name: rows[i][4], reason: '신청 전체취소' }));

    s.getRange(i + 1, statusCol + 1).setValue('cancelled');
    s.getRange(i + 1, (isOld ? 11 : 13) + 1).setValue(new Date().toLocaleString('ko-KR'));

    return { ok: true };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 항목별 취소 (부분취소) ────────────────────────────────────
function cancelItems({ id, indices }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld    = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol = isOld ? 8 : 10;
    const statusCol= isOld ? 10 : 12;

    const status = String(rows[i][statusCol] || '');
    if (status === 'distributed') return { ok: false, error: '배부 완료된 신청은 항목 취소가 불가합니다.' };
    if (status === 'cancelled')   return { ok: false, error: '이미 전체 취소된 신청입니다.' };

    let items = [];
    try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch(e) {}

    // 선택 항목 취소 처리 및 재고 복구
    indices.forEach(idx => {
      if (items[idx] && !items[idx].cancelled && !items[idx].distributed) {
        items[idx].cancelled = true;
        restoreStock(items[idx].num, items[idx].qty, { reqId: id, name: rows[i][4], reason: '항목 취소(관리자)' });
      }
    });

    // 전체 취소됐으면 cancelled, 나머지가 모두 승인됐으면 approved
    const activeItems  = items.filter(item => !item.cancelled);
    const allCancelled = activeItems.length === 0;
    const allApproved  = !allCancelled && activeItems.every(item => item.approved);
    let newStatus = status;
    if (allCancelled) newStatus = 'cancelled';
    else if (allApproved && status === 'pending') newStatus = 'approved';

    s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items));
    s.getRange(i + 1, statusCol + 1).setValue(newStatus);
    s.getRange(i + 1, (isOld ? 11 : 13) + 1).setValue(new Date().toLocaleString('ko-KR'));

    // 취소 후 나머지가 모두 승인돼서 approved로 전환 시 이메일
    if (allApproved && status === 'pending') {
      const recipientEmail = rows[i][6];
      if (recipientEmail) {
        sendNotificationEmail(recipientEmail, rows[i][4], rows[i][2], rows[i][3], rows[i][7], activeItems, 'approved', '');
      }
    }

    return { ok: true, status: newStatus };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 유저 취소신청 ────────────────────────────────────────────────
function requestCancelItem({ id, idx }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld    = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol = isOld ? 8 : 10;

    let items = [];
    try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch(e) {}

    if (!items[idx]) return { ok: false, error: '항목을 찾을 수 없습니다.' };
    if (items[idx].cancelled) return { ok: false, error: '이미 취소된 항목입니다.' };
    if (items[idx].distributed) return { ok: false, error: '이미 배부된 항목은 취소신청할 수 없습니다.' };

    items[idx].cancelRequested = true;
    s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items));
    s.getRange(i + 1, (isOld ? 11 : 13) + 1).setValue(new Date().toLocaleString('ko-KR'));

    return { ok: true };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 관리자 취소확정 ───────────────────────────────────────────────
function confirmCancelItem({ id, idx }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld    = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol = isOld ? 8  : 10;
    const statusCol= isOld ? 10 : 12;

    let items = [];
    try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch(e) {}

    if (!items[idx]) return { ok: false, error: '항목을 찾을 수 없습니다.' };

    // 재고 복구 후 취소 확정
    restoreStock(items[idx].num, items[idx].qty, { reqId: id, name: rows[i][4], reason: '취소 확정(관리자)' });
    items[idx].cancelled = true;
    delete items[idx].cancelRequested;

    // 전체 취소 여부 확인
    const activeItems = items.filter(item => !item.cancelled);
    const newStatus = activeItems.length === 0 ? 'cancelled' : String(rows[i][statusCol] || '');

    s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items));
    s.getRange(i + 1, statusCol + 1).setValue(newStatus);
    s.getRange(i + 1, (isOld ? 11 : 13) + 1).setValue(new Date().toLocaleString('ko-KR'));

    return { ok: true };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 관리자 취소신청 거절 ──────────────────────────────────────────
function rejectCancelItem({ id, idx }) {
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;

    const isOld    = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const itemsCol = isOld ? 8 : 10;

    let items = [];
    try { items = JSON.parse(rows[i][itemsCol] || '[]'); } catch(e) {}

    if (!items[idx]) return { ok: false, error: '항목을 찾을 수 없습니다.' };

    delete items[idx].cancelRequested;
    s.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items));
    s.getRange(i + 1, (isOld ? 11 : 13) + 1).setValue(new Date().toLocaleString('ko-KR'));

    return { ok: true };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

// ── 결과보고서 구글 드라이브 업로드 ─────────────────────────────
function uploadReport({ id, listFile, listFileName, docFile, docFileName }) {
  const links = [];

  if (listFile && listFileName) {
    const blob = Utilities.newBlob(Utilities.base64Decode(listFile), MimeType.MICROSOFT_EXCEL, listFileName);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    links.push('명단: ' + file.getUrl());
  }

  if (docFile && docFileName) {
    const mime = docFileName.endsWith('.pdf') ? MimeType.PDF : MimeType.MICROSOFT_WORD;
    const blob = Utilities.newBlob(Utilities.base64Decode(docFile), mime, docFileName);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    links.push('보고서: ' + file.getUrl());
  }

  if (links.length === 0) return { ok: false, error: '업로드할 파일이 없습니다.' };

  // 신청 시트의 관리자메모 칸에 링크 저장
  const s = sheet(SHEET_REQ);
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;
    const isOld   = String(rows[i][8]).trim().startsWith('[') || String(rows[i][8]).trim().startsWith('{');
    const noteCol = isOld ? 12 : 14;
    const existing = String(rows[i][noteCol] || '');
    s.getRange(i + 1, noteCol + 1).setValue(existing ? existing + '\n' + links.join('\n') : links.join('\n'));
    return { ok: true, links };
  }
  return { ok: false, error: '신청을 찾을 수 없습니다.' };
}

function testAll() {
  Logger.log('=== 전체 시스템 진단 ===');

  // 1. 신청 시트 확인
  const reqRows = sheet(SHEET_REQ).getDataRange().getValues();
  Logger.log('신청 시트 행 수: ' + reqRows.length + ' (헤더 포함)');
  if (reqRows.length > 1) {
    Logger.log('헤더: ' + JSON.stringify(reqRows[0]));
    for (let i = 1; i < Math.min(4, reqRows.length); i++) {
      Logger.log(`행${i}: ID=${reqRows[i][0]}, 이름=${reqRows[i][4]}, r[8]="${reqRows[i][8]}", r[10]="${String(reqRows[i][10]).substring(0,30)}"`);
    }
  } else {
    Logger.log('❌ 신청 시트가 비어있음!');
  }

  // 2. getRequests 직접 호출
  const result = getRequests({ name: '', role: 'admin' });
  Logger.log('getRequests 결과: ok=' + result.ok + ', 건수=' + result.requests.length);
  if (result.requests.length > 0) {
    result.requests.forEach(r => {
      Logger.log(`  → ${r.name}, status=${r.status}, items=${r.items.length}건`);
    });
  }

  // 3. 재고 시트 확인
  const stkRows = sheet(SHEET_STK).getDataRange().getValues();
  Logger.log('재고 시트 행 수: ' + stkRows.length);
  Logger.log('재고 헤더: ' + JSON.stringify(stkRows[0]));
}

function testEmail() {
  const myEmail = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to:       myEmail,
    subject:  '[테스트] 카카오프렌즈 배부시스템 이메일',
    htmlBody: '<h2>이메일 발송 테스트 성공! ✅</h2><p>이 메일이 왔다면 이메일 발송 기능이 정상 작동합니다.</p>'
  });
  Logger.log('테스트 이메일 발송 완료: ' + myEmail);
}

// ── 피드백 제출 ───────────────────────────────────────────────
function submitFeedback({ name, dept, rating, category, text }) {
  if (!text) return { ok: false, error: '의견을 입력해주세요.' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName('피드백');
  if (!s) {
    s = ss.insertSheet('피드백');
    s.appendRow(['작성일시', '이름', '부서', '별점', '분류', '의견']);
    s.getRange(1, 1, 1, 6).setBackground('#3C1E1E').setFontColor('#FEE500').setFontWeight('bold');
  }
  s.appendRow([new Date().toLocaleString('ko-KR'), name || '', dept || '', rating, category, text]);
  return { ok: true };
}

// ── 피드백 조회 (관리자용) ────────────────────────────────────
function getFeedback() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName('피드백');
  if (!s || s.getLastRow() <= 1) return { ok: true, feedbacks: [] };
  const rows = s.getDataRange().getValues().slice(1).reverse();
  const feedbacks = rows.map(r => ({
    createdAt: r[0], name: r[1], dept: r[2],
    rating: Number(r[3]), category: r[4], text: r[5]
  }));
  return { ok: true, feedbacks };
}

// ── 유틸 ──────────────────────────────────────────────────────
function sheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// 시트가 날짜처럼 보이는 문자열을 Date 객체로 자동 변환해버리는 문제를 방지:
// 항상 'yyyy-MM-dd' 형태의 순수 문자열로 정규화
function formatDateOnly(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  const s = String(v).trim();
  const m = s.match(/^\d{4}-\d{2}-\d{2}/);
  if (m) return m[0];
  const d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  return s;
}

function sha256(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8
  );
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// ── 배부현황 리포트 시트 생성 (피벗: 제품×팀) ───────────────────
// Apps Script 편집기에서 직접 실행: generateReport()
function generateReport() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 기존 시트 삭제 후 재생성
  const existing = ss.getSheetByName('배부현황');
  if (existing) ss.deleteSheet(existing);
  const report = ss.insertSheet('배부현황');

  // 재고 데이터 읽기 (제품번호 → {original, current})
  const stockMap = {};
  sheet(SHEET_STK).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    const num = String(r[0]).padStart(3, '0');
    const original = Number(r[3]) || Number(r[2]) || 0;
    const current  = Number(r[2]) || 0;
    stockMap[num] = { original, current };
  });

  // 신청 데이터 읽기
  const reqRows = sheet(SHEET_REQ).getDataRange().getValues().slice(1);

  // { "제품번호|제품명" : { "본부-팀(이름)" : qty } }
  const pivot = {};
  const colKeySet = new Set();

  reqRows.forEach(r => {
    if (!r[0]) return;
    const isOld    = String(r[8]).trim().startsWith('[') || String(r[8]).trim().startsWith('{');
    const itemsCol = isOld ? 8 : 10;
    const statusCol= isOld ? 10 : 12;
    const status   = String(r[statusCol] || '');
    if (status === 'cancelled' || status === 'rejected') return;

    const dept = String(r[2] || '').trim();
    const team = String(r[3] || '').trim();
    const name = String(r[4] || '').trim();
    const colKey = [dept, team, name].filter(Boolean).join(' / ');

    let items = [];
    try { items = JSON.parse(r[itemsCol] || '[]'); } catch(e) {}

    items.forEach(item => {
      if (item.cancelled) return;
      const rowKey = `${String(item.num).padStart(3,'0')}|${item.name}`;
      if (!pivot[rowKey]) pivot[rowKey] = {};
      pivot[rowKey][colKey] = (pivot[rowKey][colKey] || 0) + Number(item.qty);
      colKeySet.add(colKey);
    });
  });

  // 열 키 정렬
  const colKeys = Array.from(colKeySet).sort();
  const totalCols = colKeys.length;

  // 헤더 행: [제품번호, 제품명, 총재고, 팀1, 팀2, ..., 배부합계, 잔여재고]
  const headerRow = ['제품번호', '제품명', '총재고', ...colKeys, '배부합계', '잔여재고'];
  report.appendRow(headerRow);

  // 헤더 스타일
  const hRange = report.getRange(1, 1, 1, headerRow.length);
  hRange.setBackground('#3C1E1E');
  hRange.setFontColor('#FEE500');
  hRange.setFontWeight('bold');
  hRange.setHorizontalAlignment('center');

  // 총재고 열 헤더 색상 구분
  report.getRange(1, 3).setBackground('#5C3D3D').setFontColor('#FFFFFF');
  // 배부합계/잔여재고 열 헤더 색상 구분
  report.getRange(1, headerRow.length - 1, 1, 2).setBackground('#5C3D3D').setFontColor('#FFFFFF');

  // 재고 시트의 전체 제품을 기준으로 rowKeys 구성 (신청 없는 제품도 포함)
  const allProductKeys = sheet(SHEET_STK).getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => {
      const num = String(r[0]).padStart(3, '0');
      const name = String(r[1] || '');
      return `${num}|${name}`;
    })
    .sort();
  const rowKeys = allProductKeys;
  const dataRows = [];
  const colTotals = new Array(totalCols).fill(0);
  let grandOriginal = 0;
  let grandDistributed = 0;

  rowKeys.forEach(rowKey => {
    const [num, name] = rowKey.split('|');
    const stock = stockMap[num] || { original: 0, current: 0 };
    const original = stock.original;

    const qtys = colKeys.map((ck, ci) => {
      const q = (pivot[rowKey] && pivot[rowKey][ck]) || 0;
      colTotals[ci] += q;
      return q;
    });
    const rowTotal = qtys.reduce((a, b) => a + b, 0);
    const remaining = original - rowTotal;

    grandOriginal += original;
    grandDistributed += rowTotal;

    dataRows.push([num, name, original, ...qtys, rowTotal, remaining]);
  });

  if (dataRows.length > 0) {
    report.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);

    // 데이터 영역 줄무늬
    for (let i = 0; i < dataRows.length; i++) {
      const bg = i % 2 === 0 ? '#FFFFFF' : '#FFF9E6';
      report.getRange(i + 2, 1, 1, headerRow.length).setBackground(bg);
    }

    // 총재고 열 강조 (연한 회색)
    report.getRange(2, 3, dataRows.length, 1).setBackground('#F0F0F0').setFontWeight('bold');

    // 배부합계 열 강조 (노란색)
    const distCol = 3 + totalCols + 1;
    report.getRange(2, distCol, dataRows.length, 1).setFontWeight('bold').setBackground('#FEE500');

    // 잔여재고 열 강조 (연두색)
    const remCol = headerRow.length;
    report.getRange(2, remCol, dataRows.length, 1).setFontWeight('bold').setBackground('#E8F5E9');

    // 잔여재고가 0이면 빨간색 표시
    dataRows.forEach((row, i) => {
      if (row[row.length - 1] <= 0) {
        report.getRange(i + 2, remCol).setBackground('#FFCDD2').setFontColor('#C62828');
      }
    });
  }

  // 합계 행
  const colTotalGrand = colTotals.reduce((a, b) => a + b, 0);
  const totalRow = ['', '합계', grandOriginal, ...colTotals, grandDistributed, grandOriginal - grandDistributed];
  const totalRowIdx = dataRows.length + 2;
  report.appendRow(totalRow);
  report.getRange(totalRowIdx, 1, 1, headerRow.length)
    .setBackground('#3C1E1E').setFontColor('#FEE500').setFontWeight('bold');

  // 열 너비
  report.setColumnWidth(1, 80);   // 제품번호
  report.setColumnWidth(2, 250);  // 제품명
  report.setColumnWidth(3, 80);   // 총재고
  for (let c = 4; c <= headerRow.length; c++) report.setColumnWidth(c, 100);

  // 틀 고정
  report.setFrozenRows(1);
  report.setFrozenColumns(3); // 제품번호+제품명+총재고 고정

  Logger.log(`배부현황 시트 생성 완료: 제품 ${rowKeys.length}종 / 팀 ${totalCols}개`);
  // 스프레드시트 메뉴에서 수동 실행한 경우에만 알림 표시 (웹앱에서 호출 시 UI 컨텍스트가 없어 예외 발생)
  try { SpreadsheetApp.getUi().alert(`배부현황 시트가 생성되었습니다.\n제품 ${rowKeys.length}종 × 팀 ${totalCols}개`); } catch (e) {}
}

// ── 수취예정일 한 달 이내 신청 건 관리자 메일 발송 ─────────────────
// 매주 월요일 자동 발송 트리거는 setupWeeklyPickupReminderTrigger()로 1회 등록
const STATUS_LABEL_KO = {
  pending: '대기중', approved: '승인완료', distributed: '배부완료',
  cancelled: '취소됨', rejected: '반려됨'
};

function sendUpcomingPickupReminder() {
  const rows = sheet(SHEET_REQ).getDataRange().getValues().slice(1);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const cutoff = new Date(today);
  cutoff.setDate(cutoff.getDate() + 30);

  const upcoming = [];

  rows.forEach(r => {
    if (!r[0]) return;

    const r8  = String(r[8]  || '').trim();
    const r10 = String(r[10] || '').trim();

    // 구형(13열) 행은 수취예정일 컬럼이 없어 대상에서 제외
    if (!(r10.startsWith('[') || r10.startsWith('{'))) return;

    const pickupDate = formatDateOnly(r[8]);
    const itemsJson  = r10;
    const status     = String(r[12] || '');
    if (!pickupDate) return;
    if (status === 'cancelled' || status === 'rejected' || status === 'distributed') return;

    const pd = new Date(pickupDate);
    if (isNaN(pd)) return;
    pd.setHours(0, 0, 0, 0);
    if (pd < today || pd > cutoff) return;

    let items = [];
    try { items = JSON.parse(itemsJson || '[]'); } catch (e) {}
    const activeItems = items.filter(i => !i.cancelled);
    if (activeItems.length === 0) return;

    upcoming.push({
      pickupDate, status,
      dept: r[2], team: r[3], name: r[4], contact: r[5],
      items: activeItems
    });
  });

  if (upcoming.length === 0) {
    Logger.log('한 달 이내 수취예정 신청 없음 - 메일 발송 생략');
    return;
  }

  upcoming.sort((a, b) => new Date(a.pickupDate) - new Date(b.pickupDate));

  const rowsHtml = upcoming.map(u => {
    const itemList = u.items.map(i => `${i.name} × ${i.qty}개`).join('<br>');
    return `<tr>
      <td style="padding:8px;border:1px solid #ddd">${u.pickupDate}</td>
      <td style="padding:8px;border:1px solid #ddd">${u.dept || ''} / ${u.team || ''}</td>
      <td style="padding:8px;border:1px solid #ddd">${u.name || ''}</td>
      <td style="padding:8px;border:1px solid #ddd">${u.contact || ''}</td>
      <td style="padding:8px;border:1px solid #ddd">${itemList}</td>
      <td style="padding:8px;border:1px solid #ddd">${STATUS_LABEL_KO[u.status] || u.status}</td>
    </tr>`;
  }).join('');

  const fmt = d => Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  const html = `
    <div style="font-family:sans-serif">
      <h2>📦 수취예정일 한 달 이내 신청 목록</h2>
      <p>오늘(${fmt(today)}) 기준, ${fmt(cutoff)}까지 수취 예정인 카카오프렌즈 GIK 신청 건입니다.</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px">
        <thead><tr style="background:#3C1E1E;color:#FEE500">
          <th style="padding:8px;border:1px solid #ddd">수취예정일</th>
          <th style="padding:8px;border:1px solid #ddd">본부/팀</th>
          <th style="padding:8px;border:1px solid #ddd">신청인</th>
          <th style="padding:8px;border:1px solid #ddd">연락처</th>
          <th style="padding:8px;border:1px solid #ddd">신청 물품/수량</th>
          <th style="padding:8px;border:1px solid #ddd">상태</th>
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <p style="color:#888;font-size:12px;margin-top:16px">매주 월요일 자동 발송되는 메일입니다.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: `[카카오프렌즈 GIK] 수취예정 한 달 이내 신청 ${upcoming.length}건`,
    htmlBody: html
  });
  Logger.log(`수취예정 리마인더 메일 발송 완료: ${upcoming.length}건`);
}

// ── Apps Script 편집기에서 1회만 실행: 매주 월요일 오전 9시 트리거 등록 ──
function setupWeeklyPickupReminderTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['sendUpcomingPickupReminder', 'syncPickupCalendar'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendUpcomingPickupReminder')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // 웹앱 사용이 없는 주에도 캘린더가 최신 상태를 유지하도록 매일 동기화
  ScriptApp.newTrigger('syncPickupCalendar')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  Logger.log('매주 월요일 오전 9시 리마인더 + 매일 오전 6시 캘린더 동기화 트리거 등록 완료');
}

// ── 수취예정일 캘린더 동기화 ──────────────────────────────────
// "카카오프렌즈 GIK 수취일정"이라는 전용 캘린더를 만들어 관리자 구글 캘린더에 표시
const PICKUP_CAL_PROP_KEY = 'PICKUP_CALENDAR_ID';
const PICKUP_CAL_NAME     = '카카오프렌즈 GIK 수취일정';

function getPickupCalendar() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(PICKUP_CAL_PROP_KEY);
  if (savedId) {
    const cal = CalendarApp.getCalendarById(savedId);
    if (cal) return cal;
  }
  const cal = CalendarApp.createCalendar(PICKUP_CAL_NAME);
  props.setProperty(PICKUP_CAL_PROP_KEY, cal.getId());
  Logger.log(`캘린더 신규 생성: ${PICKUP_CAL_NAME} (${cal.getId()})`);
  return cal;
}

// 이 캘린더는 시스템 전용이므로 매번 전체 삭제 후 재생성해 최신 상태를 유지
function syncPickupCalendar() {
  const cal = getPickupCalendar();

  const rangeStart = new Date();
  rangeStart.setDate(rangeStart.getDate() - 14); // 지난 2주까지 포함 (놓친 수취 파악용)
  rangeStart.setHours(0, 0, 0, 0);
  const rangeEnd = new Date();
  rangeEnd.setDate(rangeEnd.getDate() + 180);
  rangeEnd.setHours(23, 59, 59, 999);

  cal.getEvents(rangeStart, rangeEnd).forEach(ev => ev.deleteEvent());

  const rows = sheet(SHEET_REQ).getDataRange().getValues().slice(1);
  let count = 0;

  rows.forEach(r => {
    if (!r[0]) return;
    const r10 = String(r[10] || '').trim();
    if (!(r10.startsWith('[') || r10.startsWith('{'))) return; // 구형 행은 수취예정일 없음

    const pickupDate = formatDateOnly(r[8]);
    const status = String(r[12] || '');
    if (!pickupDate) return;
    if (status === 'cancelled' || status === 'rejected' || status === 'distributed') return;

    const pd = new Date(pickupDate);
    if (isNaN(pd)) return;
    if (pd < rangeStart || pd > rangeEnd) return;

    let items = [];
    try { items = JSON.parse(r10 || '[]'); } catch (e) {}
    const activeItems = items.filter(i => !i.cancelled);
    if (activeItems.length === 0) return;

    const dept = r[2], team = r[3], name = r[4], contact = r[5];
    const itemSummary = activeItems.map(i => `${i.name}×${i.qty}`).join(', ');
    const title = `📦 ${name} (${dept}/${team}) 수취예정`;
    const desc = [
      `신청인: ${name}`,
      `연락처: ${contact}`,
      `본부/팀: ${dept} / ${team}`,
      `상태: ${STATUS_LABEL_KO[status] || status}`,
      `물품: ${itemSummary}`
    ].join('\n');

    cal.createAllDayEvent(title, pd, { description: desc });
    count++;
  });

  Logger.log(`수취일정 캘린더 동기화 완료: ${count}건`);
}
