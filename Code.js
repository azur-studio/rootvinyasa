/* ================================================================
   [Code.gs] - V5.12 (최종 통합 완성본)
   ─────────────────────────────────────────
   ✅ 변경/수정 내역
   1. 선택 휴강 보상(+1) 완전 제거
      → "오면 1회 차감, 안 오면 자동 연장"만 적용
   2. addHoliday / removeHoliday 복구
      → 관리자 화면에서 휴강 추가·삭제 가능
   3. "특정" 회원 정책 수정 (VIP 표현 제거)
      → 5월까지 : 5회 = 1개월 (쉬어가기 = totalWeeks ÷ 5)
      → 6월부터 : 4회 = 1개월 (쉬어가기 = totalWeeks ÷ 4)
   4. specialVipRuleActive 방향 오류 수정
      → V5.8에서 >=로 뒤집혔던 것을 <로 원복
   5. 메모 저장 형식 유지
      → [2026-03-28 관리자]: 메모내용
   ================================================================ */


// ──────────────────────────────────────────────
// 1. 전역 상수
// ──────────────────────────────────────────────
// SS_ID, ADMIN_PIN 은 GAS 스크립트 속성에서 읽습니다.
// 설정: GAS 편집기 → 프로젝트 설정 → 스크립트 속성 탭
//   속성 이름  SS_ID    → 값: 스프레드시트 ID
//   속성 이름  ADMIN_PIN → 값: 관리자 PIN
var _scriptProps_ = PropertiesService.getScriptProperties();
var SS_ID        = _scriptProps_.getProperty('SS_ID');
var SHEET_DB_NEW = '통합 수강 DB (신규)';
var SHEET_LOG    = '결제 및 신청 로그 (신규)';
var SHEET_HOLIDAY= '휴강 캘린더 (신규)';
var ADMIN_PIN    = _scriptProps_.getProperty('ADMIN_PIN');

var REGEX_DATE_SEP = new RegExp('[./]', 'g');
var REGEX_DATE_8   = new RegExp('^\\d{8}$');
var REGEX_NON_NUM  = new RegExp('[^0-9]', 'g');


// ──────────────────────────────────────────────
// 2. 날짜 유틸 함수
// ──────────────────────────────────────────────
function parseSafeDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  var s = String(val).trim().replace(/\s+/g, '').replace(/년/g, '-').replace(/월/g, '-').replace(/일/g, '');
  s = s.replace(REGEX_DATE_SEP, '-');
  if (REGEX_DATE_8.test(s)) return new Date(parseInt(s.substring(0,4)), parseInt(s.substring(4,6))-1, parseInt(s.substring(6,8)));
  var parts = s.split('-');
  if (parts.length === 3) {
    var y = parseInt(parts[0], 10);
    if (y < 100) y += 2000;
    return new Date(y, parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  var fallback = new Date(s);
  return isNaN(fallback.getTime()) ? null : fallback;
}

function formatDateSafe(d) {
  var dt = parseSafeDate(d);
  return (!dt || isNaN(dt.getTime())) ? '' : Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getSafeString(val) {
  return (val instanceof Date) ? formatDateSafe(val) : String(val || '');
}

function attachJosa(word) {
  if (!word) return '';
  var lastChar = word.charCodeAt(word.length - 1);
  if (lastChar < 0xAC00 || lastChar > 0xD7A3) return word + '(으)로 인해';
  var jong = (lastChar - 0xAC00) % 28;
  return (jong > 0 && jong !== 8) ? word + '으로 인해' : word + '로 인해';
}


// ──────────────────────────────────────────────
// 3. 핵심 날짜 계산 함수
// ──────────────────────────────────────────────
function recalcDates(startDateStr, totalWeeks, excludeList) {
  var startObj = parseSafeDate(startDateStr);
  if (!startObj || isNaN(startObj.getTime())) return [];
  var excl = excludeList || [];
  var dates = [];
  var curr = new Date(startObj.getTime());
  var maxIter = totalWeeks + excl.length + 104;
  while (dates.length < totalWeeks && maxIter-- > 0) {
    var dStr = formatDateSafe(curr);
    if (excl.indexOf(dStr) === -1) dates.push(dStr);
    curr.setDate(curr.getDate() + 7);
  }
  return dates;
}

function getOriginalStartDate(datesG, datesH) {
  var all = datesG.concat(datesH).filter(Boolean);
  if (all.length === 0) return null;
  all.sort(function(a, b){ return parseSafeDate(a) - parseSafeDate(b); });
  return all[0];
}


// ──────────────────────────────────────────────
// 4. "특정" 회원 쉬어가기 한도 계산 헬퍼
//    6월부터 4회 = 1개월 기준 통일
// ──────────────────────────────────────────────
function calcMaxPauses(type, totalWeeks) {
  return Math.floor(totalWeeks / 4);
}


// ──────────────────────────────────────────────
// 5. 휴강 데이터 로더
// ──────────────────────────────────────────────
function loadHolidays(ss) {
  var result = { all: [], mandatory: [], optional: [], reasons: {}, types: {} };
  var hSheet = ss.getSheetByName(SHEET_HOLIDAY);
  if (!hSheet) return result;
  var hData = hSheet.getDataRange().getValues();
  for (var i = 1; i < hData.length; i++) {
    if (!hData[i][0]) continue;
    var dStr = formatDateSafe(hData[i][0]);
    if (!dStr) continue;
    var reason = hData[i][1] ? String(hData[i][1]) : '';
    var type   = hData[i][2] ? String(hData[i][2]).trim() : '일반';
    if (type !== '선택') type = '일반';
    result.all.push(dStr);
    result.types[dStr] = type;
    if (reason) result.reasons[dStr] = attachJosa(reason);
    if (type === '선택') result.optional.push(dStr);
    else result.mandatory.push(dStr);
  }
  return result;
}


// ──────────────────────────────────────────────
// 6. 선택 휴강 참석 기록 읽기 (메모장에서 추출)
//    ✅ [V5.11] 보상 없이 참석 여부만 기록
// ──────────────────────────────────────────────
function getAttendedOptionals(memo) {
  var dates = [];
  if (!memo) return dates;
  var lines = memo.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var match = lines[i].match(/\[자율참석: (\d{4}-\d{2}-\d{2})\]/);
    if (match) dates.push(match[1]);
  }
  return dates;
}


// ──────────────────────────────────────────────
// 7. 전체 DB 재계산
//    ✅ [V5.11] 자율참석 날짜는 휴강에서 제외
// ──────────────────────────────────────────────
function rebuildAllMemberDates() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
  if (!dbSheet) return;
  var holidays = loadHolidays(ss);
  var data = dbSheet.getDataRange().getValues();
  var newDatesMap = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    var datesG = getSafeString(data[i][6]).split(',').map(function(s){return s.trim();}).filter(String);
    var datesH = getSafeString(data[i][7]).split(',').map(function(s){return s.trim();}).filter(String);
    var totalWeeks = Number(data[i][5]);
    var originalStart = getOriginalStartDate(datesG, datesH);
    if (!originalStart || totalWeeks <= 0) continue;
    var attendedOpts = getAttendedOptionals(String(data[i][8] || ''));
    // 자율참석으로 확정된 날짜는 제외 목록에서 빠짐 (수강일로 포함)
    var excludeList = holidays.all.filter(function(d) {
      return d !== originalStart && attendedOpts.indexOf(d) === -1;
    }).concat(datesH);
    newDatesMap[i] = recalcDates(originalStart, totalWeeks, excludeList);
  }
    /* @LOGIC_BATCH_IO_OPTIMIZATION */
  var updateData = [];
  
  for (var i = 1; i < data.length; i++) {
    if (newDatesMap[i]) {
      // 계산된 새 날짜 삽입
      updateData.push(["'" + newDatesMap[i].join(', ')]); 
    } else {
      // 변경 없는 기존 날짜 데이터(G열, 인덱스 6) 유지
      updateData.push([data[i][6]]); 
    }
  }
  
  // 단 1번의 API 통신으로 시트 전체 일괄 업데이트
  if (updateData.length > 0) {
    dbSheet.getRange(2, 7, updateData.length, 1).setValues(updateData);
  }

}


// ──────────────────────────────────────────────
// 8. 트리거 함수
// ──────────────────────────────────────────────
function handleHolidayChange(e) {
  if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() === SHEET_HOLIDAY) {
    rebuildAllMemberDates();
  }
}

function handleSpreadsheetEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();

  // 설정 시트에서 시간 관련 항목 변경 시 트리거 자동 재설정
  if (sheet.getName() === SHEET_SETTINGS) {
    var changedKey = String(sheet.getRange(e.range.getRow(), 1).getValue()).trim();
    if (changedKey === '문자 발송 시간' || changedKey === '문자 생성 시간' || changedKey === '브리핑 시간') {
      setupAllTriggers();
    }
    return;
  }

  // 결제 확인 체크박스 처리
  if (sheet.getName() === SHEET_LOG && e.range.getColumn() === 8 && e.range.getRow() > 1) {
    var row = e.range.getRow();
    if (e.range.getValue() === true) {
      if (confirmPaymentAdmin(row)) {
        var logSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_LOG);
        var rowData = logSheet.getRange(row, 1, 1, 9).getValues()[0];
        sendPaymentConfirmedSMS(
          String(rowData[1]),
          String(rowData[2]).replace(/[^0-9]/g, ''),
          String(rowData[3]),
          String(rowData[4]),
          String(rowData[6])
        );
        sheet.getRange(row, 1, 1, 9).setBackground('#c8e6c9');
      }
    } else {
      cancelPaymentAdmin(row);
      sheet.getRange(row, 1, 1, 9).setBackground('#fff9c4');
    }
  }
}



// ──────────────────────────────────────────────
// 9. 휴강 추가·삭제 (관리자 화면용)
//    addHoliday: 섹션 42에 정의 (SMS 자동화 포함 완전판)
//    removeHoliday: 아래 정의
// ──────────────────────────────────────────────

function removeHoliday(dateStr) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var hSheet = ss.getSheetByName(SHEET_HOLIDAY);
    if (!hSheet) return false;
    var data = hSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (formatDateSafe(data[i][0]) === dateStr) {
        hSheet.deleteRow(i + 1);
        rebuildAllMemberDates();
        return true;
      }
    }
    return false;
  } catch(ex) { return false; }
}


/* @LOGIC_CONCURRENCY_LOCK */
/* @LOGIC_CONCURRENCY_LOCK & BACKEND_LOOKUP */
/* @LOGIC_CONCURRENCY_LOCK & BACKEND_LOOKUP (최종 교정본) */
function submitApplication(payload) {
  var lock = LockService.getScriptLock();

  /* ── 입력 검증: 악의적 페이로드/오류 입력 차단 ── */
  try {
    if (!payload || typeof payload !== 'object') return false;
    var nm = String(payload.name || '').trim();
    if (!/^[가-힣A-Za-z\s]{2,10}$/.test(nm)) return false; // 이름: 2~10자 한글/영문
    payload.name = nm;
    var fp = String(payload.fullPhone || '').replace(/[^0-9]/g, '');
    if (payload.isNew) {
      if (!/^010\d{8}$/.test(fp)) return false; // 신규: 010-XXXX-XXXX 11자리
    }
    payload.fullPhone = fp;
    var amt = Number(payload.amount);
    if (!isFinite(amt) || amt < 0 || amt > 5000000) return false; // 금액 한도
    payload.amount = amt;
    var allowedOptions = ['정규','연장','특정','특별','원데이','신규','없음'];
    if (payload.option && allowedOptions.indexOf(String(payload.option)) === -1 && String(payload.option).length > 30) return false;
    if (payload.memo && String(payload.memo).length > 500) return false; // 메모 길이 제한
    if (payload.dates && Array.isArray(payload.dates)) {
      if (payload.dates.length > 100) return false; // 일정 폭주 방지
      for (var di = 0; di < payload.dates.length; di++) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(String(payload.dates[di]))) return false;
      }
    }
  } catch (vErr) {
    return false;
  }

  try {
    lock.waitLock(5000);
    // 당일 신청 마감 시간 체크
var startDateCheck = parseSafeDate(payload.startDate);
if (startDateCheck) {
  var todayMidnight = new Date(); todayMidnight.setHours(0, 0, 0, 0);
  var startMidnight = new Date(startDateCheck.getTime()); startMidnight.setHours(0, 0, 0, 0);
  if (startMidnight.getTime() === todayMidnight.getTime()) {
    var deadline = String(getSetting('당일 신청 마감', '12:40'));
    var dParts = deadline.split(':');
    var dHour = Number(dParts[0]); var dMin = Number(dParts[1] || 0);
    var now = new Date();
    if (now.getHours() > dHour || (now.getHours() === dHour && now.getMinutes() >= dMin)) {
      lock.releaseLock();
      return false; // 마감 시간 초과
    }
  }
}
    var ss = SpreadsheetApp.openById(SS_ID);
    var logSheet = ss.getSheetByName(SHEET_LOG);
    var finalPhoneToSave = "";
    
    if (payload.isNew) {
      // 신규 회원은 전체 번호 저장
      finalPhoneToSave = payload.fullPhone;
    } else {
      // 기존 회원은 역추적 시작
      var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
      if (dbSheet) {
        var dbData = dbSheet.getDataRange().getValues();
        for (var i = 1; i < dbData.length; i++) {
          var rowName = String(dbData[i][1]).trim();
          var rowPhone = String(dbData[i][2]).replace(/[^0-9]/g, '');
          
          if (rowName === payload.name) {
            var dbLast4 = rowPhone.slice(-4);
            var maskStr = String(payload.phoneMask).trim();
            var isMatch = false;
            
            // 핵심 교정: indexOf를 통한 '패턴 시작' 검증 (12** 의 12와 1260 매칭)
            if (maskStr.indexOf('*') !== -1) {
              var digitsOnly = maskStr.replace(/[^0-9]/g, ''); // "12**" -> "12"
              if (dbLast4.indexOf(digitsOnly) === 0) {
                isMatch = true;
              }
            } else {
              if (dbLast4 === maskStr) {
                isMatch = true;
              }
            }
            
            // 매칭 성공 시 원본 번호 인양
            if (isMatch) {
              finalPhoneToSave = rowPhone;
              break;
            }
          }
        }
      }
      
      // 혹시라도 못 찾았을 경우의 최후 보루
      if (!finalPhoneToSave) {
        finalPhoneToSave = payload.phoneMask;
      }
    }
    
    // 로그 시트에 기록
    logSheet.insertRowAfter(1);
    var targetRange = logSheet.getRange(2, 1, 1, 9);
    targetRange.setValues([[
      new Date(), 
      payload.name, 
      finalPhoneToSave ? "'" + finalPhoneToSave : '', // 찾아낸 완벽한 진짜 번호를 기록
      payload.option, 
      payload.startDate + ' ~ ' + payload.endDate, 
      payload.amount, 
      payload.memo || '', 
      false, 
      ''
    ]]);
    
    logSheet.getRange(2, 1).setNumberFormat('yyyy. MM. dd HH:mm');
    logSheet.getRange(2, 6).setNumberFormat('#,##0');
    // 신규 신청 = 미확정(FALSE) → 연노랑
logSheet.getRange(2, 1, 1, 9).setBackground('#fff9c4');

    return true;
    
  } catch(ex) {
    Logger.log(ex);
    return false; 
  } finally {
    lock.releaseLock();
  }
}





// ──────────────────────────────────────────────
// 11. 결제 확인 (관리자)
//    ✅ [수정] 특정 회원 5회 옵션 인식 추가
// ──────────────────────────────────────────────
function confirmPaymentAdmin(rowIdx) {
  try {
    var ss       = SpreadsheetApp.openById(SS_ID);
    var logSheet = ss.getSheetByName(SHEET_LOG);
    var dbSheet  = ss.getSheetByName(SHEET_DB_NEW);
    var row = logSheet.getRange(rowIdx, 1, 1, 9).getValues()[0];
    if (String(row[8] || '').trim()) return false;

    var name = String(row[1]);
    var phone = String(row[2]).replace(REGEX_NON_NUM, '');
    var option = String(row[3]);
    var amount = Number(row[5]) || 0;
    var startDateStr = String(row[4]).split('~')[0].trim();

    // 특정 회원 다회차(8회·12회) 및 구권(5회) 포함 주수 계산
    var weeks = 4;
    if (option.indexOf('12주') >= 0 || option.indexOf('12회') >= 0) weeks = 12;
    else if (option.indexOf('8주') >= 0 || option.indexOf('8회') >= 0) weeks = 8;
    else if (option.indexOf('5회') >= 0 || option.indexOf('5주') >= 0) weeks = 5;
    else if (option.indexOf('원데이') >= 0) weeks = 1;

    var classify = option.indexOf('연장') >= 0 ? '연장' : (option.indexOf('원데이') >= 0 ? '원데이' : '정규');
    // 특정 회원이거나 옵션에 '특별'/'특정' 포함 시 classify = '특정'
    if (option.indexOf('특별') >= 0 || option.indexOf('특정') >= 0) classify = '특정';

    // ── 원데이 할인 처리 ────────────────────────────────────────────
    // option에 '원데이 할인' 포함 시: 로그에서 14일 이내 확정 원데이 확인 후 차감
    // weeks(쉬어가기 기준) = 원본 그대로 / weeksForDates(실제 날짜 수) = weeks - 1
    var weeksForDates = weeks;
    if (option.indexOf('원데이 할인') >= 0 && classify !== '원데이') {
      var odDisc = _findRecentOnedayLog(ss, name, phone);
      if (odDisc.found && odDisc.perPersonPrice > 0) {
        amount = Math.max(0, amount - odDisc.perPersonPrice);
        weeksForDates = Math.max(1, weeks - 1);
        Logger.log('[confirmPaymentAdmin] 원데이 할인 적용: -' + odDisc.perPersonPrice + '원, dates=' + weeksForDates + '회');
        // 다인 원데이일 경우 강사에게 알림 (동반인 수동 처리 안내)
        if (odDisc.isMultiPerson) {
          try {
            var instrPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
            if (instrPhone) {
              enqueueSMS(instrPhone, '최승훈',
                '[ROOT] 원데이 할인 확인 필요\n' + name + '님이 2인 원데이 결제 후 정규 패스 신청했습니다.\n동반인의 원데이 할인은 별도 수동 처리가 필요합니다.',
                '시스템알림', SMS_STATUS.IMMEDIATE);
            }
          } catch(odAlertEx) { Logger.log('[confirmPaymentAdmin] 다인알림 실패: ' + odAlertEx); }
        }
      }
    }
    // ────────────────────────────────────────────────────────────────
    // 기존에 특정 회원이었던 경우 유지
    var dbData = dbSheet.getDataRange().getValues();
    for (var k = 1; k < dbData.length; k++) {
      var rPhone = String(dbData[k][2]).replace(REGEX_NON_NUM, '');
      if (String(dbData[k][1]) === name && rPhone === phone && String(dbData[k][3]) === '특정') {
        classify = '특정';
        break;
      }
    }

    var holidayData = loadHolidays(ss);
    var excludeList = holidayData.all.filter(function(d) { return d !== startDateStr; });
    var datesArray = recalcDates(startDateStr, weeksForDates, excludeList); // 원데이 할인 시 1회 적게
    var datesString = datesArray.join(', ');

    var passKey = name + phone.slice(-4) + '_' + startDateStr.split('-').join('').slice(4) + '_' + new Date().getTime().toString().slice(-4);

// 교체 후 (2줄)
dbSheet.insertRowAfter(1);
var targetDBRange = dbSheet.getRange(2, 1, 1, 9);
    targetDBRange.setValues([[passKey, name, "'" + phone, classify, amount, weeks, "'" + datesString, '', row[6] || '']]);
    targetDBRange.setFontColor(null);

    logSheet.getRange(rowIdx, 8).setValue(true);
    logSheet.getRange(rowIdx, 9).setValue(passKey);
    return true;
  } catch(ex) { return false; }
}


// ──────────────────────────────────────────────
// 11-W. 관리자 결제확인 + SMS 래퍼 (Admin.html 전용)
//   confirmPaymentAdmin 성공 시 SMS까지 발송
// ──────────────────────────────────────────────
function confirmPaymentAdminWithSMS(rowIdx) {
  try {
    var confirmed = confirmPaymentAdmin(rowIdx);
    if (!confirmed) return false;
    var ss = SpreadsheetApp.openById(SS_ID);
    var row = ss.getSheetByName(SHEET_LOG).getRange(rowIdx, 1, 1, 9).getValues()[0];
    sendPaymentConfirmedSMS(
      String(row[1]),
      String(row[2]).replace(/[^0-9]/g, ''),
      String(row[3]),
      String(row[4]),
      String(row[6])
    );
    return true;
  } catch(ex) {
    Logger.log('[confirmPaymentAdminWithSMS] ' + ex);
    return false;
  }
}


// ──────────────────────────────────────────────
// 11-A. 원데이 할인 조회 헬퍼
//   최근 14일 이내 확정된 원데이 결제 기록 반환
//   1인당 가격(logAmount ÷ 인원)을 반환해 35,000원 하드코딩 방지
// ──────────────────────────────────────────────
function _findRecentOnedayLog(ss, name, phone) {
  try {
    var logSheet = ss.getSheetByName(SHEET_LOG);
    if (!logSheet) return { found: false };
    var logData = logSheet.getDataRange().getValues();
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var cleanPhone = String(phone).replace(/[^0-9]/g, '');

    for (var i = 1; i < logData.length; i++) {
      var row = logData[i];
      if (String(row[1]).trim() !== name) continue;
      if (String(row[2]).replace(/[^0-9]/g, '') !== cleanPhone) continue;
      if (row[7] !== true) continue;                              // 결제확정 여부
      var opt = String(row[3]);
      if (opt.indexOf('원데이') === -1) continue;

      var classDateStr = String(row[4] || '').split('~')[0].trim();
      var classDate = parseSafeDate(classDateStr);
      if (!classDate) continue;
      classDate.setHours(0, 0, 0, 0);
      if (classDate > today) continue;                            // 미래 수업 제외
      var diffDays = Math.floor((today - classDate) / 86400000);
      if (diffDays > 14) continue;                               // 14일 초과 제외

      var persons = 1;
      var pm = opt.match(/(\d+)명/);
      if (pm) persons = parseInt(pm[1]);
      var logAmount = Number(row[5]) || 0;
      var perPersonPrice = persons > 0 ? Math.round(logAmount / persons) : logAmount;
      if (perPersonPrice <= 0) continue;

      return {
        found: true,
        perPersonPrice: perPersonPrice,
        classDate: classDateStr,
        isMultiPerson: persons >= 2
      };
    }
  } catch(e) { Logger.log('[_findRecentOnedayLog] ' + e); }
  return { found: false };
}


// ──────────────────────────────────────────────
// 12. 결제 취소 (관리자)
// ──────────────────────────────────────────────
function cancelPaymentAdmin(rowIdx) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var logSheet = ss.getSheetByName(SHEET_LOG);
    var dbSheet  = ss.getSheetByName(SHEET_DB_NEW);
    var passKey = String(logSheet.getRange(rowIdx, 9).getValue()).trim();
    if (!passKey) return false;

    var dbData = dbSheet.getDataRange().getValues();
    for (var i = dbData.length - 1; i >= 1; i--) {
      if (String(dbData[i][0]) === passKey) { dbSheet.deleteRow(i + 1); break; }
    }
    logSheet.getRange(rowIdx, 8).setValue(false);
    logSheet.getRange(rowIdx, 9).clearContent();
    return true;
  } catch(ex) { return false; }
}


// ──────────────────────────────────────────────
// 13. 스프레드시트 UI 자동 세팅 (1회성 실행)
// ──────────────────────────────────────────────
function setupSpreadsheetUI() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheets = ss.getSheets();
  sheets.forEach(function(s) {
    s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).setFontColor(null);
    s.setFrozenRows(1);
  });

  var sheetHoliday = ss.getSheetByName(SHEET_HOLIDAY);
  if (sheetHoliday) {
    sheetHoliday.getRange('A2:C').setBackground('#FFFDE7');
    var dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
    sheetHoliday.getRange('A2:A').setDataValidation(dateRule);
    sheetHoliday.getRange('A2:A').setNumberFormat('yyyy-MM-dd');
    var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['일반', '선택'], true).setAllowInvalid(false).build();
    sheetHoliday.getRange('C2:C').setDataValidation(typeRule);
  }

  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (sheetLog) {
    sheetLog.getRange('A2:G').setBackground('#F9F9F9');
    sheetLog.getRange('H2:I').setBackground('#FFFDE7');
    sheetLog.getRange('A2:A').setNumberFormat('yyyy-MM-dd HH:mm');
    sheetLog.getRange('F2:F').setNumberFormat('#,##0');
    sheetLog.getRange('H2:H').insertCheckboxes();
  }

  var sheetDB = ss.getSheetByName(SHEET_DB_NEW);
  if (sheetDB) {
    sheetDB.getRange('A2:A').setBackground('#F9F9F9');
    sheetDB.getRange('B2:F').setBackground('#FFFDE7');
    sheetDB.getRange('G2:H').setBackground('#F9F9F9').setNumberFormat('@');
    sheetDB.getRange('I2:I').setBackground('#FFFDE7');
    sheetDB.getRange('E2:E').setNumberFormat('#,##0');
  }
}


// ──────────────────────────────────────────────
// 14. 기본 라우터 및 유틸
// ──────────────────────────────────────────────
function doGet(e) {
  // ── API 모드: fn 파라미터가 있으면 JSON 응답 (GitHub Pages → GAS 직접 호출) ──
  if (e.parameter && e.parameter.fn) {
    return _handleApiGet_(e);
  }
  // ── HTML 모드: 기존 방식 (GAS URL 직접 접속 시) ──
  var template = e.parameter.page === 'admin' ? 'Admin' : 'index';
  return HtmlService.createTemplateFromFile(template).evaluate()
    .setTitle('루트빈야사')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* GitHub Pages에서 fetch()로 호출하는 API 라우터 */
function _handleApiGet_(e) {
  try {
    var fn = e.parameter.fn;
    var dataParam = e.parameter.data;
    var args = dataParam ? JSON.parse(dataParam) : [];
    var result;

    if      (fn === 'getInitialData')        { result = getInitialData(); }
    else if (fn === 'verifyCollisionMember') { result = verifyCollisionMember(args[0], args[1]); }
    else if (fn === 'submitApplication')     { result = submitApplication(args[0]); }
    else { throw new Error('알 수 없는 함수: ' + fn); }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (ex) {
    Logger.log('[_handleApiGet_] ' + ex);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(ex) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/* ──────────────────────────────────────────────
 * 관리자 인증: PIN → 서버 서명 토큰 (HMAC-SHA256, TTL 8h)
 * 클라이언트 localStorage에 토큰만 저장. 위변조 시 거부, 만료 시 재로그인.
 * ─────────────────────────────────────────────── */
var ADMIN_TOKEN_TTL_MS = 8 * 60 * 60 * 1000; // 8시간

function _getAdminTokenSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('ADMIN_TOKEN_SECRET');
  if (!s) {
    s = Utilities.getUuid() + '-' + Utilities.getUuid();
    props.setProperty('ADMIN_TOKEN_SECRET', s);
  }
  return s;
}
function _signAdminToken_(payload) {
  var secret = _getAdminTokenSecret_();
  var raw = Utilities.computeHmacSha256Signature(payload, secret);
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '');
}
function _issueAdminToken_() {
  var exp = Date.now() + ADMIN_TOKEN_TTL_MS;
  var payload = 'v1.' + exp;
  var sig = _signAdminToken_(payload);
  return payload + '.' + sig;
}
function _verifyAdminToken_(token) {
  if (!token || typeof token !== 'string') return false;
  var parts = token.split('.');
  if (parts.length !== 3 || parts[0] !== 'v1') return false;
  var exp = parseInt(parts[1], 10);
  if (!isFinite(exp) || exp < Date.now()) return false;
  var expected = _signAdminToken_('v1.' + exp);
  /* 상수시간 비교 */
  if (expected.length !== parts[2].length) return false;
  var diff = 0;
  for (var i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ parts[2].charCodeAt(i);
  return diff === 0;
}

/* PIN 검증 → 일치 시 서명 토큰 발급. 잘못된 PIN은 비-truthy 반환 */
function verifyAdminPin(inputPin) {
  if (String(inputPin) !== ADMIN_PIN) return false;
  return { token: _issueAdminToken_(), exp: Date.now() + ADMIN_TOKEN_TTL_MS };
}
/* 클라이언트가 저장한 토큰 검증 (페이지 로드 시 호출) */
function verifyAdminToken(token) {
  return _verifyAdminToken_(token);
}


// ──────────────────────────────────────────────
// 15. 미결 결제 목록 조회
// ──────────────────────────────────────────────
// ──────────────────────────────────────────────
// 15. 미결 결제 목록 조회
//     ✅ 원데이 수업일 지난 건 자동 제외 + 최신순
// ──────────────────────────────────────────────
function getPendingPayments() {
  var logSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_LOG);
  var data = logSheet.getDataRange().getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var res = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === true || !data[i][1]) continue;  // 확정됐거나 빈 행 제외

    var option   = String(data[i][3]);
    var schedule = String(data[i][4]);

    // 원데이: 수업일이 오늘보다 이전이면 매칭 대상에서 제외
    if (option.indexOf('원데이') !== -1) {
      var classDateStr = schedule.split('~')[0].trim();
      var classDate = parseSafeDate(classDateStr);
      if (classDate && classDate < today) continue;
    }

    res.push({
      rowIdx:   i + 1,
      reqDate:  data[i][0],           // 신청 시각 (Date 객체)
      name:     String(data[i][1]),
      phone:    String(data[i][2]),
      option:   option,
      schedule: schedule,
      amount:   String(data[i][5]),
      memo:     String(data[i][6])
    });
  }

  // 최신 신청 순 정렬 (신청 시각 내림차순)
  res.sort(function(a, b) {
    var da = a.reqDate instanceof Date ? a.reqDate : new Date(a.reqDate);
    var db = b.reqDate instanceof Date ? b.reqDate : new Date(b.reqDate);
    return db - da;
  });

  // reqDate를 문자열로 변환 (기존 호환)
  res.forEach(function(r) {
    r.reqDate = formatDateSafe(r.reqDate);
  });

  return res;
}


// ──────────────────────────────────────────────
// 16. 관리자 대시보드 데이터
//    ✅ [V5.11] 선택 휴강일 잠재 참석자 목록 포함
//    ✅ [수정] maxPauses에 특정 회원 정책 반영
// ──────────────────────────────────────────────
function getAdminDashboardData(weekOffset) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var targetSat = new Date(today);
  targetSat.setDate(targetSat.getDate() + ((6 - today.getDay() + 7) % 7));
  var holidayData = loadHolidays(ss);
  var targetSatStr = formatDateSafe(targetSat);

  // 일반 휴강이면 다음 주로 넘김 (기준 주 계산)
  while (holidayData.mandatory.indexOf(targetSatStr) !== -1) {
    targetSat.setDate(targetSat.getDate() + 7);
    targetSatStr = formatDateSafe(targetSat);
  }

  // 주 네비게이션 오프셋 적용
  if (weekOffset) {
    targetSat.setDate(targetSat.getDate() + weekOffset * 7);
    targetSatStr = formatDateSafe(targetSat);
  }

  var isHol = holidayData.all.indexOf(targetSatStr) !== -1;
  var res = {
    nextSaturday: targetSatStr,
    isHoliday: isHol,
    holidayReason: isHol ? (holidayData.reasons[targetSatStr] || '') : '',
    holidayType: isHol ? (holidayData.types[targetSatStr] || '') : '',
    attendees: [],
    paused: [],
    potentialAttendees: []  // 선택 휴강일 때 참석 가능 회원
  };

  var dbData = ss.getSheetByName(SHEET_DB_NEW).getDataRange().getValues();
  for (var i = 1; i < dbData.length; i++) {
    if (!dbData[i][0]) continue;
    var datesG = getSafeString(dbData[i][6]).split(',').map(function(s){return s.trim();}).filter(String);
    var datesH = getSafeString(dbData[i][7]).split(',').map(function(s){return s.trim();}).filter(String);
    var totalWeeks = Number(dbData[i][5]);
    var type = String(dbData[i][3]);
    var obj = {
      passKey: dbData[i][0],
      name: dbData[i][1],
      phone: String(dbData[i][2]).replace(REGEX_NON_NUM, ''),
      type: type,
      endDate: datesG[datesG.length - 1] || '',
      usedPauses: datesH.length,
      maxPauses: calcMaxPauses(type, totalWeeks)
    };
    if (datesG.indexOf(targetSatStr) !== -1) {
      res.attendees.push(obj);
    } else if (datesH.indexOf(targetSatStr) !== -1) {
      res.paused.push(obj);
    } else if (res.holidayType === '선택') {
      // 선택 휴강일 때: 수강 기간 중인 회원은 잠재 참석자로 분류
      if (datesG.length > 0 && parseSafeDate(datesG[0]) <= targetSat && parseSafeDate(datesG[datesG.length - 1]) >= targetSat) {
        res.potentialAttendees.push(obj);
      }
    }
  }
  return res;
}


// ──────────────────────────────────────────────
// 17. 쉬어가기 토글 (관리자)
//    ✅ [수정] maxPauses에 특정 회원 정책 반영
// ──────────────────────────────────────────────
function togglePauseAdmin(passKey, saturdayDate, isAdding) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    var holidayData = loadHolidays(ss);
    var data = dbSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(passKey)) continue;
      var datesG = getSafeString(data[i][6]).split(',').map(function(s){return s.trim();}).filter(String);
      var datesH = getSafeString(data[i][7]).split(',').map(function(s){return s.trim();}).filter(String);
      var totalWeeks = Number(data[i][5]);
      var type = String(data[i][3]);
      var originalStart = getOriginalStartDate(datesG, datesH);
      if (!originalStart) return { success: false, reason: '시작일을 확인할 수 없습니다.' };

      var maxPauses = calcMaxPauses(type, totalWeeks);

      if (isAdding) {
        if (datesH.indexOf(saturdayDate) !== -1) return { success: false, reason: '이미 쉬어가기 처리된 날짜입니다.' };
        if (datesH.length >= maxPauses) return { success: false, reason: '쉬어가기 가능 횟수를 초과했습니다.' };
        datesH.push(saturdayDate);
      } else {
        var idxH = datesH.indexOf(saturdayDate);
        if (idxH === -1) return { success: false, reason: '쉬어가기 목록에 없는 날짜입니다.' };
        datesH.splice(idxH, 1);
      }

      var attendedOpts = getAttendedOptionals(String(data[i][8] || ''));
      var excludeList = holidayData.all.filter(function(d) {
        return d !== originalStart && attendedOpts.indexOf(d) === -1;
      }).concat(datesH);
      var newDatesG = recalcDates(originalStart, totalWeeks, excludeList);

      dbSheet.getRange(i + 1, 7).setValue("'" + newDatesG.join(', '));
      dbSheet.getRange(i + 1, 8).setValue(datesH.length > 0 ? "'" + datesH.join(', ') : '');
      return { success: true };
    }
    return { success: false, reason: '패스를 찾을 수 없습니다.' };
  } catch(ex) { return { success: false, reason: ex.toString() }; }
}


// ──────────────────────────────────────────────
// 18. 선택 휴강 참석 확정/취소 (관리자)
//    ✅ [V5.11] 보상(+1) 없음. 오직 참석 여부만 기록
// ──────────────────────────────────────────────
function toggleOptionalAttendee(passKey, targetDate, isAdding) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    var holidayData = loadHolidays(ss);
    var data = dbSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(passKey)) continue;
      var datesG = getSafeString(data[i][6]).split(',').map(function(s){return s.trim();}).filter(String);
      var datesH = getSafeString(data[i][7]).split(',').map(function(s){return s.trim();}).filter(String);
      var memo = String(data[i][8] || '');
      var totalWeeks = Number(data[i][5]);
      var originalStart = getOriginalStartDate(datesG, datesH);
      var tag = '[자율참석: ' + targetDate + ']';

      if (isAdding) {
        if (memo.indexOf(tag) === -1) memo = memo ? memo + '\n' + tag : tag;
      } else {
        memo = memo.replace('\n' + tag, '').replace(tag, '').trim();
      }

      var attendedOpts = getAttendedOptionals(memo);
      var excludeList = holidayData.all.filter(function(d) {
        return d !== originalStart && attendedOpts.indexOf(d) === -1;
      }).concat(datesH);
      var newDatesG = recalcDates(originalStart, totalWeeks, excludeList);

      dbSheet.getRange(i + 1, 7).setValue("'" + newDatesG.join(', '));
      dbSheet.getRange(i + 1, 9).setValue(memo);
      return { success: true };
    }
    return { success: false };
  } catch(ex) { return { success: false, reason: ex.toString() }; }
}


// ──────────────────────────────────────────────
// 19. 전체 회원 목록 조회 (관리자)
//    ✅ [수정] maxPauses에 특정 회원 정책 반영
// ──────────────────────────────────────────────
function getAdminMembersList() {
  try {
    var ss      = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    var dbData  = dbSheet.getDataRange().getValues();
    var membersMap = {};
    var today = new Date(); today.setHours(0, 0, 0, 0);

    for (var i = 1; i < dbData.length; i++) {
      var row = dbData[i];
      if (!row[0]) continue;
      var phone = String(row[2]).replace(REGEX_NON_NUM, '');
      var type  = String(row[3]);
      var datesG = getSafeString(row[6]).split(',').map(function(s){return s.trim();}).filter(String);
      var datesH = getSafeString(row[7]).split(',').map(function(s){return s.trim();}).filter(String);
      var totalWeeks = Number(row[5]);

      var status = '만료'; var remW = 0; var remD = 0;
      if (datesG.length > 0) {
        var lastDateObj = parseSafeDate(datesG[datesG.length - 1]);
        if (parseSafeDate(datesG[0]) > today) {
          status = '수강예정';
        } else if (lastDateObj >= today) {
          status = '수강중';
          remW = datesG.filter(function(d){ return parseSafeDate(d) >= today; }).length;
          remD = Math.floor((lastDateObj - today) / (1000 * 60 * 60 * 24));
        }
        // 특정 만료 14일 초과 시 자동 정규 강등
        if (type === '특정' && status === '만료') {
          var diffDays = Math.floor((today - lastDateObj) / (1000 * 60 * 60 * 24));
          if (diffDays > 14) {
            type = '정규';
            dbSheet.getRange(i + 1, 4).setValue('정규');
          }
        }
      }

      var passObj = {
        passKey: String(row[0]),
        type: type,
        amount: Number(row[4]),
        totalWeeks: totalWeeks,
        usedPauses: datesH.length,
        maxPauses: calcMaxPauses(type, totalWeeks),
        startDate: datesG[0] || '',
        finalExp: datesG[datesG.length - 1] || '',
        status: status,
        remainingDays: remD,
        remainingWeeks: remW,
        upcomingDates: datesG,
        pauseHistory: datesH
      };
      var mKey = String(row[1]) + '_' + phone;
      if (!membersMap[mKey]) membersMap[mKey] = { name: String(row[1]), phone: phone, memo: String(row[8] || ''), passes: [] };
      membersMap[mKey].passes.push(passObj);
    }
    return Object.values(membersMap);
  } catch(ex) { return []; }
}


// ──────────────────────────────────────────────
// 20. 마스킹 맵 빌더 (수강 신청 화면용)
// ──────────────────────────────────────────────
function buildMaskMap(members) {
  var map2 = {}, map3 = {};
  members.forEach(function(m) {
    if (!m.phone || m.phone.length < 4) return;
    var key2 = m.name + '_' + m.phone.slice(-4, -2);
    var key3 = m.name + '_' + m.phone.slice(-4, -1);
    map2[key2] = (map2[key2] || 0) + 1;
    map3[key3] = (map3[key3] || 0) + 1;
  });
  return { count2: map2, count3: map3 };
}


// ──────────────────────────────────────────────
// 21. 초기 데이터 (수강 신청 화면 로드 시)
// ✅ [수정] 동일인(다중 결제자) 우선 병합 후 마스킹 처리하여 충돌 오류 완벽 제거
// ──────────────────────────────────────────────
function getInitialData() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var result = { holidays: [], holidayReasons: {}, holidayTypes: {}, members: [] };
  var today = new Date(); today.setHours(0, 0, 0, 0);

  try {
    var holidayData = loadHolidays(ss);
    result.holidays = holidayData.all;
    result.holidayReasons = holidayData.reasons;
    result.holidayTypes = holidayData.types;

    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    if (!dbSheet) return result;
    var dbData = dbSheet.getDataRange().getValues();

    // 1단계: 이름과 전화번호가 동일한 데이터를 '고유 회원(1명)'으로 완벽히 병합 (압축)
    var groupedMap = {};

    for (var i = 1; i < dbData.length; i++) {
      var row = dbData[i];
      if (!row[0] || !row[1]) continue;

      var name = String(row[1]).trim();
      var phone = String(row[2]).replace(REGEX_NON_NUM, '');
      if (phone.length < 4) continue; // 전화번호가 없는 비정상 데이터 무시

      var classify = String(row[3]);
      var datesG = getSafeString(row[6]).split(',').map(function(s){return s.trim();}).filter(String);

      var status = '만료';
      if (datesG.length > 0) {
        var firstDate = parseSafeDate(datesG[0]);
        var lastDate = parseSafeDate(datesG[datesG.length - 1]);
        if (firstDate > today) status = '수강예정';
        else if (lastDate >= today) status = '수강중';
      }

      var isEligible = false;
      var expDateStr = datesG.length > 0 ? datesG[datesG.length - 1] : '';
      if (expDateStr) {
        var expObj = parseSafeDate(expDateStr);
        var diffDays = Math.floor((today - expObj) / (1000 * 60 * 60 * 24));
        if (status === '수강중' || status === '수강예정' || (status === '만료' && diffDays <= 14)) isEligible = true;
        if (classify === '특정' && status === '만료' && diffDays > 14) classify = '정규';
      }

      var uniqueKey = name + '_' + phone;

      // 새 회원 등록 또는 기존 회원 데이터 최신화 (병합)
      if (!groupedMap[uniqueKey]) {
        groupedMap[uniqueKey] = {
          name: name,
          phone: phone,
          type: classify,
          expDate: expDateStr,
          status: status,
          isEligible: isEligible,
          occupiedDates: datesG.slice()
        };
      } else {
        var existing = groupedMap[uniqueKey];
        // 수강일 합치기 (중복 제거)
        datesG.forEach(function(d){ if (existing.occupiedDates.indexOf(d) === -1) existing.occupiedDates.push(d); });
        // 자격 및 만료일 업데이트
        if (isEligible) existing.isEligible = true;
        if (expDateStr > existing.expDate) existing.expDate = expDateStr;
        if (classify === '특정') existing.type = '특정';
      }
    }

    // 1-B단계: 원데이 할인 가능 여부 조회 (14일 이내 확정 원데이)
    var onedayDiscountMap = {};
    try {
      var odLog = ss.getSheetByName(SHEET_LOG);
      if (odLog) {
        var odData = odLog.getDataRange().getValues();
        var odToday = new Date(); odToday.setHours(0, 0, 0, 0);
        for (var oi = 1; oi < odData.length; oi++) {
          var odRow = odData[oi];
          if (odRow[7] !== true) continue;
          var odOpt = String(odRow[3]);
          if (odOpt.indexOf('원데이') === -1) continue;
          var odName = String(odRow[1]).trim();
          var odPhone = String(odRow[2]).replace(/[^0-9]/g, '');
          if (!odName || odPhone.length < 4) continue;
          var odDateStr = String(odRow[4] || '').split('~')[0].trim();
          var odDate = parseSafeDate(odDateStr);
          if (!odDate) continue;
          odDate.setHours(0, 0, 0, 0);
          if (odDate > odToday) continue;
          var odDiff = Math.floor((odToday - odDate) / 86400000);
          if (odDiff > 14) continue;
          var odPersons = 1;
          var odPm = odOpt.match(/(\d+)명/);
          if (odPm) odPersons = parseInt(odPm[1]);
          var odAmt = Number(odRow[5]) || 0;
          var odPerPerson = odPersons > 0 ? Math.round(odAmt / odPersons) : odAmt;
          if (odPerPerson <= 0) continue;
          var odKey = odName + '_' + odPhone;
          if (!onedayDiscountMap[odKey] || odDateStr > onedayDiscountMap[odKey].date) {
            onedayDiscountMap[odKey] = { price: odPerPerson, date: odDateStr };
          }
        }
      }
    } catch(odE) { Logger.log('[getInitialData onedayDiscount] ' + odE); }

    // 2단계: 압축된 고유 회원 목록을 배열로 변환
    var uniqueMembersArr = [];
    for (var key in groupedMap) {
      uniqueMembersArr.push(groupedMap[key]);
    }

    // 3단계: 고유 회원 배열을 바탕으로 마스킹 맵 빌드 (진짜 동명이인만 검출)
    var maskMaps = buildMaskMap(uniqueMembersArr);

    // 4단계: 동명이인 충돌 여부를 판단하여 결과값 생성
    uniqueMembersArr.forEach(function(m) {
      var k2 = m.name + '_' + m.phone.slice(-4, -2);
      var k3 = m.name + '_' + m.phone.slice(-4, -1);
      var share2 = maskMaps.count2[k2] || 0;
      var share3 = maskMaps.count3[k3] || 0;

      var finalMask = '';
      var isCollision = false;

      // 할당 로직 — 동명이인을 구분할 수 있는 *최소* 자릿수만 공개
      //   끝2 unique → "XX**"
      //   끝3 unique → "XXX*"
      //   끝3도 충돌 → 마스크 비움 (last-4 노출 금지) + 서버 검증 경로로 유도
      if (share2 === 1) finalMask = m.phone.slice(-4, -2) + '**';
      else if (share3 === 1) finalMask = m.phone.slice(-4, -1) + '*';
      else { finalMask = ''; isCollision = true; }

      var odInfo = onedayDiscountMap[m.name + '_' + m.phone];
      result.members.push({
        name: m.name,
        phoneMask: finalMask,
        isCollision: isCollision,
        type: m.type,
        expDate: m.expDate,
        isEligible: m.isEligible,
        occupiedDates: m.occupiedDates,
        onedayDiscount: odInfo ? odInfo.price : 0  // 1인당 원데이 할인 금액 (0=없음)
      });
    });

  } catch(ex) { Logger.log(ex); }

  return result;
}


// ──────────────────────────────────────────────
// 21-A. 동명이인 충돌 — 끝 4자리 서버 검증
//   클라이언트가 last-4 를 모른 채(getInitialData 가 비워서 보냄) 사용자 입력으로만
//   본인 식별을 마쳐야 하므로, 서버에서 DB 조회 후 매칭된 회원만 반환.
//   매칭 실패/모호 시 null. 매칭된 회원에게만 last-4 를 phoneMask 로 돌려준다
//   (사용자 본인이 직접 입력한 값이므로 신규 정보 노출 아님).
// ──────────────────────────────────────────────
function verifyCollisionMember(name, last4) {
  try {
    name = String(name || '').trim();
    last4 = String(last4 || '').replace(REGEX_NON_NUM, '');
    if (!name || last4.length !== 4) return null;

    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    if (!dbSheet) return null;
    var dbData = dbSheet.getDataRange().getValues();
    var today = new Date(); today.setHours(0, 0, 0, 0);

    // 같은 name + last-4 인 row 들을 모아 병합 (getInitialData 와 동일 로직)
    var found = {};
    for (var i = 1; i < dbData.length; i++) {
      var row = dbData[i];
      if (!row[0] || !row[1]) continue;
      var rName = String(row[1]).trim();
      var rPhone = String(row[2]).replace(REGEX_NON_NUM, '');
      if (rPhone.length < 4) continue;
      if (rName !== name || rPhone.slice(-4) !== last4) continue;

      var classify = String(row[3]);
      var datesG = getSafeString(row[6]).split(',').map(function(s){return s.trim();}).filter(String);

      var status = '만료';
      if (datesG.length > 0) {
        var firstDate = parseSafeDate(datesG[0]);
        var lastDate = parseSafeDate(datesG[datesG.length - 1]);
        if (firstDate > today) status = '수강예정';
        else if (lastDate >= today) status = '수강중';
      }
      var isEligible = false;
      var expDateStr = datesG.length > 0 ? datesG[datesG.length - 1] : '';
      if (expDateStr) {
        var expObj = parseSafeDate(expDateStr);
        var diffDays = Math.floor((today - expObj) / 86400000);
        if (status === '수강중' || status === '수강예정' || (status === '만료' && diffDays <= 14)) isEligible = true;
        if (classify === '특정' && status === '만료' && diffDays > 14) classify = '정규';
      }
      var uniqueKey = rName + '_' + rPhone;
      if (!found[uniqueKey]) {
        found[uniqueKey] = {
          name: rName, phone: rPhone, type: classify, expDate: expDateStr,
          status: status, isEligible: isEligible,
          occupiedDates: datesG.slice()
        };
      } else {
        var e = found[uniqueKey];
        datesG.forEach(function(d){ if (e.occupiedDates.indexOf(d) === -1) e.occupiedDates.push(d); });
        if (isEligible) e.isEligible = true;
        if (expDateStr > e.expDate) e.expDate = expDateStr;
        if (classify === '특정') e.type = '특정';
      }
    }

    var keys = Object.keys(found);
    if (keys.length === 0) return null;
    if (keys.length > 1) {
      // 같은 이름 + 같은 last-4 인 회원이 2명 이상 — 시스템 한계(매우 드묾)
      Logger.log('[verifyCollisionMember] 같은 이름+끝4자리 ' + keys.length + '명: ' + name + '/' + last4);
      return null;
    }
    var m = found[keys[0]];
    return {
      name: m.name,
      phoneMask: last4,        // 사용자가 직접 입력한 값을 그대로 반환
      isCollision: false,      // 인증되었으므로 이후 흐름에서는 충돌 아님
      type: m.type,
      expDate: m.expDate,
      isEligible: m.isEligible,
      occupiedDates: m.occupiedDates
    };
  } catch(ex) {
    Logger.log('[verifyCollisionMember] ' + ex);
    return null;
  }
}


// ──────────────────────────────────────────────
// 22. 회원 상세 조회 (관리자)
// ──────────────────────────────────────────────
function getAdminMemberDetail(name, phone) {
  try {
    var all = getAdminMembersList();
    var found = all.filter(function(m){ return m.name === name && m.phone === phone; });
    return found.length > 0 ? found[0] : null;
  } catch(ex) { return null; }
}


// ──────────────────────────────────────────────
// 23. 수강권 긴급 수정 (관리자)
// ──────────────────────────────────────────────
function emergencyEditPass(passKey, newType, newAmount, newDatesStr, newPausesStr) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    var data = dbSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(passKey)) continue;
      dbSheet.getRange(i + 1, 4).setValue(newType);
      dbSheet.getRange(i + 1, 5).setValue(newAmount);
      dbSheet.getRange(i + 1, 7).setValue("'" + newDatesStr);
      dbSheet.getRange(i + 1, 8).setValue(newPausesStr ? "'" + newPausesStr : '');
      return true;
    }
    return false;
  } catch(ex) { return false; }
}


// ──────────────────────────────────────────────
// 24. 회원 메모 저장 (관리자)
//    ✅ [유지] [날짜 관리자]: 내용 형식 (V5.8 방식)
// ──────────────────────────────────────────────
function updateMemberMemoAdmin(name, fullPhone, newMemo) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    var data = dbSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowPhone = String(data[i][2]).replace(REGEX_NON_NUM, '');
      if (String(data[i][1]) === name && rowPhone === fullPhone) {
        var oldMemo = String(data[i][8] || '');
        var combined = (oldMemo ? oldMemo + '\n' : '') + '[' + formatDateSafe(new Date()) + ' 관리자]: ' + newMemo;
        dbSheet.getRange(i + 1, 9).setValue(combined.trim());
        return true;
      }
    }
    return false;
  } catch(ex) { return false; }
}
// ══════════════════════════════════════════════════════════════
// [SMS 자동화 시스템] V1.0 — 섹션 24 닫는 괄호(}) 직후에 붙여넣기
//
// ⚠️ 기존 Gemini 블록 + 맨 아래 중복 doPost/onEdit 전부 삭제 후 교체
//
// 적용 순서:
// 1. Code.gs에서 섹션 24 아래의 모든 코드 삭제
// 2. 이 파일 내용 전체 붙여넣기
// 3. GAS → 프로젝트 설정 → 스크립트 속성:
//    GEMINI_API_KEY = 발급받은 키
//    INSTRUCTOR_PHONE = 01043130150
//    MACRODROID_WEBHOOK = https://trigger.macrodroid.com/1a6d5e4b-6484-46d6-ae82-8b5c0410f044/send_sms
// 4. setupSMSSheets() 한 번 실행 (시트 생성)
// 5. setupAllTriggers() 한 번 실행 (트리거 설정)
// ══════════════════════════════════════════════════════════════


// ──────────────────────────────────────────────
// 25. SMS 시스템 상수
// ──────────────────────────────────────────────
var SHEET_SETTINGS     = '⚙️ 설정';
var SHEET_PROMPTS      = '✏️ Gemini 프롬프트';
var SHEET_SMS_QUEUE    = '📨 문자 발송 대기열';
var SHEET_SMS_ARCHIVE  = '📦 대기열 아카이브';

var GEMINI_MODEL = 'gemini-2.0-flash';

var BOOKING_LINK = 'https://m.site.naver.com/1SFZj'; // 수강 신청 링크

var SMS_STATUS = {
  IMMEDIATE: '즉시발송',
  WAITING:   '대기',
  APPROVAL:  '승인대기',
  PROCESSING:'처리중',
  DONE:      '발송완료',
  FAILED:    '발송실패'
};

function checkAvailableModels() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log(res.getContentText());
}

// ──────────────────────────────────────────────
// 26. 설정 시트 읽기
// ──────────────────────────────────────────────
function getSettings() {
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_SETTINGS);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = data[i][1];
    if (key) settings[key] = val;
  }
  return settings;
}

function getSetting(key, fallback) {
  var s = getSettings();
  return s[key] !== undefined ? s[key] : fallback;
}


// ──────────────────────────────────────────────
// 27. Gemini 프롬프트 시트 읽기
// ──────────────────────────────────────────────
function getGeminiPrompt(type) {
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_PROMPTS);
  if (!sheet) return '';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === type) return String(data[i][1] || '');
  }
  // ── 금주 공지 행이 없으면 자동 생성 ──────────────────
  if (type === '금주 공지') {
    try {
      var lastRow = sheet.getLastRow() + 1;
      sheet.getRange(lastRow, 1, 1, 3).setValues([[
        '금주 공지',
        '',
        '이번 주 전날알림에만 포함할 한 줄 공지. 없으면 비워두세요.'
      ]]);
      sheet.getRange(lastRow, 1).setBackground('#f5f5f5').setFontWeight('bold');
      sheet.getRange(lastRow, 2).setBackground('#fff9c4').setWrap(true);
      sheet.getRange(lastRow, 3).setBackground('#f5f5f5').setFontColor('#888888').setWrap(true);
      sheet.setRowHeight(lastRow, 80);
      Logger.log('[getGeminiPrompt] 금주 공지 행 자동 생성 (행 ' + lastRow + ')');
    } catch(e) {
      Logger.log('[getGeminiPrompt] 금주 공지 행 생성 실패: ' + e);
    }
  }
  return '';
}


// ──────────────────────────────────────────────
// 28. Gemini API 호출
// ──────────────────────────────────────────────
// useSearch: bool — Grounding 검색 사용 여부
// forceJSON: bool — responseMimeType application/json 강제
function callGemini(prompt, useSearch, forceJSON) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY 스크립트 속성 없음');

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  var genCfg = { temperature: 0.5, maxOutputTokens: 1024 };
  if (forceJSON) genCfg.responseMimeType = 'application/json';

  var body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: genCfg
  };
  if (useSearch) body.tools = [{ googleSearch: {} }];

  var maxRetries = 3;
  var waitMs = [5000, 15000, 30000];

  for (var attempt = 0; attempt < maxRetries; attempt++) {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    var statusCode = res.getResponseCode();
    var rawText    = res.getContentText();
    var json;
    try { json = JSON.parse(rawText); }
    catch(pe) { throw new Error('Gemini: 응답 파싱 실패 (' + statusCode + ') ' + rawText.slice(0,200)); }

    if (json.error) {
      var errMsg = json.error.message || String(json.error);
      // 429 할당량 초과 / 503 과부하 → 재시도
      if ((statusCode === 429 || statusCode === 503 ||
           errMsg.indexOf('RESOURCE_EXHAUSTED') !== -1 ||
           errMsg.indexOf('overloaded') !== -1 ||
           errMsg.indexOf('high demand') !== -1) && attempt < maxRetries - 1) {
        Logger.log('[Gemini] HTTP ' + statusCode + ' 재시도 ' + (attempt+1) + '/' + maxRetries + ' — ' + errMsg);
        Utilities.sleep(waitMs[attempt]);
        continue;
      }
      throw new Error('Gemini API 오류: ' + errMsg);
    }

    var candidate = json.candidates && json.candidates[0];
    if (!candidate || !candidate.content) throw new Error('Gemini: candidate 없음');

    // STOP 이외의 finishReason 재시도
    if (candidate.finishReason && candidate.finishReason !== 'STOP') {
      if (attempt < maxRetries - 1) {
        Logger.log('[Gemini] finishReason=' + candidate.finishReason + ', 재시도 ' + (attempt+1));
        Utilities.sleep(waitMs[attempt]);
        continue;
      }
      throw new Error('Gemini: 응답 잘림 (finishReason: ' + candidate.finishReason + ')');
    }

    return candidate.content.parts
      .filter(function(p) { return p.text; })
      .map(function(p) { return p.text; })
      .join('').trim();
  }

  throw new Error('Gemini: 최대 재시도 횟수 초과');
}

function callGeminiJSON(prompt) {
  var raw = '';
  try {
    raw = callGemini(prompt, false, true); // forceJSON=true
    // forceJSON 모드면 깨끗한 JSON이 오지만 방어적으로 마크다운 제거
    var cleaned = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/,'').trim();
    // 배열 or 객체 시작 위치 찾기
    var start = cleaned.search(/[\[{]/);
    if (start > 0) cleaned = cleaned.slice(start);
    return JSON.parse(cleaned);
  } catch(ex) {
    Logger.log('[callGeminiJSON] 파싱 실패 — ' + ex + '\n원본: ' + String(raw).slice(0, 500));
    return [];
  }
}



// ──────────────────────────────────────────────
// 29. 숫자 → 한국어 변환
// ──────────────────────────────────────────────
var KOREAN_NUMS = ['영','한','두','세','네','다섯','여섯','일곱','여덟','아홉','열',
  '열한','열두','열세','열네','열다섯','열여섯','열일곱','열여덟','열아홉','스물',
  '스물한','스물두','스물세','스물네','스물다섯','스물여섯','스물일곱','스물여덟','스물아홉','서른'];

function toKoreanNum(n) {
  n = Number(n);
  if (n >= 0 && n <= 30 && KOREAN_NUMS[n]) return KOREAN_NUMS[n];
  return String(n);
}

function toKoreanWeeks(weeks) {
  return toKoreanNum(weeks) + ' 주';
}

function toKoreanPause(count) {
  return toKoreanNum(count) + ' 번';
}


// ──────────────────────────────────────────────
// 30. 대기열에 문자 추가 (핵심 함수)
// ──────────────────────────────────────────────
function enqueueSMS(phone, name, message, type, status, scheduledTime) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_SMS_QUEUE);
    if (!sheet) { Logger.log('대기열 시트 없음'); return false; }

    var cleanPhone = String(phone).replace(/[^0-9]/g, '');
    if (cleanPhone.indexOf('10') === 0 && cleanPhone.length >= 10) cleanPhone = '0' + cleanPhone;

    var lastRow = sheet.getLastRow();
    var seqNum = lastRow >= 2 ? lastRow : 1;
    var finalStatus = status || SMS_STATUS.WAITING;
    var schedTime = scheduledTime || '';

    // WAITING 중복 체크 — 같은 전화+유형이 이미 대기 중이면 스킵
    if (finalStatus === SMS_STATUS.WAITING) {
      var existingData = sheet.getDataRange().getValues();
      for (var _ck = 1; _ck < existingData.length; _ck++) {
        if (String(existingData[_ck][2]).replace(/[^0-9]/g, '') === cleanPhone
            && existingData[_ck][3] === type
            && existingData[_ck][5] === SMS_STATUS.WAITING) {
          Logger.log('[enqueueSMS] 중복 스킵: ' + name + ' / ' + type);
          return false;
        }
      }
    }

    sheet.insertRowAfter(1);
    var newRow = 2;
    sheet.getRange(newRow, 1, 1, 8).setValues([[seqNum, name, "'" + cleanPhone, type, message, finalStatus, schedTime, '']]);
    sheet.getRange(newRow, 5).setWrap(true);

    // 상태별 배경색 + 글자색 명시 (헤더 흰 글씨 상속 방지)
    var colors = {};
    colors[SMS_STATUS.IMMEDIATE]  = '#ffebee';  // 연빨강
    colors[SMS_STATUS.WAITING]    = '#e8f5e9';  // 연초록
    colors[SMS_STATUS.APPROVAL]   = '#fff8e1';  // 연노랑
    var bg = colors[finalStatus] || '#ffffff';
    sheet.getRange(newRow, 1, 1, 8).setBackground(bg);
    sheet.getRange(newRow, 1, 1, 8).setFontColor('#1A1A17');

    // 즉시발송이면 바로 처리
    if (finalStatus === SMS_STATUS.IMMEDIATE) {
      _sendQueueRow(sheet, newRow);
    }

    return true;
  } catch(ex) {
    Logger.log('[enqueueSMS] ' + ex);
    return false;
  }
}


// ──────────────────────────────────────────────
// 31. 대기열 행 발송 처리
// ──────────────────────────────────────────────
function _sendQueueRow(sheet, row) {
  try {
    var rawPhone = String(sheet.getRange(row, 3).getValue());
    var phone = rawPhone.replace(/[^0-9]/g, '');
    if (phone.indexOf('10') === 0 && phone.length >= 10) phone = '0' + phone;
    var msg = String(sheet.getRange(row, 5).getValue());

    if (!phone || !msg) {
      sheet.getRange(row, 6).setValue(SMS_STATUS.FAILED);
      sheet.getRange(row, 8).setValue('번호 또는 내용 없음');
      return false;
    }

    sheet.getRange(row, 6).setValue(SMS_STATUS.PROCESSING);

    var webhookUrl = PropertiesService.getScriptProperties().getProperty('MACRODROID_WEBHOOK');
    if (!webhookUrl) {
      sheet.getRange(row, 6).setValue(SMS_STATUS.FAILED);
      sheet.getRange(row, 8).setValue('MACRODROID_WEBHOOK 미설정');
      return false;
    }

    var payload = '[PHONE]' + phone + '[SMS]' + msg;
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'text/plain; charset=utf-8',
      payload: payload,
      muteHttpExceptions: true
    });

    sheet.getRange(row, 6).setValue(SMS_STATUS.DONE);
    sheet.getRange(row, 8).setValue(new Date());
    sheet.getRange(row, 1, 1, 8).setBackground(null);
sheet.getRange(row, 1, 1, 8).setFontColor(null);
    return true;
  } catch(ex) {
    sheet.getRange(row, 6).setValue(SMS_STATUS.FAILED);
    sheet.getRange(row, 8).setValue(ex.toString().substring(0, 100));
    return false;
  }
}


// ──────────────────────────────────────────────
// 32. sendSmsOnEdit — 즉시발송 감지 (설치형 트리거)
//     ✅ 기존 sendSmsOnEdit 교체
// ──────────────────────────────────────────────
function sendSmsOnEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_SMS_QUEUE) return;

  // F열(6) 상태가 "즉시발송"으로 변경됐을 때
  if (e.range.getColumn() === 6 && String(e.value) === SMS_STATUS.IMMEDIATE) {
    _sendQueueRow(sheet, e.range.getRow());
  }
}


// ──────────────────────────────────────────────
// 33. 예약 대기 항목 일괄 발송
//     매일 트리거로 실행 (설정 시트의 발송 시간)
// ──────────────────────────────────────────────
function processScheduledQueue() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_SMS_QUEUE);
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][5]) === SMS_STATUS.WAITING) {
        _sendQueueRow(sheet, i + 1);
        // 3초 대기 — MacroDroid 누락 방지 + 배터리
        Utilities.sleep(3000);
      }
    }
  } catch(ex) { Logger.log('[processScheduledQueue] ' + ex); }
}


// ──────────────────────────────────────────────
// 34. 대기열 아카이브 (매일 새벽 실행)
// ──────────────────────────────────────────────
function archiveQueue() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var qSheet = ss.getSheetByName(SHEET_SMS_QUEUE);
    var aSheet = ss.getSheetByName(SHEET_SMS_ARCHIVE);
    if (!qSheet || !aSheet) return;

    var retainDays = Number(getSetting('대기열 보관 기간', 30));
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - retainDays);

    var data = qSheet.getDataRange().getValues();
    var rowsToDelete = [];

    for (var i = data.length - 1; i >= 1; i--) {
      var status = String(data[i][5]);
      var doneTime = data[i][7];
      if (status !== SMS_STATUS.DONE && status !== SMS_STATUS.FAILED) continue;
      if (!doneTime) continue;

      var doneDate = (doneTime instanceof Date) ? doneTime : new Date(doneTime);
      if (isNaN(doneDate.getTime())) continue;

      if (doneDate < cutoff) {
        aSheet.appendRow(data[i]);
        rowsToDelete.push(i + 1);
      }
    }

    // 역순 삭제 (행 번호 꼬임 방지)
    for (var j = 0; j < rowsToDelete.length; j++) {
      qSheet.deleteRow(rowsToDelete[j]);
    }
  } catch(ex) { Logger.log('[archiveQueue] ' + ex); }
}


// ──────────────────────────────────────────────
// 35. 결제확정 문자 (템플릿 방식)
//     ✅ [V5.12] 정규(신규)/원데이 → 상세 안내 포함
//               연장/특정 → 간결 버전 유지
// ──────────────────────────────────────────────
function getSMSTemplate(type, vars) {
  var s             = getSettings();
  var closing       = String(s['끝인사'] || '평온한 하루 되시기 바랍니다.');
  var studioMain    = getSMSStudioInfoMain();
  var studioParking = getSMSParkingInfo();

  var templates = {
    '결제확정정규신규':
  '[Root Vinyasa]\n{이름}님, 확인했습니다.\n\n' +
  '{시작일}부터 {총회차}번이에요.\n' +
  '사정이 생기시면 {쉬어가기}은 쉬어갈 수 있어요.' +
  studioMain + '\n' +
  studioParking + '\n\n' +
  '{메모확인}\n\n' + closing,

    '결제확정정규연장특정':
      '[Root Vinyasa]\n{이름}님, 확인했습니다.\n\n' +
      '{시작일} 토요일에 뵙겠습니다.\n' +
      '{기간} 동안 함께하게 됩니다.\n' +
      '사정이 생기시면 {쉬어가기}은 쉬어갈 수 있어요.\n' +
      '{메모확인}\n\n' + closing,

    '결제확정원데이혼자':
      '[Root Vinyasa]\n{이름}님, 확인했습니다.\n\n' +
      '{시작일} 토요일에 뵙겠습니다.' +
      studioMain + '\n{메모확인}' +
      studioParking + '\n\n' + closing,

    '결제확정원데이인↑':
      '[Root Vinyasa]\n{이름}님, 확인했습니다.\n\n' +
      '{시작일} 토요일, {인원}분 오시는 날이에요.' +
      studioMain + '\n{메모확인}' +
      studioParking + '\n\n' + closing,

    '결제확정원데이할인':
      '[Root Vinyasa]\n{이름}님, 확인했습니다.\n\n' +
      '{시작일}부터 {총회차}번이에요.\n' +
      '(원데이 1회 포함 {기간}권 · {할인금액}원 차감)\n' +
      '사정이 생기시면 {쉬어가기}은 쉬어갈 수 있어요.' +
      studioMain + '\n' +
      studioParking + '\n\n' +
      '{메모확인}\n\n' + closing
  };

  var tmpl = templates[type];
  if (!tmpl) return null;
  for (var key in vars) {
    tmpl = tmpl.replace(new RegExp('\\{' + key + '\\}', 'g'), String(vars[key]));
  }
  tmpl = tmpl.replace(/\{[^}]+\}/g, '').replace(/\n{3,}/g, '\n\n').trim();
  return tmpl;
}


// ──────────────────────────────────────────────
// 36. 결제 확인 → 문자 발송 연결
//     confirmPaymentAdmin() 호출 후 실행
// ──────────────────────────────────────────────

function sendPaymentConfirmedSMS(name, phone, option, schedule, memo) {
  try {
    var memoLine = (memo && memo.trim()) ? '작성해주신 메모도 확인했습니다.' : '';
    var startDate = String(schedule || '').split('~')[0].trim();
    var endDate   = String(schedule || '').split('~')[1] ? String(schedule || '').split('~')[1].trim() : '';
    var d = parseSafeDate(startDate);
    var dateLabel = d ? ((d.getMonth() + 1) + '월 ' + d.getDate() + '일') : startDate;

    // 교체 핵심: 메세지 변수를 최상단에 단 한 번만 선언하여 충돌 방지
    var msg = '';

    if (option.indexOf('원데이') !== -1) {
      // ── 원데이 ──
      var persons = 1;
      var m = option.match(/(\d+)명/);
      if (m) persons = parseInt(m[1]);

      if (persons >= 2) {
        msg = getSMSTemplate('결제확정원데이인↑', { '이름': name, '시작일': dateLabel, '인원': toKoreanNum(persons), '메모확인': memoLine });
      } else {
        msg = getSMSTemplate('결제확정원데이혼자', { '이름': name, '시작일': dateLabel, '메모확인': memoLine });
      }
      enqueueSMS(phone, name, msg, '결제확정', SMS_STATUS.IMMEDIATE);

    } else {
      // ── 정규 / 연장 / 특정 ──
      var weeks = 4;
      if (option.indexOf('12주') >= 0 || option.indexOf('12회') >= 0) weeks = 12;
      else if (option.indexOf('8주') >= 0 || option.indexOf('8회') >= 0) weeks = 8;
      else if (option.indexOf('5회') >= 0 || option.indexOf('5주') >= 0) weeks = 5;

      var isExtension = option.indexOf('연장') >= 0;
      var isSpecial   = option.indexOf('특별') >= 0 || option.indexOf('특정') >= 0;
      var type        = isSpecial ? '특정' : '정규';

      // 기존에 특정 회원이었는지 DB 확인
      if (!isSpecial) {
        try {
          var dbSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_DB_NEW);
          var dbData  = dbSheet.getDataRange().getValues();
          var cleanPhone = String(phone).replace(/[^0-9]/g, '');
          for (var k = 1; k < dbData.length; k++) {
            var rPhone = String(dbData[k][2]).replace(/[^0-9]/g, '');
            if (String(dbData[k][1]) === name && rPhone === cleanPhone && String(dbData[k][3]) === '특정') {
              isSpecial = true; type = '특정'; break;
            }
          }
        } catch(ex) { Logger.log('[sendPaymentConfirmedSMS DB lookup] ' + ex); }
      }

      var maxP = calcMaxPauses(type, weeks); // 쉬어가기는 원본 weeks 기준
      var monthMap = { 4: '한 달', 5: '한 달', 8: '두 달', 12: '석 달' };
      var periodLabel = monthMap[weeks] || toKoreanWeeks(weeks);

      // 원데이 할인 적용 여부 확인 (로그 조회)
      var onedayDiscSMS = 0;
      if (option.indexOf('원데이 할인') >= 0) {
        try {
          var odSmsSS = SpreadsheetApp.openById(SS_ID);
          var odSmsRes = _findRecentOnedayLog(odSmsSS, name, String(phone).replace(/[^0-9]/g,''));
          if (odSmsRes.found) onedayDiscSMS = odSmsRes.perPersonPrice;
        } catch(odSmsEx) { Logger.log('[sendPaymentConfirmedSMS onedayDisc] ' + odSmsEx); }
      }

      if (onedayDiscSMS > 0) {
        // 원데이 할인 문자: 실제 수업 횟수 = weeks - 1, 쉬어가기 = 원본 weeks 기준
        msg = getSMSTemplate('결제확정원데이할인', {
          '이름': name,
          '시작일': dateLabel,
          '기간': periodLabel,
          '총회차': toKoreanNum(weeks - 1),
          '쉬어가기': toKoreanPause(maxP),
          '할인금액': onedayDiscSMS.toLocaleString(),
          '메모확인': memoLine
        });
      } else if (isExtension || isSpecial) {
        msg = getSMSTemplate('결제확정정규연장특정', {
          '이름': name,
          '시작일': dateLabel,
          '기간': periodLabel,
          '총회차': toKoreanNum(weeks),
          '쉬어가기': toKoreanPause(maxP),
          '메모확인': memoLine
        });
      } else {
        msg = getSMSTemplate('결제확정정규신규', {
          '이름': name,
          '시작일': dateLabel,
          '기간': periodLabel,
          '총회차': toKoreanNum(weeks),
          '쉬어가기': toKoreanPause(maxP),
          '메모확인': memoLine
        });
      }
      enqueueSMS(phone, name, msg, '결제확정', SMS_STATUS.IMMEDIATE);
    }
  } catch(ex) { Logger.log('[sendPaymentConfirmedSMS] ' + ex); }
}


// ──────────────────────────────────────────────
// 37. KB 입금 감지 (doPost) — 깨끗한 단일 버전
// ──────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    // 1. 데이터 수신 (쿼리 파라미터 방식 전용)

    // MacroDroid의 쿼리 파라미터 이름(fullText)을 직접 읽어옵니다.
    var rawText = (e && e.parameter) ? e.parameter.fullText : "";

    // 데이터가 비어있는 경우(테스트 실행 등) 처리
    if (!rawText) {
      Logger.log("수신된 데이터가 없습니다.");
      return ContentService.createTextOutput("NO_DATA_RECEIVED");
    }

    // 2. 텍스트 정규화 (줄바꿈, 다중 공백을 단일 공백 1개로 압축)
    var fullData = rawText.replace(/\s+/g, ' ').trim();

    // 3. 입금 정보 추출 (전자금융입금 기준)
    var depositor = '';
    var amount = 0;
    // 1순위: "이름 전자금융입금 금액" 패턴 (이름 1~20자, 공백 허용)
    var matchData = fullData.match(/([가-힣A-Za-z\s]{1,20}?)\s*전자금융입금\s+([\d,]+)/);
    // 2순위: "이름 입금 금액" 패턴 (타행이체 등 대체 표현)
    if (!matchData) matchData = fullData.match(/([가-힣A-Za-z\s]{1,20}?)\s*입금\s+([\d,]+)/);

    if (matchData) {
      depositor = matchData[1].replace(/\s+/g, '').trim();
      amount = Number(matchData[2].replace(/,/g, ''));
    }

    // [이 아래부터는 기존의 '파싱 실패 처리' 및 '매칭 로직'을 그대로 사용하세요]


    // 파싱 실패 → 강사님한테 원본 전달
    if (!depositor || amount === 0) {
      var instructorPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
      if (instructorPhone) {
        enqueueSMS(instructorPhone, '최승훈', 
          '[ROOT 입금 감지]\n파싱 실패\n\n원본: ' + fullData.substring(0, 120),
          '시스템알림', SMS_STATUS.IMMEDIATE);
      }
      return ContentService.createTextOutput('PARSE_FAIL');
    }

    // 미결제 대기열 대조
    var pending = getPendingPayments();
    var candidates = pending.filter(function(c) { return Number(c.amount) === amount; });
    var matched = null;

    // 1차: 이름 직접 매칭
    for (var i = 0; i < candidates.length; i++) {
      var cName = String(candidates[i].name).replace(/\s/g, '');
      var dName = depositor.replace(/\s/g, '');
      if (cName === dName || cName.indexOf(dName) !== -1 || dName.indexOf(cName) !== -1) {
        matched = candidates[i];
        break;
      }
    }

    // 2차: 후보 여러 명이면 Gemini 판단
    if (!matched && candidates.length > 0) {
      try {
        var list = candidates.map(function(c, idx) {
          return (idx + 1) + '. ' + c.name + ' / 메모: ' + (c.memo || '없음');
        }).join('\n');

        var r = callGeminiJSON(
          '입금자명: ' + depositor + '\n후보:\n' + list +
          '\n\n어느 신청자의 입금인지 판단. JSON만.\n{"index":1} (1부터, 없으면 0)'
        );
        var idx = Number(r.index || 0);
        if (idx >= 1 && idx <= candidates.length) matched = candidates[idx - 1];
      } catch(ex) { Logger.log('[doPost Gemini] ' + ex); }
    }

    // 매칭 성공
    if (matched) {
      var confirmed = confirmPaymentAdmin(matched.rowIdx);
      if (!confirmed) {
        var ipPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
        if (ipPhone) {
          enqueueSMS(ipPhone, '최승훈',
            '[ROOT 확인 실패]\n입금자: ' + depositor + '\n금액: ' + amount.toLocaleString() + '원\n' +
            '이름 매칭은 됐으나 DB 저장 실패.\n대시보드에서 수동 처리 필요.',
            '시스템알림', SMS_STATUS.IMMEDIATE);
        }
        return ContentService.createTextOutput('CONFIRM_FAIL');
      }
      sendPaymentConfirmedSMS(matched.name, matched.phone, matched.option, matched.schedule, matched.memo);
      return ContentService.createTextOutput('OK');
    }

    // 매칭 실패 → 강사님한테 알림
    var instructorPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
    if (instructorPhone) {
      enqueueSMS(instructorPhone, '최승훈',
        '[ROOT 미확인 입금]\n입금자: ' + depositor + '\n금액: ' + amount.toLocaleString() + '원\n일치하는 신청이 없습니다.',
        '시스템알림', SMS_STATUS.IMMEDIATE);
    }
    return ContentService.createTextOutput('NO_MATCH');

  } catch(ex) {
    Logger.log('[doPost] ' + ex);
    return ContentService.createTextOutput('ERROR');
  } finally {
    lock.releaseLock();
  }
}



// ──────────────────────────────────────────────
// 38. 매일 배치 — Gemini 일괄 문자 생성
// ──────────────────────────────────────────────
function dailyBatch() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    if (!dbSheet) return;

    var settings = getSettings();
    var today = new Date(); today.setHours(0, 0, 0, 0);

    var data = dbSheet.getDataRange().getValues();

    // 만료후 안부 대상 수집
    var followUpMembers = [];
    var sent = {};

    if (String(settings['수강 끝난 회원 안부']) === 'ON') {
      var followUpDays = String(settings['안부 며칠째에'] || '1,6,13')
        .split(',').map(function(s){ return Number(s.trim()); });

      for (var i = 1; i < data.length; i++) {
        var name  = String(data[i][1] || '').trim();
        var phone = String(data[i][2] || '').replace(/[^0-9]/g, '');
        var type  = String(data[i][3] || '');
        var totalW = Number(data[i][5] || 0);
        var datesG = getSafeString(data[i][6]).split(',').map(function(s){return s.trim();}).filter(Boolean);
        var datesH = getSafeString(data[i][7]).split(',').map(function(s){return s.trim();}).filter(Boolean);
        var memo   = String(data[i][8] || '');
        if (!name || phone.length < 10 || datesG.length === 0) continue;

        var key = name + '_' + phone;
        var finalObj = parseSafeDate(datesG[datesG.length - 1]);
        if (!finalObj) continue;

        var maxP = calcMaxPauses(type, totalW);
        var daysAfter = Math.floor((today - finalObj) / 86400000);

        if (followUpDays.indexOf(daysAfter) !== -1 && !sent[key + '_fu']) {
          followUpMembers.push({
            name: name, phone: phone, type: type,
            daysAfter: daysAfter,
            totalClasses: totalW,
            usedPauses: datesH.length,
            maxPauses: maxP,
            memo: memo.replace(/\[신체노트\]:[^\n]*/g, '').replace(/\[자율참석:[^\]]+\]/g, '').trim()
          });
          sent[key + '_fu'] = true;
        }
      }

      if (followUpMembers.length > 0) {
        _generateFollowUpMessages(followUpMembers, settings);
      }
    }

    // DB 서식 업데이트 (만료=회색, 임박=노랑, 예정=파랑)
    updateDBFormatting();

  } catch(ex) { Logger.log('[dailyBatch] ' + ex); }
}

// 헬퍼: 날짜가 수강 기간 범위 안인지
function _isInRange(dateStr, datesG) {
  if (datesG.length === 0) return false;
  var d = parseSafeDate(dateStr);
  var first = parseSafeDate(datesG[0]);
  var last = parseSafeDate(datesG[datesG.length - 1]);
  return d && first && last && d >= first && d <= last;
}


// ──────────────────────────────────────────────
// 39. Gemini 전날알림 생성 (로직 보존 + 무비용 최적화)
// ──────────────────────────────────────────────
function _generateReminderMessages(members, settings) {
  try {
    var stylePrompt = getGeminiPrompt('문자 스타일');
    var reminderPrompt = getGeminiPrompt('전날알림 지시');
    var weatherPrompt = getGeminiPrompt('날씨/오시는 길 지시');

    var location = settings['스튜디오 위치'] || '서울 용산구';
    var classDate = members[0].classDate;
    var classTime = settings['수업 시간'] || '10:30~11:40';

    // 1. 날씨 조회 (무비용 유지를 위해 검색 도구는 OFF 처리 권장)
    var weatherInfo = '';
    try {
      weatherInfo = callGemini(
        location + ' ' + classDate + ' ' + classTime + ' 시간대 날씨 예상.\n' +
        (weatherPrompt || '기온, 일교차 등 참고 정보.') +
        '\n사실만 간결하게 3줄 이내.',
        false // 할당량 보존을 위해 검색 기능 해제
      );
    } catch(ex) { Logger.log('[weather] ' + ex); }

    // 2. 인지적 분할 (3명씩 묶음 처리)
    var chunkSize = 3;
    for (var i = 0; i < members.length; i += chunkSize) {
      var chunk = members.slice(i, i + chunkSize);
      
      var memberData = chunk.map(function(m) {
  return {
    name: m.name,
    remaining: m.remainingClasses,
    totalWeeks: m.totalClasses,
    pauseRemaining: m.pauseStatus.remaining,
    isFirst: m.flags.isFirstClass,
    nextPassPurchased: m.flags.nextPassPurchased,
    holidays: m.upcomingHolidays,
    note: m.memberNote
  };
});

      // 사용자 요청 원본 로직 전체 수용
      var prompt =
        '요가 스튜디오 Root Vinyasa 문자 도우미.\n\n' +
        '[시트 설정 내용]\n- 스타일: ' + (stylePrompt || '') + '\n- 세부지시: ' + (reminderPrompt || '') + '\n\n' +
        '[내일 날씨/상황]\n' + (weatherInfo || '정보 없음') + '\n\n' +
        '[대상 회원 데이터]\n' + JSON.stringify(memberData, null, 2) + '\n\n' +
        'JSON 배열만 반환. [{"name":"..","message":".."}]';

      var results = callGeminiJSON(prompt);

      for (var j = 0; j < results.length; j++) {
        var r = results[j];
        var member = chunk.filter(function(m){ return m.name === r.name; })[0];
        if (member) {
          enqueueSMS(member.phone, member.name, r.message, '전날알림', SMS_STATUS.WAITING);
        }
      }

      // 3. API 안정성을 위한 20초 휴식
      if (i + chunkSize < members.length) {
        Utilities.sleep(20000); 
      }
    }
  } catch(ex) { Logger.log('[_generateReminderMessages] ' + ex); }
}





// ──────────────────────────────────────────────
// 41. Gemini 만료후 안부 생성 (승인대기)
// ──────────────────────────────────────────────
function _generateFollowUpMessages(members, settings) {
  try {
    var stylePrompt = getGeminiPrompt('문자 스타일');
    var followUpPrompt = getGeminiPrompt('만료후 안부 지시');

    var memberData = members.map(function(m) {
      return { name: m.name, daysAfter: m.daysAfter, type: m.type, totalClasses: m.totalClasses, memo: m.memo || '' };
    });

    var prompt =
      '당신은 요가 스튜디오 Root Vinyasa 문자 담당자입니다.\n' +
      '수강이 끝난 회원에게 따뜻한 안부 문자를 작성해주세요.\n\n' +

      '[문자 구조 — 반드시 이 순서와 형식을 따를 것]\n' +
      '1) [Root Vinyasa]\n' +
      '2) 안녕하세요, {이름}님!\n' +
      '3) (빈 줄)\n' +
      '4) 안부 메시지: 짧은 문장 1~2개. 13~15자 내외 지점에서 \\n 줄바꿈.\n' +
      '   종료 후 며칠이 지났는지에 따라 거리감을 조절:\n' +
      '   1일→따뜻하고 가깝게, 6일→자연스럽게, 13일→가볍고 담담하게.\n' +
      '5) (빈 줄)\n' +
      '6) 계절감 한 줄 + \\n편안한 하루 보내세요.\n\n' +

      '[절대 금지]\n' +
      '- 이모지\n' +
      '- "다음 패스 기다릴게요", "언제든 돌아오세요" 등 재등록 유도 표현\n' +
      '- "함께해 주셔서 좋았어요" 등 작별·판매성 문구\n' +
      '- 내일 뵙겠습니다 (안부 문자에는 사용 금지)\n' +
      '- "~하세요", "~해보세요" 등 지시형 말투\n' +
      '- 숫자 숫자 표기 (1→한 번, 2→두 번 ...)\n\n' +

      '[브랜드 스타일]\n' + (stylePrompt || '1인칭 현재형. 짧은 문장. 이모지 없음. 관찰하고 안내하는 말투.') + '\n\n' +

      '[안부 규칙]\n' + (followUpPrompt || '안부 위주, 판매 금지. 거리감 조절: 1일=따뜻, 6일=자연, 13일=가볍게.') + '\n\n' +

      '[회원 데이터]\n' + JSON.stringify(memberData) + '\n\n' +
      '출력: JSON 배열만. 앞뒤 텍스트 없이.\n' +
      '[{"name":"..","message":".."}]';

    var results = callGeminiJSON(prompt);
    for (var i = 0; i < results.length; i++) {
      var member = members.filter(function(m){ return m.name === results[i].name; })[0];
      if (member) {
        // 만료후 안부 — 재등록 링크 첨부
        var fuMsg = (results[i].message || '').trim() + '\n\n신청 · ' + BOOKING_LINK;
        enqueueSMS(member.phone, member.name, fuMsg, '만료후안부', SMS_STATUS.APPROVAL);
      }
    }
  } catch(ex) { Logger.log('[_generateFollowUpMessages] ' + ex); }
}


// ──────────────────────────────────────────────
// 42. 휴강 등록 시 문자 처리 (addHoliday 교체)
//     ✅ 타이밍 분기 + 원데이 예외
// ──────────────────────────────────────────────
function addHoliday(dateStr, reason, type) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var hSheet = ss.getSheetByName(SHEET_HOLIDAY);
    if (!hSheet) return false;
    var safeType = (type === '선택') ? '선택' : '일반';
    var dateObj = parseSafeDate(dateStr);
    if (!dateObj) return false;

    // 휴강 캘린더에 추가
    hSheet.appendRow([dateObj, reason || '', safeType]);
    rebuildAllMemberDates();

    // 남은 일수 계산
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var daysUntil = Math.floor((dateObj - today) / 86400000);
    var holidayLabel = (dateObj.getMonth()+1) + '월 ' + dateObj.getDate() + '일';

    // 대상 회원 수집 (해당 휴강일이 수강 기간에 포함되는 회원만)
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    var dbData = dbSheet.getDataRange().getValues();
    var targets = [];
    var onedayTargets = [];

    for (var i = 1; i < dbData.length; i++) {
      var name = String(dbData[i][1] || '').trim();
      var phone = String(dbData[i][2] || '').replace(/[^0-9]/g, '');
      var mType = String(dbData[i][3] || '');
      var datesG = getSafeString(dbData[i][6]).split(',').map(function(s){return s.trim();}).filter(Boolean);
      if (!name || phone.length < 10 || datesG.length === 0) continue;

      // 이 회원의 수강 기간에 휴강일이 포함되는지
      if (!_isInRange(dateStr, datesG) && datesG.indexOf(dateStr) === -1) continue;

      if (mType === '원데이') {
        onedayTargets.push({ name: name, phone: phone });
      } else {
        targets.push({ name: name, phone: phone });
      }
    }

    // 원데이 회원 → 강사님한테 알림
    if (onedayTargets.length > 0) {
      var instructorPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
      if (instructorPhone) {
        var names = onedayTargets.map(function(t){ return t.name; }).join(', ');
        enqueueSMS(instructorPhone, '최승훈',
          '[ROOT] 원데이 회원 휴강 안내 필요\n' + holidayLabel + ' ' + (reason || '') + '\n대상: ' + names,
          '시스템알림', SMS_STATUS.IMMEDIATE);
      }
    }

    // 정규/특정 회원 → 타이밍별 처리
    if (targets.length === 0) return true;

    if (daysUntil >= 7) {
      // 7일 이상 → 별도 안내 없음, 전날알림에서 자동 포함
      return true;
    }

    // 7일 미만 → 별도 휴강 안내 발송
    var status = (daysUntil <= 1) ? SMS_STATUS.IMMEDIATE : SMS_STATUS.WAITING;

    // Gemini로 휴강 안내 생성
    var stylePrompt = getGeminiPrompt('문자 스타일');
    var holidayPrompt = getGeminiPrompt('휴강안내 지시');
    var isUrgent = daysUntil <= 0;

    var prompt =
      '당신은 요가 스튜디오 Root Vinyasa 문자 담당자입니다.\n' +
      '아래 회원들에게 휴강 안내 문자를 작성해주세요.\n\n' +

      '[문자 구조 — 반드시 이 순서와 형식을 따를 것]\n' +
      '1) [Root Vinyasa]\n' +
      '2) 안녕하세요, {이름}님!\n' +
      '3) (빈 줄)\n' +
      '4) 안내 메시지: 짧은 문장 1~2개. 13~15자 내외 지점에서 \\n 줄바꿈.\n' +
      '   긴급도에 따라 어조 조절 (당일→양해·미안함 필수, 전날→양해 포함, 2~6일→담담하게).\n' +
      '5) (빈 줄)\n' +
      '6) 정보 블록: 각 항목은 반드시 별도 줄. "- 항목: 값" 형식.\n' +
      '   반드시 포함: - 휴강일: {날짜} / - 수업 횟수: 차감 없음\n' +
      '   선택 휴강이면: - 쉬셔도 오셔도 됩니다\n' +
      '7) (빈 줄)\n' +
      '8) 계절감 한 줄 + \\n불편 드려 죄송해요. (긴급 시) 또는 \\n감사해요.\n\n' +

      '[절대 금지]\n' +
      '- 이모지\n' +
      '- "~하세요", "~해보세요" 등 지시형 말투\n' +
      '- 두 항목을 한 줄에 나열 (반드시 줄바꿈 분리)\n\n' +

      '[브랜드 스타일]\n' + (stylePrompt || '1인칭 현재형. 짧은 문장. 이모지 없음. 관찰하고 안내하는 말투.') + '\n\n' +

      '[휴강안내 규칙]\n' + (holidayPrompt || '일반: 수업 쉬고 차감 없음. 선택: 오셔도 쉬셔도 됨. 긴급: 미안함 필수.') + '\n\n' +

      '[상황]\n' +
      '- 휴강일: ' + holidayLabel + '\n' +
      '- 사유: ' + (reason || '없음') + '\n' +
      '- 유형: ' + safeType + (safeType === '선택' ? ' (오셔도 되고, 쉬셔도 됩니다)' : '') + '\n' +
      '- 긴급도: ' + (isUrgent ? '당일 긴급 — 양해와 미안함 표현 필수' : (daysUntil === 1 ? '내일 — 양해 표현 포함' : '2~6일 전')) + '\n' +
      '- 수업 한 번은 차감되지 않음 (그대로 남아있음)\n\n' +

      '[대상 회원]\n' + JSON.stringify(targets.map(function(t){ return t.name; })) + '\n\n' +

      '출력: JSON 배열만. 앞뒤 텍스트 없이.\n' +
      '[{"name":"..","message":".."}]';

    var results = callGeminiJSON(prompt);
    for (var j = 0; j < results.length; j++) {
      var target = targets.filter(function(t){ return t.name === results[j].name; })[0];
      if (target) {
        enqueueSMS(target.phone, target.name, results[j].message, '휴강안내', status);
      }
    }

    return true;
  } catch(ex) {
    Logger.log('[addHoliday] ' + ex);
    return false;
  }
}


// ──────────────────────────────────────────────
// 43. 금요 브리핑 (강사님 전용)
// ──────────────────────────────────────────────
function fridayBriefing() {
  try {
    var settings = getSettings();
    if (String(settings['금요 브리핑 (강사용)']) !== 'ON') return;

    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    if (!dbSheet) return;
    var data = dbSheet.getDataRange().getValues();
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var tomorrow = formatDateSafe(new Date(today.getTime() + 86400000));
    var lines = ['[ROOT 내일 수업 브리핑]'];
    var count = 0;

    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][1] || '').trim();
      var datesG = getSafeString(data[i][6]).split(',').map(function(s){return s.trim();}).filter(Boolean);
      var memo = String(data[i][8] || '');
      if (!name || datesG.indexOf(tomorrow) === -1) continue;

      var attended = datesG.filter(function(d){ return parseSafeDate(d) <= today; }).length + 1;
      var milestone = attended === 1 ? ' (첫 수업)' :
                      attended % 10 === 0 ? ' (' + attended + '번째)' : '';
      var line = '· ' + name + milestone;
      var bodyNote = _extractBodyNote(memo);
      if (bodyNote) line += '\n  주의: ' + bodyNote;
      lines.push(line);
      count++;
    }

    if (count === 0) return;
    lines.splice(1, 0, '총 ' + count + '명\n');

    var instructorPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
    if (instructorPhone) {
      enqueueSMS(instructorPhone, '최승훈', lines.join('\n'), '금요브리핑', SMS_STATUS.IMMEDIATE);
    }
  } catch(ex) { Logger.log('[fridayBriefing] ' + ex); }
}

function _extractBodyNote(memo) {
  if (!memo) return null;
  var kw = ['디스크','허리','무릎','어깨','손목','발목','수술','임신','부상','통증','다쳐'];
  var lines = memo.split('\n');
  for (var i = 0; i < lines.length; i++) {
    for (var j = 0; j < kw.length; j++) {
      if (lines[i].indexOf(kw[j]) !== -1) {
        return lines[i].replace(/\[\d{4}-\d{2}-\d{2}[^\]]*\]:/g, '').trim();
      }
    }
  }
  return null;
}


// ──────────────────────────────────────────────
// 44. 신청 검증 (Gemini) — 기존 유지, 정리
// ──────────────────────────────────────────────
function verifyApplication(payload) {
  try {
    var masked = _maskPhone(payload.fullPhone || payload.phoneMask || '');
    var prompt =
      '요가 스튜디오 수강 신청 검수.\nJSON만 반환.\n\n' +
      '신청자: ' + payload.name + ' / ' + masked + '\n' +
      '패스: ' + payload.option + ' / ' + payload.amount + '원\n' +
      '메모: ' + (payload.memo || '없음') + '\n\n' +
      '금액표: 4주 120000 / 8주 230000 / 12주 330000 / 원데이 35000×인원 / 연장4주 110000 / 연장8주 210000 / 연장12주 300000 / 특별 100000\n\n' +
      '{"amount_match":bool,"phone_valid":bool,"memo_body_flag":bool,"alert_message":str|null}';

    var r = callGeminiJSON(prompt);
    var ss = SpreadsheetApp.openById(SS_ID);
    var logSheet = ss.getSheetByName(SHEET_LOG);
    var hasIssue = !r.amount_match || !r.phone_valid || r.memo_body_flag;

    logSheet.getRange(2, 10).setValue(hasIssue ? '검토필요' : '이상없음');
    logSheet.getRange(2, 11).setValue(r.alert_message || '');

    if (hasIssue && r.alert_message) {
      var instructorPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
      if (instructorPhone) {
        enqueueSMS(instructorPhone, '최승훈',
          '[ROOT 신청 검토]\n' + payload.name + ' / ' + masked + '\n' + payload.option + '\n' + r.alert_message,
          '시스템알림', SMS_STATUS.IMMEDIATE);
      }
    }
    return r;
  } catch(ex) {
    Logger.log('[verifyApplication] ' + ex);
    return null;
  }
}

function _maskPhone(phone) {
  var d = String(phone || '').replace(/[^0-9]/g, '');
  if (d.length < 8) return '010-****-????';
  return d.slice(0, 3) + '-****-' + d.slice(-4);
}


// ──────────────────────────────────────────────
// 45. 트리거 설정 (설정 시트 기반)
// ──────────────────────────────────────────────
function setupAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  var settings = getSettings();
  var sendHour  = Number(settings['문자 발송 시간'] || 9);
  var prepHour  = Number(settings['문자 생성 시간'] || 8); // 발송 1시간 전에 준비
  var briefHour = Number(settings['브리핑 시간'] || 20);

  // 발송 1시간 전 — 대상 회원 목록 준비
  ScriptApp.newTrigger('prepareReminderQueue')
    .timeBased().everyDays(1).atHour(prepHour).create();

// 매일 — 만료후 안부 + DB 서식 업데이트
  ScriptApp.newTrigger('dailyBatch')
    .timeBased().everyDays(1).atHour(prepHour).create();

  // 매분 — 한 명씩 순차 처리
  ScriptApp.newTrigger('processOneReminderMember')
    .timeBased().everyMinutes(1).create();

  // 매일 — 대기 항목 실제 발송
  ScriptApp.newTrigger('processScheduledQueue')
    .timeBased().everyDays(1).atHour(sendHour).nearMinute(15).create();

  // 매일 새벽 3시 — 아카이브 정리
  ScriptApp.newTrigger('archiveQueue')
    .timeBased().everyDays(1).atHour(3).create();

  // 금요일 — 강사 브리핑
  if (String(settings['금요 브리핑 (강사용)']) === 'ON') {
    ScriptApp.newTrigger('fridayBriefing')
      .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(briefHour).create();
  }

  // onEdit — 즉시발송 감지
  ScriptApp.newTrigger('sendSmsOnEdit')
    .forSpreadsheet(SS_ID).onEdit().create();

  // onEdit — 결제 확인 체크박스
  ScriptApp.newTrigger('handleSpreadsheetEdit')
    .forSpreadsheet(SS_ID).onEdit().create();

  Logger.log('✅ 트리거 설정 완료');
  Logger.log('  prepareReminderQueue: 매일 ' + prepHour + '시');
  Logger.log('  processOneReminderMember: 매분');
  Logger.log('  processScheduledQueue: 매일 ' + sendHour + '시');
  Logger.log('  archiveQueue: 매일 3시');
  Logger.log('  fridayBriefing: 금요일 ' + briefHour + '시');
}


// ──────────────────────────────────────────────
// 46. 시트 초기 셋업 (1회 실행)
// ──────────────────────────────────────────────
function setupSMSSheets() {
  var ss = SpreadsheetApp.openById(SS_ID);

  // ── ⚙️ 설정 시트 ──
  var sSheet = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  sSheet.clear();
  sSheet.getRange('B:B').setNumberFormat('@');

  var settingsData = [
    ['항목', '값', '설명'],
    // ── 기본 정보 ──
    ['강사명',           '최승훈',           '문자에 표시되는 이름'],
    ['강사 전화번호',    '01043130150',      '시스템 알림 수신 번호'],
    // ── 수업 시간 ──
    ['수업 시작 시간',   '10:30',            '수업 시작 (HH:MM)'],
    ['수업 종료 시간',   '11:40',            '수업 종료 (HH:MM)'],
    ['입장 시간',        '10:15',            '수업 전 입장 가능 시각 (결제확정 문자에 표시)'],
    ['당일 신청 마감',   '12:40',            '이 시각 이후 당일을 시작일로 신청 불가 (수업 종료 + 1시간 권장)'],
    // ── 스튜디오 정보 (결제확정 문자에 반영) ──
    ['스튜디오 위치',    '서울 용산구',      '날씨·오시는 길 Gemini 조회용 (구 단위)'],
    ['스튜디오 주소',    '서울 용산구 백범로77길 37 3층', '결제확정 문자에 그대로 표시'],
    ['스튜디오명',       '스튜디오쿨라',     '결제확정 문자에 표시'],
    ['역 안내',          '효창공원역 도보 5분', '결제확정 문자에 표시'],
    ['주차 안내',        '주차 전용 없음 · 주차 앱으로 근처 확인 가능', '결제확정 문자에 표시'],
    ['수건 안내',        '땀이 많은 편이시면 수건 지참', '결제확정 문자에 표시 (비워두면 미표시)'],
    // ── 자동화 설정 ──
    ['문자 생성 시간',   '9',    'Gemini가 대기열에 넣는 시각 (0~23)'],
    ['문자 발송 시간',   '9',    '대기열 실제 발송 시각 (0~23)'],
    ['휴강 안내 며칠 전부터', '28', '이 기간 안 휴강은 전날알림에 자동 포함'],
    ['수강 끝난 회원 안부',  'ON',  'ON이면 수강 완료 회원에게 안부 문자'],
    ['안부 며칠째에',    '1, 6, 13', '수강 끝난 후 며칠째에 안부 보낼지 (쉼표 구분)'],
    ['금요 브리핑 (강사용)', 'ON', 'ON이면 금요일 저녁 내일 수업 요약'],
    ['브리핑 시간',      '20',   '금요 브리핑 시각 (0~23)'],
    ['대기열 보관 기간', '30',   '발송 완료 후 며칠 뒤 아카이브로 이동']
  ];

  sSheet.getRange(1, 1, settingsData.length, 3).setValues(settingsData);
  sSheet.getRange(1, 1, 1, 3).setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold');
  sSheet.getRange(2, 1, settingsData.length - 1, 1).setBackground('#f5f5f5').setFontWeight('bold');
  sSheet.getRange(2, 2, settingsData.length - 1, 1).setBackground('#fffde7');
  sSheet.getRange(2, 3, settingsData.length - 1, 1).setBackground('#f5f5f5').setFontColor('#888888');
  sSheet.setColumnWidth(1, 220);
  sSheet.setColumnWidth(2, 280);
  sSheet.setColumnWidth(3, 380);
  sSheet.setFrozenRows(1);

  // 드롭다운: 시간 (0~23) — 생성(행15), 발송(행16), 브리핑(행21)
  var hourList = [];
  for (var h = 0; h <= 23; h++) hourList.push(String(h));
  var hourRule = SpreadsheetApp.newDataValidation().requireValueInList(hourList, true).setAllowInvalid(false).build();
  [15, 16, 21].forEach(function(r) { sSheet.getRange(r, 2).setDataValidation(hourRule); });

  // 드롭다운: ON/OFF — 안부(행18), 브리핑(행20)
  var onOffRule = SpreadsheetApp.newDataValidation().requireValueInList(['ON', 'OFF'], true).setAllowInvalid(false).build();
  [18, 20].forEach(function(r) { sSheet.getRange(r, 2).setDataValidation(onOffRule); });

  // ── ✏️ Gemini 프롬프트 시트 ──
  var pSheet = ss.getSheetByName(SHEET_PROMPTS) || ss.insertSheet(SHEET_PROMPTS);
  pSheet.clear();
  var promptData = [
    ['용도', '프롬프트 (자유 편집)'],
    ['문자 스타일',
      '1인칭, 현재형. 짧은 문장. 여백이 언어의 일부.\n관찰하고 안내하는 말투. 이모지 없음.\n~하세요/~해보세요 (지시) 금지.\n함께해요/공유해주세요 (유도) 금지.\n당신은 소중합니다 (추상 웰니스) 금지.\n판매성·작별성 문구 금지.\n숫자는 한국어로 (1→한 번, 2→두 번, 3→세 번).\n13~15자 내외에서 줄바꿈 — 문자에서 줄이 너무 길면 읽기 어려움.\n항목 여러 개는 반드시 줄바꿈으로 분리 (·나 ,로 한 줄 나열 금지).\n받는 사람이 "나만 받는 문자" 같은 느낌.'],
    ['전날알림 지시',
      '내일 수업 안내. 매주 같은 문자가 아니라 그 주의 상황에 따라 구성이 달라져야 함.\n\n[남은 수업 안내]\n4회 이상: 정보 블록에 담백하게 "- 남은 수업: N번" 한 줄.\n3회: 자연스럽게 녹이거나 정보 블록에 포함.\n2회: 부드럽게 예고. "두 번 남아있어요. 이후 패스는 편하게 말씀해 주세요."\n1회(마지막): 메시지 핵심. "어느덧 마지막 수업이네요." 등. 따뜻한 마무리.\n단, 다음 패스 이미 결제했으면 만료·남은 횟수 관련 문구 전부 제외.\n\n[쉬어가기]\n남아있을 때: 정보 블록에 "- 남은 쉬어가기: N번". 뒤에 "(사정이 생기시면 자정까지\\n문자로 알려주세요.)" 추가.\n3회 이하면 더 강조.\n소진: 쉬어가기 언급 없이.\n\n첫 수업이면 "내일 첫 수업이에요. 기쁜 마음으로 기다리고 있어요." 류의 환영 메시지.\n휴강 예정 있으면 정보 블록에 "- {날짜} 휴강" 한 줄.'],
    ['날씨/오시는 길 지시',
      '수업 시간 기준 일교차, 옷차림, 빙판길, 우산, 미세먼지.\n꽃이 피면 오시는 길에 꽃 구경, 하늘이 맑으면 하늘 보기.\n단순 날씨 보고가 아니라 오시는 길의 경험을 한 줄로.'],
    ['만료후 안부 지시',
      '안부 위주, 판매·재등록 권유 문구 금지.\n거리감 조절: 만료 1일 후=따뜻하고 가깝게, 6일=자연스럽게, 13일=가볍고 담담하게.\n"함께해 주셔서 좋았어요", "다시 오시길 기다릴게요" 등 작별·판매성 문구 금지.\n짧고 담백하게. 계절감 한 줄로 마무리. 재등록 링크는 별도 줄로 자동 첨부됨.'],
    ['휴강안내 지시',
      '일반: 수업이 쉽니다 + 사유 + 수업 한 번은 남아있음.\n선택: 오셔도 되고 쉬셔도 됨 + 쉬시면 차감 없음.\n긴급(당일/전날): 양해와 미안함 표현 필수.\n갑작스러운 안내 드려 죄송하다는 마음.'],
    ['결제확정 (미정)', '⚠️ 추후 확정 — 포함할 내용 정리 후 반영'],
    ['금주 공지', '', '이번 주 전날알림에만 포함할 한 줄 공지.\n예) "다음 주 6/7은 휴강입니다"\n예) "이번 주 매트 세탁으로 개인 매트 지참 부탁드려요"\n없으면 비워두세요. Gemini가 정보 블록에 자연스럽게 녹여줍니다.']
  ];
  pSheet.getRange(1, 1, promptData.length, 2).setValues(promptData);
  pSheet.getRange(1, 1, 1, 2).setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold');
  pSheet.getRange(2, 1, promptData.length - 1, 1).setBackground('#f5f5f5').setFontWeight('bold');
  pSheet.getRange(2, 2, promptData.length - 1, 1).setBackground('#fffde7').setWrap(true);
  pSheet.setColumnWidth(1, 180);
  pSheet.setColumnWidth(2, 600);
  pSheet.setFrozenRows(1);
  pSheet.setRowHeights(2, promptData.length - 1, 120);

  // ── 📨 문자 발송 대기열 ──
  var qSheet = ss.getSheetByName(SHEET_SMS_QUEUE) || ss.insertSheet(SHEET_SMS_QUEUE);
  qSheet.clear();
  var qHeaders = ['번호', '이름', '수신번호', '유형', '문자내용', '상태', '예약시간', '완료시각'];
  qSheet.getRange(1, 1, 1, 8).setValues([qHeaders]);
  qSheet.getRange(1, 1, 1, 8).setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold');
  qSheet.setFrozenRows(1);
  qSheet.setColumnWidths(1, 1, 50);
  qSheet.setColumnWidth(2, 80);
  qSheet.setColumnWidth(3, 120);
  qSheet.setColumnWidth(4, 100);
  qSheet.setColumnWidth(5, 400);
  qSheet.setColumnWidth(6, 80);
  qSheet.setColumnWidth(7, 100);
  qSheet.setColumnWidth(8, 150);
  qSheet.getRange('C:C').setNumberFormat('@');
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['즉시발송', '대기', '승인대기', '처리중', '발송완료', '발송실패'], true)
    .setAllowInvalid(true).build();
  qSheet.getRange('F2:F').setDataValidation(statusRule);

  // ── 📦 대기열 아카이브 ──
  var aSheet = ss.getSheetByName(SHEET_SMS_ARCHIVE) || ss.insertSheet(SHEET_SMS_ARCHIVE);
  aSheet.clear();
  aSheet.getRange(1, 1, 1, 8).setValues([qHeaders]);
  aSheet.getRange(1, 1, 1, 8).setBackground('#455a64').setFontColor('#ffffff').setFontWeight('bold');
  aSheet.setFrozenRows(1);

  Logger.log('✅ SMS 시트 4개 생성/초기화 완료');
}


// ──────────────────────────────────────────────
// 47. 통합 DB 서식 업데이트 (수강중/만료/임박 구분)
//     dailyBatch에서 자동 호출
// ──────────────────────────────────────────────
function updateDBFormatting() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    if (!dbSheet) return;
    var data = dbSheet.getDataRange().getValues();
    var today = new Date(); today.setHours(0, 0, 0, 0);

    for (var i = 1; i < data.length; i++) {
      var datesG = getSafeString(data[i][6]).split(',').map(function(s){return s.trim();}).filter(Boolean);
      if (datesG.length === 0) continue;

      var lastDate = parseSafeDate(datesG[datesG.length - 1]);
      var firstDate = parseSafeDate(datesG[0]);
      if (!lastDate) continue;

      var row = i + 1;
      var remaining = datesG.filter(function(d){ return parseSafeDate(d) >= today; }).length;

      if (lastDate < today) {
        // 만료 — 연회색 + 글자 회색
        dbSheet.getRange(row, 1, 1, 9).setBackground('#f5f5f5').setFontColor('#999999');
      } else if (remaining <= 3 && remaining > 0) {
        // 임박 (3회 이하 남음) — 연노랑
        dbSheet.getRange(row, 1, 1, 9).setBackground('#fff8e1').setFontColor(null);
      } else if (firstDate > today) {
        // 수강 예정 — 연파랑
        dbSheet.getRange(row, 1, 1, 9).setBackground('#e3f2fd').setFontColor(null);
      } else {
        // 수강 중 — 기본
        dbSheet.getRange(row, 1, 1, 9).setBackground(null).setFontColor(null);
      }
    }
  } catch(ex) { Logger.log('[updateDBFormatting] ' + ex); }
}

// ──────────────────────────────────────────────
// 결제 로그 기존 행 색상 일괄 적용 (1회 실행)
// ──────────────────────────────────────────────
function applyLogColors() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) { Logger.log('로그 시트 없음'); return; }
  var data = logSheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var row = i + 1;
    if (!data[i][1]) continue; // 빈 행 스킵
    var checked = data[i][7];
    if (checked === true) {
      logSheet.getRange(row, 1, 1, 9).setBackground('#c8e6c9'); // 연초록
    } else {
      logSheet.getRange(row, 1, 1, 9).setBackground('#fff9c4'); // 연노랑
    }
    count++;
  }
  Logger.log('✅ 로그 색상 적용: ' + count + '행');
}

// ──────────────────────────────────────────────
// 결제 로그 필터 추가 (1회 실행)
// ──────────────────────────────────────────────
function addLogFilter() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) { Logger.log('로그 시트 없음'); return; }
  var existing = logSheet.getFilter();
  if (existing) existing.remove();
  logSheet.getDataRange().createFilter();
  Logger.log('✅ 결제 로그 필터 추가 완료 (H열로 TRUE/FALSE 필터링 가능)');
}

function getSMSStudioInfoMain() {
  var s = getSettings();
  var admitTime  = s['입장 시간']      || '10:15';
  var startTime  = s['수업 시작 시간'] || '10:30';
  var exitTime   = s['퇴실 시간']      || '12:00';
  var address    = s['스튜디오 주소']  || '서울 용산구 백범로77길 37 3층';
  var studioName = s['스튜디오명']     || '스튜디오쿨라';
  var station    = s['역 안내']        || '효창공원역 도보 5분';

  return '\n\n입장 ' + admitTime + ' / 수업 ' + startTime + '\n보통 70분, 길면 ' + exitTime + '까지 이어질 수 있어요.\n\n' +
    studioName + '\n' + address + '\n' + station;
}

function getSMSParkingInfo() {
  var s = getSettings();
  var towel = String(s['수건 안내'] || '');
  var parking = '주차공간이 없어요. 주차 앱으로 근처를 확인해주세요.';
  if (towel) parking += '\n' + '땀이 많은 편이면 수건을 챙겨주세요.';
  return '\n\n' + parking;
}


function fixQueueDisplay() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(SHEET_SMS_QUEUE);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  // 글자색 검정으로 초기화
  sheet.getRange(2, 1, lastRow - 1, 8).setFontColor('#000000');
  Logger.log('✅ 대기열 글자색 복구 완료');
}

// ──────────────────────────────────────────────
// 48. 전날알림 순차 준비 (발송 1시간 전 실행)
// ──────────────────────────────────────────────
function prepareReminderQueue() {
  try {
    // ── 당일 중복 실행 방지 ──
    // setupAllTriggers()로 기존+신규 트리거가 공존하거나
    // 수동 재실행 시 같은 날 두 번 돌아 중복 대기열이 생기는 것을 막음
    var _guardProps = PropertiesService.getScriptProperties();
    var _todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (_guardProps.getProperty('REMINDER_QUEUE_DATE') === _todayKey) {
      Logger.log('[prepareReminderQueue] 오늘 이미 실행됨 (' + _todayKey + '), 스킵');
      return;
    }
    _guardProps.setProperty('REMINDER_QUEUE_DATE', _todayKey);

    var ss = SpreadsheetApp.openById(SS_ID);
    var dbSheet = ss.getSheetByName(SHEET_DB_NEW);
    if (!dbSheet) return;

    var data = dbSheet.getDataRange().getValues();
    var holidayData = loadHolidays(ss);
    var settings = getSettings();
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var tomorrow = new Date(today.getTime() + 86400000);
    var tomorrowStr = formatDateSafe(tomorrow);
    var holidayThreshold = Number(settings['휴강 안내 며칠 전부터'] || 28);

    // 다음 수강권 결제 확인용
    var logSheet = ss.getSheetByName(SHEET_LOG);
    var logData = logSheet ? logSheet.getDataRange().getValues() : [];
    var confirmedNames = {};
    for (var lg = 1; lg < logData.length; lg++) {
      if (logData[lg][7] === true) {
        var lgName = String(logData[lg][1]).trim();
        var lgPhone = String(logData[lg][2]).replace(/[^0-9]/g, '');
        var lgStart = parseSafeDate(String(logData[lg][4] || '').split('~')[0].trim());
        if (lgStart && lgStart >= today) {
          confirmedNames[lgName + '_' + lgPhone] = true;
        }
      }
    }

    var targets = [];
    var seen = {};

    for (var i = 1; i < data.length; i++) {
      var name  = String(data[i][1] || '').trim();
      var phone = String(data[i][2] || '').replace(/[^0-9]/g, '');
      var datesG = getSafeString(data[i][6]).split(',').map(function(s){return s.trim();}).filter(Boolean);
      if (!name || phone.length < 10) continue;
      if (datesG.indexOf(tomorrowStr) === -1) continue;

      var key = name + '_' + phone;
      if (seen[key]) continue;
      seen[key] = true;

      var datesH = getSafeString(data[i][7]).split(',').map(function(s){return s.trim();}).filter(Boolean);
      var totalW = Number(data[i][5] || 0);
      var type   = String(data[i][3] || '');
      var memo   = String(data[i][8] || '');

      // 원데이: 최근 결제확정 받았으면 스킵
      if (type === '원데이') {
        var skip = false;
        for (var lg2 = 1; lg2 < logData.length; lg2++) {
          if (String(logData[lg2][1]).trim() === name && logData[lg2][7] === true) {
            var confirmDate = logData[lg2][0];
            if (confirmDate instanceof Date) {
              var confirmDay = new Date(confirmDate); confirmDay.setHours(0,0,0,0);
              if (Math.floor((today - confirmDay) / 86400000) <= 1) { skip = true; break; }
            }
          }
        }
        if (skip) continue;
      }

      // 예정 휴강 목록
      var upcomingHolidays = [];
      for (var h = 0; h < holidayData.all.length; h++) {
        var hDate = parseSafeDate(holidayData.all[h]);
        if (!hDate) continue;
        var daysAway = Math.floor((hDate - today) / 86400000);
        if (daysAway > 0 && daysAway <= holidayThreshold) {
          if (datesG.indexOf(holidayData.all[h]) !== -1 || _isInRange(holidayData.all[h], datesG)) {
            upcomingHolidays.push({
              date: (hDate.getMonth()+1) + '/' + hDate.getDate(),
              reason: holidayData.reasons[holidayData.all[h]] || '',
              type: holidayData.types[holidayData.all[h]] || '일반',
              daysAway: daysAway
            });
          }
        }
      }

      targets.push({
        name: name,
        phone: phone,
        type: type,
        totalW: totalW,
        datesG: datesG,
        datesH: datesH,
        memo: memo,
        tomorrowStr: tomorrowStr,
        nextPassPurchased: !!confirmedNames[key],
        upcomingHolidays: upcomingHolidays
      });
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty('REMINDER_QUEUE', JSON.stringify(targets));
    props.setProperty('REMINDER_RETRIES', '0');
    // 날씨는 processOneReminderMember 첫 실행 시 조회 (생성 시간에 맞춰 최신 정보 반영)

    // 강사에게 준비 시작 알림
    var instructorPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
    if (instructorPhone && targets.length > 0) {
      enqueueSMS(instructorPhone, '최승훈',
        '[ROOT] 전날알림 준비 시작\n내일 수업 대상 ' + targets.length + '명\n문자 생성을 시작합니다.',
        '시스템알림', SMS_STATUS.IMMEDIATE);
    }

    Logger.log('[prepareReminderQueue] 대상 ' + targets.length + '명 저장완료');
  } catch(ex) {
    Logger.log('[prepareReminderQueue] ' + ex);
  }
}


// ──────────────────────────────────────────────
// 49-A-0. 계절별 감성 한 줄 (폴백·Gemini 공통)
// ──────────────────────────────────────────────
// ──────────────────────────────────────────────
// 날씨 조회 — 2단계 폴백
//   1차: Gemini 검색 → 자연스러운 한 줄 문장
//   2차: wttr.in 무료 API (키 불필요) → _buildWeatherSentence()
//   실패 시 빈 문자열 (buildFallbackReminder에서 계절 문장으로 대체)
// ──────────────────────────────────────────────
function _fetchWeatherNote(tomorrowStr, location) {
  var loc = location || '서울 용산구';
  var dateObj = parseSafeDate(tomorrowStr);
  var dateLabel = dateObj
    ? (dateObj.getMonth()+1) + '월 ' + dateObj.getDate() + '일 토요일'
    : (tomorrowStr || '내일');

  // 1차: Gemini 검색
  try {
    var prompt =
      loc + ' ' + dateLabel + ' 오전 10시 30분 기준 날씨로,' +
      ' 요가 수업에 오는 참석자를 위한 한 줄 날씨 문장을 써주세요.\n' +
      '날씨 상황에 맞게 오시는 길 조언을 자연스럽게 담아주세요.\n' +
      '이모지 없음. 완성된 한 줄 문장만 반환. (설명 없이)\n' +
      '예(맑음): "맑고 상쾌한 아침이 될 것 같아요."\n' +
      '예(비): "비가 내릴 것 같아요. 우산 챙기시면 좋을 것 같아요."\n' +
      '예(쌀쌀): "아침 기온이 낮아요. 따뜻하게 입고 오세요."';
    var result = callGemini(prompt, true, false);
    if (result && result.trim().length > 8) {
      return result.trim().split('\n')[0];
    }
  } catch(ex) { Logger.log('[_fetchWeatherNote] Gemini 실패: ' + ex); }

  // 2차: wttr.in 무료 날씨 API (API 키 불필요)
  try {
    var url = 'https://wttr.in/Yongsan,Seoul?format=j1';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      var w = JSON.parse(resp.getContentText());
      // weather[1] = 내일 예보
      var tmr = w.weather && w.weather[1];
      if (tmr) {
        // 수업 시작 시간(10:30)에 가장 가까운 시간대 선택
        // wttr.in hourly: 0, 300, 600, 900, 1200 (3시간 단위)
        // 900 = 오전 9시 → 수업 직전 시간대로 가장 적합
        var hourly = tmr.hourly || [];
        var am = hourly.filter(function(h){ return Number(h.time) === 900; })[0]
               || hourly.filter(function(h){ return Number(h.time) === 1200; })[0]
               || hourly.filter(function(h){ return Number(h.time) >= 600; })[0]
               || hourly[0];
        var tempC   = Number(am ? am.tempC : tmr.mintempC);
        var rainMM  = Number(am ? (am.precipMM || 0) : 0);
        var descKo  = (am && am.lang_ko && am.lang_ko[0]) ? am.lang_ko[0].value
                    : (am && am.weatherDesc && am.weatherDesc[0] ? am.weatherDesc[0].value : '');
        var sentence = _buildWeatherSentence(tempC, descKo, rainMM);
        Logger.log('[_fetchWeatherNote] wttr.in 성공: ' + sentence);
        return sentence;
      }
    }
  } catch(ex2) { Logger.log('[_fetchWeatherNote] wttr.in 실패: ' + ex2); }

  // 3차: 빈 문자열 → 호출부에서 계절 문장으로 대체
  return '';
}

// 날씨 데이터 → 한국어 한 줄 문장
function _buildWeatherSentence(tempC, descKo, rainMM) {
  var d = String(descKo).toLowerCase();
  var isRain = rainMM > 0.5 || d.indexOf('비') !== -1 || d.indexOf('rain') !== -1 || d.indexOf('drizzle') !== -1 || d.indexOf('가랑') !== -1;
  var isSnow = d.indexOf('snow') !== -1 || d.indexOf('눈') !== -1;
  var isSunny = d.indexOf('맑') !== -1 || d.indexOf('화창') !== -1 || d.indexOf('sunny') !== -1 || d.indexOf('clear') !== -1;
  if (isSnow)              return '눈이 내릴 수 있어요. 미끄럽지 않게 조심해서 오세요.';
  if (isRain && rainMM > 5) return '비가 꽤 내릴 것 같아요. 우산 꼭 챙기세요.';
  if (isRain)              return '비가 살짝 내릴 수 있어요. 우산 챙기시면 좋을 것 같아요.';
  if (tempC < 5)           return '아침 기온이 꽤 낮아요. 따뜻하게 입고 오세요.';
  if (tempC < 10)          return '아침에 살짝 쌀쌀해요. 겉옷 하나 챙기시면 좋아요.';
  if (tempC < 16)          return '선선한 아침이에요. 가볍게 겉옷 하나 걸치고 오시면 딱 좋을 것 같아요.';
  if (tempC < 24 && isSunny) return '맑고 따뜻한 아침이에요. 오시는 길이 기분 좋을 것 같아요.';
  if (tempC < 24)          return '포근한 아침이에요. 편안하게 오실 수 있을 것 같아요.';
  return '아침부터 기온이 꽤 높아요. 시원하게 입고 오시는 걸 추천해요.';
}

// 계절감 끝인사 — 월별 3~4가지, 주차로 순환
// 날씨 정보 없을 때 사용하는 완결형 한 줄 문장
function _getSeasonalClosingVaried(tomorrowStr) {
  var d = parseSafeDate(tomorrowStr) || new Date();
  var m = d.getMonth() + 1;
  var weekSeed = Math.floor((d.getDate() - 1) / 7); // 0~4
  var opts;
  if (m <= 2)        opts = ['겨울 아침, 따뜻하게 챙겨 입고 오세요.', '이 계절의 차가운 공기도 나름의 맛이 있어요.', '손이 시리더라도 발걸음은 가볍게 오세요.'];
  else if (m === 3)  opts = ['봄이 슬그머니 다가오고 있어요. 오시는 길 살짝 쌀쌀하니 겉옷 챙기세요.', '어디선가 꽃 소식이 들려오는 것 같아요. 오시는 길 기대돼요.', '3월 아침은 아직 살짝 쌀쌀해요. 따뜻하게 입고 오세요.'];
  else if (m <= 5)   opts = ['요즘 아침 공기가 참 상쾌해요. 기분 좋게 오실 수 있을 것 같아요.', '날씨가 부쩍 따뜻해졌어요. 가볍게 오시면 딱 좋을 것 같아요.', '봄 하늘이 예쁜 요즘이에요. 오시는 길 하늘 한번 올려다보세요.', '이 계절 아침은 정말 기분이 좋아요. 편안하게 오세요.'];
  else if (m === 6)  opts = ['이른 아침 바람이 아직은 시원해요. 상쾌하게 오실 수 있을 것 같아요.', '초여름 아침 햇살이 기분 좋게 쏟아져요.', '여름의 시작이 느껴지는 요즘이에요. 시원하게 입고 오세요.'];
  else if (m <= 8)   opts = ['날이 많이 더워졌어요. 시원하게 입고 오시는 걸 추천해요.', '무더운 날씨지만 수업 후엔 개운해요. 물 챙겨 오세요.', '아침부터 기온이 꽤 높아요. 가볍게 입고 오세요.'];
  else if (m === 9)  opts = ['선선한 바람이 느껴지는 계절이에요. 오시는 길 기분 좋을 것 같아요.', '가을이 천천히 오고 있어요. 아침 공기 한번 즐기며 오세요.', '9월 아침 공기가 달라졌어요. 오시는 길 가볍게 오세요.'];
  else if (m === 10) opts = ['가을 하늘이 유난히 맑은 요즘이에요. 오시는 길 기분 좋을 것 같아요.', '오는 길에 단풍도 구경하세요. 기분 좋은 아침이 될 것 같아요.', '10월 아침 공기가 참 좋아요. 상쾌하게 오세요.'];
  else if (m === 11) opts = ['단풍이 깊어지는 계절이에요. 따뜻하게 입고 오세요.', '가을 막바지, 아침 기온이 낮아졌어요. 겉옷 챙기세요.', '11월 아침 공기가 제법 차가워요. 따뜻하게 챙겨 입고 오세요.'];
  else               opts = ['연말이 다가오는 요즘이에요. 따뜻하게 입고 오세요.', '한 해가 마무리되어 가는 12월이에요. 오시는 길 따뜻하게 오세요.', '기온이 많이 낮아졌어요. 따뜻하게 챙겨 입고 오세요.'];
  return opts[weekSeed % opts.length];
}


// ──────────────────────────────────────────────
// 49-A. 전날알림 템플릿 폴백 빌더
//       Gemini 실패 시 사용자 스타일 그대로 재현
//
//  구조:
//    [Root Vinyasa]
//    안녕하세요, {이름}님!
//    (blank)
//    {상황별 메인 메시지 — 두 줄 줄바꿈}
//    (blank)
//    {정보 블록 — - 항목: 값}
//    {(사정이 생기시면 자정까지\n문자로 알려주세요.)}
//    (blank)
//    {계절 감성}\n오시는 길 기분 좋게 오세요.
//    (blank)
//    내일 뵙겠습니다!
// ──────────────────────────────────────────────
function buildFallbackReminder(member, memberData) {
  var name      = member.name;
  var remaining = memberData.remaining;
  var isFirst   = memberData.isFirst;
  var pauseLeft = memberData.pauseRemaining;
  var maxPauses = calcMaxPauses(member.type, member.totalW);
  var isOneDay  = member.totalW === 1;
  var holidays  = memberData.holidays || [];

  var timeKor  = '10시 30분';
  var hasPause = !isOneDay && maxPauses > 0 && pauseLeft > 0;
  var isLast   = !isOneDay && remaining === 1;

  var lines = [];

  // ── 헤더 + 인사 ───────────────────────────────────
  lines.push('[Root Vinyasa]');
  lines.push('안녕하세요, ' + name + '님!');

  // ── 수업 시간 안내 (담담하게, 감성 표현 없이) ─────
  if (isFirst) {
    lines.push('내일 오전 ' + timeKor + '\n첫 수업이 예정되어 있어요.');
  } else if (isOneDay) {
    lines.push('내일 오전 ' + timeKor + '\n루트 빈야사 원데이 수업이에요.');
  } else if (isLast) {
    lines.push('내일 오전 ' + timeKor + '\n이번 패스 마지막 수업이에요.');
  } else {
    lines.push('내일 오전 ' + timeKor + '\n루트 빈야사 수업이 예정되어 있어요.');
  }

  // ── 정보 블록 ─────────────────────────────────────
  var infoItems = [];

  // 예정 휴강
  var nearHols = holidays.filter(function(h) { return h.daysAway > 0 && h.daysAway <= 21; });
  if (nearHols.length > 0) {
    var h0 = nearHols[0];
    var rp = h0.reason ? h0.reason.replace(/로 인해|으로 인해/g, '').trim() + '로 ' : '';
    if (h0.type === '선택') {
      infoItems.push('- ' + h0.date + ' ' + rp + '선택 휴강\n  (오셔도, 쉬셔도 됩니다)');
    } else {
      infoItems.push('- ' + h0.date + ' ' + rp + '휴강');
    }
  }

  // 남은 수업 (마지막·첫 수업·원데이 제외)
  if (!isOneDay && !isLast && !isFirst && remaining > 0) {
    infoItems.push('- 남은 수업: ' + toKoreanNum(remaining) + '번');
  }

  // 남은 쉬어가기 (소진이면 생략)
  if (hasPause) {
    infoItems.push('- 남은 쉬어가기: ' + toKoreanNum(pauseLeft) + '번');
  }

  // 이후 패스 등록 링크 (마지막일 때)
  if (isLast) {
    infoItems.push('- 이후 패스 등록\n' + BOOKING_LINK);
  }

  if (infoItems.length > 0) {
    lines.push('');
    lines.push(infoItems.join('\n'));
    // 쉬어가기 있으면 괄호 안내
    // 링크 다음에 오면 빈 줄 한 칸 추가 (기봉님 스타일)
    if (hasPause) {
      var lastItem = infoItems[infoItems.length - 1];
      if (lastItem.indexOf('http') !== -1) lines.push('');
      lines.push('(사정이 생기시면 자정까지\n문자로 알려주세요.)');
    }
  }

  // ── 금주 특별 공지 (별도 단락) ──────────────────
  var specialNotice = getGeminiPrompt('금주 공지').trim();
  if (specialNotice) {
    lines.push('');
    lines.push(specialNotice);
  }

  // ── 날씨/계절 + 마무리 ─────────────────────────
  // 날씨 문장은 그 자체로 완결 — "내일 뵙겠습니다!"는 날씨 바로 다음, 빈 줄 없음
  var weatherClosing = '';
  try { weatherClosing = PropertiesService.getScriptProperties().getProperty('WEATHER_NOTE') || ''; } catch(e) {}
  var closingLine = weatherClosing || _getSeasonalClosingVaried(member.tomorrowStr);
  lines.push('');
  lines.push(closingLine);
  lines.push('내일 뵙겠습니다!');

  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}


// ──────────────────────────────────────────────
// 49. 1분마다 한 명씩 처리 (과부하 방지)
//     한도: 1분에 최대 10회 시도, 실패 시 3회 재시도
//     전체 실패 시 강사에게 문자
// ──────────────────────────────────────────────
function processOneReminderMember() {
  // 동시 실행 방지 — 중복 처리·중복 알림·중복 enqueue 모두 이것으로 해결
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('[processOneReminderMember] 다른 실행 중, 스킵');
    return;
  }

  try {

  var MAX_RETRIES = 3;
  var MAX_PER_MIN = 10;

  var props = PropertiesService.getScriptProperties();
  var queueRaw = props.getProperty('REMINDER_QUEUE');
  if (!queueRaw) return;

  var queue;
  try { queue = JSON.parse(queueRaw); } catch(ex) { return; }
  if (!queue || queue.length === 0) {
    props.deleteProperty('REMINDER_QUEUE');
    props.deleteProperty('REMINDER_RETRIES');
    return;
  }

  var retries = Number(props.getProperty('REMINDER_RETRIES') || '0');

  // 날씨: 오늘 날짜 기준으로 아직 조회 안 했으면 지금 조회 (문자 생성 시각 기준)
  var _todayForWeather = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var weatherNote = props.getProperty('WEATHER_NOTE') || '';
  if (props.getProperty('WEATHER_NOTE_DATE') !== _todayForWeather) {
    var _wSettings = getSettings();
    var _wLocation = _wSettings['스튜디오 위치'] || '서울 용산구';
    var _wTomorrow = new Date(); _wTomorrow.setDate(_wTomorrow.getDate() + 1);
    var _wTomorrowStr = Utilities.formatDate(_wTomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    weatherNote = _fetchWeatherNote(_wTomorrowStr, _wLocation);
    props.setProperty('WEATHER_NOTE', weatherNote);
    props.setProperty('WEATHER_NOTE_DATE', _todayForWeather);
    Logger.log('[processOneReminderMember] 날씨 조회: ' + (weatherNote || '실패'));
  }

  var member = queue[0];
  var stylePrompt    = getGeminiPrompt('문자 스타일');
  var reminderPrompt = getGeminiPrompt('전날알림 지시');
  var specialNotice  = getGeminiPrompt('금주 공지').trim();
  var today = new Date(); today.setHours(0, 0, 0, 0);

  var maxP = calcMaxPauses(member.type, member.totalW);
  var remaining = member.datesG.filter(function(d){ return parseSafeDate(d) >= today; }).length;
  var pauseRemaining = Math.max(0, maxP - member.datesH.length);

  var memberData = {
    name: member.name,
    remaining: remaining,
    totalWeeks: member.totalW,
    pauseRemaining: pauseRemaining,
    isFirst: member.datesG[0] === member.tomorrowStr,
    nextPassPurchased: !!member.nextPassPurchased,
    holidays: member.upcomingHolidays || [],
    note: _extractBodyNote(member.memo) || ''
  };

  var settings = getSettings();
  var classTimeRm = settings['수업 시작 시간'] || '10:30';

  var prompt =
    '당신은 요가 스튜디오 Root Vinyasa 문자 담당자입니다.\n' +
    '아래 회원에게 보낼 내일(' + (member.tomorrowStr || '토요일') + ') 수업 전날알림 문자를 작성해주세요.\n\n' +

    '[문자 구조 — 아래 템플릿을 그대로 따를 것. 빈 줄 절대 생략 금지.]\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '[Root Vinyasa]\n' +
    '안녕하세요, {이름}님!\n' +
    '{수업 안내 — 담담·간결. 13~15자에서 \\n 줄바꿈.\n' +
    ' 일반  → 내일 오전 {시간}\\n루트 빈야사 수업이 예정되어 있어요.\n' +
    ' 첫수업 → 내일 오전 {시간}\\n첫 수업이 예정되어 있어요.\n' +
    ' 마지막 → 내일 오전 {시간}\\n이번 패스 마지막 수업이에요.\n' +
    ' 원데이 → 내일 오전 {시간}\\n루트 빈야사 원데이 수업이에요.}\n' +
    '\n' +
    '{정보 블록 — 각 항목 별도 줄. "- 항목: 값" 형식. 두 항목 한 줄 금지.}\n' +
    '{쉬어가기 남아있으면: (사정이 생기시면 자정까지\\n문자로 알려주세요.)}\n' +
    (specialNotice
      ? '\n' +
        '{공지 단락 — 아래 공지 내용 그대로. 정보 블록 항목(-)이 아닌 별도 단락. 13~15자에서 \\n 줄바꿈.}\n'
      : '') +
    '\n' +
    '{날씨/계절 — 두 줄 이내. 13자에서 \\n. 완결 문장. "오시는 길~" 고정 표현 금지. 매주 다른 구조.}\n' +
    '내일 뵙겠습니다!\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +

    '[절대 금지]\n' +
    '- 이모지\n' +
    '- "기쁜 마음으로", "설레는 마음으로", "오래 기다렸어요" 등 감성 인사 표현 (3번 수업 안내에 절대 금지)\n' +
    '- "함께해 주셔서", "다음 패스도 기다리고 있어요", "기다릴게요" 등 판매성·작별성 문구\n' +
    '- "~하세요", "~해보세요" 등 지시형 말투\n' +
    '- "~해요, ~해요" 반복 어미 연속\n' +
    '- 항목 두 개 이상을 ·나 ,로 한 줄에 나열 (반드시 줄바꿈 분리)\n' +
    '- "쉬어가기 2번", "쉬어가기 한 번(남음)" 처럼 수량 표현 모호하게 쓰기\n' +
    '  → 반드시 "남은 쉬어가기: 두 번" 처럼 "남은"을 앞에 붙여 명확히\n' +
    '- 숫자 숫자 표기 (1→한 번, 2→두 번, 3→세 번, 4→네 번 ...)\n\n' +

    '[브랜드 스타일]\n' +
    (stylePrompt || '1인칭 현재형. 짧은 문장. 이모지 없음. 관찰하고 안내하는 말투.') + '\n\n' +

    '[상황별 작성 규칙]\n' +
    (reminderPrompt || '남은 횟수 자연스럽게. 쉬어가기 안내. 첫 수업이면 환영.') +
    (specialNotice
      ? '\n\n[이번 주 특별 공지 — 반드시 포함]\n' + specialNotice + '\n' +
        '→ 위 문자 구조 8)번 공지 단락 위치에 넣어주세요.\n' +
        '  정보 블록 항목(-)으로 쓰지 말 것. 별도 단락으로. 원문 내용·어투 유지.\n' +
        '  강조·느낌표 금지. 13~15자 내외에서 \\n 줄바꿈.'
      : '') + '\n\n' +

    '[수업 시작 시간] ' + classTimeRm + '\n\n' +
    '[내일 날씨]\n' + (weatherNote || '날씨 정보 없음 — 계절감으로 대체') + '\n\n' +
    '[회원 데이터]\n' + JSON.stringify(memberData, null, 2) + '\n\n' +
    '출력: 아래 JSON만. 앞뒤 텍스트 없이. \\n은 실제 줄바꿈이 아닌 이스케이프 문자로.\n' +
    '{"name":"' + member.name + '","message":"문자내용"}';

  var success = false;
  var callCount = 0;

  while (callCount < MAX_PER_MIN && retries < MAX_RETRIES) {
    try {
      callCount++;
      // forceJSON=true 로 Gemini에 JSON 응답 강제
      var raw = callGemini(prompt, false, true);
      var cleaned = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/,'').trim();
      // JSON 시작 위치 찾기 (앞에 텍스트가 붙었을 때 방어)
      var si = cleaned.search(/[\[{]/);
      if (si > 0) cleaned = cleaned.slice(si);
      var result = JSON.parse(cleaned);
      if (!result.message) throw new Error('message 필드 없음');
      // 마지막 수업이면 신청 링크 첨부
      var finalMsg = result.message;
      if (remaining === 1 && member.totalW !== 1) {
        finalMsg = finalMsg.trim() + '\n\n신청 · ' + BOOKING_LINK;
      }
      enqueueSMS(member.phone, member.name, finalMsg, '전날알림', SMS_STATUS.WAITING);
      success = true;
      break;
    } catch(ex) {
      retries++;
      Logger.log('[processOneReminderMember] 재시도 ' + retries + '/' + MAX_RETRIES + ' — ' + member.name + ' — ' + ex);
      if (retries < MAX_RETRIES) Utilities.sleep(4000);
    }
  }

  if (success) {
    queue.shift();
    props.setProperty('REMINDER_QUEUE', JSON.stringify(queue));
    props.setProperty('REMINDER_RETRIES', '0');
    Logger.log('[processOneReminderMember] 완료: ' + member.name + ' / 남은 대상: ' + queue.length + '명');
    if (queue.length === 0) clearSpecialNotice();  // 마지막 멤버 처리 완료 → 공지 자동 초기화
  } else if (retries >= MAX_RETRIES) {
    // Gemini 3회 실패 → 템플릿 폴백으로 대체 발송
    try {
      var fallbackMsg = buildFallbackReminder(member, memberData);
      enqueueSMS(member.phone, member.name, fallbackMsg, '전날알림(템플릿)', SMS_STATUS.WAITING);
      Logger.log('[processOneReminderMember] 템플릿 폴백 발송: ' + member.name);
    } catch(fbEx) {
      // 템플릿도 실패하면(예외적 상황) 강사에게 알림
      Logger.log('[processOneReminderMember] 템플릿 폴백 실패: ' + fbEx);
      var instructorPhone = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PHONE');
      if (instructorPhone) {
        enqueueSMS(instructorPhone, '최승훈',
          '[ROOT 알림]\n' + member.name + '님 전날알림 생성 실패 (3회 시도).\n수동으로 확인해주세요.',
          '시스템알림', SMS_STATUS.IMMEDIATE);
      }
    }
    queue.shift();
    props.setProperty('REMINDER_QUEUE', JSON.stringify(queue));
    props.setProperty('REMINDER_RETRIES', '0');
    if (queue.length === 0) clearSpecialNotice();  // 마지막 멤버 처리 완료 → 공지 자동 초기화
  } else {
    props.setProperty('REMINDER_RETRIES', String(retries));
  }

  } finally {
    lock.releaseLock();
  }
}


// ──────────────────────────────────────────────
// 50. 금주 공지 자동 초기화
//     전날알림 배치 완료 시 B열(공지 내용)을 자동으로 비워 다음 주 중복 발송 방지
// ──────────────────────────────────────────────
function clearSpecialNotice() {
  try {
    var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_PROMPTS);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === '금주 공지') {
        if (String(data[i][1] || '').trim() === '') return;  // 이미 비어있으면 스킵
        sheet.getRange(i + 1, 2).clearContent();
        Logger.log('[clearSpecialNotice] 금주 공지 초기화 완료 (행 ' + (i + 1) + ')');
        return;
      }
    }
  } catch(e) {
    Logger.log('[clearSpecialNotice] 실패: ' + e);
  }
}

// ──────────────────────────────────────────────
// 51. 금주 공지 행 추가 (1회 실행용 마이그레이션)
//     → Apps Script 편집기에서 addSpecialNoticeRow 실행
// ──────────────────────────────────────────────
function addSpecialNoticeRow() {
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_PROMPTS);
  if (!sheet) { Logger.log('프롬프트 시트 없음'); return; }
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === '금주 공지') {
      Logger.log('이미 존재합니다 (행 ' + (i + 1) + ')');
      return;
    }
  }
  var newRow = [
    '금주 공지',
    '',
    '이번 주 전날알림에만 포함할 한 줄 공지.\n예) "다음 주 6/7은 휴강입니다"\n예) "이번 주 매트 세탁으로 개인 매트 지참 부탁드려요"\n없으면 비워두세요. Gemini가 정보 블록에 자연스럽게 녹여줍니다.'
  ];
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 3).setValues([newRow]);
  sheet.getRange(lastRow, 1).setBackground('#f5f5f5').setFontWeight('bold');
  sheet.getRange(lastRow, 2).setBackground('#fff9c4').setWrap(true);  // 연노랑 강조
  sheet.getRange(lastRow, 3).setBackground('#f5f5f5').setFontColor('#888888').setWrap(true);
  sheet.setRowHeight(lastRow, 100);
  Logger.log('금주 공지 행 추가 완료 (행 ' + lastRow + ')');
}