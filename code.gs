/**
 * 설정 및 규칙 상수
 */
const CONFIG = {
  SHEET_NAME: '협곡 티어',
  START_ROW: 2,
  COLS: 5, // A(닉), B(MMR), C(주), D(부1), E(부2)
  COL_IDX: {
    NICK: 0,
    MMR: 1,
    PRIMARY: 2,
    SECONDARY1: 3,
    SECONDARY2: 4
  }
};

// ====== TEAM RESULT STORAGE ======
const TEAM_RESULT = {
  SHEET_NAME: '팀결과',
  START_ROW: 2,          // 1행 헤더 가정
  COL_RESULT_ID: 1,      // A
  COL_CREATED_AT: 2,     // B
  COL_PAYLOAD_JSON: 3,   // C
  COL_CREATED_BY: 4,     // D (optional)
  COL_NOTE: 5            // E (optional)
};


// ====== SOURCE SYNC (원본 → 로컬 캐시) ======
const SOURCE = {
  SPREADSHEET_ID: '1392OMwDE1TG-xhxDZAvYfOCe9OTujbjYiRYlBl0br8s',
  RANGE_AJ: '협곡 티어!A:J', // 원본에서 가져올 범위
  // 원본 컬럼 매핑(0-based index)
  // A(0)=닉, D(3)=MMR, I(8)=주, J(9)=부1, 부2는 없어서 빈칸으로 둠
  MAP: { NICK: 0, MMR: 3, PRIMARY: 8, SECONDARY1: 9, SECONDARY2: null }
};


const TIER_RULE = {
  CELESTIAL_MIN: 630,
  UNDERWORLD_MAX: 280
};

// 포지션 매핑 데이터 전역화
const POSITION_MAP = {
  'TOP': 'TOP', '탑': 'TOP',
  'JUG': 'JUG', 'JG': 'JUG', 'JUNGLE': 'JUG', '정글': 'JUG',
  'MID': 'MID', '미드': 'MID',
  'ADC': 'ADC', 'BOT': 'ADC', '원딜': 'ADC', '바텀': 'ADC',
  'SUP': 'SUP', 'SUPPORT': 'SUP', '서폿': 'SUP', '서포터': 'SUP',
  'ALL': 'ALL', 'ALLROUNDER': 'ALL', '올라운더': 'ALL', '올': 'ALL'
};

/**
 * 프론트엔드에서 선수 명단을 가져오는 메인 함수
 */
function getPlayers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss?.getSheetByName(CONFIG.SHEET_NAME);
    if (!sh) throw new Error(`시트 "${CONFIG.SHEET_NAME}"를 찾을 수 없습니다.`);

    const lastRow = sh.getLastRow();
    if (lastRow < CONFIG.START_ROW) {
      return { ok: true, players: [], meta: { count: 0 } };
    }

    const values = sh.getRange(
      CONFIG.START_ROW, 
      1, 
      lastRow - CONFIG.START_ROW + 1, 
      CONFIG.COLS
    ).getValues();

    const seen = new Set();
    const players = values.reduce((acc, row, i) => {
      const nick = _asString(row[CONFIG.COL_IDX.NICK]);
      if (!nick) return acc;

      // 중복 검사
      const nickKey = nick.toLowerCase();
      if (seen.has(nickKey)) {
        throw new Error(`닉네임 중복: "${nick}" (행: ${CONFIG.START_ROW + i})`);
      }
      seen.add(nickKey);

      // 데이터 정제 및 객체 생성
      const mmr = _asNumber(row[CONFIG.COL_IDX.MMR]);
      if (mmr === null) return acc;

      acc.push({
        nick,
        mmr,
        tier: _tierByMMR(mmr),
        primary: _normalizePos(row[CONFIG.COL_IDX.PRIMARY]) || 'ALL',
        secondary: [
          _normalizePos(row[CONFIG.COL_IDX.SECONDARY1]),
          _normalizePos(row[CONFIG.COL_IDX.SECONDARY2])
        ].filter(Boolean)
      });

      return acc;
    }, []);

    return {
      ok: true,
      players,
      meta: { count: players.length, sheet: CONFIG.SHEET_NAME }
    };

  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * 원본 스프레드시트 → 현재 스프레드시트(웹앱) '협곡 티어' 시트로 즉시 동기화
 * - 로컬 시트 A:E (닉, MMR, 주, 부1, 부2)로 값을 "덮어쓰기"
 * - 헤더는 1행 유지, 데이터는 START_ROW부터 갱신
 */
function syncTierFromSource() {
  try {
    const dstSS = SpreadsheetApp.getActiveSpreadsheet();
    const dstSh = dstSS.getSheetByName(CONFIG.SHEET_NAME);
    if (!dstSh) throw new Error(`대상 시트 "${CONFIG.SHEET_NAME}"를 찾을 수 없습니다.`);

    const srcSS = SpreadsheetApp.openById(SOURCE.SPREADSHEET_ID);
    const srcSh = srcSS.getSheetByName('협곡 티어');
    if (!srcSh) throw new Error('원본에서 시트 "협곡 티어"를 찾을 수 없습니다.');

    const srcValues = srcSh.getRange(SOURCE.RANGE_AJ).getValues(); // includes header row
    if (!srcValues || srcValues.length < 2) {
      // 데이터 없음 → 로컬 데이터 영역만 비움
      _clearTierData_(dstSh);
      return { ok: true, wrote: 0, cleared: true };
    }

    // 1행 헤더 제외
    const body = srcValues.slice(1);

    // 유효 행만 변환 (닉네임 있는 행)
    const out = [];
    for (let i = 0; i < body.length; i++) {
      const row = body[i];
      const nick = _asString(row[SOURCE.MAP.NICK]);
      if (!nick) continue;

      const mmr = row[SOURCE.MAP.MMR];
      const primary = row[SOURCE.MAP.PRIMARY];
      const secondary1 = row[SOURCE.MAP.SECONDARY1];
      const secondary2 = (SOURCE.MAP.SECONDARY2 === null) ? '' : row[SOURCE.MAP.SECONDARY2];

      out.push([
        nick,
        _asNumber(mmr),                 // 숫자 변환
        _asString(primary) || 'ALL',
        _asString(secondary1),
        _asString(secondary2)
      ]);
    }

    // 로컬 데이터 영역 클리어 후 쓰기
    _clearTierData_(dstSh);
    if (out.length) {
      dstSh.getRange(CONFIG.START_ROW, 1, out.length, CONFIG.COLS).setValues(out);
    }

    // 값 복사 후 계산 반영(체감용)
    SpreadsheetApp.flush();

    return { ok: true, wrote: out.length, sheet: CONFIG.SHEET_NAME };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

/** 로컬 '협곡 티어' 데이터 영역(A:E, START_ROW~lastRow) 비우기 */
function _clearTierData_(sh) {
  const last = sh.getLastRow();
  if (last < CONFIG.START_ROW) return;
  sh.getRange(CONFIG.START_ROW, 1, last - CONFIG.START_ROW + 1, CONFIG.COLS).clearContent();
}


/* ----------------- Helpers ----------------- */

function _tierByMMR(mmr) {
  if (mmr >= TIER_RULE.CELESTIAL_MIN) return '천상계';
  if (mmr <= TIER_RULE.UNDERWORLD_MAX) return '지하계';
  return '중간계';
}

function _asString(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function _asNumber(v) {
  if (typeof v === 'number' && !isNaN(v)) return v;
  const s = _asString(v);
  const n = Number(s);
  return s && Number.isFinite(n) ? n : null;
}

function _normalizePos(s) {
  const raw = _asString(s);
  if (!raw) return '';
  return POSITION_MAP[raw.toUpperCase()] || POSITION_MAP[raw] || '';
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';

  if (page === 'result') {
    const t = HtmlService.createTemplateFromFile('result'); // 나중에 만들 파일
    t.resultId = (e && e.parameter && e.parameter.id) || '';
    return t.evaluate()
      .setTitle('팀 결과')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 기본 index
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('내전 팀 밸런싱')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveTeamResult(payload, createdBy, note) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(TEAM_RESULT.SHEET_NAME);
    if (!sh) throw new Error(`시트 "${TEAM_RESULT.SHEET_NAME}"를 찾을 수 없습니다.`);

    // payload 검증(최소한 object/array만 허용)
    if (payload === null || payload === undefined) {
      throw new Error('payload가 비어있습니다.');
    }
    if (typeof payload !== 'object') {
      throw new Error('payload는 object 또는 array여야 합니다.');
    }

    const id = _newResultId_();
    const now = new Date();
    const json = JSON.stringify(payload);

    // appendRow는 열 개수 맞추기 위해 5칸으로 통일
    sh.appendRow([
      id,
      now,
      json,
      _asString(createdBy),
      _asString(note)
    ]);

    SpreadsheetApp.flush();

    return {
      ok: true,
      resultId: id,
      url: _buildResultUrl_(id),
      createdAt: now.toISOString()
    };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

function getTeamResultById(resultId) {
  try {
    const id = _asString(resultId);
    if (!id) throw new Error('resultId가 비어있습니다.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(TEAM_RESULT.SHEET_NAME);
    if (!sh) throw new Error(`시트 "${TEAM_RESULT.SHEET_NAME}"를 찾을 수 없습니다.`);

    const lastRow = sh.getLastRow();
    if (lastRow < TEAM_RESULT.START_ROW) throw new Error('저장된 팀 결과가 없습니다.');

    // A:C 정도만 읽으면 충분
    const numRows = lastRow - TEAM_RESULT.START_ROW + 1;
    const values = sh.getRange(TEAM_RESULT.START_ROW, 1, numRows, 3).getValues();

    for (let i = 0; i < values.length; i++) {
      const rowId = _asString(values[i][0]);
      if (rowId === id) {
        const createdAt = values[i][1];
        const payloadJson = _asString(values[i][2]);
        const payload = payloadJson ? JSON.parse(payloadJson) : null;

        return {
          ok: true,
          resultId: id,
          createdAt: createdAt instanceof Date ? createdAt.toISOString() : String(createdAt),
          payload
        };
      }
    }

    return { ok: false, error: '해당 resultId를 찾을 수 없습니다.' };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

function _buildResultUrl_(resultId) {
  const base = ScriptApp.getService().getUrl();
  const sep = base.includes('?') ? '&' : '?';
  return `${base}${sep}page=result&id=${encodeURIComponent(resultId)}`;
}


function _newResultId_() {
  const now = new Date();
  const ts = Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  const rand = Math.random().toString(36).slice(2, 6);
  return `${ts}_${rand}`;
}

/* =============================================
   TEAM FORTUNE (A:상 / B:하 / C:등급 1~3)
   ============================================= */

const FORTUNE2_CONFIG = {
  SHEET_NAME: '팀별운세',
  START_ROW: 2,   // 헤더 있으면 2로
  COLS: { TOP: 1, BOTTOM: 2, GRADE: 3 } // A,B,C
};

function getTeamFortunePool() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss?.getSheetByName(FORTUNE2_CONFIG.SHEET_NAME);
    if (!sh) throw new Error(`시트 "${FORTUNE2_CONFIG.SHEET_NAME}"를 찾을 수 없습니다.`);

    const lastRow = sh.getLastRow();
    if (lastRow < FORTUNE2_CONFIG.START_ROW) {
      return { ok: true, items: [], meta: { count: 0, sheet: FORTUNE2_CONFIG.SHEET_NAME } };
    }

    const numRows = lastRow - FORTUNE2_CONFIG.START_ROW + 1;
    const values = sh.getRange(FORTUNE2_CONFIG.START_ROW, 1, numRows, 3).getValues();

    const seen = new Set(); // 중복 비허용(상+하+등급 기준)
    const items = [];

    for (let i = 0; i < values.length; i++) {
      const top = _asString(values[i][0]);
      const bottom = _asString(values[i][1]);
      const grade = _asNumber(values[i][2]);

      if (!top && !bottom) continue;
      if (![1,2,3].includes(grade)) {
        throw new Error(`등급 오류: "${values[i][2]}" (행: ${FORTUNE2_CONFIG.START_ROW + i})`);
      }

      const key = `${top}||${bottom}||${grade}`.toLowerCase();
      if (seen.has(key)) {
        throw new Error(`운세 중복: "${top} / ${bottom}" (등급 ${grade}) (행: ${FORTUNE2_CONFIG.START_ROW + i})`);
      }
      seen.add(key);

      items.push({ top, bottom, grade });
    }

    return { ok: true, items, meta: { count: items.length, sheet: FORTUNE2_CONFIG.SHEET_NAME } };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function drawTeamFortunesByRules(teamCount) {
  try {
    const n = Number(teamCount);
    if (!Number.isFinite(n) || n <= 0) throw new Error('teamCount가 올바르지 않습니다.');

    const res = getTeamFortunePool();
    if (!res || !res.ok) throw new Error(res?.error || '운세 불러오기 실패');

    // ✅ grade 정규화 + 헤더/불량행 제거
    const normalizeGrade = (g) => {
      const s = String(g ?? '').trim();
      if (!s) return NaN;
      if (s === '등급') return NaN;              // 헤더 방어
      const num = Number(s.replace(/[^\d.-]/g, '')); // "1등급" 같은 경우도 대비
      return Number.isFinite(num) ? num : NaN;
    };

    const normalizeItem = (x) => {
      const grade = normalizeGrade(x?.grade);
      if (![1, 2, 3].includes(grade)) return null;

      // top/bottom 키가 다를 수도 있어서 방어적으로 처리
      const top = (x?.top ?? x?.TOP ?? x?.t ?? '').toString();
      const bottom = (x?.bottom ?? x?.BOTTOM ?? x?.b ?? '').toString();

      return { top, bottom, grade };
    };

    const pool = (Array.isArray(res.items) ? res.items : [])
      .map(normalizeItem)
      .filter(Boolean);

    if (pool.length < n) throw new Error(`운세가 부족합니다: ${pool.length}개 / 필요 ${n}개`);

    // 등급별 버킷
    const g1 = pool.filter(x => x.grade === 1);
    const g2 = pool.filter(x => x.grade === 2);
    const g3 = pool.filter(x => x.grade === 3);

    // 중복 방지용: 선택된 항목을 pool에서 제거
    const picked = [];

    const pickOneFrom = (arr) => {
      if (!arr.length) return null;
      const idx = Math.floor(Math.random() * arr.length);
      return arr.splice(idx, 1)[0];
    };

    const removeFromPool = (item) => {
      const idx = pool.indexOf(item);
      if (idx >= 0) pool.splice(idx, 1);
    };

    // 1) 필수 뽑기
    if (n === 2) {
      // 3등급 1개 필수
      const must3 = pickOneFrom(g3);
      if (!must3) throw new Error('규칙 충족 불가: 3등급 운세가 없습니다.');
      picked.push(must3); removeFromPool(must3);

      // 1 또는 2 등급 1개 필수
      const g12 = g1.concat(g2);
      const must12 = pickOneFrom(g12);
      if (!must12) throw new Error('규칙 충족 불가: 1/2등급 운세가 없습니다.');
      picked.push(must12); removeFromPool(must12);

    } else if (n >= 4) {
      const must1 = pickOneFrom(g1);
      if (!must1) throw new Error('규칙 충족 불가: 1등급 운세가 없습니다.');
      picked.push(must1); removeFromPool(must1);

      const must3 = pickOneFrom(g3);
      if (!must3) throw new Error('규칙 충족 불가: 3등급 운세가 없습니다.');
      picked.push(must3); removeFromPool(must3);
    }

    // 2) 나머지 랜덤 (전체 pool에서 중복 없이)
    const remain = n - picked.length;
    if (pool.length < remain) throw new Error('규칙 적용 후 남은 운세가 부족합니다.');

    // Fisher-Yates shuffle 후 remain개
    for (let i = pool.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const t = pool[i]; pool[i] = pool[j]; pool[j] = t;
    }
    picked.push(...pool.slice(0, remain));

    // (선택) 섞어서 팀에 배정되게 하고 싶으면 한 번 더 shuffle
    for (let i = picked.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const t = picked[i]; picked[i] = picked[j]; picked[j] = t;
    }

    return {
      ok: true,
      fortunes: picked,
      meta: { picked: picked.length, total: (Array.isArray(res.items) ? res.items.length : 0), sheet: FORTUNE2_CONFIG.SHEET_NAME }
    };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

// ✅ (중요) 결과 페이지가 이 함수를 호출할 가능성이 높아서,
// 스텁(TODO)로 두면 링크 페이지가 무조건 실패함.
function getTeamResult(resultId) {
  // 기존에 구현해둔 시트조회 로직 재사용
  return getTeamResultById(resultId);
}



