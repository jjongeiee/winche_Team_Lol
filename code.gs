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

const AUCTION = {
  SESSION_SHEET: 'AuctionSessions',
  TARGET_SHEET: 'AuctionTargets',
  PARTICIPANT_SHEET: 'AuctionParticipants',
  BID_SHEET: 'AuctionBids'
};

const AUCTION_ROUND = {
  DURATION_MS: 7700
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
    const t = HtmlService.createTemplateFromFile('result');
    t.resultId = (e && e.parameter && e.parameter.id) || '';
    return t.evaluate()
      .setTitle('팀 결과')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 🔹 경매 페이지
  if (page === 'auction') {
    const t = HtmlService.createTemplateFromFile('auction');
    t.resultId = (e && e.parameter && e.parameter.id) || '';
    return t.evaluate()
      .setTitle('경매')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

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

// ====== MATCH DB ======
const MATCH_DB = {
  SHEET_NAME: 'DB',
  HEADER_ROW: 1,
  START_ROW: 2,
  REQUIRED_HEADERS: [
    'match_id',
    '닉네임',
    '픽',
    '진영',
    '포지션',
    '승/패',
    'Kill',
    'Death',
    'Assist',
    'CS',
    '골드',
    '데미지',
    '시야점수',
    '플레이 시간',
    '플레이 시간 변환'
  ]
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

function getTeamResult(resultId) {
  
  return getTeamResultById(resultId);
}

/**
 * DB 탭의 전적 데이터를 헤더 기준으로 읽어서 객체 배열로 반환
 * - 헤더명 기준으로 읽기 때문에 열 순서가 바뀌어도 비교적 안전
 * - 빈 닉네임 행은 건너뜀
 */
function getMatchDbRows() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss?.getSheetByName(MATCH_DB.SHEET_NAME);
    if (!sh) throw new Error(`시트 "${MATCH_DB.SHEET_NAME}"를 찾을 수 없습니다.`);

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < MATCH_DB.START_ROW) {
      return {
        ok: true,
        rows: [],
        meta: { count: 0, sheet: MATCH_DB.SHEET_NAME }
      };
    }

    const headers = sh
      .getRange(MATCH_DB.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0]
      .map(h => _asString(h));

    const headerMap = _makeHeaderMap_(headers);

    // 필수 헤더 검사
    const missing = MATCH_DB.REQUIRED_HEADERS.filter(h => !(h in headerMap));
    if (missing.length) {
      throw new Error(`DB 헤더 누락: ${missing.join(', ')}`);
    }

    const numRows = lastRow - MATCH_DB.START_ROW + 1;
    const values = sh.getRange(MATCH_DB.START_ROW, 1, numRows, lastCol).getValues();

    const rows = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      const nick = _asString(row[headerMap['닉네임']]);
      if (!nick) continue;

      rows.push({
        matchId: _asString(row[headerMap['match_id']]),
        nick,
        champ: _asString(row[headerMap['픽']]),
        side: _asString(row[headerMap['진영']]),
        position: _normalizePos(row[headerMap['포지션']]),
        result: _asString(row[headerMap['승/패']]),
        kill: _asNumber(row[headerMap['Kill']]) ?? 0,
        death: _asNumber(row[headerMap['Death']]) ?? 0,
        assist: _asNumber(row[headerMap['Assist']]) ?? 0,
        cs: _asNumber(row[headerMap['CS']]),
        gold: _asNumber(row[headerMap['골드']]),
        damage: _asNumber(row[headerMap['데미지']]),
        vision: _asNumber(row[headerMap['시야점수']]),
        playTimeRaw: _asString(row[headerMap['플레이 시간']]),
        playTimeText: _asString(row[headerMap['플레이 시간 변환']]),
        _rowNumber: MATCH_DB.START_ROW + i
      });
    }

    return {
      ok: true,
      rows,
      meta: {
        count: rows.length,
        sheet: MATCH_DB.SHEET_NAME
      }
    };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

/**
 * 특정 닉네임의 전적만 가져오기
 * - match_id 최신순 정렬
 */
function getMatchDbRowsByNick(nick) {
  try {
    const targetNick = _asString(nick);
    if (!targetNick) throw new Error('nick이 비어있습니다.');

    const res = getMatchDbRows();
    if (!res.ok) throw new Error(res.error || 'DB 읽기 실패');

    const rows = res.rows
      .filter(r => r.nick === targetNick)
      .sort((a, b) => String(b.matchId).localeCompare(String(a.matchId)));

    return {
      ok: true,
      nick: targetNick,
      rows,
      meta: {
        count: rows.length,
        sheet: MATCH_DB.SHEET_NAME
      }
    };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

/**
 * 특정 닉네임의 우측 패널용 프로필 집계
 * - MMR/티어/주포는 기존 getPlayers() 결과를 우선 사용
 * - 승률/최근10/챔프TOP3/KDA 등은 DB 전적으로 계산
 */
function getPlayerProfileByNick(nick) {
  try {
    const targetNick = _asString(nick);
    if (!targetNick) throw new Error('nick이 비어있습니다.');

    // 1) 기본 플레이어 정보(MMR/티어/주포)
    const playersRes = getPlayers();
    if (!playersRes || !playersRes.ok) {
      throw new Error(playersRes?.error || '플레이어 기본정보 불러오기 실패');
    }

    const player = (playersRes.players || []).find(p => _asString(p.nick) === targetNick) || null;

    // 2) 전적 정보
    const matchRes = getMatchDbRowsByNick(targetNick);
    if (!matchRes || !matchRes.ok) {
      throw new Error(matchRes?.error || '전적 불러오기 실패');
    }

    const rows = Array.isArray(matchRes.rows) ? matchRes.rows : [];
    const games = rows.length;

    // 전적이 하나도 없더라도 기본 프로필은 내려주도록 처리
    const wrAll = _calcWinRate_(rows);
    const wr10 = _calcWinRate_(rows.slice(0, 10));
    const mainPosFromDb = _calcMainPosition_(rows);
    const champs = _calcTopChamps_(rows, 3);
    const kda = _calcAverageKda_(rows);
    const avgVision = _calcAverageVision_(rows);
    const dpm = _calcAverageDpm_(rows);

    const mmr = player?.mmr ?? null;
    const tierName = player?.tier || (mmr !== null ? _tierByMMR(mmr) : '중간계');
    const tierCls = _tierClass_(tierName);

    return {
      ok: true,
      profile: {
        nick: targetNick,
        mainLine: _roleLabel(player?.primary || mainPosFromDb || 'ALL'),
        mmr,
        tierName,
        tierCls,
        wrAll,
        wr10,
        games,
        champs,
        avgKda: kda.avgKda,
        avgKill: kda.avgKill,
        avgDeath: kda.avgDeath,
        avgAssist: kda.avgAssist,
        avgVision,
        dpm
      },
      meta: {
        matchCount: games,
        hasPlayerBase: !!player
      }
    };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

function _makeHeaderMap_(headers) {
  const map = {};
  headers.forEach((name, idx) => {
    const key = _asString(name);
    if (key) map[key] = idx;
  });
  return map;
}

function _calcWinRate_(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return null;

  let wins = 0;
  for (let i = 0; i < list.length; i++) {
    if (_isWin_(list[i]?.result)) wins++;
  }
  return wins / list.length; // 0~1 소수
}

function _isWin_(result) {
  const v = _asString(result);
  return v === '승' || v.toUpperCase() === 'WIN';
}

function _calcMainPosition_(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return 'ALL';

  const counts = {};
  for (let i = 0; i < list.length; i++) {
    const pos = _normalizePos(list[i]?.position);
    if (!pos) continue;
    counts[pos] = (counts[pos] || 0) + 1;
  }

  const ordered = Object.entries(counts)
    .sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1];
      return _positionOrderValue_(a[0]) - _positionOrderValue_(b[0]);
    });

  return ordered.length ? ordered[0][0] : 'ALL';
}

function _calcTopChamps_(rows, limit) {
  const list = Array.isArray(rows) ? rows : [];
  const champMap = new Map();

  for (let i = 0; i < list.length; i++) {
    const row = list[i];
    const champ = _asString(row?.champ);
    if (!champ) continue;

    if (!champMap.has(champ)) {
      champMap.set(champ, { name: champ, wins: 0, games: 0 });
    }

    const item = champMap.get(champ);
    item.games += 1;
    if (_isWin_(row?.result)) item.wins += 1;
  }

  return Array.from(champMap.values())
    .sort((a, b) => {
      if (b.games !== a.games) return b.games - a.games;
      const aWr = a.games ? a.wins / a.games : 0;
      const bWr = b.games ? b.wins / b.games : 0;
      if (bWr !== aWr) return bWr - aWr;
      return a.name.localeCompare(b.name, 'ko');
    })
    .slice(0, limit || 3)
    .map(x => ({
      name: x.name,
      wr: x.games ? x.wins / x.games : null,
      games: x.games
    }));
}

function _calcAverageKda_(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    return {
      avgKill: null,
      avgDeath: null,
      avgAssist: null,
      avgKda: null
    };
  }

  let sumK = 0;
  let sumD = 0;
  let sumA = 0;

  for (let i = 0; i < list.length; i++) {
    sumK += _asNumber(list[i]?.kill) ?? 0;
    sumD += _asNumber(list[i]?.death) ?? 0;
    sumA += _asNumber(list[i]?.assist) ?? 0;
  }

  const n = list.length;
  const avgKill = sumK / n;
  const avgDeath = sumD / n;
  const avgAssist = sumA / n;
  const avgKda = (avgKill + avgAssist) / Math.max(avgDeath, 1);

  return {
    avgKill,
    avgDeath,
    avgAssist,
    avgKda
  };
}

function _calcAverageVision_(rows) {
  const list = Array.isArray(rows) ? rows : [];
  let sum = 0;
  let count = 0;

  for (let i = 0; i < list.length; i++) {
    const v = _asNumber(list[i]?.vision);
    if (v === null) continue; // 빈값은 제외
    sum += v;
    count += 1;
  }

  return count > 0 ? (sum / count) : null;
}

function _calcAverageDpm_(rows) {
  const list = Array.isArray(rows) ? rows : [];
  let sum = 0;
  let count = 0;

  for (let i = 0; i < list.length; i++) {
    const damage = _asNumber(list[i]?.damage);
    if (damage === null) continue; // 데미지 빈값은 제외

    const minutes = _playTimeTextToMinutes_(list[i]?.playTimeText);
    if (!Number.isFinite(minutes) || minutes <= 0) continue;

    sum += (damage / minutes);
    count += 1;
  }

  return count > 0 ? (sum / count) : null;
}

function _playTimeTextToMinutes_(text) {
  const raw = _asString(text);
  if (!raw) return null;

  const m = raw.match(/^(\d+):(\d{1,2})$/);
  if (!m) return null;

  const mm = Number(m[1]);
  const ss = Number(m[2]);
  if (!Number.isFinite(mm) || !Number.isFinite(ss)) return null;

  return mm + (ss / 60);
}

function _tierClass_(tierName) {
  const v = _asString(tierName);
  if (v === '천상계') return 'cel';
  if (v === '지하계') return 'und';
  return 'mid';
}

function _roleLabel(role) {
  const v = _normalizePos(role);
  const map = {
    TOP: '탑',
    JUG: '정글',
    MID: '미드',
    ADC: '원딜',
    SUP: '서포터',
    ALL: '올라운더'
  };
  return map[v] || '올라운더';
}

function _positionOrderValue_(pos) {
  const map = {
    TOP: 1,
    JUG: 2,
    MID: 3,
    ADC: 4,
    SUP: 5,
    ALL: 99
  };
  return map[_normalizePos(pos)] || 999;
}

/* =========================================================
   AUCTION: resultId → 경매 대상 플레이어 목록 생성
   ========================================================= */

function getAuctionPlayersFromResult(resultId) {
  try {
    const res = getTeamResultById(resultId);
    if (!res || !res.ok) {
      throw new Error(res?.error || '팀결과 조회 실패');
    }

    const payload = res.payload;
    if (!payload || !Array.isArray(payload.teams)) {
      throw new Error('팀 데이터가 없습니다.');
    }

    const players = [];

    payload.teams.forEach(team => {
      const teamId = team.id ?? null;
      const slots = team.slots || {};

      Object.keys(slots).forEach(role => {
        const nick = _asString(slots[role]);
        if (!nick) return;

        players.push({
          nick,
          teamId,
          role
        });
      });
    });

    if (!players.length) {
      throw new Error('플레이어 목록이 비어 있습니다.');
    }

    // 랜덤 순서 생성
    const shuffled = _shuffleArray_(players).map((p, i) => ({
      ...p,
      orderNo: i + 1
    }));

    return {
      ok: true,
      resultId,
      players: shuffled,
      meta: {
        count: shuffled.length
      }
    };

  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
}

/* =========================================================
   HELPER: 배열 랜덤 셔플
   ========================================================= */

function _shuffleArray_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = a[i];
    a[i] = a[j];
    a[j] = t;
  }
  return a;
}

function createAuctionSession(resultId) {
  try {

    const playersRes = getAuctionPlayersFromResult(resultId);
    if (!playersRes.ok) throw new Error(playersRes.error);

    const players = playersRes.players;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sessionSh = ss.getSheetByName(AUCTION.SESSION_SHEET);
    const targetSh = ss.getSheetByName(AUCTION.TARGET_SHEET);

    if (!sessionSh) throw new Error('AuctionSessions 시트 없음');
    if (!targetSh) throw new Error('AuctionTargets 시트 없음');

    const auctionId = _newAuctionId_();

    // 세션 저장
    sessionSh.appendRow([
      auctionId,
      resultId,
      'waiting',
      new Date(),
      1
    ]);

    // 플레이어 저장
    const rows = players.map(p => [
      auctionId,
      p.orderNo,
      p.nick,
      p.teamId,
      p.role,
      ''
    ]);

    targetSh.getRange(
      targetSh.getLastRow() + 1,
      1,
      rows.length,
      rows[0].length
    ).setValues(rows);

    return {
      ok: true,
      auctionId,
      playerCount: rows.length
    };

  } catch (err) {
    return { ok:false, error: err.message };
  }
}

function _newAuctionId_() {
  const now = new Date();
  const ts = Utilities.formatDate(now,'Asia/Seoul','yyyyMMdd_HHmmss');
  const rand = Math.random().toString(36).slice(2,5);
  return `auction_${ts}_${rand}`;
}

function startAuctionRound(auctionId, orderNo, roundNo) {

  const endTime = Date.now() + AUCTION_ROUND.DURATION_MS;

  const cache = CacheService.getScriptCache();

  cache.put(
    `auction_round_${auctionId}`,
    JSON.stringify({
      orderNo,
      roundNo,
      endTime
    }),
    30
  );

  return {
    ok:true,
    orderNo,
    roundNo,
    endTime
  };
}

function submitBid(auctionId, bidderCode, amount) {
  try {

    const cache = CacheService.getScriptCache();
    const stateRaw = cache.get(`auction_round_${auctionId}`);
    if (!stateRaw) throw new Error('라운드 없음');

    const state = JSON.parse(stateRaw);

    const bidAmount = Number(amount);
    if (!Number.isFinite(bidAmount) || bidAmount < 0) {
      throw new Error('입찰 금액 오류');
    }

    // 이전 최고가 조회
    const prevMax = _getRoundMaxBid_(auctionId, state.orderNo, state.roundNo);

    if (prevMax !== null && bidAmount < prevMax) {
      throw new Error(`최소 입찰가: ${prevMax}`);
    }

    if (Date.now() > state.endTime) {
      throw new Error('입찰 마감');
    }

    // 🔴 여기 추가: 사용자 포인트 확인
    const userRes = getAuctionUser(auctionId, bidderCode);

    if (!userRes.ok) {
      throw new Error('참가자 등록 안됨');
    }

    if (bidAmount > userRes.points) {
      throw new Error('포인트 부족');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(AUCTION.BID_SHEET);

    sh.appendRow([
      auctionId,
      state.orderNo,
      state.roundNo,
      bidderCode,
      bidAmount,
      new Date()
    ]);

    return { ok:true };

  } catch(err) {
    return { ok:false, error:err.message };
  }
}

function finishAuctionRound(auctionId) {

  const cache = CacheService.getScriptCache();
  const stateRaw = cache.get(`auction_round_${auctionId}`);
  if (!stateRaw) return { ok:false, error:'라운드 없음' };

  const state = JSON.parse(stateRaw);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(AUCTION.BID_SHEET);

  const values = sh.getDataRange().getValues();

  const bids = values
    .filter(r =>
      r[0] === auctionId &&
      r[1] === state.orderNo &&
      r[2] === state.roundNo
    );

  if (!bids.length) {
    return { ok:true, result:'no_bid' };
  }

  const latest = {};

  bids.forEach(b => {
    latest[b[3]] = Number(b[4]);
  });

  const entries = Object.entries(latest);

  entries.sort((a,b)=>b[1]-a[1]);

  const max = entries[0][1];
  const winners = entries.filter(e=>e[1]===max);

  if (winners.length === 1) {

    return {
      ok:true,
      result:'win',
      bidder:winners[0][0],
      amount:max
    };

  }

  return {
    ok:true,
    result:'tie',
    bidders:winners.map(w=>w[0])
  };
}

/* =========================================================
   AUCTION STATE 조회
   ========================================================= */

function getAuctionState(auctionId) {

  try {

    const cache = CacheService.getScriptCache();
    const stateRaw = cache.get(`auction_round_${auctionId}`);

    if (!stateRaw) {
      return {
        ok: true,
        running: false
      };
    }

    const state = JSON.parse(stateRaw);

    const now = Date.now();
    const remain = Math.max(0, state.endTime - now);

    return {
      ok: true,
      running: true,
      orderNo: state.orderNo,
      roundNo: state.roundNo,
      remainMs: remain
    };

  } catch (err) {

    return {
      ok: false,
      error: err.message
    };

  }

}


/* =========================================================
   현재 경매 대상 플레이어 조회
   ========================================================= */

function getAuctionCurrentTarget(auctionId) {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(AUCTION.TARGET_SHEET);

    if (!sh) throw new Error('AuctionTargets 시트 없음');

    const values = sh.getDataRange().getValues();

    const sessionRes = _getAuctionSession_(auctionId);

    if (!sessionRes) throw new Error('세션 없음');

    const orderNo = sessionRes.currentIndex;

    const row = values.find(r =>
      r[0] === auctionId &&
      Number(r[1]) === Number(orderNo)
    );

    if (!row) {
      return { ok:false, error:'플레이어 없음' };
    }

    return {
      ok:true,
      orderNo,
      player: {
        nick: row[2],
        teamId: row[3],
        role: row[4]
      }
    };

  } catch(err) {

    return { ok:false, error:err.message };

  }

}


/* =========================================================
   경매 세션 조회
   ========================================================= */

function _getAuctionSession_(auctionId) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(AUCTION.SESSION_SHEET);

  if (!sh) return null;

  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    if (values[i][0] === auctionId) {

      return {
        row: i + 1,
        resultId: values[i][1],
        status: values[i][2],
        createdAt: values[i][3],
        currentIndex: values[i][4]
      };

    }

  }

  return null;

}

function _getRoundMaxBid_(auctionId, orderNo, roundNo) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(AUCTION.BID_SHEET);

  if (!sh) return null;

  const values = sh.getDataRange().getValues();

  let max = null;

  for (let i = 1; i < values.length; i++) {

    if (
      values[i][0] === auctionId &&
      Number(values[i][1]) === Number(orderNo) &&
      Number(values[i][2]) === Number(roundNo)
    ) {

      const amount = Number(values[i][4]);

      if (max === null || amount > max) {
        max = amount;
      }

    }

  }

  return max;

}

function joinAuction(auctionId, nick) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AuctionUsers");

  if (!sh) throw new Error("AuctionUsers 시트 없음");

  const values = sh.getDataRange().getValues();

  for (let i=1;i<values.length;i++){

    if(values[i][0]===auctionId && values[i][1]===nick){
      return {ok:true};
    }

  }

  sh.appendRow([
    auctionId,
    nick,
    0,
    new Date()
  ]);

  return {ok:true};
}

function getAuctionUser(auctionId,nick){

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName("AuctionUsers");

  const values=sh.getDataRange().getValues();

  for(let i=1;i<values.length;i++){

    if(values[i][0]===auctionId && values[i][1]===nick){

      return{
        ok:true,
        nick:nick,
        points:Number(values[i][2])
      }

    }

  }

  return{ok:false}

}
