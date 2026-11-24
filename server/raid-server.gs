/**
 * 화평 레이드 서버 (Google Apps Script 웹앱)
 * ------------------------------------------------------------
 * - 기존 피드백 서버(feedback_router.gs)와는 완전히 별개입니다. 건드리지 않습니다.
 * - 이 스크립트는 "새 구글 시트"에 붙여 넣습니다. (시트 메뉴 > 확장 프로그램 > Apps Script)
 * - 시트에 탭 3개가 자동으로 만들어집니다: 명단 / 설정 / 공격기록
 *     명단   : A열 이름, B열 반 (예: 홍길동 | 고1-1반)   ← 직접 채우세요. GitHub에 올리지 마세요.
 *     설정   : 주차별 보스 최대 체력 (자동 생성, 필요하면 숫자만 고치세요)
 *     공격기록: 공격이 들어올 때마다 한 줄씩 쌓입니다.
 * - 규칙: (이름 + 주차) 1회만 반영, 보스당 공격력 상한 40, 같은 요청(rid) 중복 무시.
 * - 이 파일에는 비밀번호·키가 없습니다. 웹앱 주소(URL)만 게임 파일의 RAID_URL에 넣으면 됩니다.
 */

var DEFAULT_MAX_HP = 2000;   // 처음 보는 주차의 보스 최대 체력 (설정 탭에서 주차별로 고칠 수 있음)
var CAP_BOSS = 40;           // 한 사람이 보스 하나에 줄 수 있는 최대 공격력
var NAME_MAX = 20;

var SHEET_ROSTER = "명단";
var SHEET_CONFIG = "설정";
var SHEET_LOG    = "공격기록";
var LOG_HEADER = ["시각","주차","이름","반","손절","소문","saved","broken","blocked","spread","wrong","tempo","sec","rid"];

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var action = String(p.action || "status");
    var week = cleanWeek_(p.week);
    if (!week) return json_({ ok: false, error: "week" });

    if (action === "status") return json_(buildStatus_(week, p.name));
    if (action === "hit") return json_(handleHit_(week, p));
    return json_({ ok: false, error: "action" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/* ---------- 공격 ---------- */
function handleHit_(week, p) {
  var name = cleanName_(p.name);
  if (!name) return { ok: false, error: "name" };
  var cut = clampInt_(p.cut, 0, CAP_BOSS);
  var rumor = clampInt_(p.rumor, 0, CAP_BOSS);
  var rid = String(p.rid || "").slice(0, 64);

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSheets_();
    var key = normName_(name);
    var rows = readLog_();
    var dup = false;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r.week === week && (r.key === key || (rid && r.rid === rid))) { dup = true; break; }
    }
    var cls = findClass_(name);
    if (!dup) {
      var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
      sh.appendRow([
        new Date(), week, name, cls, cut, rumor,
        clampInt_(p.saved, 0, 999), clampInt_(p.broken, 0, 999),
        clampInt_(p.blocked, 0, 999), clampInt_(p.spread, 0, 999), clampInt_(p.wrong, 0, 999),
        clampInt_(p.tempo, 0, 100), clampInt_(p.sec, 0, 9999), rid
      ]);
    }
    var st = buildStatus_(week, name);
    st.cls = cls;
    st.applied = dup ? { cut: 0, rumor: 0 } : { cut: cut, rumor: rumor };
    st.duplicate = dup;
    return st;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 상태 ---------- */
function buildStatus_(week, nameOrNull) {
  ensureSheets_();
  var rows = readLog_().filter(function (r) { return r.week === week; });
  var max = getMax_(week);
  var cutSum = 0, rumorSum = 0, cutBy = "", rumorBy = "";
  var byClass = {};
  var played = false;
  var key = nameOrNull ? normName_(cleanName_(nameOrNull)) : "";

  rows.forEach(function (r) {
    cutSum += r.cut; rumorSum += r.rumor;
    // 보스를 쓰러뜨린 마지막 일격의 반 (체력이 처음 0 이하가 된 순간)
    if (!cutBy && cutSum >= max.cut) cutBy = r.cls || "";
    if (!rumorBy && rumorSum >= max.rumor) rumorBy = r.cls || "";
    var c = r.cls || "미확인";
    if (!byClass[c]) byClass[c] = { name: c, damage: 0, players: 0 };
    byClass[c].damage += r.cut + r.rumor;
    byClass[c].players += 1;
    if (key && r.key === key) played = true;
  });

  var classes = Object.keys(byClass).map(function (k) { return byClass[k]; });
  return {
    ok: true,
    bosses: {
      cut:   { hp: Math.max(0, max.cut - cutSum),     max: max.cut,   by: cutBy },
      rumor: { hp: Math.max(0, max.rumor - rumorSum), max: max.rumor, by: rumorBy }
    },
    classes: classes,
    players: rows.length,
    played: played
  };
}

/* ---------- 시트 ---------- */
function ensureSheets_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_ROSTER);
  if (!sh) { sh = ss.insertSheet(SHEET_ROSTER); sh.appendRow(["이름", "반"]); }
  sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) { sh = ss.insertSheet(SHEET_CONFIG); sh.appendRow(["주차", "손절 괴물 최대체력", "소문 괴물 최대체력"]); }
  sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) { sh = ss.insertSheet(SHEET_LOG); sh.appendRow(LOG_HEADER); }
}

function getMax_(week) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === week) {
      return { cut: positive_(v[i][1]), rumor: positive_(v[i][2]) };
    }
  }
  sh.appendRow([week, DEFAULT_MAX_HP, DEFAULT_MAX_HP]);   // 처음 보는 주차는 기본값으로 자동 등록
  return { cut: DEFAULT_MAX_HP, rumor: DEFAULT_MAX_HP };
}

function readLog_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    out.push({
      week: String(r[1]).trim(), key: normName_(String(r[2])), cls: String(r[3] || ""),
      cut: clampInt_(r[4], 0, CAP_BOSS), rumor: clampInt_(r[5], 0, CAP_BOSS), rid: String(r[13] || "")
    });
  }
  return out;
}

function findClass_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  var v = sh.getDataRange().getValues();
  var key = normName_(name);
  for (var i = 1; i < v.length; i++) {
    if (normName_(String(v[i][0])) === key) return String(v[i][1] || "").trim();
  }
  return "";
}

/* ---------- 도우미 ---------- */
function cleanWeek_(w) { w = String(w || "").trim(); return /^\d{4,8}$/.test(w) ? w : ""; }
function cleanName_(n) { return String(n || "").replace(/[\u0000-\u001f]/g, "").trim().slice(0, NAME_MAX); }
function normName_(n) { return String(n || "").replace(/\s+/g, "").toLowerCase(); }
function clampInt_(x, lo, hi) { x = Math.round(Number(x)); if (!isFinite(x)) x = 0; return Math.max(lo, Math.min(hi, x)); }
function positive_(x) { x = Math.round(Number(x)); return x > 0 ? x : DEFAULT_MAX_HP; }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
