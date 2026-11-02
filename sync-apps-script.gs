// 시드니 가족여행 앱 - 가족 공유 동기화 창구
// 이 스크립트는 이 스프레드시트에만 데이터를 읽고 씁니다.
// 가족 코드는 앱에서 처음 연결할 때 정한 값이 저장되며, 이후 같은 코드를 입력한 기기만 읽고 쓸 수 있습니다.
// 가족 코드를 바꾸려면: 프로젝트 설정 > 스크립트 속성에서 FAMILY_CODE를 지우고 앱에서 다시 연결하세요.

var CAT = {flight:'항공', hotel:'숙소', tour:'투어·입장', transport:'교통', food:'외식', grocery:'장보기', show:'오페라하우스 투어', shopping:'쇼핑', etc:'예비·기타'};

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  var props = PropertiesService.getScriptProperties();
  var code = props.getProperty('FAMILY_CODE');
  var given = String(p.code || '').trim();
  if (!given || given.length < 4) return out_({ok: false, error: 'code_required'});
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!code) { props.setProperty('FAMILY_CODE', given); code = given; }
    if (given !== code) return out_({ok: false, error: 'wrong_code'});
    var st = load_();
    if (p.action === 'sync' && p.state) {
      st = merge_(st, p.state);
      save_(st);
      writeTable_(st);
    }
    return out_({ok: true, state: st, serverTime: Date.now()});
  } finally {
    lock.releaseLock();
  }
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function empty_() { return {expenses: {}, budget: null, checks: {}, rate: null}; }

function load_() {
  var sh = sheet_('_data');
  var n = sh.getLastRow();
  if (n < 1) return empty_();
  var s = sh.getRange(1, 1, n, 1).getValues().map(function (r) { return r[0]; }).join('');
  try { var o = JSON.parse(s); return o && o.expenses ? o : empty_(); } catch (e) { return empty_(); }
}

function save_(st) {
  var sh = sheet_('_data');
  var s = JSON.stringify(st);
  var chunks = [];
  for (var i = 0; i < s.length; i += 40000) chunks.push([s.slice(i, i + 40000)]);
  sh.clearContents();
  if (chunks.length) sh.getRange(1, 1, chunks.length, 1).setNumberFormat('@').setValues(chunks);
}

function newer_(a, b) {
  if (!a) return b || null;
  if (!b) return a;
  return (Number(b.u) || 0) > (Number(a.u) || 0) ? b : a;
}

function keys_(a, b) {
  var k = {};
  Object.keys(a || {}).forEach(function (x) { k[x] = 1; });
  Object.keys(b || {}).forEach(function (x) { k[x] = 1; });
  return Object.keys(k);
}

function merge_(a, b) {
  var r = empty_();
  keys_(a.expenses, b.expenses).forEach(function (id) { r.expenses[id] = newer_((a.expenses || {})[id], (b.expenses || {})[id]); });
  keys_(a.checks, b.checks).forEach(function (id) { r.checks[id] = newer_((a.checks || {})[id], (b.checks || {})[id]); });
  r.budget = newer_(a.budget, b.budget);
  r.rate = newer_(a.rate, b.rate);
  return r;
}

function writeTable_(st) {
  var sh = sheet_('지출 내역');
  var rows = [['날짜', '결제한 날', '카테고리', '내용', '금액', '통화', '호주달러 환산', '결제', '입력 시각']];
  var list = Object.keys(st.expenses).map(function (k) { return st.expenses[k]; })
    .filter(function (e) { return e && !e.del; })
    .sort(function (x, y) {
      var a = x.date === 'pre' ? '0' + (x.pdate || '') : x.date;
      var b = y.date === 'pre' ? '0' + (y.pdate || '') : y.date;
      return String(a).localeCompare(String(b));
    });
  list.forEach(function (e) {
    rows.push([e.date === 'pre' ? '출발 전' : e.date, e.pdate || '', CAT[e.cat] || e.cat, e.desc || '', Number(e.amt) || 0, e.cur === 'KRW' ? '원' : 'A$', Number(e.aud) || 0, e.pay === 'cash' ? '현금' : '카드', e.at || '']);
  });
  sh.clearContents();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
}
