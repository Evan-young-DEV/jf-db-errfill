// DB 오류줄 자동 채움 — GitHub Actions 5분마다 (PC 꺼져도 작동)
// 캠페인 DB 탭에서 연락처는 있는데 A열 id / ad_name 이 빈(오류) 줄 → 메타 원본 리드에서 같은 연락처·같은 캠페인 DB 찾아 빈 칸만 채움
// 비밀값은 전부 GitHub Secrets(SHEETS, META, GSA). 로그에 개인정보 출력 금지.
const { google } = require('googleapis');
const SHEETS = JSON.parse(process.env.SHEETS), META = JSON.parse(process.env.META), GSA = JSON.parse(process.env.GSA);
const DRY = process.argv.includes('--dry');
const LOG_TAB = '_오류채움_log';
const SKIP_TAB = /광고비|단가|제외|광고변경|보관|raw|계정|승인|log/;
const META_COLS = ['id', 'created_time', 'ad_id', 'ad_name', 'adset_id', 'adset_name', 'campaign_id', 'campaign_name', 'form_id', 'form_name', 'is_organic', 'platform'];
const GV = 'https://graph.facebook.com/v21.0/';
const sh = google.sheets({ version: 'v4', auth: new google.auth.GoogleAuth({ credentials: GSA, scopes: ['https://www.googleapis.com/auth/spreadsheets'] }) });

const phone10 = (s) => { let d = String(s || '').replace(/\D/g, ''); if (d.startsWith('82')) d = '0' + d.slice(2); return /^01\d{8,9}$/.test(d) ? d.slice(-10) : ''; };
const colL = (n) => { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
function fmtTime(t, off) {
  const ms = Date.parse(String(t).replace(/([+-]\d{2})(\d{2})$/, '$1:$2'));
  const sign = off[0] === '-' ? -1 : 1, d = new Date(ms + sign * (Number(off.slice(1, 3)) * 60 + Number(off.slice(4, 6))) * 60e3), p = (n) => String(n).padStart(2, '0');
  return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}T${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}${off}`;
}
async function api(path, params, token) {
  const u = new URL(GV + path); Object.entries(params).forEach(([k, v]) => u.searchParams.set(k, v)); u.searchParams.set('access_token', token);
  const j = await (await fetch(u)).json(); if (j.error) throw new Error(j.error.message); return j;
}
async function getAll(path, params, token) {
  let j = await api(path, params, token), out = j.data || [];
  while (j.paging && j.paging.next) { j = await (await fetch(j.paging.next)).json(); if (j.error) break; out = out.concat(j.data || []); }
  return out;
}

// 살아있음 표시: 숨김탭 _오류채움_log 의 K1:K2 에 마지막 실행 시각(KST)·오류줄 수 — 멈췄는지 시트에서 바로 확인
async function heartbeat(n) {
  if (DRY) return;
  const ss = SHEETS[0];
  const m = (await sh.spreadsheets.get({ spreadsheetId: ss, fields: 'sheets.properties.title' })).data.sheets.map((s) => s.properties.title);
  if (!m.includes(LOG_TAB)) {
    await sh.spreadsheets.batchUpdate({ spreadsheetId: ss, requestBody: { requests: [{ addSheet: { properties: { title: LOG_TAB, hidden: true } } }] } });
    await sh.spreadsheets.values.update({ spreadsheetId: ss, range: `'${LOG_TAB}'!A1`, valueInputOption: 'RAW', requestBody: { values: [['채운 시각(UTC)', '시트', '탭', '행', '연락처', 'id', 'ad_name', 'campaign_name', '채운 칸']] } });
  }
  const kst = new Date(Date.now() + 9 * 3600e3).toISOString().replace('T', ' ').slice(0, 16);
  await sh.spreadsheets.values.update({ spreadsheetId: ss, range: `'${LOG_TAB}'!K1:L2`, valueInputOption: 'RAW', requestBody: { values: [['마지막 실행(KST)', kst], ['그때 오류줄', n]] } });
}

async function findBadRows(simulateGood) {
  const out = [];
  for (const ssId of SHEETS) {
    let tabs;
    try { tabs = (await sh.spreadsheets.get({ spreadsheetId: ssId, fields: 'sheets.properties.title' })).data.sheets.map((s) => s.properties.title).filter((t) => !SKIP_TAB.test(t)); } catch (e) { console.log('시트 열기 실패', e.message); continue; }
    if (!tabs.length) continue;
    const vr = (await sh.spreadsheets.values.batchGet({ spreadsheetId: ssId, ranges: tabs.map((t) => `'${t}'!A1:Z`) })).data.valueRanges;
    vr.forEach((r, i) => {
      const v = r.values || []; if (v.length < 2) return;
      const header = v[0].map((h) => String(h).trim()), tab = tabs[i];
      const pi = header.indexOf('phone_number'), ai = header.indexOf('ad_name'), ti = header.indexOf('created_time'), fi = header.indexOf('form_id'), ci = header.indexOf('campaign_name');
      if (pi < 0 || ai < 0 || ti < 0) return;
      let offset = null; for (let k = v.length - 1; k > 0 && !offset; k--) { const m = String(v[k][ti] || '').match(/([+-]\d{2}:\d{2})$/); if (m) offset = m[1]; }
      v.slice(1).forEach((x, k) => {
        const p = phone10(x[pi]); if (!p) return;
        const a = String(x[0] || '').trim();
        const isBad = !a || !/^l:\d+$/.test(a) || !String(x[ai] || '').trim();
        if (!isBad && !simulateGood) return;
        const forms = {}, camps = {};
        v.slice(1).forEach((y, m) => { if (m === k) return; if (fi >= 0 && y[fi]) forms[String(y[fi]).replace(/^f:/, '')] = 1; if (ci >= 0 && y[ci]) camps[y[ci]] = 1; });
        out.push({ ssId, tab, row: k + 2, header, pi, phone: p, rowVals: x, offset: offset || '-05:00', forms, camps, tabKey: tab.split(/[_\s-]/)[0], isBad });
      });
    });
  }
  return out;
}

async function getLeads(days, allAds) {
  const since = Math.floor(Date.now() / 1000) - days * 86400, map = {}, formName = {};
  const st = allAds ? ['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'ARCHIVED', 'DELETED', 'WITH_ISSUES'] : ['ACTIVE', 'WITH_ISSUES'];
  for (const m of META) for (const act of m.accounts) {
    let ads = [];
    try { ads = await getAll(`act_${act}/ads`, { fields: 'id,name,adset{id,name},campaign{id,name}', limit: 500, filtering: JSON.stringify([{ field: 'effective_status', operator: 'IN', value: st }]) }, m.token); } catch (e) { console.log('광고 조회 실패', m.name, act, e.message); continue; }
    for (const ad of ads) {
      let ls = [];
      try { ls = await getAll(`${ad.id}/leads`, { fields: 'id,created_time,form_id,field_data,is_organic,platform', limit: 500, filtering: JSON.stringify([{ field: 'time_created', operator: 'GREATER_THAN', value: since }]) }, m.token); } catch (e) { continue; }
      for (const l of ls) {
        const ph = (l.field_data || []).find((f) => /phone/i.test(f.name)); const p = phone10(ph && ph.values && ph.values[0]); if (!p) continue;
        if (l.form_id && formName[l.form_id] == null) { try { formName[l.form_id] = (await api(l.form_id, { fields: 'name' }, m.token)).name || ''; } catch (e) { formName[l.form_id] = ''; } }
        (map[p] = map[p] || []).push({ id: l.id, created_time: l.created_time, ad_id: ad.id, ad_name: ad.name, adset_id: ad.adset?.id, adset_name: ad.adset?.name, campaign_id: ad.campaign?.id, campaign_name: ad.campaign?.name, form_id: l.form_id, form_name: formName[l.form_id] || '', is_organic: !!l.is_organic, platform: l.platform || '' });
      }
    }
  }
  return map;
}
// 같은 번호가 여러 캠페인이면: 탭 양식(form_id) → 탭 캠페인명 → 탭 이름 앞부분(상품명), 그래도 여러 개면 최신. 후보 여러 개인데 하나도 안 맞으면 안 채움
function pickLead(list, b) {
  if (!list || !list.length) return null;
  const by = (f) => list.filter(f).sort((x, y) => (x.created_time < y.created_time ? 1 : -1));
  let c = by((l) => b.forms[l.form_id]); if (c.length) return c[0];
  c = by((l) => b.camps[l.campaign_name]); if (c.length) return c[0];
  c = by((l) => String(l.campaign_name || '').split(/[_\s-]/)[0] === b.tabKey); if (c.length) return c[0];
  return list.length === 1 ? list[0] : null;
}

async function once() {
  // --test: 정상 줄을 오류줄로 가정하고 맞게 찾는지 검증(쓰기 없음)
  if (process.argv.includes('--test')) {
    const rows = (await findBadRows(true)).filter((b) => !b.isBad && String(b.rowVals[1] || '') >= new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10));
    const leads = await getLeads(30, true);
    let ok = 0, wrong = 0, none = 0;
    rows.forEach((b) => { const L = pickLead(leads[b.phone], b); if (!L) none++; else if ('l:' + L.id === b.rowVals[0]) ok++; else wrong++; });
    console.log('검증 대상', rows.length, '정확', ok, '틀림', wrong, '못찾음', none);
    return;
  }
  const bad = await findBadRows(false);
  console.log('오류줄', bad.length);
  await heartbeat(bad.length);
  if (!bad.length) return;
  const leads = await getLeads(3, false);
  if (bad.some((b) => !pickLead(leads[b.phone], b))) { // 최근 3일 켜진 광고에서 못 찾으면 30일·전체 광고로 넓게
    const wide = await getLeads(30, true);
    for (const [k, v] of Object.entries(wide)) { const cur = leads[k] || []; leads[k] = cur.concat(v.filter((x) => !cur.some((y) => y.id === x.id))); }
  }
  const data = {}, logs = [];
  for (const b of bad) {
    const L = pickLead(leads[b.phone], b); if (!L) continue;
    const val = { id: 'l:' + L.id, created_time: fmtTime(L.created_time, b.offset), ad_id: 'ag:' + L.ad_id, ad_name: L.ad_name, adset_id: 'as:' + L.adset_id, adset_name: L.adset_name, campaign_id: 'c:' + L.campaign_id, campaign_name: L.campaign_name, form_id: 'f:' + L.form_id, form_name: L.form_name, is_organic: String(L.is_organic), platform: L.platform };
    const filled = [];
    for (const k of META_COLS) {
      const ci = k === 'id' ? 0 : b.header.indexOf(k);
      if (ci < 0 || String(b.rowVals[ci] || '').trim() !== '' || !val[k]) continue;
      (data[b.ssId] = data[b.ssId] || []).push({ range: `'${b.tab}'!${colL(ci)}${b.row}`, values: [[val[k]]] }); filled.push(k);
    }
    if (filled.length) logs.push([new Date().toISOString(), b.ssId === SHEETS[0] ? '11월' : '기타', b.tab, b.row, b.rowVals[b.pi], val.id, val.ad_name, val.campaign_name, filled.join(',')]);
  }
  if (DRY) { console.log('DRY 채울 줄', logs.length); return; }
  for (const [ssId, d] of Object.entries(data)) await sh.spreadsheets.values.batchUpdate({ spreadsheetId: ssId, requestBody: { valueInputOption: 'RAW', data: d } });
  if (logs.length) {
    const ss = SHEETS[0], m = (await sh.spreadsheets.get({ spreadsheetId: ss, fields: 'sheets.properties.title' })).data.sheets.map((s) => s.properties.title);
    if (!m.includes(LOG_TAB)) {
      await sh.spreadsheets.batchUpdate({ spreadsheetId: ss, requestBody: { requests: [{ addSheet: { properties: { title: LOG_TAB, hidden: true } } }] } });
      await sh.spreadsheets.values.update({ spreadsheetId: ss, range: `'${LOG_TAB}'!A1`, valueInputOption: 'RAW', requestBody: { values: [['채운 시각(UTC)', '시트', '탭', '행', '연락처', 'id', 'ad_name', 'campaign_name', '채운 칸']] } });
    }
    await sh.spreadsheets.values.append({ spreadsheetId: ss, range: `'${LOG_TAB}'!A1`, valueInputOption: 'RAW', requestBody: { values: logs } });
  }
  console.log('채운 줄', logs.length, '/ 못 찾은 줄', bad.length - logs.length);
}

// 한 번 실행 = 55분 동안 2분마다 반복 (GitHub 예약 실행 지연과 상관없이 끊김 없이 확인). --once 면 1회만
const LOOP_MIN = Number(process.env.LOOP_MIN || 55), EVERY_SEC = Number(process.env.EVERY_SEC || 120);
(async () => {
  const end = Date.now() + LOOP_MIN * 60e3;
  for (;;) {
    try { await once(); } catch (e) { console.log('ERR', e.message); }
    if (process.argv.includes('--once') || process.argv.includes('--test') || Date.now() + EVERY_SEC * 1e3 > end) break;
    await new Promise((r) => setTimeout(r, EVERY_SEC * 1e3));
  }
})();
