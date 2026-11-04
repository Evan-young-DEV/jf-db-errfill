// 메타 지출 raw + 광고비 탭 자동 갱신 — GitHub Actions (PC 꺼져도 작동)
// 어제·오늘 지출만 다시 받아 _메타_<계정>_raw 에 합치고, 11/1~어제 합계를 <상품>_광고비_1102 탭에 기입 (합계>0일 때만)
// 비밀값은 전부 GitHub Secrets(SHEETS, META, GSA). 로그에 토큰·개인정보 출력 금지.
const { google } = require('googleapis');
const SHEETS = JSON.parse(process.env.SHEETS), META = JSON.parse(process.env.META), GSA = JSON.parse(process.env.GSA);
const NOV = SHEETS[0], LOG_TAB = '_오류채움_log', START = 46327; // 11/1
const sh = google.sheets({ version: 'v4', auth: new google.auth.GoogleAuth({ credentials: GSA, scopes: ['https://www.googleapis.com/auth/spreadsheets'] }) });
const kstDate = (off = 0) => new Date(Date.now() + 9 * 3600e3 + off * 864e5).toISOString().slice(0, 10);
const serial = (d) => Math.round((Date.parse(d + 'T00:00:00Z') - Date.UTC(1899, 11, 30)) / 864e5);
const get = async (range) => (await sh.spreadsheets.values.get({ spreadsheetId: NOV, range, valueRenderOption: 'UNFORMATTED_VALUE' })).data.values || [];

async function pull(m, since, until) {
  const rows = [];
  for (const id of m.accounts) {
    const u = new URL(`https://graph.facebook.com/v21.0/act_${id}/insights`);
    Object.entries({ level: 'campaign', time_increment: 1, limit: 500, fields: 'campaign_name,spend,clicks,impressions,actions', time_range: JSON.stringify({ since, until }), access_token: m.token }).forEach(([k, v]) => u.searchParams.set(k, v));
    let url = u;
    while (url) {
      const j = await (await fetch(url)).json();
      if (j.error) throw new Error(`${m.name} ${id} ${j.error.message}`);
      for (const r of j.data || []) {
        const ac = r.actions || [], lead = ac.find((x) => x.action_type === 'lead') || ac.find((x) => x.action_type === 'onsite_conversion.lead_grouped');
        rows.push([serial(r.date_start), r.campaign_name, +r.spend, +r.clicks || 0, +r.impressions || 0, lead ? +lead.value : 0, id]);
      }
      url = j.paging && j.paging.next ? j.paging.next : null;
    }
  }
  return rows;
}

async function once() {
  const titles = (await sh.spreadsheets.get({ spreadsheetId: NOV, fields: 'sheets.properties.title' })).data.sheets.map((s) => s.properties.title);
  const cut = serial(kstDate(-1)), yest = cut, errs = [], sum = {};
  for (const m of META) {
    const raw = `_메타_${m.name}_raw`;
    if (!titles.includes(raw)) continue;
    try {
      const fresh = await pull(m, kstDate(-1), kstDate(0));
      const old = (await get(`'${raw}'!A2:G`));
      if (!old.length) { errs.push(`${raw} 비어 있어 건너뜀`); continue; } // 다른 작업이 지우는 순간이면 과거 기록을 잃지 않게 쓰지 않음
      const all = old.filter((r) => !(Number(r[0]) >= cut)).concat(fresh).sort((a, b) => a[0] - b[0] || b[2] - a[2]);
      const pad = Array.from({ length: Math.max(0, old.length - all.length) }, () => ['', '', '', '', '', '', '']);
      await sh.spreadsheets.values.update({ spreadsheetId: NOV, range: `'${raw}'!A2`, valueInputOption: 'RAW', requestBody: { values: all.concat(pad) } });
      for (const r of all) {
        const d = Number(r[0]), cost = Number(r[2]);
        if (!(d >= START && d <= yest) || !(cost > 0)) continue;
        const tab = `${String(r[1]).split(/[_\s-]/)[0].trim()}_광고비_1102`;
        if (!titles.includes(tab)) continue;
        sum[tab] = sum[tab] || {}; sum[tab][d] = (sum[tab][d] || 0) + cost;
      }
    } catch (e) { errs.push(String(e.message).slice(0, 120)); }
  }
  const data = [];
  for (const [tab, byDay] of Object.entries(sum)) {
    const v = await get(`'${tab}'!A1:B40`);
    for (const [d, amt] of Object.entries(byDay)) {
      const i = v.findIndex((r) => r[0] === Number(d));
      if (i >= 0 && Number(v[i][1] || 0) !== amt) data.push({ range: `'${tab}'!B${i + 1}`, values: [[amt]] });
    }
  }
  if (data.length) await sh.spreadsheets.values.batchUpdate({ spreadsheetId: NOV, requestBody: { valueInputOption: 'RAW', data } });
  const kst = new Date(Date.now() + 9 * 3600e3).toISOString().replace('T', ' ').slice(0, 16);
  if (titles.includes(LOG_TAB)) await sh.spreadsheets.values.update({ spreadsheetId: NOV, range: `'${LOG_TAB}'!N1:O2`, valueInputOption: 'RAW', requestBody: { values: [['메타 지출 자동 마지막 실행(KST)', kst], ['오류', errs.join(' / ')]] } });
  console.log(kst, '광고비 변경', data.length, '오류', errs.length, errs.join(' / '));
}

const LOOP_MIN = Number(process.env.LOOP_MIN || 55), EVERY_SEC = Number(process.env.EVERY_SEC || 300);
(async () => {
  const end = Date.now() + LOOP_MIN * 60e3;
  for (;;) {
    try { await once(); } catch (e) { console.log('ERR', String(e.message).slice(0, 200)); }
    if (process.argv.includes('--once') || Date.now() + EVERY_SEC * 1e3 > end) break;
    await new Promise((r) => setTimeout(r, EVERY_SEC * 1e3));
  }
})();
