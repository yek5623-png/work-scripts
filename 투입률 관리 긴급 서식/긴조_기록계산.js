export const AUTO_FLAG = '긴조_자동기록';
export const DAY_NAME = 'AA누적_운영일';
export const INSTALL_NAME = 'AA누적_설치버전';
export const AA_VERSION = '20261114.1';
export const ENABLED_SETTING = 'ginjo.auto.enabled.v1';
export const LAST_EDIT_SETTING = 'ginjo.auto.lastEdit.v1';
export const PLANNED_FORMULA = '=LET(source,수식근무!A2#,names,CHOOSECOLS(source,3),starts,CHOOSECOLS(source,10),ends,CHOOSECOLS(source,11),reason,CHOOSECOLS(source,12),rows,HSTACK(names,starts,ends,reason),valid,(names<>"")*(starts>=0)*(ends>=0)*(starts<>ends),picked,UNIQUE(FILTER(rows,valid,HSTACK("",-1,-1,""))),MAP(CHOOSECOLS(picked,1),CHOOSECOLS(picked,2),CHOOSECOLS(picked,3),CHOOSECOLS(picked,4),LAMBDA(person,beginMin,endMin,note,IF(person="","",person&" "&TEXT(beginMin/1440,"hh:mm")&"~"&TEXT(endMin/1440,"hh:mm")&IF(note="","","("&note&")")))))';

export const pad = n => String(n).padStart(2, '0');
export const key = value => String(value ?? '').trim().toLowerCase();
export const serialAt = ms => ms / 86400000 + 25569 + 9 / 24;

export function businessTime(serial) {
  if (!Number.isFinite(serial) || serial < 60 || serial > 100000) throw new Error('계산 기준 시각을 읽지 못했습니다.');
  const day = Math.floor(serial);
  const hour24 = Math.min(23, Math.floor((serial - day) * 24 + 1e-7));
  const businessDay = day - (hour24 < 5 ? 1 : 0);
  const date = new Date(Date.UTC(1899, 11, 30) + businessDay * 86400000);
  return { dateKey: date.toISOString().slice(0, 10), dateText: `${date.getUTCMonth() + 1}/${date.getUTCDate()}`, hour: hour24 < 5 ? hour24 + 24 : hour24 };
}

export function stampAt(ms) {
  const date = new Date(ms + 9 * 3600000);
  return { ...businessTime(serialAt(ms)), serial: serialAt(ms), hhmm: `${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}`, iso: date.toISOString().slice(0, 19) + '+09:00' };
}

// Excel change events can contain a rectangle or multiple changed areas.
export function includesA1(address) {
  return String(address ?? '').split(',').some(part => {
    const local = part.slice(part.lastIndexOf('!') + 1).replaceAll('$', '').trim().toUpperCase();
    if (local === 'A:A' || local === '1:1') return true;
    const match = local.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    return !!match && match[1] === 'A' && Number(match[2]) === 1;
  });
}

export function assertSnapshot(snapshot, capturedAt) {
  if (typeof snapshot.raw !== 'string' || !snapshot.raw.trim()) throw new Error('A1 상태 원문이 비어 있습니다.');
  if (snapshot.manualClock !== '' && snapshot.manualClock != null) throw new Error('수식설정 B2에 수기 기준 시각이 있습니다. 자동 기록을 쓰려면 B2를 비워 주세요.');
  const expected = stampAt(capturedAt), actual = businessTime(snapshot.clock);
  if (expected.dateKey !== actual.dateKey || expected.hour !== actual.hour) throw new Error('붙여넣기 시각과 계산 시간대가 다릅니다. 시간대가 바뀌었다면 A1 원문을 다시 붙여 주세요.');
  if (!snapshot.parsed.some(row => key(row[1]))) throw new Error('A1에서 상담사 상태를 읽지 못했습니다. AA 기록은 유지했습니다.');
  if (snapshot.roster.length !== snapshot.mask.length) throw new Error('근무표와 판정 행 수가 다릅니다. AA 기록은 유지했습니다.');
  return expected;
}

export function summarize(snapshot, hour) {
  const scheduled = new Set(snapshot.roster.filter((_, i) => snapshot.mask[i]?.[0] === true || snapshot.mask[i]?.[0] === 1).map(row => key(row[1])).filter(Boolean));
  const ot = new Set(snapshot.ot.map(row => key(row[0])).filter(id => id && !scheduled.has(id)));
  const present = new Set(snapshot.present.map(row => key(row[1])).filter(Boolean));
  const missing = new Set(snapshot.missing.map(row => key(row[0])).filter(Boolean));
  const missingNormal = [...missing].filter(id => !ot.has(id)).length;
  const missingOt = [...ot].filter(id => missing.has(id)).length;
  const presentOt = [...ot].filter(id => present.has(id)).length;
  if ([...ot].some(id => missing.has(id) && present.has(id))) throw new Error('OT 출근·미출근 판정이 충돌합니다. AA 기록은 유지했습니다.');
  const verdict = missingNormal ? `미출근 ${missingNormal}명` : missingOt ? `OT 미출근 ${missingOt}명` : '결점 없음';
  let summary = `(${pad(hour)}~${pad(hour + 1)}) ${verdict}`;
  if (ot.size) summary += ` / OT ${ot.size}명` + (missingOt ? presentOt ? ` (출근 ${presentOt}명, 미출근 ${missingOt}명)` : ` (미출근 ${missingOt}명)` : '');
  return summary;
}

export function formulaString(value) { return `="${String(value).replaceAll('"', '""')}"`; }
export function fromFormula(value) { return String(value ?? '').replace(/^="|"$/g, '').replaceAll('""', '"'); }

export function journalPlan(snapshot, capturedAt) {
  const stamp = assertSnapshot(snapshot, capturedAt);
  const summary = summarize(snapshot, stamp.hour);
  const storedDay = fromFormula(snapshot.dayFormula);
  if (storedDay && !/^\d{4}-\d{2}-\d{2}$/.test(storedDay)) throw new Error('AA 운영일 저장값이 손상되어 이전 기록을 유지했습니다.');
  const installed = fromFormula(snapshot.installFormula) === AA_VERSION;
  const sameDay = storedDay ? storedDay === stamp.dateKey : snapshot.header.startsWith(stamp.dateText + ' SM ');
  const rows = Array.from({ length: 24 }, () => ['']);
  if (sameDay) {
    if (installed) snapshot.previous.slice(0, 24).forEach((row, i) => { rows[i][0] = String(row[0] ?? ''); });
    else for (const row of snapshot.previous) {
      const text = String(row[0] ?? ''), match = text.match(/^\(\s*(\d{1,2})\s*[-~～]\s*(\d{1,2})\s*\)/);
      if (match && +match[1] >= 5 && +match[1] <= 28 && +match[2] === +match[1] + 1) rows[+match[1] - 5][0] = `(${pad(+match[1])}~${pad(+match[2])})` + text.slice(match[0].length);
    }
  }
  rows[stamp.hour - 5][0] = summary;
  return { stamp, summary, rows, needsInstall: !installed, header: [[`${stamp.dateText} SM ${pad(stamp.hour)}시 투입률 공유드립니다.`], [''], ['[긴급조치]']] };
}

// Last write wins. A second paste invalidates any older asynchronous calculation.
export function latestQueue({ process, onError, debounce = 160, timer = setTimeout, cancel = clearTimeout }) {
  let revision = 0, pending = null, running = false, scheduled = null, stopped = false;
  async function drain() {
    if (running || stopped) return;
    running = true;
    try {
      while (pending && !stopped) {
        const task = pending; pending = null;
        const isCurrent = () => !stopped && task.revision === revision;
        try { await process(task, isCurrent); } catch (error) { if (isCurrent()) onError(error); }
      }
    } finally { running = false; }
  }
  return {
    push(task) { if (stopped) return; revision++; pending = { ...task, revision }; if (scheduled) cancel(scheduled); scheduled = timer(() => { scheduled = null; void drain(); }, debounce); },
    flush: drain,
    stop() { stopped = true; revision++; pending = null; if (scheduled) cancel(scheduled); },
  };
}
