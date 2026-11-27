import { AUTO_FLAG, DAY_NAME, INSTALL_NAME, AA_VERSION, ENABLED_SETTING, PLANNED_FORMULA, serialAt, stampAt, includesA1, journalPlan, formulaString, latestQueue } from './긴조_기록계산.js';

const HELPERS = [['parsed', '수식원문', 'D2'], ['roster', '수식근무', 'A2'], ['mask', '수식근무', 'P2'], ['ot', '수식OT', 'G2'], ['present', '수식판정', 'P2'], ['missing', '수식명단', 'AB2']];
const LAST_NAME = '긴조_마지막입력';

export function createAutomation({ Excel, Office, notify = () => {}, now = Date.now, delay = ms => new Promise(resolve => setTimeout(resolve, ms)) }) {
  let registration = null, active = false, queue = null, changingMode = false;
  const report = (state, message, detail = '') => notify({ state, message, detail });
  const stale = isCurrent => { if (!isCurrent()) { const error = new Error('새 입력으로 바뀌어 이전 계산을 취소했습니다.'); error.stale = true; throw error; } };
  const putName = (ctx, loaded, name, formula) => loaded && !loaded.isNullObject ? loaded.formula = formula : ctx.workbook.names.add(name, formula);

  async function verifyWorkbook(ctx) {
    const required = [...new Set(['상태출력', '수식설정', '데일리연결', ...HELPERS.map(item => item[1])])];
    const sheets = required.map(name => ({ name, sheet: ctx.workbook.worksheets.getItemOrNullObject(name) }));
    sheets.forEach(item => item.sheet.load('name'));
    await ctx.sync();
    const missing = sheets.find(item => item.sheet.isNullObject);
    if (missing) throw new Error(`${missing.name} 탭이 없습니다. 긴조 수식 서식을 먼저 준비해 주세요.`);
    const config = ctx.workbook.worksheets.getItem('데일리연결');
    const tab = config.getRange('B9').load('values');
    const file = config.getRange('B21').load('values');
    await ctx.sync();
    if (!String(tab.values[0][0]).startsWith('긴급 ') || !String(file.values[0][0]).includes('긴급조치 DailySheet.')) throw new Error('긴조 전용 기능입니다. 인앱·귀책 파일에서는 켜지 않습니다.');
    return ctx.workbook.worksheets.getItem('상태출력');
  }

  async function saveEnabled(enabled) {
    Office.context.document.settings.set(ENABLED_SETTING, enabled);
    await new Promise((resolve, reject) => Office.context.document.settings.saveAsync(result => {
      if (String(result.status).toLowerCase() === 'succeeded') resolve();
      else reject(new Error('자동 시작 설정을 저장하지 못했습니다. 파일의 편집 권한을 확인해 주세요.'));
    }));
  }

  async function snapshot(ctx, isCurrent) {
    const status = ctx.workbook.worksheets.getItem('상태출력');
    const input = status.getRange('A1').load('values');
    const clock = ctx.workbook.worksheets.getItem('수식설정').getRange('B2:B3').load('values');
    const header = status.getRange('AA3').load('values');
    const previous = status.getRange('AA6:AA32').load('values');
    const names = {
      day: ctx.workbook.names.getItemOrNullObject(DAY_NAME),
      install: ctx.workbook.names.getItemOrNullObject(INSTALL_NAME),
      last: ctx.workbook.names.getItemOrNullObject(LAST_NAME),
    };
    Object.values(names).forEach(item => item.load('formula'));
    const ranges = HELPERS.map(([key, sheet, address]) => {
      const anchor = ctx.workbook.worksheets.getItem(sheet).getRange(address).load('formulas');
      return { key, anchor, spill: anchor.getSpillingToRangeOrNullObject().load('address') };
    });
    await ctx.sync(); stale(isCurrent);
    ranges.forEach(item => { item.range = item.spill.isNullObject ? item.anchor : item.spill; item.range.load('values,valueTypes'); });
    await ctx.sync(); stale(isCurrent);
    const result = { raw: input.values[0][0], manualClock: clock.values[0][0], clock: clock.values[1][0], header: String(header.values[0][0] ?? ''), previous: previous.values, dayFormula: names.day.isNullObject ? '' : names.day.formula, installFormula: names.install.isNullObject ? '' : names.install.formula };
    for (const item of ranges) {
      if (!item.anchor.formulas[0][0]) throw new Error('계산 수식이 빠져 있습니다. 기존 수식 설치 상태를 확인해 주세요.');
      if (item.range.valueTypes.some(row => row.some(type => String(type).toLowerCase() === 'error'))) throw new Error('계산 결과에 오류가 있습니다. D5 작성 시각은 저장하고 AA 기록은 유지했습니다.');
      result[item.key] = item.range.values;
    }
    return { result, names, status, input };
  }

  async function processEdit(task, isCurrent) {
    report('working', 'A1 변경 감지 · 계산 확인 중', stampAt(task.at).hhmm);
    await Excel.run(async ctx => {
      stale(isCurrent);
      const status = ctx.workbook.worksheets.getItem('상태출력');
      const rawCell = status.getRange('A1').load('values');
      await ctx.sync(); stale(isCurrent);
      const raw = rawCell.values[0][0];
      if (typeof task.expected === 'string' && task.expected !== raw) return;
      if (typeof raw !== 'string' || !raw.trim()) {
        status.getRange('D5').values = [['']];
        await ctx.sync();
        report('waiting', 'A1이 비어 있습니다', '이전 AA 누적 기록은 유지합니다.');
        return;
      }
      status.getRange('D5').values = [[serialAt(task.at)]];
      status.getRange('D5').numberFormat = [['hh:mm']];
      ctx.application.calculate('Recalculate');
      ctx.application.load('calculationState');
      await ctx.sync(); stale(isCurrent);
      for (let attempt = 0; String(ctx.application.calculationState).toLowerCase() !== 'done'; attempt++) {
        if (attempt >= 12) throw new Error('계산이 아직 끝나지 않았습니다. AA 기록은 유지했습니다. 계산 후 A1 원문을 다시 붙여 주세요.');
        await delay(250); stale(isCurrent);
        ctx.application.load('calculationState'); await ctx.sync();
      }
      const loaded = await snapshot(ctx, isCurrent);
      if (loaded.result.raw !== raw) return;
      const plan = journalPlan(loaded.result, task.at);
      // Re-read A1 immediately before committing to avoid journaling an older paste.
      loaded.input.load('values'); await ctx.sync(); stale(isCurrent);
      if (loaded.input.values[0][0] !== raw) return;
      if (plan.needsInstall) {
        loaded.status.getRange('AA6:AA32').clear('Contents');
        loaded.status.getRange('AA34:AA2003').clear('Contents');
        loaded.status.getRange('AA34').formulas = [[PLANNED_FORMULA]];
      }
      loaded.status.getRange('AA6:AA29').values = plan.rows;
      loaded.status.getRange('AA3:AA5').values = plan.header;
      loaded.status.getRange('AA33').values = [['[데일리 결점]']];
      putName(ctx, loaded.names.day, DAY_NAME, formulaString(plan.stamp.dateKey));
      if (plan.needsInstall) putName(ctx, loaded.names.install, INSTALL_NAME, formulaString(AA_VERSION));
      putName(ctx, loaded.names.last, LAST_NAME, formulaString(plan.stamp.iso));
      await ctx.sync();
      if (isCurrent()) report('on', '자동 기록 중', `${plan.stamp.hhmm} 작성 · ${plan.summary}`);
    });
  }

  function changed(event) {
    if (!active || String(event.source).toLowerCase() === 'remote' || !includesA1(event.address)) return;
    queue.push({ at: now(), expected: event.details?.valueAfter });
  }

  async function start({ persist = true } = {}) {
    if (active || changingMode) return;
    changingMode = true;
    try {
      if (!Office.context.requirements.isSetSupported('ExcelApi', '1.12')) throw new Error('ExcelApi 1.12 이상이 필요합니다. 현재 Excel 웹 계정에서 지원되는지 확인해 주세요.');
      if (!Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) throw new Error('자동 재시작을 위한 SharedRuntime 1.1을 지원하지 않습니다.');
      queue = latestQueue({ process: processEdit, onError(error) { if (!error.stale) report('error', 'AA 자동 기록을 완료하지 못했습니다', error.message || '계산 결과를 확인해 주세요.'); } });
      await Excel.run(async ctx => {
        const status = await verifyWorkbook(ctx);
        registration = status.onChanged.add(changed);
        const flag = ctx.workbook.names.getItemOrNullObject(AUTO_FLAG).load('formula');
        await ctx.sync();
        putName(ctx, flag, AUTO_FLAG, '=TRUE');
        await ctx.sync();
      });
      active = true;
      if (persist) {
        await saveEnabled(true);
        await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
      }
      report('on', '자동 기록 중', '이제 상태출력 A1에 원문을 붙여 주세요. 다른 칸은 감지하지 않습니다.');
    } catch (error) {
      active = false; queue?.stop();
      if (registration) { try { await Excel.run(registration.context, async ctx => { registration.remove(); await ctx.sync(); }); } catch { /* The host may have unloaded the handler. */ } registration = null; }
      try { await Excel.run(async ctx => { const flag = ctx.workbook.names.getItemOrNullObject(AUTO_FLAG).load('formula'); await ctx.sync(); if (!flag.isNullObject) flag.formula = '=FALSE'; await ctx.sync(); }); } catch { /* Surface the original activation error. */ }
      if (persist) { try { await saveEnabled(false); } catch { /* Keep the activation error visible. */ } }
      report('error', '자동 기록을 켜지 못했습니다', error.message || '추가 기능 권한을 확인해 주세요.');
      throw error;
    } finally { changingMode = false; }
  }

  async function stop() {
    if (changingMode) return;
    changingMode = true;
    active = false; queue?.stop();
    try {
      if (registration) { await Excel.run(registration.context, async ctx => { registration.remove(); await ctx.sync(); }); registration = null; }
      await Excel.run(async ctx => { const flag = ctx.workbook.names.getItemOrNullObject(AUTO_FLAG).load('formula'); await ctx.sync(); if (!flag.isNullObject) flag.formula = '=FALSE'; await ctx.sync(); });
      await saveEnabled(false);
      await Office.addin.setStartupBehavior(Office.StartupBehavior.none);
      report('off', '자동 기록 꺼짐', '다시 켜면 다음 A1 변경부터 기록합니다.');
    } catch (error) { report('error', '자동 기록 해제를 확인해 주세요', error.message); throw error; }
    finally { changingMode = false; }
  }
  return { start, stop, changed, isActive: () => active, flushPending: () => queue?.flush() };
}

async function boot() {
  const badge = document.getElementById('badge'), title = document.getElementById('status'), detail = document.getElementById('detail');
  const startButton = document.getElementById('start'), stopButton = document.getElementById('stop');
  function notify(result) {
    badge.dataset.state = result.state;
    badge.textContent = { on: '켜짐', off: '꺼짐', error: '확인 필요', working: '기록 중', waiting: '입력 대기' }[result.state] || '대기';
    title.textContent = result.message; detail.textContent = result.detail;
  }
  if (!window.Office) { notify({ state: 'error', message: 'Excel 안에서 열어 주세요', detail: '설치 안내 페이지에서 추가 기능 설치 파일을 먼저 받으세요.' }); return; }
  const info = await Office.onReady();
  if (info.host !== Office.HostType.Excel) { notify({ state: 'error', message: 'Excel 전용 추가 기능입니다', detail: '긴조 수식 파일에서 추가 기능을 여세요.' }); return; }
  const automation = createAutomation({ Excel: window.Excel, Office: window.Office, notify });
  startButton.disabled = false; stopButton.disabled = false;
  startButton.addEventListener('click', () => { void automation.start().catch(() => {}); });
  stopButton.addEventListener('click', () => { void automation.stop().catch(() => {}); });
  if (Office.context.document.settings.get(ENABLED_SETTING) === true) void automation.start({ persist: false }).catch(() => {});
  else notify({ state: 'off', message: '처음 한 번 자동 기록을 켜 주세요', detail: '켜진 뒤에는 A1에 붙여넣기만 하면 됩니다.' });
}

if (typeof window !== 'undefined' && typeof document !== 'undefined') {
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => { void boot(); });
  else void boot();
}
