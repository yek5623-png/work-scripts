// [인앱] 데일리 읽기 스크립트 (인앱귀책 투입률)
// 데일리 원본을 수정하지 않고 읽기만 합니다.
// 반환된 result를 Power Automate에서 데일리.json에 저장합니다.
// 20261121: 속도 개선 — 시트마다 큰 범위 2번만 읽습니다 (예전엔 시트마다 7~15번).
//  - 모드 "" (기본): 이번 달 파일에서 오늘 데일리 + 내일 시트 전체 + 향후 9일 예상을 읽습니다.
//    다음달예상 에 다음 달 파일에서 읽은 결과를 넣으면, 이번 달 파일에 없는 날짜(시트 없음)를 그것으로 채웁니다.
//  - 모드 "예상만": 다음 달 파일에 쓰는 모드입니다. 오늘 시트가 없어도 오류 없이,
//    이 파일에 있는 향후 9일 날짜의 시간대 목표·OT 명단·명단만 돌려줍니다.
// ★ 이 데일리 엑셀의 시트 이름 앞부분입니다. 예: "인앱전담 11(23)" 시트면 "인앱전담"
const 시트머리말 = "인앱전담";
// 위 이름의 시트가 없으면 예전 이름("긴급 11(23)")도 찾아봅니다. 필요 없으면 ""로 비워 두세요.
const 예전머리말 = "긴급";

function main(
  workbook: ExcelScript.Workbook,
  기준시각?: string,
  모드?: string,
  다음달예상?: string
): string {
  // 빈칸 대신 "-" 를 넣어도 빈칸으로 봅니다.
  const blank = (v?: string) => { const s = (v || "").trim(); return s === "-" || s === "없음" ? "" : s; };
  기준시각 = blank(기준시각);
  모드 = blank(모드);
  다음달예상 = blank(다음달예상);
  const at = businessClock(기준시각 || new Date().toISOString());

  if (모드 === "예상만") {
    const r = readForecast(workbook, at.date);
    return JSON.stringify({
      type: "ginjo-forecast-only",
      version: 1,
      date: at.date,
      forecast: r.forecast.filter(item => item.error !== "시트 없음"),
      next: r.next
    });
  }

  const found = findSheet(workbook, at.date);
  const name = found.name;
  const sheet = found.sheet;

  if (!sheet) {
    throw new Error("원본에 " + sheetName(at.date) + " 시트가 없습니다. 원본 월 파일 선택을 확인하세요.");
  }

  const today = readBlocks(sheet);
  if (today.bad.length > 0) {
    throw new Error(name + "!" + today.bad.join(", ") + "에 엑셀 오류가 있습니다.");
  }
  if (!today.blocks[0].values.some(row => String(row[0]).trim())) {
    throw new Error("데일리 사번(C열)이 비어 있습니다.");
  }

  // 운영일 기준 +1~+9일 (내일은 시트 전체도 같이)
  const r = readForecast(workbook, at.date);
  const forecast = r.forecast;
  let next: FullDay | undefined = r.next;

  // 이번 달 파일에 없는 날짜는 다음 달 파일 결과로 채웁니다.
  if (다음달예상) {
    try {
      const extraText = 다음달예상.charCodeAt(0) === 65279 ? 다음달예상.slice(1) : 다음달예상;
      const extra = JSON.parse(extraText.trim()) as { forecast?: ForecastItem[]; next?: FullDay };
      for (const item of extra.forecast || []) {
        const i = forecast.findIndex(f => f.date === item.date);
        if (i >= 0 && forecast[i].error === "시트 없음" && !item.error) forecast[i] = item;
      }
      if (!next && extra.next && extra.next.date === addDays(at.date, 1)) next = extra.next;
    } catch (error) {
      // 다음달예상을 읽지 못하면 건너뜁니다.
    }
  }

  return JSON.stringify({
    type: "ginjo-daily-blocks",
    version: 1,
    date: at.date,
    capturedAt: new Date().toISOString(),
    sourceSheet: name,
    blocks: today.blocks,
    forecast: forecast,
    next: next
  });
}

type Cell = string | number | boolean;

interface ForecastItem {
  date: string;
  sheetName: string;
  address?: string;
  values?: Cell[][];
  otValues?: Cell[][];
  staff?: Cell[][];
  error?: string;
}

interface FullDay {
  date: string;
  sourceSheet: string;
  blocks: { address: string; values: Cell[][] }[];
}

// 엑셀 오류 값(#N/A 등)인지
function isErr(v: Cell): boolean {
  if (typeof v !== "string" || v.charAt(0) !== "#") return false;
  const errs = ["#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!"];
  return errs.some(e => v.indexOf(e) === 0);
}
function cols(rows: Cell[][], from: number, to: number): Cell[][] {
  return rows.map(row => row.slice(from, to + 1));
}
function hasErr(rows: Cell[][]): boolean {
  return rows.some(row => row.some(v => isErr(v)));
}

// 한 시트를 2번만 읽어서 예전과 같은 7개 블록으로 나눕니다.
// C6:R140 → C(0) E:G(2~4) I:K(6~8) L:M(9~10) R(15)
// AY6:CD29 → AY:BC(0~4) BL:CD(13~31)
function readBlocks(sheet: ExcelScript.Worksheet): { blocks: { address: string; values: Cell[][] }[]; bad: string[] } {
  const body = sheet.getRange("C6:R140").getValues() as Cell[][];
  const hour = sheet.getRange("AY6:CD29").getValues() as Cell[][];
  const blocks = [
    { address: "C6:C140", values: cols(body, 0, 0) },
    { address: "E6:G140", values: cols(body, 2, 4) },
    { address: "I6:K140", values: cols(body, 6, 8) },
    { address: "L6:M140", values: cols(body, 9, 10) },
    { address: "R6:R140", values: cols(body, 15, 15) },
    { address: "AY6:BC29", values: cols(hour, 0, 4) },
    { address: "BL6:CD29", values: cols(hour, 13, 31) }
  ];
  const bad = blocks.filter(b => hasErr(b.values)).map(b => b.address);
  return { blocks, bad };
}

function readForecast(workbook: ExcelScript.Workbook, today: string): { forecast: ForecastItem[]; next: FullDay | undefined } {
  const forecast: ForecastItem[] = [];
  let next: FullDay | undefined = undefined;

  for (let n = 1; n <= 9; n++) {
    const date = addDays(today, n);
    const found = findSheet(workbook, date);
    const name = found.name;
    const sheet = found.sheet;

    if (!sheet) {
      forecast.push({ date: date, sheetName: name, error: "시트 없음" });
      continue;
    }

    const day = readBlocks(sheet);
    const get = (address: string) => day.blocks.find(b => b.address === address)!.values;

    // 실시간 배정용 명단: E(이름) G(시프트) J(근무시간) — 이름 있는 행만
    const eg = get("E6:G140"), ik = get("I6:K140");
    const staff: Cell[][] = [];
    for (let i = 0; i < eg.length; i++) {
      if (String(eg[i][0] ?? "").trim() === "") continue;
      staff.push([eg[i][0], eg[i][2], ik[i][1]]);
    }

    // 내일 시트는 05:00 바로 넘어가도록 전체를 담아 둡니다.
    if (n === 1 && day.bad.length === 0 && get("C6:C140").some(row => String(row[0]).trim())) {
      next = { date: date, sourceSheet: name, blocks: day.blocks };
    }

    const values = get("AY6:BC29");
    if (hasErr(values)) {
      forecast.push({ date: date, sheetName: name, error: "엑셀 오류", staff: staff });
      continue;
    }
    const otValues = get("BL6:CD29");

    forecast.push({
      date: date,
      sheetName: name,
      address: "AY6:BC29",
      values: values,
      otValues: hasErr(otValues) ? undefined : otValues,
      staff: staff
    });
  }

  return { forecast, next };
}

function sheetName(date: string): string {
  return 시트머리말 + " " + date.slice(5, 7) + "(" + Number(date.slice(8, 10)) + ")";
}

// 시트머리말 시트를 먼저 찾고, 없으면 예전머리말 시트를 찾습니다.
function findSheet(workbook: ExcelScript.Workbook, date: string): { sheet: ExcelScript.Worksheet | undefined; name: string } {
  const name = sheetName(date);
  const sheet = workbook.getWorksheet(name);
  if (sheet || !예전머리말) return { sheet: sheet, name: name };
  const oldName = 예전머리말 + " " + date.slice(5, 7) + "(" + Number(date.slice(8, 10)) + ")";
  const oldSheet = workbook.getWorksheet(oldName);
  return oldSheet ? { sheet: oldSheet, name: oldName } : { sheet: undefined, name: name };
}

function addDays(date: string, n: number): string {
  const d = new Date(date + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function businessClock(value: string): { date: string } {
  const ms = Date.parse(value);

  if (!Number.isFinite(ms)) {
    throw new Error("기준 시각을 확인하세요.");
  }

  const local = new Date(ms + 9 * 3600000);
  const day = new Date(local);

  // 한국 시각 새벽 00~04시는 전날 데일리를 사용합니다.
  if (local.getUTCHours() < 5) {
    day.setUTCDate(day.getUTCDate() - 1);
  }

  return {
    date: day.toISOString().slice(0, 10)
  };
}