function main(workbook: ExcelScript.Workbook): string {
  // ===== 설정 =====
  const 계산모드_자동으로_바꾸기 = false; // true면 '수동 계산'일 때 자동 계산으로 바꿔 줍니다
  const 재계산_시험 = true;               // 전체 재계산 전·후 결과를 비교합니다
  const A1_앞부분_글자수 = 20;            // 로그에 남길 A1 앞부분 글자 수 (0이면 남기지 않음)

  const app = workbook.getApplication();
  const rows: string[][] = [];
  const add = (item: string, value: string, verdict: string): void => {
    rows.push([String(rows.length + 1), item, value, verdict]);
    console.log("[" + verdict + "] " + item + " : " + value);
  };
  const step = (item: string, fn: () => string[]): void => {
    try {
      const r = fn();
      add(item, r[0], r[1]);
    } catch (e) {
      add(item, "읽기 실패: " + String(e), "오류");
    }
  };

  step("실행 시각(UTC)", () => [new Date().toISOString(), "정보"]);
  step("탭 목록", () => [workbook.getWorksheets().map(s => s.getName()).join(", "), "정보"]);

  step("계산 모드", () => {
    const mode = app.getCalculationMode();
    if (mode === ExcelScript.CalculationMode.manual) {
      if (계산모드_자동으로_바꾸기) {
        app.setCalculationMode(ExcelScript.CalculationMode.automatic);
        return ["수동(manual)이어서 자동(automatic)으로 바꿨습니다", "수정함"];
      }
      return ["수동(manual): A1을 바꿔도 계산되지 않는 원인일 수 있습니다", "오류"];
    }
    return [String(mode), mode === ExcelScript.CalculationMode.automatic ? "정상" : "주의"];
  });

  step("통합문서 보호", () => {
    const p = workbook.getProtection().getProtected();
    return [p ? "보호됨" : "해제", p ? "주의" : "정상"];
  });

  const needed = ["상태출력", "수식설정", "데일리", "데일리연결", "수식원문", "수식근무", "수식OT", "수식판정", "수식명단"];
  step("필수 탭", () => {
    const missing = needed.filter(n => !workbook.getWorksheet(n));
    return missing.length > 0 ? ["없음: " + missing.join(", "), "오류"] : ["모두 있음", "정상"];
  });
  step("탭 보호·숨김", () => {
    const out: string[] = [];
    let bad = false;
    for (const n of needed) {
      const s = workbook.getWorksheet(n);
      if (!s) continue;
      const prot = s.getProtection().getProtected();
      out.push(n + ":" + (prot ? "보호" : "해제") + "/" + s.getVisibility());
      if (prot) bad = true;
    }
    return [out.join(", "), bad ? "주의" : "정상"];
  });

  const status = workbook.getWorksheet("상태출력");
  if (status) {
    step("상태출력!A1", () => {
      const r = status.getRange("A1");
      const v = r.getValue();
      const f = r.getFormula();
      const text = String(v === undefined || v === null ? "" : v);
      const lines = text.split(/\r?\n/).filter(x => x.trim() !== "").length;
      const isFormula = typeof f === "string" && f.startsWith("=");
      const head = A1_앞부분_글자수 > 0 ? " / 앞부분: " + text.slice(0, A1_앞부분_글자수).replace(/\s+/g, " ") : "";
      const verdict = text.trim() === "" ? "오류" : (isFormula ? "주의" : "정상");
      return ["타입 " + typeof v + ", 값종류 " + r.getValueType() + ", 길이 " + text.length + "자, 줄 " + lines + "개, 수식 " + (isFormula ? "예" : "아니오") + head, verdict];
    });
    step("상태출력!D5 / 수식설정!B2·B3", () => {
      const st = workbook.getWorksheet("수식설정");
      if (!st) return ["수식설정 탭 없음", "오류"];
      const d5 = status.getRange("D5");
      const b2 = st.getRange("B2").getValue();
      const b3 = st.getRange("B3").getValue();
      const b3f = st.getRange("B3").getFormula();
      return ["D5=" + d5.getText() + "(" + String(d5.getValue()) + "), B2=" + String(b2) + ", B3=" + String(b3) + ", B3수식=" + String(b3f), typeof b3 === "number" ? "정상" : "오류"];
    });
    step("상태출력 AA3~AA6 / AA34", () => {
      const aa = status.getRange("AA3:AA6").getTexts().map(r => r[0]).join(" / ");
      const f34 = status.getRange("AA34").getFormula();
      return [aa + " | AA34수식 " + (typeof f34 === "string" && f34.startsWith("=") ? "있음" : "없음"), "정보"];
    });
  }

  step("데일리!A1", () => {
    const d = workbook.getWorksheet("데일리");
    if (!d) return ["데일리 탭 없음", "오류"];
    const r = d.getRange("A1");
    return ["값 " + String(r.getValue()) + ", 표시 " + r.getText(), "정보"];
  });

  step("이름 정의", () => {
    const names = ["긴조_자동기록", "AA누적_설치버전", "AA누적_운영일", "긴조_클라우드최신접수", "긴조_마지막입력"];
    const out = names.map(n => {
      const it = workbook.getNamedItem(n);
      return n + "=" + (it ? it.getFormula() : "(없음)");
    });
    return [out.join(" | "), "정보"];
  });

  const spills: string[][] = [["수식원문", "D2"], ["수식근무", "A2"], ["수식근무", "P2"], ["수식OT", "G2"], ["수식판정", "P2"], ["수식명단", "AB2"]];
  for (const sp of spills) {
    step("수식 " + sp[0] + "!" + sp[1], () => checkSpill(workbook, sp[0], sp[1]));
  }

  if (재계산_시험) {
    step("재계산 시험", () => {
      const before = spillRows(workbook, "수식원문", "D2") + "/" + spillRows(workbook, "수식판정", "P2");
      const t0 = Date.now();
      app.calculate(ExcelScript.CalculationType.full);
      const ms = Date.now() - t0;
      const after = spillRows(workbook, "수식원문", "D2") + "/" + spillRows(workbook, "수식판정", "P2");
      return ["값 있는 행(원문/판정) 전 " + before + " → 후 " + after + ", " + ms + "ms",
        before === after ? "정보" : "오류(재계산하니 결과가 바뀜: 자동 계산이 꺼져 있었음)"];
    });
  }

  const errCount = rows.filter(r => r[3].startsWith("오류")).length;
  add("요약", "오류 " + errCount + "개 (판정이 '오류'인 줄을 확인하세요)", errCount > 0 ? "오류" : "정상");

  const old = workbook.getWorksheet("디버그로그");
  if (old) old.delete();
  const sheet = workbook.addWorksheet("디버그로그");
  const data: string[][] = [["#", "항목", "내용", "판정"]].concat(rows);
  const area = sheet.getRangeByIndexes(0, 0, data.length, 4);
  area.setNumberFormat("@");
  area.setValues(data);
  sheet.getRange("A:A").getFormat().setColumnWidth(30);
  sheet.getRange("B:B").getFormat().setColumnWidth(170);
  sheet.getRange("C:C").getFormat().setColumnWidth(700);
  sheet.getRange("C:C").getFormat().setWrapText(true);
  sheet.getRange("D:D").getFormat().setColumnWidth(90);
  sheet.activate();

  return data.map(r => r.join(" | ")).join("\n");
}

function spillRows(workbook: ExcelScript.Workbook, sheetName: string, address: string): number {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) return -1;
  const anchor = sheet.getRange(address);
  const spilled = anchor.getSpillingToRange();
  const values = (spilled ? spilled : anchor).getValues();
  return values.filter(row => row.some(v => String(v).trim() !== "")).length;
}

function checkSpill(workbook: ExcelScript.Workbook, sheetName: string, address: string): string[] {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) return ["탭 없음", "오류"];
  const anchor = sheet.getRange(address);
  const formula = anchor.getFormula();
  if (typeof formula !== "string" || !formula.startsWith("=")) {
    return ["수식이 없습니다(현재 값: " + String(anchor.getValue()) + ")", "오류"];
  }
  const spilled = anchor.getSpillingToRange();
  const range = spilled ? spilled : anchor;
  const values = range.getValues();
  const types = range.getValueTypes();
  let errorCount = 0;
  let firstError = "";
  for (let r = 0; r < types.length; r++) {
    for (let c = 0; c < types[r].length; c++) {
      if (types[r][c] === ExcelScript.RangeValueType.error) {
        errorCount++;
        if (firstError === "") firstError = String(values[r][c]) + " (" + (r + 1) + "행 " + (c + 1) + "열)";
      }
    }
  }
  const filled = values.filter(row => row.some(v => String(v).trim() !== "")).length;
  const info = (spilled ? "스필 " : "단일셀 ") + values.length + "행×" + (values[0] ? values[0].length : 0) + "열, 값 있는 행 " + filled
    + (errorCount > 0 ? ", 오류 " + errorCount + "개: " + firstError : "") + " | 수식 " + formula.slice(0, 50);
  return [info, errorCount > 0 ? "오류" : (filled === 0 ? "주의" : "정상")];
}
