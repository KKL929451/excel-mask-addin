/* Office.js browser add-in V3. Formula cells are never changed. */
(() => {
  "use strict";

  const CFG = {
    sheet: "Main",
    passwordCell: "S1",
    statusCell: "T1",
    password: "D", // IMPORTANT: visible in source. Visual masking only, not security.
    blocks: [{ address: "B2:G1048576", colIndex: 1 },
             { address: "J2:O1048576", colIndex: 9 }],
    chunkRows: 150,
    maxRowsPerBlock: 30000
  };

  const $ = (id) => document.getElementById(id);
  let operationQueue = Promise.resolve();

  function display(message, kind = "info") {
    const box = $("status");
    box.textContent = message;
    box.dataset.kind = kind;
  }

  function queueOperation(fn) {
    operationQueue = operationQueue.catch(() => {}).then(fn);
    return operationQueue.catch((e) => {
      console.error(e);
      display(`Error: ${e && e.message ? e.message : e}`, "error");
    });
  }

  function changedPasswordCell(address) {
    // Event address can contain worksheet names, $ signs and (for paste) a range.
    const local = address.split("!").pop().replace(/\$/g, "").toUpperCase();
    const parts = local.split(":");
    function coords(cell) {
      const m = /^([A-Z]+)([0-9]+)$/.exec(cell);
      if (!m) return null;
      let c = 0;
      for (const l of m[1]) c = c * 26 + l.charCodeAt(0) - 64;
      return { row: Number(m[2]), col: c };
    }
    const a = coords(parts[0]);
    const b = coords(parts[1] || parts[0]);
    return !!(a && b && Math.min(a.row, b.row) <= 1 &&
      Math.max(a.row, b.row) >= 1 && Math.min(a.col, b.col) <= 19 &&
      Math.max(a.col, b.col) >= 19);
  }

  async function readState() {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CFG.sheet);
      const pwd = sheet.getRange(CFG.passwordCell);
      const state = sheet.getRange(CFG.statusCell);
      pwd.load("values");
      state.load("values");
      await context.sync();
      return { password: String(pwd.values[0][0] ?? ""), status: String(state.values[0][0] ?? "") };
    });
  }

  async function refreshState() {
    const state = await readState();
    $("workbook-status").textContent = state.status || "NOT INITIALIZED";
    display(`Workbook state: ${state.status || "not initialized"}. S1 is ${state.password ? "filled" : "empty"}.`, "info");
  }

  async function applyPassword(forceLock = false) {
    display("Reading workbook and preparing changes…", "info");
    const outcome = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CFG.sheet);
      const pwd = sheet.getRange(CFG.passwordCell);
      const stateCell = sheet.getRange(CFG.statusCell);
      pwd.load("values");
      stateCell.load("values");
      await context.sync();

      const state = String(stateCell.values[0][0] ?? "");
      const supplied = String(pwd.values[0][0] ?? "");
      if (state !== "LOCKED" && state !== "UNLOCKED") {
        throw new Error("T1 must be LOCKED or UNLOCKED. On a fresh, readable test workbook, initialize it with the desktop VBA SetupProtection macro first.");
      }
      const wantLocked = forceLock || supplied !== CFG.password;
      if ((wantLocked && state === "LOCKED") || (!wantLocked && state === "UNLOCKED")) {
        if (forceLock) pwd.clear(Excel.ClearApplyTo.contents);
        await context.sync();
        return { count: 0, state, unchanged: true };
      }

      // Find last used row across the two selected blocks (ignores formatting).
      const used = CFG.blocks.map(b => sheet.getRange(b.address).getUsedRangeOrNullObject(true));
      await context.sync();
      for (let i = 0; i < used.length; i++) {
        if (!used[i].isNullObject) used[i].load("rowIndex,rowCount");
      }
      await context.sync();
      // Keep bounds in matching order (including empty blocks).
      const lastRows = used.map(x => x.isNullObject ? 1 : x.rowIndex + x.rowCount);
      const requests = [];
      for (let i = 0; i < CFG.blocks.length; i++) {
        const last = lastRows[i];
        if (last < 2) continue;
        if (last > CFG.maxRowsPerBlock) {
          throw new Error(`Block ${CFG.blocks[i].address} has data down to row ${last}; adjust maxRowsPerBlock in app.js or reduce the data range.`);
        }
        for (let row = 2; row <= last; row += CFG.chunkRows) {
          const count = Math.min(CFG.chunkRows, last - row + 1);
          const range = sheet.getRangeByIndexes(row - 1, CFG.blocks[i].colIndex, count, 6);
          range.load("values,formulas");
          requests.push({ range, row, startCol: CFG.blocks[i].colIndex, count });
        }
      }
      await context.sync();

      // Prevalidate *every* output before writing anything.
      const jobs = [];
      let countChanged = 0;
      let countSkippedFormulas = 0;
      for (const req of requests) {
        const values = [];
        const textFormats = [];
        let changedHere = false;
        for (let r = 0; r < req.count; r++) {
          const valueRow = [];
          const formatsRow = [];
          for (let c = 0; c < 6; c++) {
            const original = req.range.values[r][c];
            const formula = req.range.formulas[r][c];
            const label = `${String.fromCharCode(65 + req.startCol + c)}${req.row + r}`;
            if (typeof formula === "string" && formula.startsWith("=")) {
              // Formula is intentionally left in place, including its result.
              // The null values below leave this cell and its formatting unchanged.
              valueRow.push(null);
              formatsRow.push(null);
              countSkippedFormulas++;
              continue;
            }
            if (original === "" || original === null || original === undefined) {
              valueRow.push(null);  // null preserves the original blank cell
              formatsRow.push(null);
              continue;
            }
            const transformed = MaskCodec.transformText(
              String(original), req.row + r, req.startCol + c + 1,
              wantLocked ? 1 : -1
            );
            if (transformed.length > 32767) throw new Error(`Cell ${label} exceeds Excel's text limit.`);
            // V3 masked output never starts with =, + or -. When
            // decrypting actual plaintext that starts with those signs,
            // Excel Range.values would interpret it as a formula, so stop.
            if (/^[=+\-]/.test(transformed)) {
              throw new Error(`Cell ${label} would start with a formula character when displayed. Restore this text safely in desktop Excel or as a text-formatted value.`);
            }
            valueRow.push(transformed);
            formatsRow.push("@");
            countChanged++;
            changedHere = true;
          }
          values.push(valueRow);
          textFormats.push(formatsRow);
        }
        if (changedHere) jobs.push({ range: req.range, values, textFormats });
      }

      // Commit the prepared modifications, preserving blanks and styling.
      // Network interruption could still result in partial writes: use a backup.
      for (const job of jobs) {
        job.range.numberFormat = job.textFormats;
        job.range.values = job.values;
      }
      stateCell.values = [[wantLocked ? "LOCKED" : "UNLOCKED"]];
      if (wantLocked) pwd.clear(Excel.ClearApplyTo.contents);
      await context.sync();
      return {
        count: countChanged,
        skipped: countSkippedFormulas,
        state: wantLocked ? "LOCKED" : "UNLOCKED",
        unchanged: false
      };
    });
    $("workbook-status").textContent = outcome.state;
    display(outcome.unchanged
      ? `Already ${outcome.state.toLowerCase()}; no shifting performed.`
      : `${outcome.state}: processed ${outcome.count} cells. ` +
        (outcome.skipped
          ? `Skipped ${outcome.skipped} formula cell(s). Their calculated values are NOT masked.`
          : `No formula cells skipped.`),
      outcome.skipped ? "info" : "ok");
  }

  function onWorksheetChanged(event) {
    if (!changedPasswordCell(event.address)) return;
    // Ignore changes made by another coauthor in this client. Simultaneous edits
    // remain inherently unsafe and should be avoided during masking.
    if (event.source === Excel.EventSource.remote) return;
    queueOperation(() => applyPassword());
  }

  async function begin() {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      display("This version of Excel does not support the worksheet change event.", "error");
      return;
    }
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(CFG.sheet);
      sheet.onChanged.add(onWorksheetChanged);
      await context.sync();
    });
    $("monitoring").textContent = "Active — keep pane open";
    await refreshState();
  }

  $("apply").addEventListener("click", () => queueOperation(() => applyPassword()));
  $("lock").addEventListener("click", () => queueOperation(() => applyPassword(true)));
  $("refresh").addEventListener("click", () => queueOperation(refreshState));

  if (typeof Office === "undefined") {
    display("Open this page through Excel > My Add-ins. It cannot access a workbook in a normal browser tab.", "error");
    return;
  }
  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) {
      display("This add-in must be opened inside Excel.", "error");
      return;
    }
    queueOperation(begin);
  });
})();
