/* Office.js task pane UI, compact layout for the existing V3 codec.
 * The masking algorithm is unchanged. Formula cells are left untouched.
 */
(() => {
  "use strict";

  const CFG = {
    sheet: "Main",
    passwordCell: "S1",
    statusCell: "T1",
    password: "D", // Visual masking only; this is visible in source code.
    blocks: [
      { address: "B2:G1048576", colIndex: 1 },
      { address: "J2:O1048576", colIndex: 9 }
    ],
    chunkRows: 150,
    maxRowsPerBlock: 30000
  };

  const $ = id => document.getElementById(id);
  let operationQueue = Promise.resolve();

  function setWorkbookStatus(raw) {
    const state = (raw || "NOT INITIALIZED").toUpperCase();
    $("workbook-status").textContent = state;
    $("workbook-status").dataset.state = state;
  }

  function clearError() {
    $("status").hidden = true;
    $("status").textContent = "";
  }

  function showError(message) {
    $("status").textContent = message;
    $("status").hidden = false;
  }

  function queueOperation(fn) {
    operationQueue = operationQueue.catch(() => {}).then(async () => {
      clearError();
      $("apply").disabled = true;
      $("lock").disabled = true;
      try {
        return await fn();
      } finally {
        $("apply").disabled = false;
        $("lock").disabled = false;
      }
    });
    return operationQueue.catch(err => {
      console.error(err);
      showError(err && err.message ? err.message : String(err));
    });
  }

  function changedPasswordCell(address) {
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

  async function refreshState() {
    await Excel.run(async context => {
      const stateCell = context.workbook.worksheets.getItem(CFG.sheet)
        .getRange(CFG.statusCell);
      stateCell.load("values");
      await context.sync();
      setWorkbookStatus(String(stateCell.values[0][0] ?? ""));
    });
  }

  // Explicit decrypt requires the right password; Encrypt always masks and clears S1.
  // Automatic S1 edits retain the original V3 toggle behavior.
  async function applyPassword(forceLock = false, requireCorrectPassword = false) {
    const outcome = await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getItem(CFG.sheet);
      const pwd = sheet.getRange(CFG.passwordCell);
      const stateCell = sheet.getRange(CFG.statusCell);
      pwd.load("values");
      stateCell.load("values");
      await context.sync();

      const state = String(stateCell.values[0][0] ?? "");
      const supplied = String(pwd.values[0][0] ?? "");
      if (state !== "LOCKED" && state !== "UNLOCKED") {
        throw new Error("The workbook is not initialized: T1 must be LOCKED or UNLOCKED.");
      }
      if (requireCorrectPassword && supplied !== CFG.password) {
        throw new Error("Enter the correct password in S1 before decrypting.");
      }
      const wantLocked = forceLock || supplied !== CFG.password;
      if ((wantLocked && state === "LOCKED") || (!wantLocked && state === "UNLOCKED")) {
        if (forceLock) pwd.clear(Excel.ClearApplyTo.contents);
        await context.sync();
        return { state, unchanged: true };
      }

      // Find last used rows, ignoring formatting.
      const used = CFG.blocks.map(b =>
        sheet.getRange(b.address).getUsedRangeOrNullObject(true));
      await context.sync();
      for (const range of used) {
        if (!range.isNullObject) range.load("rowIndex,rowCount");
      }
      await context.sync();
      const lastRows = used.map(x =>
        x.isNullObject ? 1 : x.rowIndex + x.rowCount);

      const requests = [];
      for (let i = 0; i < CFG.blocks.length; i++) {
        const last = lastRows[i];
        if (last < 2) continue;
        if (last > CFG.maxRowsPerBlock) {
          throw new Error(`Data extends to row ${last}, beyond the configured limit.`);
        }
        for (let row = 2; row <= last; row += CFG.chunkRows) {
          const count = Math.min(CFG.chunkRows, last - row + 1);
          const range = sheet.getRangeByIndexes(
            row - 1, CFG.blocks[i].colIndex, count, 6);
          range.load("values,formulas");
          requests.push({range, row, startCol: CFG.blocks[i].colIndex, count});
        }
      }
      await context.sync();

      // Prepare all changes before writing. Null preserves formula/blank cells.
      const jobs = [];
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
              valueRow.push(null);
              formatsRow.push(null);
              continue;
            }
            if (original === "" || original === null || original === undefined) {
              valueRow.push(null);
              formatsRow.push(null);
              continue;
            }
            const transformed = MaskCodec.transformText(
              String(original), req.row + r, req.startCol + c + 1,
              wantLocked ? 1 : -1
            );
            if (transformed.length > 32767) {
              throw new Error(`Cell ${label} exceeds Excel's character limit.`);
            }
            // Keep the original V3 safeguard against formula interpretation.
            if (/^[=+\-]/.test(transformed)) {
              throw new Error(`Cell ${label} would be interpreted as a formula. Restore it with desktop Excel.`);
            }
            valueRow.push(transformed);
            formatsRow.push("@");
            changedHere = true;
          }
          values.push(valueRow);
          textFormats.push(formatsRow);
        }
        if (changedHere) jobs.push({range: req.range, values, textFormats});
      }

      // As with V3, interruptions during writes can cause partial updates.
      for (const job of jobs) {
        job.range.numberFormat = job.textFormats;
        job.range.values = job.values;
      }
      stateCell.values = [[wantLocked ? "LOCKED" : "UNLOCKED"]];
      if (wantLocked) pwd.clear(Excel.ClearApplyTo.contents);
      await context.sync();
      return { state: wantLocked ? "LOCKED" : "UNLOCKED", unchanged: false };
    });
    setWorkbookStatus(outcome.state);
  }

  function onWorksheetChanged(event) {
    if (!changedPasswordCell(event.address)) return;
    if (event.source === Excel.EventSource.remote) return;
    queueOperation(() => applyPassword());
  }

  async function begin() {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      setWorkbookStatus("UNAVAILABLE");
      throw new Error("This Excel version cannot monitor worksheet changes.");
    }
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getItem(CFG.sheet);
      sheet.onChanged.add(onWorksheetChanged);
      await context.sync();
    });
    await refreshState();
  }

  $("apply").addEventListener("click", () =>
    queueOperation(() => applyPassword(false, true)));
  $("lock").addEventListener("click", () =>
    queueOperation(() => applyPassword(true)));

  if (typeof Office === "undefined") {
    setWorkbookStatus("UNAVAILABLE");
    showError("Open this add-in from Excel.");
    return;
  }
  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) {
      setWorkbookStatus("UNAVAILABLE");
      showError("Open this add-in from Excel.");
      return;
    }
    queueOperation(begin);
  });
})();
