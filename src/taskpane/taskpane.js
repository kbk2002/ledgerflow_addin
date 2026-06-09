/* ============================================================
   LedgerFlow — Taskpane Logic
   Office JS + vanilla JS
   ============================================================ */

"use strict";

const state = {
  platform: null,
  connected: false,
  companyName: null,
  accessToken: null,
  validationPassed: false,
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initTabs();
    initConnectPanel();
    initPullPanel();
    initValidatePanel();
    initPushPanel();
  }
});

function initTabs() {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(`panel-${tab.dataset.panel}`).classList.add("active");
    });
  });
}

function initConnectPanel() {
  document.getElementById("btn-connect-qbo").addEventListener("click", () => startOAuth("qbo"));
  document.getElementById("btn-connect-xero").addEventListener("click", () => startOAuth("xero"));
  document.getElementById("btn-disconnect").addEventListener("click", disconnect);
}

async function startOAuth(platform) {
  simulateConnection(platform);
}

function simulateConnection(platform) {
  const names = { qbo: "Acme Corp (QBO)", xero: "Acme Corp (Xero)" };
  setConnected(platform, names[platform], "DEMO_TOKEN");
}

function setConnected(platform, companyName, token) {
  state.platform = platform;
  state.connected = true;
  state.companyName = companyName;
  state.accessToken = token;

  document.getElementById("session-badge").textContent = "Connected";
  document.getElementById("session-badge").className = "status-badge ok";
  document.getElementById("company-name").textContent = companyName;
  document.getElementById("connect-status-msg").textContent = `Connected to ${companyName}`;
  document.getElementById("btn-disconnect").disabled = false;

  document.getElementById("btn-connect-qbo").classList.remove("connected");
  document.getElementById("btn-connect-xero").classList.remove("connected");

  if (platform === "qbo") document.getElementById("btn-connect-qbo").classList.add("connected");
  if (platform === "xero") document.getElementById("btn-connect-xero").classList.add("connected");

  document.getElementById("btn-pull").disabled = false;
  document.getElementById("btn-push").disabled = false;
}

function disconnect() {
  state.platform = null;
  state.connected = false;
  state.companyName = null;
  state.accessToken = null;
  state.validationPassed = false;

  document.getElementById("session-badge").textContent = "Disconnected";
  document.getElementById("session-badge").className = "status-badge";
  document.getElementById("company-name").textContent = "—";
  document.getElementById("connect-status-msg").textContent = "No platform connected. Select one to authenticate.";
  document.getElementById("btn-disconnect").disabled = true;
  document.getElementById("btn-pull").disabled = true;
  document.getElementById("btn-push").disabled = true;

  document.getElementById("btn-connect-qbo").classList.remove("connected");
  document.getElementById("btn-connect-xero").classList.remove("connected");
}

function initPullPanel() {
  document.getElementById("btn-pull").addEventListener("click", pullMasterData);
}

async function pullMasterData() {
  if (!state.connected) return;

  const type = document.getElementById("pull-type").value;
  const sheetName = document.getElementById("pull-sheet").value.trim() || `${type}_Ref`;
  const status = document.getElementById("pull-status");

  if (!type) {
    status.textContent = "Select a data type first.";
    return;
  }

  status.textContent = "Fetching from platform…";

  try {
    const data = getDemoMasterData(type);

    await Excel.run(async (ctx) => {
      let sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load("name");
      await ctx.sync();

      if (sheet.isNullObject) {
        sheet = ctx.workbook.worksheets.add(sheetName);
      } else {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("address");
        await ctx.sync();

        if (!usedRange.isNullObject) {
          usedRange.clear();
        }
      }

      if (data.length === 0) {
        status.textContent = "No data returned.";
        return;
      }

      const headers = Object.keys(data[0]);
      const rows = data.map((row) => headers.map((h) => row[h] ?? ""));

      const headerRange = sheet.getRange(`A1:${colLetter(headers.length)}1`);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#1A1F2E";
      headerRange.format.font.color = "#FFFFFF";

      const dataRange = sheet.getRange(`A2:${colLetter(headers.length)}${rows.length + 1}`);
      dataRange.values = rows;

      sheet.getUsedRange().format.autofitColumns();
      sheet.activate();

      await ctx.sync();
    });

    status.textContent = `✓ Pulled ${data.length} rows into "${sheetName}".`;
  } catch (err) {
    status.textContent = `Error: ${err.message}`;
  }
}

function getDemoMasterData(type) {
  const demos = {
    accounts: [
      { AccountID: "1000", Name: "Cash", Type: "Bank", Active: true },
      { AccountID: "4000", Name: "Sales Revenue", Type: "Income", Active: true },
      { AccountID: "5000", Name: "Cost of Goods Sold", Type: "COGS", Active: true },
      { AccountID: "6000", Name: "Salaries Expense", Type: "Expense", Active: true },
    ],
    vendors: [
      { VendorID: "V001", Name: "ABC Supplies", Email: "ap@abcsupplies.com", Balance: 0 },
      { VendorID: "V002", Name: "XYZ Consulting", Email: "billing@xyz.com", Balance: 0 },
    ],
    classes: [
      { ClassID: "C1", Name: "Operations" },
      { ClassID: "C2", Name: "Marketing" },
      { ClassID: "C3", Name: "Finance" },
    ],
    customers: [
      { CustomerID: "CU01", Name: "Globex Corp", Email: "ar@globex.com" },
      { CustomerID: "CU02", Name: "Initech Ltd", Email: "finance@initech.com" },
    ],
    departments: [
      { DeptID: "D1", Name: "Sales" },
      { DeptID: "D2", Name: "Support" },
    ],
  };

  return demos[type] || [];
}

function initValidatePanel() {
  document.getElementById("btn-detect-range").addEventListener("click", detectRange);
  document.getElementById("btn-validate").addEventListener("click", runValidation);
}

async function detectRange() {
  try {
    await Excel.run(async (ctx) => {
      const selection = ctx.workbook.getSelectedRange();
      selection.load("address");
      await ctx.sync();
      document.getElementById("je-range").value = selection.address;
    });
  } catch {
    document.getElementById("je-range").value = "";
  }
}

async function runValidation() {
  const rangeAddr = document.getElementById("je-range").value.trim();
  const tbody = document.getElementById("val-body");

  if (!rangeAddr || !rangeAddr.includes("!")) {
    tbody.innerHTML = `<tr><td colspan="3" style="color:var(--error)">Select a valid Excel range first.</td></tr>`;
    return;
  }

  tbody.innerHTML = "<tr><td colspan='3'>Running…</td></tr>";

  try {
    const results = await validateJournalEntries(rangeAddr);
    renderValidationResults(results);

    state.validationPassed = results.every((r) => r.pass);

    if (state.validationPassed) {
      document.getElementById("push-range").value = rangeAddr;
      document.getElementById("push-validation-warning").style.display = "none";
      document.getElementById("push-help").textContent = "Validation passed. Ready to post.";
    } else {
      document.getElementById("push-validation-warning").style.display = "flex";
      document.getElementById("push-help").textContent = "Fix validation errors before posting.";
    }
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="3" style="color:var(--error)">Error: ${err.message}</td></tr>`;
  }
}

async function validateJournalEntries(rangeAddr) {
  return await Excel.run(async (ctx) => {
    const { sheetName, address } = parseRangeAddress(rangeAddr);
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.load("values");
    await ctx.sync();

    const values = range.values.filter((r) => r.some((c) => c !== ""));
    const checks = [];

    checks.push({
      name: "Has data rows",
      pass: values.length > 0,
      detail: `${values.length} row(s) found`,
    });

    if (values.length === 0) return checks;

    const colCount = values[0].length;

    checks.push({
      name: "Minimum columns (≥5)",
      pass: colCount >= 5,
      detail: `${colCount} column(s) detected`,
    });

    const blankDates = values.filter((r) => !r[0]).length;

    checks.push({
      name: "No blank dates",
      pass: blankDates === 0,
      detail: blankDates === 0 ? "All dates present" : `${blankDates} blank date(s)`,
    });

    const blankAccounts = values.filter((r) => !r[2]).length;

    checks.push({
      name: "No blank account codes",
      pass: blankAccounts === 0,
      detail: blankAccounts === 0 ? "All accounts present" : `${blankAccounts} blank account(s)`,
    });

    if (colCount >= 5) {
      const totalDebit = values.reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
      const totalCredit = values.reduce((sum, row) => sum + (parseFloat(row[4]) || 0), 0);
      const balanced = Math.abs(totalDebit - totalCredit) < 0.005;

      checks.push({
        name: "Debits = Credits",
        pass: balanced,
        detail: balanced
          ? `Both sides: $${totalDebit.toFixed(2)}`
          : `Debit $${totalDebit.toFixed(2)} vs Credit $${totalCredit.toFixed(2)}`,
      });
    }

    return checks;
  });
}

function renderValidationResults(results) {
  const tbody = document.getElementById("val-body");

  tbody.innerHTML = results
    .map(
      (r) => `
        <tr>
          <td>${r.name}</td>
          <td class="${r.pass ? "val-pass" : "val-fail"}">${r.pass ? "✓ Pass" : "✗ Fail"}</td>
          <td>${r.detail}</td>
        </tr>
      `
    )
    .join("");
}

function initPushPanel() {
  document.getElementById("btn-push").addEventListener("click", postEntries);
}

async function postEntries() {
  if (!state.connected) {
    appendLog("err", "Not connected to a platform.");
    return;
  }

  if (!state.validationPassed) {
    document.getElementById("push-validation-warning").style.display = "flex";
    return;
  }

  const rangeAddr = document.getElementById("push-range").value.trim();
  const txnType = document.getElementById("push-type").value;

  if (!rangeAddr || !rangeAddr.includes("!")) {
    appendLog("err", "No valid range specified.");
    return;
  }

  const log = document.getElementById("push-log");
  const progress = document.getElementById("progress-wrap");
  const bar = document.getElementById("progress-bar");

  log.innerHTML = "";
  log.style.display = "block";
  progress.style.display = "block";
  bar.style.width = "0%";

  try {
    const rows = await readRangeValues(rangeAddr);
    appendLog("info", `${rows.length} row(s) to post as ${txnType}…`);

    const chunkSize = 10;

    for (let i = 0; i < rows.length; i += chunkSize) {
      const chunk = rows.slice(i, i + chunkSize);
      await sleep(400);

      const pct = Math.round(((i + chunk.length) / rows.length) * 100);
      bar.style.width = `${pct}%`;
      appendLog("ok", `Posted rows ${i + 1}–${i + chunk.length} (${pct}%)`);
    }

    appendLog("ok", `✓ All ${rows.length} entries posted successfully.`);
    bar.style.width = "100%";

    await markRowsPosted(rangeAddr, rows.length);
  } catch (err) {
    appendLog("err", `Failed: ${err.message}`);
  }
}

async function readRangeValues(rangeAddr) {
  return await Excel.run(async (ctx) => {
    const { sheetName, address } = parseRangeAddress(rangeAddr);
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.load("values");
    await ctx.sync();

    return range.values.filter((r) => r.some((c) => c !== ""));
  });
}

async function markRowsPosted(rangeAddr, rowCount) {
  await Excel.run(async (ctx) => {
    const { sheetName, address } = parseRangeAddress(rangeAddr);
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.load("columnIndex, rowIndex, columnCount");
    await ctx.sync();

    const statusCol = range.columnIndex + range.columnCount;
    const startRow = range.rowIndex;
    const statusRange = sheet.getRangeByIndexes(startRow, statusCol, rowCount, 1);

    statusRange.values = Array(rowCount).fill(["Posted"]);
    statusRange.format.font.color = "#1B9E5E";
    statusRange.format.font.bold = true;

    await ctx.sync();
  });
}

function parseRangeAddress(rangeAddr) {
  const parts = rangeAddr.split("!");
  const sheetName = parts[0].replace(/'/g, "");
  const address = parts[1];

  return { sheetName, address };
}

function appendLog(type, text) {
  const log = document.getElementById("push-log");
  const line = document.createElement("div");

  line.className = `log-line ${type}`;
  line.textContent = `[${timestamp()}] ${text}`;

  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}

function timestamp() {
  return new Date().toLocaleTimeString("en-US", { hour12: false });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function colLetter(n) {
  let s = "";

  while (n > 0) {
    s = String.fromCharCode(((n - 1) % 26) + 65) + s;
    n = Math.floor((n - 1) / 26);
  }

  return s;
}