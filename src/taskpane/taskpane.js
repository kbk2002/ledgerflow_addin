/* ============================================================
   LedgerFlow — Taskpane Logic
   Office JS + vanilla JS
   ============================================================ */

"use strict";

const API_BASE_URL = "https://ledgerflow-addin.vercel.app";

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
initMappingPanel();
initTemplatePanel();
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
  try {
    const res = await fetch(
      `${API_BASE_URL}/api/auth?platform=${platform}`
    );

    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Connection failed");
    }

    setConnected(
      platform,
      data.companyName,
      data.accessToken
    );
  } catch (err) {
    showConnectError(err.message);
  }
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

const refreshBtn = document.getElementById("btn-refresh-all");
if (refreshBtn) {
  refreshBtn.disabled = false;
}
}
updateConnectionManager();

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

const refreshBtn = document.getElementById("btn-refresh-all");
if (refreshBtn) {
  refreshBtn.disabled = true;
}
  document.getElementById("btn-connect-qbo").classList.remove("connected");
  document.getElementById("btn-connect-xero").classList.remove("connected");
}
updateConnectionManager();
function showConnectError(message) {
  document.getElementById(
    "connect-status-msg"
  ).textContent = `Error: ${message}`;
}

function initPullPanel() {
  document.getElementById("btn-pull").addEventListener("click", pullMasterData);

  const refreshBtn = document.getElementById("btn-refresh-all");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", refreshAllReferenceData);
  }
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
   const res = await fetch(
  `${API_BASE_URL}/api/masterdata?platform=${state.platform}&type=${type}`
);

const payload = await res.json();

if (!res.ok) {
  throw new Error(
    payload.error || "Failed to pull master data"
  );
}

const data = payload.data;

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
async function writeDataToSheet(sheetName, data) {
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

    if (!data || data.length === 0) {
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

    await ctx.sync();
  });
}
async function refreshAllReferenceData() {
  if (!state.connected) return;

  const status = document.getElementById("pull-status");
  status.textContent = "Refreshing all reference data…";

  const datasets = [
    { type: "accounts", sheetName: "Accounts_Ref" },
    { type: "vendors", sheetName: "Vendors_Ref" },
    { type: "customers", sheetName: "Customers_Ref" }
  ];

  try {
    for (const item of datasets) {
      const res = await fetch(
        `${API_BASE_URL}/api/masterdata?platform=${state.platform}&type=${item.type}`
      );

      const payload = await res.json();

      if (!res.ok) {
        throw new Error(payload.error || `Failed to refresh ${item.type}`);
      }

      await writeDataToSheet(item.sheetName, payload.data);
    }

await createDashboard();

status.textContent = "✓ Refreshed Accounts, Vendors, and Customers.";
  } catch (err) {
    status.textContent = `Error: ${err.message}`;
  }
}
async function createDashboard() {
  await Excel.run(async (ctx) => {

    let sheet =
      ctx.workbook.worksheets.getItemOrNullObject(
        "LedgerFlow_Dashboard"
      );

    sheet.load("name");
    await ctx.sync();

    if (sheet.isNullObject) {
      sheet = ctx.workbook.worksheets.add(
        "LedgerFlow_Dashboard"
      );
    }

    const dashboard = [
      ["Metric", "Value"],
      ["Platform", state.platform || "Unknown"],
      ["Accounts", 4],
      ["Vendors", 2],
      ["Customers", 2],
      ["Last Refresh", new Date().toLocaleString()]
    ];

    const range =
      sheet.getRange(`A1:B${dashboard.length}`);

    range.values = dashboard;

    const header =
      sheet.getRange("A1:B1");

    header.format.font.bold = true;
    header.format.fill.color = "#1A1F2E";
    header.format.font.color = "#FFFFFF";

    sheet.getUsedRange().format.autofitColumns();

    await ctx.sync();
  });
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
await clearValidationHighlights(rangeAddr);

const results = await validateJournalEntries(rangeAddr);
renderValidationResults(results);

const hasErrors = results.some((r) => !r.pass);

if (hasErrors) {
  await highlightValidationErrors(rangeAddr);
}

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

    const mappedAccounts = await getMappedAccounts();

const dataRows = values[0][0] === "Date" ? values.slice(1) : values;

const unmappedAccounts = dataRows.filter((row) => {
  const account = String(row[2] || "").trim();
  return account && !mappedAccounts.includes(account);
});

    checks.push({
      name: "Account mapping validation",
      pass: unmappedAccounts.length === 0,
      detail:
        unmappedAccounts.length === 0
          ? "All accounts are mapped"
          : `${unmappedAccounts.length} unmapped account row(s)`,
    });

    if (colCount >= 5) {
      const totalDebit = dataRows.reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
      const totalCredit = dataRows.reduce((sum, row) => sum + (parseFloat(row[4]) || 0), 0);
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
function initMappingPanel() {
  const btn = document.getElementById("btn-create-mapping");

  if (btn) {
    btn.addEventListener("click", createMappingSheet);
  }
}

async function createMappingSheet() {
  const status = document.getElementById("mapping-status");

  if (status) {
    status.textContent = "Creating mapping sheet...";
  }

  try {
    await Excel.run(async (ctx) => {
      let sheet = ctx.workbook.worksheets.getItemOrNullObject("LedgerFlow_Mappings");

      sheet.load("name");
      await ctx.sync();

      if (sheet.isNullObject) {
        sheet = ctx.workbook.worksheets.add("LedgerFlow_Mappings");
      }

      sheet.getRange("A1:D1").values = [[
        "Excel Account",
        "ERP Account",
        "Platform",
        "Status"
      ]];

      sheet.getRange("A2:D5").values = [
        ["6000", "Office Expense", "QBO", "Mapped"],
        ["6100", "Rent Expense", "QBO", "Mapped"],
        ["6200", "Payroll Expense", "QBO", "Mapped"],
        ["1000", "Cash", "QBO", "Mapped"]
      ];

      sheet.getRange("A1:D1").format.font.bold = true;

      sheet.getUsedRange().format.autofitColumns();

      sheet.activate();

      await ctx.sync();
    });

    if (status) {
      status.textContent = "✓ Mapping sheet created";
    }

  } catch (err) {

    if (status) {
      status.textContent = err.message;
    }

    console.error(err);
  }
}
function initTemplatePanel() {
  const journalBtn = document.getElementById("btn-journal-template");
  const payrollBtn = document.getElementById("btn-payroll-template");
  const vendorBtn = document.getElementById("btn-vendor-template");

  if (journalBtn) {
    journalBtn.addEventListener("click", () => createTemplate("journal"));
  }

  if (payrollBtn) {
    payrollBtn.addEventListener("click", () => createTemplate("payroll"));
  }

  if (vendorBtn) {
    vendorBtn.addEventListener("click", () => createTemplate("vendor"));
  }
}

async function createTemplate(type) {
  const status = document.getElementById("template-status");

  const templates = {
    journal: {
      sheetName: "JE_Template",
      headers: ["Date", "Memo", "Account", "Debit", "Credit"],
      rows: [
        ["06/10/2026", "Office Supplies", "6000", 100, 0],
        ["06/10/2026", "Office Supplies", "1000", 0, 100]
      ]
    },
    payroll: {
      sheetName: "Payroll_Template",
      headers: ["Employee", "Account", "Amount", "Date"],
      rows: [
        ["John Smith", "6200", 1500, "06/10/2026"],
        ["Jane Doe", "6200", 1750, "06/10/2026"]
      ]
    },
    vendor: {
      sheetName: "Vendor_Bill_Template",
      headers: ["Vendor", "Account", "Amount", "Bill Date"],
      rows: [
        ["ABC Supplies", "6000", 250, "06/10/2026"],
        ["XYZ Consulting", "6100", 500, "06/10/2026"]
      ]
    }
  };

  const template = templates[type];

  try {
    if (status) status.textContent = "Creating template…";

    await Excel.run(async (ctx) => {
      let sheet = ctx.workbook.worksheets.getItemOrNullObject(template.sheetName);
      sheet.load("name");
      await ctx.sync();

      if (sheet.isNullObject) {
        sheet = ctx.workbook.worksheets.add(template.sheetName);
      } else {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("address");
        await ctx.sync();

        if (!usedRange.isNullObject) {
          usedRange.clear();
        }
      }

      const endCol = colLetter(template.headers.length);

      const headerRange = sheet.getRange(`A1:${endCol}1`);
      headerRange.values = [template.headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#1A1F2E";
      headerRange.format.font.color = "#FFFFFF";

      const dataRange = sheet.getRange(`A2:${endCol}${template.rows.length + 1}`);
      dataRange.values = template.rows;

      sheet.getUsedRange().format.autofitColumns();
      sheet.activate();

      await ctx.sync();
    });

    if (status) status.textContent = `✓ ${template.sheetName} created.`;
  } catch (err) {
    if (status) status.textContent = `Error: ${err.message}`;
  }
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
     const alreadyPosted = await hasPostedRows(rangeAddr);

if (alreadyPosted) {
  appendLog("err", "Some selected rows are already marked as Posted. Posting blocked.");
  return;
}
    const rows = await readRangeValues(rangeAddr);
    appendLog("info", `${rows.length} row(s) to post as ${txnType}…`);

    const chunkSize = 10;

for (let i = 0; i < rows.length; i += chunkSize) {
  const chunk = rows.slice(i, i + chunkSize);

  const res = await fetch(`${API_BASE_URL}/api/journals`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      platform: state.platform,
      transactionType: txnType,
      entries: chunk
    })
  });

  const payload = await res.json();

  if (!res.ok) {
    throw new Error(payload.error || "Posting failed");
  }

  const pct = Math.round(((i + chunk.length) / rows.length) * 100);
  bar.style.width = `${pct}%`;
  appendLog("ok", `Posted ${payload.postedCount} row(s) to ${payload.platform} (${pct}%)`);
}

    appendLog("ok", `✓ All ${rows.length} entries posted successfully.`);
    bar.style.width = "100%";

    await markRowsPosted(rangeAddr, rows.length, true);
     await writeAuditLog({
  platform: state.platform,
  transactionType: txnType,
  rowsPosted: rows.length,
  status: "Success"
});
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

    const rows = range.values.filter((r) => r.some((c) => c !== ""));

    // If first row looks like headers, skip it
    const firstRow = rows[0] || [];
    const firstCell = String(firstRow[0] || "").toLowerCase();

    if (firstCell === "date") {
      return rows.slice(1);
    }

    return rows;
  });
}

async function markRowsPosted(rangeAddr, rowCount, skipHeader = false) {
  await Excel.run(async (ctx) => {
    const { sheetName, address } = parseRangeAddress(rangeAddr);
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.load("columnIndex, rowIndex, columnCount");
    await ctx.sync();

    const statusCol = range.columnIndex + range.columnCount;
    const startRow = skipHeader ? range.rowIndex + 1 : range.rowIndex;

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
async function writeAuditLog({ platform, transactionType, rowsPosted, status }) {
  await Excel.run(async (ctx) => {
    let sheet = ctx.workbook.worksheets.getItemOrNullObject("LedgerFlow_Log");
    sheet.load("name");
    await ctx.sync();

    if (sheet.isNullObject) {
      sheet = ctx.workbook.worksheets.add("LedgerFlow_Log");

      const headerRange = sheet.getRange("A1:E1");
      headerRange.values = [["Timestamp", "Platform", "Transaction Type", "Rows Posted", "Status"]];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#1A1F2E";
      headerRange.format.font.color = "#FFFFFF";
    }

    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("rowCount");
    await ctx.sync();

    const nextRow = usedRange.isNullObject ? 1 : usedRange.rowCount;

    const logRange = sheet.getRangeByIndexes(nextRow, 0, 1, 5);
    logRange.values = [[
      new Date().toLocaleString(),
      platform,
      transactionType,
      rowsPosted,
      status
    ]];

    sheet.getUsedRange().format.autofitColumns();

    await ctx.sync();
  });
}
async function hasPostedRows(rangeAddr) {
  return await Excel.run(async (ctx) => {
    const { sheetName, address } = parseRangeAddress(rangeAddr);
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.load("rowIndex, columnIndex, rowCount, columnCount");
    await ctx.sync();

    const statusCol = range.columnIndex + range.columnCount;
    const startRow = range.rowIndex + 1;
    const dataRowCount = Math.max(range.rowCount - 1, 0);

    if (dataRowCount === 0) return false;

    const statusRange = sheet.getRangeByIndexes(startRow, statusCol, dataRowCount, 1);
    statusRange.load("values");
    await ctx.sync();

    return statusRange.values.some((row) =>
      String(row[0] || "").toLowerCase() === "posted"
    );
  });
}
async function clearValidationHighlights(rangeAddr) {
  await Excel.run(async (ctx) => {
    const { sheetName, address } = parseRangeAddress(rangeAddr);
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.format.fill.clear();

    await ctx.sync();
  });
}
async function highlightValidationErrors(rangeAddr) {
  await Excel.run(async (ctx) => {
    const { sheetName, address } = parseRangeAddress(rangeAddr);
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);

    range.load("values, rowIndex, columnIndex");
    await ctx.sync();

    const rows = range.values;
    if (rows.length < 2) return;

    const headers = rows[0].map((h) => String(h || "").trim().toLowerCase());

    const idx = {
      date: headers.indexOf("date"),
      account: headers.indexOf("account"),
      debit: headers.indexOf("debit"),
      credit: headers.indexOf("credit"),
    };

    const errorColor = "#F8D7DA";

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];

      const date = row[idx.date];
      const account = row[idx.account];
      const debitRaw = row[idx.debit];
      const creditRaw = row[idx.credit];

      if (!date && idx.date >= 0) {
        sheet.getCell(range.rowIndex + i, range.columnIndex + idx.date).format.fill.color = errorColor;
      }

      if (!account && idx.account >= 0) {
        sheet.getCell(range.rowIndex + i, range.columnIndex + idx.account).format.fill.color = errorColor;
      }

      if (idx.debit >= 0 && debitRaw !== "" && debitRaw !== null && isNaN(parseFloat(debitRaw))) {
        sheet.getCell(range.rowIndex + i, range.columnIndex + idx.debit).format.fill.color = errorColor;
      }

      if (idx.credit >= 0 && creditRaw !== "" && creditRaw !== null && isNaN(parseFloat(creditRaw))) {
        sheet.getCell(range.rowIndex + i, range.columnIndex + idx.credit).format.fill.color = errorColor;
      }

      const debit = parseFloat(debitRaw) || 0;
      const credit = parseFloat(creditRaw) || 0;

      if (debit > 0 && credit > 0) {
        if (idx.debit >= 0) sheet.getCell(range.rowIndex + i, range.columnIndex + idx.debit).format.fill.color = errorColor;
        if (idx.credit >= 0) sheet.getCell(range.rowIndex + i, range.columnIndex + idx.credit).format.fill.color = errorColor;
      }

      if (debit === 0 && credit === 0) {
        if (idx.debit >= 0) sheet.getCell(range.rowIndex + i, range.columnIndex + idx.debit).format.fill.color = errorColor;
        if (idx.credit >= 0) sheet.getCell(range.rowIndex + i, range.columnIndex + idx.credit).format.fill.color = errorColor;
      }
    }

    await ctx.sync();
  });
}
async function getMappedAccounts() {
  return await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject("LedgerFlow_Mappings");

    sheet.load("name");
    await ctx.sync();

    if (sheet.isNullObject) {
      return [];
    }

    const range = sheet.getUsedRange();
    range.load("values");
    await ctx.sync();

    return range.values
      .slice(1)
      .map((row) => String(row[0] || "").trim())
      .filter(Boolean);
  });
}
function updateConnectionManager() {
  const qboStatus = document.getElementById("conn-qbo-status");
  const qboCompany = document.getElementById("conn-qbo-company");
  const xeroStatus = document.getElementById("conn-xero-status");
  const xeroCompany = document.getElementById("conn-xero-company");
  const lastSync = document.getElementById("conn-last-sync");

  if (!qboStatus || !xeroStatus) return;

  qboStatus.textContent = "Disconnected";
  qboCompany.textContent = "—";
  xeroStatus.textContent = "Disconnected";
  xeroCompany.textContent = "—";

  if (state.connected && state.platform === "qbo") {
    qboStatus.textContent = "Connected";
    qboCompany.textContent = state.companyName || "Acme Corp (QBO)";
  }

  if (state.connected && state.platform === "xero") {
    xeroStatus.textContent = "Connected";
    xeroCompany.textContent = state.companyName || "Acme Corp (Xero)";
  }

  if (lastSync) {
    lastSync.textContent = `Last sync: ${new Date().toLocaleString()}`;
  }
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
