const REQUIRED_COLUMNS = [
  "Order ID",
  "都道府県",
  "品番",
  "数量",
  "ご住所",
  "お電話番号",
  "受取人様名",
  "備考＜内容確認＞",
  "メーカー名",
];

const DETAIL_START_ROW = 21;
const DETAIL_END_ROW = 28;
const TEMPLATE_URL = "./assets/asplund_template.xlsx";

const runBtn = document.getElementById("runBtn");
const csvFileInput = document.getElementById("csvFile");
const manufacturerInput = document.getElementById("manufacturer");
const summary = document.getElementById("summary");
const generatedList = document.getElementById("generatedList");
const issuesBody = document.getElementById("issuesBody");

runBtn.addEventListener("click", onRun);

function normalizeHeader(name) {
  return String(name || "").replace(/\ufeff/g, "").trim();
}

function normalizeValue(value) {
  return String(value ?? "").trim();
}

function toNumber(text) {
  const cleaned = normalizeValue(text).replace(/,/g, "");
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function normalizePhoneNumber(rawPhone) {
  const normalized = normalizeValue(rawPhone);
  if (!normalized) return "";
  let digits = normalized.replace(/\D/g, "");
  if (!digits) return "";

  if (normalized.startsWith("+81") && digits.startsWith("81")) {
    digits = digits.slice(2);
  } else if (normalized.startsWith("0081") && digits.startsWith("0081")) {
    digits = digits.slice(4);
  } else if (digits.startsWith("81") && !digits.startsWith("0") && digits.length >= 10) {
    digits = digits.slice(2);
  }

  if (digits && !digits.startsWith("0")) digits = `0${digits}`;
  return digits;
}

function buildNote(orderId, noteText) {
  const n = normalizeValue(noteText);
  if (!n) return `弊社注文番号：${orderId}`;
  return `弊社注文番号：${orderId}\n${n}着でお願いいたします。`;
}

function safeOrderId(orderId) {
  return normalizeValue(orderId).replace("#", "").replace(/[^0-9A-Za-z_-]/g, "_");
}

function setCellValue(ws, addr, value) {
  const cell = ws[addr] || {};
  delete cell.f;
  if (value === null || value === undefined || value === "") {
    cell.t = "s";
    cell.v = "";
    ws[addr] = cell;
    return;
  }
  if (typeof value === "number") {
    cell.t = "n";
    cell.v = value;
  } else {
    cell.t = "s";
    cell.v = String(value);
  }
  ws[addr] = cell;
}

function setCellFormula(ws, addr, formula) {
  const cell = ws[addr] || {};
  cell.f = formula;
  delete cell.v;
  delete cell.w;
  ws[addr] = cell;
}

function renderIssues(issues) {
  if (!issues.length) {
    issuesBody.innerHTML = `<tr><td colspan="3">No issues.</td></tr>`;
    return;
  }
  issuesBody.innerHTML = issues
    .map(
      (x) =>
        `<tr><td>${x.csvRow}</td><td>${escapeHtml(x.orderId)}</td><td>${escapeHtml(x.message)}</td></tr>`
    )
    .join("");
}

function renderGenerated(files) {
  if (!files.length) {
    generatedList.innerHTML = "<li>No files generated.</li>";
    return;
  }
  generatedList.innerHTML = files.map((f) => `<li>${escapeHtml(f)}</li>`).join("");
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function parseCsv(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const fields = (results.meta.fields || []).map(normalizeHeader);
        const rows = (results.data || []).map((obj) => {
          const row = {};
          Object.entries(obj || {}).forEach(([k, v]) => {
            row[normalizeHeader(k)] = normalizeValue(v);
          });
          return row;
        });
        resolve({ fields, rows });
      },
      error: reject,
    });
  });
}

function extractMasterData(templateWb) {
  const listSheet = templateWb.Sheets["リスト"];
  const shipSheet = templateWb.Sheets["送料対応表"];
  if (!listSheet || !shipSheet) throw new Error("Template missing required sheets");

  const listRows = XLSX.utils.sheet_to_json(listSheet, { header: 1, raw: true, defval: "" });
  const shipRows = XLSX.utils.sheet_to_json(shipSheet, { header: 1, raw: true, defval: "" });

  const validCodes = new Set();
  for (let i = 1; i < listRows.length; i += 1) {
    const raw = listRows[i][0];
    if (raw === null || raw === undefined || raw === "") continue;
    const code = normalizeValue(Number.isInteger(raw) ? raw : (Number.isFinite(raw) ? Math.trunc(raw) : raw));
    validCodes.add(code);
  }

  const validPrefs = new Set();
  for (let i = 1; i < shipRows.length; i += 1) {
    const pref = normalizeValue(shipRows[i][0]);
    if (pref) validPrefs.add(pref);
  }
  return { validCodes, validPrefs };
}

function validateRows(rows, fields, validCodes, validPrefs, manufacturer) {
  const issues = [];
  const groups = new Map();

  const missingColumns = REQUIRED_COLUMNS.filter((c) => !fields.includes(c));
  if (missingColumns.length) {
    issues.push({ csvRow: 1, orderId: "-", message: `Missing required columns: ${missingColumns.join(", ")}` });
    return { groups, issues };
  }

  rows.forEach((row, idx) => {
    const csvRow = idx + 2;
    if (normalizeValue(row["メーカー名"]) !== manufacturer) return;

    const orderId = normalizeValue(row["Order ID"]);
    if (!orderId) {
      issues.push({ csvRow, orderId: "-", message: "Order ID is empty" });
      return;
    }

    const code = normalizeValue(row["品番"]);
    if (!validCodes.has(code)) {
      issues.push({ csvRow, orderId, message: `Unknown 品番: ${code}` });
      return;
    }

    const pref = normalizeValue(row["都道府県"]);
    if (!validPrefs.has(pref)) {
      issues.push({ csvRow, orderId, message: `Unknown 都道府県: ${pref}` });
      return;
    }

    const qty = toNumber(row["数量"]);
    if (qty === null || qty <= 0) {
      issues.push({ csvRow, orderId, message: `Invalid 数量: ${normalizeValue(row["数量"])}` });
      return;
    }

    if (!groups.has(orderId)) groups.set(orderId, []);
    groups.get(orderId).push({ csvRow, row });
  });

  const invalidOrders = new Set();
  for (const [orderId, items] of groups.entries()) {
    if (items.length > DETAIL_END_ROW - DETAIL_START_ROW + 1) {
      issues.push({
        csvRow: items[0].csvRow,
        orderId,
        message: `Too many detail lines: ${items.length} (max 8)`,
      });
      invalidOrders.add(orderId);
      continue;
    }

    const keys = ["都道府県", "ご住所", "お電話番号", "受取人様名", "備考＜内容確認＞"];
    const baseline = {};
    keys.forEach((k) => { baseline[k] = normalizeValue(items[0].row[k]); });

    for (let i = 1; i < items.length; i += 1) {
      const item = items[i];
      for (const k of keys) {
        if (normalizeValue(item.row[k]) !== baseline[k]) {
          issues.push({ csvRow: item.csvRow, orderId, message: `Inconsistent ${k} in the same order` });
          invalidOrders.add(orderId);
          break;
        }
      }
    }
  }

  invalidOrders.forEach((orderId) => groups.delete(orderId));
  return { groups, issues };
}

function applyOrderToWorkbook(wb, orderId, items) {
  const ws = wb.Sheets["発注書フォーマット"];
  if (!ws) throw new Error("Template missing 発注書フォーマット sheet");
  const first = items[0].row;

  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");

  setCellValue(ws, "I3", `${yyyy}-${mm}-${dd}`);
  setCellValue(ws, "D17", normalizeValue(first["都道府県"]));
  setCellValue(ws, "C35", normalizeValue(first["ご住所"]));
  setCellValue(ws, "C36", normalizePhoneNumber(first["お電話番号"]));
  setCellValue(ws, "C37", normalizeValue(first["受取人様名"]));
  setCellValue(ws, "G40", buildNote(orderId, first["備考＜内容確認＞"]));

  for (let rowNo = DETAIL_START_ROW; rowNo <= DETAIL_END_ROW; rowNo += 1) {
    setCellFormula(ws, `C${rowNo}`, `IFERROR(VLOOKUP(A${rowNo},'リスト'!A:T,3,0),"")`);
    setCellFormula(ws, `G${rowNo}`, `IFERROR(VLOOKUP(A${rowNo},'リスト'!A:T,5,0),"")`);
    setCellFormula(ws, `I${rowNo}`, `IFERROR(VLOOKUP(A${rowNo},'リスト'!A:T,8,0),"")`);
    setCellFormula(ws, `J${rowNo}`, `IFERROR(H${rowNo}*I${rowNo},"")`);
    setCellValue(ws, `A${rowNo}`, "");
    setCellValue(ws, `H${rowNo}`, "");
  }

  items.forEach((item, idx) => {
    const rowNo = DETAIL_START_ROW + idx;
    const qty = toNumber(item.row["数量"]) || 0;
    setCellValue(ws, `A${rowNo}`, normalizeValue(item.row["品番"]));
    setCellValue(ws, `H${rowNo}`, Number.isInteger(qty) ? qty : qty);
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function onRun() {
  const file = csvFileInput.files?.[0];
  const manufacturer = normalizeValue(manufacturerInput.value || "ASPLUND");
  if (!file) {
    alert("CSVファイルを選択してください。");
    return;
  }
  if (!manufacturer) {
    alert("メーカー名を入力してください。");
    return;
  }

  runBtn.disabled = true;
  summary.textContent = "Processing...";
  generatedList.innerHTML = "";
  issuesBody.innerHTML = `<tr><td colspan="3">Processing...</td></tr>`;

  try {
    const [{ fields, rows }, templateResp] = await Promise.all([
      parseCsv(file),
      fetch(TEMPLATE_URL),
    ]);
    if (!templateResp.ok) throw new Error(`Template fetch failed: ${templateResp.status}`);

    const templateBuffer = await templateResp.arrayBuffer();
    const templateWb = XLSX.read(templateBuffer, { type: "array", cellFormula: true, cellStyles: true });
    const { validCodes, validPrefs } = extractMasterData(templateWb);

    const { groups, issues } = validateRows(rows, fields, validCodes, validPrefs, manufacturer);
    renderIssues(issues);

    const zip = new JSZip();
    const generatedFiles = [];

    const now = new Date();
    const ymd = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
    for (const [orderId, items] of groups.entries()) {
      const wb = XLSX.read(templateBuffer.slice(0), { type: "array", cellFormula: true, cellStyles: true });
      applyOrderToWorkbook(wb, orderId, items);

      const safeId = safeOrderId(orderId);
      const fileName = `発注書_${manufacturer}_${safeId}_${ymd}.xlsx`;
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      zip.file(fileName, wbout);
      generatedFiles.push(fileName);
    }

    renderGenerated(generatedFiles);
    summary.textContent = `Generated: ${generatedFiles.length} / Issues: ${issues.length}`;

    if (generatedFiles.length > 0) {
      const zipBlob = await zip.generateAsync({ type: "blob" });
      downloadBlob(zipBlob, `asplund_orders_${ymd}.zip`);
    }
  } catch (err) {
    console.error(err);
    summary.textContent = `Error: ${err.message || err}`;
    issuesBody.innerHTML = `<tr><td colspan="3">${escapeHtml(err.message || String(err))}</td></tr>`;
  } finally {
    runBtn.disabled = false;
  }
}
