const TIMELINE_SHEET_NAME = "TIMELINE";
const TIMELINE_REQUIRED_HEADERS = ["Title", "Start", "End", "Increment"];
const DATA_REQUIRED_HEADERS = ["Start Yr", "Event", "Paradigm"];

function toInteger(value) {
  if (value == null || String(value).trim() === "") {
    return null;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) && Number.isInteger(parsed) ? parsed : null;
}

function cleanParadigm(value) {
  return String(value == null ? "" : value).trim();
}

export function normalizeParadigmKey(value) {
  return cleanParadigm(value).toLowerCase();
}

function normalizeHeader(value) {
  return String(value == null ? "" : value).trim().replace(/\s+/g, " ").toLowerCase();
}

function isBlankCell(value) {
  return value == null || String(value).trim() === "";
}

function readSheetTable(workbook, sheetName, requiredHeaders) {
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`Workbook is missing sheet "${sheetName}".`);
  }

  const table = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: true,
    blankrows: false,
    defval: null
  });
  if (!table.length) {
    throw new Error(`Sheet "${sheetName}" is empty.`);
  }

  const headerIndexes = new Map();
  const headerRow = table[0] || [];
  for (let index = 0; index < headerRow.length; index += 1) {
    const normalized = normalizeHeader(headerRow[index]);
    if (normalized && !headerIndexes.has(normalized)) {
      headerIndexes.set(normalized, index);
    }
  }

  const missing = requiredHeaders.filter((header) => !headerIndexes.has(normalizeHeader(header)));
  if (missing.length) {
    throw new Error(`Sheet "${sheetName}" is missing required columns: ${missing.join(", ")}.`);
  }

  return {
    rows: table.slice(1),
    columnIndex: (header) => headerIndexes.get(normalizeHeader(header))
  };
}

function parseTimelineSheet(workbook) {
  const table = readSheetTable(workbook, TIMELINE_SHEET_NAME, TIMELINE_REQUIRED_HEADERS);
  const titleIndex = table.columnIndex("Title");
  const startIndex = table.columnIndex("Start");
  const endIndex = table.columnIndex("End");
  const incrementIndex = table.columnIndex("Increment");
  const populatedRows = table.rows.filter((row) => (
    !isBlankCell(row[titleIndex]) ||
    !isBlankCell(row[startIndex]) ||
    !isBlankCell(row[endIndex]) ||
    !isBlankCell(row[incrementIndex])
  ));

  if (populatedRows.length !== 1) {
    throw new Error('Sheet "TIMELINE" must contain exactly one populated configuration row.');
  }

  const row = populatedRows[0];
  const title = String(row[titleIndex] == null ? "" : row[titleIndex]).trim();
  const startYear = toInteger(row[startIndex]);
  const endYear = toInteger(row[endIndex]);
  const increment = toInteger(row[incrementIndex]);
  if (!title) {
    throw new Error("TIMELINE Title must not be empty.");
  }
  if (startYear === null || endYear === null || increment === null) {
    throw new Error("TIMELINE Start, End, and Increment must be finite integers.");
  }
  if (startYear >= endYear) {
    throw new Error("TIMELINE Start must be earlier than End.");
  }
  if (increment <= 0) {
    throw new Error("TIMELINE Increment must be a positive integer.");
  }

  return { title, startYear, endYear, increment };
}

function parseDataSheet(workbook, sheetName) {
  const table = readSheetTable(workbook, sheetName, DATA_REQUIRED_HEADERS);
  const yearIndex = table.columnIndex("Start Yr");
  const eventIndex = table.columnIndex("Event");
  const paradigmIndex = table.columnIndex("Paradigm");
  const events = [];

  for (let index = 0; index < table.rows.length; index += 1) {
    const row = table.rows[index] || [];
    const rawYear = row[yearIndex];
    const event = String(row[eventIndex] == null ? "" : row[eventIndex]).trim();
    const paradigm = cleanParadigm(row[paradigmIndex]);
    if (isBlankCell(rawYear) && !event && !paradigm) {
      continue;
    }

    const startYear = toInteger(rawYear);
    if (startYear === null) {
      throw new Error(`Sheet "${sheetName}" has an invalid Start Yr near data row ${index + 2}.`);
    }
    if (!event) {
      throw new Error(`Sheet "${sheetName}" has an empty Event near data row ${index + 2}.`);
    }

    events.push({ startYear, event, paradigm });
  }

  if (!events.length) {
    throw new Error(`Sheet "${sheetName}" has no usable event rows.`);
  }
  return events;
}

export function parseWorkbookData(workbook) {
  if (!Array.isArray(workbook.SheetNames) || !workbook.SheetNames.includes(TIMELINE_SHEET_NAME)) {
    throw new Error('Workbook must contain a sheet named "TIMELINE".');
  }

  const timeline = parseTimelineSheet(workbook);
  const tabs = new Map();
  for (const sheetName of workbook.SheetNames) {
    if (sheetName !== TIMELINE_SHEET_NAME) {
      tabs.set(sheetName, parseDataSheet(workbook, sheetName));
    }
  }

  if (!tabs.size) {
    throw new Error("Workbook does not contain any data sheets.");
  }
  return { timeline, tabs };
}
