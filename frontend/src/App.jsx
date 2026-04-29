import { useCallback, useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";

const apiBase = "/api";

function stripAnsi(value) {
  if (!value) return "";
  return value.replace(/\u001b\[[0-9;]*m/g, "");
}

function extractDeviceCodePrompt(output) {
  const cleanOutput = stripAnsi(output);
  if (!cleanOutput) return null;

  const sentenceMatch = cleanOutput.match(
    /To sign in,\s*use a web browser to open the page\s*(https:\/\/[^\s]+)\s*and enter the code\s*([A-Z0-9-]+)\s*to authenticate\.?/i
  );
  if (sentenceMatch) {
    return {
      url: sentenceMatch[1],
      code: sentenceMatch[2]
    };
  }

  const urlMatch = cleanOutput.match(
    /https:\/\/(?:microsoft\.com\/devicelogin|login\.microsoft(?:online)?\.com\/device|login\.microsoftonline\.com\/common\/oauth2\/deviceauth)[^\s]*/i
  );
  const codeMatch = cleanOutput.match(/\b[A-Z0-9]{4,}(?:-[A-Z0-9]{4,})+\b|\b[A-Z0-9]{8,10}\b/);

  if (!urlMatch && !codeMatch) {
    return null;
  }

  return {
    url: urlMatch?.[0] || "https://microsoft.com/devicelogin",
    code: codeMatch?.[0] || ""
  };
}

async function parseApiResponse(response) {
  const contentType = response.headers.get("content-type") || "";

  if (contentType.includes("application/json")) {
    return response.json();
  }

  const text = await response.text();
  const compactText = text.replace(/\s+/g, " ").trim();
  throw new Error(
    `API returned ${response.status} ${response.statusText}: ${compactText.slice(0, 180) || "Empty response"}`
  );
}

function normalizeDefaults(fields) {
  return fields.reduce((acc, field) => {
    acc[field.id] = field.defaultValue ?? (field.type === "checkbox" ? false : "");
    return acc;
  }, {});
}

function formatDate(value) {
  if (!value) return "Pending";
  return new Date(value).toLocaleString();
}

function formatDuration(value) {
  if (!value && value !== 0) return "Pending";
  const totalSeconds = Math.max(0, Math.round(value / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}m ${String(seconds).padStart(2, "0")}s`;
}

function formatRunTenant(run) {
  const tenant = run?.tenantHint || run?.payload?.tenantId;
  return tenant ? String(tenant).trim() : "Auto-detect";
}

function buildRunHistoryQuery({ limit, offset, status, scriptId, tenantId, dateFrom, dateTo }) {
  const params = new URLSearchParams();
  params.set("limit", String(limit));
  params.set("offset", String(offset));
  if (status && status !== "all") params.set("status", status);
  if (scriptId && scriptId !== "all") params.set("scriptId", scriptId);
  if (tenantId) params.set("tenantId", tenantId);
  if (dateFrom) params.set("dateFrom", `${dateFrom}T00:00:00.000Z`);
  if (dateTo) params.set("dateTo", `${dateTo}T23:59:59.999Z`);
  return params.toString();
}

function formatRelativeTime(value, nowMs = Date.now()) {
  if (!value) return "Pending";

  const timestamp = typeof value === "number" ? value : new Date(value).getTime();
  if (!Number.isFinite(timestamp)) {
    return "Pending";
  }

  const diffSeconds = Math.max(0, Math.round((nowMs - timestamp) / 1000));
  if (diffSeconds < 5) return "just now";
  if (diffSeconds < 60) return `${diffSeconds}s ago`;

  const minutes = Math.floor(diffSeconds / 60);
  const seconds = diffSeconds % 60;
  if (minutes < 60) {
    return `${minutes}m ${String(seconds).padStart(2, "0")}s ago`;
  }

  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;
  return `${hours}h ${String(remainingMinutes).padStart(2, "0")}m ago`;
}

function getLiveDurationMs(run, nowMs = Date.now()) {
  if (!run) {
    return null;
  }

  if (run.durationMs !== null && run.durationMs !== undefined && !["running", "queued", "canceling"].includes(run.status)) {
    return run.durationMs;
  }

  const anchor = run.startedAt || run.requestedAt;
  if (!anchor) {
    return run.durationMs ?? null;
  }

  return Math.max(0, nowMs - new Date(anchor).getTime());
}

function formatEstimatedRuntime(minutes) {
  if (!minutes) {
    return "Varies by tenant size";
  }

  return minutes === 1 ? "Usually about 1 minute" : `Usually about ${minutes} minutes`;
}

function formatOutputText(value, fallback) {
  const clean = stripAnsi(value || "").trim();
  return clean || fallback;
}

function formatFileSize(value) {
  const size = Number(value);
  if (!Number.isFinite(size) || size <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB", "TB"];
  const exponent = Math.min(Math.floor(Math.log(size) / Math.log(1024)), units.length - 1);
  const formatted = size / 1024 ** exponent;
  const digits = formatted >= 10 || exponent === 0 ? 0 : 1;
  return `${formatted.toFixed(digits)} ${units[exponent]}`;
}

function formatDateInputDisplay(value) {
  if (!value) {
    return "mm/dd/yyyy";
  }

  const [year, month, day] = String(value).split("-");
  if (!year || !month || !day) {
    return value;
  }

  return `${month}/${day}/${year}`;
}

function parseInputDateValue(value) {
  if (!value) {
    return null;
  }

  const [year, month, day] = String(value).split("-").map(Number);
  if (!year || !month || !day) {
    return null;
  }

  return new Date(year, month - 1, day);
}

function toInputDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonths(date, delta) {
  return new Date(date.getFullYear(), date.getMonth() + delta, 1);
}

function sameDay(left, right) {
  return (
    left?.getFullYear?.() === right?.getFullYear?.() &&
    left?.getMonth?.() === right?.getMonth?.() &&
    left?.getDate?.() === right?.getDate?.()
  );
}

function buildCalendarDays(monthDate) {
  const monthStart = startOfMonth(monthDate);
  const gridStart = new Date(monthStart);
  gridStart.setDate(monthStart.getDate() - monthStart.getDay());

  return Array.from({ length: 42 }, (_, index) => {
    const date = new Date(gridStart);
    date.setDate(gridStart.getDate() + index);
    return date;
  });
}

function formatCalendarHeading(date) {
  return date.toLocaleDateString(undefined, {
    month: "long",
    year: "numeric"
  });
}

function useFloatingLayer(open, anchorRef, panelRef, { matchWidth = false, minWidth = 0, estimatedHeight = 320 } = {}) {
  const [style, setStyle] = useState(null);

  useEffect(() => {
    if (!open) {
      setStyle(null);
      return undefined;
    }

    const updatePosition = () => {
      const anchorRect = anchorRef.current?.getBoundingClientRect();
      const panelRect = panelRef.current?.getBoundingClientRect();

      if (!anchorRect) {
        return;
      }

      const viewportPadding = 12;
      const desiredWidth = matchWidth
        ? anchorRect.width
        : Math.max(minWidth, anchorRect.width, panelRect?.width || 0);
      const desiredHeight = panelRect?.height || 0;
      const availableBelow = Math.max(0, Math.floor(window.innerHeight - anchorRect.bottom - viewportPadding));
      const availableAbove = Math.max(0, Math.floor(anchorRect.top - viewportPadding));
      const preferredHeight = desiredHeight || estimatedHeight;
      const placeBelow = availableBelow >= preferredHeight;
      const maxHeight = placeBelow ? availableBelow : availableAbove;

      let left = anchorRect.left;
      if (left + desiredWidth > window.innerWidth - viewportPadding) {
        left = Math.max(viewportPadding, window.innerWidth - viewportPadding - desiredWidth);
      }

      let top = anchorRect.bottom;
      if (!placeBelow) {
        const renderHeight = Math.min(desiredHeight || estimatedHeight, maxHeight || preferredHeight);
        top = Math.max(viewportPadding, anchorRect.top - renderHeight);
      }

      setStyle({
        position: "fixed",
        top: `${Math.round(top)}px`,
        left: `${Math.round(left)}px`,
        width: matchWidth ? `${Math.round(anchorRect.width)}px` : undefined,
        minWidth: `${Math.round(Math.max(minWidth, anchorRect.width))}px`,
        maxHeight: `${Math.round(maxHeight)}px`
      });
    };

    updatePosition();
    window.addEventListener("resize", updatePosition);
    window.addEventListener("scroll", updatePosition, true);

    return () => {
      window.removeEventListener("resize", updatePosition);
      window.removeEventListener("scroll", updatePosition, true);
    };
  }, [anchorRef, estimatedHeight, matchWidth, minWidth, open, panelRef]);

  return style;
}

async function copyText(text) {
  if (!text) return false;

  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // Fall back to document.execCommand below.
  }

  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    textarea.style.pointerEvents = "none";
    document.body.appendChild(textarea);
    textarea.select();
    textarea.setSelectionRange(0, textarea.value.length);
    const copied = document.execCommand("copy");
    document.body.removeChild(textarea);
    return copied;
  } catch {
    return false;
  }
}

function groupScriptsByCategory(scripts) {
  return scripts.reduce((groups, script) => {
    const category = script.category || "Other";
    if (!groups[category]) {
      groups[category] = [];
    }
    groups[category].push(script);
    return groups;
  }, {});
}

function parseUserList(value) {
  const entries = String(value || "")
    .split(/[\n,]+/)
    .map((item) => item.trim())
    .filter(Boolean);
  const unique = [...new Set(entries)];
  return {
    entries,
    unique,
    duplicates: entries.filter((item, index) => entries.indexOf(item) !== index)
  };
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function sanitizeFileName(value) {
  return String(value || "report")
    .replace(/[^A-Za-z0-9]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "") || "report";
}

function sanitizeWorksheetName(value, usedNames) {
  const normalized = String(value || "Sheet")
    .replace(/[\[\]\\/*?:]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 31) || "Sheet";

  let candidate = normalized;
  let suffix = 2;
  while (usedNames.has(candidate)) {
    const base = normalized.slice(0, Math.max(1, 31 - String(suffix).length - 1)).trim();
    candidate = `${base}-${suffix}`;
    suffix += 1;
  }

  usedNames.add(candidate);
  return candidate;
}

function stringifyReportValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (Array.isArray(value)) {
    return value.map((entry) => stringifyReportValue(entry)).filter(Boolean).join("\n");
  }

  if (typeof value === "boolean") {
    return value ? "Yes" : "No";
  }

  if (typeof value === "object") {
    return JSON.stringify(value);
  }

  return String(value);
}

function isSpreadsheetNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return true;
  }

  const text = stringifyReportValue(value).trim();
  return Boolean(text) && /^-?(?:\d+|\d*\.\d+)$/.test(text) && !/^0\d+/.test(text);
}

function getColumnLabel(columnIndex) {
  let value = columnIndex + 1;
  let label = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function createWorksheetCellXml(value, rowIndex, columnIndex, styleIndex = 0) {
  const cellReference = `${getColumnLabel(columnIndex)}${rowIndex}`;

  if (isSpreadsheetNumber(value)) {
    return `<c r="${cellReference}" s="${styleIndex}"><v>${escapeXml(stringifyReportValue(value))}</v></c>`;
  }

  return `<c r="${cellReference}" s="${styleIndex}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(stringifyReportValue(value))}</t></is></c>`;
}

function buildXlsxWorksheetXml(sheet) {
  const allRows = [];

  if (sheet.header?.length) {
    allRows.push(sheet.header.map((value) => ({ value, styleIndex: 1 })));
  }

  (sheet.rows || []).forEach((row) => {
    allRows.push(row.map((value) => ({ value, styleIndex: 0 })));
  });

  const rowXml = allRows
    .map((cells, rowIndex) => {
      const cellsXml = cells
        .map((cell, columnIndex) => createWorksheetCellXml(cell.value, rowIndex + 1, columnIndex, cell.styleIndex))
        .join("");

      return `<row r="${rowIndex + 1}">${cellsXml}</row>`;
    })
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowXml}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
}

function buildXlsxWorkbookXml(sheets) {
  const sheetXml = sheets
    .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetXml}</sheets>
</workbook>`;
}

function buildXlsxWorkbookRelsXml(sheets) {
  const sheetRels = sheets
    .map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`)
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetRels}
  <Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function buildXlsxStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment vertical="top" wrapText="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
}

function buildXlsxContentTypesXml(sheets) {
  const worksheetOverrides = sheets
    .map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${worksheetOverrides}
</Types>`;
}

function buildXlsxRootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function createCrc32Table() {
  const table = new Uint32Array(256);

  for (let index = 0; index < 256; index += 1) {
    let crc = index;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc & 1) ? (0xedb88320 ^ (crc >>> 1)) : (crc >>> 1);
    }
    table[index] = crc >>> 0;
  }

  return table;
}

const CRC32_TABLE = createCrc32Table();

function computeCrc32(bytes) {
  let crc = 0xffffffff;

  for (let index = 0; index < bytes.length; index += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }

  return (crc ^ 0xffffffff) >>> 0;
}

function toUtf8Bytes(value) {
  return new TextEncoder().encode(value);
}

function concatUint8Arrays(chunks) {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const combined = new Uint8Array(totalLength);
  let offset = 0;

  chunks.forEach((chunk) => {
    combined.set(chunk, offset);
    offset += chunk.length;
  });

  return combined;
}

function setUint16(view, offset, value) {
  view.setUint16(offset, value, true);
}

function setUint32(view, offset, value) {
  view.setUint32(offset, value >>> 0, true);
}

function getDosDateTime(value) {
  const date = value instanceof Date && !Number.isNaN(value.getTime()) ? value : new Date();
  const year = Math.max(1980, date.getFullYear());
  const dosDate = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  const dosTime = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);

  return { dosDate, dosTime };
}

function createStoredZip(files) {
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;

  files.forEach((file) => {
    const nameBytes = toUtf8Bytes(file.name);
    const dataBytes = file.data instanceof Uint8Array ? file.data : toUtf8Bytes(file.data);
    const { dosDate, dosTime } = getDosDateTime(file.lastModified);
    const crc32 = computeCrc32(dataBytes);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    setUint32(localView, 0, 0x04034b50);
    setUint16(localView, 4, 20);
    setUint16(localView, 6, 0);
    setUint16(localView, 8, 0);
    setUint16(localView, 10, dosTime);
    setUint16(localView, 12, dosDate);
    setUint32(localView, 14, crc32);
    setUint32(localView, 18, dataBytes.length);
    setUint32(localView, 22, dataBytes.length);
    setUint16(localView, 26, nameBytes.length);
    setUint16(localView, 28, 0);
    localHeader.set(nameBytes, 30);

    localChunks.push(localHeader, dataBytes);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    setUint32(centralView, 0, 0x02014b50);
    setUint16(centralView, 4, 20);
    setUint16(centralView, 6, 20);
    setUint16(centralView, 8, 0);
    setUint16(centralView, 10, 0);
    setUint16(centralView, 12, dosTime);
    setUint16(centralView, 14, dosDate);
    setUint32(centralView, 16, crc32);
    setUint32(centralView, 20, dataBytes.length);
    setUint32(centralView, 24, dataBytes.length);
    setUint16(centralView, 28, nameBytes.length);
    setUint16(centralView, 30, 0);
    setUint16(centralView, 32, 0);
    setUint16(centralView, 34, 0);
    setUint16(centralView, 36, 0);
    setUint32(centralView, 38, 0);
    setUint32(centralView, 42, offset);
    centralHeader.set(nameBytes, 46);

    centralChunks.push(centralHeader);
    offset += localHeader.length + dataBytes.length;
  });

  const centralDirectory = concatUint8Arrays(centralChunks);
  const endOfCentralDirectory = new Uint8Array(22);
  const endView = new DataView(endOfCentralDirectory.buffer);
  setUint32(endView, 0, 0x06054b50);
  setUint16(endView, 4, 0);
  setUint16(endView, 6, 0);
  setUint16(endView, 8, files.length);
  setUint16(endView, 10, files.length);
  setUint32(endView, 12, centralDirectory.length);
  setUint32(endView, 16, offset);
  setUint16(endView, 20, 0);

  return concatUint8Arrays([...localChunks, centralDirectory, endOfCentralDirectory]);
}

function extractReportDataFromHtml(html) {
  const marker = "const DATA =";
  const start = html.indexOf(marker);

  if (start < 0) {
    throw new Error("Unable to locate the report data in the HTML export.");
  }

  const jsonStart = html.indexOf("{", start);
  if (jsonStart < 0) {
    throw new Error("Unable to parse the embedded report data.");
  }

  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = jsonStart; index < html.length; index += 1) {
    const character = html[index];

    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (character === "\\") {
        escaped = true;
      } else if (character === "\"") {
        inString = false;
      }
      continue;
    }

    if (character === "\"") {
      inString = true;
      continue;
    }

    if (character === "{") {
      depth += 1;
    } else if (character === "}") {
      depth -= 1;
      if (depth === 0) {
        return JSON.parse(html.slice(jsonStart, index + 1));
      }
    }
  }

  throw new Error("The embedded report data appears to be incomplete.");
}

function buildSummaryRows(reportData) {
  const rows = [];

  if (reportData.title) rows.push(["Title", reportData.title]);
  if (reportData.tenant) rows.push(["Tenant", reportData.tenant]);
  if (reportData.subtitle) rows.push(["Subtitle", reportData.subtitle]);
  if (reportData.reportDate) rows.push(["Generated", reportData.reportDate]);

  if (Array.isArray(reportData.stripItems)) {
    reportData.stripItems.forEach((item) => {
      rows.push([item.label || "Strip Item", item.value ?? ""]);
    });
  }

  if (Array.isArray(reportData.kpis)) {
    reportData.kpis.forEach((item) => {
      rows.push([`KPI: ${item.label || "Value"}`, [item.value, item.sub].filter(Boolean).join(" | ")]);
    });
  }

  if (Array.isArray(reportData.services)) {
    rows.push(["Services", reportData.services.length]);
    if (reportData.totalItems !== undefined) rows.push(["Total Items", reportData.totalItems]);
    if (reportData.activeItems !== undefined) rows.push(["Active Items", reportData.activeItems]);
    if (reportData.totalGB !== undefined) rows.push(["Total Used (GB)", reportData.totalGB]);

    reportData.services.forEach((service) => {
      rows.push([`${service.Name} Items`, service.TotalItems ?? ""]);
      rows.push([`${service.Name} Active`, service.ActiveItems ?? ""]);
      rows.push([`${service.Name} Used (GB)`, service.TotalGB ?? ""]);
    });
  }

  return rows;
}

function buildWorkbookSheetsFromReport(reportData) {
  const usedNames = new Set();
  const sheets = [];
  const summaryRows = buildSummaryRows(reportData);

  sheets.push({
    name: sanitizeWorksheetName("Summary", usedNames),
    header: ["Field", "Value"],
    rows: summaryRows.length ? summaryRows : [["Report", reportData.title || "M365 Report"]]
  });

  if (Array.isArray(reportData.sections)) {
    reportData.sections.forEach((section, index) => {
      const sheetName = sanitizeWorksheetName(section.title || `Section ${index + 1}`, usedNames);
      if (Array.isArray(section.columns) && section.columns.length > 0) {
        const header = section.columns.map((column) => column.header || column.key || "Value");
        const rows = Array.isArray(section.rows)
          ? section.rows.map((row) =>
            section.columns.map((column) => stringifyReportValue(row?.[column.key]))
          )
          : [];

        sheets.push({
          name: sheetName,
          header,
          rows: rows.length ? rows : [header.map(() => "")]
        });
        return;
      }

      sheets.push({
        name: sheetName,
        header: ["Details"],
        rows: [[section.text || "No tabular data in this section."]]
      });
    });
  }

  if (Array.isArray(reportData.services)) {
    reportData.services.forEach((service, index) => {
      const sheetName = sanitizeWorksheetName(service.Name || `Service ${index + 1}`, usedNames);
      const rows = Array.isArray(service.Rows)
        ? service.Rows.map((row) => [
          row.DisplayName ?? "",
          row.Principal ?? "",
          row.UsedGB ?? "",
          row.Url ?? ""
        ])
        : [];

      sheets.push({
        name: sheetName,
        header: ["Name", "Principal", "Used (GB)", "URL"],
        rows: rows.length ? rows : [["", "", "", ""]]
      });
    });
  }

  return sheets;
}

function createExcelExportBlob(html) {
  const reportData = extractReportDataFromHtml(html);
  const sheets = buildWorkbookSheetsFromReport(reportData);
  const workbookFiles = [
    {
      name: "[Content_Types].xml",
      data: buildXlsxContentTypesXml(sheets)
    },
    {
      name: "_rels/.rels",
      data: buildXlsxRootRelsXml()
    },
    {
      name: "xl/workbook.xml",
      data: buildXlsxWorkbookXml(sheets)
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      data: buildXlsxWorkbookRelsXml(sheets)
    },
    {
      name: "xl/styles.xml",
      data: buildXlsxStylesXml()
    },
    ...sheets.map((sheet, index) => ({
      name: `xl/worksheets/sheet${index + 1}.xml`,
      data: buildXlsxWorksheetXml(sheet)
    }))
  ];

  return new Blob([createStoredZip(workbookFiles)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function parseReportTimestamp(value) {
  const normalized = String(value || "").trim();
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);

  if (match) {
    const [, year, month, day, hour, minute, second = "00"] = match;
    return {
      datePart: `${year}-${month}-${day}`,
      timePart: `${hour}-${minute}-${second}`
    };
  }

  const date = normalized ? new Date(normalized) : new Date();
  if (Number.isNaN(date.getTime())) {
    const fallback = new Date();
    return {
      datePart: fallback.toISOString().slice(0, 10),
      timePart: fallback.toTimeString().slice(0, 8).replace(/:/g, "-")
    };
  }

  return {
    datePart: date.toISOString().slice(0, 10),
    timePart: date.toTimeString().slice(0, 8).replace(/:/g, "-")
  };
}

function buildReportExportFileName({ tenantName, scriptName, reportDate, extension }) {
  const { datePart, timePart } = parseReportTimestamp(reportDate);
  const tenantPart = sanitizeFileName(tenantName || "Tenant");
  const scriptPart = sanitizeFileName(scriptName || "Report");
  const extensionPart = String(extension || "txt").replace(/^\./, "");
  return `${tenantPart}-${scriptPart}-${datePart}-${timePart}.${extensionPart}`;
}

function downloadBlob(blob, fileName) {
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.setTimeout(() => window.URL.revokeObjectURL(url), 1000);
}

function prepareHtmlForPdf(html, fileName) {
  const printStyles = `
    <style>
      @page { margin: 12mm; }
      body::before { display: none !important; }
      .topbar { position: static !important; backdrop-filter: none !important; }
      .page { max-width: none !important; padding: 0 !important; }
      .table-scroll { max-height: none !important; overflow: visible !important; }
      a { color: inherit !important; text-decoration: none !important; }
    </style>
  `;

  let preparedHtml = html;

  if (fileName && preparedHtml.includes("<title>") && preparedHtml.includes("</title>")) {
    preparedHtml = preparedHtml.replace(/<title>[\s\S]*?<\/title>/i, `<title>${escapeXml(fileName)}</title>`);
  }

  if (preparedHtml.includes("</head>")) {
    return preparedHtml.replace("</head>", `${printStyles}</head>`);
  }

  return `${printStyles}${preparedHtml}`;
}

function openPdfPrintDialog(html, fileName) {
  return new Promise((resolve, reject) => {
    const printableHtml = prepareHtmlForPdf(html, fileName);
    const blob = new Blob([printableHtml], { type: "text/html" });
    const url = window.URL.createObjectURL(blob);
    const iframe = document.createElement("iframe");

    iframe.style.position = "fixed";
    iframe.style.right = "0";
    iframe.style.bottom = "0";
    iframe.style.width = "0";
    iframe.style.height = "0";
    iframe.style.border = "0";
    iframe.style.opacity = "0";
    iframe.setAttribute("aria-hidden", "true");

    let settled = false;

    const cleanup = () => {
      window.setTimeout(() => {
        window.URL.revokeObjectURL(url);
        iframe.remove();
      }, 1000);
    };

    const finish = () => {
      if (settled) {
        return;
      }

      settled = true;
      cleanup();
      resolve();
    };

    const fail = (error) => {
      if (settled) {
        return;
      }

      settled = true;
      cleanup();
      reject(error);
    };

    iframe.onload = () => {
      const printWindow = iframe.contentWindow;
      if (!printWindow) {
        fail(new Error("Unable to open the report for PDF export."));
        return;
      }

      printWindow.onafterprint = finish;
      window.setTimeout(() => {
        try {
          printWindow.focus();
          printWindow.print();
          window.setTimeout(finish, 1500);
        } catch (error) {
          fail(error);
        }
      }, 150);
    };

    iframe.onerror = () => {
      fail(new Error("Unable to load the printable report."));
    };

    iframe.src = url;
    document.body.appendChild(iframe);
  });
}

async function fetchHtmlReport(url) {
  if (!url) {
    throw new Error("No HTML preview is available for this run yet.");
  }

  const response = await fetch(url);
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();

  if (!response.ok) {
    if (contentType.includes("application/json")) {
      try {
        const parsed = JSON.parse(text);
        throw new Error(parsed.message || "Failed to load the HTML report.");
      } catch (error) {
        if (error instanceof Error && error.message) {
          throw error;
        }
      }
    }

    throw new Error(`Failed to load the HTML report (${response.status}).`);
  }

  return text;
}

function CategoryIcon({ category }) {
  const normalized = (category || "Other").toLowerCase();
  let path = "M12 3.5a8.5 8.5 0 1 0 8.5 8.5A8.51 8.51 0 0 0 12 3.5Zm0 4.25a1.25 1.25 0 1 1-1.25 1.25A1.25 1.25 0 0 1 12 7.75Zm2.25 8.5h-4.5v-1.5h.75v-2.25h-.75V11h3v3.75h1.5Z";

  if (normalized === "identity") {
    path = "M12 3.5c-2.9 0-5.25 2.35-5.25 5.25S9.1 14 12 14s5.25-2.35 5.25-5.25S14.9 3.5 12 3.5Zm0 12c-3.43 0-6.38 1.9-7.88 4.7l1.33.8c1.23-2.3 3.68-3.75 6.55-3.75s5.32 1.45 6.55 3.75l1.33-.8c-1.5-2.8-4.45-4.7-7.88-4.7Z";
  } else if (normalized === "exchange") {
    path = "M4 6.25A2.25 2.25 0 0 1 6.25 4h11.5A2.25 2.25 0 0 1 20 6.25v11.5A2.25 2.25 0 0 1 17.75 20H6.25A2.25 2.25 0 0 1 4 17.75ZM7 8v1.5h10V8Zm0 3.25v1.5h6.5v-1.5Zm0 3.25V16h8v-1.5Z";
  } else if (normalized === "security") {
    path = "M12 3.5 5.5 6v4.75c0 4.15 2.8 7.97 6.5 8.95 3.7-.98 6.5-4.8 6.5-8.95V6Zm0 4a2.5 2.5 0 0 1 2.5 2.5v.5h.5v4h-6v-4h.5V10A2.5 2.5 0 0 1 12 7.5Zm0 1.5A1 1 0 0 0 11 10v.5h2V10a1 1 0 0 0-1-1Z";
  } else if (normalized === "sharepoint") {
    path = "M7 5.25A3.75 3.75 0 1 0 10.75 9 3.75 3.75 0 0 0 7 5.25Zm10 2A2.75 2.75 0 1 0 19.75 10 2.75 2.75 0 0 0 17 7.25ZM8 14c-3.05 0-5.65 1.68-7 4.17L2.28 19C3.38 16.96 5.5 15.5 8 15.5c1.13 0 2.18.3 3.08.83l.73-1.31A7.97 7.97 0 0 0 8 14Zm9.25.5c-2.14 0-4.02 1.07-5.16 2.7l1.23.86c.87-1.25 2.31-2.06 3.93-2.06 1.28 0 2.45.5 3.33 1.32l1.02-1.1a6.25 6.25 0 0 0-4.35-1.72Z";
  } else if (normalized === "teams") {
    path = "M6 6.25A2.25 2.25 0 0 1 8.25 4h5.5A2.25 2.25 0 0 1 16 6.25v11.5A2.25 2.25 0 0 1 13.75 20h-5.5A2.25 2.25 0 0 1 6 17.75ZM9 8v1.5h1.75V16h1.5V9.5H14V8Zm9.25.5a1.75 1.75 0 1 0 0 3.5 1.75 1.75 0 0 0 0-3.5Zm-1.5 4.75a2.75 2.75 0 0 0-2.75 2.75V17h5.5v-1a2.75 2.75 0 0 0-2.75-2.75Z";
  } else if (normalized === "reporting") {
    path = "M5 19.25h14v1.5H3.5V5H5Zm2-2.5V10h1.75v6.75Zm4 0V7.5h1.75v9.25Zm4 0V12h1.75v4.75Z";
  } else if (normalized === "licensing") {
    path = "M12 3.75 5 6.5v5.25c0 4.1 2.72 7.87 7 8.75 4.28-.88 7-4.65 7-8.75V6.5Zm-1.25 5h2.5v2.5h2.5v2.5h-2.5v2.5h-2.5v-2.5h-2.5v-2.5h2.5Z";
  } else if (normalized === "incident response") {
    path = "M12 4 3.5 19.5h17ZM12 9a.75.75 0 0 1 .75.75v4.5a.75.75 0 0 1-1.5 0v-4.5A.75.75 0 0 1 12 9Zm0 8a1 1 0 1 1 0-2 1 1 0 0 1 0 2Z";
  } else if (normalized === "operations") {
    path = "m12 4 1.15 2.53 2.78.33-1.87 2 0.5 2.74L12 10.35 9.44 11.6l.5-2.74-1.87-2 2.78-.33Zm-6.5 8 1.15 2.53 2.78.33-1.87 2 .5 2.74L5.5 18.35 2.94 19.6l.5-2.74-1.87-2 2.78-.33Zm13 0 1.15 2.53 2.78.33-1.87 2 .5 2.74-2.56-1.25-2.56 1.25.5-2.74-1.87-2 2.78-.33Z";
  } else if (normalized === "collaboration") {
    path = "M8 6.25A2.25 2.25 0 1 0 10.25 8.5 2.25 2.25 0 0 0 8 6.25Zm8 0A2.25 2.25 0 1 0 18.25 8.5 2.25 2.25 0 0 0 16 6.25ZM4.75 17A3.25 3.25 0 0 1 8 13.75h.5A3.25 3.25 0 0 1 11.75 17v.75h-7Zm7.5.75V17a4.7 4.7 0 0 0-.77-2.56 3.23 3.23 0 0 1 1.77-.69H16A3.25 3.25 0 0 1 19.25 17v.75Z";
  } else if (normalized === "devices") {
    path = "M7.25 4h9.5A2.25 2.25 0 0 1 19 6.25v11.5A2.25 2.25 0 0 1 16.75 20h-9.5A2.25 2.25 0 0 1 5 17.75V6.25A2.25 2.25 0 0 1 7.25 4ZM11 16.5h2v-1.25h-2Z";
  }

  return (
    <span className="catalog-group-icon" aria-hidden="true">
      <svg viewBox="0 0 24 24" focusable="false">
        <path d={path} fill="currentColor" />
      </svg>
    </span>
  );
}

function getFavoriteScriptIds() {
  try {
    const raw = window.localStorage.getItem("m365-toolbox-favorites");
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function getThemePreference() {
  try {
    return window.localStorage.getItem("m365-toolbox-theme") || "light";
  } catch {
    return "light";
  }
}

function createLocalId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }
  return `local-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getCompanySettings() {
  try {
    const raw = window.localStorage.getItem("m365-toolbox-companies");
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed)
      ? parsed
          .map((company) => ({
            id: company.id || createLocalId(),
            name: String(company.name || "").trim(),
            tenant: String(company.tenant || "").trim()
          }))
          .filter((company) => company.name && company.tenant)
      : [];
  } catch {
    return [];
  }
}

function findCompanyTenant(value, companies) {
  const query = String(value || "").trim().toLowerCase();
  if (!query) {
    return "";
  }

  const match = companies.find(
    (company) =>
      company.name.toLowerCase() === query ||
      company.tenant.toLowerCase() === query
  );

  return match?.tenant || value;
}

function escapeCsvCell(value) {
  const text = String(value ?? "");
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function parseCsvRows(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const nextChar = text[index + 1];

    if (quoted) {
      if (char === '"' && nextChar === '"') {
        cell += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        cell += char;
      }
      continue;
    }

    if (char === '"') {
      quoted = true;
    } else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (char !== "\r") {
      cell += char;
    }
  }

  row.push(cell);
  if (row.some((value) => String(value).trim())) {
    rows.push(row);
  }

  return rows;
}

function parseCompaniesCsv(text) {
  const rows = parseCsvRows(text);
  if (!rows.length) {
    return [];
  }

  const firstRow = rows[0].map((value) => String(value).trim().toLowerCase());
  const hasHeader = firstRow.some((value) => value.includes("company") || value.includes("tenant") || value.includes("domain"));
  const dataRows = hasHeader ? rows.slice(1) : rows;

  return dataRows
    .map((row) => ({
      id: createLocalId(),
      name: String(row[0] || "").trim(),
      tenant: String(row[1] || "").trim()
    }))
    .filter((company) => company.name && company.tenant);
}

function TenantLookupField({ field, value, onChange, companies }) {
  const [focused, setFocused] = useState(false);
  const rootRef = useRef(null);
  const menuRef = useRef(null);
  const query = String(value || "").trim().toLowerCase();
  const suggestions = companies
    .filter((company) => {
      if (!query) {
        return true;
      }
      return (
        company.name.toLowerCase().includes(query) ||
        company.tenant.toLowerCase().includes(query)
      );
    })
    .slice(0, 50);
  const menuStyle = useFloatingLayer(focused && suggestions.length > 0, rootRef, menuRef, {
    matchWidth: true,
    estimatedHeight: 320
  });

  return (
    <div className="form-field tenant-lookup-field" ref={rootRef}>
      <span>{field.label}</span>
      <input
        type="text"
        placeholder={companies.length ? "Type company, tenant ID, or domain" : field.placeholder || ""}
        value={value ?? ""}
        onFocus={() => setFocused(true)}
        onBlur={() => window.setTimeout(() => setFocused(false), 120)}
        onChange={(event) => onChange(field.id, event.target.value)}
      />
      {focused && suggestions.length && menuStyle
        ? createPortal(
            <div className="tenant-lookup-menu" ref={menuRef} style={menuStyle}>
              {suggestions.map((company) => (
                <button
                  key={company.id}
                  type="button"
                  className="tenant-lookup-option"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => {
                    onChange(field.id, company.tenant);
                    setFocused(false);
                  }}
                >
                  <span className="tenant-lookup-name">{company.name}</span>
                  <span className="tenant-lookup-domain">{company.tenant}</span>
                </button>
              ))}
            </div>,
            document.body
          )
        : null}
      {field.helpText ? <small>{field.helpText}</small> : null}
    </div>
  );
}

function InfoTooltip({ label, children }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);
  const panelRef = useRef(null);
  const panelStyle = useFloatingLayer(open, rootRef, panelRef, {
    minWidth: 300,
    estimatedHeight: 110
  });

  return (
    <span
      className="info-tooltip"
      ref={rootRef}
      onMouseEnter={() => setOpen(true)}
      onMouseLeave={() => setOpen(false)}
      onFocus={() => setOpen(true)}
      onBlur={() => setOpen(false)}
    >
      <button type="button" className="info-icon-btn" aria-label={label}>
        i
      </button>
      {open && panelStyle
        ? createPortal(
            <span className="info-tooltip-panel" ref={panelRef} style={panelStyle} role="tooltip">
              {children}
            </span>,
            document.body
          )
        : null}
    </span>
  );
}

function Field({ field, value, onChange, companies = [] }) {
  if (field.id === "tenantId") {
    return <TenantLookupField field={field} value={value} onChange={onChange} companies={companies} />;
  }

  if (field.type === "checkbox") {
    return (
      <label className="checkbox-field">
        <input
          type="checkbox"
          checked={Boolean(value)}
          onChange={(event) => onChange(field.id, event.target.checked)}
        />
        <span>{field.label}</span>
      </label>
    );
  }

  if (field.type === "textarea") {
    return (
      <label className="form-field">
        <span>{field.label}</span>
        <textarea
          rows={4}
          placeholder={field.placeholder || ""}
          value={value || ""}
          onChange={(event) => onChange(field.id, event.target.value)}
        />
        {field.helpText ? <small>{field.helpText}</small> : null}
      </label>
    );
  }

  if (field.type === "multiselect") {
    const selected = Array.isArray(value) ? value : [];
    return (
      <div className="form-field">
        <span>{field.label}</span>
        <div className={`multiselect-grid${field.id === "actions" ? " multiselect-grid-wide" : ""}`}>
          {field.options.map((option) => (
            <label key={option} className="checkbox-field">
              <input
                type="checkbox"
                checked={selected.includes(option)}
                onChange={(event) => {
                  const next = event.target.checked
                    ? [...selected, option]
                    : selected.filter((entry) => entry !== option);
                  onChange(field.id, next);
                }}
              />
              <span>{option}</span>
            </label>
          ))}
        </div>
      </div>
    );
  }

  return (
    <label className="form-field">
      <span>{field.label}</span>
      <input
        type={field.type === "number" ? "number" : field.type === "password" ? "password" : "text"}
        min={field.min}
        max={field.max}
        placeholder={field.placeholder || ""}
        value={value ?? ""}
        onChange={(event) => onChange(field.id, field.type === "number" ? Number(event.target.value) : event.target.value)}
      />
      {field.helpText ? <small>{field.helpText}</small> : null}
    </label>
  );
}

function SelectField({ label, value, options, onChange }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);
  const menuRef = useRef(null);
  const selectedOption = options.find((option) => option.value === value) || options[0] || null;
  const estimatedMenuHeight = Math.min(520, Math.max(180, options.length * 38 + 20));
  const menuStyle = useFloatingLayer(open, rootRef, menuRef, {
    matchWidth: true,
    estimatedHeight: estimatedMenuHeight
  });

  useEffect(() => {
    if (!open) {
      return undefined;
    }

    const handlePointerDown = (event) => {
      if (!rootRef.current?.contains(event.target) && !menuRef.current?.contains(event.target)) {
        setOpen(false);
      }
    };

    const handleEscape = (event) => {
      if (event.key === "Escape") {
        setOpen(false);
      }
    };

    window.addEventListener("mousedown", handlePointerDown);
    window.addEventListener("keydown", handleEscape);

    return () => {
      window.removeEventListener("mousedown", handlePointerDown);
      window.removeEventListener("keydown", handleEscape);
    };
  }, [open]);

  return (
    <div className="form-field select-field" ref={rootRef}>
      <span>{label}</span>
      <button
        type="button"
        className={`select-trigger${open ? " open" : ""}`}
        onClick={() => setOpen((current) => !current)}
        aria-haspopup="listbox"
        aria-expanded={open}
      >
        <span className="select-trigger-label">{selectedOption?.label || ""}</span>
        <span className="select-trigger-chevron" aria-hidden="true">▾</span>
      </button>
      {open && menuStyle
        ? createPortal(
            <div className="select-menu" role="listbox" aria-label={label} ref={menuRef} style={menuStyle}>
              {options.map((option) => (
                <button
                  key={option.value}
                  type="button"
                  role="option"
                  aria-selected={option.value === value}
                  className={`select-option${option.value === value ? " active" : ""}`}
                  onClick={() => {
                    onChange(option.value);
                    setOpen(false);
                  }}
                >
                  {option.label}
                </button>
              ))}
            </div>,
            document.body
          )
        : null}
    </div>
  );
}

function DateField({ label, value, onChange }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);
  const panelRef = useRef(null);
  const selectedDate = parseInputDateValue(value);
  const [viewMonth, setViewMonth] = useState(() => startOfMonth(selectedDate || new Date()));
  const panelStyle = useFloatingLayer(open, rootRef, panelRef, { minWidth: 280, estimatedHeight: 360 });
  const today = new Date();
  const days = buildCalendarDays(viewMonth);

  useEffect(() => {
    if (selectedDate) {
      setViewMonth(startOfMonth(selectedDate));
    }
  }, [value]);

  useEffect(() => {
    if (!open) {
      return undefined;
    }

    const handlePointerDown = (event) => {
      if (!rootRef.current?.contains(event.target) && !panelRef.current?.contains(event.target)) {
        setOpen(false);
      }
    };

    const handleEscape = (event) => {
      if (event.key === "Escape") {
        setOpen(false);
      }
    };

    window.addEventListener("mousedown", handlePointerDown);
    window.addEventListener("keydown", handleEscape);

    return () => {
      window.removeEventListener("mousedown", handlePointerDown);
      window.removeEventListener("keydown", handleEscape);
    };
  }, [open]);

  return (
    <div className="form-field date-field" ref={rootRef}>
      <span>{label}</span>
      <button
        type="button"
        className={`select-trigger date-trigger${open ? " open" : ""}${value ? "" : " placeholder"}`}
        onClick={() => setOpen((current) => !current)}
        aria-haspopup="dialog"
        aria-expanded={open}
      >
        <span className="select-trigger-label">{formatDateInputDisplay(value)}</span>
        <span className="date-trigger-icon" aria-hidden="true">◷</span>
      </button>
      {open && panelStyle
        ? createPortal(
            <div className="date-popover" ref={panelRef} style={panelStyle} role="dialog" aria-label={`${label} date`}>
              <div className="date-popover-header">
                <button
                  type="button"
                  className="date-nav-btn"
                  onClick={() => setViewMonth((current) => addMonths(current, -1))}
                  aria-label="Previous month"
                >
                  ‹
                </button>
                <div className="date-popover-title">{formatCalendarHeading(viewMonth)}</div>
                <button
                  type="button"
                  className="date-nav-btn"
                  onClick={() => setViewMonth((current) => addMonths(current, 1))}
                  aria-label="Next month"
                >
                  ›
                </button>
              </div>
              <div className="date-weekdays">
                {["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"].map((day) => (
                  <span key={day}>{day}</span>
                ))}
              </div>
              <div className="date-grid">
                {days.map((date) => {
                  const inCurrentMonth = date.getMonth() === viewMonth.getMonth();
                  const isSelected = selectedDate ? sameDay(date, selectedDate) : false;
                  const isToday = sameDay(date, today);

                  return (
                    <button
                      key={date.toISOString()}
                      type="button"
                      className={`date-cell${inCurrentMonth ? "" : " muted"}${isSelected ? " active" : ""}${isToday ? " today" : ""}`}
                      onClick={() => {
                        onChange(toInputDateValue(date));
                        setOpen(false);
                      }}
                    >
                      {date.getDate()}
                    </button>
                  );
                })}
              </div>
              <div className="date-popover-actions">
                <button
                  type="button"
                  className="date-action-btn"
                  onClick={() => {
                    onChange("");
                    setOpen(false);
                  }}
                >
                  Clear
                </button>
                <button
                  type="button"
                  className="date-action-btn primary"
                  onClick={() => {
                    const nextValue = toInputDateValue(today);
                    onChange(nextValue);
                    setViewMonth(startOfMonth(today));
                    setOpen(false);
                  }}
                >
                  Today
                </button>
              </div>
            </div>,
            document.body
          )
        : null}
    </div>
  );
}

export function App() {
  const reportCardRef = useRef(null);
  const companyImportInputRef = useRef(null);
  const nextToastIdRef = useRef(0);
  const toastTimersRef = useRef(new Map());
  const [sidebarWidth, setSidebarWidth] = useState(280);
  const [isResizingSidebar, setIsResizingSidebar] = useState(false);
  const [scripts, setScripts] = useState([]);
  const [selectedScript, setSelectedScript] = useState(null);
  const [formValues, setFormValues] = useState({});
  const [runs, setRuns] = useState([]);
  const [activeRun, setActiveRun] = useState(null);
  const [artifacts, setArtifacts] = useState([]);
  const [status, setStatus] = useState(null);
  const [statusUpdatedAt, setStatusUpdatedAt] = useState("");
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [toasts, setToasts] = useState([]);
  const [exportingFormat, setExportingFormat] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [canceling, setCanceling] = useState(false);
  const [nowMs, setNowMs] = useState(() => Date.now());
  const [runDetailsOpen, setRunDetailsOpen] = useState(true);
  const [recentRunsOpen, setRecentRunsOpen] = useState(true);
  const [devicePromptDismissed, setDevicePromptDismissed] = useState(false);
  const [expandedCategories, setExpandedCategories] = useState({});
  const [scriptSearch, setScriptSearch] = useState("");
  const [favoriteScriptIds, setFavoriteScriptIds] = useState(() => getFavoriteScriptIds());
  const [companies, setCompanies] = useState(() => getCompanySettings());
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [companyDraft, setCompanyDraft] = useState({ name: "", tenant: "" });
  const [favoritesOnly, setFavoritesOnly] = useState(false);
  const [modeFilter, setModeFilter] = useState("all");
  const [theme, setTheme] = useState(() => getThemePreference());
  const [runStatusFilter, setRunStatusFilter] = useState("all");
  const [runScriptFilter, setRunScriptFilter] = useState("all");
  const [runTenantFilter, setRunTenantFilter] = useState("");
  const [runDateFrom, setRunDateFrom] = useState("");
  const [runDateTo, setRunDateTo] = useState("");
  const [runOffset, setRunOffset] = useState(0);
  const [runTotal, setRunTotal] = useState(0);
  const [runsLoading, setRunsLoading] = useState(false);
  const runPageSize = 12;
  const activeRunIsBusy = Boolean(activeRun && ["running", "queued", "canceling"].includes(activeRun.status));

  const loadRuns = useCallback(async (nextOffset = runOffset) => {
    setRunsLoading(true);
    try {
      const query = buildRunHistoryQuery({
        limit: runPageSize,
        offset: nextOffset,
        status: runStatusFilter,
        scriptId: runScriptFilter,
        tenantId: runTenantFilter.trim(),
        dateFrom: runDateFrom,
        dateTo: runDateTo
      });
      const runsResponse = await fetch(`${apiBase}/runs?${query}`);
      const runsData = await parseApiResponse(runsResponse);
      if (!runsResponse.ok) {
        throw new Error(runsData.message || "Failed to load run history.");
      }
      setRuns(runsData.items || []);
      setRunTotal(runsData.total ?? 0);
      setRunOffset(runsData.offset ?? nextOffset);
    } finally {
      setRunsLoading(false);
    }
  }, [runDateFrom, runDateTo, runOffset, runScriptFilter, runStatusFilter, runTenantFilter]);

  useEffect(() => {
    const load = async () => {
      const [scriptsResponse, statusResponse] = await Promise.all([
        fetch(`${apiBase}/scripts`),
        fetch(`${apiBase}/status`)
      ]);

      const scriptsData = await parseApiResponse(scriptsResponse);
      const statusData = await parseApiResponse(statusResponse);
      setScripts(scriptsData);
      setStatus(statusData);
      setStatusUpdatedAt(new Date().toISOString());
      setSelectedScript(null);
      setFormValues({});
      setExpandedCategories({});
    };

    load().catch((loadError) => setError(loadError.message));
  }, []);

  useEffect(() => {
    loadRuns(runOffset).catch((loadError) => setError(loadError.message));
  }, [loadRuns, runOffset]);

  useEffect(() => {
    setRunOffset(0);
  }, [runStatusFilter, runScriptFilter, runTenantFilter, runDateFrom, runDateTo]);

  useEffect(() => {
    document.body.dataset.theme = theme;

    try {
      window.localStorage.setItem("m365-toolbox-theme", theme);
    } catch {
      // Ignore storage errors.
    }
  }, [theme]);

  useEffect(() => () => {
    toastTimersRef.current.forEach((timer) => window.clearTimeout(timer));
    toastTimersRef.current.clear();
  }, []);

  const dismissToast = (toastId) => {
    const timer = toastTimersRef.current.get(toastId);
    if (timer) {
      window.clearTimeout(timer);
      toastTimersRef.current.delete(toastId);
    }
    setToasts((current) => current.filter((toast) => toast.id !== toastId));
  };

  const pushToast = (kind, message) => {
    if (!message) {
      return;
    }

    const toastId = nextToastIdRef.current++;
    setToasts((current) => [...current, { id: toastId, kind, message }].slice(-4));

    const timeoutMs = kind === "error" ? 6000 : 3200;
    const timer = window.setTimeout(() => {
      setToasts((current) => current.filter((toast) => toast.id !== toastId));
      toastTimersRef.current.delete(toastId);
    }, timeoutMs);
    toastTimersRef.current.set(toastId, timer);
  };

  useEffect(() => {
    if (!error) {
      return;
    }
    pushToast("error", error);
    setError("");
  }, [error]);

  useEffect(() => {
    if (!success) {
      return;
    }
    pushToast("success", success);
    setSuccess("");
  }, [success]);

  useEffect(() => {
    if (!activeRun || !["running", "queued", "canceling"].includes(activeRun.status)) {
      return undefined;
    }

    const timer = window.setInterval(async () => {
      try {
        const response = await fetch(`${apiBase}/runs/${activeRun.id}`);
        const data = await parseApiResponse(response);
        if (!response.ok) {
          throw new Error(data.message || "Failed to refresh the active run.");
        }
        setActiveRun(data);
        await loadRuns(runOffset);
      } catch (pollError) {
        setError(pollError.message);
      }
    }, 2000);

    return () => window.clearInterval(timer);
  }, [activeRun, loadRuns, runOffset]);

  useEffect(() => {
    if (!activeRunIsBusy) {
      return undefined;
    }

    const timer = window.setInterval(() => {
      setNowMs(Date.now());
    }, 1000);

    return () => window.clearInterval(timer);
  }, [activeRunIsBusy]);

  useEffect(() => {
    if (!activeRun?.id) {
      setArtifacts([]);
      return undefined;
    }

    let cancelled = false;

    const loadArtifacts = async () => {
      try {
        const response = await fetch(`${apiBase}/runs/${activeRun.id}/artifacts`);
        const data = await parseApiResponse(response);
        if (!cancelled) {
          setArtifacts(data);
        }
      } catch {
        if (!cancelled) {
          setArtifacts([]);
        }
      }
    };

    loadArtifacts();

    return () => {
      cancelled = true;
    };
  }, [activeRun?.id, activeRun?.status]);

  useEffect(() => {
    if (activeRun?.status === "completed") {
      setSuccess(`Completed ${activeRun.scriptName} successfully.`);
    }
  }, [activeRun?.status, activeRun?.scriptName]);

  useEffect(() => {
    const hasHtmlArtifact = artifacts.some((artifact) => artifact.type === "html");
    if (activeRun?.id && hasHtmlArtifact) {
      setRunDetailsOpen(false);
      setRecentRunsOpen(false);
    }
  }, [activeRun?.id, artifacts]);

  useEffect(() => {
    const hasHtmlArtifact = artifacts.some((artifact) => artifact.type === "html");
    if (!activeRun?.id || !hasHtmlArtifact) {
      return;
    }

    const frame = window.requestAnimationFrame(() => {
      reportCardRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      reportCardRef.current?.focus({ preventScroll: true });
    });

    return () => window.cancelAnimationFrame(frame);
  }, [activeRun?.id, artifacts]);

  useEffect(() => {
    if (activeRun?.status === "running") {
      setDevicePromptDismissed(false);
    }
  }, [activeRun?.id, activeRun?.status]);

  useEffect(() => {
    if (!isResizingSidebar) {
      return undefined;
    }

    const handlePointerMove = (event) => {
      const nextWidth = Math.min(Math.max(event.clientX, 220), 420);
      setSidebarWidth(nextWidth);
    };

    const handlePointerUp = () => {
      setIsResizingSidebar(false);
    };

    window.addEventListener("mousemove", handlePointerMove);
    window.addEventListener("mouseup", handlePointerUp);

    return () => {
      window.removeEventListener("mousemove", handlePointerMove);
      window.removeEventListener("mouseup", handlePointerUp);
    };
  }, [isResizingSidebar]);

  useEffect(() => {
    try {
      window.localStorage.setItem("m365-toolbox-favorites", JSON.stringify(favoriteScriptIds));
    } catch {
      // Ignore storage errors.
    }
  }, [favoriteScriptIds]);

  useEffect(() => {
    try {
      window.localStorage.setItem("m365-toolbox-companies", JSON.stringify(companies));
    } catch {
      // Ignore storage errors.
    }
  }, [companies]);

  const handleScriptSelect = (script) => {
    setSettingsOpen(false);
    setSelectedScript(script);
    setFormValues(normalizeDefaults(script.fields));
    setExpandedCategories({ [script.category || "Other"]: true });
    setError("");
    setSuccess("");
    setActiveRun(null);
    setArtifacts([]);
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
  };

  const handleGoHome = () => {
    setSettingsOpen(false);
    setSelectedScript(null);
    setFormValues({});
    setExpandedCategories({});
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
  };

  const handleOpenSettings = () => {
    setSettingsOpen(true);
    setSelectedScript(null);
    setFormValues({});
    setActiveRun(null);
    setArtifacts([]);
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
  };

  const handleAddCompany = (event) => {
    event.preventDefault();
    const name = companyDraft.name.trim();
    const tenant = companyDraft.tenant.trim();

    if (!name || !tenant) {
      setError("Company name and tenant ID or domain are required.");
      return;
    }

    const duplicate = companies.some(
      (company) =>
        company.name.toLowerCase() === name.toLowerCase() ||
        company.tenant.toLowerCase() === tenant.toLowerCase()
    );

    if (duplicate) {
      setError("That company name or tenant value already exists.");
      return;
    }

    setCompanies((current) => [...current, { id: createLocalId(), name, tenant }]);
    setCompanyDraft({ name: "", tenant: "" });
    setSuccess(`Added ${name}.`);
  };

  const handleRemoveCompany = (companyId) => {
    setCompanies((current) => current.filter((company) => company.id !== companyId));
  };

  const handleExportCompanies = () => {
    const rows = [
      ["Company Name", "Tenant ID or Domain"],
      ...companies.map((company) => [company.name, company.tenant])
    ];
    const csv = rows.map((row) => row.map(escapeCsvCell).join(",")).join("\r\n");
    downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), "m365-toolbox-companies.csv");
    setSuccess("Companies exported.");
  };

  const handleImportCompanies = async (event) => {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (!file) {
      return;
    }

    try {
      const imported = parseCompaniesCsv(await file.text());
      if (!imported.length) {
        throw new Error("No companies found in the CSV file.");
      }

      let added = 0;
      setCompanies((current) => {
        const seen = new Set(current.flatMap((company) => [
          company.name.toLowerCase(),
          company.tenant.toLowerCase()
        ]));
        const next = [...current];

        for (const company of imported) {
          const nameKey = company.name.toLowerCase();
          const tenantKey = company.tenant.toLowerCase();
          if (seen.has(nameKey) || seen.has(tenantKey)) {
            continue;
          }
          seen.add(nameKey);
          seen.add(tenantKey);
          next.push(company);
          added += 1;
        }

        return next;
      });
      setSuccess(`Imported ${added} compan${added === 1 ? "y" : "ies"}.`);
    } catch (importError) {
      setError(importError.message);
    }
  };

  const handleOpenRun = async (run) => {
    setSettingsOpen(false);
    setArtifacts([]);
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
    const matchingScript = scripts.find((script) => script.id === run?.scriptId);
    if (matchingScript) {
      setSelectedScript(matchingScript);
    }
    if (!run?.id) {
      setActiveRun(null);
      return;
    }
    try {
      const response = await fetch(`${apiBase}/runs/${run.id}`);
      const fullRun = await parseApiResponse(response);
      if (!response.ok) {
        throw new Error(fullRun.message || "Failed to load run details.");
      }
      setActiveRun(fullRun);
    } catch (openError) {
      setError(openError.message);
    }
  };

  const handleChange = (fieldId, nextValue) => {
    setFormValues((current) => ({ ...current, [fieldId]: nextValue }));
  };

  const handleApplyActionProfile = (profile) => {
    if (!profile?.actions?.length) {
      return;
    }

    setFormValues((current) => ({
      ...current,
      actions: [...profile.actions]
    }));
    setSuccess(`Applied ${profile.label} action profile.`);
  };

  const handleToggleFavorite = (scriptId) => {
    setFavoriteScriptIds((current) =>
      current.includes(scriptId)
        ? current.filter((id) => id !== scriptId)
        : [...current, scriptId]
    );
  };

  const handleExportReport = async (format) => {
    if (!activeRun?.id) {
      return;
    }

    setExportingFormat(format);
    setError("");
    setSuccess("");

    try {
      const html = await fetchHtmlReport(activeRun.artifacts?.htmlPreviewUrl);
      const reportData = extractReportDataFromHtml(html);
      const fileName = buildReportExportFileName({
        tenantName: reportData.tenant || activeRun.payload?.tenantId || "Tenant",
        scriptName: activeRun.scriptName || reportData.title || "Report",
        reportDate: reportData.reportDate || activeRun.finishedAt || activeRun.startedAt || activeRun.requestedAt,
        extension: format === "excel" ? "xlsx" : format
      });

      if (format === "html") {
        downloadBlob(new Blob([html], { type: "text/html;charset=utf-8" }), fileName);
        setSuccess("HTML export downloaded.");
        return;
      }

      if (format === "excel") {
        downloadBlob(createExcelExportBlob(html), fileName);
        setSuccess("Excel export downloaded.");
        return;
      }

      if (format === "pdf") {
        await openPdfPrintDialog(html, fileName);
        setSuccess("Print dialog opened. Choose Save as PDF.");
        return;
      }

      throw new Error("Unsupported export format requested.");
    } catch (exportError) {
      setError(exportError.message);
    } finally {
      setExportingFormat("");
    }
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (!selectedScript) return;

    const approvalConfirmed = !selectedScript.approvalRequired || window.confirm(
      `This script is marked as remediation and can change tenant state.\n\nDo you want to approve and launch "${selectedScript.name}"?`
    );

    if (!approvalConfirmed) {
      return;
    }

    setSubmitting(true);
    setError("");
    setSuccess("");

    try {
      const resolvedFormValues = { ...formValues };
      if (selectedScript.fields.some((field) => field.id === "tenantId")) {
        resolvedFormValues.tenantId = findCompanyTenant(formValues.tenantId, companies);
      }

      const response = await fetch(`${apiBase}/scripts/${selectedScript.id}/run`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...resolvedFormValues, approvalConfirmed })
      });

      const data = await parseApiResponse(response);
      if (!response.ok) {
        throw new Error(data.message || "Failed to start run.");
      }
      setActiveRun(data);
      setDevicePromptDismissed(false);
      setRunDetailsOpen(true);
      setRecentRunsOpen(true);
      setSuccess(data.status === "queued" ? "Run queued successfully." : "Run started successfully.");
      await loadRuns(0);
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      setSubmitting(false);
    }
  };

  const handleCancelRun = async () => {
    if (!activeRun?.id) return;

    setCanceling(true);
    setError("");
    setSuccess("");

    try {
      const response = await fetch(`${apiBase}/runs/${activeRun.id}/cancel`, {
        method: "POST"
      });
      const data = await parseApiResponse(response);
      if (!response.ok) {
        throw new Error(data.message || "Failed to cancel run.");
      }
      setActiveRun(data);
      setSuccess("Run cancellation requested.");
      await loadRuns(runOffset);
    } catch (cancelError) {
      setError(cancelError.message);
    } finally {
      setCanceling(false);
    }
  };

  const handleRerun = async (run) => {
    if (!run?.scriptId) return;
    if (run?.canRerun === false) {
      setError("This run used redacted parameters, so it cannot be re-launched directly from history.");
      return;
    }

    const matchingScript = scripts.find((script) => script.id === run.scriptId);
    if (!matchingScript) {
      setError("The original script is no longer available in the catalog.");
      return;
    }

    const approvalConfirmed = !matchingScript.approvalRequired || window.confirm(
      `This script is marked as remediation and can change tenant state.\n\nDo you want to approve and launch "${matchingScript.name}" again with the previous inputs?`
    );

    if (!approvalConfirmed) {
      return;
    }

    setSubmitting(true);
    setError("");
    setSuccess("");

    try {
      setSelectedScript(matchingScript);
      setFormValues(normalizeDefaults(matchingScript.fields));
      const response = await fetch(`${apiBase}/scripts/${matchingScript.id}/run`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...(run.payload || {}), approvalConfirmed })
      });

      const data = await parseApiResponse(response);
      if (!response.ok) {
        throw new Error(data.message || "Failed to re-run script.");
      }

      setFormValues({ ...(run.payload || {}) });
      setActiveRun(data);
      setDevicePromptDismissed(false);
      setRunDetailsOpen(true);
      setRecentRunsOpen(true);
      setSuccess(data.status === "queued" ? "Run re-queued successfully." : "Run started again with the previous inputs.");
      await loadRuns(0);
    } catch (rerunError) {
      setError(rerunError.message);
    } finally {
      setSubmitting(false);
    }
  };

  const devicePrompt = extractDeviceCodePrompt([activeRun?.stdout, activeRun?.stderr].filter(Boolean).join("\n"));
  const showDevicePrompt = Boolean(devicePrompt?.code) && activeRun?.status === "running" && !devicePromptDismissed;
  const activeScriptDefinition = activeRun
    ? scripts.find((script) => script.id === activeRun.scriptId) || selectedScript
    : selectedScript;
  const activeRunDurationMs = activeRun ? getLiveDurationMs(activeRun, nowMs) : null;
  const activeRunLastActivityMs = activeRun?.lastActivityAt
    ? Math.max(0, nowMs - new Date(activeRun.lastActivityAt).getTime())
    : null;
  const activeRunHeartbeatVisible = activeRun?.status === "running" && activeRunLastActivityMs !== null && activeRunLastActivityMs >= 20_000;
  const activeRunBannerVisible = Boolean(activeRun && ["running", "queued", "canceling"].includes(activeRun.status));
  const activeRunQueueLabel = activeRun?.status === "queued" && activeRun.queuePosition
    ? `Position ${activeRun.queuePosition} of ${activeRun.queueSize || activeRun.queuePosition}`
    : activeRun?.status === "queued"
      ? "Waiting for queue update"
      : activeRun?.status === "canceling"
        ? "Stop requested"
        : "Active now";
  const activeRunBannerTitle = activeRun?.status === "queued"
    ? "Run queued"
    : activeRun?.status === "canceling"
      ? "Canceling run"
      : "Run in progress";
  const activeRunHeartbeatText = activeRun?.status === "queued"
    ? "This run is waiting for a free execution slot. You can leave the page open while the queue advances."
    : activeRun?.status === "canceling"
      ? "Waiting for PowerShell to stop cleanly. Some commands take a little time to exit."
      : activeRunHeartbeatVisible
        ? "Still running. Some Microsoft 365, Graph, and Exchange queries can take a few minutes before the next output appears."
        : "";
  const runStatusOptions = [
    { value: "all", label: "All" },
    { value: "queued", label: "Queued" },
    { value: "running", label: "Running" },
    { value: "canceling", label: "Canceling" },
    { value: "completed", label: "Completed" },
    { value: "failed", label: "Failed" },
    { value: "canceled", label: "Canceled" }
  ];
  const runScriptOptions = [
    { value: "all", label: "All scripts" },
    ...scripts.map((script) => ({ value: script.id, label: script.name }))
  ];
  const normalizedSearch = scriptSearch.trim().toLowerCase();
  const runCountsByScriptId = runs.reduce((acc, run) => {
    acc[run.scriptId] = (acc[run.scriptId] || 0) + 1;
    return acc;
  }, {});
  const favoriteScripts = scripts.filter((script) => favoriteScriptIds.includes(script.id));
  const recentFavoriteScripts = favoriteScripts.slice(0, 3);
  const mostUsedScripts = [...scripts]
    .filter((script) => runCountsByScriptId[script.id])
    .sort((a, b) => (runCountsByScriptId[b.id] || 0) - (runCountsByScriptId[a.id] || 0))
    .slice(0, 3);
  const hasHtmlArtifact = artifacts.some((artifact) => artifact.type === "html");
  const htmlPreviewUrl = activeRun?.artifacts?.htmlPreviewUrl || null;
  const bundleUrl = activeRun?.artifacts?.bundleUrl || null;
  const compromisedTargetPreview = selectedScript?.id === "m365-compromised-account-remediation"
    ? parseUserList(formValues.userPrincipalName)
    : null;
  const filteredScripts = scripts.filter((script) => {
    const matchesSearch = !normalizedSearch || [
      script.name,
      script.category,
      script.summary,
      script.description,
      script.id
      ]
      .filter(Boolean)
      .some((value) => value.toLowerCase().includes(normalizedSearch));
    const matchesFavorite = !favoritesOnly || favoriteScriptIds.includes(script.id);
    const matchesMode = modeFilter === "all" || script.mode === modeFilter;
    return matchesSearch && matchesFavorite && matchesMode;
  });
  const scriptGroups = groupScriptsByCategory(filteredScripts);
  const sortedCategories = Object.keys(scriptGroups).sort((a, b) => a.localeCompare(b));

  const renderReportExportActions = ({ includePreview = false } = {}) => {
    if (!activeRun?.id || !htmlPreviewUrl) {
      return null;
    }

    return (
      <>
        <button
          type="button"
          className="filter-btn active-all"
          onClick={() => handleExportReport("html")}
          disabled={Boolean(exportingFormat)}
        >
          {exportingFormat === "html" ? "Downloading..." : "HTML"}
        </button>
        <button
          type="button"
          className="filter-btn"
          onClick={() => handleExportReport("excel")}
          disabled={Boolean(exportingFormat)}
        >
          {exportingFormat === "excel" ? "Building..." : "Excel"}
        </button>
        <button
          type="button"
          className="filter-btn"
          onClick={() => handleExportReport("pdf")}
          disabled={Boolean(exportingFormat)}
        >
          {exportingFormat === "pdf" ? "Preparing..." : "PDF"}
        </button>
        {includePreview ? (
          <a className="filter-btn" href={htmlPreviewUrl} target="_blank" rel="noreferrer">
            Preview
          </a>
        ) : null}
        {bundleUrl ? (
          <a className="filter-btn" href={bundleUrl}>
            ZIP
          </a>
        ) : null}
      </>
    );
  };

  const renderRecentRunsContent = () => (
    <div className="card-body">
      <div className="recent-runs-filters">
        <SelectField
          label="Status"
          value={runStatusFilter}
          options={runStatusOptions}
          onChange={setRunStatusFilter}
        />
        <SelectField
          label="Script"
          value={runScriptFilter}
          options={runScriptOptions}
          onChange={setRunScriptFilter}
        />
        <label className="form-field">
          <span>Tenant</span>
          <input value={runTenantFilter} onChange={(event) => setRunTenantFilter(event.target.value)} placeholder="Tenant ID or domain" />
        </label>
        <DateField label="From" value={runDateFrom} onChange={setRunDateFrom} />
        <DateField label="To" value={runDateTo} onChange={setRunDateTo} />
      </div>
      {runsLoading ? (
        <div className="empty-row">Loading run history...</div>
      ) : runs.length === 0 ? (
        <div className="empty-row">No runs yet. Launch a report to start building persistent history.</div>
      ) : (
        <>
          <div className="table-scroll recent-runs-scroll">
            <table>
              <thead>
                <tr>
                  <th>Script</th>
                  <th>Tenant</th>
                  <th>Status</th>
                  <th>Requested</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                {runs.map((run) => (
                  <tr key={run.id}>
                    <td>{run.scriptName}</td>
                    <td>{formatRunTenant(run)}</td>
                    <td>
                      <span className={`pill ${run.status === "completed" ? "badge-ok" : run.status === "failed" || run.status === "canceled" || run.status === "interrupted" ? "badge-crit" : "badge-warn"}`}>
                        {run.status}
                      </span>
                      {run.status === "queued" && run.queuePosition ? (
                        <div className="table-subtext">Queue {run.queuePosition} of {run.queueSize || run.queuePosition}</div>
                      ) : null}
                    </td>
                    <td>{formatDate(run.requestedAt || run.startedAt)}</td>
                    <td className="table-actions">
                      <button type="button" className="filter-btn active-all" onClick={() => handleOpenRun(run)}>
                        Open
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="panel-toolbar" style={{ marginTop: "1rem" }}>
            <div className="method-label">
              Showing {Math.min(runOffset + 1, runTotal)}-{Math.min(runOffset + runs.length, runTotal)} of {runTotal}
            </div>
            <div className="run-actions">
              <button type="button" className="filter-btn" disabled={runOffset === 0} onClick={() => setRunOffset(Math.max(0, runOffset - runPageSize))}>
                Previous
              </button>
              <button type="button" className="filter-btn" disabled={runOffset + runPageSize >= runTotal} onClick={() => setRunOffset(runOffset + runPageSize)}>
                Next
              </button>
            </div>
          </div>
        </>
      )}
    </div>
  );

  const renderRecentRunsCard = () => {
    return (
      <div className={`card recent-runs-card ${recentRunsOpen ? "" : "card-collapsed"}`}>
        <button type="button" className="card-header card-header-button" onClick={() => setRecentRunsOpen((current) => !current)}>
          <span className="card-title">Recent Runs</span>
          <span className="card-badge badge-neutral">{runTotal}</span>
          <span className="card-chevron">{recentRunsOpen ? "▾" : "▸"}</span>
        </button>
        {recentRunsOpen ? renderRecentRunsContent() : null}
      </div>
    );
  };

  const renderSettingsPage = () => (
    <div className="dash-page settings-page">
      <div className="sections">
        <div className="card">
          <div className="card-header">
            <span className="card-title">Settings</span>
            <span className="card-badge badge-neutral">{companies.length} companies</span>
          </div>
          <div className="card-body">
            <form className="company-form" onSubmit={handleAddCompany}>
              <label className="form-field">
                <span>Company Name</span>
                <input
                  value={companyDraft.name}
                  onChange={(event) => setCompanyDraft((current) => ({ ...current, name: event.target.value }))}
                  placeholder="Contoso"
                />
              </label>
              <label className="form-field">
                <span>Tenant ID or Domain</span>
                <input
                  value={companyDraft.tenant}
                  onChange={(event) => setCompanyDraft((current) => ({ ...current, tenant: event.target.value }))}
                  placeholder="contoso.onmicrosoft.com"
                />
              </label>
              <button type="submit" className="add-btn">Add Company</button>
            </form>
          </div>
        </div>

        <div className="card">
          <div className="card-header">
            <span className="card-title">Companies</span>
            <div className="run-actions">
              <InfoTooltip label="Company CSV format">
                CSV format: Company Name,Tenant ID or Domain. Example: Contoso,contoso.onmicrosoft.com. Wrap company names with commas in quotes.
              </InfoTooltip>
              <input
                ref={companyImportInputRef}
                type="file"
                accept=".csv,text/csv"
                className="visually-hidden"
                onChange={handleImportCompanies}
              />
              <button type="button" className="filter-btn" onClick={() => companyImportInputRef.current?.click()}>
                Import CSV
              </button>
              <button type="button" className="filter-btn" onClick={handleExportCompanies} disabled={!companies.length}>
                Export CSV
              </button>
            </div>
          </div>
          <div className="card-body">
            {companies.length ? (
              <div className="company-list">
                {companies.map((company) => (
                  <div key={company.id} className="company-item">
                    <div className="tenant-avatar">{company.name.slice(0, 2).toUpperCase()}</div>
                    <div className="tenant-info">
                      <div className="tenant-name">{company.name}</div>
                      <div className="tenant-meta">{company.tenant}</div>
                    </div>
                    <button type="button" className="filter-btn destructive" onClick={() => handleRemoveCompany(company.id)}>
                      Remove
                    </button>
                  </div>
                ))}
              </div>
            ) : (
              <div className="empty-row">Add companies here, then type a company name, tenant ID, or domain in any script tenant field.</div>
            )}
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div className="app-shell">
      <div className="toast-stack" aria-live="polite" aria-atomic="true">
        {toasts.map((toast) => (
          <div key={toast.id} className={`toast toast-${toast.kind}`} role="status">
            <span>{toast.message}</span>
            <button type="button" className="toast-close" onClick={() => dismissToast(toast.id)} aria-label="Dismiss notification">
              ×
            </button>
          </div>
        ))}
      </div>
      {showDevicePrompt ? (
        <div className="modal-backdrop" role="presentation">
          <div className="modal-card" role="dialog" aria-modal="true" aria-labelledby="device-auth-title">
            <div className="card-header">
              <span className="card-title" id="device-auth-title">Microsoft Authentication Required</span>
              <button type="button" className="modal-close" onClick={() => setDevicePromptDismissed(true)}>
                Close
              </button>
            </div>
            <div className="card-body modal-body">
              <p className="empty-sub">Open the Microsoft device login page and enter this code to continue the script run.</p>
              <div className="device-code-row">
                <div className="device-code-box">{devicePrompt.code}</div>
                <button
                  type="button"
                  className="device-copy-btn"
                  onClick={async () => {
                    const copied = await copyText(devicePrompt.code);
                    setSuccess(copied ? "Device code copied." : "Unable to copy device code automatically.");
                  }}
                  aria-label="Copy device code"
                  title="Copy device code"
                >
                  <svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M9 9.75A2.25 2.25 0 0 1 11.25 7.5h7.5A2.25 2.25 0 0 1 21 9.75v9A2.25 2.25 0 0 1 18.75 21h-7.5A2.25 2.25 0 0 1 9 18.75Zm-6-4.5A2.25 2.25 0 0 1 5.25 3h7.5A2.25 2.25 0 0 1 15 5.25V6.5h-1.5V5.25a.75.75 0 0 0-.75-.75h-7.5a.75.75 0 0 0-.75.75v9a.75.75 0 0 0 .75.75H6.5v1.5H5.25A2.25 2.25 0 0 1 3 14.25Z" fill="currentColor" />
                  </svg>
                </button>
              </div>
              <a className="filter-btn active-all" href={devicePrompt.url} target="_blank" rel="noreferrer">
                Open Microsoft Device Login
              </a>
            </div>
          </div>
        </div>
      ) : null}

      <header className="topbar">
        <button type="button" className="topbar-logo topbar-home-btn" onClick={handleGoHome}>
          M365 Toolbox
        </button>
        <div className="topbar-title">Web-based PowerShell operations for Microsoft 365</div>
        <div className="topbar-right">
          <button
            type="button"
            className={`theme-switch${theme === "dark" ? " dark" : ""}`}
            onClick={() => setTheme((current) => current === "dark" ? "light" : "dark")}
            aria-label={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
            aria-pressed={theme === "dark"}
          >
            <span className="theme-switch-icon sun" aria-hidden="true">
              <svg viewBox="0 0 24 24">
                <path d="M12 6.5A5.5 5.5 0 1 0 17.5 12 5.51 5.51 0 0 0 12 6.5Zm0-3.5h1.5v2.25H12Zm0 15.75h1.5V21H12ZM3 11.25h2.25v1.5H3Zm15.75 0H21v1.5h-2.25ZM5.47 4.41l1.06-1.06 1.59 1.59-1.06 1.06Zm10.41 10.41 1.06-1.06 1.59 1.59-1.06 1.06ZM4.41 18.53l-1.06-1.06 1.59-1.59 1.06 1.06Zm10.41-10.41-1.06-1.06 1.59-1.59 1.06 1.06Z" fill="currentColor" />
              </svg>
            </span>
            <span className="theme-switch-track">
              <span className="theme-switch-thumb" />
            </span>
            <span className="theme-switch-icon moon" aria-hidden="true">
              <svg viewBox="0 0 24 24">
                <path d="M14.73 3.2a8.84 8.84 0 0 0 0 17.6 8.5 8.5 0 0 0 6.07-2.54 9.62 9.62 0 0 1-3.02.48A9.75 9.75 0 0 1 8.04 9a9.62 9.62 0 0 1 .48-3.02A8.5 8.5 0 0 0 5.98 12a8.84 8.84 0 0 0 8.75 8.8A8.84 8.84 0 0 0 14.73 3.2Z" fill="currentColor" />
              </svg>
            </span>
          </button>
          <div className="topbar-count">{scripts.length} script{scripts.length === 1 ? "" : "s"}</div>
          <div className="topbar-count">{runTotal} run{runTotal === 1 ? "" : "s"}</div>
        </div>
      </header>

      <div className="layout" style={{ "--sidebar-w": `${sidebarWidth}px` }}>
        <aside className="sidebar">
          <div className="sidebar-header">
            <div className="sidebar-label">Script Catalog</div>
            <label className="sidebar-search">
              <input
                type="text"
                placeholder="Search scripts..."
                value={scriptSearch}
                onChange={(event) => setScriptSearch(event.target.value)}
              />
            </label>
            <div className="sidebar-filter-group">
              <div className="sidebar-filter-label">Filters</div>
              <div className="chip-row">
                <button
                  type="button"
                  className={favoritesOnly ? "filter-chip active" : "filter-chip"}
                  onClick={() => setFavoritesOnly((current) => !current)}
                >
                  Favorites Only
                </button>
                <button type="button" className={modeFilter === "all" ? "filter-chip active" : "filter-chip"} onClick={() => setModeFilter("all")}>
                  All Modes
                </button>
                <button type="button" className={modeFilter === "read-only" ? "filter-chip active" : "filter-chip"} onClick={() => setModeFilter("read-only")}>
                  Read-only
                </button>
                <button type="button" className={modeFilter === "remediation" ? "filter-chip active" : "filter-chip"} onClick={() => setModeFilter("remediation")}>
                  Remediation
                </button>
              </div>
            </div>
          </div>
          <div className="tenant-list">
            {sortedCategories.length === 0 ? (
              <div className="empty-row">No scripts match your search.</div>
            ) : null}
            {sortedCategories.map((category) => {
              const isExpanded = normalizedSearch ? true : Boolean(expandedCategories[category]);
              return (
                <div key={category} className="catalog-group">
                  <button
                    type="button"
                    className={`catalog-group-header${isExpanded ? " expanded" : ""}`}
                    onClick={() =>
                      setExpandedCategories(isExpanded ? {} : { [category]: true })
                    }
                  >
                    <span className="catalog-group-title">
                      <CategoryIcon category={category} />
                      <span>{category}</span>
                    </span>
                    <span className="catalog-group-count">{scriptGroups[category].length}</span>
                    <span className="catalog-group-chevron">{isExpanded ? "▾" : "▸"}</span>
                  </button>

                  {isExpanded ? (
                    <div className="catalog-group-items">
                      {scriptGroups[category].map((script) => (
                        <div
                          key={script.id}
                          className={selectedScript?.id === script.id ? "tenant-item active" : "tenant-item"}
                          onClick={() => handleScriptSelect(script)}
                          role="button"
                          tabIndex={0}
                          onKeyDown={(event) => {
                            if (event.key === "Enter" || event.key === " ") {
                              handleScriptSelect(script);
                            }
                          }}
                        >
                          <div className="tenant-avatar">{script.name.slice(0, 2).toUpperCase()}</div>
                          <div className="tenant-info">
                            <div className="tenant-name">{script.name}</div>
                            <div className="tenant-meta">{script.summary}</div>
                            <div className="tenant-tags">
                              <span className={`mini-pill ${script.mode === "remediation" ? "badge-crit" : "badge-ok"}`}>{script.mode}</span>
                              <span className="mini-pill badge-neutral">{script.estimatedRuntimeMinutes} min</span>
                            </div>
                          </div>
                          <button
                            type="button"
                            className={favoriteScriptIds.includes(script.id) ? "favorite-btn active" : "favorite-btn"}
                            onClick={(event) => {
                              event.stopPropagation();
                              handleToggleFavorite(script.id);
                            }}
                          >
                            Fav
                          </button>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </div>
              );
            })}
          </div>

          <div className="sidebar-footer">
            <div>{runTotal === 0 ? "No runs yet." : `${runTotal} tracked run${runTotal === 1 ? "" : "s"}`}</div>
            <div className="sidebar-footer-links">
              <a
                className="sidebar-repo-link"
                href="https://github.com/albercuba/M365-Toolbox"
                target="_blank"
                rel="noreferrer"
                aria-label="Open GitHub repository"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <path
                    d="M12 2C6.48 2 2 6.58 2 12.24c0 4.53 2.87 8.37 6.84 9.73.5.1.68-.22.68-.49 0-.24-.01-1.04-.01-1.89-2.78.62-3.37-1.21-3.37-1.21-.45-1.2-1.11-1.52-1.11-1.52-.91-.64.07-.63.07-.63 1 .07 1.53 1.06 1.53 1.06.9 1.58 2.35 1.12 2.92.86.09-.67.35-1.12.63-1.37-2.22-.26-4.56-1.15-4.56-5.13 0-1.13.39-2.05 1.03-2.77-.1-.26-.45-1.31.1-2.74 0 0 .84-.28 2.75 1.06A9.3 9.3 0 0 1 12 6.84c.85 0 1.71.12 2.51.35 1.91-1.34 2.75-1.06 2.75-1.06.55 1.43.2 2.48.1 2.74.64.72 1.03 1.64 1.03 2.77 0 3.99-2.34 4.87-4.57 5.12.36.32.68.95.68 1.92 0 1.39-.01 2.51-.01 2.85 0 .27.18.59.69.49A10.26 10.26 0 0 0 22 12.24C22 6.58 17.52 2 12 2Z"
                    fill="currentColor"
                  />
                </svg>
                <span>GitHub</span>
              </a>
              <button
                type="button"
                className={`sidebar-repo-link sidebar-settings-link${settingsOpen ? " active" : ""}`}
                onClick={handleOpenSettings}
                aria-label="Open settings"
                title="Settings"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <path
                    d="M19.43 12.98c.04-.32.07-.65.07-.98s-.02-.66-.07-.98l2.05-1.6a.5.5 0 0 0 .12-.64l-1.94-3.36a.5.5 0 0 0-.61-.22l-2.42.98a7.2 7.2 0 0 0-1.7-.98L14.56 2.6A.5.5 0 0 0 14.07 2h-4.14a.5.5 0 0 0-.49.4L9.07 5a7.2 7.2 0 0 0-1.7.98L4.95 5a.5.5 0 0 0-.61.22L2.4 8.58a.5.5 0 0 0 .12.64l2.05 1.6c-.04.32-.07.65-.07.98s.02.66.07.98l-2.05 1.6a.5.5 0 0 0-.12.64l1.94 3.36a.5.5 0 0 0 .61.22l2.42-.98c.52.4 1.09.73 1.7.98l.37 2.6a.5.5 0 0 0 .49.4h4.14a.5.5 0 0 0 .49-.4l.37-2.6c.61-.25 1.18-.58 1.7-.98l2.42.98a.5.5 0 0 0 .61-.22l1.94-3.36a.5.5 0 0 0-.12-.64Zm-7.43 2.27A3.25 3.25 0 1 1 15.25 12 3.25 3.25 0 0 1 12 15.25Z"
                    fill="currentColor"
                  />
                </svg>
                <span>Settings</span>
              </button>
            </div>
          </div>
        </aside>
        <div
          className={`sidebar-resizer${isResizingSidebar ? " active" : ""}`}
          role="separator"
          aria-orientation="vertical"
          aria-label="Resize script catalog"
          onMouseDown={() => setIsResizingSidebar(true)}
        />

        <main className="main">
          {activeRunBannerVisible ? (
            <div className={`run-banner ${activeRun?.status === "queued" ? "tone-warn" : activeRun?.status === "canceling" ? "tone-crit" : "tone-active"}`}>
              <div className="run-banner-head">
                <div className="run-banner-title-wrap">
                  <span className={`run-banner-spinner${activeRun?.status === "canceling" ? " canceling" : ""}`} aria-hidden="true" />
                  <div>
                    <div className="run-banner-eyebrow">{activeRun?.scriptName || "Script run"}</div>
                    <div className="run-banner-title">{activeRunBannerTitle}</div>
                  </div>
                </div>
                <span className={`pill ${activeRun?.status === "canceling" ? "badge-crit" : "badge-warn"}`}>
                  {activeRun?.status}
                </span>
              </div>
              <div className="run-banner-grid">
                <div className="run-banner-item">
                  <div className="method-label">Current Step</div>
                  <div className="method-count">{activeRun?.currentStep || "Waiting for logs"}</div>
                </div>
                <div className="run-banner-item">
                  <div className="method-label">Elapsed</div>
                  <div className="method-count">{formatDuration(activeRunDurationMs)}</div>
                </div>
                <div className="run-banner-item">
                  <div className="method-label">Last Activity</div>
                  <div className="method-count">{formatRelativeTime(activeRun?.lastActivityAt, nowMs)}</div>
                </div>
                <div className="run-banner-item">
                  <div className="method-label">{activeRun?.status === "queued" ? "Queue" : "Runtime Hint"}</div>
                  <div className="method-count">
                    {activeRun?.status === "queued"
                      ? activeRunQueueLabel
                      : formatEstimatedRuntime(activeScriptDefinition?.estimatedRuntimeMinutes)}
                  </div>
                </div>
              </div>
              {activeRunHeartbeatText ? (
                <div className="run-banner-note">{activeRunHeartbeatText}</div>
              ) : null}
            </div>
          ) : null}
          {settingsOpen ? (
            renderSettingsPage()
          ) : selectedScript ? (
            <>
              <div className="dash-topstrip">
                <div className="strip-item">
                  <div className="strip-label">Script</div>
                  <div className="strip-value">{selectedScript.name}</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Mode</div>
                  <div className="strip-value">{selectedScript.mode}</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Estimated Runtime</div>
                  <div className="strip-value">{selectedScript.estimatedRuntimeMinutes} min</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Script Runs</div>
                  <div className="strip-value">{runCountsByScriptId[selectedScript.id] || 0}</div>
                </div>
              </div>

              <div className="dash-page">
                <div className="sections">
                  <div className="card">
                    <div className="card-header">
                      <span className="card-title">Script Details</span>
                      <span className={`card-badge ${selectedScript.mode === "remediation" ? "badge-crit" : "badge-ok"}`}>{selectedScript.mode}</span>
                      {selectedScript.approvalRequired ? <span className="card-badge badge-warn">approval required</span> : null}
                    </div>
                    <div className="card-body">
                      <div className="method-grid">
                        <div className="method-item method-item-selected">
                          <div className="method-info">
                            <div className="method-label">Summary</div>
                            <div className="method-count">{selectedScript.summary}</div>
                          </div>
                        </div>
                      </div>
                      <div className="empty-row" style={{ marginTop: "0.85rem" }}>{selectedScript.description}</div>
                    </div>
                  </div>

                  <div className="manage-workspace">
                    <div className="sections">
                      <div className="card">
                        <div className="card-header">
                          <span className="card-title">Run Script</span>
                          <span className={`card-badge ${selectedScript.mode === "remediation" ? "badge-crit" : "badge-ok"}`}>{selectedScript.mode}</span>
                        </div>
                        <div className="card-body">
                          {selectedScript.approvalRequired ? (
                            <div className="approval-banner">This remediation workflow requires an approval confirmation before launch.</div>
                          ) : null}
                          <div className="runtime-hint">
                            {formatEstimatedRuntime(selectedScript.estimatedRuntimeMinutes)}. Larger tenants, Exchange-heavy reports, and first-time module activity can take longer.
                          </div>
                          {selectedScript.actionProfiles?.length ? (
                            <div className="manage-form-panel">
                              <h4>Action Profiles</h4>
                              <div className="shortcut-grid">
                                {selectedScript.actionProfiles.map((profile) => (
                                  <button
                                    key={profile.id}
                                    type="button"
                                    className="shortcut-link"
                                    onClick={() => handleApplyActionProfile(profile)}
                                  >
                                    <strong>{profile.label}</strong>
                                    <span className="shortcut-link-meta">{profile.description}</span>
                                  </button>
                                ))}
                              </div>
                            </div>
                          ) : null}
                          {selectedScript.highImpactActions?.length ? (
                            <div className="approval-banner">
                              High-impact actions: {selectedScript.highImpactActions.join(", ")}
                            </div>
                          ) : null}
                          {compromisedTargetPreview ? (
                            <div className="manage-form-panel">
                              <h4>Target Preview</h4>
                              <div className="quick-summary-grid">
                                <div className="quick-summary-item">
                                  <div className="method-label">Targets Entered</div>
                                  <div className="method-count">{compromisedTargetPreview.entries.length}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Unique Targets</div>
                                  <div className="method-count">{compromisedTargetPreview.unique.length}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Duplicates</div>
                                  <div className="method-count">{compromisedTargetPreview.duplicates.length || "None"}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Tenant Scope</div>
                                  <div className="method-count">{formValues.tenantId || "Auto-detect at sign-in"}</div>
                                </div>
                              </div>
                              <div className="empty-row compact" style={{ marginTop: "0.75rem" }}>
                                Live tenant validation such as user existence, mailbox presence, guest status, and sync state is surfaced in the per-user results after the run begins.
                              </div>
                            </div>
                          ) : null}
                          <form className="settings-row" onSubmit={handleSubmit}>
                            {selectedScript.fields.map((field) => (
                              <Field key={field.id} field={field} value={formValues[field.id]} onChange={handleChange} companies={companies} />
                            ))}
                            <button className="add-btn" type="submit" disabled={submitting}>
                              {submitting ? "Starting..." : selectedScript.approvalRequired ? "Approve and Run" : "Run in Toolbox"}
                            </button>
                          </form>
                        </div>
                      </div>

                      <div className={`card ${runDetailsOpen ? "" : "card-collapsed"}`}>
                        <button type="button" className="card-header card-header-button" onClick={() => setRunDetailsOpen((current) => !current)}>
                          <span className="card-title">Run Details</span>
                          <span className={`card-badge ${activeRun ? (activeRun.status === "completed" ? "badge-ok" : activeRun.status === "failed" || activeRun.status === "canceled" || activeRun.status === "interrupted" ? "badge-crit" : "badge-warn") : "badge-neutral"}`}>
                            {activeRun ? activeRun.status : "idle"}
                          </span>
                          <span className="card-chevron">{runDetailsOpen ? "▾" : "▸"}</span>
                        </button>
                        {runDetailsOpen ? (
                        <div className="card-body">
                          {!activeRun ? <div className="empty-row">Select a recent run or start a new one to inspect the output.</div> : null}
                          {activeRun ? (
                            <>
                              <div className="quick-summary-grid">
                                <div className="quick-summary-item">
                                  <div className="method-label">Requested</div>
                                  <div className="method-count">{formatDate(activeRun.requestedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Started</div>
                                  <div className="method-count">{formatDate(activeRun.startedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Finished</div>
                                  <div className="method-count">{formatDate(activeRun.finishedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Duration</div>
                                  <div className="method-count">{formatDuration(activeRunDurationMs)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Current Step</div>
                                  <div className="method-count">{activeRun.currentStep || "Waiting for logs"}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Last Activity</div>
                                  <div className="method-count">{formatRelativeTime(activeRun.lastActivityAt, nowMs)}</div>
                                </div>
                                {activeRun.status === "queued" ? (
                                  <div className="quick-summary-item">
                                    <div className="method-label">Queue</div>
                                    <div className="method-count">{activeRunQueueLabel}</div>
                                  </div>
                                ) : null}
                                {activeRun.status !== "queued" ? (
                                  <div className="quick-summary-item">
                                    <div className="method-label">Runtime Hint</div>
                                    <div className="method-count">{formatEstimatedRuntime(activeScriptDefinition?.estimatedRuntimeMinutes)}</div>
                                  </div>
                                ) : null}
                                <div className="quick-summary-item">
                                  <div className="method-label">Exit Code</div>
                                  <div className="method-count">{activeRun.exitCode ?? "Pending"}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Approval</div>
                                  <div className="method-count">{activeRun.approval?.status || "not recorded"}</div>
                                </div>
                              </div>
                              {activeRun.errorSummary ? (
                                <div className="flash flash-error soft">{activeRun.errorSummary}</div>
                              ) : null}
                              {!activeRun.errorSummary && activeRunHeartbeatText ? (
                                <div className="flash soft">{activeRunHeartbeatText}</div>
                              ) : null}
                              <div className="manage-form-panel" style={{ marginTop: "1rem" }}>
                                <div className="panel-toolbar">
                                  <h4>Command</h4>
                                  <div className="run-actions">
                                  <button
                                    type="button"
                                    className="filter-btn active-all"
                                    onClick={async () => {
                                      const copied = await copyText(activeRun.command);
                                      setSuccess(copied ? "Command copied." : "Unable to copy command automatically.");
                                    }}
                                  >
                                    Copy Command
                                  </button>
                                  {activeRun.payload ? (
                                    <button
                                      type="button"
                                      className="filter-btn"
                                      onClick={() => handleRerun(activeRun)}
                                      disabled={
                                        submitting ||
                                        activeRun.canRerun === false ||
                                        ["running", "queued", "canceling"].includes(activeRun.status)
                                      }
                                    >
                                      Re-Run
                                    </button>
                                  ) : null}
                                    {["running", "queued", "canceling"].includes(activeRun.status) ? (
                                      <button type="button" className="filter-btn destructive" onClick={handleCancelRun} disabled={canceling}>
                                        {canceling ? "Canceling..." : "Cancel Run"}
                                      </button>
                                    ) : null}
                                  </div>
                                </div>
                                {activeRun.canRerun === false ? (
                                  <div className="flash soft" style={{ marginBottom: "0.75rem" }}>
                                    This run stored redacted parameters only. Enter fresh credentials in the form before running it again.
                                  </div>
                                ) : null}
                                <pre className="manage-response">{activeRun.command}</pre>
                              </div>
                              <div className="manage-form-panel">
                                <h4>Structured Log</h4>
                                <div className="run-log-list">
                                  {activeRun.logs?.length ? activeRun.logs.slice(-12).reverse().map((entry) => (
                                    <div key={entry.id} className={`run-log-entry level-${entry.level}`}>
                                      <span className="run-log-time">{formatDate(entry.timestamp)}</span>
                                      <span className="run-log-message">{entry.message}</span>
                                    </div>
                                  )) : (
                                    <div className="empty-row compact">No structured log entries yet.</div>
                                  )}
                                </div>
                              </div>
                              <div className="manage-form-panel">
                                <div className="panel-toolbar">
                                  <h4>Stdout</h4>
                                  {activeRun.stdout ? (
                                    <button
                                      type="button"
                                      className="filter-btn"
                                      onClick={async () => {
                                        const copied = await copyText(formatOutputText(activeRun.stdout, ""));
                                        setSuccess(copied ? "Stdout copied." : "Unable to copy stdout automatically.");
                                      }}
                                    >
                                      Copy Output
                                    </button>
                                  ) : null}
                                </div>
                                <pre className="manage-response">{formatOutputText(activeRun.stdout, "No stdout yet.")}</pre>
                              </div>
                              <div className="manage-form-panel">
                                <div className="panel-toolbar">
                                  <h4>Stderr</h4>
                                  {activeRun.stderr ? (
                                    <button
                                      type="button"
                                      className="filter-btn"
                                      onClick={async () => {
                                        const copied = await copyText(formatOutputText(activeRun.stderr, ""));
                                        setSuccess(copied ? "Error output copied." : "Unable to copy error output automatically.");
                                      }}
                                    >
                                      Copy Error
                                    </button>
                                  ) : null}
                                </div>
                                <pre className="manage-response">{formatOutputText(activeRun.stderr, "No stderr.")}</pre>
                              </div>
                            </>
                          ) : null}
                        </div>
                        ) : null}
                      </div>

                      <div className="card">
                        <div className="card-header">
                          <span className="card-title">Artifacts</span>
                          <span className="card-badge badge-neutral">{artifacts.length}</span>
                        </div>
                        <div className="card-body">
                          {artifacts.length === 0 ? (
                            <div className="empty-row">Artifacts will appear here when a run exports HTML, CSV, XLSX, log, or text files.</div>
                          ) : (
                            <div className="artifact-list">
                              {artifacts.map((artifact) => (
                                <div key={artifact.id} className="artifact-item">
                                  <div>
                                    <div className="artifact-name">{artifact.name}</div>
                                    <div className="artifact-meta">
                                      {artifact.type.toUpperCase()} | {formatFileSize(artifact.size)} | {formatDate(artifact.createdAt)}
                                    </div>
                                  </div>
                                  <div className="artifact-actions">
                                    {artifact.type === "html" ? (
                                      renderReportExportActions({ includePreview: true })
                                    ) : (
                                      <a className="filter-btn active-all" href={artifact.downloadUrl}>
                                        Download
                                      </a>
                                    )}
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>
                      {activeRun?.id && hasHtmlArtifact ? (
                        <div className="card" ref={reportCardRef} tabIndex={-1}>
                          <div className="card-header">
                            <span className="card-title">HTML Report</span>
                            <span className="card-badge badge-ok">preview</span>
                            {renderReportExportActions()}
                          </div>
                          <div className="card-body report-card-body">
                            <iframe
                              title="HTML report preview"
                              className="report-frame"
                              src={htmlPreviewUrl}
                              sandbox="allow-same-origin allow-scripts"
                            />
                          </div>
                        </div>
                      ) : null}
                    </div>
                  </div>
                </div>
              </div>
            </>
          ) : (
            <div className="dash-page">
              <div className="sections">
                <div className="card">
                  <div className="card-header">
                    <span className="card-title">Start Here</span>
                    <span className="card-badge badge-neutral">home</span>
                  </div>
                  <div className="card-body">
                    <div className="empty-row">Browse categories on the left, use filters to separate read-only and remediation scripts, and open any script to see prerequisites, examples, and runtime expectations before launch.</div>
                  </div>
                </div>

                <div className="card">
                  <div className="card-header">
                    <span className="card-title">Dashboard</span>
                    <span className="card-badge badge-neutral">{scripts.length} scripts</span>
                  </div>
                  <div className="card-body">
                    <div className="method-grid">
                      <div className="method-item method-item-selected">
                        <div className="method-info">
                          <div className="method-label">Catalog</div>
                          <div className="method-count">{scripts.length} scripts across {Object.keys(groupScriptsByCategory(scripts)).length} categories.</div>
                        </div>
                      </div>
                      <div className="method-item">
                        <div className="method-info">
                          <div className="method-label">Persistent Runs</div>
                          <div className="method-count">{runTotal} tracked runs survive backend restarts.</div>
                        </div>
                      </div>
                      <div className="method-item">
                        <div className="method-info">
                          <div className="method-label">Favorites</div>
                          <div className="method-count">{favoriteScriptIds.length} saved favorites ready for quick access.</div>
                        </div>
                      </div>
                      <div className="method-item">
                        <div className="method-info">
                          <div className="method-label">Backend Status</div>
                          <div className="method-count">
                            {status?.execution?.available ? "Worker path ready" : "Execution degraded"} | {status?.redis?.available ? "Redis connected" : "Redis unavailable"}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="card">
                  <div className="card-header">
                    <span className="card-title">Shortcuts</span>
                    <button
                      type="button"
                      className="filter-btn"
                      onClick={() =>
                        fetch(`${apiBase}/status`)
                          .then(parseApiResponse)
                          .then((data) => {
                            setStatus(data);
                            setStatusUpdatedAt(new Date().toISOString());
                          })
                          .catch((loadError) => setError(loadError.message))
                      }
                    >
                      Refresh Status
                    </button>
                  </div>
                  <div className="card-body">
                    <div className="shortcut-grid">
                      <div className="shortcut-card">
                        <div className="method-label">Recent Favorites</div>
                        {recentFavoriteScripts.length ? recentFavoriteScripts.map((script) => (
                          <button key={script.id} type="button" className="shortcut-link" onClick={() => handleScriptSelect(script)}>
                            {script.name}
                          </button>
                        )) : <div className="empty-row compact">Mark a script as favorite to pin it here.</div>}
                      </div>
                      <div className="shortcut-card">
                        <div className="method-label">Most Used</div>
                        {mostUsedScripts.length ? mostUsedScripts.map((script) => (
                          <button key={script.id} type="button" className="shortcut-link" onClick={() => handleScriptSelect(script)}>
                            {script.name}
                          </button>
                        )) : <div className="empty-row compact">Most-used shortcuts appear after a few runs.</div>}
                      </div>
                      <div className="shortcut-card">
                        <div className="method-label">Health Snapshot</div>
                        <div className="status-list">
                          <div>Output path: {status?.paths?.outputWritable ? "writable" : "unavailable"}</div>
                          <div>Scripts mount: {status?.paths?.scriptsMounted ? "mounted" : "missing"}</div>
                          <div>Worker: {status?.worker?.fresh ? `${status.worker.status || "ready"}` : "missing or stale"}</div>
                          <div>Queue: {status?.queue ? `${status.queue.active || 0} active / ${status.queue.waiting || 0} waiting / ${status.queue.deadLetter || 0} dead-letter` : "pending"}</div>
                          <div>Redis: {status?.redis?.available ? "connected" : "unavailable"}</div>
                          <div>Last updated: {statusUpdatedAt ? formatDate(statusUpdatedAt) : "Pending"}</div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                {renderRecentRunsCard()}

              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
