const HEADER_ALIASES = {
  uniqueid: "uniqueId",
  documentnumber: "documentNumber",
  shorttitle: "shortTitle",
  longtitle: "longTitle",
  notes: "notes",
  parentdoc: "parentDoc",
  parentdocument: "parentDoc",
  status: "status",
};

const REQUIRED_FIELDS = [
  "uniqueId",
  "documentNumber",
  "shortTitle",
  "longTitle",
  "notes",
  "parentDoc",
  "status",
];

export function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

export function cellToText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

export function slugify(value) {
  const slug = String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return slug || "unknown";
}

export function formatTimestamp(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day} ${hour}:${minute}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

export function statusPalette(status) {
  const normalized = String(status ?? "").toLowerCase();
  if (normalized.includes("approved") && !normalized.includes("await")) {
    return { accent: "#1f6f50", fill: "#d7efe4" };
  }
  if (normalized.includes("await") || normalized.includes("review") || normalized.includes("pending")) {
    return { accent: "#9a6700", fill: "#fbebc8" };
  }
  if (normalized.includes("draft")) {
    return { accent: "#6b4f0f", fill: "#f3e7c9" };
  }
  if (normalized.includes("super") || normalized.includes("obsolete") || normalized.includes("retired")) {
    return { accent: "#8a2f39", fill: "#f4d9dd" };
  }
  if (normalized.includes("archive")) {
    return { accent: "#4d5c63", fill: "#dde5e8" };
  }
  return { accent: "#215c91", fill: "#dbe8f5" };
}

export function noteFromCell(value, hyperlinkTarget = "") {
  const text = cellToText(value);
  if (!text) {
    return null;
  }
  const url = cellToText(hyperlinkTarget);
  if (url) {
    return { text, url };
  }
  if (/^https?:\/\//i.test(text)) {
    return { text: "Open note", url: text };
  }
  return { text, url: null };
}

export function extractHeaderMap(headerRow) {
  const headerMap = {};
  for (let columnIndex = 0; columnIndex < headerRow.length; columnIndex += 1) {
    const canonical = HEADER_ALIASES[normalizeHeader(headerRow[columnIndex])];
    if (canonical) {
      headerMap[canonical] = columnIndex;
    }
  }

  const missing = REQUIRED_FIELDS.filter((field) => !(field in headerMap));
  if (missing.length > 0) {
    throw new Error(
      `Missing required columns: ${missing.join(", ")}. Expected: ${REQUIRED_FIELDS.join(", ")}`
    );
  }
  return headerMap;
}

export function createTreeFromWorksheetRows(rows, options = {}) {
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error("The active worksheet is empty.");
  }

  const headerMap = extractHeaderMap(rows[0] ?? []);
  const documents = [];
  const noteTargetsByRow = options.noteTargetsByRow ?? {};

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    if (!row.some((cell) => cellToText(cell))) {
      continue;
    }

    const documentNumber = cellToText(row[headerMap.documentNumber]);
    if (!documentNumber) {
      continue;
    }

    const sourceRow = rowIndex + 1;
    documents.push({
      uniqueId: cellToText(row[headerMap.uniqueId]),
      documentNumber,
      shortTitle: cellToText(row[headerMap.shortTitle]),
      longTitle: cellToText(row[headerMap.longTitle]),
      note: noteFromCell(row[headerMap.notes], noteTargetsByRow[sourceRow]),
      parentDoc: cellToText(row[headerMap.parentDoc]) || null,
      status: cellToText(row[headerMap.status]) || "Unknown",
      sourceRow,
      children: [],
    });
  }

  return buildTree(documents);
}

export function buildTree(documents) {
  if (!Array.isArray(documents) || documents.length === 0) {
    throw new Error("No document rows were found in the worksheet.");
  }

  const uniqueIds = new Map();
  const documentNumbers = new Map();
  for (const document of documents) {
    if (!document.uniqueId) {
      throw new Error(`Row ${document.sourceRow} is missing a Unique ID.`);
    }
    if (uniqueIds.has(document.uniqueId)) {
      throw new Error(
        `Duplicate Unique ID '${document.uniqueId}' found in rows ${uniqueIds.get(document.uniqueId).sourceRow} and ${document.sourceRow}.`
      );
    }
    if (documentNumbers.has(document.documentNumber)) {
      throw new Error(
        `Duplicate document number '${document.documentNumber}' found in rows ${documentNumbers.get(document.documentNumber).sourceRow} and ${document.sourceRow}.`
      );
    }
    uniqueIds.set(document.uniqueId, document);
    documentNumbers.set(document.documentNumber, document);
  }

  const roots = [];
  const unresolvedParents = [];
  for (const document of documents) {
    document.children = [];
    if (!document.parentDoc) {
      roots.push(document);
      continue;
    }
    if (document.parentDoc === document.documentNumber) {
      throw new Error(`Row ${document.sourceRow} references itself as Parent Doc.`);
    }
    const parent = documentNumbers.get(document.parentDoc);
    if (!parent) {
      unresolvedParents.push(`${document.documentNumber} -> ${document.parentDoc} (row ${document.sourceRow})`);
      continue;
    }
    parent.children.push(document);
  }

  if (unresolvedParents.length > 0) {
    throw new Error(`Parent Doc references missing document numbers: ${unresolvedParents.sort().join("; ")}`);
  }

  detectCycles(documentNumbers);

  for (const document of documents) {
    document.children.sort((left, right) => left.documentNumber.localeCompare(right.documentNumber));
  }
  roots.sort((left, right) => left.documentNumber.localeCompare(right.documentNumber));

  return {
    roots,
    documents,
    statuses: [...new Set(documents.map((document) => document.status))].sort((left, right) =>
      left.localeCompare(right, undefined, { sensitivity: "base" })
    ),
  };
}

function detectCycles(documentNumbers) {
  const state = new Map();

  function visit(document, trail) {
    const currentState = state.get(document.documentNumber) ?? 0;
    if (currentState === 1) {
      throw new Error(`Cycle detected in Parent Doc relationships: ${[...trail, document.documentNumber].join(" -> ")}`);
    }
    if (currentState === 2) {
      return;
    }

    state.set(document.documentNumber, 1);
    for (const child of document.children) {
      visit(child, [...trail, document.documentNumber]);
    }
    state.set(document.documentNumber, 2);
  }

  for (const document of documentNumbers.values()) {
    if ((state.get(document.documentNumber) ?? 0) === 0) {
      visit(document, []);
    }
  }
}

function renderLegendItem(status) {
  const palette = statusPalette(status);
  return (
    `<div class="legend-item" style="--status-accent:${palette.accent}; --status-fill:${palette.fill};">` +
    '<span class="legend-swatch" aria-hidden="true"></span>' +
    `${escapeHtml(status)}</div>`
  );
}

function renderNode(document) {
  const palette = statusPalette(document.status);
  let noteLink = "";
  if (document.note) {
    if (document.note.url) {
      noteLink = `<a class="meta-chip" href="${escapeHtml(document.note.url)}" target="_blank" rel="noreferrer">Notes</a>`;
    } else {
      noteLink = `<a class="meta-chip" href="#note-${slugify(document.uniqueId)}">Notes</a>`;
    }
  }

  return (
    `<article class="node" style="--status-accent:${palette.accent}; --status-fill:${palette.fill};">` +
    `<div class="node-code">${escapeHtml(document.documentNumber)}</div>` +
    `<h2>${escapeHtml(document.shortTitle)}</h2>` +
    `<p>${escapeHtml(document.longTitle)}</p>` +
    '<div class="node-meta">' +
    `<span class="meta-chip">ID ${escapeHtml(document.uniqueId)}</span>` +
    `<span class="meta-chip">${escapeHtml(document.status)}</span>` +
    noteLink +
    "</div>" +
    "</article>"
  );
}

function renderBranch(document) {
  const children = document.children.map((child) => renderBranch(child)).join("");
  return `<li>${renderNode(document)}${children ? `<ul>${children}</ul>` : ""}</li>`;
}

function renderNote(document) {
  const noteBody = document.note?.url
    ? `<p><a href="${escapeHtml(document.note.url)}" target="_blank" rel="noreferrer">${escapeHtml(document.note.url)}</a></p>`
    : `<p>${escapeHtml(document.note?.text ?? "")}</p>`;

  return (
    `<article class="note" id="note-${slugify(document.uniqueId)}">` +
    `<p class="note-title">${escapeHtml(document.documentNumber)} · ${escapeHtml(document.shortTitle)}</p>` +
    noteBody +
    "</article>"
  );
}

function renderNotesPanel(tree) {
  const notesHtml = tree.documents
    .filter((document) => document.note)
    .map((document) => renderNote(document))
    .join("");

  if (!notesHtml) {
    return "";
  }

  return `<section class="panel notes"><h2>Notes</h2>${notesHtml}</section>`;
}

function reportStyles() {
  return `:root {
  --page-bg: linear-gradient(180deg, #f7f2e8 0%, #edf1f3 100%);
  --ink: #1e2528;
  --muted: #5d6a70;
  --panel: rgba(255, 255, 255, 0.84);
  --line: #8ea1aa;
  --shadow: 0 18px 44px rgba(40, 53, 61, 0.12);
  --radius: 18px;
}
* {
  box-sizing: border-box;
}
.report-page {
  color: var(--ink);
  font-family: "Aptos", "Segoe UI", sans-serif;
}
.report-page a {
  color: #0d5c8d;
}
.report-page .hero {
  display: grid;
  gap: 14px;
  margin-bottom: 28px;
}
.report-page .eyebrow {
  text-transform: uppercase;
  letter-spacing: 0.16em;
  font-size: 0.78rem;
  color: var(--muted);
  margin: 0;
}
.report-page h1 {
  margin: 0;
  font-family: Georgia, "Times New Roman", serif;
  font-size: clamp(2rem, 3vw, 3.3rem);
  line-height: 1.04;
}
.report-page .summary {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  color: var(--muted);
}
.report-page .pill {
  background: rgba(255, 255, 255, 0.72);
  border: 1px solid rgba(141, 160, 171, 0.5);
  border-radius: 999px;
  padding: 8px 12px;
  box-shadow: var(--shadow);
}
.report-page .panel {
  background: var(--panel);
  border: 1px solid rgba(141, 160, 171, 0.45);
  border-radius: 24px;
  box-shadow: var(--shadow);
  backdrop-filter: blur(12px);
}
.report-page .legend {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  padding: 18px;
  margin-bottom: 22px;
}
.report-page .legend-item {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 8px 12px;
  border-radius: 999px;
  background: rgba(255, 255, 255, 0.72);
  border: 1px solid rgba(141, 160, 171, 0.35);
}
.report-page .legend-swatch {
  width: 14px;
  height: 14px;
  border-radius: 999px;
  border: 2px solid var(--status-accent);
  background: var(--status-fill);
}
.report-page .tree-wrap {
  overflow-x: auto;
  padding: 22px;
}
.report-page .tree,
.report-page .tree ul {
  list-style: none;
  margin: 0;
  padding: 0;
}
.report-page .tree {
  display: grid;
  justify-items: center;
  gap: 18px;
}
.report-page .tree > li {
  width: fit-content;
}
.report-page .tree ul {
  display: flex;
  justify-content: center;
  gap: 0;
  padding-top: 34px;
  position: relative;
}
.report-page .tree ul::before {
  content: "";
  position: absolute;
  top: 16px;
  left: 50%;
  width: 2px;
  height: 18px;
  background: var(--line);
  transform: translateX(-50%);
}
.report-page .tree li {
  text-align: center;
  position: relative;
  padding: 0 18px;
}
.report-page .tree li::before,
.report-page .tree li::after {
  content: "";
  position: absolute;
  top: 16px;
  width: 50%;
  height: 2px;
  background: var(--line);
}
.report-page .tree li::before {
  right: 50%;
}
.report-page .tree li::after {
  left: 50%;
}
.report-page .tree li:only-child::before,
.report-page .tree li:only-child::after,
.report-page .tree > li::before,
.report-page .tree > li::after {
  display: none;
}
.report-page .tree li:first-child::before,
.report-page .tree li:last-child::after {
  display: none;
}
.report-page .tree li > .node::before {
  content: "";
  position: absolute;
  top: -18px;
  left: 50%;
  width: 2px;
  height: 18px;
  background: var(--line);
  transform: translateX(-50%);
}
.report-page .tree > li > .node::before {
  display: none;
}
.report-page .node {
  width: min(300px, 84vw);
  text-align: left;
  padding: 16px 16px 14px;
  border-radius: var(--radius);
  border: 2px solid var(--status-accent);
  background: linear-gradient(180deg, rgba(255,255,255,0.95), var(--status-fill));
  position: relative;
  box-shadow: var(--shadow);
}
.report-page .node-code {
  font-size: 0.78rem;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: var(--muted);
  margin-bottom: 8px;
}
.report-page .node h2 {
  font-size: 1.05rem;
  margin: 0 0 6px;
}
.report-page .node p {
  margin: 0;
  color: #334247;
  line-height: 1.45;
}
.report-page .node-meta {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  margin-top: 14px;
  font-size: 0.8rem;
}
.report-page .meta-chip {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 10px;
  border-radius: 999px;
  background: rgba(255,255,255,0.76);
  border: 1px solid rgba(30, 37, 40, 0.08);
  color: #314045;
  text-decoration: none;
}
.report-page .notes {
  margin-top: 26px;
  padding: 22px;
}
.report-page .notes h2 {
  margin-top: 0;
  font-family: Georgia, "Times New Roman", serif;
}
.report-page .note {
  padding: 14px 0;
  border-top: 1px solid rgba(141, 160, 171, 0.35);
}
.report-page .note:first-of-type {
  border-top: 0;
  padding-top: 0;
}
.report-page .note-title {
  margin: 0 0 6px;
  font-weight: 700;
}
@media (max-width: 800px) {
  .report-page .tree-wrap {
    padding: 14px;
  }
  .report-page .tree ul {
    flex-direction: column;
    align-items: center;
    gap: 26px;
  }
  .report-page .tree ul::before,
  .report-page .tree li::before,
  .report-page .tree li::after {
    display: none;
  }
  .report-page .tree li > .node::before {
    top: -22px;
    height: 22px;
  }
}`;
}

function reportMarkup(tree, options = {}) {
  const title = options.title ?? "Document Tree";
  const generatedAt = options.generatedAt ?? formatTimestamp();
  const worksheetName = cellToText(options.worksheetName);

  return `
    <header class="hero">
      <p class="eyebrow">Document Tree Generator</p>
      <h1>${escapeHtml(title)}</h1>
      <div class="summary">
        <span class="pill">${tree.documents.length} documents</span>
        <span class="pill">${tree.roots.length} root documents</span>
        <span class="pill">${tree.statuses.length} status values</span>
        ${worksheetName ? `<span class="pill">Sheet ${escapeHtml(worksheetName)}</span>` : ""}
        <span class="pill">Generated ${escapeHtml(generatedAt)}</span>
      </div>
    </header>
    <section class="panel legend" aria-label="Status legend">
      ${tree.statuses.map((status) => renderLegendItem(status)).join("")}
    </section>
    <section class="panel tree-wrap">
      <ul class="tree">
        ${tree.roots.map((document) => renderBranch(document)).join("")}
      </ul>
    </section>
    ${renderNotesPanel(tree)}
  `;
}

export function renderReportMarkup(tree, options = {}) {
  return `<style>${reportStyles()}</style><div class="report-page">${reportMarkup(tree, options)}</div>`;
}

export function renderStandaloneHtml(tree, options = {}) {
  const title = options.title ?? "Document Tree";
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(title)}</title>
  <style>
    body {
      margin: 0;
      background: linear-gradient(180deg, #f7f2e8 0%, #edf1f3 100%);
    }
    .page-shell {
      max-width: 1600px;
      margin: 0 auto;
      padding: 32px 20px 56px;
    }
    ${reportStyles()}
  </style>
</head>
<body>
  <main class="page-shell">
    <div class="report-page">
      ${reportMarkup(tree, options)}
    </div>
  </main>
</body>
</html>`;
}
