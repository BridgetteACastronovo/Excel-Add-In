import {
  createTreeFromWorksheetRows,
  extractHeaderMap,
  formatTimestamp,
  renderReportMarkup,
  renderStandaloneHtml,
  slugify,
} from "./doctree.mjs";

const state = {
  lastExport: null,
};

const ui = {};

Office.onReady((info) => {
  cacheElements();
  wireEvents();

  if (info.host !== Office.HostType.Excel) {
    renderError("This add-in only runs inside Microsoft Excel.");
    return;
  }

  setStatus("Ready. Open the worksheet you want to map, then generate the tree.");
});

function cacheElements() {
  ui.titleInput = document.getElementById("report-title");
  ui.generateButton = document.getElementById("generate-button");
  ui.exportButton = document.getElementById("export-button");
  ui.statusMessage = document.getElementById("status-message");
  ui.preview = document.getElementById("preview");
  ui.sheetName = document.getElementById("sheet-name");
}

function wireEvents() {
  ui.generateButton.addEventListener("click", () => {
    void handleGenerate();
  });
  ui.exportButton.addEventListener("click", handleExport);
}

async function handleGenerate() {
  setBusy(true);
  setStatus("Reading the active worksheet...");

  try {
    const worksheetData = await readActiveWorksheetData();
    const tree = createTreeFromWorksheetRows(worksheetData.rows, {
      noteTargetsByRow: worksheetData.noteTargetsByRow,
    });
    const title = ui.titleInput.value.trim() || `${worksheetData.sheetName} Document Tree`;
    const generatedAt = formatTimestamp();

    ui.preview.innerHTML = renderReportMarkup(tree, {
      title,
      worksheetName: worksheetData.sheetName,
      generatedAt,
    });
    ui.preview.classList.remove("preview-empty");
    ui.sheetName.textContent = worksheetData.sheetName;

    state.lastExport = {
      fileName: `${slugify(title)}.html`,
      html: renderStandaloneHtml(tree, {
        title,
        worksheetName: worksheetData.sheetName,
        generatedAt,
      }),
    };
    ui.exportButton.disabled = false;
    setStatus(`Generated ${tree.documents.length} documents from ${worksheetData.sheetName}.`);
  } catch (error) {
    state.lastExport = null;
    ui.exportButton.disabled = true;
    renderError(error instanceof Error ? error.message : String(error));
  } finally {
    setBusy(false);
  }
}

async function readActiveWorksheetData() {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = worksheet.getUsedRangeOrNullObject(true);
    worksheet.load("name");
    usedRange.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    if (usedRange.isNullObject || usedRange.rowCount === 0 || usedRange.columnCount === 0) {
      throw new Error("The active worksheet is empty.");
    }

    const rows = usedRange.values.map((row) => [...row]);
    const noteTargetsByRow = {};

    try {
      const headerMap = extractHeaderMap(rows[0] ?? []);
      const noteColumnIndex = headerMap.notes;
      const noteCells = [];

      for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
        const cell = usedRange.getCell(rowIndex, noteColumnIndex);
        cell.load("hyperlink");
        noteCells.push({ cell, sourceRow: rowIndex + 1 });
      }

      if (noteCells.length > 0) {
        await context.sync();
        for (const noteCell of noteCells) {
          const address = noteCell.cell.hyperlink?.address?.trim();
          if (address) {
            noteTargetsByRow[noteCell.sourceRow] = address;
          }
        }
      }
    } catch {
      // Header validation is handled by the shared parser; skip hyperlink extraction if headers are incomplete.
    }

    return {
      sheetName: worksheet.name,
      rows,
      noteTargetsByRow,
    };
  });
}

function handleExport() {
  if (!state.lastExport) {
    return;
  }

  const blob = new Blob([state.lastExport.html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = state.lastExport.fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
  setStatus(`Exported ${state.lastExport.fileName}.`);
}

function setBusy(isBusy) {
  ui.generateButton.disabled = isBusy;
  ui.generateButton.textContent = isBusy ? "Generating..." : "Generate tree";
}

function setStatus(message) {
  ui.statusMessage.textContent = message;
  ui.statusMessage.classList.remove("status-error");
}

function renderError(message) {
  ui.preview.innerHTML = `
    <div class="error-card">
      <p class="error-label">Could not generate tree</p>
      <p class="error-message">${escapeHtml(message)}</p>
    </div>
  `;
  ui.preview.classList.remove("preview-empty");
  ui.statusMessage.textContent = message;
  ui.statusMessage.classList.add("status-error");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
