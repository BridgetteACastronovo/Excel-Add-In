import test from "node:test";
import assert from "node:assert/strict";

import {
  createTreeFromWorksheetRows,
  extractHeaderMap,
  renderStandaloneHtml,
} from "../taskpane/doctree.mjs";

const SAMPLE_ROWS = [
  ["Unique ID", "Document number", "Short Title", "Long Title", "Notes", "Parent Doc", "Status"],
  ["DOC-001", "QM-000", "Quality Manual", "Corporate Quality Manual", "", "", "Approved"],
  ["DOC-002", "PR-101", "Control of Documents", "Procedure for Control of Documents and Records", "https://example.com/procedures/control-of-documents", "QM-000", "Approved"],
  ["DOC-003", "WI-101", "Document Review", "Work Instruction for Periodic Document Review", "Used during annual review cycles.", "PR-101", "Awaiting Approval"],
  ["DOC-004", "FR-101", "Review Checklist", "Form for Document Review Checklist", "", "WI-101", "Draft"],
];

test("extractHeaderMap accepts the Excel template headers", () => {
  const headerMap = extractHeaderMap(SAMPLE_ROWS[0]);
  assert.equal(headerMap.uniqueId, 0);
  assert.equal(headerMap.documentNumber, 1);
  assert.equal(headerMap.parentDoc, 5);
});

test("createTreeFromWorksheetRows builds a sorted tree and standalone html", () => {
  const tree = createTreeFromWorksheetRows(SAMPLE_ROWS, {
    noteTargetsByRow: {
      3: "https://example.com/procedures/control-of-documents",
    },
  });

  assert.equal(tree.documents.length, 4);
  assert.equal(tree.roots.length, 1);
  assert.equal(tree.roots[0].documentNumber, "QM-000");
  assert.equal(tree.roots[0].children[0].documentNumber, "PR-101");

  const html = renderStandaloneHtml(tree, {
    title: "Quality Management Documents",
    worksheetName: "Documents",
    generatedAt: "2026-03-13 09:00",
  });
  assert.match(html, /Quality Management Documents/);
  assert.match(html, /Sheet Documents/);
  assert.match(html, /href="https:\/\/example.com\/procedures\/control-of-documents"/);
});

test("createTreeFromWorksheetRows fails on unresolved parent references", () => {
  const brokenRows = [...SAMPLE_ROWS];
  brokenRows[2] = [...brokenRows[2]];
  brokenRows[2][5] = "UNKNOWN-DOC";

  assert.throws(
    () => createTreeFromWorksheetRows(brokenRows),
    /Parent Doc references missing document numbers/
  );
});
