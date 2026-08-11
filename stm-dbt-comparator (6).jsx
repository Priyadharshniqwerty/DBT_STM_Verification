import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  Upload,
  FileSpreadsheet,
  FileCode2,
  ClipboardList,
  GitCompare,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  RefreshCw,
  Loader2,
  Search,
  ChevronDown,
  PlusCircle,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* Design tokens — "as-built" blueprint / drafting-table aesthetic     */
/* ------------------------------------------------------------------ */
const C = {
  inkDark: "#071B2E",
  ink: "#0E2C4A",
  inkLine: "#2F5C86",
  inkLineSoft: "#1B3E60",
  paper: "#F4F7F9",
  paperPanel: "#FFFFFF",
  paperLine: "#D6E1E9",
  textLight: "#E7F0F7",
  textMuted: "#8FAAC0",
  textMutedDark: "#5B7386",
  textDark: "#152233",
  accent: "#59C2E0",
  accentDim: "#2E5A6E",
  match: "#2F9E76",
  matchBg: "#E4F5EE",
  warn: "#C97F14",
  warnBg: "#FBF0DC",
  miss: "#C13F35",
  missBg: "#FBE7E5",
  extra: "#7259B0",
  extraBg: "#EFEAFA",
};

const FONTS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
`;

/* ------------------------------------------------------------------ */
/* Small helpers                                                       */
/* ------------------------------------------------------------------ */
function fuzzyIndex(headerRow, keywords) {
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i] || "").toLowerCase();
    if (keywords.some((k) => h.includes(k))) return i;
  }
  return -1;
}

function cleanHeaderCell(v, i) {
  const s = String(v ?? "").trim();
  return s || `Column ${i + 1}`;
}

function rowIsBlank(row) {
  return !row.some((c) => String(c ?? "").trim() !== "");
}

/* Parse a single STM tab following the documented layout:
   row1 = section/group labels (merged cells — forward-filled here),
   row2 = sub-labels (unused directly, kept implicit in row3's text),
   row3 = actual column headers, data starts row4
   (all 1-indexed -> 0-indexed 2 and 3). */
function parseStmSheet(workbook, tabName) {
  const sheetNames = workbook.SheetNames;
  const match = sheetNames.find(
    (n) => n.trim().toLowerCase() === tabName.trim().toLowerCase()
  );
  if (!match) {
    return { error: `No sheet tab named "${tabName}" was found in the workbook.` };
  }
  const sheet = workbook.Sheets[match];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  if (aoa.length < 4) {
    return { error: `Sheet "${match}" doesn't have enough rows for the expected 3-row header + data layout.` };
  }
  const groupRowRaw = aoa[0] || [];
  const subRowRaw = aoa[1] || [];
  const headerRow = aoa[2];
  const headers = headerRow.map((v, i) => cleanHeaderCell(v, i));

  // Row 1's group/section labels sit in merged cells, so only the first
  // column of each group actually has text — forward-fill it across the
  // blank cells that follow so every column knows which section it's in.
  const groups = [];
  let lastGroup = "";
  for (let i = 0; i < headers.length; i++) {
    const g = String(groupRowRaw[i] ?? "").trim();
    if (g) lastGroup = g;
    groups.push(lastGroup);
  }

  // Build a key for every column and GUARANTEE it's unique — some STM tabs
  // repeat a sub-header (e.g. "Datatype") under more than one section, and
  // if the group text for those columns ever comes back blank or identical,
  // a plain "group :: header" string can still collide, silently letting
  // one column's data overwrite another's. A per-occurrence counter closes
  // that gap entirely regardless of what row1 actually contains.
  const seenCount = {};
  const keys = headers.map((h, i) => {
    const base = groups[i] ? `${groups[i]} :: ${h}` : h;
    seenCount[base] = (seenCount[base] || 0) + 1;
    return seenCount[base] > 1 ? `${base} #${seenCount[base]}` : base;
  });

  const dataRows = aoa.slice(3).filter((r) => !rowIsBlank(r));
  // SCD Type is read directly from Column I (index 8, 0-indexed: A=0...I=8)
  // by raw position, the same way Join/Where and Business Rules are read —
  // independent of whatever the header text at that column says, so it
  // can never be lost to a header-matching mismatch across different tabs.
  const SCD_TYPE_COL_INDEX = 8;
  const rows = dataRows.map((r) => {
    const obj = {};
    keys.forEach((k, i) => {
      obj[k] = r[i] !== undefined ? String(r[i]).trim() : "";
    });
    obj["SCD Type (Column I)"] = r[SCD_TYPE_COL_INDEX] !== undefined ? String(r[SCD_TYPE_COL_INDEX]).trim() : "";
    return obj;
  });

  // The Join/Where clause is looked up independently of the header-based
  // grouping above — it's searched for across all three header rows (it
  // may live in the group row, the sub-label row, or the header row itself)
  // and its values are read straight from the raw sheet data by column
  // index, so it can never be lost to a header-matching mismatch.
  let joinWhereColIndex = -1;
  outer: for (const row of [groupRowRaw, subRowRaw, headerRow]) {
    for (let i = 0; i < row.length; i++) {
      const v = String(row[i] ?? "").toLowerCase();
      if (JOIN_WHERE_KEYS.some((k) => v.includes(k))) {
        joinWhereColIndex = i;
        break outer;
      }
    }
  }
  const joinWhereValues =
    joinWhereColIndex === -1
      ? []
      : [...new Set(dataRows.map((r) => String(r[joinWhereColIndex] ?? "").trim()).filter(Boolean))];

  return { sheetName: match, headers, groups, keys, rows, joinWhereColIndex, joinWhereValues };
}

function parseMappingWorkbookRows(aoa) {
  if (!aoa || aoa.length < 2) return [];
  const headerRow = aoa[0];
  let modelIdx = fuzzyIndex(headerRow, ["dbt model", "model name", "model"]);
  let tabIdx = fuzzyIndex(headerRow, ["stm sheet", "sheet tab", "tab name", "tab"]);
  if (modelIdx === -1) modelIdx = 0;
  if (tabIdx === -1) tabIdx = modelIdx === 0 ? 1 : 0;
  return aoa
    .slice(1)
    .map((r) => ({ model: String(r[modelIdx] ?? "").trim(), tab: String(r[tabIdx] ?? "").trim() }))
    .filter((r) => r.model && r.tab);
}

/* Looks for a "Business Rules" tab inside the STM workbook (searched by
   name, since it's a workbook-wide reference sheet rather than a
   per-model tab) and reads it positionally: Column B = the rule name
   (matching the STM's "General Rule Applied" text), Column D = the
   actual macro/expression that implements that rule. Used by the
   cleansing-rule columns appended to the end of the Column Presence &
   Mapping Check table — if this tab isn't found, those columns still
   populate, just with an explicit note that no Business Rules mapping
   was available. */
function parseBusinessRulesSheet(workbook) {
  if (!workbook) return { found: false };
  const sheetName = workbook.SheetNames.find((n) => /business\s*rule/i.test(n));
  if (!sheetName) return { found: false };
  const sheet = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  const rules = {};
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const ruleName = String(row[1] ?? "").trim(); // Column B
    const macro = String(row[3] ?? "").trim(); // Column D
    if (ruleName) rules[ruleName] = macro;
  }
  return { found: true, sheetName, rules };
}

/* The fixed section groups the "STM Document Details" pane displays, in
   this order, each with its sub-columns in this order. Matching is fuzzy
   on both the section (row1) and sub-header (row3) text so small wording
   differences in the actual STM tab don't break the layout. */
const STM_GROUPS = [
  {
    group: "Target detail",
    groupKeys: ["target detail", "target"],
    fields: [
      { label: "Target Extract", keys: ["target extract"] },
      { label: "Seq", keys: ["seq"] },
      { label: "Column Name", keys: ["column name"] },
      { label: "Datatype", keys: ["datatype"] },
      { label: "SCD Type", keys: ["scd"] },
    ],
  },
  {
    group: "Guidewire PolicyCenter",
    groupKeys: ["guidewire", "policycenter", "gwpc"],
    fields: [
      { label: "Source Table", keys: ["source table", "table"] },
      { label: "Datatype", keys: ["datatype"] },
      { label: "Column Mapping", keys: ["column mapping", "mapping"] },
    ],
  },
  {
    group: "Validation (Data Quality)",
    groupKeys: ["validation", "data quality"],
    fields: [
      { label: "Validation", keys: ["validation"] },
      { label: "Error Action", keys: ["error action", "error"] },
    ],
  },
  {
    group: "Transformation / Cleansing",
    groupKeys: ["transformation", "cleansing", "rules"],
    fields: [
      { label: "Rules", keys: ["rules"] },
      { label: "General Rule Applied", keys: ["general rule"] },
    ],
  },
  {
    group: "Requirement",
    groupKeys: ["requirement"],
    fields: [{ label: "Reason Added", keys: ["reason added", "reason"] }],
  },
  {
    group: "Model Target Definitions",
    groupKeys: ["model target definition"],
    fields: [
      { label: "Name", keys: ["name"] },
      { label: "Code", keys: ["code"] },
      { label: "Domain", keys: ["domain"] },
      { label: "Logical Data Type", keys: ["logical data type", "logical"] },
      { label: "Comment", keys: ["comment"] },
    ],
  },
];

/* The From/Join/Where clause is pulled out of the grouped table entirely
   and shown once, at the top of the pane, as "Join/Where Condition" —
   it describes how the source table is joined, not a per-column value. */
const JOIN_WHERE_KEYS = [
  "from / join / where",
  "from/join/where",
  "join / where",
  "join/where",
  "where clause",
  "join condition",
  "join clause",
  "from clause",
];

/* Builds the grouped column plan for the table from a parsed sheet's
   headers/groups/keys. joinWhereColIndex (found during parsing) is
   excluded here since it's shown separately, not as a grouped column. */
function buildGroupedPlan(headers, groups, keys, joinWhereColIndex) {
  const used = new Set();
  if (joinWhereColIndex !== -1) used.add(joinWhereColIndex);

  const groupedPlan = [];
  STM_GROUPS.forEach((gDef) => {
    const cols = [];
    gDef.fields.forEach((fDef) => {
      let idx = headers.findIndex((h, i) => {
        if (used.has(i)) return false;
        const lh = h.toLowerCase();
        if (!fDef.keys.some((k) => lh.includes(k))) return false;
        const grp = (groups[i] || "").toLowerCase();
        return !grp || gDef.groupKeys.some((k) => grp.includes(k));
      });
      // Fall back to a field-only match if this tab's row1 text doesn't
      // resemble the expected section name at all.
      if (idx === -1) {
        idx = headers.findIndex((h, i) => !used.has(i) && fDef.keys.some((k) => h.toLowerCase().includes(k)));
      }
      if (idx !== -1) {
        used.add(idx);
        cols.push({ label: fDef.label, key: keys[idx] });
      }
    });
    if (cols.length) groupedPlan.push({ group: gDef.group, columns: cols });
  });

  const leftover = [];
  headers.forEach((h, i) => {
    if (!used.has(i)) leftover.push({ label: h, key: keys[i] });
  });
  if (leftover.length) groupedPlan.push({ group: "Other", columns: leftover });

  return groupedPlan;
}

/* ------------------------------------------------------------------ */
/* Promise-wrapped file readers, shared by single-file upload and the  */
/* "select a folder" bulk-upload flow.                                 */
/* ------------------------------------------------------------------ */
function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => resolve(ev.target.result);
    reader.onerror = () => reject(new Error(`Couldn't read ${file.name}`));
    reader.readAsText(file);
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => resolve(ev.target.result);
    reader.onerror = () => reject(new Error(`Couldn't read ${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}

/* Reads a file as base64 (no data: URL prefix) — used to persist the STM
   workbook to storage, since window.storage only accepts strings. */
function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => resolve(String(ev.target.result).split(",")[1] || "");
    reader.onerror = () => reject(new Error(`Couldn't read ${file.name}`));
    reader.readAsDataURL(file);
  });
}

/* Parses a mapping file (csv or xlsx) into { model, tab } rows. Shared by
   the manual mapping-file upload and the folder-select bulk flow. */
async function parseMappingFile(file) {
  if (/\.csv$/i.test(file.name)) {
    const text = await readFileAsText(file);
    const parsed = Papa.parse(text, { skipEmptyLines: true });
    return parseMappingWorkbookRows(parsed.data);
  }
  const buf = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  return parseMappingWorkbookRows(aoa);
}

/* ------------------------------------------------------------------ */
/* Minimal markdown renderer (headings, bold, code, lists, tables)     */
/* ------------------------------------------------------------------ */
function inline(text, keyPrefix) {
  const parts = [];
  const regex = /(\*\*[^*]+\*\*|`[^`]+`)/g;
  let last = 0;
  let m;
  let idx = 0;
  while ((m = regex.exec(text))) {
    if (m.index > last) parts.push(text.slice(last, m.index));
    const token = m[0];
    if (token.startsWith("**")) {
      parts.push(
        <strong key={`${keyPrefix}-b${idx++}`} style={{ color: C.textDark, fontWeight: 600 }}>
          {token.slice(2, -2)}
        </strong>
      );
    } else {
      parts.push(
        <code
          key={`${keyPrefix}-c${idx++}`}
          style={{
            fontFamily: "'IBM Plex Mono', monospace",
            background: "#EEF3F6",
            border: `1px solid ${C.paperLine}`,
            borderRadius: 4,
            padding: "1px 5px",
            fontSize: "0.85em",
          }}
        >
          {token.slice(1, -1)}
        </code>
      );
    }
    last = m.index + token.length;
  }
  if (last < text.length) parts.push(text.slice(last));
  return parts;
}

function statusFromEmoji(text) {
  const t = text.trim();
  if (/^✅/u.test(t)) return { icon: "✓", color: C.match, bg: C.matchBg, rest: t.replace(/^✅\s*/u, "") };
  if (/^⚠️/u.test(t)) return { icon: "!", color: C.warn, bg: C.warnBg, rest: t.replace(/^⚠️\s*/u, "") };
  if (/^❌/u.test(t)) return { icon: "✕", color: C.miss, bg: C.missBg, rest: t.replace(/^❌\s*/u, "") };
  if (/^➕/u.test(t)) return { icon: "+", color: C.extra, bg: C.extraBg, rest: t.replace(/^➕\s*/u, "") };
  return null;
}

function MdList({ items, keyPrefix }) {
  return (
    <ul style={{ listStyle: "none", margin: "0 0 14px 0", padding: 0 }}>
      {items.map((raw, i) => {
        const status = statusFromEmoji(raw);
        const text = status ? status.rest : raw.replace(/^[-*]\s+/, "");
        return (
          <li
            key={`${keyPrefix}-li${i}`}
            style={{
              display: "flex",
              alignItems: "flex-start",
              gap: 10,
              padding: "7px 0",
              borderBottom: `1px solid ${C.paperLine}`,
              fontSize: 14,
              lineHeight: 1.5,
              color: C.textDark,
            }}
          >
            <span
              style={{
                flexShrink: 0,
                width: 20,
                height: 20,
                borderRadius: 5,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontFamily: "'IBM Plex Mono', monospace",
                fontSize: 12,
                fontWeight: 600,
                marginTop: 1,
                color: status ? status.color : C.accentDim,
                background: status ? status.bg : "#EAF2F6",
              }}
            >
              {status ? status.icon : "•"}
            </span>
            <span>{inline(text, `${keyPrefix}-li${i}`)}</span>
          </li>
        );
      })}
    </ul>
  );
}

function splitTableRow(line) {
  const trimmed = line.trim().replace(/^\|/, "").replace(/\|$/, "");
  // Split only on pipes that AREN'T escaped with a backslash — Markdown's
  // convention for a literal "|" inside a table cell. Without this, a cell
  // containing SQL's "||" concatenation operator (or any other literal "|")
  // gets misread as extra column separators, distorting the whole row.
  return trimmed.split(/(?<!\\)\|/).map((c) => c.trim().replace(/\\\|/g, "|"));
}

function MdTable({ lines, keyPrefix }) {
  const rows = lines.map(splitTableRow);
  const isSep = (cells) => cells.every((c) => /^:?-{2,}:?$/.test(c));
  const header = rows[0];
  const body = isSep(rows[1] || []) ? rows.slice(2) : rows.slice(1);
  return (
    <div style={{ overflowX: "auto", margin: "10px 0 18px 0", border: `1px solid ${C.paperLine}`, borderRadius: 8 }}>
      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
        <thead>
          <tr>
            {header.map((h, i) => (
              <th
                key={`${keyPrefix}-th${i}`}
                style={{
                  textAlign: "left",
                  padding: "8px 12px",
                  background: C.ink,
                  color: C.textLight,
                  fontFamily: "'Space Grotesk', sans-serif",
                  fontWeight: 600,
                  fontSize: 12,
                  letterSpacing: 0.3,
                  whiteSpace: "nowrap",
                }}
              >
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {body.map((r, ri) => (
            <tr key={`${keyPrefix}-tr${ri}`} style={{ background: ri % 2 ? "#FAFCFD" : C.paperPanel }}>
              {r.map((cell, ci) => {
                const status = statusFromEmoji(cell);
                return (
                  <td
                    key={`${keyPrefix}-td${ri}-${ci}`}
                    style={{
                      padding: "7px 12px",
                      borderTop: `1px solid ${C.paperLine}`,
                      color: status ? status.color : C.textDark,
                      fontWeight: status ? 600 : 400,
                      verticalAlign: "top",
                      background: status ? status.bg : "transparent",
                    }}
                  >
                    {status ? (
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                        <span
                          style={{
                            width: 16,
                            height: 16,
                            borderRadius: 4,
                            background: status.bg,
                            color: status.color,
                            display: "inline-flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: 10,
                            fontFamily: "'IBM Plex Mono', monospace",
                          }}
                        >
                          {status.icon}
                        </span>
                        {inline(status.rest, `${keyPrefix}-td${ri}-${ci}`)}
                      </span>
                    ) : (
                      inline(cell, `${keyPrefix}-td${ri}-${ci}`)
                    )}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function renderMarkdown(md) {
  const lines = (md || "").split("\n");
  const out = [];
  let i = 0;
  let key = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (/^\s*#{1,4}\s+/.test(line)) {
      const level = line.match(/^\s*(#{1,4})/)[1].length;
      const text = line.replace(/^\s*#{1,4}\s+/, "");
      const sizes = { 1: 20, 2: 17, 3: 15, 4: 14 };
      out.push(
        <div
          key={`h${key++}`}
          style={{
            fontFamily: "'Space Grotesk', sans-serif",
            fontWeight: 600,
            fontSize: sizes[level] || 14,
            color: C.ink,
            marginTop: level <= 2 ? 22 : 14,
            marginBottom: 8,
            paddingBottom: level <= 2 ? 8 : 0,
            borderBottom: level <= 2 ? `2px solid ${C.paperLine}` : "none",
          }}
        >
          {inline(text, `h${key}`)}
        </div>
      );
      i++;
      continue;
    }
    if (/^\s*\|.*\|\s*$/.test(line)) {
      const tableLines = [];
      while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) {
        tableLines.push(lines[i]);
        i++;
      }
      out.push(<MdTable key={`t${key++}`} lines={tableLines} keyPrefix={`t${key}`} />);
      continue;
    }
    if (/^\s*[-*]\s+/.test(line) || /^\s*(✅|⚠️|❌|➕)/u.test(line)) {
      const items = [];
      while (
        i < lines.length &&
        (/^\s*[-*]\s+/.test(lines[i]) || /^\s*(✅|⚠️|❌|➕)/u.test(lines[i]))
      ) {
        items.push(lines[i].trim());
        i++;
      }
      out.push(<MdList key={`l${key++}`} items={items} keyPrefix={`l${key}`} />);
      continue;
    }
    if (line.trim() === "") {
      i++;
      continue;
    }
    const paraLines = [line];
    i++;
    while (
      i < lines.length &&
      lines[i].trim() !== "" &&
      !/^\s*#{1,4}\s+/.test(lines[i]) &&
      !/^\s*\|.*\|\s*$/.test(lines[i]) &&
      !/^\s*[-*]\s+/.test(lines[i]) &&
      !/^\s*(✅|⚠️|❌|➕)/u.test(lines[i])
    ) {
      paraLines.push(lines[i]);
      i++;
    }
    out.push(
      <p key={`p${key++}`} style={{ fontSize: 14, lineHeight: 1.6, color: C.textDark, margin: "0 0 12px 0" }}>
        {inline(paraLines.join(" "), `p${key}`)}
      </p>
    );
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Comparison prompt builder                                           */
/* ------------------------------------------------------------------ */
function buildPromptPreamble({ stmRows, sheetName, dbtModelName, dbtSql, businessRules }) {
  const businessRulesSection = businessRules && businessRules.found
    ? `### INPUT 3: Business Rules mapping (from the "${businessRules.sheetName}" tab in the STM workbook)
Column B of that tab is the rule name (matches the STM's "General Rule Applied" text); Column D ("Macro used") is only the macro-TYPE argument for that rule (e.g. \`VARCHAR_NOKEY\`) — the full expected DBT call is \`m_cleanse('<Column D value>', '<the target column's own name>')\`, NOT Column D by itself.

${JSON.stringify(businessRules.rules, null, 2)}`
    : `### INPUT 3: Business Rules mapping
No "Business Rules" tab was found in the uploaded STM workbook. Treat every column as having no defined cleansing macro — i.e. a direct load is expected for every column unless the STM's "General Rule Applied" text plus the DBT SQL itself clearly indicate otherwise.`;

  return `You are a meticulous Data Engineering QA reviewer specializing in Guidewire PolicyCenter (GWPC) to DBT ETL migrations. Compare the STM specification against the DBT model code and report discrepancies precisely.

### INPUT 1: STM Specification — sheet tab "${sheetName}" (parsed rows)
Each row is a JSON object keyed by its actual column headers as found in the sheet (Target Column Name, Target Datatype, SCD Type, Source Table(s), mapping expression, join/where clause, data quality rule, validation/error action, general rule applied, etc. — whichever of these exist as columns in this tab).

${JSON.stringify(stmRows, null, 2)}

### INPUT 2: DBT Model Code — model "${dbtModelName}"
\`\`\`sql
${dbtSql}
\`\`\`

${businessRulesSection}
`;
}

/* Part 1 of the comparison prompt — sections 1-2. These are the two
   sections that need to see every column at once (cross-referencing one
   column's mapping against every other column's), so they stay together
   in one call. Split out into its own call (run in parallel with Part 2)
   so each response only has to carry roughly half the report. */
function buildPromptPart1({ stmRows, sheetName, dbtModelName, dbtSql, businessRules }) {
  return buildPromptPreamble({ stmRows, sheetName, dbtModelName, dbtSql, businessRules }) + `
### YOUR TASK

1. **Entity Verification** — compare the join structure (FROM/JOIN tables, join types, and aliases) stated in the STM's Join/Where condition against what's actually implemented in the DBT model's SQL. Build a table with these columns, in this order: **S.No | Table Name | Join Type in STM | Alias in STM | Join Type in DBT | Alias in DBT | Status**.

   **Parse the STM side first:** find the Join/Where condition text among the STM rows above (it may appear under a key containing "Join" or "Where"). From that text, identify every source table it references: which one is the base table (the one after \`FROM\`, with no join keyword), and which ones are brought in via \`INNER JOIN\`, \`LEFT JOIN\`, \`RIGHT JOIN\`, \`FULL JOIN\`, or \`CROSS JOIN\` — plus the alias each table is given there. Example of the expected shape:
   | Table Name | Join Type | Alias |
   |---|---|---|
   | Table A | Base table | A |
   | Table B | Left Join | B |

   **Parse the DBT side the same way:** read the actual \`FROM\`/\`JOIN\` clauses in the DBT SQL to get the same three facts per table — base table vs. join type, and the alias actually used in the code.

   **Then, for every table found in the STM's join structure:**
   - Look for that same table (by actual table name, not alias) anywhere in the DBT SQL's FROM/JOIN clauses.
   - If found: fill in Join Type in DBT and Alias in DBT from what the SQL actually does, and set Status:
     - "✅ Match" — same join type AND same alias in both.
     - "⚠️ Alias mismatch" — join type matches but the alias differs (state both aliases in the Status text).
     - "⚠️ Join type mismatch" — alias matches but the join type differs (state both types).
     - "⚠️ Join type and alias mismatch" — both differ (state all four values).
   - If NOT found anywhere in the DBT SQL's FROM/JOIN clauses: Join Type in DBT = "—", Alias in DBT = "—", Status = "❌ Missing in DBT".

   **Then, for every table that appears in the DBT SQL's FROM/JOIN clauses but is NOT mentioned anywhere in the STM's Join/Where condition:** add it as its own row — Table Name = that DBT table, Join Type in STM = "—", Alias in STM = "—", Status = "❌ Extra table in DBT, not referenced in STM".

   Number every row sequentially in S.No. After the table, add a one-line **Summary**, e.g. "All 3 STM tables are present in DBT with matching join type and alias" or "2 of 3 STM tables match; 1 has an alias mismatch; 1 extra table exists in DBT that isn't in the STM."

2. **Column Presence & Column Mapping Check** — mandatory, and just as important as section 1. Build a single table with EXACTLY these columns, in this order: **S.No | STM Target Column | DBT Column/Alias Name | Column Name Status | Column Mapping in STM | Column Mapping in DBT | Column Mapping Status | Cleansing Rule in STM | Cleansing Rule in DBT | Cleansing Rule Status | SCD Type in STM | SCD Type in DBT | SCD Status**.

   **Name matching — keep this simple, no inference:**
   - List every single target column from the STM rows above, with no omissions. For each one, check whether the exact same column/alias name (trivial casing/whitespace differences are fine) exists in the DBT model's SELECT list.
     - If it exists in DBT: DBT Column/Alias Name = that name, Column Name Status = "✅ Present in Column Name".
     - If it does NOT exist in DBT under that name: DBT Column/Alias Name = "—", Column Name Status = "❌ Missing in DBT".
   - Then, for every column/alias in the DBT model's SELECT list that does NOT match any STM target column by name: add it as its own row — STM Target Column = "—", DBT Column/Alias Name = that DBT column name, Column Name Status = "❌ Invalid column in DBT".
   - Do NOT attempt to guess that a differently-named column is "actually" a renamed match, and do NOT suggest renames. If the names don't match exactly, that's Missing/Invalid — nothing more. Leave any judgment about renamed/equivalent columns to the person reading the report.
   - Number every row sequentially in S.No, starting at 1, across the full table (STM-sourced rows and DBT-only rows share one sequential numbering).

   **Column Mapping comparison — applies to EVERY row in the table, no exceptions, including Missing/Invalid rows:**
   - Column Mapping in STM = the STM row's Column Mapping / source-table-and-column / expression text exactly as written in the sheet, **excluding any alias** — just the underlying source/expression. If this row has no STM column at all (a DBT-only row), this is "—".
   - Column Mapping in DBT = the actual SQL expression/logic that produces that column in the DBT model (source table, source column, transformation, casts — as literally implemented), **excluding the column alias** — e.g. write \`job.CloseDate\`, not \`job.CloseDate AS ROW_PROC_DTS\`; strip any trailing \`AS <alias>\` before putting it in the table. **Treat the entire expression as ONE atomic mapping**, no matter how many lines, operators, CASTs, or concatenation segments (\`||\`) it spans in the raw SQL — e.g. a multi-part key like \`'GWPC' || '-' || CAST(polper.PERIODID AS VARCHAR) || '-' || 'Truck' || '-' || CAST(catruck.FIXEDID AS VARCHAR)\` is a single concatenation mapping, not several separate ones. **Write it out as one flat single line** in the table cell — collapse any line breaks and extra indentation from the source SQL into single spaces — since a literal line break inside a Markdown table cell breaks the table's rendering. If this row has no DBT column matched by name at all, this is "—".
   - Column Mapping Status — run this exact three-step check for EVERY row, no shortcuts and never leave it blank or write "N/A":
     1. **Same-row check.** If both Column Mapping in STM and Column Mapping in DBT are present (not "—") for this row, and they express the same underlying logic (same source table/column, same transformation, same casts): "✅ Match".
     2. **Cross-check against every other DBT mapping.** If step 1 didn't produce a match (either because the same-row expressions differ, or because one side was "—" for this row — e.g. a "Missing in DBT" row still has an STM mapping to check), take this row's Column Mapping in STM (if it has one) and compare it against the Column Mapping in DBT value of **every other row in this table** — for example, take \`LIC_JURS_CD\`'s STM mapping and check it against every other row's DBT mapping value, one by one. If any of them express the same logic: "⚠️ This mapping logic is matching with '<that other row's DBT column name>' in DBT".
     3. **Cross-check against every other STM mapping.** If step 2 found nothing, take this row's Column Mapping in DBT (if it has one) and compare it against the Column Mapping in STM value of every other row in the table. If any of them express the same logic: "⚠️ This mapping logic is matching with '<that other row's STM column name>' in STM".
     4. **Only if steps 1–3 all find nothing**: "❌ Not Match". This is the only fallback status — never output "N/A" or leave the cell empty, for any row, for any reason.

   **Cleansing rule comparison — applies to EVERY row in the table, no exceptions (use INPUT 3, the Business Rules mapping, which appears earlier in this prompt):**
   - Cleansing Rule in STM = that STM row's "General Rule Applied" value exactly as written (or "—" if blank, or "—" for a DBT-only row that has no STM column at all) — this is a rule NAME like "General Foreign Key Rule 1", not a macro call.
   - Determine what DBT is expected to do: look up Cleansing Rule in STM's text (the rule name) in INPUT 3's Business Rules mapping to get its "Macro used" value (e.g. \`VARCHAR_NOKEY\`).
     - If Cleansing Rule in STM is blank, OR it's non-blank but its looked-up "Macro used" value in INPUT 3 is blank/not found: the column is expected to be a **direct load** (selected straight from source, no cleansing macro).
     - If Cleansing Rule in STM is non-blank AND its looked-up "Macro used" value in INPUT 3 is non-blank: **that "Macro used" value is only the macro-TYPE argument, not the whole expected call.** Construct the full expected DBT expression yourself as \`m_cleanse('<Macro used value>', '<this row's own STM Target Column name>')\` — e.g. for General Foreign Key Rule 1 on column \`POL_KEY\`, the expected call is \`m_cleanse('VARCHAR_NOKEY', 'POL_KEY')\`. The column name argument is always THIS row's own target column name, never a hardcoded or borrowed one. Some columns may have this wrapped in an outer function too (e.g. \`UPPER(m_cleanse(...))\`) — that outer wrapping is fine and still counts as a match, since the wrapping is a per-column stylistic choice, not part of the rule itself; what actually matters is that the macro type and its column argument are correct.
   - Cleansing Rule in DBT = the actual cleansing logic found in the DBT SQL for this column: the macro call and its arguments if one is used (including any outer wrapping like \`UPPER(...)\`), "Direct load" if selected with no cleansing macro, or "—" if the column isn't present in DBT at all.
   - Cleansing Rule Status:
     - "✅ Match" if what DBT actually does (the constructed macro call — allowing for outer-wrapping differences — or direct load) matches the expectation above.
     - "❌ Not a Match" if a macro was expected but DBT uses a different macro, a different column argument, or no macro at all; if a direct load was expected but a macro was applied anyway; or if the column isn't present in DBT at all.
     - "—" only for a DBT-only row that has no STM column at all (nothing to check).

   **SCD Type comparison — applies to EVERY row in the table, no exceptions:**
   - SCD Type in STM = this row's "SCD Type (Column I)" field, exactly as written in the sheet (e.g. \`1\`, \`2\`, \`N/A\`) — this was read directly from Column I of the sheet, independent of whatever the header text there says. Use "—" if blank, or for a DBT-only row that has no STM column at all.
   - Find which of the DBT model's SCD column lists — \`scd1_cols\`, \`scd2_cols\`, or \`key_cols\` (however they're defined in the SQL, e.g. Jinja \`{% set scd1_cols = [...] %}\` blocks or similar) — actually contain this row's column name (matched the same way as STM Target Column vs. DBT Column/Alias Name above). SCD Type in DBT = the name of whichever block it's actually found in (e.g. \`scd1_cols\`), "Not found in any block" if it's in none of them, or a list of all of them if it appears in more than one (that's itself an anomaly to flag) — or "—" if this column isn't present in DBT at all.
   - SCD Status — **do not use the words "Match" or "Not a Match" anywhere in this column.** Describe what's actually happening instead:
     - If SCD Type in STM is \`1\`: the expected block is \`scd1_cols\`. If found there (and only there): "✅ Present in scd1_cols as expected". Otherwise: "❌ Expected in scd1_cols but " + state where it's actually found instead (another block, more than one block, or "not found in any block").
     - If SCD Type in STM is \`2\`: same idea, expected block is \`scd2_cols\`.
     - If SCD Type in STM is \`N/A\`: the expected block is \`key_cols\`. **Always call this row out explicitly regardless of whether it's correct** — state which block it's actually found in either way: "⚠️ N/A in STM — present in key_cols as expected" if correctly placed, or "❌ N/A in STM — expected in key_cols but " + where it's actually found instead.
     - If SCD Type in STM is anything else (blank, or a value that isn't \`1\`, \`2\`, or \`N/A\`): "❌ Unrecognized SCD Type value '<value>' in STM — cannot verify against DBT".
     - If this is a DBT-only row with no STM column, or the column is entirely absent from DBT: "—".

   After the table, add a one-line **Summary** stating the counts, e.g. "12 of 15 STM columns are present in DBT; 3 are missing; 1 DBT column is invalid; of all 16 rows, 9 have matching column mapping logic, 3 have logic matching a different column, and 4 do not match at all; separately, 10 of 15 columns match their expected cleansing behavior and 5 do not; and 13 of 15 columns are correctly classified in their expected SCD block while 2 are not."

Be literal — verify against what the code actually does, not what a column name implies. Do not assume unmapped/NULL fields are errors if the STM explicitly marks them as not applicable for this coverable type — check any "Model Target Definitions" / general-rule comment column before flagging.

CRITICAL FORMATTING RULE: Markdown tables use "|" as the column separator. Any literal "|" character inside a table cell's content — including SQL's "||" concatenation operator, which will appear often in the Column Mapping in DBT column — MUST be escaped as "\\|" (backslash-pipe) so it is not misread as extra column separators.

Output format: structured Markdown with ONLY the 2 sections above (use ##  headings), each finding as a bullet point prefixed with ✅, ⚠️, or ❌ as appropriate, plus tables as Markdown tables. Do not add any other sections — a second, separate call handles Risk & Recommendations.`;
}

/* Part 2 of the comparison prompt — Risk & Recommendations only.
   Checklist Verification, the standalone Discrepancy Table, Extra/
   Unmapped DBT Logic, and the standalone General Rule Logic Validation
   table have all been removed from the report — General Rule Logic
   Validation's three columns (Cleansing Rule in STM / in DBT / Status)
   now live at the end of Part 1's Column Presence & Mapping Check table
   instead of their own section. */
function buildPromptPart2({ stmRows, sheetName, dbtModelName, dbtSql, businessRules }) {
  return buildPromptPreamble({ stmRows, sheetName, dbtModelName, dbtSql, businessRules }) + `
### YOUR TASK

**Risk & Recommendations** — for every mismatch you would find between the STM and the DBT model (column presence, column mapping logic, or cleansing rule logic), state the downstream data-quality risk and a concrete corrected DBT snippet. Cover the same ground a full column-by-column and cleansing-rule review of this model would surface — don't wait for a separate table to point you at issues, work it out directly from INPUT 1/2/3 above.

Be literal — verify against what the code actually does, not what a column name implies. Do not assume unmapped/NULL fields are errors if the STM explicitly marks them as not applicable for this coverable type.

CRITICAL FORMATTING RULE: Markdown tables use "|" as the column separator. Any literal "|" character inside a table cell's content — including SQL's "||" concatenation operator — MUST be escaped as "\\|" (backslash-pipe).

Output format: structured Markdown with a single "## Risk & Recommendations" heading, each finding as a bullet point prefixed with ⚠️ or ❌ as appropriate. Do not add any other sections — a separate call already generated the Entity Verification and Column Presence & Mapping Check sections.`;
}

/* ------------------------------------------------------------------ */
/* Main app                                                             */
/* ------------------------------------------------------------------ */
export default function App() {
  const [stmWorkbook, setStmWorkbook] = useState(null);
  const [stmFileName, setStmFileName] = useState("");
  const [stmError, setStmError] = useState("");
  const [stmFromStorage, setStmFromStorage] = useState(false);
  const [stmTooLargeToSave, setStmTooLargeToSave] = useState(false);

  const [folderError, setFolderError] = useState("");
  const [folderSummary, setFolderSummary] = useState(null); // { mappingFound, sqlTotal, sqlMatched }
  const [folderLoading, setFolderLoading] = useState(false);

  const [mappingRows, setMappingRows] = useState([]);
  const [mappingFileName, setMappingFileName] = useState("");
  const [mappingError, setMappingError] = useState("");
  const [mappingFromStorage, setMappingFromStorage] = useState(false);
  const [mappingStorageChecked, setMappingStorageChecked] = useState(false);
  const [mappingSource, setMappingSource] = useState(""); // "folder" | "manual" | ""

  const [selectedModel, setSelectedModel] = useState("");
  const [stmCache, setStmCache] = useState({}); // tabName -> {headers, rows, sheetName, error}

  const [dbtSql, setDbtSql] = useState("");
  const [dbtFileName, setDbtFileName] = useState("");
  const [dbtSqlByModel, setDbtSqlByModel] = useState({}); // model -> { sql, fileName }
  const [sqlModelSuggestion, setSqlModelSuggestion] = useState(null); // { model, sql, fileName } awaiting user confirmation

  const [activeTab, setActiveTab] = useState("details");
  const [comparisonCache, setComparisonCache] = useState({}); // model -> markdown
  const [comparing, setComparing] = useState(false);
  const [compareError, setCompareError] = useState("");

  const [search, setSearch] = useState("");

  const stmInputRef = useRef(null);
  const mappingInputRef = useRef(null);
  const sqlInputRef = useRef(null);
  const folderInputRef = useRef(null);

  /* Load the mapping file from persistent storage once, on mount, so it
     doesn't have to be re-uploaded every session. The upload button stays
     available so it can still be replaced whenever the mapping changes. */
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const stored = await window.storage?.get?.("model-tab-mapping", false);
        if (!cancelled && stored?.value) {
          const parsed = JSON.parse(stored.value);
          if (parsed?.rows?.length) {
            setMappingRows(parsed.rows);
            setMappingFileName(parsed.fileName || "saved mapping");
            setMappingFromStorage(true);
            setMappingSource("manual");
          }
        }
      } catch (e) {
        // no saved mapping yet — that's fine, user uploads one below
      } finally {
        if (!cancelled) setMappingStorageChecked(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const saveMappingToStorage = useCallback(async (rows, fileName) => {
    try {
      await window.storage?.set?.("model-tab-mapping", JSON.stringify({ rows, fileName }), false);
    } catch (e) {
      // non-fatal — mapping still works for this session even if save fails
    }
  }, []);

  /* Load a previously-saved STM workbook from storage on mount, same idea
     as the mapping file. Best-effort: storage values are capped (~5MB),
     so a large workbook simply won't have saved — that's fine, the app
     falls back to asking for a normal upload with no functional loss. */
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const stored = await window.storage?.get?.("stm-workbook", false);
        if (!cancelled && stored?.value) {
          const parsed = JSON.parse(stored.value);
          if (parsed?.base64) {
            const wb = XLSX.read(parsed.base64, { type: "base64" });
            setStmWorkbook(wb);
            setStmFileName(parsed.fileName || "saved STM workbook");
            setStmFromStorage(true);
          }
        }
      } catch (e) {
        // no saved workbook, or it didn't fit — user uploads normally
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const saveStmToStorage = useCallback(async (file) => {
    try {
      const base64 = await readFileAsBase64(file);
      // Storage values are capped around 5MB; skip the save rather than
      // risk a failed/partial write for large workbooks.
      if (base64.length > 4.5 * 1024 * 1024) {
        setStmTooLargeToSave(true);
        return;
      }
      setStmTooLargeToSave(false);
      await window.storage?.set?.("stm-workbook", JSON.stringify({ base64, fileName: file.name }), false);
    } catch (e) {
      setStmTooLargeToSave(true);
    }
  }, []);

  const resolvedTab = useMemo(() => {
    const m = mappingRows.find((r) => r.model === selectedModel);
    return m ? m.tab : "";
  }, [mappingRows, selectedModel]);

  const stmParsed = resolvedTab ? stmCache[resolvedTab] : null;

  const filteredRows = useMemo(() => {
    if (!stmParsed || !stmParsed.rows) return [];
    if (!search.trim()) return stmParsed.rows;
    const q = search.toLowerCase();
    return stmParsed.rows.filter((r) => Object.values(r).some((v) => String(v).toLowerCase().includes(q)));
  }, [stmParsed, search]);

  /* ---- file handlers ---- */
  const handleStmUpload = useCallback(
    (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      setStmError("");
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const wb = XLSX.read(ev.target.result, { type: "array" });
          setStmWorkbook(wb);
          setStmFileName(file.name);
          setStmFromStorage(false);
          setStmCache({});
          saveStmToStorage(file);
        } catch (err) {
          setStmError("Couldn't read that workbook. Make sure it's a valid .xlsx or .xlsm file.");
        }
      };
      reader.onerror = () => setStmError("Couldn't read that file from disk.");
      reader.readAsArrayBuffer(file);
    },
    [saveStmToStorage]
  );

  const handleMappingUpload = useCallback((e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setMappingError("");
    const isCsv = /\.csv$/i.test(file.name);
    if (isCsv) {
      const reader = new FileReader();
      reader.onload = (ev) => {
        const parsed = Papa.parse(ev.target.result, { skipEmptyLines: true });
        const rows = parseMappingWorkbookRows(parsed.data);
        if (!rows.length) {
          setMappingError("No model/tab rows could be found in that CSV.");
          return;
        }
        setMappingRows(rows);
        setMappingFileName(file.name);
        setMappingFromStorage(false);
        setMappingSource("manual");
        saveMappingToStorage(rows, file.name);
      };
      reader.readAsText(file);
    } else {
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const wb = XLSX.read(ev.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          const rows = parseMappingWorkbookRows(aoa);
          if (!rows.length) {
            setMappingError("No model/tab rows could be found in that workbook.");
            return;
          }
          setMappingRows(rows);
          setMappingFileName(file.name);
          setMappingFromStorage(false);
          setMappingSource("manual");
          saveMappingToStorage(rows, file.name);
        } catch (err) {
          setMappingError("Couldn't read that mapping file.");
        }
      };
      reader.readAsArrayBuffer(file);
    }
  }, [saveMappingToStorage]);

  /* Bulk "select a folder" flow: point at a project folder once (e.g. one
     that contains a mapping file plus a "Queries" subfolder of .sql files)
     and it reads everything inside — including subfolders — in one shot.
     This is a one-click multi-file read via the browser's folder picker,
     not a persistent link to the folder: browsers don't allow a web app to
     silently watch your local disk, so you click this again whenever the
     folder's contents change. The STM workbook is deliberately NOT part of
     this — that one you said you'll keep updating by hand. */
  const handleFolderSelect = useCallback(
    (e) => {
      const files = Array.from(e.target.files || []);
      if (!files.length) return;
      setFolderError("");
      setFolderSummary(null);
      setFolderLoading(true);

      (async () => {
        try {
          const sqlFiles = files.filter((f) => /\.sql$/i.test(f.name));
          const mappingCandidates = files.filter(
            (f) => /\.(xlsx|xlsm|csv)$/i.test(f.name) && !/\.sql$/i.test(f.name)
          );
          const mappingFile =
            mappingCandidates.find((f) => /map/i.test(f.name)) || mappingCandidates[0] || null;

          let effectiveMappingRows = mappingRows;
          let mappingFound = false;

          if (mappingFile) {
            const rows = await parseMappingFile(mappingFile);
            if (rows.length) {
              effectiveMappingRows = rows;
              mappingFound = true;
              setMappingRows(rows);
              setMappingFileName(mappingFile.name);
              setMappingFromStorage(false);
              setMappingSource("folder");
              saveMappingToStorage(rows, mappingFile.name);
            }
          }

          const matchedEntries = {};
          let matchedCount = 0;
          for (const f of sqlFiles) {
            const baseName = f.name.replace(/\.[^.]+$/, "").trim().toLowerCase();
            const model = effectiveMappingRows.find((r) => r.model.trim().toLowerCase() === baseName)?.model;
            if (!model) continue;
            const text = await readFileAsText(f);
            matchedEntries[model] = { sql: text, fileName: f.name };
            matchedCount++;
          }

          if (matchedCount) {
            setDbtSqlByModel((prev) => ({ ...prev, ...matchedEntries }));
            if (selectedModel && matchedEntries[selectedModel]) {
              setDbtSql(matchedEntries[selectedModel].sql);
              setDbtFileName(matchedEntries[selectedModel].fileName);
            }
          }

          setFolderSummary({ mappingFound, sqlTotal: sqlFiles.length, sqlMatched: matchedCount });
        } catch (err) {
          setFolderError("Couldn't read that folder. Try selecting it again.");
        } finally {
          setFolderLoading(false);
        }
      })();
    },
    [mappingRows, selectedModel, saveMappingToStorage]
  );

  /* ---- resolve + parse STM tab whenever model changes ---- */
  const ensureStmParsed = useCallback(
    (tab) => {
      if (!tab || !stmWorkbook) return;
      if (stmCache[tab]) return;
      const result = parseStmSheet(stmWorkbook, tab);
      setStmCache((prev) => ({ ...prev, [tab]: result }));
    },
    [stmWorkbook, stmCache]
  );

  const handleSelectModel = (model) => {
    // Stash whatever SQL is currently in the box under the model that's
    // being switched away from, so coming back to it later restores it.
    setDbtSqlByModel((prev) =>
      selectedModel ? { ...prev, [selectedModel]: { sql: dbtSql, fileName: dbtFileName } } : prev
    );
    setSelectedModel(model);
    setSearch("");
    setSqlModelSuggestion(null);
    // SQL & model are linked: picking a model auto-populates whatever SQL
    // was previously provided for it (upload or paste), if any.
    const cached = dbtSqlByModel[model];
    setDbtSql(cached ? cached.sql : "");
    setDbtFileName(cached ? cached.fileName : "");
    const m = mappingRows.find((r) => r.model === model);
    if (m) ensureStmParsed(m.tab);
    setActiveTab("details");
  };

  /* Keep the per-model SQL cache in sync as the user types/pastes, so
     switching models and back doesn't lose in-progress edits. */
  const handleDbtSqlChange = useCallback(
    (value) => {
      setDbtSql(value);
      if (selectedModel) {
        setDbtSqlByModel((prev) => ({ ...prev, [selectedModel]: { sql: value, fileName: dbtFileName } }));
      }
    },
    [selectedModel, dbtFileName]
  );

  const handleSqlUpload = useCallback(
    (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        const text = ev.target.result;
        const baseName = file.name.replace(/\.[^.]+$/, "").trim().toLowerCase();
        // SQL & model selection are linked, but only ever auto-switched
        // with the user's confirmation — never silently.
        const matchedModel = mappingRows.find((r) => r.model.trim().toLowerCase() === baseName)?.model;
        if (matchedModel && matchedModel !== selectedModel) {
          setSqlModelSuggestion({ model: matchedModel, sql: text, fileName: file.name });
          return;
        }
        setDbtSql(text);
        setDbtFileName(file.name);
        if (selectedModel) {
          setDbtSqlByModel((prev) => ({ ...prev, [selectedModel]: { sql: text, fileName: file.name } }));
        }
      };
      reader.readAsText(file);
    },
    [mappingRows, selectedModel]
  );

  const acceptSqlModelSuggestion = useCallback(() => {
    if (!sqlModelSuggestion) return;
    const { model, sql, fileName } = sqlModelSuggestion;
    setDbtSqlByModel((prev) =>
      selectedModel ? { ...prev, [selectedModel]: { sql: dbtSql, fileName: dbtFileName } } : prev
    );
    setSelectedModel(model);
    setSearch("");
    setDbtSql(sql);
    setDbtFileName(fileName);
    setDbtSqlByModel((prev) => ({ ...prev, [model]: { sql, fileName } }));
    const m = mappingRows.find((r) => r.model === model);
    if (m) ensureStmParsed(m.tab);
    setActiveTab("details");
    setSqlModelSuggestion(null);
  }, [sqlModelSuggestion, selectedModel, dbtSql, dbtFileName, mappingRows, ensureStmParsed]);

  const dismissSqlModelSuggestion = useCallback(() => setSqlModelSuggestion(null), []);

  const canCompare = stmParsed && stmParsed.rows && !stmParsed.error && dbtSql.trim().length > 0 && selectedModel;

  const runCompare = useCallback(
    async (force) => {
      if (!canCompare) return;
      if (!force && comparisonCache[selectedModel]) {
        setActiveTab("comparison");
        return;
      }
      setComparing(true);
      setCompareError("");
      try {
        // Re-read the STM workbook fresh every time a comparison starts,
        // rather than relying on whatever was cached when the model was
        // first selected — the mapping file is a one-time upload, but the
        // STM workbook's contents are what must always be current.
        const freshStm = parseStmSheet(stmWorkbook, resolvedTab);
        if (freshStm.error) {
          setCompareError(freshStm.error);
          setComparing(false);
          return;
        }
        setStmCache((prev) => ({ ...prev, [resolvedTab]: freshStm }));
        const businessRules = parseBusinessRulesSheet(stmWorkbook);
        const promptArgs = {
          stmRows: freshStm.rows,
          sheetName: freshStm.sheetName,
          dbtModelName: selectedModel,
          dbtSql,
          businessRules,
        };
        const callPart = async (prompt) => {
          const response = await fetch("https://api.anthropic.com/v1/messages", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              model: "claude-sonnet-4-6",
              max_tokens: 16000,
              messages: [{ role: "user", content: prompt }],
            }),
          });
          const data = await response.json();
          return (data.content || [])
            .map((b) => (b.type === "text" ? b.text : ""))
            .filter(Boolean)
            .join("\n");
        };
        // Run both halves of the report in parallel — Part 1 (Entity
        // Verification + Column Presence & Mapping Check) needs to see
        // every column at once to cross-reference, Part 2 (General Rule
        // Logic Validation + Checklist + Discrepancy + Extra + Risk)
        // doesn't depend on Part 1's output, so splitting them means
        // each response only has to carry roughly half the report.
        const [text1, text2] = await Promise.all([
          callPart(buildPromptPart1(promptArgs)),
          callPart(buildPromptPart2(promptArgs)),
        ]);
        const text = [text1, text2].filter(Boolean).join("\n\n");
        if (!text) throw new Error("empty response");
        setComparisonCache((prev) => ({ ...prev, [selectedModel]: text }));
        setActiveTab("comparison");
      } catch (err) {
        setCompareError("The comparison call failed. Please try again.");
      } finally {
        setComparing(false);
      }
    },
    [canCompare, comparisonCache, selectedModel, stmWorkbook, resolvedTab, dbtSql]
  );

  const comparisonMarkdown = comparisonCache[selectedModel] || "";

  /* ------------------------------------------------------------------ */
  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans', sans-serif",
        background: C.paper,
        minHeight: "100vh",
        color: C.textDark,
      }}
    >
      <style>{FONTS}</style>

      {/* Top bar */}
      <div
        style={{
          background: C.inkDark,
          borderBottom: `1px solid ${C.inkLineSoft}`,
          backgroundImage:
            "linear-gradient(rgba(255,255,255,0.035) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.035) 1px, transparent 1px)",
          backgroundSize: "22px 22px",
        }}
      >
        <div style={{ padding: "16px 24px 14px 24px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
            <div
              style={{
                width: 30,
                height: 30,
                borderRadius: 7,
                background: C.accent,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                flexShrink: 0,
              }}
            >
              <GitCompare size={17} color={C.inkDark} />
            </div>
            <div>
              <div
                style={{
                  fontFamily: "'Space Grotesk', sans-serif",
                  fontWeight: 700,
                  fontSize: 17,
                  color: C.textLight,
                  letterSpacing: 0.2,
                }}
              >
                STM ↔ DBT Verifier
              </div>
              <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 1 }}>
                Source-to-Target Mapping vs. model implementation, side by side
              </div>
            </div>
          </div>

          <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" }}>
            {/* Step 1: STM workbook — always a manual, single-file upload */}
            <ToolbarButton
              icon={<FileSpreadsheet size={14} />}
              label={stmFileName ? `${stmFileName}${stmFromStorage ? " (saved)" : ""}` : "1 · Upload STM workbook"}
              sub={
                stmFromStorage
                  ? "loaded from last session — click to replace"
                  : stmTooLargeToSave
                  ? ".xlsx / .xlsm — too large to auto-save, re-upload each session"
                  : ".xlsx / .xlsm — saved automatically when it fits"
              }
              onClick={() => stmInputRef.current?.click()}
              done={!!stmFileName}
            />
            <input ref={stmInputRef} type="file" accept=".xlsx,.xlsm" onChange={handleStmUpload} style={{ display: "none" }} />

            {/* Step 2: Model folder — the .sql files (and, if present, the mapping file) */}
            <ToolbarButton
              icon={folderLoading ? <Loader2 size={14} className="animate-spin" /> : <Upload size={14} />}
              label={
                folderLoading
                  ? "Reading folder…"
                  : folderSummary
                  ? `Model folder loaded · ${folderSummary.sqlMatched}/${folderSummary.sqlTotal} SQL matched`
                  : "2 · Upload model folder (.sql files)"
              }
              sub="reads every .sql file inside, subfolders included — also checks for the mapping file"
              onClick={() => folderInputRef.current?.click()}
              done={!!folderSummary}
            />
            <input
              ref={folderInputRef}
              type="file"
              webkitdirectory=""
              directory=""
              multiple
              onChange={handleFolderSelect}
              style={{ display: "none" }}
            />

            {/* Step 3: STM–Model mapping — auto-detected from the folder above if it's
                in there, otherwise a manual upload button appears instead. */}
            {mappingRows.length > 0 && mappingSource === "folder" ? (
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  padding: "8px 12px",
                  borderRadius: 8,
                  border: `1px solid ${C.accent}`,
                  background: C.accentDim,
                  fontSize: 12.5,
                  color: C.textLight,
                }}
              >
                <ClipboardList size={14} color={C.accent} />
                <span>
                  3 · Mapping: {mappingFileName} <span style={{ color: C.textMuted }}>(auto-detected in folder)</span>
                </span>
                <button
                  onClick={() => mappingInputRef.current?.click()}
                  style={{
                    border: "none",
                    background: "transparent",
                    color: C.accent,
                    fontSize: 11.5,
                    textDecoration: "underline",
                    cursor: "pointer",
                    padding: 0,
                  }}
                >
                  replace manually
                </button>
              </div>
            ) : (
              <ToolbarButton
                icon={<ClipboardList size={14} />}
                label={
                  mappingFileName
                    ? `${mappingFileName}${mappingFromStorage ? " (saved)" : ""}`
                    : "3 · Upload STM–Model mapping"
                }
                sub={
                  mappingFromStorage
                    ? "loaded from last session — click to replace"
                    : folderSummary && !folderSummary.mappingFound
                    ? "not found in the model folder — upload it manually"
                    : ".xlsx / .csv"
                }
                onClick={() => mappingInputRef.current?.click()}
                done={!!mappingFileName}
              />
            )}
            <input
              ref={mappingInputRef}
              type="file"
              accept=".xlsx,.xlsm,.csv"
              onChange={handleMappingUpload}
              style={{ display: "none" }}
            />

            <div style={{ width: 1, height: 28, background: C.inkLineSoft, margin: "0 2px" }} />

            <ModelDropdown
              options={mappingRows.map((r) => r.model)}
              value={selectedModel}
              onChange={handleSelectModel}
              disabled={!mappingRows.length}
            />

            {/* Manual fallback for a single model's SQL, used when the model
                folder didn't contain a matching .sql file for it. */}
            <ToolbarButton
              icon={<FileCode2 size={14} />}
              label={dbtFileName || "Upload this model's .sql manually"}
              sub="only needed if the model folder didn't have a matching file"
              onClick={() => sqlInputRef.current?.click()}
              done={!!dbtFileName}
            />
            <input ref={sqlInputRef} type="file" accept=".sql,.txt" onChange={handleSqlUpload} style={{ display: "none" }} />

            <button
              onClick={() => runCompare(false)}
              disabled={!canCompare || comparing}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 7,
                padding: "9px 16px",
                borderRadius: 8,
                border: "none",
                fontFamily: "'Space Grotesk', sans-serif",
                fontWeight: 600,
                fontSize: 13,
                cursor: canCompare && !comparing ? "pointer" : "not-allowed",
                background: canCompare ? C.accent : C.inkLineSoft,
                color: canCompare ? C.inkDark : C.textMuted,
                opacity: comparing ? 0.7 : 1,
              }}
            >
              {comparing ? <Loader2 size={15} className="animate-spin" /> : <GitCompare size={15} />}
              {comparing ? "Comparing…" : "Compare"}
            </button>
          </div>

          {(stmError || mappingError || folderError) && (
            <div style={{ marginTop: 10, fontSize: 12, color: "#F3A79E" }}>{stmError || mappingError || folderError}</div>
          )}

          {folderSummary && !folderError && (
            <div style={{ marginTop: 10, fontSize: 12, color: C.textMuted }}>
              {folderSummary.mappingFound && "Mapping file found and loaded. "}
              {folderSummary.sqlTotal
                ? `${folderSummary.sqlMatched} of ${folderSummary.sqlTotal} .sql files in that folder matched a model name and were loaded.${
                    folderSummary.sqlMatched < folderSummary.sqlTotal
                      ? " Unmatched files' names didn't match any model in the mapping."
                      : ""
                  }`
                : "No .sql files were found in that folder."}
            </div>
          )}

          {sqlModelSuggestion && (
            <div
              style={{
                marginTop: 10,
                display: "flex",
                alignItems: "center",
                gap: 10,
                fontSize: 12.5,
                color: C.textLight,
                background: C.accentDim,
                border: `1px solid ${C.accent}`,
                borderRadius: 8,
                padding: "8px 12px",
              }}
            >
              <span>
                "{sqlModelSuggestion.fileName}" looks like it belongs to <strong>{sqlModelSuggestion.model}</strong>, not the
                currently selected model. Switch to it?
              </span>
              <button
                onClick={acceptSqlModelSuggestion}
                style={{
                  padding: "4px 10px",
                  borderRadius: 6,
                  border: "none",
                  background: C.accent,
                  color: C.inkDark,
                  fontWeight: 600,
                  fontSize: 12,
                  cursor: "pointer",
                  flexShrink: 0,
                }}
              >
                Switch model
              </button>
              <button
                onClick={dismissSqlModelSuggestion}
                style={{
                  padding: "4px 10px",
                  borderRadius: 6,
                  border: `1px solid ${C.inkLineSoft}`,
                  background: "transparent",
                  color: C.textMuted,
                  fontSize: 12,
                  cursor: "pointer",
                  flexShrink: 0,
                }}
              >
                Keep current
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Body */}
      <div style={{ display: "flex", minHeight: "calc(100vh - 132px)" }}>
        {/* Left nav */}
        <div
          style={{
            width: 216,
            flexShrink: 0,
            background: C.ink,
            borderRight: `1px solid ${C.inkLineSoft}`,
            padding: "18px 12px",
          }}
        >
          <div style={{ fontSize: 10.5, letterSpacing: 1, color: C.textMuted, fontWeight: 600, padding: "0 8px 8px 8px" }}>
            {selectedModel ? selectedModel : "NO MODEL SELECTED"}
          </div>
          <NavItem
            active={activeTab === "details"}
            onClick={() => setActiveTab("details")}
            icon={<ClipboardList size={15} />}
            label="STM Document Details"
          />
          <NavItem
            active={activeTab === "comparison"}
            onClick={() => setActiveTab("comparison")}
            icon={<GitCompare size={15} />}
            label="Comparison Result"
            badge={comparisonMarkdown ? "•" : null}
          />

          <div style={{ marginTop: 22, padding: "12px 10px", borderRadius: 8, background: C.inkDark, border: `1px solid ${C.inkLineSoft}` }}>
            <div style={{ fontSize: 10.5, color: C.textMuted, lineHeight: 1.6 }}>
              <StepLine done={!!stmFileName} label="Upload STM workbook" />
              <StepLine done={!!mappingRows.length} label="Upload mapping file" />
              <StepLine done={!!selectedModel} label="Select a DBT model" />
              <StepLine done={!!dbtSql.trim()} label="Provide DBT SQL" />
              <StepLine done={!!comparisonMarkdown} label="Run comparison" />
            </div>
          </div>
        </div>

        {/* Main content */}
        <div style={{ flex: 1, padding: "22px 26px", overflow: "auto" }}>
          {activeTab === "details" && (
            <DetailsPane
              selectedModel={selectedModel}
              resolvedTab={resolvedTab}
              stmParsed={stmParsed}
              filteredRows={filteredRows}
              search={search}
              setSearch={setSearch}
              dbtSql={dbtSql}
              setDbtSql={handleDbtSqlChange}
              stmWorkbook={stmWorkbook}
              mappingRows={mappingRows}
              folderUsed={!!folderSummary}
            />
          )}
          {activeTab === "comparison" && (
            <ComparisonPane
              comparing={comparing}
              compareError={compareError}
              comparisonMarkdown={comparisonMarkdown}
              canCompare={canCompare}
              selectedModel={selectedModel}
              onRerun={() => runCompare(true)}
            />
          )}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Sub-components                                                       */
/* ------------------------------------------------------------------ */
function ToolbarButton({ icon, label, sub, onClick, done }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 8,
        padding: "8px 12px",
        borderRadius: 8,
        border: `1px solid ${done ? C.accent : C.inkLineSoft}`,
        background: done ? C.accentDim : "transparent",
        color: done ? C.textLight : C.textMuted,
        fontSize: 12.5,
        cursor: "pointer",
        maxWidth: 220,
      }}
      title={sub}
    >
      <span style={{ color: done ? C.accent : C.textMuted, flexShrink: 0 }}>{icon}</span>
      <span
        style={{
          overflow: "hidden",
          textOverflow: "ellipsis",
          whiteSpace: "nowrap",
          textAlign: "left",
        }}
      >
        {label}
      </span>
    </button>
  );
}

function ModelDropdown({ options, value, onChange, disabled }) {
  return (
    <div style={{ position: "relative" }}>
      <select
        value={value}
        disabled={disabled}
        onChange={(e) => onChange(e.target.value)}
        style={{
          appearance: "none",
          padding: "9px 30px 9px 12px",
          borderRadius: 8,
          border: `1px solid ${value ? C.accent : C.inkLineSoft}`,
          background: disabled ? "transparent" : C.ink,
          color: disabled ? C.textMuted : C.textLight,
          fontSize: 12.5,
          fontFamily: "'IBM Plex Mono', monospace",
          minWidth: 220,
          cursor: disabled ? "not-allowed" : "pointer",
        }}
      >
        <option value="">{disabled ? "Load mapping file first…" : "Select DBT model…"}</option>
        {options.map((m) => (
          <option key={m} value={m}>
            {m}
          </option>
        ))}
      </select>
      <ChevronDown size={14} color={C.textMuted} style={{ position: "absolute", right: 10, top: 11, pointerEvents: "none" }} />
    </div>
  );
}

function NavItem({ active, onClick, icon, label, badge }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 9,
        width: "100%",
        textAlign: "left",
        padding: "10px 10px",
        borderRadius: 7,
        border: "none",
        background: active ? C.accentDim : "transparent",
        color: active ? C.textLight : C.textMuted,
        fontSize: 13,
        fontWeight: active ? 600 : 400,
        cursor: "pointer",
        marginBottom: 4,
      }}
    >
      <span style={{ color: active ? C.accent : C.textMuted }}>{icon}</span>
      <span style={{ flex: 1 }}>{label}</span>
      {badge && <span style={{ color: C.accent, fontSize: 16, lineHeight: 1 }}>{badge}</span>}
    </button>
  );
}

function StepLine({ done, label }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7, padding: "2px 0" }}>
      <span
        style={{
          width: 6,
          height: 6,
          borderRadius: 3,
          background: done ? C.accent : C.inkLineSoft,
          flexShrink: 0,
        }}
      />
      <span style={{ color: done ? C.textLight : C.textMuted }}>{label}</span>
    </div>
  );
}

function DetailsPane({
  selectedModel,
  resolvedTab,
  stmParsed,
  filteredRows,
  search,
  setSearch,
  dbtSql,
  setDbtSql,
  stmWorkbook,
  mappingRows,
  folderUsed,
}) {
  if (!selectedModel) {
    return (
      <EmptyState
        title="Pick a model to inspect its spec"
        body="Upload the STM workbook, then choose a DBT model from the dropdown above. STM Document Details organizes that model's STM rows into fixed sections — Target detail, Guidewire PolicyCenter, Validation (Data Quality), Transformation / Cleansing, Requirement, and Model Target Definitions — with the Join/Where condition called out separately at the top."
      />
    );
  }
  if (!stmWorkbook) {
    return <EmptyState title="STM workbook not loaded" body="Upload the STM .xlsx/.xlsm file so its sheet tabs can be parsed." />;
  }
  if (!stmParsed) {
    return <EmptyState title="Resolving sheet tab…" body={`Looking for a tab named "${resolvedTab}".`} />;
  }
  if (stmParsed.error) {
    return <EmptyState title="Couldn't parse this tab" body={stmParsed.error} warn />;
  }

  const groupedPlan = buildGroupedPlan(stmParsed.headers, stmParsed.groups, stmParsed.keys, stmParsed.joinWhereColIndex);
  const totalCols = groupedPlan.reduce((n, g) => n + g.columns.length, 0);
  const joinWhereValues = stmParsed.joinWhereValues || [];

  return (
    <div>
      <SectionHeader
        eyebrow={`SHEET TAB · ${stmParsed.sheetName}`}
        title={selectedModel}
        meta={`${stmParsed.rows.length} mapped column${stmParsed.rows.length === 1 ? "" : "s"}`}
      />

      {joinWhereValues.length > 0 && (
        <div
          style={{
            marginTop: 14,
            padding: "12px 14px",
            borderRadius: 10,
            border: `1px solid ${C.accent}`,
            background: C.accentDim,
          }}
        >
          <div style={{ fontSize: 10.5, letterSpacing: 1, color: C.accent, fontWeight: 600, marginBottom: 6 }}>
            JOIN/WHERE CONDITION
          </div>
          {joinWhereValues.map((v, i) => (
            <div
              key={i}
              style={{
                fontFamily: "'IBM Plex Mono', monospace",
                fontSize: 12.5,
                color: C.textLight,
                lineHeight: 1.6,
                whiteSpace: "pre-wrap",
                marginBottom: i < joinWhereValues.length - 1 ? 8 : 0,
              }}
            >
              {v}
            </div>
          ))}
        </div>
      )}

      <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "14px 0" }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            padding: "8px 12px",
            borderRadius: 8,
            border: `1px solid ${C.paperLine}`,
            background: C.paperPanel,
            flex: 1,
            maxWidth: 340,
          }}
        >
          <Search size={14} color={C.textMutedDark} />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Filter rows (column name, table, rule…)"
            style={{ border: "none", outline: "none", fontSize: 13, flex: 1, fontFamily: "inherit" }}
          />
        </div>
      </div>

      <ScrollSyncedTable groupedPlan={groupedPlan} filteredRows={filteredRows} totalCols={totalCols} search={search} />

      <div style={{ marginTop: 22 }}>
        <SectionHeader eyebrow="DBT MODEL CODE" title="Paste or upload the .sql for this model" meta="" small />
        {folderUsed && !dbtSql.trim() && (
          <div
            style={{
              marginTop: 10,
              padding: "8px 12px",
              borderRadius: 8,
              border: `1px solid ${C.warn}`,
              background: C.warnBg,
              color: C.warn,
              fontSize: 12.5,
            }}
          >
            No .sql file for "{selectedModel}" was found in the uploaded model folder — upload it manually using the
            button in the top bar, or paste it in below.
          </div>
        )}
        <textarea
          value={dbtSql}
          onChange={(e) => setDbtSql(e.target.value)}
          placeholder={`-- paste the full ${selectedModel}.sql here (CTEs, mapping, cleansing, SCD config)…`}
          style={{
            width: "100%",
            minHeight: 220,
            marginTop: 10,
            padding: 14,
            borderRadius: 10,
            border: `1px solid ${C.paperLine}`,
            background: C.paperPanel,
            fontFamily: "'IBM Plex Mono', monospace",
            fontSize: 12.5,
            lineHeight: 1.6,
            resize: "vertical",
            boxSizing: "border-box",
          }}
        />
      </div>
    </div>
  );
}

/* A wide grouped table needs a horizontal scrollbar, but the browser only
   puts one at the very bottom of the scroll area — on a long table that
   means scrolling all the way down just to scroll sideways. This renders
   a second, slim scrollbar above the table that's kept in sync with the
   real one below, so horizontal scrolling works from the top too. */
function ScrollSyncedTable({ groupedPlan, filteredRows, totalCols, search }) {
  const topRef = useRef(null);
  const bottomRef = useRef(null);
  const [scrollWidth, setScrollWidth] = useState(0);
  const syncing = useRef(false);

  useEffect(() => {
    const el = bottomRef.current;
    if (!el) return;
    const update = () => setScrollWidth(el.scrollWidth);
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, [groupedPlan, filteredRows]);

  const onTopScroll = () => {
    if (syncing.current) return;
    syncing.current = true;
    if (bottomRef.current && topRef.current) bottomRef.current.scrollLeft = topRef.current.scrollLeft;
    syncing.current = false;
  };
  const onBottomScroll = () => {
    if (syncing.current) return;
    syncing.current = true;
    if (topRef.current && bottomRef.current) topRef.current.scrollLeft = bottomRef.current.scrollLeft;
    syncing.current = false;
  };

  return (
    <div>
      <div
        ref={topRef}
        onScroll={onTopScroll}
        style={{ overflowX: "auto", overflowY: "hidden", height: 14, marginBottom: -1 }}
      >
        <div style={{ width: scrollWidth, height: 1 }} />
      </div>
      <div
        ref={bottomRef}
        onScroll={onBottomScroll}
        style={{
          overflowX: "auto",
          border: `1px solid ${C.paperLine}`,
          borderRadius: 10,
          background: C.paperPanel,
        }}
      >
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12.5 }}>
          <thead>
            <tr>
              {groupedPlan.map((g) => (
                <th
                  key={g.group}
                  colSpan={g.columns.length}
                  style={{
                    position: "sticky",
                    top: 0,
                    textAlign: "left",
                    padding: "8px 12px",
                    background: C.inkDark,
                    color: C.accent,
                    fontFamily: "'Space Grotesk', sans-serif",
                    fontWeight: 600,
                    fontSize: 11,
                    letterSpacing: 0.4,
                    whiteSpace: "nowrap",
                    borderLeft: `1px solid ${C.inkLineSoft}`,
                  }}
                >
                  {g.group.toUpperCase()}
                </th>
              ))}
            </tr>
            <tr>
              {groupedPlan.map((g) =>
                g.columns.map((col, ci) => (
                  <th
                    key={`${col.key}-${ci}`}
                    style={{
                      position: "sticky",
                      top: 29,
                      textAlign: "left",
                      padding: "9px 12px",
                      background: C.ink,
                      color: C.textLight,
                      fontFamily: "'Space Grotesk', sans-serif",
                      fontWeight: 600,
                      fontSize: 11.5,
                      whiteSpace: "nowrap",
                      letterSpacing: 0.2,
                      borderLeft: ci === 0 ? `1px solid ${C.inkLineSoft}` : "none",
                    }}
                  >
                    {col.label}
                  </th>
                ))
              )}
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((r, i) => (
              <tr key={i} style={{ background: i % 2 ? "#FAFCFD" : C.paperPanel }}>
                {groupedPlan.map((g) =>
                  g.columns.map((col, ci) => (
                    <td
                      key={`${col.key}-${ci}`}
                      style={{
                        padding: "7px 12px",
                        borderTop: `1px solid ${C.paperLine}`,
                        borderLeft: ci === 0 ? `1px solid ${C.paperLine}` : "none",
                        whiteSpace: "nowrap",
                        maxWidth: 320,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        fontFamily: /column|table|mapping|rule/i.test(col.label)
                          ? "'IBM Plex Mono', monospace"
                          : "inherit",
                      }}
                      title={r[col.key]}
                    >
                      {r[col.key] || "—"}
                    </td>
                  ))
                )}
              </tr>
            ))}
            {filteredRows.length === 0 && (
              <tr>
                <td colSpan={totalCols} style={{ padding: 18, color: C.textMutedDark, textAlign: "center" }}>
                  No rows match "{search}".
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ComparisonPane({ comparing, compareError, comparisonMarkdown, canCompare, selectedModel, onRerun }) {
  if (comparing) {
    return (
      <EmptyState
        title="Running the comparison…"
        body={`Checking ${selectedModel}'s DBT logic against its STM spec column by column. This usually takes a few seconds.`}
        loading
      />
    );
  }
  if (compareError) {
    return <EmptyState title="Comparison failed" body={compareError} warn />;
  }
  if (!comparisonMarkdown) {
    return (
      <EmptyState
        title="No comparison run yet"
        body={
          canCompare
            ? "Everything needed is loaded — hit Compare in the top bar to generate the discrepancy report."
            : "Upload the STM workbook, the mapping file, select a model, and provide its DBT SQL, then hit Compare."
        }
      />
    );
  }
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
        <SectionHeader eyebrow="AI-GENERATED" title={`Comparison — ${selectedModel}`} meta="" small />
        <button
          onClick={onRerun}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "7px 12px",
            borderRadius: 7,
            border: `1px solid ${C.paperLine}`,
            background: C.paperPanel,
            fontSize: 12,
            color: C.textMutedDark,
            cursor: "pointer",
          }}
        >
          <RefreshCw size={13} /> Re-run comparison
        </button>
      </div>
      <div
        style={{
          background: C.paperPanel,
          border: `1px solid ${C.paperLine}`,
          borderRadius: 10,
          padding: "18px 20px",
        }}
      >
        {renderMarkdown(comparisonMarkdown)}
      </div>
    </div>
  );
}

function SectionHeader({ eyebrow, title, meta, small }) {
  return (
    <div>
      <div style={{ fontSize: 10.5, letterSpacing: 1, color: C.accentDim, fontWeight: 600, marginBottom: 3 }}>{eyebrow}</div>
      <div
        style={{
          display: "flex",
          alignItems: "baseline",
          gap: 10,
          fontFamily: "'Space Grotesk', sans-serif",
          fontWeight: 600,
          fontSize: small ? 15 : 19,
          color: C.textDark,
        }}
      >
        {title}
        {meta && <span style={{ fontSize: 12, fontWeight: 400, color: C.textMutedDark }}>{meta}</span>}
      </div>
    </div>
  );
}

function EmptyState({ title, body, warn, loading }) {
  return (
    <div
      style={{
        marginTop: 40,
        maxWidth: 480,
        marginLeft: "auto",
        marginRight: "auto",
        textAlign: "center",
        padding: "36px 30px",
        borderRadius: 12,
        border: `1px dashed ${C.paperLine}`,
        background: C.paperPanel,
      }}
    >
      <div style={{ display: "flex", justifyContent: "center", marginBottom: 14 }}>
        {loading ? (
          <Loader2 size={26} color={C.accent} className="animate-spin" />
        ) : warn ? (
          <AlertTriangle size={26} color={C.warn} />
        ) : (
          <PlusCircle size={26} color={C.textMutedDark} />
        )}
      </div>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, fontSize: 16, marginBottom: 8, color: C.textDark }}>
        {title}
      </div>
      <div style={{ fontSize: 13, color: C.textMutedDark, lineHeight: 1.6 }}>{body}</div>
    </div>
  );
}
