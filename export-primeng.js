// ============================================
// PrimeNG Table Export - Excel Report Code Generator
// ============================================
// Generates Angular component code using PrimeNG's <p-table>
// ============================================

function exportToNgPrime({ ws: wsJson, workbook: wbJson, meta, images = [] }) {
  const ws = buildWorksheetFromJson(wsJson);
  const workbook = buildWorkbookFromJson(wbJson);
  const themeColors = getThemeColors(workbook);
  const sections = buildSections(meta, ws.rowCount || ws.lastRow?.number || 0);

  // ===== FIND SECTIONS =====
  const titleSection = sections.find(s => s.key === 'title');
  const columnHeaderSection = sections.find(s => s.key === 'columnheader');
  const detailSection = sections.find(s => s.key === 'detail');
  const pageHeaderSection = sections.find(s => s.key === 'pageheader');
  const pageFooterSection = sections.find(s => s.key === 'pagefooter');
  const summarySection = sections.find(s => s.key === 'summary');

  // ===== BUILD MERGE MAP (colspan/rowspan) =====
  const mergeMap = {};
  if (ws.model?.merges) {
    ws.model.merges.forEach(m => {
      const [start, end] = m.split(':');
      const s = decodeAddr(start);
      const e = decodeAddr(end);
      const masterKey = `${s.r},${s.c}`;
      const colspan = e.c - s.c + 1;
      const rowspan = e.r - s.r + 1;
      mergeMap[masterKey] = { colspan, rowspan };
      // Mark all covered cells (skip)
      for (let r = s.r; r <= e.r; r++) {
        for (let c = s.c; c <= e.c; c++) {
          if (r !== s.r || c !== s.c) {
            mergeMap[`${r},${c}`] = { skip: true };
          }
        }
      }
    });
  }

  // ===== BUILD HEADER ROWS (with colspan/rowspan) =====
  const headerRows = [];
  const columns = [];

  if (columnHeaderSection) {
    for (let r = columnHeaderSection.start; r <= columnHeaderSection.end; r++) {
      const row = ws.getRow(r);
      const cells = [];
      for (let c = 1; c <= ws.columnCount; c++) {
        const key = `${r},${c}`;
        const info = mergeMap[key];
        if (info?.skip) continue; // skip non-master cells in a merge

        const cell = row.getCell(c);
        const text = (cell.text || '').trim();
        cells.push({
          text,
          colspan: info?.colspan || 1,
          rowspan: info?.rowspan || 1,
          align: cell.alignment?.horizontal || 'left',
          style: buildNgStyle(cell, themeColors),
          field: text ? toCamelCase(text) : `col${c}`
        });
      }
      headerRows.push(cells);
    }

    // Columns = last header row (after resolving merges)
    // Expand colspan > 1 cells into multiple column entries
    // Deduplicate field names (e.g., two "Last Year" columns → lastYear, lastYear_2)
    const lastRow = headerRows[headerRows.length - 1];
    const usedFields = {};
    lastRow.forEach(cell => {
      const span = cell.colspan || 1;
      for (let i = 0; i < span; i++) {
        let baseField = cell.field + (span > 1 ? `_${i + 1}` : '');
        // Deduplicate field name
        if (usedFields[baseField]) {
          let counter = 2;
          while (usedFields[baseField + `_${counter}`]) counter++;
          baseField = baseField + `_${counter}`;
        }
        usedFields[baseField] = true;
        columns.push({
          field: baseField,
          header: span > 1 ? `${cell.text} ${i + 1}` : (cell.text || `Column ${columns.length + 1}`),
          align: cell.align
        });
      }
    });
  } else {
    // No columnheader section — generate generic columns
    for (let c = 1; c <= ws.columnCount; c++) {
      columns.push({
        field: `col${c}`,
        header: `Column ${c}`,
        align: 'left'
      });
    }
    // Single header row from generic column names
    headerRows.push(columns.map(col => ({
      text: col.header,
      colspan: 1,
      rowspan: 1,
      field: col.field,
      align: 'left'
    })));
  }

  // ===== COLLECT DETAIL TEMPLATE ROWS (multi-line detail) =====
  // Every row of the detail section is a template row that repeats per record.
  const detailTemplateRows = [];
  if (detailSection) {
    for (let r = detailSection.start; r <= detailSection.end; r++) {
      const row = ws.getRow(r);
      const fields = [];
      row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
        const text = (cell.text || '').trim();
        const fieldMatch = text.match(/\{\{(.+?)\}\}/);
        const paramMatch = text.match(/^[pP]\{\{(.+?)\}\}$/);
        if (fieldMatch) {
          let rawField = fieldMatch[1].trim();
          let fieldName = rawField;
          let pipeOverride = null;
          // รองรับ syntax: {{fieldName:pipeType}}
          // pipeType = number | date | datetime | text
          // ถ้ากำหนด pipeType และไม่ใช่ 'general' → ใช้ pipeType แทน inferNgPipe
          // fieldName ยังคงเป็นชื่อฟีลด์เดิมเสมอ
          const colonIdx = rawField.indexOf(':');
          if (colonIdx > 0) {
            fieldName = rawField.substring(0, colonIdx).trim();
            const typePart = rawField.substring(colonIdx + 1).trim();
            if (typePart && typePart.toLowerCase() !== 'general') {
              pipeOverride = typePart.toLowerCase();
            }
          }
          // infer pipe type จาก cell number format (Excel format)
          const pipeFromNumFmt = inferPipeFromNumFmt(cell.numFmt);
          fields.push({
            colIndex: cIdx,
            fieldName: fieldName,
            pipeOverride: pipeOverride,
            pipeFromNumFmt: pipeFromNumFmt,
            type: 'field'
          });
        } else if (paramMatch) {
          fields.push({
            colIndex: cIdx,
            fieldName: paramMatch[1].trim(),
            type: 'parameter'
          });
        } else if (text) {
          fields.push({
            colIndex: cIdx,
            fieldName: text,
            type: 'static'
          });
        }
      });
      detailTemplateRows.push({ rowIndex: r, fields });
    }
  }

  // Flattened fields (kept for getNgType compatibility)
  const detailFields = detailTemplateRows.flatMap(tr => tr.fields);

  // ===== COLLECT PARAMETERS =====
  const params = [];
  const allParams = collectAllParams(ws);
  allParams.forEach(p => {
    if (!params.find(x => x.name === p.name)) {
      params.push(p);
    }
  });

  // ===== BUILD DETAIL CELL MAP (rowIdx → colIndex → cell info) =====
  // Uses the actual {{fieldName}} / static text from each detail template row
  // so the body renders exactly what the Excel detail section defines.
  const detailCellMap = {};
  detailTemplateRows.forEach((tr, ri) => {
    detailCellMap[ri] = detailCellMap[ri] || {};
    tr.fields.forEach(df => {
      if (df.type === 'field') {
        // Priority: explicit :pipeType > cell numFmt inference > inferNgPipe(fieldName)
        let pipe = null;
        if (df.pipeOverride) pipe = df.pipeOverride;
        else if (df.pipeFromNumFmt) pipe = df.pipeFromNumFmt;
        detailCellMap[ri][df.colIndex] = { kind: 'field', fieldName: df.fieldName, pipe };
      } else if (df.type === 'parameter') {
        detailCellMap[ri][df.colIndex] = { kind: 'field', fieldName: df.fieldName, pipe: null };
      } else {
        detailCellMap[ri][df.colIndex] = { kind: 'static', text: df.fieldName };
      }
    });
  });

  // ===== BUILD DETAIL STYLE MAP (rowIdx → colIndex → inline CSS) =====
  // Captures font/background colors from each detail template row
  const detailStyleMap = {};
  if (detailSection) {
    for (let r = detailSection.start; r <= detailSection.end; r++) {
      const rowIdx = r - detailSection.start;
      detailStyleMap[rowIdx] = detailStyleMap[rowIdx] || {};
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
        // Skip Column A markers and empty cells — just capture visible styles
        const style = buildNgStyle(cell, themeColors);
        const css = styleObjToCss(style);
        if (css) {
          detailStyleMap[rowIdx][cIdx] = css;
        }
      });
    }
  }

  // ===== COUNT TOTAL DATA COLUMNS =====
  // Resolve last header row to count actual data columns (expand colspan)
  let dataColCount = 0;
  if (headerRows.length > 0) {
    const lastRow = headerRows[headerRows.length - 1];
    lastRow.forEach(cell => { dataColCount += cell.colspan; });
  } else {
    dataColCount = columns.length;
  }

  // ===== HTML TEMPLATE (.html) =====
  let html = [];

  // Title section
  if (titleSection) {
    html.push(`<!-- Title Section -->`);
    html.push(`<div class="report-title">`);
    html.push(`  <ng-container *ngIf="params">`);

    for (let r = titleSection.start; r <= titleSection.end; r++) {
      const row = ws.getRow(r);
      let cellsHtml = [];
      row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
        if (cIdx === 1) return;
        const text = (cell.text || '').trim();
        if (!text) return;
        const processed = processNgText(text);
        cellsHtml.push(`    <span>${processed}</span>`);
      });
      if (cellsHtml.length > 0) {
        html.push(`  <div class="title-row">${cellsHtml.join(' ')}</div>`);
      }
    }
    html.push(`  </ng-container>`);
    html.push(`</div>`);
    html.push(``);
  }

  // Page Header section
  if (pageHeaderSection) {
    html.push(`<!-- Page Header Section -->`);
    html.push(`<div class="page-header">`);
    for (let r = pageHeaderSection.start; r <= pageHeaderSection.end; r++) {
      const row = ws.getRow(r);
      let cellsHtml = [];
      row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
        if (cIdx === 1) return;
        const text = (cell.text || '').trim();
        if (!text) return;
        const processed = processNgText(text);
        cellsHtml.push(`    <span>${processed}</span>`);
      });
      if (cellsHtml.length > 0) {
        html.push(`  <div class="header-row">${cellsHtml.join(' ')}</div>`);
      }
    }
    html.push(`</div>`);
    html.push(``);
  }

  // PrimeNG Table
  html.push(`<!-- PrimeNG Table -->`);
  html.push(`<p-table [value]="data" [tableStyle]="{'min-width': '50rem'}" [resizableColumns]="true">`);
  html.push(``);

  // Header (multi-row with colspan/rowspan + Excel styles)
  html.push(`  <ng-template pTemplate="header">`);
  headerRows.forEach((row, ri) => {
    html.push(`    <tr>`);
    row.forEach(cell => {
      const attrs = [];
      if (cell.colspan > 1) attrs.push(`colspan="${cell.colspan}"`);
      if (cell.rowspan > 1) attrs.push(`rowspan="${cell.rowspan}"`);
      const inlineStyle = styleObjToCss(cell.style);
      if (inlineStyle) attrs.push(`style="${inlineStyle}"`);
      const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';
      html.push(`      <th${attrStr}>${escapeHtml(cell.text)}</th>`);
    });
    html.push(`    </tr>`);
  });
  html.push(`  </ng-template>`);
  html.push(``);

  // Body (with Excel styles from each detail template row — multi-line detail)
  // One <tr> is emitted per detail template row, so every row repeats per record.
  const detailRowCount = detailTemplateRows.length > 0 ? detailTemplateRows.length : 1;
  html.push(`  <ng-template pTemplate="body" let-row>`);
  for (let ri = 0; ri < detailRowCount; ri++) {
    html.push(`    <tr>`);
    columns.forEach((col, idx) => {
      const wsCol = idx + 1; // 1-based column index in cloned worksheet
      const cellInfo = (detailCellMap[ri] && detailCellMap[ri][wsCol]) || null;
      const inlineStyle = (detailStyleMap[ri] && detailStyleMap[ri][wsCol]) || '';
      const styleAttr = inlineStyle ? ` style="${inlineStyle}"` : '';
      let expr;
      let alignClass = '';
      if (cellInfo && cellInfo.kind === 'static') {
        // Static text repeats literally in every record (HTML-escaped).
        // Static cells never contain {{...}} (those are classified as 'field'),
        // so only escapeHtml is needed — processNgText would double-escape '@'.
        expr = escapeHtml(cellInfo.text);
      } else {
        const fieldName = (cellInfo && cellInfo.fieldName) || col.field;
        const pipe = (cellInfo && cellInfo.pipe) || inferNgPipe(fieldName);
        alignClass = pipe === 'number' ? 'text-right' : (pipe === 'date' || pipe === 'datetime' ? 'text-center' : '');
        if (pipe === 'number') {
          expr = `{{row.${fieldName} | number : '1.2-2'}}`;
        } else if (pipe === 'datetime') {
          expr = `{{row.${fieldName} | date:'dd/MM/yyyy HH:mm'}}`;
        } else if (pipe === 'date') {
          expr = `{{row.${fieldName} | date:'dd/MM/yyyy'}}`;
        } else {
          expr = `{{row.${fieldName}}}`;
        }
      }
      const classAttr = alignClass ? ` class="${alignClass}"` : '';
      html.push(`      <td${styleAttr}${classAttr}>${expr}</td>`);
    });
    html.push(`    </tr>`);
  }
  html.push(`  </ng-template>`);
  html.push(``);

  // Footer (from summary section, with Excel styles)
  // Note: cloned worksheet has NO Column A (section key was removed),
  // so column indices start at 1 = first data column.
  if (summarySection) {
    html.push(`  <ng-template pTemplate="footer">`);
    for (let r = summarySection.start; r <= summarySection.end; r++) {
      const row = ws.getRow(r);
      html.push(`    <tr>`);
      // Collect all non-empty summary cells with their styles
      let summaryCells = [];
      row.eachCell({ includeEmpty: false }, (cell, cIdx) => {
        const text = (cell.text || '').trim();
        if (!text) return;
        const processed = processNgText(text);
        const style = buildNgStyle(cell, themeColors);
        const css = styleObjToCss(style);
        summaryCells.push({ text: processed, col: cIdx, style: css });
      });
      // Render with colspan: each cell spans the gap to the next cell
      summaryCells.forEach((cell, idx) => {
        const nextCol = idx < summaryCells.length - 1 ? summaryCells[idx + 1].col : dataColCount + 1;
        const span = nextCol - cell.col;
        const colspan = span > 1 ? ` colspan="${span}"` : '';
        const styleAttr = cell.style ? ` style="${cell.style}"` : '';
        const alignClass = idx === 0 ? ' class="text-right"' : '';
        html.push(`      <td${colspan}${styleAttr}${alignClass}>${cell.text}</td>`);
      });
      html.push(`    </tr>`);
    }
    html.push(`  </ng-template>`);
    html.push(``);
  }

  html.push(`</p-table>`);
  html.push(``);

  // Page Footer section
  if (pageFooterSection) {
    html.push(`<!-- Page Footer Section -->`);
    html.push(`<div class="page-footer">`);
    for (let r = pageFooterSection.start; r <= pageFooterSection.end; r++) {
      const row = ws.getRow(r);
      let cellsHtml = [];
      row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
        if (cIdx === 1) return;
        const text = (cell.text || '').trim();
        if (!text) return;
        const processed = processNgText(text);
        cellsHtml.push(`    <span>${processed}</span>`);
      });
      if (cellsHtml.length > 0) {
        html.push(`  <div class="footer-row">${cellsHtml.join(' ')}</div>`);
      }
    }
    html.push(`</div>`);
  }

  // ===== BUILD TYPESCRIPT COMPONENT CLASS =====
  let ts = [];

  ts.push(`import { Component, Input } from '@angular/core';`);
  ts.push(`import { CommonModule } from '@angular/common';`);
  ts.push(`import { TableModule } from 'primeng/table';`);
  ts.push(``);

  ts.push(`@Component({`);
  ts.push(`  selector: 'app-report-table',`);
  ts.push(`  templateUrl: './report-table.component.html',`);
  ts.push(`  standalone: true,`);
  ts.push(`  imports: [TableModule, CommonModule],`);
  ts.push(`})`);
  ts.push(`export class ReportTableComponent {`);
  ts.push(`  @Input() data: any[] = [];`);
  ts.push(`}`);

  // ===== COMBINE OUTPUT =====
  let output = [];
  output.push(`<!-- ============================================ -->`);
  output.push(`<!-- EXCEL REPORT CODE GENERATOR — PrimeNG Table -->`);
  output.push(`<!-- ============================================ -->`);
  output.push(`<!-- Save as: report-table.component.html       -->`);
  output.push(`<!-- ============================================ -->`);
  output.push(``);
  output.push(html.join('\n'));
  output.push(``);
  output.push(`<!-- ============================================ -->`);
  output.push(`/* Save as: report-table.component.ts              */`);
  output.push(`/* ============================================ */`);
  output.push(``);
  output.push(ts.join('\n'));

  return output.join('\n');
}

// ============================================
// Helper Functions
// ============================================

function toCamelCase(str) {
  if (!str) return '';
  return str
    .replace(/[^a-zA-Z0-9\u0E00-\u0E7F ]/g, ' ')
    .split(/\s+/)
    .map((word, i) => {
      if (i === 0) return word.toLowerCase();
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join('')
    .replace(/[^a-zA-Z0-9_]/g, '');
}

function escapeNgString(str) {
  if (!str) return '';
  return String(str)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '');
}

function processNgText(text) {
  if (!text) return '';
  // Replace {{field}} with Angular interpolation using row context
  // Replace P{{param}} with params?.paramName
  let result = String(text);

  // P{{paramName}} → {{ params?.paramName }}
  result = result.replace(/[pP]\{\{(.*?)\}\}/g, (_, name) => {
    return `{{ params?.${toCamelCase(name.trim())} }}`;
  });

  // {{fieldName}} → {{ row.fieldName }}
  result = result.replace(/\{\{(.*?)\}\}/g, (_, name) => {
    return `{{ row?.${toCamelCase(name.trim())} }}`;
  });

  // Escape @ to &#64; to avoid Angular control flow syntax conflicts (Angular 17+)
  result = result.replace(/@/g, '&#64;');

  return result;
}

function buildNgStyle(cell, themeColors) {
  const style = {};

  if (cell.font) {
    if (cell.font.bold) style.fontWeight = 'bold';
    if (cell.font.italic) style.fontStyle = 'italic';
    if (cell.font.size) style.fontSize = `${cell.font.size}pt`;
    if (cell.font.name) style.fontFamily = cell.font.name;
    if (cell.font.color) {
      const argb = toARGB(cell.font.color, themeColors);
      if (argb) style.color = argbToHex(argb);
    }
  }

  if (cell.fill?.fgColor) {
    const argb = toARGB(cell.fill.fgColor, themeColors);
    if (argb) style.backgroundColor = argbToHex(argb);
  }

  if (cell.alignment?.horizontal) {
    style.textAlign = cell.alignment.horizontal;
  }

  return style;
}



function getNgType(col, detailFields) {
  const field = detailFields.find(f => toCamelCase(
    f.fieldName || ''
  ) === col.field);
  
  if (field) {
    switch (field.type) {
      case 'field': return { type: 'any', optional: true };
      case 'parameter': return { type: 'any', optional: true };
      default: return { type: 'string', optional: false };
    }
  }
  return { type: 'string', optional: false };
}

function mapToNgType(javaType) {
  switch (javaType) {
    case 'java.lang.Integer': return { type: 'number', optional: false };
    case 'java.math.BigDecimal': return { type: 'number', optional: false };
    case 'java.util.Date': return { type: 'Date', optional: false };
    case 'java.lang.Boolean': return { type: 'boolean', optional: false };
    default: return { type: 'string', optional: false };
  }
}

/**
 * Infer Angular pipe type from Excel cell number format (numFmt)
 * Examples: "0.00" -> number, "dd/mm/yyyy" -> date, "General" -> null (fallback)
 */
function inferPipeFromNumFmt(numFmt) {
  if (!numFmt) return null;
  const fmt = String(numFmt).trim();
  if (!fmt || fmt.toLowerCase() === 'general' || fmt === '@') return null;

  const lower = fmt.toLowerCase();

  // Detect date patterns (y=year, d=day)
  const isDate = /[yd]/i.test(fmt);

  // Detect time patterns:
  // - h or hh is always hours
  // - s or ss is always seconds
  // - m or mm is minutes ONLY when near h (before) or s (after)
  //   (avoid matching "mm" as month in "dd/mm/yyyy")
  const hasHour = lower.includes('h');
  const hasSecond = lower.includes('s');
  const hasMinuteAfterHour = /h.*m/.test(lower) || /m.*s/.test(lower);
  const hasTime = hasHour || hasSecond || hasMinuteAfterHour;

  if (isDate && hasTime) return 'datetime';
  if (isDate) return 'date';
  if (hasTime) return 'datetime';

  // Number/Currency/Percentage formats
  if (/[0#.$%,$]/.test(fmt)) return 'number';

  return null;
}

/**
 * Infer Angular pipe expression based on field name conventions
 */
function inferNgPipe(fieldName) {
  if (!fieldName) return 'text';
  const name = fieldName.toLowerCase();
  // Datetime fields (date + time)
  if (name.includes('time') || name.endsWith('date')) {
    return 'datetime';
  }
  // Date-only fields
  if (name.includes('date') || name === 'created' || name === 'updated') {
    return 'date';
  }
  // Number fields (exclude short/noisy patterns like 'no', 'sort')
  if (name.includes('tax') || name.includes('total') || name.includes('amount') ||
      name.includes('sum') || name.includes('price') || name.includes('cost') ||
      name.includes('fee') || name.includes('rate') || name.includes('qty') ||
      name.includes('quantity') || name.includes('balance') || name.includes('net') ||
      name.includes('gross') || name.includes('discount') || name.includes('vat') ||
      name.includes('percent') || name.includes('count') || name.includes('number')) {
    return 'number';
  }
  return 'text';
}

function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

/**
 * Convert a style object (from buildNgStyle) to an inline CSS string
 */
function styleObjToCss(style) {
  if (!style || Object.keys(style).length === 0) return '';
  const parts = [];
  for (const [key, value] of Object.entries(style)) {
    // Skip font-size and font-family — let PrimeNG theme handle those
    if (key === 'fontSize' || key === 'fontFamily') continue;
    const cssKey = key.replace(/([A-Z])/g, '-$1').toLowerCase();
    parts.push(`${cssKey}: ${value}`);
  }
  return parts.join('; ');
}

function buildInlineStyle(cell, themeColors) {
  return styleObjToCss(buildNgStyle(cell, themeColors));
}

/**
 * Collect all P{{param}} references from the entire worksheet
 */
function collectAllParams(ws) {
  const params = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = (cell.text || '').trim();
      const match = text.match(/[pP]\{\{(.*?)\}\}/);
      if (match) {
        const name = match[1].trim();
        if (!params.find(p => p.name === name)) {
          params.push({
            name: name,
            type: inferFieldType(cell)
          });
        }
      }
    });
  });
  return params;
}
