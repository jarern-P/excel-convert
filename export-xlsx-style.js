// ============================================
// xlsx-js-style Export - Excel Report Code Generator
// ============================================
// Generates JavaScript code using:
//   import * as XLSX from 'xlsx-js-style';
// ============================================

function exportToXlsxStyle({ ws: wsJson, workbook: wbJson, meta, images = [] }) {
  const ws = buildWorksheetFromJson(wsJson);
  const workbook = buildWorkbookFromJson(wbJson);
  const themeColors = getThemeColors(workbook);
  const sections = buildSections(meta, ws.rowCount || ws.lastRow?.number || 0);

  // ===== BUILD STYLE FUNCTIONS (for code generation) =====

  function buildFontStr(cell) {
    if (!cell.font) return 'null';
    const parts = [];
    if (cell.font.name) parts.push(`name: "${cell.font.name}"`);
    if (cell.font.size) parts.push(`sz: ${cell.font.size}`);
    if (cell.font.bold) parts.push(`bold: true`);
    if (cell.font.italic) parts.push(`italic: true`);
    if (cell.font.underline) parts.push(`underline: true`);
    if (cell.font.strike) parts.push(`strike: true`);
    if (cell.font.color) {
      const argb = toARGB(cell.font.color, themeColors);
      if (argb) parts.push(`color: { rgb: "${argb}" }`);
    }
    if (parts.length === 0) return 'null';
    return `{ ${parts.join(', ')} }`;
  }

  function buildFillStr(cell) {
    if (!cell.fill?.fgColor) return 'null';
    const argb = toARGB(cell.fill.fgColor, themeColors);
    if (!argb) return 'null';
    return `{ patternType: "solid", fgColor: { rgb: "${argb}" } }`;
  }

  function buildAlignmentStr(cell) {
    if (!cell.alignment) return 'null';
    const parts = [];
    if (cell.alignment.horizontal) parts.push(`horizontal: "${cell.alignment.horizontal}"`);
    if (cell.alignment.vertical) parts.push(`vertical: "${cell.alignment.vertical}"`);

    // xlsx-js-style uses "center" for vertical center (ExcelJS uses "middle")
    for (let i = 0; i < parts.length; i++) {
      if (parts[i].startsWith('vertical: "middle"')) {
        parts[i] = 'vertical: "center"';
      }
    }

    if (cell.alignment.wrapText) parts.push(`wrapText: true`);
    if (parts.length === 0) return 'null';
    return `{ ${parts.join(', ')} }`;
  }

  function buildBorderSideStr(side) {
    if (!side?.style) return 'null';
    const parts = [`style: "${side.style}"`];
    if (side.color) {
      const argb = toARGB(side.color, themeColors);
      if (argb) parts.push(`color: { rgb: "${argb}" }`);
    }
    return `{ ${parts.join(', ')} }`;
  }

  function buildBorderStr(cell) {
    if (!cell.border) return 'null';
    const parts = [];
    const top = buildBorderSideStr(cell.border.top);
    if (top !== 'null') parts.push(`top: ${top}`);
    const bottom = buildBorderSideStr(cell.border.bottom);
    if (bottom !== 'null') parts.push(`bottom: ${bottom}`);
    const left = buildBorderSideStr(cell.border.left);
    if (left !== 'null') parts.push(`left: ${left}`);
    const right = buildBorderSideStr(cell.border.right);
    if (right !== 'null') parts.push(`right: ${right}`);
    if (parts.length === 0) return 'null';
    return `{ ${parts.join(', ')} }`;
  }

  function buildStyleObjStr(cell) {
    const fontStr = buildFontStr(cell);
    const fillStr = buildFillStr(cell);
    const alignmentStr = buildAlignmentStr(cell);
    const borderStr = buildBorderStr(cell);
    const numFmt = cell.numFmt || null;

    const parts = [];
    if (fontStr !== 'null') parts.push(`font: ${fontStr}`);
    if (fillStr !== 'null') parts.push(`fill: ${fillStr}`);
    if (alignmentStr !== 'null') parts.push(`alignment: ${alignmentStr}`);
    if (borderStr !== 'null') parts.push(`border: ${borderStr}`);
    if (numFmt) parts.push(`numFmt: "${numFmt.replace(/"/g, '\\"')}"`);

    if (parts.length === 0) return 'null';
    return `{ ${parts.join(', ')} }`;
  }

  function getCellValueStr(cell) {
    if (cell.value == null) return 'null';
    if (typeof cell.value === 'number') return String(cell.value);
    if (typeof cell.value === 'boolean') return String(cell.value);
    if (typeof cell.value === 'object' && cell.value.richText) {
      const text = cell.value.richText.map(x => x.text || '').join('');
      return JSON.stringify(text);
    }
    if (typeof cell.value === 'object' && cell.value.text) {
      return JSON.stringify(cell.value.text);
    }
    if (typeof cell.value === 'object' && (cell.value.formula || cell.value.sharedFormula)) {
      const result = cell.value.result;
      if (typeof result === 'number') return String(result);
      return JSON.stringify(String(result ?? ''));
    }
    return JSON.stringify(String(cell.value));
  }

  // ===== LOCAL HELPER: convert 0-based {r, c} to A1 string =====
  function encodeCellRef(r, c) {
    let col = '';
    let n = c;
    while (n >= 0) {
      col = String.fromCharCode(65 + (n % 26)) + col;
      n = Math.floor(n / 26) - 1;
    }
    return col + (r + 1);
  }

  // ============================================================
  // PASS 1: COLLECT METADATA (fields, params, styles, merge map)
  // ============================================================

  const allFields = [];
  const allParams = new Set();
  const detailSection = sections.find(s => s.key === 'detail');

  // Collect fields and params from worksheet
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      // const text = (cell.text || '').trim();
      // const text = String(cell?.text ?? '').trim();
      const value = cell?.value;
      const text = value == null ? '' : String(value).trim();
      const fieldMatch = text.match(/^\{\{(.+?)\}\}$/);
      const paramMatch = text.match(/^[pP]\{\{(.+?)\}\}$/);
      if (fieldMatch) {
        const name = fieldMatch[1].trim();
        if (!allFields.find(f => f === name)) allFields.push(name);
      }
      if (paramMatch) {
        allParams.add(paramMatch[1].trim());
      }
    });
  });

  // Also collect from cell notes (RPPRINTIF/RP)
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const noteText = getCellNoteText(cell);
      if (noteText) {
        const rpFormula = parseRPPRINTIF(noteText);
        if (rpFormula) {
          const paramRegex = /[pP]\{\{(.*?)\}\}/g;
          let m;
          while ((m = paramRegex.exec(rpFormula.expression)) !== null) {
            allParams.add(m[1].trim());
          }
        }
      }
    });
  });

  // Build style cache (varName -> styleJS string)
  const styleCache = new Map();
  let styleIdx = 0;

  function getStyleVarName(cell) {
    const key = JSON.stringify({
      font: cell.font ? {
        name: cell.font.name,
        sz: cell.font.size,
        bold: cell.font.bold,
        italic: cell.font.italic,
        underline: cell.font.underline,
        strike: cell.font.strike,
        color: cell.font.color ? toARGB(cell.font.color, themeColors) : null
      } : null,
      fill: cell.fill?.fgColor ? toARGB(cell.fill.fgColor, themeColors) : null,
      alignment: cell.alignment ? {
        h: cell.alignment.horizontal,
        v: cell.alignment.vertical,
        wrap: cell.alignment.wrapText
      } : null,
      border: cell.border ? {
        top: cell.border.top?.style,
        topColor: cell.border.top?.color ? toARGB(cell.border.top.color, themeColors) : null,
        bottom: cell.border.bottom?.style,
        bottomColor: cell.border.bottom?.color ? toARGB(cell.border.bottom.color, themeColors) : null,
        left: cell.border.left?.style,
        leftColor: cell.border.left?.color ? toARGB(cell.border.left.color, themeColors) : null,
        right: cell.border.right?.style,
        rightColor: cell.border.right?.color ? toARGB(cell.border.right.color, themeColors) : null,
        diagonal: cell.border.diagonal?.style,
        diagonalColor: cell.border.diagonal?.color ? toARGB(cell.border.diagonal.color, themeColors) : null,
        diagonalUp: cell.border.diagonalUp,
        diagonalDown: cell.border.diagonalDown
      } : null,
      numFmt: cell.numFmt
    });
    if (!styleCache.has(key)) {
      styleCache.set(key, `cellStyle_${styleIdx++}`);
    }
    return styleCache.get(key);
  }

  // Collect all unique styles (use full column loop to catch all cells)
  const allStyleDeclarations = new Map();
  sections.forEach(section => {
    for (let r = section.start; r <= section.end; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = row.getCell(c);

        // Include slave cells of merged ranges so their border styles are captured
        const styleStr = buildStyleObjStr(cell);
        if (styleStr !== 'null') {
          const varName = getStyleVarName(cell);
          allStyleDeclarations.set(varName, styleStr);
        }
      }
    }
  });

  // Build merge map
  const mergeMap = {};
  if (ws.model?.merges) {
    ws.model.merges.forEach(m => {
      const [start, end] = m.split(':');
      const s = decodeAddr(start);
      const e = decodeAddr(end);
      mergeMap[`${s.r},${s.c}`] = { startRow: s.r, startCol: s.c, endRow: e.r, endCol: e.c };
    });
  }

  // ============================================================
  // BORDER PROPAGATION: Apply borders from preceding section to
  // detail cells that lack borders, so the table grid is complete.
  // ============================================================
  // Helper: a border object is "real" when at least one side
  // has a style defined (empty {} objects from ExcelJS are ignored).
  function hasActualBorder(border) {
    return !!(border && (border.top || border.bottom || border.left || border.right || border.diagonal));
  }

  if (detailSection) {
    // Find the nearest section before the detail section that
    // actually has real border definitions on its cells.  Typically
    // this is the COLUMNHEADER band.
    let borderSource = null;
    for (const s of sections) {
      if (s.key.toLowerCase() !== 'detail' && s.end < detailSection.start) {
        let sectionHasBorders = false;
        for (let r = s.start; r <= s.end && !sectionHasBorders; r++) {
          const row = ws.getRow(r);
          for (let c = 1; c <= ws.columnCount && !sectionHasBorders; c++) {
            if (hasActualBorder(row.getCell(c).border)) sectionHasBorders = true;
          }
        }
        if (sectionHasBorders) borderSource = s; // keep the latest one
      }
    }

    if (borderSource) {
      // Collect per-column border from the border source.
      // Iterate all rows so the row closest to the detail wins;
      // only overwrite with borders that have actual sides.
      const columnBorders = {};
      for (let r = borderSource.start; r <= borderSource.end; r++) {
        const srcRow = ws.getRow(r);
        for (let c = 1; c <= ws.columnCount; c++) {
          const srcCell = srcRow.getCell(c);
          if (hasActualBorder(srcCell.border)) {
            columnBorders[c] = JSON.parse(JSON.stringify(srcCell.border));
          }
        }
      }

      // Apply those borders to every detail template cell that
      // does NOT already have real borders of its own.
      if (Object.keys(columnBorders).length > 0) {
        for (let dr = detailSection.start; dr <= detailSection.end; dr++) {
          const tplRow = ws.getRow(dr);
          for (let c = 1; c <= ws.columnCount; c++) {
            const cell = tplRow.getCell(c);
            if (!hasActualBorder(cell.border) && columnBorders[c]) {
              cell.border = JSON.parse(JSON.stringify(columnBorders[c]));

              // Re-register the style so the generated code
              // picks up the new border.
              const styleStr = buildStyleObjStr(cell);
              if (styleStr !== 'null') {
                const varName = getStyleVarName(cell);
                allStyleDeclarations.set(varName, styleStr);
              }
            }
          }
        }
      }
    }
  }

  // ============================================================
  // PASS 2: GENERATE CODE — two buffers
  //   dataCode  : ws_data.push(...) statements (emitted first)
  //   styleCode : ws["A1"].s = ... statements (emitted after sheet creation)
  // ============================================================

  const dataCode = [];
  const styleCode = [];
  let detailRows = [];            // per-template-row configs: { fieldMap, paramMap, styleVars, rpConfig, rowHeight }
  let detailStartRowVar = '0';

  sections.forEach(section => {
    const sectionKey = section.key.toLowerCase();
    const isDetail = sectionKey === 'detail';

    dataCode.push(``);
    dataCode.push(`// ==========================`);
    dataCode.push(`// SECTION: ${section.key.toUpperCase()}`);
    dataCode.push(`// Rows ${section.start} - ${section.end}`);
    dataCode.push(`// ==========================`);

    // Only the FIRST detail section is the repeating template; additional
    // 'detail' markers (e.g. sample multi-group sheets) must not re-emit the loop.
    if (isDetail && section === detailSection) {
      // ===== DETAIL: DATA PUSHING (multi-line: one row per detail template row) =====
      const rowConfigs = [];
      for (let dr = section.start; dr <= section.end; dr++) {
        const templateRow = ws.getRow(dr);
        const fieldMap = [];
        const paramMap = [];
        const staticMap = [];
        const styleVars = [];
        const rpConfig = [];

        for (let c = 1; c <= ws.columnCount; c++) {
          const cell = templateRow.getCell(c);
          const value = cell?.value;
          const text = value == null ? '' : String(value).trim();
          let field = null;
          let param = null;
          let staticText = null;
          const fieldMatch = text.match(/^\{\{(.+?)\}\}$/);
          const paramMatch = text.match(/^[pP]\{\{(.+?)\}\}$/);
          if (paramMatch) {
            param = paramMatch[1].trim();
          } else if (fieldMatch) {
            field = fieldMatch[1].trim();
          } else if (text) {
            // Static text (no {{field}}/P{{param}}) repeats literally on every record
            staticText = text;
          }

          const styleStr = buildStyleObjStr(cell);
          const styleVar = getStyleVarName(cell);
          styleVars.push(styleStr !== 'null' ? `"${styleVar}"` : 'null');

          const noteText = getCellNoteText(cell);
          const rpFormula = parseRPPRINTIF(noteText);
          let hasPrintWhen = false, hasTextExpr = false;
          let printWhenExpr = 'null', textExpr = 'null';
          if (rpFormula) {
            if (rpFormula.type === 'printWhen') {
              hasPrintWhen = true;
              printWhenExpr = convertToJsExpression(rpFormula.expression);
            } else if (rpFormula.type === 'textExpression') {
              hasTextExpr = true;
              textExpr = convertToJsExpression(rpFormula.expression);
            }
          }
          rpConfig.push({ hasPrintWhen, hasTextExpr, printWhen: printWhenExpr, textExpr });

          fieldMap.push(field ? `"${field}"` : 'null');
          paramMap.push(param ? `"${param}"` : 'null');
          staticMap.push(staticText ? JSON.stringify(staticText) : 'null');
        }

        rowConfigs.push({
          fieldMap,
          paramMap,
          staticMap,
          styleVars,
          rpConfig,
          rowHeight: templateRow.height || 15
        });
      }

      detailRows = rowConfigs;

      // ===== Emit 2D arrays for the generated script =====
      const fmt1D = (arr) => `[${arr.join(', ')}]`;
      dataCode.push(`// Template configuration (rows ${section.start}-${section.end} — ${detailRows.length} template row(s) per record)`);
      dataCode.push(`const detailRowCount = ${detailRows.length};`);
      dataCode.push(`const fieldMap = [${detailRows.map(r => fmt1D(r.fieldMap)).join(', ')}];`);
      dataCode.push(`const paramMap = [${detailRows.map(r => fmt1D(r.paramMap)).join(', ')}];`);
      dataCode.push(`const staticMap = [${detailRows.map(r => fmt1D(r.staticMap)).join(', ')}];`);
      dataCode.push(`const detailRowHeights = [${detailRows.map(r => r.rowHeight).join(', ')}];`);
      dataCode.push(``);
      dataCode.push(`const detailStartRow = ws_data.length;`);
      dataCode.push(``);
      dataCode.push(`// Loop through data and populate rows (${detailRows.length} line(s) per record)`);
      dataCode.push(`data.forEach((row, idx) => {`);
      dataCode.push(`  for (let t = 0; t < detailRowCount; t++) {`);
      dataCode.push(`    const rowData = [];`);
      dataCode.push(`    for (let i = 0; i < fieldMap[t].length; i++) {`);
      dataCode.push(`      let value = '';`);
      dataCode.push(`      const staticKey = staticMap[t][i];`);
      dataCode.push(`      if (staticKey) {`);
      dataCode.push(`        value = staticKey;`);
      dataCode.push(`      } else {`);
      dataCode.push(`        const paramKey = paramMap[t][i];`);
      dataCode.push(`        if (paramKey) {`);
      dataCode.push(`          value = params[paramKey];`);
      dataCode.push(`        } else {`);
      dataCode.push(`          const fieldKey = fieldMap[t][i];`);
      dataCode.push(`          if (fieldKey) {`);
      dataCode.push(`            value = row[fieldKey] !== undefined ? row[fieldKey] : '';`);
      dataCode.push(`          }`);
      dataCode.push(`        }`);
      dataCode.push(`      }`);
      dataCode.push(`      rowData.push(value);`);
      dataCode.push(`    }`);
      dataCode.push(`    ws_data.push(rowData);`);
      dataCode.push(`  }`);
      dataCode.push(`});`);

    } else {
      // ===== STATIC SECTION: DATA PUSHING =====
      for (let r = section.start; r <= section.end; r++) {
        const row = ws.getRow(r);
        const rowData = [];

        for (let c = 1; c <= ws.columnCount; c++) {
          const cell = row.getCell(c);
          const isSlave = cell.isMerged && cell.master && cell.address !== cell.master.address;
          if (isSlave) {
            rowData.push('null');
            continue;
          }

          const noteText = getCellNoteText(cell);
          const rpFormula = parseRPPRINTIF(noteText);
          if (rpFormula?.type === 'textExpression') {
            rowData.push(convertToJsExpression(rpFormula.expression));
          } else {
            rowData.push(getCellValueStr(cell));
          }
        }

        dataCode.push(`// Row ${r}`);
        dataCode.push(`ws_data.push([`);
        rowData.forEach((v, i) => {
          dataCode.push(`  ${v}${i < rowData.length - 1 ? ',' : ''}`);
        });
        dataCode.push(`]);`);
      }
    }
  });

  // ============================================================
  // STYLE APPLICATION CODE (emitted AFTER ws creation)
  // ============================================================

  styleCode.push(``);
  styleCode.push(`// ==========================`);
  styleCode.push(`// APPLY CELL STYLES`);
  styleCode.push(`// ==========================`);

  // Static section styles (use full column loop to catch all cells)
  sections.forEach(section => {
    if (section.key.toLowerCase() === 'detail') return;
    for (let r = section.start; r <= section.end; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = row.getCell(c);
        // Include slave cells of merged ranges so their border styles are applied

        const styleStr = buildStyleObjStr(cell);
        if (styleStr === 'null') continue; // skip cells with no style

        const styleVar = getStyleVarName(cell);
        const cellRef = encodeCellRef(r - 1, c - 1);
        const noteText = getCellNoteText(cell);
        const rpFormula = parseRPPRINTIF(noteText);

        // Ensure cell exists before setting style (fixes 'Cannot set properties of undefined' error)
        if (rpFormula?.type === 'printWhen') {
          const condition = convertToJsExpression(rpFormula.expression);
          styleCode.push(`// ${cellRef} - RPPRINTIF`);
          styleCode.push(`if (${condition}) {`);
          styleCode.push(`  if (!ws["${cellRef}"]) ws["${cellRef}"] = { v: '', t: 's' };`);
          styleCode.push(`  ws["${cellRef}"].s = ${styleVar};`);
          styleCode.push(`}`);
        } else {
          styleCode.push(`if (!ws["${cellRef}"]) ws["${cellRef}"] = { v: '', t: 's' };`);
          styleCode.push(`ws["${cellRef}"].s = ${styleVar};`);
        }
      }
    }
  });

  // Detail section styles (applied after ws is created)
  if (detailSection) {
    const hasDetailStyles = detailRows.some(r => r.styleVars.some(v => v !== 'null'));
    if (hasDetailStyles) {
      styleCode.push(``);
      styleCode.push(`// Apply styles to detail cells (${detailRows.length} template row(s) per record)`);
      styleCode.push(`data.forEach((row, idx) => {`);
      styleCode.push(`  for (let t = 0; t < detailRowCount; t++) {`);
      styleCode.push(`    const r = detailStartRow + idx * detailRowCount + t;`);
      styleCode.push(`    for (let i = 0; i < fieldMap[t].length; i++) {`);
      styleCode.push(`      const cellRef = XLSX.utils.encode_cell({ r, c: i });`);
      styleCode.push(`      if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };`);

      detailRows.forEach((rc, t) => {
        const hasRp = rc.rpConfig.some(c => c.hasPrintWhen || c.hasTextExpr);
        if (hasRp) {
          rc.rpConfig.forEach((rp, i) => {
            if (rp.hasPrintWhen) {
              styleCode.push(`      if (t === ${t} && i === ${i} && (${rp.printWhen})) ws[cellRef].s = ${rc.styleVars[i].replace(/"/g, '')};`);
            } else if (rp.hasTextExpr) {
              styleCode.push(`      if (t === ${t} && i === ${i}) { ws[cellRef].v = ${rp.textExpr}; ws[cellRef].s = ${rc.styleVars[i].replace(/"/g, '')}; }`);
            } else if (rc.styleVars[i] !== 'null') {
              styleCode.push(`      if (t === ${t} && i === ${i}) ws[cellRef].s = ${rc.styleVars[i].replace(/"/g, '')};`);
            }
          });
        } else {
          rc.styleVars.forEach((v, i) => {
            if (v !== 'null') {
              styleCode.push(`      if (t === ${t} && i === ${i}) ws[cellRef].s = ${v.replace(/"/g, '')};`);
            }
          });
        }
      });

      styleCode.push(`    }`);
      styleCode.push(`  }`);
      styleCode.push(`  // Row heights for detail rows`);
      styleCode.push(`  if (!ws["!rows"]) ws["!rows"] = [];`);
      styleCode.push(`  for (let t = 0; t < detailRowCount; t++) {`);
      styleCode.push(`    ws["!rows"][detailStartRow + idx * detailRowCount + t] = { hpt: detailRowHeights[t] };`);
      styleCode.push(`  }`);
      styleCode.push(`});`);
    }
  }

  // ============================================================
  // ASSEMBLE FINAL CODE (correct order)
  // ============================================================

  let code = [];

  code.push(`// ============================================`);
  code.push(`// GENERATED BY Excel Report Code Generator`);
  code.push(`// Library: xlsx-js-style (https://www.npmjs.com/package/xlsx-js-style)`);
  code.push(`// ============================================`);
  code.push(`//`);
  code.push(`// Installation:`);
  code.push(`//   npm install xlsx-js-style`);
  code.push(`//`);
  code.push(`// Usage:`);
  code.push(`//   1. Set your data in the 'data' array below`);
  code.push(`//   2. Set any parameters in the 'params' object`);
  code.push(`//   3. Run this script with Node.js`);
  code.push(`// ============================================`);
  code.push(``);
  code.push(`import * as XLSX from 'xlsx-js-style';`);
  code.push(``);
  code.push(`exportExcelReport(){`);
  code.push(`// ==========================`);
  code.push(`// CONFIGURATION`);
  code.push(`// ==========================`);
  code.push(`// Replace with your actual data array`);
  if (allFields.length > 0) {
    code.push(`// Expected fields: ${allFields.map(f => `"${f}"`).join(', ')}`);
  }
  code.push(`let data :any[] = [];`);
  code.push(``);

  if (detailSection) {
    const paramList = allParams.size > 0
      ? Array.from(allParams).map(n => `  "${n}": null`).join(',\n')
      : `  // Add your parameter keys here, e.g. "paramName": null`;
    code.push(`// Set your parameter values here`);
    code.push(`const params: Record<string, any> = {`);
    code.push(`${paramList}`);
    code.push(`};`);
    code.push(``);
  }

  code.push(`// ==========================`);
  code.push(`// CREATE WORKBOOK`);
  code.push(`// ==========================`);
  code.push(`const wb = XLSX.utils.book_new();`);
  code.push(`const ws_data = [];`);
  code.push(``);

  // Style declarations
  if (allStyleDeclarations.size > 0) {
    code.push(`// ==========================`);
    code.push(`// CELL STYLE DEFINITIONS`);
    code.push(`// ==========================`);
    allStyleDeclarations.forEach((styleStr, varName) => {
      code.push(`const ${varName} = ${styleStr};`);
    });
    code.push(``);
  }

  // DATA PUSHING CODE (ws_data.push) — does NOT reference ws
  code.push(`// ==========================`);
  code.push(`// BUILD WORKSHEET DATA`);
  code.push(`// ==========================`);
  code.push(...dataCode);
  code.push(``);

  // CREATE SHEET
  code.push(`// ==========================`);
  code.push(`// CREATE WORKSHEET`);
  code.push(`// ==========================`);
  code.push(`const ws = XLSX.utils.aoa_to_sheet(ws_data);`);
  code.push(``);

  // STYLE APPLICATION CODE (references ws — emitted AFTER ws creation)
  code.push(...styleCode);
  code.push(``);

  // ===== COLUMN WIDTHS =====
  const DEFAULT_COL_WIDTH = ws.properties?.defaultColWidth ?? 8.43;
  let hasCustomWidth = false;
  const colWidths = [];
  for (let i = 1; i <= ws.columnCount; i++) {
    const col = ws.getColumn(i);
    const width = col.width || DEFAULT_COL_WIDTH;
    colWidths.push({ width, isDefault: Math.abs(width - DEFAULT_COL_WIDTH) < 0.1 });
    if (!hasCustomWidth && Math.abs(width - DEFAULT_COL_WIDTH) >= 0.1) hasCustomWidth = true;
  }

  if (hasCustomWidth) {
    code.push(`// ==========================`);
    code.push(`// COLUMN WIDTHS`);
    code.push(`// ==========================`);
    code.push(`ws["!cols"] = [`);
    colWidths.forEach((cw, i) => {
      code.push(`  { wch: ${Math.round(cw.width)} }${i < colWidths.length - 1 ? ',' : ''}`);
    });
    code.push(`];`);
    code.push(``);
  }

  // ===== ROW HEIGHTS =====
  const DEFAULT_ROW_HEIGHT = ws.properties?.defaultRowHeight ?? 15;
  let hasCustomHeight = false;
  const rowHeights = [];

  for (let r = 1; r <= (ws.rowCount || ws.lastRow?.number || 0); r++) {
    const row = ws.getRow(r);
    const h = row.height || DEFAULT_ROW_HEIGHT;
    rowHeights.push({ height: h, isDefault: Math.abs(h - DEFAULT_ROW_HEIGHT) < 0.1 });
    if (!hasCustomHeight && Math.abs(h - DEFAULT_ROW_HEIGHT) >= 0.1) hasCustomHeight = true;
  }

  if (hasCustomHeight) {
    code.push(`// ==========================`);
    code.push(`// ROW HEIGHTS`);
    code.push(`// ==========================`);
    code.push(`ws["!rows"] = [`);
    rowHeights.forEach((rh, i) => {
      if (rh.isDefault) {
        code.push(`  {}${i < rowHeights.length - 1 ? ',' : ''}`);
      } else {
        code.push(`  { hpt: ${rh.height} }${i < rowHeights.length - 1 ? ',' : ''}`);
      }
    });
    code.push(`];`);
    code.push(``);
  }

  // ===== MERGES =====
  if (ws.model?.merges && ws.model.merges.length > 0) {
    code.push(`// ==========================`);
    code.push(`// MERGED CELLS`);
    code.push(`// ==========================`);
    code.push(`ws["!merges"] = [`);
    ws.model.merges.forEach((m, i) => {
      const [start, end] = m.split(':');
      const s = decodeAddr(start);
      const e = decodeAddr(end);
      code.push(`  { s: { r: ${s.r - 1}, c: ${s.c - 1} }, e: { r: ${e.r - 1}, c: ${e.c - 1} } }${i < ws.model.merges.length - 1 ? ',' : ''}`);
    });
    code.push(`];`);
    code.push(``);
  }

  // ===== FINAL ASSEMBLY =====
  code.push(`// ==========================`);
  code.push(`// ADD SHEET TO WORKBOOK & SAVE`);
  code.push(`// ==========================`);
  code.push(`const sheetName = "Configuration_Form";`);
  code.push(`XLSX.utils.book_append_sheet(wb, ws, sheetName);`);
  code.push(``);
code.push(`// Generate buffer (Uint8Array/Buffer for binary Blob compatibility)`);
code.push(`const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });`);
  code.push(``);
  code.push(`// Dynamic filename with current date`);
  code.push(`const now = new Date();`);
  code.push(`const year = now.getFullYear();`);
  code.push(`const month = String(now.getMonth() + 1).padStart(2, '0');`);
  code.push(`const day = String(now.getDate()).padStart(2, '0');`);
  code.push(`const fileName = \`Configuration_Form_\${year}\${month}\${day}.xlsx\`;`);
  code.push(``);
  code.push(`// ==========================`);
  code.push(`// SAVE FILE TO MACHINE`);
  code.push(`// ==========================`);
  code.push(`// Option 1: FileSaver.js (npm install file-saver)`);
  code.push(`//import { saveAs } from 'file-saver';`);
  code.push(`const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); // buffer is Uint8Array, Blob handles it correctly`);
  code.push(`saveAs(blob, fileName);`);
  code.push(``);
  code.push(`// Option 2: Blazor/.NET JS interop`);
  code.push(`// await JSRuntime.InvokeVoidAsync("saveAsFile", fileName, Convert.ToBase64String(buffer));`);
  code.push(``);
  code.push(`// Option 3: Direct download via anchor element`);
  code.push(`// const url = URL.createObjectURL(blob);`);
  code.push(`// const a = document.createElement('a');`);
  code.push(`// a.href = url;`);
  code.push(`// a.download = fileName;`);
  code.push(`// document.body.appendChild(a);`);
  code.push(`// a.click();`);
  code.push(`// document.body.removeChild(a);`);
  code.push(`// setTimeout(() => URL.revokeObjectURL(url), 1000);`);
  code.push(``);
  code.push(`return buffer;`);
  code.push(`}`);

  // Images note
  if (images && images.length > 0) {
    code.push(`// ==========================`);
    code.push(`// Note: xlsx-js-style does not support embedding images.`);
    code.push(`// Use exceljs (ExcelJS.Workbook) if you need image support.`);
    code.push(`// ==========================`);
  }

  return code.join('\n');
}

// ============================================
// HELPER: Convert expressions to JavaScript
// ============================================

function convertToJsExpression(expr) {
  if (!expr) return 'null';

  let result = String(expr);

  // P{{paramName}} → params["paramName"]
  result = result.replace(/[pP]\{\{(.*?)\}\}/g, (_, name) => {
    return `params["${name.trim()}"]`;
  });

  // {{fieldName}} → row["fieldName"]
  result = result.replace(/\{\{(.*?)\}\}/g, (_, name) => {
    return `row["${name.trim()}"]`;
  });

  return result;
}
