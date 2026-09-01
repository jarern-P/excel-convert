// ============================================
// iTextSharp PDF Export - Code Generator
// ============================================
// Generates C# code using iTextSharp for full-featured PDF export
// with styles, fonts, colors, borders, images, and dynamic data.
// ============================================

function exportToItextSharp({ ws: wsJson, workbook: wbJson, meta, images = [] }) {
    const ws = buildWorksheetFromJson(wsJson);
    const workbook = buildWorkbookFromJson(wbJson);
    let code = [];
    const themeColors = getThemeColors(workbook);

    // Font & color variable counters (unique per method scope — no global caching)
    let fontIndex = 0;
    let colorIndex = 0;

    function buildFontCached(cell, lines) {
        if (!cell.font) return null;

        const fontVar = `font_${fontIndex++}`;
        const fontName = cell.font.name || 'Arial';
        const fontSize = parseFloat(cell.font.size) || 11;

        lines.push(`var ${fontVar} = FontFactory.GetFont("${fontName}", ${fontSize}f);`);

        if (cell.font.bold && cell.font.italic) {
            lines.push(`${fontVar}.SetStyle(Font.BOLDITALIC);`);
        } else if (cell.font.bold) {
            lines.push(`${fontVar}.SetStyle(Font.BOLD);`);
        } else if (cell.font.italic) {
            lines.push(`${fontVar}.SetStyle(Font.ITALIC);`);
        }

        if (cell.font.underline) {
            lines.push(`${fontVar}.SetStyle(${fontVar}.GetStyle() | Font.UNDERLINE);`);
        }
        if (cell.font.strike) {
            lines.push(`${fontVar}.SetStyle(${fontVar}.GetStyle() | Font.STRIKETHRU);`);
        }

        if (cell.font?.color) {
            const hex = toARGB(cell.font.color, themeColors).substring(2);
            const r = parseInt(hex.substring(0, 2), 16);
            const g = parseInt(hex.substring(2, 4), 16);
            const b = parseInt(hex.substring(4, 6), 16);
            const colorVar = buildColorCached(r, g, b, lines);
            lines.push(`${fontVar}.Color = ${colorVar};`);
        }

        return fontVar;
    }

    function buildColorCached(r, g, b, lines) {
        const colorVar = `clr_${colorIndex++}`;
        lines.push(`var ${colorVar} = new BaseColor(${r}, ${g}, ${b});`);
        return colorVar;
    }

    // ===== STYLE CACHE =====
    const styleCache = new Map();
    let styleIndex = 0;

    function mapBorderStyle(style) {
        const map = {
            thin: '0.5f',
            medium: '1.5f',
            thick: '2.5f',
            dotted: '0.5f',
            dashed: '0.75f',
            double: '1.5f',
            hair: '0.25f',
            dashDot: '0.75f',
            dashDotDot: '0.75f',
            mediumDashed: '1.5f',
            mediumDashDot: '1.5f',
            mediumDashDotDot: '1.5f',
            slantDashDot: '1.0f'
        };
        return map[style] || '0.5f';
    }

    function getStyleKey(cell, themeColors) {
        function normColor(c) {
            const hex = c ? toARGB(c, themeColors) : '';
            return hex || 'FF000000';
        }

        function normH(h) { return h || 'left'; }
        function normV(v) { return v || 'middle'; }
        function normBorder(side) { return side?.style || ''; }

        return JSON.stringify({
            h: normH(cell.alignment?.horizontal),
            v: normV(cell.alignment?.vertical),
            wrap: !!cell.alignment?.wrapText,
            b: {
                t: normBorder(cell.border?.top),
                b: normBorder(cell.border?.bottom),
                l: normBorder(cell.border?.left),
                r: normBorder(cell.border?.right)
            },
            borderColor: cell.border?.top?.color ? toARGB(cell.border.top.color, themeColors) : null,
            font: {
                bold: !!cell.font?.bold,
                size: cell.font?.size || 11,
                name: cell.font?.name || 'Arial',
                italic: !!cell.font?.italic,
                underline: !!cell.font?.underline,
                strike: !!cell.font?.strike,
                color: normColor(cell.font?.color)
            },
            fill: cell.fill?.fgColor ? toARGB(cell.fill.fgColor, themeColors) : '',
            numFmt: cell.numFmt || ''
        });
    }

    function buildStyle(cell) {
        const styleVar = `cellStyle_${styleIndex++}`;
        let lines = [];

        lines.push(`// Build styled cell: ${styleVar}`);
        lines.push(`PdfPCell ${styleVar}(string text)`);
        lines.push(`{`);

        // Font
        const fontVar = buildFontCached(cell, lines);
        if (fontVar) {
            lines.push(`    var phrase = new Phrase(text ?? "", ${fontVar});`);
        } else {
            lines.push(`    var phrase = new Phrase(text ?? "");`);
        }

        lines.push(`    var cell = new PdfPCell(phrase);`);

        // Alignment
        const alignMap = { center: 'Element.ALIGN_CENTER', left: 'Element.ALIGN_LEFT', right: 'Element.ALIGN_RIGHT' };
        if (cell.alignment?.horizontal) {
            lines.push(`    cell.HorizontalAlignment = ${alignMap[cell.alignment.horizontal] || 'Element.ALIGN_LEFT'};`);
        }

        const vAlignMap = { middle: 'Element.ALIGN_MIDDLE', top: 'Element.ALIGN_TOP', bottom: 'Element.ALIGN_BOTTOM' };
        if (cell.alignment?.vertical) {
            lines.push(`    cell.VerticalAlignment = ${vAlignMap[cell.alignment.vertical] || 'Element.ALIGN_MIDDLE'};`);
        }

        // Wrap text — default to wrapping to prevent text from overflowing into adjacent cells
        lines.push(`    cell.NoWrap = false;`);

        // Borders
        if (cell.border) {
            const borderSides = [
                { key: 'top', prop: 'BorderWidthTop' },
                { key: 'bottom', prop: 'BorderWidthBottom' },
                { key: 'left', prop: 'BorderWidthLeft' },
                { key: 'right', prop: 'BorderWidthRight' }
            ];

            borderSides.forEach(side => {
                if (cell.border[side.key]) {
                    const width = mapBorderStyle(cell.border[side.key].style);
                    lines.push(`    cell.${side.prop} = ${width};`);
                } else {
                    lines.push(`    cell.${side.prop} = 0f;`);
                }
            });

            // Border color from any side
            const colorSide = cell.border.top || cell.border.bottom || cell.border.left || cell.border.right;
            if (colorSide?.color) {
                const hex = toARGB(colorSide.color, themeColors).substring(2);
                const r = parseInt(hex.substring(0, 2), 16);
                const g = parseInt(hex.substring(2, 4), 16);
                const b = parseInt(hex.substring(4, 6), 16);
                const colorVar = buildColorCached(r, g, b, lines);
                lines.push(`    cell.BorderColor = ${colorVar};`);
            }
        } else {
            lines.push(`    cell.Border = Rectangle.NO_BORDER;`);
        }

        // Background color
        if (cell.fill?.fgColor) {
            const hex = toARGB(cell.fill.fgColor, themeColors).substring(2);
            const r = parseInt(hex.substring(0, 2), 16);
            const g = parseInt(hex.substring(2, 4), 16);
            const b = parseInt(hex.substring(4, 6), 16);
            const bgColorVar = buildColorCached(r, g, b, lines);
            lines.push(`    cell.BackgroundColor = ${bgColorVar};`);
        }

        // Padding
        lines.push(`    cell.Padding = 3f;`);

        lines.push(`    return cell;`);
        lines.push(`}`);
        lines.push(``);

        return { styleVar, lines };
    }

    // ===== DOCUMENT SETUP =====
    code.push(`// ============================================`);
    code.push(`// GENERATED BY Excel Report Code Generator`);
    code.push(`// Library: iTextSharp (https://www.nuget.org/packages/iTextSharp/)`);
    code.push(`// ============================================`);
    code.push(`//`);
    code.push(`// Installation:`);
    code.push(`//   Install-Package iTextSharp`);
    code.push(`//`);
    code.push(`// Required usings:`);
    code.push(`//   using System;`);
    code.push(`//   using System.IO;`);
    code.push(`//   using System.Collections.Generic;`);
    code.push(`//   using iTextSharp.text;`);
    code.push(`//   using iTextSharp.text.pdf;`);
    code.push(`// ============================================`);
    code.push(`//`);
    code.push(`// iTextSharp Parameters — set these before running the export`);
    code.push(`var Params = new Dictionary<string, dynamic>();`);
    code.push(``);
    // Calculate total column width in points for automatic orientation detection
    let totalWidthPt = 0;
    for (let i = 1; i <= ws.columnCount; i++) {
        const col = ws.getColumn(i);
        totalWidthPt += getColWidthPt(col.width);
    }
    const portraitUsableWidth = 520;
    const isLandscape = totalWidthPt > portraitUsableWidth;

    code.push(`    // Automatic page orientation (total width: ${totalWidthPt}pt, ${isLandscape ? 'Landscape' : 'Portrait'})`);
    code.push(`    float totalWidthPt = ${totalWidthPt}f;`);
    code.push(`    bool isLandscape = totalWidthPt > ${portraitUsableWidth}f;`);
    code.push(`    var pageSize = isLandscape ? PageSize.A4.Rotate() : PageSize.A4;`);
    code.push(`    var document = new Document(pageSize, 20, 20, 20, 20);`);
    code.push(`var ms = new MemoryStream();`);
    code.push(`var writer = PdfWriter.GetInstance(document, ms);`);
    code.push(`document.Open();`);
    code.push(``);

    // ===== IMAGES =====
    if (images && images.length > 0) {
        const media = workbook.model?.media || [];
        media.forEach((m, i) => {
            if (!m?.buffer) return;
            const base64 = toBase64(m.buffer);
            code.push(`    var imageBytes${i} = Convert.FromBase64String("${base64}");`);
        });

        images.forEach((img, idx) => {
            const tl = img.range?.tl;
            const br = img.range?.br;
            if (!tl || !br) return;
            const imgIndex = img.imageId ?? idx;

            code.push(`
    var img_${idx} = Image.GetInstance(imageBytes${imgIndex});
    img_${idx}.ScaleToFit(${(br.nativeColOff || 0) / 9525 * 0.75}f, ${(br.nativeRowOff || 0) / 9525 * 0.75}f);
    document.Add(img_${idx});
            `);
        });
    }

    // ===== SECTIONS =====
    const sections = buildSections(meta, ws.rowCount || ws.lastRow?.number || 0);
    const numCols = ws.columnCount;

    // Calculate column widths for the PDF table
    const colWidths = [];
    for (let c = 1; c <= numCols; c++) {
        const col = ws.getColumn(c);
        colWidths.push(getColWidthPt(col.width));
    }

    code.push(`    // ===== PDF TABLE =====`);
    code.push(`    var table = new PdfPTable(${numCols});`);
    code.push(`    table.WidthPercentage = 100;`);
    code.push(`    table.SetWidths(new float[] { ${colWidths.join('f, ')}f });`);
    code.push(``);

    const detailSection = sections.find(s => s.key === 'detail');

    sections.forEach(section => {
        const sectionKey = section.key.toLowerCase();
        // Only the FIRST detail section is the repeating template; additional
        // 'detail' markers (e.g. sample multi-group sheets) must not re-emit the loop.
        const isDetail = sectionKey === 'detail' && section === detailSection;
        const isStaticSection = !isDetail;

        code.push(``);
        code.push(`    // =======================`);
        code.push(`    // SECTION: ${section.key.toUpperCase()}`);
        code.push(`    // ROW ${section.start} - ${section.end}`);
        code.push(`    // =======================`);

        if (isDetail) {
            // ===== DETAIL SECTION: Dynamic data loop =====
            // MULTI-LINE SUPPORT: every row of the detail section becomes a template row.
            // Each data record renders ALL template rows (section.start .. section.end) in order.
            const detailRows = [];
            const rpParamNames = new Set();

            for (let dr = section.start; dr <= section.end; dr++) {
                const templateRow = ws.getRow(dr);
                const config = parseDetailTemplate(templateRow);

                const fieldMap = config.map(col => col.field ? `"${col.field}"` : 'null');
                const paramMap = config.map(col => col.param ? `"${col.param}"` : 'null');
                // Static text (no {{field}}/P{{param}}) repeats literally on every record.
                // Escape backslashes FIRST, then \r\n\t, then quotes so multi-line cells
                // don't emit raw newlines inside the generated C# string literal.
                const staticMap = config.map(col => col.staticText
                    ? `"${String(col.staticText)
                        .replace(/\\/g, '\\\\')
                        .replace(/\r/g, '\\r')
                        .replace(/\n/g, '\\n')
                        .replace(/\t/g, '\\t')
                        .replace(/"/g, '\\"')}"`
                    : 'null');

                const styleMap = [];
                const printWhenMap = [];
                const textExprMap = [];

                templateRow.eachCell({ includeEmpty: true }, (cell, cIdx) => {
                    const key = getStyleKey(cell, themeColors);

                    if (!styleCache.has(key)) {
                        const { styleVar, lines } = buildStyle(cell);
                        styleCache.set(key, styleVar);
                        code.push(...lines.map(l => `    ${l}`));
                    }

                    styleMap.push(styleCache.get(key));

                    // RPPRINTIF / RP from cell note
                    const noteText = getCellNoteText(cell);
                    const rpFormula = parseRPPRINTIF(noteText);
                    if (rpFormula) {
                        const paramRegex = /[pP]\{\{(.*?)\}\}/g;
                        let m;
                        while ((m = paramRegex.exec(rpFormula.expression)) !== null) {
                            rpParamNames.add(m[1].trim());
                        }

                        if (rpFormula.type === 'printWhen') {
                            printWhenMap.push(convertToCSharpExpression(rpFormula.expression));
                            textExprMap.push('null');
                        } else if (rpFormula.type === 'textExpression') {
                            printWhenMap.push('null');
                            textExprMap.push(convertToCSharpExpression(rpFormula.expression));
                        }
                    } else {
                        printWhenMap.push('null');
                        textExprMap.push('null');
                    }
                });

                // Merge span
                const mergeSpan = [];
                let skip = 0;
                templateRow.eachCell({ includeEmpty: true }, (cell, cIdx) => {
                    if (skip > 0) {
                        skip--;
                        mergeSpan.push(0);
                        return;
                    }

                    if (cell.isMerged && cell.master && cell.address === cell.master.address) {
                        const merges = ws.model.merges || [];
                        let span = 1;
                        merges.forEach(m => {
                            const [start, end] = m.split(':');
                            if (start === cell.address) {
                                const s = decodeAddr(start);
                                const e = decodeAddr(end);
                                span = e.c - s.c + 1;
                            }
                        });
                        mergeSpan.push(span);
                        skip = span - 1;
                    } else {
                        mergeSpan.push(1);
                    }
                });

                detailRows.push({
                    fieldMap,
                    paramMap,
                    staticMap,
                    styleMap,
                    mergeSpan,
                    printWhenMap,
                    textExprMap
                });
            }

            let rpParamDeclText = '';
            if (rpParamNames.size > 0) {
                const paramList = Array.from(rpParamNames).map(n => `Params["${n}"]`).join(', ');
                rpParamDeclText = `\n    // RPPRINTIF uses Params dictionary — set values before running: ${paramList}\n`;
            }

            // ===== 2D arrays: index [templateRow][cell] =====
            const fmtStrArr = (arr) => `new string[] { ${arr.join(', ')} }`;
            const fmtIntArr = (arr) => `new int[] { ${arr.join(', ')} }`;
            const fmtPrintArr = (arr) => `new Func<Dictionary<string, object>, bool>[] { ${arr.map(c => c !== 'null' ? `dtRow => ${c}` : 'null').join(', ')} }`;
            const fmtTextArr = (arr) => `new Func<Dictionary<string, object>, object>[] { ${arr.map(c => c !== 'null' ? `dtRow => (object)(${c})` : 'null').join(', ')} }`;
            const fmtStyleArr = (arr) => `new Func<string, PdfPCell>[] { ${arr.join(', ')} }`;

            code.push(`
    // ===== DETAIL LOOP (${detailRows.length} template row(s) per record) =====
    string[][] fieldMap =
    {
        ${detailRows.map(r => fmtStrArr(r.fieldMap)).join(',\n        ')}
    };

    string[][] paramMap =
    {
        ${detailRows.map(r => fmtStrArr(r.paramMap)).join(',\n        ')}
    };

    string[][] staticMap =
    {
        ${detailRows.map(r => fmtStrArr(r.staticMap)).join(',\n        ')}
    };

    int[][] mergeSpan =
    {
        ${detailRows.map(r => fmtIntArr(r.mergeSpan)).join(',\n        ')}
    };

    Func<string, PdfPCell>[][] styleFuncs =
    {
        ${detailRows.map(r => fmtStyleArr(r.styleMap)).join(',\n        ')}
    };

    Func<Dictionary<string, object>, bool>[][] printWhenMap =
    {
        ${detailRows.map(r => fmtPrintArr(r.printWhenMap)).join(',\n        ')}
    };

    Func<Dictionary<string, object>, object>[][] textExprMap =
    {
        ${detailRows.map(r => fmtTextArr(r.textExprMap)).join(',\n        ')}
    };
    ${rpParamDeclText}`);

            // Data conversion
            code.push(`
    // ===== CONVERT DATA SOURCE =====
    IEnumerable<Dictionary<string, object>> ConvertToDictList(object data)
    {
        if (data is DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                var dict = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                    dict[col.ColumnName] = row[col];
                yield return dict;
            }
        }
        else if (data is System.Collections.IEnumerable enumerable)
        {
            foreach (var item in enumerable)
            {
                if (item is Dictionary<string, object> d)
                    yield return d;
                else
                {
                    var dict = new Dictionary<string, object>();
                    foreach (var prop in item.GetType().GetProperties())
                        dict[prop.Name] = prop.GetValue(item);
                    yield return dict;
                }
            }
        }
    }

    var _data = ConvertToDictList(P_Data);

    foreach (var dtRow in _data)
    {
        for (int t = 0; t < fieldMap.Length; t++)
        {
            for (int i = 0; i < fieldMap[t].Length; i++)
            {
                object value = null;
                bool hasTextExpr = i < textExprMap[t].Length && textExprMap[t][i] != null;
                bool hasPrintWhen = i < printWhenMap[t].Length && printWhenMap[t][i] != null;
                bool shouldPrint = !hasPrintWhen || printWhenMap[t][i](dtRow);

                if (!shouldPrint) continue;

                if (hasTextExpr)
                {
                    value = textExprMap[t][i](dtRow);
                }
                else
                {
                    var staticVal = staticMap[t][i];
                    if (!string.IsNullOrEmpty(staticVal))
                        value = staticVal;
                    else
                    {
                        var param = paramMap[t][i];
                        if (!string.IsNullOrEmpty(param))
                            value = Params[param];
                        else
                        {
                            var field = fieldMap[t][i];
                            if (!string.IsNullOrEmpty(field) && dtRow.ContainsKey(field))
                                value = dtRow[field];
                        }
                    }
                }

                string textValue = value?.ToString() ?? "";

                PdfPCell cell;
                if (styleFuncs[t][i] != null)
                    cell = styleFuncs[t][i](textValue);
                else
                {
                    cell = new PdfPCell(new Phrase(textValue));
                    cell.Padding = 3f;
                    cell.Border = Rectangle.NO_BORDER;
                }

                if (mergeSpan[t][i] > 1)
                    cell.Colspan = mergeSpan[t][i];
                table.AddCell(cell);
            }
        }
    }`);
        } else {
            // ===== STATIC SECTION =====

            // Build global merge map from worksheet for rowspan/colspan handling
            const mergeMap = [];
            if (ws.model.merges) {
                ws.model.merges.forEach(m => {
                    const [start, end] = m.split(':');
                    const s = decodeAddr(start);
                    const e = decodeAddr(end);
                    mergeMap.push({
                        startRow: s.r,
                        startCol: s.c,
                        colSpan: e.c - s.c + 1,
                        rowSpan: e.r - s.r + 1
                    });
                });
            }

            for (let rIdx = section.start; rIdx <= section.end; rIdx++) {
                const row = ws.getRow(rIdx);

                // Check which columns are occupied by rowspan from previous rows
                const occupiedByRowspan = {};
                mergeMap.forEach(mm => {
                    if (mm.rowSpan > 1 && rIdx > mm.startRow && rIdx < mm.startRow + mm.rowSpan) {
                        for (let c = mm.startCol; c < mm.startCol + mm.colSpan; c++) {
                            occupiedByRowspan[c] = true;
                        }
                    }
                });

                // Build merge span map for this row (column 1 to numCols)
                const mergeSpan = [];
                let skip = 0;
                for (let c = 1; c <= numCols; c++) {
                    if (occupiedByRowspan[c]) {
                        mergeSpan.push(-1); // occupied by rowspan from above
                        continue;
                    }
                    if (skip > 0) {
                        skip--;
                        mergeSpan.push(0);
                        continue;
                    }

                    const cell = row.getCell(c);

                    if (cell.isMerged && cell.master) {
                        if (cell.address === cell.master.address) {
                            // This cell IS the master of a merge
                            const mm = mergeMap.find(m => m.startRow === rIdx && m.startCol === c);
                            const colSpan = mm ? mm.colSpan : 1;
                            mergeSpan.push(colSpan);
                            skip = colSpan - 1;
                        } else {
                            // Non-master merged cell (part of horiz or vert merge) → skip
                            mergeSpan.push(0);
                        }
                    } else {
                        mergeSpan.push(1);
                    }
                }

                // Collect cells (skip non-master merged cells & rowspan-occupied)
                const rowCells = [];
                for (let c = 1; c <= numCols; c++) {
                    const span = mergeSpan[c - 1];
                    if (span === 0 || span === -1) continue;

                    const cell = row.getCell(c);

                    const key = getStyleKey(cell, themeColors);
                    if (!styleCache.has(key)) {
                        const { styleVar, lines } = buildStyle(cell);
                        styleCache.set(key, styleVar);
                        code.push(...lines.map(l => `    ${l}`));
                    }
                    const styleVar = styleCache.get(key);

                    let valueCode, hasValue;
                    if (cell.value != null) {
                        hasValue = true;
                        if (typeof cell.value === 'number') {
                            valueCode = `"${cell.value}"`;
                        } else {
                            const text = (cell.text || '').replace(/\r\n/g, '\\n').replace(/"/g, '\\\\"');
                            const paramMatch = text.match(/^[pP]\{\{(.+?)\}\}$/);
                            if (paramMatch) {
                                valueCode = `Params["${paramMatch[1]}"]?.ToString() ?? ""`;
                            } else {
                                valueCode = `"${text}"`;
                            }
                        }
                    } else {
                        hasValue = false;
                        valueCode = '""';
                    }

                    // Look up colSpan and rowSpan for this cell
                    const mm = mergeMap.find(m => m.startRow === rIdx && m.startCol === c);
                    const colSpan = mm ? mm.colSpan : 1;
                    const rowSpan = mm ? mm.rowSpan : 1;

                    rowCells.push({ cIdx: c, styleVar, hasValue, valueCode, colSpan, rowSpan });
                }

                // Generate cells
                let ci = 0;
                while (ci < rowCells.length) {
                    const cur = rowCells[ci];

                    // Handle merged cell with colspan/rowspan
                    if (cur.colSpan > 1 || cur.rowSpan > 1) {
                        const styleVar = cur.styleVar;
                        code.push(`    // Row ${rIdx} Col ${cur.cIdx} (colspan: ${cur.colSpan}, rowspan: ${cur.rowSpan})`);
                        code.push(`    {`);
                        code.push(`        var c = ${styleVar}(${cur.valueCode});`);
                        if (cur.colSpan > 1) {
                            code.push(`        c.Colspan = ${cur.colSpan};`);
                        }
                        if (cur.rowSpan > 1) {
                            code.push(`        c.Rowspan = ${cur.rowSpan};`);
                        }
                        code.push(`        table.AddCell(c);`);
                        code.push(`    }`);
                        ci++;
                        continue;
                    }

                    // Group consecutive duplicate cells (colSpan=1, rowSpan=1 only)
                    let end = ci;
                    while (
                        end + 1 < rowCells.length &&
                        rowCells[end + 1].colSpan === 1 &&
                        rowCells[end + 1].rowSpan === 1 &&
                        rowCells[end + 1].hasValue === cur.hasValue &&
                        rowCells[end + 1].valueCode === cur.valueCode &&
                        rowCells[end + 1].styleVar === cur.styleVar
                    ) {
                        end++;
                    }
                    const count = end - ci + 1;

                    const styleVar = cur.styleVar;
                    const firstCol = cur.cIdx;

                    if (styleVar) {
                        if (count > 1) {
                            code.push(`    // Row ${rIdx} Col ${firstCol}-${firstCol + count - 1}`);
                            code.push(`    for (int i = 0; i < ${count}; i++)`);
                            code.push(`    {`);
                            code.push(`        var c = ${styleVar}(${cur.valueCode});`);
                            code.push(`        table.AddCell(c);`);
                            code.push(`    }`);
                        } else {
                            code.push(`    // Row ${rIdx} Col ${firstCol}`);
                            code.push(`    {`);
                            code.push(`        var c = ${styleVar}(${cur.valueCode});`);
                            code.push(`        table.AddCell(c);`);
                            code.push(`    }`);
                        }
                    } else {
                        if (count > 1) {
                            code.push(`    // Row ${rIdx} Col ${firstCol}-${firstCol + count - 1}`);
                            code.push(`    for (int i = 0; i < ${count}; i++)`);
                            code.push(`        table.AddCell(new PdfPCell(new Phrase(${cur.valueCode})) { Padding = 3f });`);
                        } else {
                            code.push(`    // Row ${rIdx} Col ${firstCol}`);
                            code.push(`    table.AddCell(new PdfPCell(new Phrase(${cur.valueCode})) { Padding = 3f });`);
                        }
                    }

                    ci = end + 1;
                }
            }
        }
    });

    // ===== MERGES (static sections only) =====
    if (ws.model.merges) {
        code.push(`
    // Note: Merged cells are handled via Colspan and Rowspan per-cell above.
`);
    }

    // ===== FINALIZE =====
    code.push(`
    document.Add(table);
    document.Close();

    return ms.ToArray();`);

    return code.join('\n');
}
