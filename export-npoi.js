function exportToNPOI({ ws: wsJson, workbook: wbJson, meta, images = [] }) {
    const ws = buildWorksheetFromJson(wsJson);
    const workbook = buildWorkbookFromJson(wbJson);
    let code = [];
    function mapBorderStyle(style) {

    const map = {
        thin: 'Thin',
        medium: 'Medium',
        thick: 'Thick',

        dotted: 'Dotted',
        dashed: 'Dashed',

        double: 'Double',

        hair: 'Hair',

        dashDot: 'DashDot',
        dashDotDot: 'DashDotDot',

        mediumDashed: 'MediumDashed',
        mediumDashDot: 'MediumDashDot',
        mediumDashDotDot: 'MediumDashDotDot',

        slantDashDot: 'SlantedDashDot'
    };

    return map[style] || 'Thin';
}
    const themeColors = getThemeColors(workbook);
    code.push(`var workbook = new XSSFWorkbook();`);
    code.push(`var sheet = workbook.CreateSheet("Sheet1");`);

    // ===== IMAGE EXPORT =====
if (images && images.length > 0) {

    code.push(`var drawing = sheet.CreateDrawingPatriarch();`);

    const media = workbook.model?.media || [];

    media.forEach((m, i) => {
        if (!m?.buffer) return;

        const base64 = toBase64(m.buffer);
        code.push(`var imageBytes${i} = Convert.FromBase64String("${base64}");`);
    });

    images.forEach((img, idx) => {

        const tl = img.range?.tl;
        const br = img.range?.br;
        if (!tl || !br) return;

        const imgIndex = img.imageId ?? idx;

        code.push(`
var anchor_${idx} = new XSSFClientAnchor(
    ${tl.nativeColOff || 0}, ${tl.nativeRowOff || 0},
    ${br.nativeColOff || 0}, ${br.nativeRowOff || 0},
    ${tl.nativeCol}, ${tl.nativeRow},
    ${br.nativeCol}, ${br.nativeRow}
);

var picIdx_${idx} = workbook.AddPicture(imageBytes${imgIndex}, PictureType.PNG);
drawing.CreatePicture(anchor_${idx}, picIdx_${idx});
        `);
    });
}

    const styleCache = new Map();
    let styleIndex = 0;

    // ===== FONT CACHE =====
    const fontCache = new Map();
    let fontIndex = 0;

    // ===== COLOR CACHE =====
    const colorCache = new Map();
    let colorIndex = 0;

    function getFontKey(cell) {
        return JSON.stringify({
            bold: !!cell.font?.bold,
            size: cell.font?.size || 11,
            name: cell.font?.name || 'Arial',
            italic: !!cell.font?.italic,
            color: cell.font?.color ? toARGB(cell.font.color, themeColors) : null
        });
    }

    function buildFontCached(cell, lines) {
        if (!cell.font) return null;

        const key = getFontKey(cell);
        if (fontCache.has(key)) {
            return fontCache.get(key);
        }

        const fontVar = `font_${fontIndex++}`;
        lines.push(`var ${fontVar} = workbook.CreateFont();`);

        if (cell.font.bold) lines.push(`${fontVar}.IsBold = true;`);
        if (cell.font.size) lines.push(`${fontVar}.FontHeightInPoints = ${cell.font.size};`);
        if (cell.font.name) lines.push(`${fontVar}.FontName = "${cell.font.name}";`);
        if (cell.font.italic) lines.push(`${fontVar}.IsItalic = true;`);

        if (cell.font?.color) {
            const hex = toARGB(cell.font.color, themeColors).substring(2);
            const r = parseInt(hex.substring(0, 2), 16);
            const g = parseInt(hex.substring(2, 4), 16);
            const b = parseInt(hex.substring(4, 6), 16);

            const colorVar = buildColorCached(r, g, b, lines);
            lines.push(`((XSSFFont)${fontVar}).SetColor(${colorVar});`);
        }

        fontCache.set(key, fontVar);
        return fontVar;
    }

    function buildColorCached(r, g, b, lines) {
        const key = `${r},${g},${b}`;
        if (colorCache.has(key)) {
            return colorCache.get(key);
        }

        const colorVar = `clr_${colorIndex++}`;
        lines.push(`var ${colorVar} = new XSSFColor();`);
        lines.push(`${colorVar}.SetRgb(new byte[] { ${r}, ${g}, ${b} });`);

        colorCache.set(key, colorVar);
        return colorVar;
    }

function getStyleKey(cell, themeColors) {

    // ===== NORMALIZE COLOR =====
    function normColor(c) {
        const hex = c ? toARGB(c, themeColors) : '';
        return hex || 'FF000000'; // default = ดำ
    }

    // ===== NORMALIZE ALIGN =====
    function normH(h) {
        return h || 'left';
    }

    function normV(v) {
        return v || 'middle';
    }

    // ===== NORMALIZE BORDER =====
    function normBorder(side) {
        // return side ? 'thin' : ''; // simplify → มี = thin, ไม่มี = ''
        return side?.style || '';
    }

    return JSON.stringify({
        h: normH(cell.alignment?.horizontal),
        v: normV(cell.alignment?.vertical),

        b: {
            t: normBorder(cell.border?.top),
            b: normBorder(cell.border?.bottom),
            l: normBorder(cell.border?.left),
            r: normBorder(cell.border?.right)
        },

        diag: cell.border?.diagonal
        ? {
            style: cell.border.diagonal.style || '',
            up: !!cell.border.diagonal.up,
            down: !!cell.border.diagonal.down
        }
        : null,

        font: {
            bold: !!cell.font?.bold,
            size: cell.font?.size || 11,   // default Excel ~11
            name: cell.font?.name || 'Arial',
            italic: !!cell.font?.italic,
            color: normColor(cell.font?.color)
        },

        fill: cell.fill?.fgColor
            ? toARGB(cell.fill.fgColor, themeColors)
            : '',

        numFmt: cell.numFmt || '' 
    });
}

    function buildStyle(cell) {
        const styleVar = `style_${styleIndex++}`;
        let lines = [];

        lines.push(`var ${styleVar} = workbook.CreateCellStyle();`);

        // ===== ALIGN =====
        const alignMap = {
            center: 'Center',
            left: 'Left',
            right: 'Right'
        };
        if (cell.alignment?.horizontal) {
            lines.push(`${styleVar}.Alignment = HorizontalAlignment.${alignMap[cell.alignment.horizontal] || 'Left'};`);
        }

        const vAlignMap = {
            middle: 'Center',
            top: 'Top',
            bottom: 'Bottom'
        };
        if (cell.alignment?.vertical) {
            lines.push(`${styleVar}.VerticalAlignment = VerticalAlignment.${vAlignMap[cell.alignment.vertical] || 'Center'};`);
        }

        // ===== WRAP TEXT =====
        if (
            cell.alignment?.wrapText ||
            (typeof cell.value === 'string' && cell.value.includes('\\n'))
        ) {
            lines.push(`${styleVar}.WrapText = true;`);
        }

        // ===== BORDER =====
        if (cell.border) {
            const borderProps = [
                { key: 'top', name: 'BorderTop' },
                { key: 'bottom', name: 'BorderBottom' },
                { key: 'left', name: 'BorderLeft' },
                { key: 'right', name: 'BorderRight' }
            ];
            borderProps.forEach(side => {
                if (cell.border[side.key]) {
                    lines.push(
                        `${styleVar}.${side.name} = BorderStyle.${mapBorderStyle(cell.border[side.key].style)};`
                    );
                }
            });
        }

        // ===== BORDER DIAGONAL =====
        if (cell.border?.diagonal) {
            const diagStyle = mapBorderStyle(cell.border.diagonal.style);
            lines.push(`${styleVar}.BorderDiagonalLineStyle = BorderStyle.${diagStyle};`);

            const diagDirMap = {
                'up': 'Forward',
                'down': 'Backward',
                'up,down': 'Both'
            };
            const dirParts = [];
            if (cell.border.diagonal.up) dirParts.push('up');
            if (cell.border.diagonal.down) dirParts.push('down');
            const dirKey = dirParts.join(',');
            if (diagDirMap[dirKey]) {
                lines.push(`${styleVar}.BorderDiagonal = BorderDiagonal.${diagDirMap[dirKey]};`);
            }
        }

        // ===== FONT (CACHED) =====
        const fontVar = buildFontCached(cell, lines);
        if (fontVar) {
            lines.push(`${styleVar}.SetFont(${fontVar});`);
        }

        // ===== BACKGROUND (COLOR CACHED) =====
        if (cell.fill?.fgColor) {
            const hex = toARGB(cell.fill.fgColor, themeColors).substring(2);

            const r = parseInt(hex.substring(0, 2), 16);
            const g = parseInt(hex.substring(2, 4), 16);
            const b = parseInt(hex.substring(4, 6), 16);

            const bgColorVar = buildColorCached(r, g, b, lines);
            lines.push(`((XSSFCellStyle)${styleVar}).SetFillForegroundColor(${bgColorVar});`);
            lines.push(`${styleVar}.FillPattern = FillPattern.SolidForeground;`);
        }

        // ===== NUMBER FORMAT =====
        if (cell.numFmt) {
            const fmt = cell.numFmt.replace(/"/g, '\\"');
            lines.push(`${styleVar}.DataFormat = workbook.CreateDataFormat().GetFormat("${fmt}");`);
        }

        return { styleVar, lines };
    }

    function getExcelColName(colIndex) {
        let name = '';
        while (colIndex > 0) {
            let rem = (colIndex - 1) % 26;
            name = String.fromCharCode(65 + rem) + name;
            colIndex = Math.floor((colIndex - 1) / 26);
        }
        return name;
    }

    // ===== COLUMN WIDTH =====
    // ws.columns.forEach((col, cIdx) => {
    //     if (col.width) {
    //         code.push(`sheet.SetColumnWidth(${cIdx}, ${Math.round(col.width * 256)});`);
    //     }
    // });
// const DEFAULT_WIDTH = 8.43;
const DEFAULT_WIDTH = ws.properties?.defaultColWidth ?? 8.43;

const colLines = ws.columns
  .map((col, i) => {
    if (!col.width) return null;
    const isDefault = Math.abs(col.width - DEFAULT_WIDTH) < 0.1;
    if (isDefault) return null;
    const colName = getExcelColName(i + 1);
    return `    { ${i}, ${Math.round(col.width)} }, // ${colName}`;
  })
  .filter(Boolean);

const cols = colLines
  .map((line, idx) => idx < colLines.length - 1 ? line : line.replace(/,\s*(?=\/\/)/, ''))
  .join("\n");

if (cols) {
    code.push(`var columnWidths = new Dictionary<int, int>
{
${cols}
};`);

    code.push(`
foreach (var col in columnWidths)
    sheet.SetColumnWidth(col.Key, col.Value * 256);
`);
}

    // // ===== CREATE CELLS =====
    code.push(`int rowIdx = 0, colIdx = 0;`);
    code.push(`ICell? cell = null;`);
    code.push(`// NPOI Parameters — set these before running the export`);
    code.push(`var Params = new Dictionary<string, dynamic>();`);
    code.push(``);
    const sections = buildSections(meta, ws.rowCount || ws.lastRow?.number || 0);
    const detailSection = sections.find(s => s.key === 'detail');

    sections.forEach(section => {
    const sectionKey = section.key.toLowerCase();
    // Only the FIRST detail section is the repeating template; additional
    // 'detail' markers (e.g. sample multi-group sheets) must not re-emit the loop.
    const isDetail = sectionKey === 'detail' && section === detailSection;
    const isStaticSection = !isDetail;

    code.push(``);
    code.push(`// =======================`);
    code.push(`// SECTION: ${section.key.toUpperCase()}`);
    code.push(`// ROW ${section.start} - ${section.end}`);
    code.push(`// =======================`);

    if (isDetail) {

        // ===== DETAIL SECTION: Dynamic data loop (DataTable or Dictionary) =====
        // MULTI-LINE SUPPORT: every row of the detail section becomes a template row.
        // Each data record renders ALL template rows (section.start .. section.end) in order.
        const detailRows = [];           // one config per template row
        const rpParamNames = new Set();  // P{{name}} references used in RPPRINTIF/RP

        for (let dr = section.start; dr <= section.end; dr++) {
            const templateRow = ws.getRow(dr);
            const config = parseDetailTemplate(templateRow);

            // ===== FIELD MAP & PARAM MAP =====
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

            // ===== STYLE MAP + RPPRINTIF / RP =====
            const styleMap = [];
            const printWhenMap = [];   // RPPRINTIF condition per cell (C# expression)
            const textExprMap = [];    // RP text expression per cell (C# expression)

            templateRow.eachCell({ includeEmpty: true }, (cell, cIdx) => {
                const key = getStyleKey(cell, themeColors);

                if (!styleCache.has(key)) {
                    const { styleVar, lines } = buildStyle(cell);
                    styleCache.set(key, styleVar);
                    code.push(...lines);
                }

                styleMap.push(styleCache.get(key));

                // ===== RPPRINTIF / RP from cell note =====
                const noteText = getCellNoteText(cell);
                const rpFormula = parseRPPRINTIF(noteText);
                if (rpFormula) {
                    // Extract P{{param}} names from the raw expression for declaration generation
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

            // ===== MERGE SPAN =====
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
                textExprMap,
                rowHeight: ws.getRow(dr).height || 0
            });
        }

        // Build parameter comment if any P{{param}} references exist
        // (actual values are read from Params dictionary at runtime)
        let rpParamDeclText = '';
        if (rpParamNames.size > 0) {
            const paramList = Array.from(rpParamNames).map(n => `Params["${n}"]`).join(', ');
            rpParamDeclText = `\n// RPPRINTIF uses Params dictionary — set values before running: ${paramList}\n`;
        }

        // ===== 2D arrays: index [templateRow][cell] =====
        const fmtStrArr = (arr) => `new string[] { ${arr.join(', ')} }`;
        const fmtStyleArr = (arr) => `new ICellStyle[] { ${arr.join(', ')} }`;
        const fmtIntArr = (arr) => `new int[] { ${arr.join(', ')} }`;
        const fmtPrintArr = (arr) => `new Func<Dictionary<string, object>, bool>[] { ${arr.map(c => c !== 'null' ? `dtRow => ${c}` : 'null').join(', ')} }`;
        const fmtTextArr = (arr) => `new Func<Dictionary<string, object>, object>[] { ${arr.map(c => c !== 'null' ? `dtRow => (object)(${c})` : 'null').join(', ')} }`;

        code.push(`// ===== DETAIL LOOP (FIELD + STYLE + MERGE SPAN + RPPRINTIF) — ${detailRows.length} template row(s) per record =====`);
        code.push(`
string[][] fieldMap =
{
    ${detailRows.map(r => fmtStrArr(r.fieldMap)).join(',\n    ')}
};

string[][] paramMap =
{
    ${detailRows.map(r => fmtStrArr(r.paramMap)).join(',\n    ')}
};

string[][] staticMap =
{
    ${detailRows.map(r => fmtStrArr(r.staticMap)).join(',\n    ')}
};

ICellStyle[][] styleMap =
{
    ${detailRows.map(r => fmtStyleArr(r.styleMap)).join(',\n    ')}
};

int[][] mergeSpan =
{
    ${detailRows.map(r => fmtIntArr(r.mergeSpan)).join(',\n    ')}
};

Func<Dictionary<string, object>, bool>[][] printWhenMap =
{
    ${detailRows.map(r => fmtPrintArr(r.printWhenMap)).join(',\n    ')}
};

Func<Dictionary<string, object>, object>[][] textExprMap =
{
    ${detailRows.map(r => fmtTextArr(r.textExprMap)).join(',\n    ')}
};

float[] detailRowHeights =
{
    ${detailRows.map(r => `${r.rowHeight}f`).join(',\n    ')}
};
${rpParamDeclText}`);

        code.push(`// ===== CONVERT ANY DATA SOURCE TO List<Dictionary<string, object>> =====`);
        code.push(`IEnumerable<Dictionary<string, object>> ConvertToDictList(object data)`);
        code.push(`{`);
        code.push(`    if (data is DataTable dt)`);
        code.push(`    {`);
        code.push(`        foreach (DataRow row in dt.Rows)`);
        code.push(`        {`);
        code.push(`            var dict = new Dictionary<string, object>();`);
        code.push(`            foreach (DataColumn col in dt.Columns)`);
        code.push(`                dict[col.ColumnName] = row[col];`);
        code.push(`            yield return dict;`);
        code.push(`        }`);
        code.push(`    }`);
        code.push(`    else if (data is System.Collections.IEnumerable enumerable)`);
        code.push(`    {`);
        code.push(`        foreach (var item in enumerable)`);
        code.push(`        {`);
        code.push(`            if (item is Dictionary<string, object> d)`);
        code.push(`            {`);
        code.push(`                yield return d;`);
        code.push(`            }`);
        code.push(`            else`);
        code.push(`            {`);
        code.push(`                var dict = new Dictionary<string, object>();`);
        code.push(`                foreach (var prop in item.GetType().GetProperties())`);
        code.push(`                    dict[prop.Name] = prop.GetValue(item);`);
        code.push(`                yield return dict;`);
        code.push(`            }`);
        code.push(`        }`);
        code.push(`    }`);
        code.push(`}`);
        code.push(``);
        code.push(`var _data = ConvertToDictList(P_Data);`);
        code.push(``);
        code.push(`foreach (var dtRow in _data)`);
        code.push(`{`);
        code.push(`    for (int t = 0; t < fieldMap.Length; t++)`);
        code.push(`    {`);
        code.push(`        IRow row = sheet.CreateRow(rowIdx++);`);
        code.push(`        if (detailRowHeights[t] > 0)`);
        code.push(`            row.HeightInPoints = detailRowHeights[t];`);
        code.push(``);
        code.push(`        for (int i = 0; i < fieldMap[t].Length; i++)`);
        code.push(`        {`);
        code.push(`            cell = row.CreateCell(i);`);
        code.push(`            object value = null;`);
        code.push(`            bool hasTextExpr = i < textExprMap[t].Length && textExprMap[t][i] != null;`);
        code.push(`            bool hasPrintWhen = i < printWhenMap[t].Length && printWhenMap[t][i] != null;`);
        code.push(`            bool shouldPrint = !hasPrintWhen || printWhenMap[t][i](dtRow);`);

        code.push(`            if (hasTextExpr)`);
        code.push(`            {`);
        code.push(`                value = textExprMap[t][i](dtRow);`);
        code.push(`            }`);        code.push(`                else`);
        code.push(`                {`);
        code.push(`                    var staticVal = staticMap[t][i];`);
        code.push(`                    if (!string.IsNullOrEmpty(staticVal))`);
        code.push(`                        value = staticVal;`);
        code.push(`                    else`);
        code.push(`                    {`);
        code.push(`                        var param = paramMap[t][i];`);
        code.push(`                        if (!string.IsNullOrEmpty(param))`);
        code.push(`                            value = Params[param];`);
        code.push(`                        else`);
        code.push(`                        {`);
        code.push(`                            var field = fieldMap[t][i];`);
        code.push(`                            if (!string.IsNullOrEmpty(field) && dtRow.ContainsKey(field))`);
        code.push(`                                value = dtRow[field];`);
        code.push(`                        }`);
        code.push(`                    }`);
        code.push(`                }`);

        code.push(`            if (shouldPrint && value != null)`);
        code.push(`            {`);
        code.push(`                if (value is decimal || value is double || value is float || value is int || value is long) cell.SetCellValue(Convert.ToDouble(value));`);
        code.push(`                else if (value is DateTime dt) cell.SetCellValue(dt);`);
        code.push(`                else cell.SetCellValue(value.ToString());`);
        code.push(`            }`);

        code.push(`            if (i < styleMap[t].Length)`);
        code.push(`                cell.CellStyle = styleMap[t][i];`);

        code.push(`
            if (i < mergeSpan[t].Length && mergeSpan[t][i] > 1)
            {
                int span = mergeSpan[t][i];

                for (int j = 1; j < span; j++)
                    row.CreateCell(i + j).CellStyle = styleMap[t][i];
                sheet.AddMergedRegion(new CellRangeAddress(row.RowNum, row.RowNum, i, i + span - 1));

                i += span - 1;
            }
        `);

        code.push(`        }`);
        code.push(`    }`);
        code.push(`}`);

    } else {

        // ===== STATIC SECTION (title, pageheader, columnheader, columnfooter, pagefooter, summary, etc.) =====
        for (let rIdx = section.start; rIdx <= section.end; rIdx++) {

            const row = ws.getRow(rIdx);
            code.push(`var row${rIdx} = sheet.CreateRow(rowIdx++);`);
            code.push(`colIdx = 0;`);

            if (row.height) {
                code.push(`row${rIdx}.HeightInPoints = ${row.height}f;`);
            }

            // Collect cells for grouping consecutive duplicates
            const rowCells = [];
            row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
                const colName = getExcelColName(cIdx);
                const cellRef = `${colName}${rIdx}`;

                // ===== STYLE =====
                const key = getStyleKey(cell, themeColors);
                if (!styleCache.has(key)) {
                    const { styleVar, lines } = buildStyle(cell);
                    styleCache.set(key, styleVar);
                    code.push(...lines);
                }
                const styleVar = styleCache.get(key);

                // ===== RPPRINTIF / RP NOT SUPPORTED in static sections =====
                // (no DataRow context for field refs, no variable declarations for params)
                let printWhenCondition = null;
                let textExpression = null;

                let valueCode, hasValue;
                if (cell.value != null) {
                    hasValue = true;
                    if (typeof cell.value === 'number') {
                        valueCode = `${cell.value}`;
                    } else {
                        const text = (cell.text || '').replace(/\r\n/g, '\n').replace(/"/g, '\\"');
                        // Check for P{{param}} pattern → parameter reference
                        const paramMatch = text.match(/^[pP]\{\{(.+?)\}\}$/);
                        if (paramMatch) {
                            valueCode = `Params["${paramMatch[1]}"]`;
                        } else {
                            valueCode = `"${text}"`;
                        }
                    }
                } else {
                    hasValue = false;
                    valueCode = null;
                }

                rowCells.push({ cIdx, colName, cellRef, styleVar, hasValue, valueCode, printWhenCondition, textExpression });
            });

            // ===== GROUP CONSECUTIVE DUPLICATE CELLS =====
            let ci = 0;
            while (ci < rowCells.length) {
                const cur = rowCells[ci];
                let end = ci;
                while (
                    end + 1 < rowCells.length &&
                    rowCells[end + 1].hasValue === cur.hasValue &&
                    rowCells[end + 1].valueCode === cur.valueCode &&
                    rowCells[end + 1].styleVar === cur.styleVar &&
                    rowCells[end + 1].printWhenCondition === cur.printWhenCondition &&
                    rowCells[end + 1].textExpression === cur.textExpression
                ) {
                    end++;
                }
                const count = end - ci + 1;

                // ===== GENERATE CODE with RPPRINTIF/RP support =====
                function generateCellCode(hasValue, valueCode, styleVar, printWhenCond, textExpr, cellRef) {
                    const effectiveValue = textExpr || valueCode;
                    const hasEffectiveValue = hasValue || !!textExpr;

                    let lines = [];

                    if (printWhenCond) {
                        // RPPRINTIF: conditionally set the value
                        if (count > 1) {
                            // For grouped cells, generate individually to avoid complex nested loops
                            lines.push(`// ${cellRef}`);
                        }
                        lines.push(`{
    var c = row${rIdx}.CreateCell(colIdx++);
    if (${printWhenCond})`);
                        if (hasEffectiveValue) {
                            lines.push(`        c.SetCellValue(${effectiveValue});`);
                        }
                        lines.push(`    c.CellStyle = ${styleVar};
}`);
                    } else if (count > 1) {
                        const firstRef = rowCells[ci].cellRef;
                        const lastRef = rowCells[end].cellRef;
                        if (hasEffectiveValue) {
                            lines.push(`for (int i = 0; i < ${count}; i++)
{
    var c = row${rIdx}.CreateCell(colIdx++);
    c.SetCellValue(${effectiveValue});
    c.CellStyle = ${styleVar};
} // ${firstRef}-${lastRef}`);
                        } else {
                            lines.push(`for (int i = 0; i < ${count}; i++)
    row${rIdx}.CreateCell(colIdx++).CellStyle = ${styleVar};
// ${firstRef}-${lastRef}`);
                        }
                    } else {
                        if (hasEffectiveValue) {
                            lines.push(`cell = row${rIdx}.CreateCell(colIdx++); cell.SetCellValue(${effectiveValue}); cell.CellStyle = ${styleVar}; // ${cellRef}`);
                        } else {
                            lines.push(`row${rIdx}.CreateCell(colIdx++).CellStyle = ${styleVar}; // ${cellRef}`);
                        }
                    }

                    return lines.join('\n');
                }

                if (count > 1 && cur.printWhenCondition) {
                    // Cannot group cells with RPPRINTIF — generate individually
                    for (let j = ci; j <= end; j++) {
                        const c = rowCells[j];
                        code.push(generateCellCode(c.hasValue, c.valueCode, c.styleVar, c.printWhenCondition, c.textExpression, c.cellRef));
                    }
                } else {
                    code.push(generateCellCode(cur.hasValue, cur.valueCode, cur.styleVar, cur.printWhenCondition, cur.textExpression, cur.cellRef));
                }

                ci = end + 1;
            }
        }
    }

});

    // ===== MERGE CELLS (STATIC ONLY) =====
if (ws.model.merges) {
    ws.model.merges.forEach(m => {
        const [start, end] = m.split(':');
        const s = decodeAddr(start);
        const e = decodeAddr(end);

        // ⭐ skip detail row
        const isDetailRow = sections.some(sec =>
            sec.key.toLowerCase() === 'detail' &&
            s.r >= sec.start && s.r <= sec.end
        );

        if (!isDetailRow) {
            code.push(`sheet.AddMergedRegion(new CellRangeAddress(${s.r - 1}, ${e.r - 1}, ${s.c - 1}, ${e.c - 1}));`);
        }
    });
}
    code.push(``);
    code.push(`using (var stream = new MemoryStream())`);
    code.push(`{`);
    code.push(`     workbook.Write(stream);`);
    code.push(`     return stream.ToArray();`);
    code.push(`}`);


    return code.join('\n');
}


