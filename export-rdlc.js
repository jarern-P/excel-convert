// ============================================
// RDLC Export - Excel Report Code Generator
// ============================================

const SNAP_GRID_PT = 0.25;

function exportToRDLC({ ws: wsJson, workbook: wbJson, meta, images = [] }) {
  const ws = buildWorksheetFromJson(wsJson);
  const workbook = buildWorkbookFromJson(wbJson);

  let xml = [];
  const themeColors = getThemeColors(workbook);

  const A4_PORTRAIT_WIDTH = 595;
  const A4_PORTRAIT_HEIGHT = 842;
  const A4_LANDSCAPE_WIDTH = 842;
  const A4_LANDSCAPE_HEIGHT = 595;

  // ===== CALCULATE TOTAL WIDTH =====
  let totalWidthPt = 0;
  for (let i = 1; i <= ws.columnCount; i++) {
    const col = ws.getColumn(i);
    totalWidthPt += getColWidthPt(col.width);
  }

  const leftMargin = 20;
  const rightMargin = 20;
  const topMargin = 20;
  const bottomMargin = 20;

  const portraitUsableWidth = 520;
  const isLandscape = totalWidthPt > portraitUsableWidth;

  const pageWidth = isLandscape ? A4_LANDSCAPE_WIDTH : A4_PORTRAIT_WIDTH;
  const pageHeight = isLandscape ? A4_LANDSCAPE_HEIGHT : A4_PORTRAIT_HEIGHT;
  const orientation = isLandscape ? 'Landscape' : 'Portrait';

  const columnWidth = pageWidth - leftMargin - rightMargin;
  const scaleX = totalWidthPt > 0 ? columnWidth / totalWidthPt : 1;
  const scaleY = 1;
  const isSnapToGrid = true;

  // ===== BUILD SECTIONS (keys are already normalized by buildSections) =====
  const sections = buildSections(meta, ws.rowCount || ws.lastRow?.number || 0);

  sections.forEach(section => {
    let bandHeight = 0;
    for (let r = section.start; r <= section.end; r++) {
      const row = ws.getRow(r);
      bandHeight += getRowHeightPt(row.height, ws.properties.defaultRowHeight);
    }
    section.height = bandHeight;
  });

  // ===== COLLECT FIELDS & PARAMETERS =====
  const detailSection = sections.find(x => x.key === 'detail');
  const pageHeaderSection = sections.find(x => x.key === 'pageheader');
  const titleSection = sections.find(x => x.key === 'title');
  const pageFooterSection = sections.find(x => x.key === 'pagefooter');
  const columnHeaderSection = sections.find(x => x.key === 'columnheader');

  const detailFields = collectFieldsFromSection(ws, detailSection);
  const detailParams = collectParamsFromSection(ws, detailSection);

  const allFields = collectFields(ws);

  // ===== BUILD XML PARTS =====
  const xmlFields = buildRdlcFields(allFields);
  const xmlParameters = buildRdlcParameters(detailParams);
  const xmlEmbeddedImages = buildEmbeddedImages(workbook, images);

  // ===== PAGE HEADER (includes title + pageheader + columnheader) =====
  const pageHeaderPt = getSectionHeight(sections, ws, 'pageheader');
  const columnHeaderPt = getSectionHeight(sections, ws, 'columnheader');
  const titlePt = getSectionHeight(sections, ws, 'title');
  const pageHeaderHeightPt = titlePt + pageHeaderPt + columnHeaderPt;
  const pageHeaderHeightCm = ptToCm(pageHeaderHeightPt > 0 ? pageHeaderHeightPt : 0);

  // Get cells from ALL page-header-related sections
  const pageHeaderImages = filterItemsBySection(images, pageHeaderSection);
  const xmlPageHeaderImages = buildImageReportItems(pageHeaderImages, {
    ws, scaleX, scaleY, startRow: pageHeaderSection?.start || 1, isSnapToGrid
  });

  const titleCells = getCellsBySection(ws, titleSection);
  const pageHeaderCells = getCellsBySection(ws, pageHeaderSection);
  const columnHeaderCells = getCellsBySection(ws, columnHeaderSection);
  const allPageHeaderCells = [...titleCells, ...pageHeaderCells, ...columnHeaderCells];

  const xmlPageHeaderTextboxes = buildTextboxItems(allPageHeaderCells, {
    ws, scaleX, scaleY, startRow: titleSection?.start || pageHeaderSection?.start || 1, themeColors, isSnapToGrid
  });

  // ===== BODY / DETAIL =====
  const bodyImages = filterItemsBySection(images, detailSection);
  const xmlBodyImages = buildImageReportItems(bodyImages, {
    ws, scaleX, scaleY, startRow: detailSection?.start || 1, isSnapToGrid
  });

  const xmlBodyTablix = buildTablix(ws, detailSection, {
    scaleX, scaleY, fields: allFields, params: detailParams, themeColors, isSnapToGrid
  });

  const detailHeightPt = getSectionHeight(sections, ws, 'detail');
  const detailHeightCm = ptToCm(detailHeightPt > 0 ? detailHeightPt : 30);

  // ===== PAGE FOOTER =====
  const pageFooterHeight = getSectionHeightCm(sections, ws, 'pagefooter');
  const pageFooterCells = getCellsBySection(ws, pageFooterSection);
  const xmlPageFooterTextboxes = buildTextboxItems(pageFooterCells, {
    ws, scaleX, scaleY, startRow: pageFooterSection?.start || 1, themeColors, isSnapToGrid
  });

  const reportWidthCm = ptToCm(columnWidth);

  // ===== CHECK FOR EMPTY ReportItems (RDLC schema requires at least 1 child element) =====
  const bodyReportItemsXml = (() => {
    const items = [xmlBodyImages, xmlBodyTablix].filter(s => s.trim()).join('\n          ');
    if (items) return `<ReportItems>\n          ${items}\n        </ReportItems>`;
    return `<ReportItems>\n          <Textbox Name="Placeholder"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value></Value></TextRun></TextRuns></Paragraph></Paragraphs><Visibility><Hidden>true</Hidden></Visibility><Style /></Textbox>\n        </ReportItems>`;
  })();

  const pageHeaderReportItemsXml = (() => {
    const items = [xmlPageHeaderImages, xmlPageHeaderTextboxes].filter(s => s.trim()).join('\n            ');
    if (items) return `<ReportItems>\n            ${items}\n          </ReportItems>`;
    return '';
  })();

  const pageFooterReportItemsXml = (() => {
    const items = xmlPageFooterTextboxes.trim();
    if (items) return `<ReportItems>\n            ${items}\n          </ReportItems>`;
    return '';
  })();

  // ===== BUILD XML =====
  xml.push(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition" xmlns:rd="http://schemas.microsoft.com/SQLServer/reporting/reportdesigner">
  <AutoRefresh>0</AutoRefresh>
  <DataSources>
    <DataSource Name="ReportDataSource">
      <ConnectionProperties>
        <DataProvider>System.Data.DataSet</DataProvider>
        <ConnectString>/* Local Connection */</ConnectString>
      </ConnectionProperties>
      <rd:DataSourceID>3809a38b-8d09-4af0-a5a7-7c611986f90c</rd:DataSourceID>
    </DataSource>
  </DataSources>
  <DataSets>
    <DataSet Name="ReportDataSet">
      <Query>
        <DataSourceName>ReportDataSource</DataSourceName>
        <CommandText>/* Local Query */</CommandText>
      </Query>
      ${xmlFields}
    </DataSet>
  </DataSets>
  ${xmlParameters}
  <ReportSections>
    <ReportSection>
      <Body>
        ${bodyReportItemsXml}
        <Height>${detailHeightCm}cm</Height>
        <Style />
      </Body>
      <Width>${reportWidthCm}cm</Width>
      <Page>
        <PageHeader>
          <Height>${pageHeaderHeightCm}cm</Height>
          ${pageHeaderReportItemsXml}
          <PrintOnFirstPage>true</PrintOnFirstPage>
          <PrintOnLastPage>true</PrintOnLastPage>
          <Style />
        </PageHeader>
        ${pageFooterSection ? `
        <PageFooter>
          <Height>${pageFooterHeight}cm</Height>
          ${pageFooterReportItemsXml}
          <PrintOnFirstPage>true</PrintOnFirstPage>
          <PrintOnLastPage>true</PrintOnLastPage>
          <Style />
        </PageFooter>` : ''}
        <PageHeight>${ptToCm(pageHeight)}cm</PageHeight>
        <PageWidth>${ptToCm(pageWidth)}cm</PageWidth>
        <LeftMargin>${ptToCm(leftMargin)}cm</LeftMargin>
        <RightMargin>${ptToCm(rightMargin)}cm</RightMargin>
        <TopMargin>${ptToCm(topMargin)}cm</TopMargin>
        <BottomMargin>${ptToCm(bottomMargin)}cm</BottomMargin>
        <Style />
      </Page>
    </ReportSection>
  </ReportSections>
  ${xmlEmbeddedImages}
  <rd:ReportUnitType>Cm</rd:ReportUnitType>
  <rd:ReportID>c3cb1c9e-ac03-4836-b2df-46a9b96ceb16</rd:ReportID>
</Report>`);

  return xml.join('\n');
}

// ============================================
// FIELD COLLECTION (detail section only)
// ============================================
function collectFieldsFromSection(ws, section) {
  if (!section) return [];
  const fields = [];

  for (let r = section.start; r <= section.end; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell) => {
      const value = getCellDisplayValue(cell);
      const expr = parseRdlcExpression(value);
      if (expr?.type === 'field') {
        const exists = fields.find(x => x.name === expr.name);
        if (!exists) {
          fields.push({
            name: expr.name,
            type: inferFieldType(cell)
          });
        }
      }
    });
  }

  return fields;
}

function collectParamsFromSection(ws, section) {
  if (!section) return [];
  const params = [];

  for (let r = section.start; r <= section.end; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell) => {
      const value = getCellDisplayValue(cell);
      const expr = parseRdlcExpression(value);
      if (expr?.type === 'parameter') {
        const exists = params.find(x => x.name === expr.name);
        if (!exists) {
          params.push({
            name: expr.name,
            type: mapParamType(inferFieldType(cell))
          });
        }
      }
    });
  }

  return params;
}

function mapParamType(fieldType) {
  switch (fieldType) {
    case 'java.lang.Integer': return 'Integer';
    case 'java.math.BigDecimal': return 'Decimal';
    case 'java.util.Date': return 'DateTime';
    default: return 'String';
  }
}

// ============================================
// RDLC EXPRESSION PARSING
// ============================================
function parseRdlcExpression(value) {
  if (value == null) return null;

  const str = String(value).trim();
  if (!str) return null;

  // P{{paramName}} → Parameter
  let match = str.match(/^[pP]\{\{(.*?)\}\}$/);
  if (match) {
    return { type: 'parameter', name: match[1].trim() };
  }

  // {{fieldName}} → Field
  match = str.match(/^\{\{(.*?)\}\}$/);
  if (match) {
    return { type: 'field', name: match[1].trim() };
  }

  // Mixed text with {{field}} or P{{param}}
  if (/\{\{/.test(str)) {
    // Split into parts: alternating text and field/parameter references
    let parts = [];
    let lastIndex = 0;
    const regex = /(?:[pP]\{\{(.*?)\}\}|\{\{(.*?)\}\})/g;
    let match;

    while ((match = regex.exec(str)) !== null) {
      // Text before this match
      if (match.index > lastIndex) {
        const text = str.slice(lastIndex, match.index);
        if (text) {
          parts.push({ type: 'text', value: text });
        }
      }
      // Field or parameter
      if (match[1]) {
        parts.push({ type: 'parameter', name: match[1].trim() });
      } else {
        parts.push({ type: 'field', name: match[2].trim() });
      }
      lastIndex = match.index + match[0].length;
    }
    // Remaining text after last match
    if (lastIndex < str.length) {
      const text = str.slice(lastIndex);
      if (text) {
        parts.push({ type: 'text', value: text });
      }
    }

    // Build expression: text parts wrapped in "...", field/param as direct references
    const exprParts = parts.map(p => {
      switch (p.type) {
        case 'text':
          return `"${escapeRdlcString(p.value)}"`;
        case 'field':
          return `Fields!${p.name}.Value`;
        case 'parameter':
          return `Parameters!${p.name}.Value`;
        default:
          return '';
      }
    });

    return {
      type: 'mixed',
      expression: exprParts.join(' & ')
    };
  }

  return null;
}

function formatRdlcExpression(expr, value) {
  // For non-expression values, return escaped value with space fallback for empty cells
  if (!expr) {
    return escapeXml(value === null || value === undefined || value === '' ? ' ' : value);
  }

  let result;
  switch (expr.type) {
    case 'field':
      result = `=Fields!${expr.name}.Value`;
      break;
    case 'parameter':
      result = `=Parameters!${expr.name}.Value`;
      break;
    case 'mixed':
      result = `=${expr.expression}`;
      break;
    default:
      result = String(value);
  }

  return escapeXml(result);
}

// ============================================
// BUILD FIELDS (RDLC XML)
// ============================================
function buildRdlcFields(fields) {
  if (!fields || fields.length === 0) return '';
  let xml = ['<Fields>'];

  fields.forEach(f => {
    const rdlcType = mapToRdlcType(f.type);
    xml.push(`<Field Name="${escapeXml(f.name)}">
  <DataField>${escapeXml(f.name)}</DataField>
  <rd:TypeName>${rdlcType}</rd:TypeName>
</Field>`);
  });

  xml.push('</Fields>');
  return xml.join('\n');
}

function mapToRdlcType(javaType) {
  switch (javaType) {
    case 'java.lang.Integer': return 'System.Int32';
    case 'java.math.BigDecimal': return 'System.Decimal';
    case 'java.util.Date': return 'System.DateTime';
    case 'java.lang.Boolean': return 'System.Boolean';
    default: return 'System.String';
  }
}

// ============================================
// BUILD PARAMETERS (RDLC XML)
// ============================================
function buildRdlcParameters(params) {
  if (!params || params.length === 0) return '';
  let xml = ['<ReportParameters>'];

  params.forEach(p => {
    xml.push(`<ReportParameter Name="${escapeXml(p.name)}">
  <DataType>${p.type || 'String'}</DataType>
  <Prompt>${escapeXml(p.name)}</Prompt>
</ReportParameter>`);
  });

  xml.push('</ReportParameters>');
  return xml.join('\n');
}

// ============================================
// EMBEDDED IMAGES
// ============================================
function buildEmbeddedImages(workbook, images = []) {
  if (!images.length) return '';
  let xml = [];
  xml.push('<EmbeddedImages>');

  images.forEach(img => {
    const image = workbook.getImage(img.imageId);
    if (!image?.buffer) return;

    const base64 = toBase64(image.buffer);
    xml.push(`<EmbeddedImage Name="${getImageName(img)}">
  <MIMEType>${escapeXml(img.mimeType || 'image/png')}</MIMEType>
  <ImageData>${base64}</ImageData>
</EmbeddedImage>`);
  });

  xml.push('</EmbeddedImages>');
  return xml.join('\n');
}

// ============================================
// IMAGE REPORT ITEMS
// ============================================
function buildImageReportItems(images = [], options = {}) {
  const { ws, scaleX = 1, scaleY = 1, startRow = 1, isSnapToGrid = false } = options;
  if (!images.length) return '';

  let xml = [];
  images.forEach((img, index) => {
    const tl = img.range?.tl;
    const br = img.range?.br;

    if (!tl || !br) return;

    let leftPt, topPt, widthPt, heightPt;

    if (isSnapToGrid) {
      const gridCols = buildGridColumns(ws, scaleX, SNAP_GRID_PT);
      const gridRows = buildGridRows(ws, scaleY, SNAP_GRID_PT);
      const baseY = gridRows[startRow - 1];

      const startCol = tl.nativeCol + 1;
      const endCol = br.nativeCol + 1;
      const startRow2 = tl.nativeRow + 1;
      const endRow2 = br.nativeRow + 1;

      leftPt = gridCols[startCol - 1];
      widthPt = gridCols[endCol] - gridCols[startCol - 1];
      topPt = gridRows[startRow2 - 1] - baseY;
      heightPt = gridRows[endRow2] - gridRows[startRow2 - 1];
    } else {
      const bounds = getImageBoundsPt(ws, img, scaleX, scaleY, startRow);
      leftPt = bounds.x;
      topPt = bounds.y;
      widthPt = bounds.width;
      heightPt = bounds.height;
    }

    const left = ptToCm(leftPt);
    const top = ptToCm(topPt);
    const width = ptToCm(widthPt);
    const height = ptToCm(heightPt);

    xml.push(`<Image Name="Image${index + 1}">
  <Source>Embedded</Source>
  <Value>${getImageName(img)}</Value>
  <Sizing>FitProportional</Sizing>
  <Top>${top}cm</Top>
  <Left>${left}cm</Left>
  <Width>${width}cm</Width>
  <Height>${height}cm</Height>
  <ZIndex>1</ZIndex>
  <Style />
</Image>`);
  });

  return xml.join('\n');
}

// ============================================
// TABLIX (single dynamic row with detail group)
// ============================================
function buildTablix(ws, section, options = {}) {
  const { scaleX = 1, scaleY = 1, fields = [], params = [], themeColors, isSnapToGrid = false } = options;

  if (!section) {
    return '';
  }

  const columns = buildTablixColumns(ws, scaleX, isSnapToGrid);

  // ===== MULTI-LINE DETAIL: one TablixRow per detail template row =====
  // Each data record renders ALL rows of the detail section (section.start .. section.end).
  const tablixRows = [];
  let totalHeightPt = 0;
  for (let r = section.start; r <= section.end; r++) {
    const cellContents = buildTablixCellContents(ws, r, fields, params, themeColors);

    // Height for this template row
    let rowHeightPt = 0;
    if (isSnapToGrid) {
      const gridRows = buildGridRows(ws, scaleY, SNAP_GRID_PT);
      rowHeightPt = gridRows[r] - gridRows[r - 1];
    } else {
      rowHeightPt = getRowHeightPt(ws.getRow(r).height, ws.properties.defaultRowHeight);
    }
    totalHeightPt += rowHeightPt;
    tablixRows.push({ heightCm: ptToCm(rowHeightPt > 0 ? rowHeightPt : 15), cells: cellContents });
  }
  const totalHeightCm = ptToCm(totalHeightPt > 0 ? totalHeightPt : 30);

  const xmlTablixRows = tablixRows.map(tr => `<TablixRow>
        <Height>${tr.heightCm}cm</Height>
        <TablixCells>
          ${tr.cells}
        </TablixCells>
      </TablixRow>`).join('\n      ');

  // Row hierarchy: DetailGroup on the group member; one plain TablixMember per row
  const rowHierarchyMembers = `<TablixMember>
        <Group Name="DetailGroup" />
        <TablixMembers>
          ${'<TablixMember />'.repeat(tablixRows.length)}
        </TablixMembers>
      </TablixMember>`;

  // Calculate total width
  let totalWidthPt = 0;
  if (isSnapToGrid) {
    const gridCols = buildGridColumns(ws, scaleX, SNAP_GRID_PT);
    totalWidthPt = gridCols[ws.columnCount];
  } else {
    for (let c = 1; c <= ws.columnCount; c++) {
      totalWidthPt += getColWidthPt(ws.getColumn(c).width) * scaleX;
    }
  }
  const totalWidthCm = ptToCm(totalWidthPt);

  return `<Tablix Name="Tablix1">
  <TablixBody>
    <TablixColumns>
      ${columns}
    </TablixColumns>
    <TablixRows>
      ${xmlTablixRows}
    </TablixRows>
  </TablixBody>
  <TablixColumnHierarchy>
    <TablixMembers>
      ${'<TablixMember />'.repeat(ws.columnCount)}
    </TablixMembers>
  </TablixColumnHierarchy>
  <TablixRowHierarchy>
    <TablixMembers>
      ${rowHierarchyMembers}
    </TablixMembers>
  </TablixRowHierarchy>
  <DataSetName>ReportDataSet</DataSetName>
  <Top>0cm</Top>
  <Left>0cm</Left>
  <Height>${totalHeightCm}cm</Height>
  <Width>${totalWidthCm}cm</Width>
  <NoRowsMessage>= "No data available"</NoRowsMessage>
  <Style />
</Tablix>`;
}

function buildTablixColumns(ws, scaleX, isSnapToGrid) {
  let xml = [];
  const gridCols = isSnapToGrid ? buildGridColumns(ws, scaleX, SNAP_GRID_PT) : null;
  for (let c = 1; c <= ws.columnCount; c++) {
    let widthPt;
    if (isSnapToGrid) {
      widthPt = gridCols[c] - gridCols[c - 1];
    } else {
      const col = ws.getColumn(c);
      widthPt = getColWidthPt(col.width) * scaleX;
    }
    xml.push(`<TablixColumn>
  <Width>${ptToCm(widthPt)}cm</Width>
</TablixColumn>`);
  }
  return xml.join('\n');
}

function buildTablixCellContents(ws, templateRow, fields, params, themeColors) {
  let xml = [];

  for (let c = 1; c <= ws.columnCount; c++) {
    const merge = getMergeRange(ws, templateRow, c);

    // If this is a slave merged cell, skip
    if (merge && (merge.startRow !== templateRow || merge.startCol !== c)) {
      xml.push('<TablixCell/>');
      continue;
    }

    const cell = ws.getRow(templateRow).getCell(c);
    const value = getCellDisplayValue(cell);
    const expr = parseRdlcExpression(value);

    const colSpan = merge ? merge.endCol - merge.startCol + 1 : 1;
    const rowSpan = merge ? merge.endRow - merge.startRow + 1 : 1;

    // ===== RPPRINTIF / RP from cell note =====
    const noteText = getCellNoteText(cell);
    const rpFormula = parseRPPRINTIF(noteText);
    let rdlcHidden = '';
    let rdlcValue = '';

    if (rpFormula) {
      if (rpFormula.type === 'printWhen') {
        const condition = convertToRdlcConditionExpression(rpFormula.expression);
        rdlcHidden = `\n      <Visibility><Hidden>=Not(${condition})</Hidden></Visibility>`;
      } else if (rpFormula.type === 'textExpression') {
        const textExpr = convertToRdlcConditionExpression(rpFormula.expression);
        rdlcValue = `=${textExpr}`;
      }
    }

    const displayValue = rdlcValue || formatRdlcExpression(expr, value);

    const textboxStyle = buildTextboxStyle(cell, themeColors);
    const textRunStyle = buildTextRunStyle(cell, themeColors);
    const paragraphStyle = buildParagraphStyle(cell);

    xml.push(`<TablixCell>
  <CellContents>
    <Textbox Name="Textbox_${templateRow}_${c}">
      <CanGrow>true</CanGrow>${rdlcHidden}
      <Paragraphs>
        <Paragraph>
          <TextRuns>
            <TextRun>
              <Value>${displayValue}</Value>
              ${textRunStyle}
            </TextRun>
          </TextRuns>
          ${paragraphStyle}
        </Paragraph>
      </Paragraphs>
      <Style>
        ${textboxStyle}
      </Style>
    </Textbox>
    ${colSpan > 1 ? `<ColSpan>${colSpan}</ColSpan>` : ''}
    ${rowSpan > 1 ? `<RowSpan>${rowSpan}</RowSpan>` : ''}
  </CellContents>
</TablixCell>`);
  }

  return xml.join('\n');
}

// ============================================
// TEXTBOX ITEMS (for page header/footer)
// ============================================
function buildTextboxItems(items, options = {}) {
  const { ws, startRow = 1, scaleX = 1, scaleY = 1, themeColors, isSnapToGrid = false } = options;
  let xml = [];
  let rendered = new Set();

  items.forEach((x, index) => {
    const cell = x.cell;
    const valueRaw = getCellDisplayValue(cell);
    const expr = parseRdlcExpression(valueRaw);

    const hasBorder = !!(cell.border?.top?.style || cell.border?.bottom?.style ||
      cell.border?.left?.style || cell.border?.right?.style);
    const hasFill = !!cell.fill && cell.fill.type !== 'none' && cell.fill.pattern !== 'none';

    if (!valueRaw && !hasBorder && !hasFill) return;

    const merge = getMergeRange(ws, x.row, x.col);
    if (merge && (merge.startRow !== x.row || merge.startCol !== x.col)) return;

    const key = `${x.row}_${x.col}`;
    if (rendered.has(key)) return;
    rendered.add(key);

    const startCol = merge?.startCol || x.col;
    const endCol = merge?.endCol || x.col;
    const startRow2 = merge?.startRow || x.row;
    const endRow = merge?.endRow || x.row;

    let leftPt, topPt, widthPt, heightPt;

    if (isSnapToGrid) {
      const gridCols = buildGridColumns(ws, scaleX, SNAP_GRID_PT);
      const gridRows = buildGridRows(ws, scaleY, SNAP_GRID_PT);
      const baseY = gridRows[startRow - 1];

      leftPt = gridCols[startCol - 1];
      widthPt = gridCols[endCol] - gridCols[startCol - 1];
      topPt = gridRows[startRow2 - 1] - baseY;
      heightPt = gridRows[endRow] - gridRows[startRow2 - 1];
    } else {
      leftPt = getColumnLeftPt(ws, startCol) * scaleX;
      topPt = getRowTopPt(ws, startRow2, startRow) * scaleY;
      widthPt = getColumnsWidthPt(ws, startCol, endCol) * scaleX;
      heightPt = getRowsHeightPt(ws, startRow2, endRow) * scaleY;
    }

    const left = ptToCm(leftPt);
    const top = ptToCm(topPt);
    const width = ptToCm(widthPt);
    const height = ptToCm(heightPt);
    const textboxStyle = buildTextboxStyle(cell, themeColors);
    const textRunStyle = buildTextRunStyle(cell, themeColors);
    const paragraphStyle = buildParagraphStyle(cell);

    // ===== RPPRINTIF / RP from cell note =====
    const noteText = getCellNoteText(cell);
    const rpFormula = parseRPPRINTIF(noteText);
    let rdlcHidden = '';
    let rdlcValue = '';

    if (rpFormula) {
      if (rpFormula.type === 'printWhen') {
        const condition = convertToRdlcConditionExpression(rpFormula.expression);
        rdlcHidden = `\n  <Visibility><Hidden>=Not(${condition})</Hidden></Visibility>`;
      } else if (rpFormula.type === 'textExpression') {
        const textExpr = convertToRdlcConditionExpression(rpFormula.expression);
        rdlcValue = `=${textExpr}`;
      }
    }

    const displayValue = rdlcValue || formatRdlcExpression(expr, valueRaw);

    xml.push(`<Textbox Name="Textbox_${index}">
  <CanGrow>true</CanGrow>${rdlcHidden}
  <Top>${top}cm</Top>
  <Left>${left}cm</Left>
  <Width>${width}cm</Width>
  <Height>${height}cm</Height>
  <Paragraphs>
    <Paragraph>
      <TextRuns>
        <TextRun>
          <Value>${displayValue}</Value>
          ${textRunStyle}
        </TextRun>
      </TextRuns>
      ${paragraphStyle}
    </Paragraph>
  </Paragraphs>
  <Style>
    ${textboxStyle}
  </Style>
</Textbox>`);
  });

  return xml.join('\n');
}

// ============================================
// GEOMETRY HELPERS
// ============================================
function getColumnLeftPt(ws, colNumber) {
  let pt = 0;
  for (let i = 1; i < colNumber; i++) {
    pt += getColWidthPt(ws.getColumn(i).width);
  }
  return pt;
}

function getRowTopPt(ws, rowNumber, startRow = 1) {
  let pt = 0;
  for (let i = startRow; i < rowNumber; i++) {
    pt += getRowHeightPt(ws.getRow(i).height, ws.properties.defaultRowHeight);
  }
  return pt;
}

function getColumnsWidthPt(ws, startCol, endCol) {
  let pt = 0;
  for (let c = startCol; c <= endCol; c++) {
    pt += getColWidthPt(ws.getColumn(c).width);
  }
  return pt;
}

function getRowsHeightPt(ws, startRow, endRow) {
  let pt = 0;
  for (let r = startRow; r <= endRow; r++) {
    pt += getRowHeightPt(ws.getRow(r).height, ws.properties.defaultRowHeight);
  }
  return pt;
}

// ============================================
// SNAP GRID HELPERS
// ============================================
function buildGridColumns(ws, scaleX = 1, grid = SNAP_GRID_PT) {
  const positions = [0];
  let currentX = 0;

  for (let c = 1; c <= ws.columnCount; c++) {
    const width = getColWidthPt(ws.getColumn(c).width) * scaleX;
    const snappedWidth = Math.round(width / grid) * grid;
    currentX += snappedWidth;
    positions[c] = currentX;
  }

  return positions;
}

function buildGridRows(ws, scaleY = 1, grid = SNAP_GRID_PT) {
  const positions = [0];
  let currentY = 0;
  const rowCount = ws.rowCount || ws.lastRow?.number || 0;

  for (let r = 1; r <= rowCount; r++) {
    const row = ws.getRow(r);
    const height = getRowHeightPt(row.height, ws.properties.defaultRowHeight) * scaleY;
    currentY += height;
    positions[r] = Math.round(currentY / grid) * grid;
  }

  return positions;
}

// ============================================
// MERGE HELPERS
// ============================================
function getMergeRange(ws, row, col) {
  const merges = ws.model.merges || [];
  for (const m of merges) {
    const range = parseRange(m);
    if (row >= range.startRow && row <= range.endRow &&
        col >= range.startCol && col <= range.endCol) {
      return range;
    }
  }
  return null;
}

function parseRange(range) {
  const [start, end] = range.split(':');
  const s = decodeCell(start);
  const e = decodeCell(end);
  return {
    startRow: s.row, startCol: s.col,
    endRow: e.row, endCol: e.col
  };
}

function decodeCell(address) {
  const match = address.match(/([A-Z]+)(\d+)/);
  const colLetters = match[1];
  const row = parseInt(match[2]);
  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { row, col };
}

// ============================================
// STYLE BUILDERS — TextRun level (Font properties)
// ============================================
function buildTextRunStyle(cell, themeColors) {
  // In RDLC, font properties MUST be at the TextRun <Style> level,
  // because inheritance from Textbox/Paragraph is NOT reliable.
  let xml = [];

  const font = cell.font || {};
  xml.push(`<FontFamily>${escapeXml(font.name || 'Calibri')}</FontFamily>`);
  xml.push(`<FontSize>${font.size || 11}pt</FontSize>`);
  if (font.bold) {
    xml.push('<FontWeight>Bold</FontWeight>');
  }
  if (font.italic) {
    xml.push('<FontStyle>Italic</FontStyle>');
  }
  if (font.underline) {
    xml.push('<TextDecoration>Underline</TextDecoration>');
  }
  if (font.strike) {
    xml.push('<TextDecoration>LineThrough</TextDecoration>');
  }
  const argbFont = toARGB(font.color, themeColors);
  if (argbFont) {
    xml.push(`<Color>${argbToHex(argbFont)}</Color>`);
  }

  return `<Style>\n          ${xml.join('\n          ')}\n        </Style>`;
}

// ============================================
// STYLE BUILDERS — Paragraph level (Alignment)
// ============================================
function buildParagraphStyle(cell) {
  // In RDLC, TextAlign and VerticalAlign go at the Paragraph <Style> level.
  let xml = [];

  if (cell.alignment) {
    const hAlign = mapHorizontalAlign(cell.alignment.horizontal || 'general');
    if (hAlign) xml.push(`<TextAlign>${hAlign}</TextAlign>`);
    const vAlign = mapVerticalAlign(cell.alignment.vertical || '');
    if (vAlign) xml.push(`<VerticalAlign>${vAlign}</VerticalAlign>`);
  } else {
    // No alignment object → use Excel defaults
    xml.push('<TextAlign>Left</TextAlign>');
    xml.push('<VerticalAlign>Bottom</VerticalAlign>');
  }

  return `<Style>\n        ${xml.join('\n        ')}\n      </Style>`;
}

// ============================================
// STYLE BUILDERS — Textbox level (Background, Border, Padding)
// ============================================
function buildTextboxStyle(cell, themeColors) {
  // At the Textbox level: BackgroundColor, Border, Padding only.
  // Font and Alignment are at TextRun and Paragraph levels respectively.
  let xml = [];

  // BACKGROUND
  const argbBg = toARGB(cell.fill?.fgColor, themeColors);
  if (argbBg) {
    xml.push(`<BackgroundColor>${argbToHex(argbBg)}</BackgroundColor>`);
  }

  // BORDER
  xml.push(buildBorderStyle(cell.border, themeColors));

  // PADDING
  xml.push(`<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>1pt</PaddingTop>
<PaddingBottom>1pt</PaddingBottom>`);

  return xml.join('\n');
}

function mapHorizontalAlign(value) {
  switch ((value || '').toLowerCase()) {
    case 'center': return 'Center';
    case 'right': return 'Right';
    case 'justify': return 'Justify';
    case 'general':
    default: return 'Left';
  }
}

function mapVerticalAlign(value) {
  switch ((value || '').toLowerCase()) {
    case 'top': return 'Top';
    case 'middle': return 'Middle';
    case 'bottom': return 'Bottom';
    default: return 'Bottom';
  }
}

function buildBorderStyle(border, themeColors) {
  if (!border || (typeof border === 'object' && Object.keys(border).length === 0)) return '';
  let xml = [];

  const sides = [
    ['Top', border.top],
    ['Bottom', border.bottom],
    ['Left', border.left],
    ['Right', border.right]
  ];

  sides.forEach(([name, side]) => {
    // Must have a valid style to render — matches JRXML behavior
    if (!side?.style) return;
    xml.push(`<${name}Border>\n  <Color>${getBorderColor(side, themeColors)}</Color>\n  <Style>${mapBorderStyle(side.style)}</Style>\n  <Width>${mapBorderWidth(side.style)}</Width>\n</${name}Border>`);
  });

  return xml.join('\n');
}

function mapBorderStyle(style) {
  // Map ExcelJS border styles to RDLC border styles
  switch (style) {
    case 'double': return 'Double';
    case 'dashed':
    case 'mediumDashed':
      return 'Dashed';
    case 'dotted': return 'Dotted';
    case 'dashDot':
    case 'mediumDashDot':
    case 'slantDashDot':
      return 'DashDot';
    case 'dashDotDot':
    case 'mediumDashDotDot':
      return 'DashDotDot';
    default:
      // thin, medium, thick, hair → Solid
      return 'Solid';
  }
}

function mapBorderWidth(style) {
  switch (style) {
    case 'hair': return '0.25pt';
    case 'thin': return '0.5pt';
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
      return '1.5pt';
    case 'thick': return '2.5pt';
    default: return '0.5pt';
  }
}

function getBorderColor(side, themeColors) {
  const argb = toARGB(side.color, themeColors);
  if (argb) {
    return argbToHex(argb);
  }
  return '#000000';
}

// ============================================
// SECTION HELPERS
// ============================================
function getSectionHeightCm(sections, ws, ...keys) {
  const pt = getSectionHeight(sections, ws, ...keys);
  return ptToCm(pt);
}

function filterItemsBySection(items, section) {
  if (!section) return [];
  return items.filter(item => {
    const tl = item.range?.tl;
    if (!tl) return false;
    const row = tl.nativeRow + 1;
    return row >= section.start && row <= section.end;
  });
}

function getCellsBySection(ws, section) {
  if (!section) return [];
  let cells = [];
  for (let r = section.start; r <= section.end; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= ws.columnCount; c++) {
      cells.push({ cell: row.getCell(c), row: r, col: c });
    }
  }
  return cells;
}

// ============================================
// UNIT CONVERSION
// ============================================
function ptToCm(pt) {
  pt = Number(pt);
  if (isNaN(pt) || !isFinite(pt)) return '0';
  return (pt * 0.0352778).toFixed(3);
}

// ============================================
// IMAGE HELPERS
// ============================================
function getImageName(img) {
  const rawName = img.name || `Image_${img.imageId}`;
  return escapeXml(String(rawName).replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, ''));
}

function getImageBoundsPt(ws, img, scaleX = 1, scaleY = 1, sectionStart = 1) {
  const EMU_PER_PIXEL = 9525;
  const tl = img.range.tl;
  const br = img.range.br;

  let x = getColumnX(ws, tl.nativeCol + 1);
  let y = getRowY(ws, sectionStart, tl.nativeRow + 1);

  x += ((tl.nativeColOff || 0) / EMU_PER_PIXEL) * 0.75;
  y += ((tl.nativeRowOff || 0) / EMU_PER_PIXEL) * 0.75;

  let x2 = getColumnX(ws, br.nativeCol + 1);
  let y2 = getRowY(ws, sectionStart, br.nativeRow + 1);

  x2 += ((br.nativeColOff || 0) / EMU_PER_PIXEL) * 0.75;
  y2 += ((br.nativeRowOff || 0) / EMU_PER_PIXEL) * 0.75;

  x *= scaleX;
  y *= scaleY;
  x2 *= scaleX;
  y2 *= scaleY;

  return { x, y, width: x2 - x, height: y2 - y };
}

function getColumnX(ws, colNumber) {
  let x = 0;
  for (let c = 1; c < colNumber; c++) {
    x += getColWidthPt(ws.getColumn(c).width);
  }
  return x;
}

function getRowY(ws, startRow, currentRow) {
  let y = 0;
  for (let r = startRow; r < currentRow; r++) {
    y += getRowHeightPt(ws.getRow(r).height, ws.properties.defaultRowHeight);
  }
  return y;
}

// ============================================
// RDLC STRING ESCAPING (for string literals in expressions)
// ============================================
function escapeRdlcString(str) {
  return String(str).replace(/"/g, '""');
}

// ============================================
// XML ESCAPING
// ============================================
function escapeXml(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
