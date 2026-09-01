// ============================================
// Entity Framework Core Export - Code Generator
// ============================================
// Generates C# code matching the SQL script pattern:
//   - Repository (Interface + Implementation)
//   - Service (Interface + Implementation with CRUD)
//   - Entity class (with PK from Excel cell notes)
//   - DTO class
//   - API Controller (with CRUD endpoints)
//   - CREATE TABLE SQL script
// ============================================
// PK detection: Excel cell note/comment containing "PK" (case-insensitive)
// ============================================

function exportToEfCore({ ws: wsJson, workbook: wbJson, meta, images = [], fileName: sourceFileName = '' }) {
  const ws = buildWorksheetFromJson(wsJson);
  const workbook = buildWorkbookFromJson(wbJson);
  const sections = buildSections(meta, ws.rowCount || ws.lastRow?.number || 0);

  const columnHeaderSection = sections.find(s => s.key === 'columnheader');
  const detailSection = sections.find(s => s.key === 'detail');
  const firstDetailRow = detailSection ? ws.getRow(detailSection.start) : null;

  // ===== INFER TABLE & ENTITY NAMES =====
  const tableName = inferTableName(sourceFileName || 'Report');
  const entityName = inferEntityName(sourceFileName || 'Report');

  // Extract EX code for controller naming
  const baseName = (sourceFileName || 'Report').replace(/\.xlsx$/i, '').trim();
  const exMatch = baseName.match(/EX\d+/);
  const exCode = exMatch ? exMatch[0] : '';
  const exNumber = exCode ? exCode.replace('EX', '') : '';
  // const exportMethodName = exNumber ? 'ExportExcelEx' + exNumber : 'ExportExcel';
    const exportMethodName = exNumber ? 'ExportExcelAsync' : 'ExportExcelAsync';
  const controllerName = exCode
    ? exCode.charAt(0).toUpperCase() + exCode.slice(1).toLowerCase()
    : entityName;

  const namespace = "excisetaxService"
  const namespaceSub = entityName + 'Service';
  const schemaName = 'dbo';

  // ===== COLLECT COLUMN DEFINITIONS FROM HEADER =====
  const columns = [];
  const pkColumns = [];

  if (columnHeaderSection) {
    const headerRowIndex = columnHeaderSection.end;
    const headerRow = ws.getRow(headerRowIndex);
    const usedFields = new Set();

    headerRow.eachCell({ includeEmpty: true }, (cell, cIdx) => {
      const text = (cell.text || '').trim();
      if (!text) return;

      let propertyName = toPascalCase(text);
      let uniqueName = propertyName;
      let counter = 1;
      while (usedFields.has(uniqueName)) {
        uniqueName = `${propertyName}_${counter++}`;
      }
      usedFields.add(uniqueName);

      // Infer C# type: priority is Excel data cell value first.
      // Only use header name as fallback when data cell is empty (no value).
      let csType;
      if (firstDetailRow) {
        const dataCell = firstDetailRow.getCell(cIdx);
        csType = inferCsType(dataCell);
        // Only fall back to name-based inference if cell has no value (empty)
        if (csType === 'string' && dataCell.value == null) {
          csType = inferCsTypeFromName(text);
        }
      } else {
        csType = inferCsType(cell);
        if (csType === 'string') {
          csType = inferCsTypeFromName(text);
        }
      }
      const isNullable = csType !== 'string' && csType !== 'byte[]';

      // ===== PK & SIZE DETECTION from cell notes in detail row =====
      // Note format examples:
      //   "pk"            → Primary Key, default size
      //   "pk size(100)"  → Primary Key, NVARCHAR(100)
      //   "pk:100"        → Primary Key, NVARCHAR(100)
      //   "size(500)"     → NVARCHAR(500)
      //   "500"           → NVARCHAR(500)
      //   "max"           → NVARCHAR(MAX)
      //   "pk max"        → Primary Key, NVARCHAR(MAX)
      let isPk = false;
      let noteMaxLength = null;
      let decimalPrecision = null;
      let decimalScale = null;
      if (firstDetailRow) {
        const detailCell = firstDetailRow.getCell(cIdx);
        const noteText = getCellNoteText(detailCell);
        const parsed = parseCellNote(noteText);
        isPk = parsed.isPk;
        noteMaxLength = parsed.maxLength;
        decimalPrecision = parsed.decimalPrecision;
        decimalScale = parsed.decimalScale;
      }

      columns.push({
        header: text,
        columnName: uniqueName,
        csType: csType,
        isNullable: isNullable,
        displayIndex: cIdx,
        isPk: isPk,
        maxLength: noteMaxLength !== null ? noteMaxLength : 255,  // From cell note or default 255
        decimalPrecision: decimalPrecision,
        decimalScale: decimalScale
      });

      if (isPk) {
        pkColumns.push(uniqueName);
      }
    });
  } else {
    for (let c = 1; c <= ws.columnCount; c++) {
      const colName = `Column${c}`;
      columns.push({
        header: colName,
        columnName: colName,
        csType: 'string',
        isNullable: false,
        displayIndex: c,
        isPk: c === 1,
        maxLength: 255
      });
      if (c === 1) pkColumns.push(colName);
    }
  }

  // If no PK found via notes, default first column as PK
  if (pkColumns.length === 0 && columns.length > 0) {
    pkColumns.push(columns[0].columnName);
    columns[0].isPk = true;
  }

  // ====================================================
  // BUILD OUTPUT
  // ====================================================
  const pkList = pkColumns.join(', ');
  const firstPkColumn = pkColumns.length > 0 ? pkColumns[0] : 'Id';

  let code = [];
  const sep = (title) => [
    '',
    '// ' + '='.repeat(40),
    '// ' + title,
    '// ' + '='.repeat(40),
    ''
  ];

  // ====================================================
  // HEADER
  // ====================================================
  code.push('// ============================================');
  code.push('// GENERATED BY Excel Report Code Generator');
  code.push('// Source: EF Core Pattern from SQL Script');
  code.push('// ============================================');
  code.push('//');
  code.push('// To use this code:');
  code.push('//   1. Replace namespace "' + namespace + '" with your actual namespace');
  code.push('//   2. Install NuGet packages:');
  code.push('//      dotnet add package Microsoft.EntityFrameworkCore');
  code.push('//      dotnet add package Microsoft.EntityFrameworkCore.SqlServer');
  code.push('//      dotnet add package Mapster');
  code.push('//      dotnet add package Dapper');
  code.push('//      dotnet add package Hangfire');
  code.push('//      dotnet add package EPPlus (for Excel import)');
  code.push('//      dotnet add package NPOI (for Excel export)');
  code.push('//      dotnet add package iTextSharp (for PDF export)');
  code.push('// ============================================');

  // ====================================================
  // ENTITY
  // ====================================================
  sep('ENTITY').forEach(l => code.push(l));

  code.push('using System;');
  code.push('using System.ComponentModel.DataAnnotations;');
  code.push('using System.ComponentModel.DataAnnotations.Schema;');
  code.push('using Microsoft.EntityFrameworkCore;');
  code.push('');
  code.push('namespace ' + namespace + '.Entitys;');
  code.push('');
  code.push('[Table("' + tableName + '")]');
  if (pkColumns.length > 0) {
    code.push('[PrimaryKey(' + pkList + ')]');
  }
  code.push('public class ' + entityName + ' : BaseEntity');
  code.push('{');
  code.push('');

  columns.forEach((col, idx) => {
    // [Column] attribute
    let colAttr = '    [Column("' + col.columnName + '", Order = ' + idx;
    // Add TypeName for decimal types with precision/scale from cell note
    if (col.csType === 'decimal' && col.decimalPrecision !== null && col.decimalScale !== null) {
      colAttr += ', TypeName = "decimal(' + col.decimalPrecision + ',' + col.decimalScale + ')"';
    }
    colAttr += ')]';
    code.push(colAttr);

    // [MaxLength] for string types
    if (col.csType === 'string' && col.maxLength > 0) {
      code.push('    [MaxLength(' + col.maxLength + ')]');
    }

    // [Required] for non-nullable strings / decimals
    if (!col.isNullable && (col.csType === 'string' || col.csType === 'decimal')) {
      code.push('    [Required]');
    }

    // Property declaration
    let typeDecl = col.csType;
    if (col.isNullable && col.csType !== 'string') {
      typeDecl += '?';
    } else if (col.csType === 'string' && col.isNullable) {
      typeDecl += '?';
    }

    let defaultVal = '';
    if (col.csType === 'string' && !col.isNullable) {
      defaultVal = ' = string.Empty;';
    }

    code.push('    public ' + typeDecl + ' ' + col.columnName + ' { get; set; }' + defaultVal);
    code.push('');
  });

  code.push('}');
  code.push('');

  // ====================================================
  // DTO
  // ====================================================
  sep('DTO').forEach(l => code.push(l));

  code.push('using System;');
  code.push('');
  code.push('namespace ' + namespace + '.DTO;');
  code.push('');
  code.push('public class ' + entityName + 'DTO');
  code.push('{');
  code.push('');

  // Base DTO properties (Id, CreatedBy, UpdatedBy)
  code.push('    public string? Id { get; set; }');
  code.push('    public string? CreatedBy { get; set; }');
  code.push('    public string? UpdatedBy { get; set; }');
  code.push('');

  columns.forEach(col => {
    // Skip if column matches any base property name (case-insensitive)
    const colNameUpper = col.columnName.toUpperCase();
    if (colNameUpper === 'ID' || colNameUpper === 'CREATEDBY' || colNameUpper === 'UPDATEDBY') return;

    // DTO properties are nullable for search/filter usage
    let typeDecl = col.csType;
    typeDecl += '?';
    code.push('    public ' + typeDecl + ' ' + col.columnName + ' { get; set; }');
  });

  code.push('}');
  code.push('');

  // ====================================================
  // MODELS
  // ====================================================
  sep('MODELS').forEach(l => code.push(l));

  code.push('using System;');
  code.push('');
  code.push('namespace ' + namespace + '.Models;');
  code.push('');
  code.push('public class ResponseModel<T>');
  code.push('{');
  code.push('    public string Status { get; set; } = string.Empty;');
  code.push('    public T? Data { get; set; }');
  code.push('    public string Message { get; set; } = string.Empty;');
  code.push('    public string? ErrorMessage { get; set; }');
  code.push('    public bool IsSuccess => Status == "OK";');
  code.push('}');
  code.push('');
  code.push('public class ServiceResult<T>');
  code.push('{');
  code.push('    public T? Data { get; }');
  code.push('    public string? ErrorMessage { get; }');
  code.push('    public bool IsSuccess => ErrorMessage == null;');
  code.push('');
  code.push('    public ServiceResult(T? data)');
  code.push('    {');
  code.push('        Data = data;');
  code.push('    }');
  code.push('');
  code.push('    public ServiceResult(string? errorMessage)');
  code.push('    {');
  code.push('        ErrorMessage = errorMessage;');
  code.push('    }');
  code.push('');
  code.push('    public static implicit operator ServiceResult<T>(T value) => new(value);');
  code.push('}');
  code.push('');

  // ====================================================
  // REPOSITORY
  // ====================================================
  sep('REPOSITORY').forEach(l => code.push(l));

  code.push('using System.Collections.Generic;');
  code.push('using System.Linq;');
  code.push('using System.Threading.Tasks;');
  code.push('using ' + namespace + '.Contexts;');
  code.push('using ' + namespace + '.Entitys;');
  code.push('using Microsoft.EntityFrameworkCore;');
  code.push('');
  code.push('namespace ' + namespace + '.Repository;');
  code.push('');
  code.push('public interface I' + entityName + 'Repository : IRepository<' + entityName + '>');
  code.push('{');
  code.push('}');
  code.push('');
  code.push('public class ' + entityName + 'Repository : Repository<' + entityName + '>, I' + entityName + 'Repository');
  code.push('{');
  code.push('    public ' + entityName + 'Repository(DataContext context) : base(context)');
  code.push('    {');
  code.push('    }');
  code.push('}');
  code.push('');

  // ====================================================
  // SERVICE
  // ====================================================
  sep('SERVICE').forEach(l => code.push(l));

  code.push('using System;');
  code.push('using System.Threading.Tasks;');
  code.push('using ' + namespace + '.Contexts;');
  code.push('using ' + namespace + '.Models;');
  code.push('using ' + namespace + '.Repository;');
  code.push('using Microsoft.Extensions.Configuration;');
  code.push('using Microsoft.Extensions.Logging;');
  code.push('using System.Linq;');
  code.push('using System.Collections.Generic;');
  code.push('using ' + namespace + '.DTO;');
  code.push('using ' + namespace + '.Entitys;');
  code.push('using Microsoft.EntityFrameworkCore;');
  code.push('using NPOI.SS.UserModel;');
  code.push('using NPOI.XSSF.UserModel;');
  code.push('using System.IO;');
  code.push('using iTextSharp.text;');
  code.push('using iTextSharp.text.pdf;');
  code.push('');
  code.push('namespace ' + namespace + '.Services;');
  code.push('');
  code.push('public interface I' + entityName + 'Service');
  code.push('{');
  code.push('    Task<ServiceResult<List<' + entityName + '>>> SearchAsync(' + entityName + 'DTO data);');
  code.push('    Task<ServiceResult<' + entityName + '>> GetByIdAsync(string id);');
  code.push('    Task<ServiceResult<' + entityName + '>> AddAsync(' + entityName + 'DTO data);');
  code.push('    Task<ServiceResult<' + entityName + '>> UpdateAsync(' + entityName + 'DTO data);');
  code.push('    Task<ServiceResult<int>> DeleteAsync(' + entityName + 'DTO data);');
  code.push('    Task<byte[]> ' + exportMethodName + '(List<' + entityName + 'DTO> data);');
  code.push('    Task<byte[]> ExportPdfAsync(List<' + entityName + 'DTO> data);');
  code.push('}');
  code.push('');
  code.push('public class ' + entityName + 'Service : AbstractService, I' + entityName + 'Service');
  code.push('{');
  code.push('    private readonly ILogger<' + entityName + 'Service> _logger;');
  code.push('    private readonly DataContext _dbContext;');
  code.push('');
  code.push('    public ' + entityName + 'Service(');
  code.push('        ILogger<' + entityName + 'Service> logger,');
  code.push('        IConfiguration configuration,');
  code.push('        DataContext dbContext');
  code.push('        ) : base(logger, configuration)');
  code.push('    {');
  code.push('        _logger = logger;');
  code.push('        _dbContext = dbContext;');
  code.push('    }');
  code.push('');

  // SearchAsync
  code.push('    public async Task<ServiceResult<List<' + entityName + '>>> SearchAsync(' + entityName + 'DTO model)');
  code.push('    {');
  code.push('        var query = _dbContext.Set<' + entityName + '>().AsQueryable();');
  code.push('');

  // Generate dynamic Where filters for each string column
  const stringColumns = columns.filter(col => col.csType === 'string');
  stringColumns.forEach(col => {
    const colNameUpper = col.columnName.toUpperCase();
    if (colNameUpper === 'ID' || colNameUpper === 'CREATEDBY' || colNameUpper === 'UPDATEDBY') return;
    code.push('        if (!string.IsNullOrWhiteSpace(model.' + col.columnName + '))');
    code.push('            query = query.Where(x => x.' + col.columnName + '!.Contains(model.' + col.columnName + '));');
  });

  code.push('');
  // Ordering
  const orderColumns = pkColumns.length > 0 ? pkColumns : [columns[0].columnName];
  orderColumns.forEach((col, idx) => {
    if (idx === 0) {
      code.push('        var data = await query.OrderBy(x => x.' + col + ')');
    } else {
      code.push('            .ThenBy(x => x.' + col + ')');
    }
  });
  code.push('            .ToListAsync();');
  code.push('');
  code.push('        return new ServiceResult<List<' + entityName + '>>(data);');
  code.push('    }');
  code.push('');

  // GetByIdAsync
  code.push('    public async Task<ServiceResult<' + entityName + '>> GetByIdAsync(string id)');
  code.push('    {');
  code.push('        var entity = await _dbContext.Set<' + entityName + '>().FirstOrDefaultAsync(x => x.Id == id);');
  code.push('        if (entity == null)');
  code.push('            return new ServiceResult<' + entityName + '>("ไม่พบข้อมูล");');
  code.push('        return entity;');
  code.push('    }');
  code.push('');

  // AddAsync
  code.push('    public async Task<ServiceResult<' + entityName + '>> AddAsync(' + entityName + 'DTO model)');
  code.push('    {');
  code.push('        var entity = new ' + entityName);
  code.push('        {');
  code.push('            Id = Guid.NewGuid().ToString(),');
  // Add all column assignments (skip Id, CreatedBy, UpdatedBy, CreatedDate)
  columns.forEach(col => {
    const cn = col.columnName.toUpperCase();
    if (cn === 'ID' || cn === 'CREATEDBY' || cn === 'UPDATEDBY' || cn === 'CREATEDDATE' || cn === 'UPDATEDDATE') return;
    code.push('            ' + col.columnName + ' = model.' + col.columnName + ',');
  });
  code.push('            CreatedDate = DateTime.Now,');
  code.push('            CreatedBy = model.CreatedBy');
  code.push('        };');
  code.push('');
  code.push('        _dbContext.Set<' + entityName + '>().Add(entity);');
  code.push('        await _dbContext.SaveChangesAsync();');
  code.push('');
  code.push('        return entity;');
  code.push('    }');
  code.push('');

  // UpdateAsync
  code.push('    public async Task<ServiceResult<' + entityName + '>> UpdateAsync(' + entityName + 'DTO model)');
  code.push('    {');
  code.push('        var entity = await _dbContext.Set<' + entityName + '>().FirstOrDefaultAsync(x => x.Id == model.Id);');
  code.push('        if (entity == null)');
  code.push('            throw new Exception($"ไม่พบข้อมูล Id : {model.Id}");');
  code.push('');
  // Update all column assignments (skip Id, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate)
  columns.forEach(col => {
    const cn = col.columnName.toUpperCase();
    if (cn === 'ID' || cn === 'CREATEDBY' || cn === 'UPDATEDBY' || cn === 'CREATEDDATE' || cn === 'UPDATEDDATE') return;
    code.push('        entity.' + col.columnName + ' = model.' + col.columnName + ';');
  });
  code.push('        entity.UpdatedDate = DateTime.Now;');
  code.push('        entity.UpdatedBy = model.UpdatedBy;');
  code.push('');
  code.push('        await _dbContext.SaveChangesAsync();');
  code.push('');
  code.push('        return entity;');
  code.push('    }');
  code.push('');

  // DeleteAsync
  code.push('    public async Task<ServiceResult<int>> DeleteAsync(' + entityName + 'DTO model)');
  code.push('    {');
  code.push('        var entity = await _dbContext.Set<' + entityName + '>().FirstOrDefaultAsync(x => x.Id == model.Id);');
  code.push('        _dbContext.Set<' + entityName + '>().Remove(entity);');
  code.push('        return await _dbContext.SaveChangesAsync();');
  code.push('    }');
  code.push('');
  code.push('    public async Task<byte[]> ' + exportMethodName + '(List<' + entityName + 'DTO> data)');
  code.push('    {');
  code.push('        var columnHeaders = new Dictionary<string, string>');
  code.push('        {');

  columns.forEach(col => {
    const isSystemField = ['ID', 'CREATEDBY', 'CREATEDDATE', 'UPDATEDBY', 'UPDATEDDATE'].includes(col.columnName.toUpperCase());
    const headerText = col.header.replace(/"/g, '\"');
    if (isSystemField) {
      code.push('            // { "' + col.columnName + '", "' + headerText + '" },');
    } else {
      code.push('            { "' + col.columnName + '", "' + headerText + '" },');
    }
  });

  code.push('        };');
  code.push('');
  code.push('        var ColIndex = new Dictionary<string, int>();');
  code.push('        var workbook = new XSSFWorkbook();');
  code.push('        ISheet sheet = workbook.CreateSheet("Sheet1");');
  code.push('');
  code.push('        // สร้าง style สำหรับหัวข้อ');
  code.push('        IFont boldFont = workbook.CreateFont();');
  code.push('        boldFont.IsBold = true;');
  code.push('');
  code.push('        ICellStyle boldStyle = workbook.CreateCellStyle();');
  code.push('        boldStyle.SetFont(boldFont);');
  code.push('        boldStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;');
  code.push('        boldStyle.FillPattern = FillPattern.SolidForeground;');
  code.push('');
  code.push('        var headerRow = sheet.CreateRow(0);');
  code.push('        int col = 0;');
  code.push('        foreach (var header in columnHeaders)');
  code.push('        {');
  code.push('            var cell = headerRow.CreateCell(col);');
  code.push('            cell.SetCellValue(header.Value);');
  code.push('            cell.CellStyle = boldStyle;');
  code.push('            ColIndex[header.Key] = col;');
  code.push('            col++;');
  code.push('        }');


  code.push('');
  code.push('        int rowIndex = 1;');
  code.push('        foreach (var item in data)');
  code.push('        {');
  code.push('            var rowData = sheet.CreateRow(rowIndex);');
  code.push('');

  columns.forEach(col => {
    const isSystemField = ['ID', 'CREATEDBY', 'CREATEDDATE', 'UPDATEDBY', 'UPDATEDDATE'].includes(col.columnName.toUpperCase());
    if (isSystemField) {
      code.push('            // rowData.CreateCell(ColIndex["' + col.columnName + '"]).SetCellValue(item.' + col.columnName + ');');
    } else {
      code.push('            rowData.CreateCell(ColIndex["' + col.columnName + '"]).SetCellValue(item.' + col.columnName + ');');
    }
  });

  code.push('');
  code.push('            rowIndex++;');
  code.push('        }');
  code.push('');
  code.push('        byte[] excelData;');
  code.push('');
  code.push('        using (var stream = new MemoryStream())');
  code.push('        {');
  code.push('            workbook.Write(stream);');
  code.push('            excelData = stream.ToArray();');
  code.push('        }');
  code.push('');
  code.push('        return excelData;');
  code.push('    }');
  code.push('');
  code.push('    public async Task<byte[]> ExportPdfAsync(List<' + entityName + 'DTO> data)');
  code.push('    {');
  code.push('        using var ms = new MemoryStream();');
  code.push('');
  code.push('        var document = new Document(PageSize.A4, 20, 20, 20, 20);');
  code.push('        PdfWriter.GetInstance(document, ms);');
  code.push('');
  code.push('        document.Open();');
  code.push('');
  code.push('        // Title');
  code.push('        var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);');
  code.push('        document.Add(new Paragraph("' + entityName + ' Report", titleFont));');
  code.push('');
  code.push('        document.Add(new Paragraph(" "));');
  code.push('        document.Add(new Paragraph($"Generate Date : {DateTime.Now:dd/MM/yyyy HH:mm}"));');
  code.push('');
  code.push('        document.Add(new Paragraph(" "));');
  code.push('');
  code.push('        // Column headers for PDF table');
  code.push('        var pdfColumnHeaders = new Dictionary<string, string>');
  code.push('        {');

  columns.forEach(col => {
    const isSystemField = ['ID', 'CREATEDBY', 'CREATEDDATE', 'UPDATEDBY', 'UPDATEDDATE'].includes(col.columnName.toUpperCase());
    const headerText = col.header.replace(/"/g, '\"');
    if (isSystemField) {
      code.push('            // { "' + col.columnName + '", "' + headerText + '" },');
    } else {
      code.push('            { "' + col.columnName + '", "' + headerText + '" },');
    }
  });

  code.push('        };');
  code.push('');
  code.push('        // Create PDF table');
  code.push('        var table = new PdfPTable(pdfColumnHeaders.Count);');
  code.push('        table.WidthPercentage = 100;');
  code.push('');
  code.push('        // Header style');
  code.push('        var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);');
  code.push('        var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);');
  code.push('');
  code.push('        // Render header row');
  code.push('        foreach (var header in pdfColumnHeaders)');
  code.push('        {');
  code.push('            var cell = new PdfPCell(new Phrase(new Chunk(header.Value, headerFont)));');
  code.push('            cell.BackgroundColor = new BaseColor(230, 230, 230);');
  code.push('            cell.HorizontalAlignment = Element.ALIGN_CENTER;');
  code.push('            cell.Padding = 5;');
  code.push('            table.AddCell(cell);');
  code.push('        }');
  code.push('');
  code.push('        // Render data rows');
  code.push('        bool alternate = false;');
  code.push('        foreach (var item in data)');
  code.push('        {');

  columns.forEach(col => {
    const isSystemField = ['ID', 'CREATEDBY', 'CREATEDDATE', 'UPDATEDBY', 'UPDATEDDATE'].includes(col.columnName.toUpperCase());
    if (isSystemField) {
      code.push('            // table.AddCell(new PdfPCell(new Phrase(new Chunk(item.' + col.columnName + '?.ToString() ?? "", dataFont))) { Padding = 3 });');
    } else {
      code.push('            table.AddCell(new PdfPCell(new Phrase(new Chunk(item.' + col.columnName + '?.ToString() ?? "", dataFont)))');
      code.push('            {');
      code.push('                Padding = 3,');
      code.push('                BackgroundColor = alternate ? new BaseColor(245, 245, 245) : BaseColor.White');
      code.push('            });');
    }
  });

  code.push('');
  code.push('            alternate = !alternate;');
  code.push('        }');
  code.push('');
  code.push('        document.Add(table);');
  code.push('');
  code.push('        document.Close();');
  code.push('');
  code.push('        return ms.ToArray();');
  code.push('    }');
  code.push('}');
  code.push('');
  code.push('');

  // ====================================================
  // CONTROLLER
  // ====================================================
  sep('CONTROLLER').forEach(l => code.push(l));

  code.push('using System;');
  code.push('using System.Threading.Tasks;');
  code.push('using Microsoft.AspNetCore.Mvc;');
  code.push('using Microsoft.Extensions.Configuration;');
  code.push('using Microsoft.Extensions.Logging;');
  code.push('using ' + namespace + '.Models;');
  code.push('using ' + namespace + '.DTO;');
  code.push('using ' + namespace + '.Services;');
  code.push('using ' + namespace + '.Repository;');
  code.push('');
  code.push('namespace excisetaxApi.Controllers');
  code.push('{');
  code.push('    /// <summary>');
  code.push('    /// ' + entityName + ' Controller');
  code.push('    /// </summary>');
  code.push('    [Route("api/' + entityName + '/[controller]")]');
  code.push('    [ApiController]');
  code.push('    public class ' + controllerName + 'Controller : AbstractController');
  code.push('    {');
  code.push('        private readonly ILogger<' + controllerName + 'Controller> _logger;');
  code.push('        private readonly I' + entityName + 'Service _' + entityName.toLowerCase() + 'Service;');
  code.push('');
  code.push('        public ' + controllerName + 'Controller(');
  code.push('            ILogger<' + controllerName + 'Controller> logger,');
  code.push('            IConfiguration configuration,');
  code.push('            IUserService userService,');
  code.push('            ISYS_UserRepo userRepo,');
  code.push('            I' + entityName + 'Service ' + entityName.toLowerCase() + 'Service)');
  code.push('            : base(logger, configuration, userService, userRepo)');
  code.push('        {');
  code.push('            _logger = logger;');
  code.push('            _' + entityName.toLowerCase() + 'Service = ' + entityName.toLowerCase() + 'Service;');
  code.push('        }');
  code.push('');
  // Search
  code.push('        [HttpPost("Search")]');
  code.push('        public async Task<IActionResult> Search([FromBody] ' + entityName + 'DTO data)');
  code.push('        {');
  code.push('            var res = new ResponseModel<dynamic>();');
  code.push('            var loginName = HttpContext.User.Identity?.Name;');
  code.push('            try');
  code.push('            {');
  code.push('                var result = await _' + entityName.toLowerCase() + 'Service.SearchAsync(data);');
  code.push('                res.Status = result.IsSuccess ? "OK" : "Error";');
  code.push('                res.Data = result.Data;');
  code.push('                res.Message = result.IsSuccess ? "Data Found" : result.ErrorMessage;');
  code.push('                return result.IsSuccess ? Ok(res) : BadRequest(res);');
  code.push('            }');
  code.push('            catch (Exception e)');
  code.push('            {');
  code.push('                res.Status = "Error";');
  code.push('                res.Message = e.Message;');
  code.push('                return BadRequest(res);');
  code.push('            }');
  code.push('        }');
  code.push('');
  // Get
  code.push('        [HttpGet("Get")]');
  code.push('        public async Task<IActionResult> Get(string id)');
  code.push('        {');
  code.push('            var res = new ResponseModel<dynamic>();');
  code.push('            try');
  code.push('            {');
  code.push('                var result = await _' + entityName.toLowerCase() + 'Service.GetByIdAsync(id);');
  code.push('                res.Status = result.IsSuccess ? "OK" : "Error";');
  code.push('                res.Data = result.Data;');
  code.push('                res.Message = result.IsSuccess ? "Data Found" : "Not found";');
  code.push('                return result.IsSuccess ? Ok(res) : NotFound(res);');
  code.push('            }');
  code.push('            catch (Exception e)');
  code.push('            {');
  code.push('                res.Status = "Error";');
  code.push('                res.Message = e.Message;');
  code.push('                return BadRequest(res);');
  code.push('            }');
  code.push('        }');
  code.push('');
  // Add
  code.push('        [HttpPost("Add")]');
  code.push('        public async Task<IActionResult> Add([FromBody] ' + entityName + 'DTO data)');
  code.push('        {');
  code.push('            var res = new ResponseModel<dynamic>();');
  code.push('            var loginName = HttpContext.User.Identity?.Name;');
  code.push('            try');
  code.push('            {');
  code.push('                data.CreatedBy = loginName;');
  code.push('                var result = await _' + entityName.toLowerCase() + 'Service.AddAsync(data);');
  code.push('                res.Status = result.IsSuccess ? "OK" : "Error";');
  code.push('                res.Data = result.Data;');
  code.push('                res.Message = result.IsSuccess ? "Data Added" : result.ErrorMessage;');
  code.push('                return result.IsSuccess ? Ok(res) : BadRequest(res);');
  code.push('            }');
  code.push('            catch (Exception e)');
  code.push('            {');
  code.push('                AddLog("Error", loginName, e.Message);');
  code.push('                return BadRequest(new { Status = "Error", Message = e.Message });');
  code.push('            }');
  code.push('        }');
  code.push('');
  // Edit (calls UpdateAsync internally)
  code.push('        [HttpPost("Edit")]');
  code.push('        public async Task<IActionResult> Edit([FromBody] ' + entityName + 'DTO data)');
  code.push('        {');
  code.push('            var res = new ResponseModel<dynamic>();');
  code.push('            var loginName = HttpContext.User.Identity?.Name;');
  code.push('            try');
  code.push('            {');
  code.push('                data.UpdatedBy = loginName;');
  code.push('                var result = await _' + entityName.toLowerCase() + 'Service.UpdateAsync(data);');
  code.push('                res.Status = result.IsSuccess ? "OK" : "Error";');
  code.push('                res.Data = result.Data;');
  code.push('                res.Message = result.IsSuccess ? "Data Updated" : result.ErrorMessage;');
  code.push('                return result.IsSuccess ? Ok(res) : BadRequest(res);');
  code.push('            }');
  code.push('            catch (Exception e)');
  code.push('            {');
  code.push('                AddLog("Error", loginName, e.Message);');
  code.push('                return BadRequest(new { Status = "Error", Message = e.Message });');
  code.push('            }');
  code.push('        }');
  code.push('');
  // Del
  code.push('        [HttpPost("Del")]');
  code.push('        public async Task<IActionResult> Del([FromBody] ' + entityName + 'DTO data)');
  code.push('        {');
  code.push('            var res = new ResponseModel<dynamic>();');
  code.push('            var loginName = HttpContext.User.Identity?.Name;');
  code.push('            try');
  code.push('            {');
  code.push('                var result = await _' + entityName.toLowerCase() + 'Service.DeleteAsync(data);');
  code.push('                res.Status = result.IsSuccess ? "OK" : "Error";');
  code.push('                res.Message = result.IsSuccess ? "Data Deleted" : result.ErrorMessage;');
  code.push('                return result.IsSuccess ? Ok(res) : BadRequest(res);');
  code.push('            }');
  code.push('            catch (Exception e)');
  code.push('            {');
  code.push('                AddLog("Error", loginName, e.Message);');
  code.push('                return BadRequest(new { Status = "Error", Message = e.Message });');
  code.push('            }');
  code.push('        }');  code.push('        }');
  code.push('');
  // Export
  code.push('        [HttpPost("ExportExcel")]');
  code.push('        public async Task<IActionResult> ExportExcel([FromBody] List<' + entityName + 'DTO> Data)');
  code.push('        {');
  code.push('            var Res = new ResponseModel<byte[]>();');
  code.push('            var LoginName = HttpContext.User.Identity.Name;');
  code.push('            try');
  code.push('            {');
  code.push('                Res.Status = "OK";');
  code.push('                Res.Data = await _' + entityName.toLowerCase() + 'Service.' + exportMethodName + '(Data);');
  code.push('                Res.Message = "Data Exported";');
  code.push('');
  code.push('                return File(Res.Data, "application/vnd.ms-excel", "ExportedData.xlsx");');
  code.push('            }');
  code.push('            catch (Exception e)');
  code.push('            {');
  code.push('                Res.Status = "Error";');
  code.push('                Res.Message = "Export data error";');
  code.push('                AddLog("Error", LoginName, e.Message);');
  code.push('                return BadRequest(Res);');
  code.push('            }');
  code.push('        }');
  code.push('');
  code.push('        [HttpPost("ExportPdf")]');
  code.push('        public async Task<IActionResult> ExportPdf([FromBody] List<' + entityName + 'DTO> Data)');
  code.push('        {');
  code.push('            var Res = new ResponseModel<byte[]>();');
  code.push('            var LoginName = HttpContext.User.Identity.Name;');
  code.push('            try');
  code.push('            {');
  code.push('                Res.Status = "OK";');
  code.push('                Res.Data = await _' + entityName.toLowerCase() + 'Service.ExportPdfAsync(Data);');
  code.push('                Res.Message = "PDF Exported";');
  code.push('');
  code.push('                return File(Res.Data, "application/pdf", "ExportedData.pdf");');
  code.push('            }');
  code.push('            catch (Exception e)');
  code.push('            {');
  code.push('                Res.Status = "Error";');
  code.push('                Res.Message = "Export PDF error";');
  code.push('                AddLog("Error", LoginName, e.Message);');
  code.push('                return BadRequest(Res);');
  code.push('            }');
  code.push('        }');
  code.push('    }');
  code.push('}');
  code.push('');

  // ====================================================
  // CREATE TABLE SQL
  // ====================================================
  sep('CREATE TABLE SQL SCRIPT').forEach(l => code.push(l));

  code.push('-- ============================================');
  code.push('-- Generated by Excel Report Code Generator');
  code.push('-- Target: SQL Server');
  code.push('-- ============================================');
  code.push('');
  code.push('-- Drop table if exists (for development)');
  code.push('IF OBJECT_ID(N"[' + schemaName + '].[' + tableName + ']", N"U") IS NOT NULL');
  code.push('    DROP TABLE [' + schemaName + '].[' + tableName + '];');
  code.push('GO');
  code.push('');
  code.push('CREATE TABLE [' + schemaName + '].[' + tableName + '] (');
  code.push('');

  const pkSqlList = pkColumns.map(c => '[' + c + ']').join(', ');

  columns.forEach((col, idx) => {
    const sqlType = mapCsTypeToSql(col.csType, col.maxLength, col.decimalPrecision, col.decimalScale);
    const nullable = col.isPk ? 'NOT NULL' : (col.isNullable ? 'NULL' : 'NOT NULL');
    const comma = idx < columns.length - 1 ? ',' : '';
    code.push('    [' + col.columnName + '] ' + sqlType + ' ' + nullable + comma);
  });

  if (pkColumns.length > 0) {
    code.push('');
    code.push('    CONSTRAINT [PK_' + tableName + '] PRIMARY KEY (' + pkSqlList + ')');
  }

  code.push(');');
  code.push('GO');
  code.push('');
  // code.push('-- ============================================');
  // code.push('-- Comments on columns (Excel headers)');
  // code.push('-- ============================================');
  // code.push('EXEC sys.sp_addextendedproperty');
  // code.push('    @name = N"MS_Description",');
  // code.push('    @value = N"Table generated from Excel: ' + escapeSqlString(sourceFileName || 'Report') + '",');
  // code.push('    @level0type = N"SCHEMA", @level0name = "' + schemaName + '",');
  // code.push('    @level1type = N"TABLE",  @level1name = "' + tableName + '";');
  // code.push('GO');
  // code.push('');

  // columns.forEach(col => {
  //   code.push('EXEC sys.sp_addextendedproperty');
  //   code.push('    @name = N"MS_Description",');
  //   code.push('    @value = N"' + escapeSqlString(col.header) + '",');
  //   code.push('    @level0type = N"SCHEMA", @level0name = "' + schemaName + '",');
  //   code.push('    @level1type = N"TABLE",  @level1name = "' + tableName + '",');
  //   code.push('    @level2type = N"COLUMN", @level2name = "' + col.columnName + '";');
  //   code.push('GO');
  //   code.push('');
  // });

  // ====================================================
  // EXCEL IMPORT CODE (Bonus)
  // ====================================================
  code.push('');
  code.push('// ============================================');
  code.push('// EXCEL IMPORT HELPER (using closedxml / EPPlus)');
  code.push('// ============================================');
  code.push('//');
  code.push('// Using EPPlus (dotnet add package EPPlus):');
  code.push('//');
  code.push('// using OfficeOpenXml;');
  code.push('// using System.IO;');
  code.push('//');
  code.push('// public async Task ImportFromExcel(Stream excelStream)');
  code.push('// {');
  code.push('//     using var package = new ExcelPackage(excelStream);');
  code.push('//     var ws = package.Workbook.Worksheets[0];');
  code.push('//     var rowCount = ws.Dimension.Rows;');
  code.push('//     var colCount = ws.Dimension.Columns;');
  code.push('//');
  code.push('//     var entities = new List<' + entityName + '>();');
  code.push('//');
  code.push('//     for (int row = 2; row <= rowCount; row++) // skip header');
  code.push('//     {');
  code.push('//         var entity = new ' + entityName + '();');
  columns.forEach((col, idx) => {
    const cellRef = 'ws.Cells[row, ' + (idx + 1) + ']';
    if (col.csType === 'string') {
      code.push('//         entity.' + col.columnName + ' = ' + cellRef + '.Text ?? string.Empty;');
    } else if (col.csType === 'int') {
      code.push('//         entity.' + col.columnName + ' = ' + cellRef + '.GetValue<int>();');
    } else if (col.csType === 'decimal') {
      code.push('//         entity.' + col.columnName + ' = ' + cellRef + '.GetValue<decimal>();');
    } else if (col.csType === 'double') {
      code.push('//         entity.' + col.columnName + ' = ' + cellRef + '.GetValue<double>();');
    } else if (col.csType === 'bool') {
      code.push('//         entity.' + col.columnName + ' = ' + cellRef + '.GetValue<bool>();');
    } else if (col.csType === 'DateTime') {
      code.push('//         entity.' + col.columnName + ' = ' + cellRef + '.GetValue<DateTime>();');
    }
  });
  code.push('//         entities.Add(entity);');
  code.push('//     }');
  code.push('//');
  code.push('//     _dbContext.Set<' + entityName + '>().AddRange(entities);');
  code.push('//     await _dbContext.SaveChangesAsync();');
  code.push('// }');
  code.push('');


  return code.join('\n');
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Infer table name from file name
 * e.g., "Master_exciseLcEd-EX5602.xlsx" → "Master_exciseLcEd"
 * e.g., "my-report.xlsx"               → "my-report"
 */
function inferTableName(fileName) {
  if (!fileName) return 'ReportEntity';
  let name = fileName
    .replace(/\.xlsx$/i, '')
    .trim();

  // Remove EX code pattern (EX followed by digits)
  name = name.replace(/EX\d+/g, '').trim();
  // Clean up leading/trailing separators left after removing EX code
  name = name.replace(/^[-_\s]+|[-_\s]+$/g, '').trim();

  return name;
}

/**
 * Infer entity name from file name.
 * Supports patterns like:
 *   "Master_Customer-EX5010.xlsx" → "MasterCustomer"
 *   "EX5010_Master.xlsx"         → "Master"
 *   "my-report.xlsx"             → "MyReport"
 */
function inferEntityName(fileName) {
  if (!fileName) return 'ReportEntity';

  let base = fileName.replace(/\.xlsx$/i, '').trim();

  // If it contains underscore, extract EX code and build entity name
  if (base.includes('_')) {
    // Extract EX code pattern (EX followed by digits) from the entire string
    const exMatch = base.match(/EX\d+/);

    // Remove the EX code to get the entity name parts
    let modBase = base;
    if (exMatch) {
      modBase = modBase.replace(exMatch[0], '').trim();
    }

    // Clean up and build PascalCase entity name
    return modBase
      .replace(/[^a-zA-Z0-9]/g, ' ')
      .split(/\s+/)
      .filter(Boolean)
      .map(w => w.charAt(0).toUpperCase() + w.slice(1))
      .join('') || 'ReportEntity';
  }

  // Otherwise, convert to PascalCase
  return base
    .replace(/[^a-zA-Z0-9]/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join('') || 'ReportEntity';
}

/**
 * Convert header text to PascalCase (C# property name)
 */
function toPascalCase(str) {
  if (!str) return 'Empty';
  return str
    .replace(/[^a-zA-Z0-9\u0E00-\u0E7F ]/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join('')
    .replace(/^[0-9]/, m => '_' + m) || 'Field';
}

/**
 * Infer C# type from Excel cell
 */
function inferCsType(cell) {
  if (!cell) return 'string';

  const numFmt = (cell.numFmt || '').toLowerCase().trim();
  const isDateFormat = /(yy|yyyy|dd|mmm)/i.test(numFmt);
  const hasTime = /h/.test(numFmt) || /m.*s/.test(numFmt) || /s$/.test(numFmt);

  if (isDateFormat || hasTime) {
    return 'DateTime';
  }

  if (cell.value != null && typeof cell.value === 'number') {
    if (Number.isInteger(cell.value)) {
      if (numFmt.includes('%')) return 'decimal';
      return 'int';
    }
    return 'decimal';
  }

  if (cell.value != null && typeof cell.value === 'boolean') {
    return 'bool';
  }

  if (numFmt.includes('%') || numFmt.includes('0.00') || numFmt.includes('#,##0')) {
    return 'decimal';
  }
  if (numFmt === '0' || numFmt === '#,##0') {
    return 'int';
  }

  return 'string';
}

/**
 * Infer C# type from column header name (fallback when Excel cell format is 'string').
 * Similar to inferNgPipe in export-primeng.js — uses name conventions.
 */
function inferCsTypeFromName(headerName) {
  if (!headerName) return 'string';
  const name = headerName.toLowerCase();

  // DateTime fields
  if (name.includes('date') || name.includes('time') || name === 'created' || name === 'updated') {
    return 'DateTime';
  }

  // Numeric fields (decimal/float)
  if (name.includes('tax') || name.includes('total') || name.includes('amount') ||
      name.includes('sum') || name.includes('price') || name.includes('cost') ||
      name.includes('fee') || name.includes('rate') || name.includes('qty') ||
      name.includes('quantity') || name.includes('balance') || name.includes('net') ||
      name.includes('gross') || name.includes('discount') || name.includes('vat') ||
      name.includes('percent') || name.includes('count') || name.includes('number')) {
    return 'decimal';
  }

  // Integer fields
  if (name.includes('year') || name.includes('rank') || name === 'no' || name.endsWith('id') || name === 'id') {
    return 'int';
  }

  // Boolean fields
  if (name.startsWith('is') || name.startsWith('has') || name.startsWith('flag') ||
      name === 'active' || name.includes('enable') || name.includes('status')) {
    return 'bool';
  }

  return 'string';
}

/**
 * Map C# type to SQL Server type
 */
function mapCsTypeToSql(csType, maxLength, decimalPrecision, decimalScale) {
  switch (csType) {
    case 'int': return 'INT';
    case 'decimal':
      if (decimalPrecision !== null && decimalScale !== null) {
        return 'DECIMAL(' + decimalPrecision + ',' + decimalScale + ')';
      }
      return 'DECIMAL(18,2)';
    case 'double':
    case 'float': return 'FLOAT';
    case 'bool': return 'BIT';
    case 'DateTime': return 'DATETIME2(7)';
    case 'string':
    default:
      if (maxLength < 0) return 'NVARCHAR(MAX)';
      if (maxLength > 4000) return 'NVARCHAR(MAX)';
      return 'NVARCHAR(' + (maxLength > 0 ? maxLength : 255) + ')';
  }
}

/**
 * Escape string for SQL N-prefixed strings (doubles single quotes)
 */
function escapeSqlString(str) {
  if (!str) return '';
  return String(str).replace(/'/g, "''");
}

/**
 * Escape string for use in C# XML comments
 */
function escapeXmlForCs(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Parse cell note text to extract PK flag, max length, and decimal precision/scale.
 *
 * Supported note formats:
 *   "pk"              → isPk=true,  maxLength=null (use default)
 *   "pk size(100)"    → isPk=true,  maxLength=100
 *   "pk:100"          → isPk=true,  maxLength=100
 *   "pk,100"          → isPk=true,  maxLength=100
 *   "pk max"          → isPk=true,  maxLength=-1  (NVARCHAR(MAX))
 *   "size(500)"       → isPk=false, maxLength=500
 *   "length(200)"     → isPk=false, maxLength=200
 *   "500"             → isPk=false, maxLength=500
 *   "max"             → isPk=false, maxLength=-1  (NVARCHAR(MAX))
 *   "size(18,6)"      → decimalPrecision=18, decimalScale=6
 *   "decimal(18,6)"   → decimalPrecision=18, decimalScale=6
 *   "dec(18,6)"       → decimalPrecision=18, decimalScale=6
 *
 * @param {string} noteText - The raw note text from the cell.
 * @returns {{ isPk: boolean, maxLength: number|null, decimalPrecision: number|null, decimalScale: number|null }}
 */
function parseCellNote(noteText) {
  const text = (noteText || '').trim();
  if (!text) return { isPk: false, maxLength: null, decimalPrecision: null, decimalScale: null };

  // Check for PK flag (case-insensitive, whole word or part of "pk")
  const isPk = /\bpk\b/i.test(text);

  let maxLength = null;
  let decimalPrecision = null;
  let decimalScale = null;

  // Check for decimal precision/scale with comma: decimal(18,6), dec(18,6), size(18,6)
  // MUST be checked BEFORE single-number patterns to avoid mis-matching "size(18" from "size(18,6)"
  const decMatch = text.match(/\b(?:decimal|dec|size)\s*\(?\s*(\d+)\s*,\s*(\d+)\s*\)?\b/i);
  if (decMatch) {
    decimalPrecision = parseInt(decMatch[1], 10);
    decimalScale = parseInt(decMatch[2], 10);
  }

  // Check for "max" keyword (NVARCHAR(MAX) → sentinel -1)
  if (/\bmax\b/i.test(text) && !/\bmaxlength\b/i.test(text)) {
    maxLength = -1;
  }

  // Check for size(N) or length(N) or maxLength(N) patterns (single number only)
  // Only run if we didn't already parse a decimal pattern (which could include "size(N,N)")
  if (decimalPrecision === null) {
    const sizeMatch = text.match(/\b(?:size|length|maxlength)\s*[:\(]\s*(\d+)\s*[\)]?\b/i);
    if (sizeMatch) {
      maxLength = parseInt(sizeMatch[1], 10);
    }
  }

  // If no keyword match, try to find a standalone number (not part of "pk")
  if (maxLength === null && decimalPrecision === null) {
    const cleaned = text.replace(/\bpk\b/gi, '').trim();
    const numMatch = cleaned.match(/^\s*(\d+)\s*$/);
    if (numMatch) {
      maxLength = parseInt(numMatch[1], 10);
    }
  }

  return { isPk, maxLength, decimalPrecision, decimalScale };
}

/**
 * Extract note/comment text from a cell reliably.
 * (Same as getCellNoteText in utils-cal.js, included here for standalone use)
 */
function getCellNoteText(cell) {
  if (!cell) return '';

  // Internal _comment (ExcelJS v6+)
  if (cell._comment?.note?.texts) {
    return cell._comment.note.texts.map(x => x.text || '').join('').trim();
  }

  // Standard note (cloned worksheet structure)
  if (cell.note) {
    if (cell.note.texts) {
      return cell.note.texts.map(x => x.text || '').join('').trim();
    }
    if (cell.note.note?.texts) {
      return cell.note.note.texts.map(x => x.text || '').join('').trim();
    }
    if (typeof cell.note === 'string') {
      return cell.note.trim();
    }
  }

  // comment property (ExcelJS fallback)
  if (cell.comment?.texts) {
    return cell.comment.texts.map(x => x.text || '').join('').trim();
  }
  if (cell.comment?.note?.texts) {
    return cell.comment.note.texts.map(x => x.text || '').join('').trim();
  }

  // model?.note (internal ExcelJS)
  if (cell.model?.note?.texts) {
    return cell.model.note.texts.map(x => x.text || '').join('').trim();
  }

  return '';
}
