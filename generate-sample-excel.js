// ============================================
// Sample Excel Generator for Group Feature Testing
// ============================================
// Run: node generate-sample-excel.js
// Output: sample-group-report.xlsx
// ============================================

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function generateSampleFiles() {
  await generateGroupReport();
  await generateMultiGroupReport();
  await generateSimpleReport();    await generateNestedGroupReport();
    console.log('\n✅ All sample Excel files created successfully!');
}

// ============================================
// Sample 1: Single Group Report
// ============================================
async function generateGroupReport() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel Report Code Generator';
  const ws = wb.addWorksheet('GroupReport');

  // Column widths
  ws.getColumn(1).width = 22;   // A: Section marker
  ws.getColumn(2).width = 18;   // B: รหัสสินค้า
  ws.getColumn(3).width = 35;   // C: ชื่อสินค้า
  ws.getColumn(4).width = 15;   // D: หมวดหมู่
  ws.getColumn(5).width = 15;   // E: ราคา
  ws.getColumn(6).width = 15;   // F: จำนวน

  // Default row height
  ws.properties.defaultRowHeight = 20;

  // ---- Row 1: TITLE ----
  ws.getCell('A1').value = 'title';
  ws.getCell('B1').value = 'รายงานสินค้าคงคลัง';
  ws.getCell('B1').font = { name: 'Sarabun', size: 18, bold: true };
  ws.mergeCells('B1:F1');
  ws.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 40;

  // ---- Row 2-3: COLUMNHEADER ----
  ws.getCell('A2').value = 'columnheader';
  ws.getCell('B2').value = 'รหัสสินค้า';
  ws.getCell('C2').value = 'ชื่อสินค้า';
  ws.getCell('D2').value = 'หมวดหมู่';
  ws.getCell('E2').value = 'ราคา';
  ws.getCell('F2').value = 'จำนวน';
  ws.getRow(2).height = 22;

  // Sub-header
  ws.getCell('A3').value = '';  // continuation of columnheader
  ws.getCell('B3').value = '(Code)';
  ws.getCell('B3').font = { italic: true, size: 9, color: { argb: 'FF666666' } };
  ws.getCell('C3').value = '(Product Name)';
  ws.getCell('C3').font = { italic: true, size: 9, color: { argb: 'FF666666' } };
  ws.getCell('D3').value = '(Category)';
  ws.getCell('D3').font = { italic: true, size: 9, color: { argb: 'FF666666' } };
  ws.getCell('E3').value = '(Price)';
  ws.getCell('E3').font = { italic: true, size: 9, color: { argb: 'FF666666' } };
  ws.getCell('F3').value = '(Qty)';
  ws.getCell('F3').font = { italic: true, size: 9, color: { argb: 'FF666666' } };
  ws.getRow(3).height = 18;

  // ---- Row 4: GROUP DEFINITION ----
  ws.getCell('A4').value = 'group:หมวดหมู่';
  ws.getCell('A4').font = { bold: true, color: { argb: 'FF991B1B' } };

  // ---- Row 5: GROUPHEADER ----
  ws.getCell('A5').value = 'groupheader';
  ws.getCell('B5').value = 'หมวดหมู่:';
  ws.getCell('C5').value = '{{หมวดหมู่}}';
  ws.getCell('C5').font = { bold: true, color: { argb: 'FF1E40AF' } };
  ws.mergeCells('C5:F5');
  ws.getRow(5).height = 25;

  // ---- Row 6-9: DETAIL ----
  ws.getCell('A6').value = 'detail';
  ws.getCell('B6').value = '{{รหัสสินค้า}}';
  ws.getCell('C6').value = '{{ชื่อสินค้า}}';
  ws.getCell('D6').value = '{{หมวดหมู่}}';
  ws.getCell('E6').value = '{{ราคา}}';
  ws.getCell('F6').value = '{{จำนวน}}';
  ws.getRow(6).height = 22;

  ws.getCell('A7').value = '';  // continuation of detail
  ws.getCell('B7').value = 'P001';
  ws.getCell('C7').value = 'โน้ตบุ๊ก Dell Inspiron';
  ws.getCell('D7').value = 'คอมพิวเตอร์';
  ws.getCell('E7').value = 25000;
  ws.getCell('E7').numFmt = '#,##0.00';
  ws.getCell('F7').value = 10;
  ws.getRow(7).height = 20;

  ws.getCell('A8').value = '';  // continuation of detail
  ws.getCell('B8').value = 'P002';
  ws.getCell('C8').value = 'เมาส์ไร้สาย Logitech';
  ws.getCell('D8').value = 'อุปกรณ์เสริม';
  ws.getCell('E8').value = 890;
  ws.getCell('E8').numFmt = '#,##0.00';
  ws.getCell('F8').value = 50;
  ws.getRow(8).height = 20;

  ws.getCell('A9').value = '';  // continuation of detail
  ws.getCell('B9').value = 'P003';
  ws.getCell('C9').value = 'คีย์บอร์ด Mechanical';
  ws.getCell('D9').value = 'อุปกรณ์เสริม';
  ws.getCell('E9').value = 3500;
  ws.getCell('E9').numFmt = '#,##0.00';
  ws.getCell('F9').value = 25;
  ws.getRow(9).height = 20;

  // ---- Row 10: GROUPFOOTER ----
  ws.getCell('A10').value = 'groupfooter';
  ws.getCell('B10').value = 'รวมจำนวน:';
  ws.getCell('B10').font = { bold: true };
  ws.getCell('F10').value = 85;
  ws.getCell('F10').font = { bold: true };
  ws.getRow(10).height = 22;

  // ---- Row 11: SUMMARY ----
  ws.getCell('A11').value = 'summary';
  ws.getCell('B11').value = 'รวมทั้งหมด:';
  ws.getCell('B11').font = { bold: true, size: 14 };
  ws.getCell('E11').value = 29390;
  ws.getCell('E11').numFmt = '#,##0.00';
  ws.getCell('E11').font = { bold: true, size: 14 };
  ws.getCell('F11').value = 85;
  ws.getCell('F11').font = { bold: true, size: 14 };
  ws.getRow(11).height = 30;

  // Add borders to header cells
  const headerCells = ['B2', 'C2', 'D2', 'E2', 'F2'];
  headerCells.forEach(ref => {
    const cell = ws.getCell(ref);
    cell.border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.font = { bold: true };
  });

  const outPath = path.join(__dirname, 'sample-group-report.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`✅ Created: ${path.basename(outPath)}`);
}

// ============================================
// Sample 2: Multiple Groups Report
// ============================================
async function generateMultiGroupReport() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel Report Code Generator';
  const ws = wb.addWorksheet('MultiGroup');

  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 15;
  ws.getColumn(3).width = 30;
  ws.getColumn(4).width = 15;
  ws.getColumn(5).width = 15;
  ws.properties.defaultRowHeight = 20;

  // Row 1: Title
  ws.getCell('A1').value = 'title';
  ws.getCell('B1').value = 'รายงานยอดขายแยกตามพนักงาน';
  ws.getCell('B1').font = { name: 'Sarabun', size: 16, bold: true };
  ws.mergeCells('B1:E1');
  ws.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 35;

  // Row 2: ColumnHeader
  ws.getCell('A2').value = 'columnheader';
  ws.getCell('B2').value = 'รหัส';
  ws.getCell('C2').value = 'ชื่อ-นามสกุล';
  ws.getCell('D2').value = 'ยอดขาย';
  ws.getCell('E2').value = 'ค่าคอมมิชชั่น';
  ws.getRow(2).height = 22;

  ['B2','C2','D2','E2'].forEach(ref => {
    ws.getCell(ref).font = { bold: true };
    ws.getCell(ref).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(ref).border = {
      top: { style: 'thin' }, bottom: { style: 'thin' },
      left: { style: 'thin' }, right: { style: 'thin' }
    };
  });

  // Row 3: group:แผนก
  ws.getCell('A3').value = 'group:แผนก';
  ws.getCell('A3').font = { bold: true, color: { argb: 'FF991B1B' } };
  ws.getCell('B3').value = '(group:แผนก)';
  ws.getCell('B3').font = { italic: true, size: 9, color: { argb: 'FF999999' } };

  // Row 4: GroupHeader
  ws.getCell('A4').value = 'groupheader';
  ws.getCell('B4').value = 'แผนก:';
  ws.getCell('C4').value = '{{แผนก}}';
  ws.getCell('C4').font = { bold: true, color: { argb: 'FF1E40AF' } };
  ws.mergeCells('C4:E4');
  ws.getRow(4).height = 24;

  // Row 5-7: Detail (first group)
  ws.getCell('A5').value = 'detail';
  ws.getCell('B5').value = '{{รหัสพนักงาน}}';
  ws.getCell('C5').value = '{{ชื่อพนักงาน}}';
  ws.getCell('D5').value = '{{ยอดขาย}}';
  ws.getCell('E5').value = '{{ค่าคอมมิชชั่น}}';
  ws.getRow(5).height = 22;

  ws.getCell('A6').value = '';
  ws.getCell('B6').value = 'EMP001';
  ws.getCell('C6').value = 'สมชาย ใจดี';
  ws.getCell('D6').value = 150000;
  ws.getCell('D6').numFmt = '#,##0.00';
  ws.getCell('E6').value = 7500;
  ws.getCell('E6').numFmt = '#,##0.00';

  ws.getCell('A7').value = '';
  ws.getCell('B7').value = 'EMP002';
  ws.getCell('C7').value = 'สมหญิง รักดี';
  ws.getCell('D7').value = 200000;
  ws.getCell('D7').numFmt = '#,##0.00';
  ws.getCell('E7').value = 10000;
  ws.getCell('E7').numFmt = '#,##0.00';

  // Row 8: GroupFooter (first)
  ws.getCell('A8').value = 'groupfooter';
  ws.getCell('B8').value = 'รวมแผนก';
  ws.getCell('B8').font = { bold: true };
  ws.getCell('D8').value = 350000;
  ws.getCell('D8').numFmt = '#,##0.00';
  ws.getCell('D8').font = { bold: true };
  ws.getCell('E8').value = 17500;
  ws.getCell('E8').numFmt = '#,##0.00';
  ws.getCell('E8').font = { bold: true };
  ws.getRow(8).height = 22;

  // Row 9: group:แผนก2 (second group)
  ws.getCell('A9').value = 'group:แผนก2';
  ws.getCell('A9').font = { bold: true, color: { argb: 'FF991B1B' } };
  ws.getCell('B9').value = '(group:แผนก2)';
  ws.getCell('B9').font = { italic: true, size: 9, color: { argb: 'FF999999' } };

  // Row 10: GroupHeader (second)
  ws.getCell('A10').value = 'groupheader';
  ws.getCell('B10').value = 'แผนก:';
  ws.getCell('C10').value = '{{แผนก}}';
  ws.getCell('C10').font = { bold: true, color: { argb: 'FF1E40AF' } };
  ws.mergeCells('C10:E10');
  ws.getRow(10).height = 24;

  // Row 11-12: Detail (second group)
  ws.getCell('A11').value = 'detail';
  ws.getCell('B11').value = '{{รหัสพนักงาน}}';
  ws.getCell('C11').value = '{{ชื่อพนักงาน}}';
  ws.getCell('D11').value = '{{ยอดขาย}}';
  ws.getCell('E11').value = '{{ค่าคอมมิชชั่น}}';
  ws.getRow(11).height = 22;

  ws.getCell('A12').value = '';
  ws.getCell('B12').value = 'EMP003';
  ws.getCell('C12').value = 'มานพ ขยันดี';
  ws.getCell('D12').value = 300000;
  ws.getCell('D12').numFmt = '#,##0.00';
  ws.getCell('E12').value = 15000;
  ws.getCell('E12').numFmt = '#,##0.00';

  // Row 13: GroupFooter (second)
  ws.getCell('A13').value = 'groupfooter';
  ws.getCell('B13').value = 'รวมแผนก 2';
  ws.getCell('B13').font = { bold: true };
  ws.getCell('D13').value = 300000;
  ws.getCell('D13').numFmt = '#,##0.00';
  ws.getCell('D13').font = { bold: true };
  ws.getCell('E13').value = 15000;
  ws.getCell('E13').numFmt = '#,##0.00';
  ws.getCell('E13').font = { bold: true };
  ws.getRow(13).height = 22;

  // Row 14: Summary
  ws.getCell('A14').value = 'summary';
  ws.getCell('B14').value = 'รวมทั้งสิ้น';
  ws.getCell('B14').font = { bold: true, size: 14 };
  ws.getCell('D14').value = 650000;
  ws.getCell('D14').numFmt = '#,##0.00';
  ws.getCell('D14').font = { bold: true, size: 14 };
  ws.getCell('E14').value = 32500;
  ws.getCell('E14').numFmt = '#,##0.00';
  ws.getCell('E14').font = { bold: true, size: 14 };
  ws.getRow(14).height = 30;

  const outPath = path.join(__dirname, 'sample-multi-group-report.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`✅ Created: ${path.basename(outPath)}`);
}

// ============================================
// Sample 3: Simple Report (no groups, for comparison)
// ============================================
async function generateSimpleReport() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel Report Code Generator';
  const ws = wb.addWorksheet('SimpleReport');

  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 18;
  ws.getColumn(3).width = 30;
  ws.getColumn(4).width = 15;
  ws.properties.defaultRowHeight = 20;

  // Row 1: Title
  ws.getCell('A1').value = 'title';
  ws.getCell('B1').value = 'รายงานพนักงาน';
  ws.getCell('B1').font = { size: 16, bold: true };
  ws.mergeCells('B1:D1');
  ws.getCell('B1').alignment = { horizontal: 'center' };
  ws.getRow(1).height = 35;

  // Row 2: ColumnHeader
  ws.getCell('A2').value = 'columnheader';
  ws.getCell('B2').value = 'รหัส';
  ws.getCell('C2').value = 'ชื่อ';
  ws.getCell('D2').value = 'เงินเดือน';
  ws.getRow(2).height = 22;
  ['B2','C2','D2'].forEach(ref => {
    ws.getCell(ref).font = { bold: true };
    ws.getCell(ref).alignment = { horizontal: 'center' };
  });

  // Row 3-4: Detail
  ws.getCell('A3').value = 'detail';
  ws.getCell('B3').value = '{{รหัส}}';
  ws.getCell('C3').value = '{{ชื่อ}}';
  ws.getCell('D3').value = '{{เงินเดือน}}';
  ws.getRow(3).height = 22;

  ws.getCell('A4').value = '';
  ws.getCell('B4').value = 'EMP001';
  ws.getCell('C4').value = 'สมชาย ใจดี';
  ws.getCell('D4').value = 45000;
  ws.getCell('D4').numFmt = '#,##0';

  // Row 5: Summary
  ws.getCell('A5').value = 'summary';
  ws.getCell('B5').value = 'รวม';
  ws.getCell('D5').value = 45000;
  ws.getCell('D5').numFmt = '#,##0';
  ws.getRow(5).height = 25;

  const outPath = path.join(__dirname, 'sample-simple-report.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`✅ Created: ${path.basename(outPath)}`);
}

// ============================================
// Sample 4: Nested Groups Report
// ============================================
// แสดงตัวอย่าง Group ซ้อนกัน 2 ระดับ:
//   ภูมิภาค (outer group) → แผนก (inner group)
// ============================================
async function generateNestedGroupReport() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Excel Report Code Generator';
    const ws = wb.addWorksheet('NestedGroup');

    // Column setup
    ws.getColumn(1).width = 22;   // A: Section marker
    ws.getColumn(2).width = 14;   // B: รหัส
    ws.getColumn(3).width = 28;   // C: ชื่อพนักงาน
    ws.getColumn(4).width = 16;   // D: ภูมิภาค
    ws.getColumn(5).width = 18;   // E: แผนก
    ws.getColumn(6).width = 18;   // F: ยอดขาย
    ws.getColumn(7).width = 22;   // G: หมายเหตุ (extra column)
    ws.properties.defaultRowHeight = 20;

    // ============================
    // Row 1: TITLE
    // ============================
    ws.getCell('A1').value = 'title';
    ws.getCell('B1').value = 'รายงานยอดขายแยกตามภูมิภาคและแผนก';
    ws.getCell('B1').font = { name: 'Sarabun', size: 18, bold: true };
    ws.mergeCells('B1:F1');
    ws.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 45;

    // ============================
    // Row 2: COLUMNHEADER
    // ============================
    ws.getCell('A2').value = 'columnheader';
    ws.getCell('B2').value = 'รหัส';
    ws.getCell('C2').value = 'ชื่อพนักงาน';
    ws.getCell('D2').value = 'ภูมิภาค';
    ws.getCell('E2').value = 'แผนก';
    ws.getCell('F2').value = 'ยอดขาย';
    ws.getCell('G2').value = 'หมายเหตุ';

    ['B2','C2','D2','E2','F2','G2'].forEach(ref => {
        const cell = ws.getCell(ref);
        cell.font = { bold: true, size: 11 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    ws.getRow(2).height = 24;

    // ============================
    // Row 3: OUTER GROUP — ภูมิภาค
    // ============================
    ws.getCell('A3').value = 'group:ภูมิภาค';
    ws.getCell('A3').font = { bold: true, color: { argb: 'FF991B1B' } };

    // ============================
    // Row 4: OUTER GROUP HEADER
    // ============================
    ws.getCell('A4').value = 'groupheader';
    ws.getCell('B4').value = 'ภูมิภาค:';
    ws.getCell('B4').font = { bold: true, size: 12 };
    ws.getCell('C4').value = '{{ภูมิภาค}}';
    ws.getCell('C4').font = { bold: true, size: 12, color: { argb: 'FF1E40AF' } };
    ws.mergeCells('C4:G4');
    ws.getRow(4).height = 28;

    // ============================
    // Row 5: INNER GROUP — แผนก
    // ============================
    ws.getCell('A5').value = 'group:แผนก';
    ws.getCell('A5').font = { bold: true, color: { argb: 'FF991B1B' } };
    ws.getCell('B5').value = '(nested under ภูมิภาค)';
    ws.getCell('B5').font = { italic: true, size: 9, color: { argb: 'FF999999' } };

    // ============================
    // Row 6: INNER GROUP HEADER
    // ============================
    ws.getCell('A6').value = 'groupheader';
    ws.getCell('B6').value = 'แผนก:';
    ws.getCell('B6').font = { bold: true, size: 11 };
    ws.getCell('C6').value = '{{แผนก}}';
    ws.getCell('C6').font = { bold: true, size: 11, color: { argb: 'FF1E40AF' } };
    ws.mergeCells('C6:G6');
    ws.getRow(6).height = 26;

    // ============================
    // Row 7: DETAIL (template)
    // ============================
    ws.getCell('A7').value = 'detail';
    ws.getCell('B7').value = '{{รหัส}}';
    ws.getCell('C7').value = '{{ชื่อพนักงาน}}';
    ws.getCell('D7').value = '{{ภูมิภาค}}';
    ws.getCell('E7').value = '{{แผนก}}';
    ws.getCell('F7').value = '{{ยอดขาย}}';
    ws.getCell('F7').numFmt = '#,##0.00';
    ws.getRow(7).height = 22;

    // ============================
    // Rows 8-15: DETAIL (actual data)
    // ============================
    const data = [
        ['EMP001', 'สมชาย ใจดี',     'ภาคเหนือ', 'ฝ่ายขาย',     50000],
        ['EMP002', 'สมหญิง รักดี',   'ภาคเหนือ', 'ฝ่ายขาย',     65000],
        ['EMP003', 'มานพ ขยันดี',    'ภาคเหนือ', 'ฝ่ายบริการ',   45000],
        ['EMP004', 'ดารา สวยงาม',    'ภาคเหนือ', 'ฝ่ายบริการ',   38000],
        ['EMP005', 'วิชัย มั่งมี',     'ภาคใต้',   'ฝ่ายขาย',     72000],
        ['EMP006', 'อาภา รุ่งเรือง',   'ภาคใต้',   'ฝ่ายขาย',     58000],
        ['EMP007', 'ประทีป แก้วดี',   'ภาคใต้',   'ฝ่ายบริการ',   42000],
        ['EMP008', 'สุดา ใจบุญ',      'ภาคใต้',   'ฝ่ายบริการ',   35000],
        ['EMP009', 'ก้อง เก่งกาจ',    'ภาคกลาง',  'ฝ่ายขาย',     88000],
        ['EMP010', 'แพรว พราวแสง',   'ภาคกลาง',  'ฝ่ายบริการ',   32000],
    ];

    data.forEach((row, idx) => {
        const r = 8 + idx;
        ws.getCell(`A${r}`).value = '';  // continuation of detail
        ws.getCell(`B${r}`).value = row[0];
        ws.getCell(`C${r}`).value = row[1];
        ws.getCell(`D${r}`).value = row[2];
        ws.getCell(`E${r}`).value = row[3];
        ws.getCell(`F${r}`).value = row[4];
        ws.getCell(`F${r}`).numFmt = '#,##0.00';
        ws.getRow(r).height = 20;

        // Add subtle borders to data cells
        ['B','C','D','E','F'].forEach(col => {
            ws.getCell(`${col}${r}`).border = {
                left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                right: { style: 'thin', color: { argb: 'FFE0E0E0' } },
            };
        });
    });

    // ============================
    // Row 18: INNER GROUP FOOTER (แผนก)
    // ============================
    ws.getCell('A18').value = 'groupfooter';
    ws.getCell('B18').value = 'รวมยอดแผนก';
    ws.getCell('B18').font = { bold: true, italic: true, color: { argb: 'FF7C3AED' } };
    ws.getCell('E18').value = '{{แผนก}}';
    ws.getCell('E18').font = { bold: true, italic: true, color: { argb: 'FF1E40AF' } };
    ws.getCell('F18').value = '{{sum_ยอดขายแผนก}}';
    ws.getCell('F18').font = { bold: true };
    ws.getCell('F18').numFmt = '#,##0.00';
    ws.mergeCells('C18:D18');
    ws.getRow(18).height = 24;

    // ============================
    // Row 19: OUTER GROUP FOOTER (ภูมิภาค)
    // ============================
    ws.getCell('A19').value = 'groupfooter';
    ws.getCell('B19').value = 'รวมยอดภูมิภาค';
    ws.getCell('B19').font = { bold: true, color: { argb: 'FFDC2626' } };
    ws.getCell('E19').value = '{{ภูมิภาค}}';
    ws.getCell('E19').font = { bold: true, color: { argb: 'FF1E40AF' } };
    ws.getCell('F19').value = '{{sum_ยอดขายภูมิภาค}}';
    ws.getCell('F19').font = { bold: true };
    ws.getCell('F19').numFmt = '#,##0.00';
    ws.mergeCells('C19:D19');
    ws.getRow(19).height = 24;

    // ============================
    // Row 20: SUMMARY
    // ============================
    ws.getCell('A20').value = 'summary';
    ws.getCell('B20').value = 'รวมทั้งสิ้น';
    ws.getCell('B20').font = { bold: true, size: 14 };
    ws.getCell('F20').value = '{{sum_ยอดขายรวม}}';
    ws.getCell('F20').font = { bold: true, size: 14 };
    ws.getCell('F20').numFmt = '#,##0.00';
    ws.mergeCells('B20:E20');
    ws.getRow(20).height = 32;

    // ============================
    // Add visual annotations for clarity
    // ============================
    // Section labels in column A with color coding
    const sectionColors = {
        'title': 'FF3B82F6',
        'columnheader': 'FF10B981',
        'group:ภูมิภาค': 'FF991B1B',
        'groupheader': 'FFF59E0B',
        'group:แผนก': 'FF991B1B',
        'detail': 'FF6B7280',
        'groupfooter': 'FFF59E0B',
        'summary': 'FF3B82F6',
    };

    for (let r = 1; r <= 20; r++) {
        const cellA = ws.getCell(`A${r}`);
        const val = String(cellA.value || '').trim();
        if (val) {
            const baseColor = sectionColors[val] || 'FF6B7280';
            cellA.font = {
                color: { argb: baseColor },
                bold: true,
                size: 9
            };
            cellA.alignment = { horizontal: 'right', vertical: 'middle' };
        }
    }

    const outPath = path.join(__dirname, 'sample-nested-group-report.xlsx');
    await wb.xlsx.writeFile(outPath);
    console.log(`✅ Created: ${path.basename(outPath)}`);
}

generateSampleFiles().catch(err => {
  console.error('❌ Error generating samples:', err.message);
  process.exit(1);
});
