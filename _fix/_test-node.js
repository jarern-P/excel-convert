// ============================================
// Multi-line detail functional test (Node)
// Loads the app's REAL browser scripts into a vm
// sandbox, replicates doExport()'s input building,
// runs all 5 patched exports on sample-group-report.xlsx
// (detail section = rows 6-9 => 4 template rows/record)
// and asserts multi-line markers.
// ============================================
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ExcelJS = require('exceljs');

const ROOT = path.join(__dirname, '..');

// ---- minimal DOM stubs (enough for app.js/utils load + escapeHtml) ----
function makeEl() {
  return {
    style: {},
    classList: { add() {}, remove() {} },
    setAttribute() {},
    appendChild() {},
    remove() {},
    click() {},
    addEventListener() {},
    textContent: '',
    innerHTML: '',
    value: '',
  };
}

const documentStub = {
  addEventListener() {},
  getElementById() { return makeEl(); },
  createElement() {
    const el = makeEl();
    Object.defineProperty(el, 'innerHTML', {
      get() { return escapeHtml(this.textContent); },
      set(v) { this.textContent = String(v).replace(/<[^>]*>/g, ''); },
    });
    return el;
  },
  head: makeEl(),
  body: makeEl(),
  createTextNode() { return makeEl(); },
};

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const sandbox = {
  console,
  setTimeout,
  clearTimeout,
  Buffer,
  Blob: typeof Blob !== 'undefined' ? Blob : undefined,
  URL: typeof URL !== 'undefined' ? URL : undefined,
  atob: typeof atob === 'function' ? atob : (s) => Buffer.from(s, 'base64').toString('binary'),
  btoa: typeof btoa === 'function' ? btoa : (s) => Buffer.from(s, 'binary').toString('base64'),
  ExcelJS,
  document: documentStub,
  navigator: {},
  location: { href: 'http://localhost/test' },
  FileReader: function () {},
  requestAnimationFrame: (cb) => setTimeout(cb, 0),
  DOMParser: function () {
    // Minimal XML parser stub: extracts a:clrScheme colors from theme XML
    // so getThemeColors() works against the real workbook theme.
    const cache = {};
    function parseTheme(xml) {
      const tags = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6'];
      const schemeBody = (xml.match(/<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/) || [])[1] || '';
      const colors = {};
      tags.forEach(t => {
        const m = schemeBody.match(new RegExp('<a:' + t + '[^>]*>([\\s\\S]*?)<\\/a:' + t + '>'));
        const body = m ? m[1] : '';
        const srgb = body.match(/<a:srgbClr val="([0-9A-Fa-f]{6})"/);
        const sys = body.match(/<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/);
        colors[t] = (srgb && srgb[1]) || (sys && sys[1]) || 'FFFFFF';
      });
      return colors;
    }
    return {
      parseFromString: function (xml) {
        const colors = parseTheme(String(xml));
        return {
          getElementsByTagName: function (name) {
            if (name === 'a:clrScheme') {
              return [{
                getElementsByTagName: function (tag) {
                  const m = String(tag).match(/^a:(.+)$/);
                  const key = m ? m[1] : tag;
                  if (!colors[key]) return [];
                  return [{
                    // This node represents a:lt1/dk1/accentN; it may contain a:srgbClr or a:sysClr
                    getElementsByTagName: function (sub) {
                      const sm = String(sub).match(/^a:(.+)$/);
                      const sk = sm ? sm[1] : sub;
                      if (sk === 'srgbClr' || sk === 'sysClr') {
                        return [{ getAttribute: function (attr) { return (attr === 'val' || attr === 'lastClr') ? colors[key] : null; } }];
                      }
                      return [];
                    },
                    getAttribute: function () { return null; },
                  }];
                },
              }];
            }
            return [];
          },
        };
      },
    };
  },
};
sandbox.window = sandbox;
sandbox.globalThis = sandbox;
vm.createContext(sandbox);

const scripts = [
  'utils-color.js',
  'utils-cal.js',
  'export-npoi.js',
  'export-jrxmlV6.js',
  'export-rdlc.js',
  'export-primeng.js',
  'export-xlsx-style.js',
  'export-efcore.js',
  'export-typescript.js',
  'export-itextsharp.js',
  'app.js',
];

for (const s of scripts) {
  const src = fs.readFileSync(path.join(ROOT, s), 'utf8');
  try {
    vm.runInContext(src, sandbox, { filename: s });
  } catch (e) {
    console.log('LOAD ERROR in ' + s + ': ' + e.message);
    process.exit(1);
  }
}

// ---- replicate doExport() flow (variables live ON the sandbox) ----
async function runFile(fileName, label) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(ROOT, fileName));
  const ws = wb.getWorksheet(1);

  sandbox.wb = wb;
  sandbox.ws = ws;
  const shifted = vm.runInContext('cloneWithoutColumnA(wb, ws).worksheet', sandbox);
  sandbox.meta = vm.runInContext('cloneWithoutColumnA(wb, ws).meta', sandbox);
  sandbox.ws2 = shifted;
  const exportWs = vm.runInContext('cloneWorksheet(ws2)', sandbox);
  sandbox.ws3 = exportWs;
  const wsJson = vm.runInContext('worksheetToJson(ws3)', sandbox);
  sandbox.wb2 = wb;
  const wbJson = vm.runInContext('workbookToJson(wb2)', sandbox);

  const input = { ws: wsJson, workbook: wbJson, meta: sandbox.meta, images: [], fileName: 'test' };
  sandbox.input = input;

  const r = {};
  function safe(fn) {
    try { return fn(); }
    catch (e) { return 'ERR:' + (e.message || String(e)).slice(0, 80); }
  }
  const npoi = safe(() => vm.runInContext('exportToNPOI(input)', sandbox));
  r.npoi = typeof npoi === 'string' ? npoi.includes('for (int t = 0; t < fieldMap.Length; t++)') : npoi;
  r.npoiDeclCount = typeof npoi === 'string' ? (npoi.match(/string\[\]\[\] fieldMap/g) || []).length : -1;
  r.npoiStatic = typeof npoi === 'string' ? npoi.includes('string[][] staticMap') : npoi;
  const itx = safe(() => vm.runInContext('exportToItextSharp(input)', sandbox));
  r.itext = typeof itx === 'string' ? itx.includes('for (int t = 0; t < fieldMap.Length; t++)') : itx;
  r.itextStatic = typeof itx === 'string' ? itx.includes('string[][] staticMap') : itx;
  const xls = safe(() => vm.runInContext('exportToXlsxStyle(input)', sandbox));
  r.xlsx = typeof xls === 'string' ? xls.includes('const detailRowCount') : xls;
  r.xlsxStatic = typeof xls === 'string' ? xls.includes('const staticMap') : xls;
  const rdlc = safe(() => vm.runInContext('exportToRDLC(input)', sandbox));
  r.rdlc = typeof rdlc === 'string' ? (rdlc.match(/<TablixRow>/g) || []).length : rdlc;
  const prim = safe(() => vm.runInContext('exportToNgPrime(input)', sandbox));
  if (typeof prim === 'string') {
    const body = prim.match(/pTemplate="body"[\s\S]*?<\/ng-template>/);
    r.prim = body ? (body[0].match(/<tr>/g) || []).length : -1;
  } else {
    r.prim = prim;
  }

  console.log(label + ' => ' + JSON.stringify(r));
  return r;
}

(async () => {
  try {
    await runFile('sample-group-report.xlsx', 'group');
    await runFile('sample-simple-report.xlsx', 'simple');
    await runFile('sample-multi-group-report.xlsx', 'multi-group');
    await runFile('sample-nested-group-report.xlsx', 'nested');
  } catch (e) {
    console.log('ERROR: ' + (e && e.stack ? e.stack : String(e)));
    process.exit(1);
  }
})();
