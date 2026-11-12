// 엑셀 파일(.xlsx)을 밖의 도구 없이 만든다.
//   SiteXlsx.make([{ name, cols:[너비…], rows:[[칸…]] }]) → Uint8Array
//   칸: 글 | null | { v:숫자, s:모양 } | { t:글, s:모양 }
//   모양: 0 보통 · 1 굵게 · 2 시각(hh:mm) · 3 걸린 시간([h]:mm) · 4 굵게+걸린 시간 · 5 날짜(yyyy-mm-dd)
(function (root) {
  var TBL = null;
  function crc32(a) {
    if (!TBL) {
      TBL = [];
      for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; TBL[n] = c >>> 0; }
    }
    var crc = 0xFFFFFFFF;
    for (var i = 0; i < a.length; i++) crc = TBL[(crc ^ a[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }
  function u8(s) { return new TextEncoder().encode(s); }
  function zip(files) {
    var parts = [], central = [], off = 0, csize = 0, total = 0;
    files.forEach(function (f) {
      var name = u8(f.name), data = f.data, crc = crc32(data), n = data.length;
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(12, 0x21, true); h.setUint32(14, crc, true); h.setUint32(18, n, true); h.setUint32(22, n, true);
      h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, n, true); c.setUint32(24, n, true);
      c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + n;
      csize += 46 + name.length;
    });
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, csize, true); e.setUint32(16, off, true);
    var all = parts.concat(central, [new Uint8Array(e.buffer)]);
    all.forEach(function (p) { total += p.length; });
    var out = new Uint8Array(total), at = 0;
    all.forEach(function (p) { out.set(p, at); at += p.length; });
    return out;
  }
  function x(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function col(i) { return String.fromCharCode(65 + i); }
  function sheetName(s, used) {
    var n = String(s || '시트').replace(/[\[\]\:\*\?\/\\']/g, ' ').slice(0, 28) || '시트', base = n, k = 2;
    while (used[n]) n = base + ' ' + (k++);
    used[n] = true;
    return n;
  }
  function sheetXml(sh) {
    var s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    s += '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>';
    if (sh.cols && sh.cols.length) {
      s += '<cols>';
      sh.cols.forEach(function (w, i) { s += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; });
      s += '</cols>';
    }
    s += '<sheetData>';
    sh.rows.forEach(function (row, r) {
      s += '<row r="' + (r + 1) + '">';
      row.forEach(function (c, i) {
        if (c == null || c === '') return;
        var ref = col(i) + (r + 1);
        if (typeof c === 'string') s += '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + x(c) + '</t></is></c>';
        else if (c.t != null) s += '<c r="' + ref + '" t="inlineStr" s="' + (c.s || 0) + '"><is><t xml:space="preserve">' + x(c.t) + '</t></is></c>';
        else s += '<c r="' + ref + '" s="' + (c.s || 0) + '"><v>' + c.v + '</v></c>';
      });
      s += '</row>';
    });
    return s + '</sheetData></worksheet>';
  }
  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<numFmts count="3"><numFmt numFmtId="164" formatCode="hh:mm"/><numFmt numFmtId="165" formatCode="[h]:mm"/><numFmt numFmtId="166" formatCode="yyyy-mm-dd"/></numFmts>'
    + '<fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/></font></fonts>'
    + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="6">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    + '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'
    + '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '</cellXfs></styleSheet>';

  function make(sheets) {
    var used = {}, names = sheets.map(function (sh) { return sheetName(sh.name, used); });
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
    var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';
    var rel = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    var files = [];
    sheets.forEach(function (sh, i) {
      var n = i + 1;
      ct += '<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      wb += '<sheet name="' + x(names[i]) + '" sheetId="' + n + '" r:id="rId' + n + '"/>';
      rel += '<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>';
      files.push({ name: 'xl/worksheets/sheet' + n + '.xml', data: u8(sheetXml(sh)) });
    });
    ct += '</Types>';
    wb += '</sheets></workbook>';
    rel += '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var root = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    return zip([
      { name: '[Content_Types].xml', data: u8(ct) },
      { name: '_rels/.rels', data: u8(root) },
      { name: 'xl/workbook.xml', data: u8(wb) },
      { name: 'xl/_rels/workbook.xml.rels', data: u8(rel) },
      { name: 'xl/styles.xml', data: u8(STYLES) }
    ].concat(files));
  }

  root.SiteXlsx = { make: make };
})(typeof window !== 'undefined' ? window : globalThis);
