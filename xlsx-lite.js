/*
 * xlsx-lite — écrit un classeur .xlsx (une feuille) sans dépendance ni réseau.
 * Reprend la petite partie de l'API SheetJS utilisée par l'application :
 *   XLSX.utils.book_new(), XLSX.utils.book_append_sheet(wb, ws, nom), XLSX.write(wb) -> ArrayBuffer
 * Cellules : { t: "n" | "s", v: valeur, f: "formule sans =", z: "0.00" | "d-mmm" }, plus ws["!ref"], ws["!cols"], ws["!merges"].
 */
(function () {
  "use strict";

  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function utf8(s) { return new TextEncoder().encode(s); }

  // Archive ZIP sans compression (méthode « stored »), suffisante pour un .xlsx
  function zip(files) {
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    var chunks = [], central = [], offset = 0;
    files.forEach(function (f) {
      var name = utf8(f.name), data = utf8(f.data), crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, dosTime, true); h.setUint16(12, dosDate, true); h.setUint32(14, crc, true);
      h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      chunks.push(new Uint8Array(h.buffer), name, data);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
      c.setUint16(12, dosTime, true); c.setUint16(14, dosDate, true); c.setUint32(16, crc, true);
      c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, name.length, true);
      c.setUint16(30, 0, true); c.setUint16(32, 0, true); c.setUint16(34, 0, true); c.setUint16(36, 0, true); c.setUint32(38, 0, true);
      c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var cdSize = central.reduce(function (s, a) { return s + a.length; }, 0);
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true); e.setUint32(16, offset, true);
    var parts = chunks.concat(central, [new Uint8Array(e.buffer)]);
    var total = parts.reduce(function (s, a) { return s + a.length; }, 0), out = new Uint8Array(total), p = 0;
    parts.forEach(function (a) { out.set(a, p); p += a.length; });
    return out;
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; })
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  }
  function colIndex(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64); return n - 1; }
  function colName(i) { var s = ""; i += 1; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

  // Styles : 0 normal · 1 nombre 0.00 · 2 date d-mmm · 3 en-tête gras · 4 gras 0.00 · 5 gras
  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="16" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function sheetXml(ws) {
    var rows = {}, lastRow = 0;
    Object.keys(ws).forEach(function (addr) {
      if (addr.charAt(0) === "!") return;
      var m = addr.match(/^([A-Z]+)(\d+)$/); if (!m) return;
      var r = +m[2], c = colIndex(m[1]);
      (rows[r] = rows[r] || []).push({ c: c, addr: addr, cell: ws[addr] });
      if (r > lastRow) lastRow = r;
    });
    // La ligne contenant « TOTAL » en colonne A est mise en gras
    var boldRows = {};
    Object.keys(rows).forEach(function (r) { rows[r].forEach(function (x) { if (x.c === 0 && x.cell.t === "s" && x.cell.v === "TOTAL") boldRows[r] = true; }); });
    var out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'];
    if (ws["!ref"]) out.push('<dimension ref="' + esc(ws["!ref"]) + '"/>');
    out.push('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>');
    if (ws["!cols"] && ws["!cols"].length) {
      out.push("<cols>");
      ws["!cols"].forEach(function (c, i) { if (c && c.wch) out.push('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (c.wch + 1) + '" customWidth="1"/>'); });
      out.push("</cols>");
    }
    out.push("<sheetData>");
    Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      out.push('<row r="' + r + '">');
      rows[r].sort(function (a, b) { return a.c - b.c; }).forEach(function (x) {
        var cell = x.cell, bold = r === 1 || boldRows[r];
        var s = cell.z === "0.00" ? (bold ? 4 : 1) : cell.z === "d-mmm" ? 2 : (bold ? (r === 1 ? 3 : 5) : 0);
        var sAttr = s ? ' s="' + s + '"' : "";
        if (cell.t === "s") {
          out.push('<c r="' + x.addr + '" t="inlineStr"' + sAttr + '><is><t xml:space="preserve">' + esc(cell.v) + "</t></is></c>");
        } else {
          var v = typeof cell.v === "number" && isFinite(cell.v) ? '<v>' + cell.v + "</v>" : "";
          var f = cell.f ? "<f>" + esc(String(cell.f).replace(/^=/, "")) + "</f>" : "";
          out.push('<c r="' + x.addr + '"' + sAttr + ">" + f + v + "</c>");
        }
      });
      out.push("</row>");
    });
    out.push("</sheetData>");
    var merges = ws["!merges"] || [];
    if (merges.length) {
      out.push('<mergeCells count="' + merges.length + '">');
      merges.forEach(function (m) { out.push('<mergeCell ref="' + colName(m.s.c) + (m.s.r + 1) + ":" + colName(m.e.c) + (m.e.r + 1) + '"/>'); });
      out.push("</mergeCells>");
    }
    out.push("</worksheet>");
    return out.join("");
  }

  function write(wb) {
    var name = (wb.SheetNames[0] || "Feuille1").replace(/[\\\/:*?"<>|\[\]]/g, "-").slice(0, 31);
    var ws = wb.Sheets[wb.SheetNames[0]] || {};
    var files = [
      { name: "[Content_Types].xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: "_rels/.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: "xl/workbook.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + esc(name) + '" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>' },
      { name: "xl/_rels/workbook.xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: "xl/styles.xml", data: STYLES },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml(ws) }
    ];
    return zip(files).buffer;
  }

  var api = {
    utils: {
      book_new: function () { return { SheetNames: [], Sheets: {} }; },
      book_append_sheet: function (wb, ws, name) { wb.SheetNames.push(name); wb.Sheets[name] = ws; }
    },
    write: write
  };
  if (typeof window !== "undefined") window.XLSX = api;
  if (typeof module !== "undefined") module.exports = api;
})();
