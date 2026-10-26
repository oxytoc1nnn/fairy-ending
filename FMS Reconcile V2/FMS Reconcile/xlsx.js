(() => {
  if (globalThis.FmsXlsx) return;

  const enc = new TextEncoder();
  const CRC = new Uint32Array(256).map((_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc32 = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };

  async function deflate(data) {
    const stream = new Blob([data]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function zip(files) {
    const now = new Date();
    const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const parts = [], central = [];
    let offset = 0;
    for (const [name, text] of files) {
      const nameBuf = enc.encode(name), raw = enc.encode(text), crc = crc32(raw), data = await deflate(raw);
      const local = new DataView(new ArrayBuffer(30));
      [[0, 0x04034b50, 4], [4, 20, 2], [6, 0x0800, 2], [8, 8, 2], [10, time, 2], [12, date, 2],
       [14, crc, 4], [18, data.length, 4], [22, raw.length, 4], [26, nameBuf.length, 2], [28, 0, 2]]
        .forEach(([o, v, n]) => (n === 4 ? local.setUint32(o, v, true) : local.setUint16(o, v, true)));
      const head = new DataView(new ArrayBuffer(46));
      [[0, 0x02014b50, 4], [4, 20, 2], [6, 20, 2], [8, 0x0800, 2], [10, 8, 2], [12, time, 2], [14, date, 2],
       [16, crc, 4], [20, data.length, 4], [24, raw.length, 4], [28, nameBuf.length, 2], [30, 0, 2],
       [32, 0, 2], [34, 0, 2], [36, 0, 2], [38, 0, 4], [42, offset, 4]]
        .forEach(([o, v, n]) => (n === 4 ? head.setUint32(o, v, true) : head.setUint16(o, v, true)));
      parts.push(new Uint8Array(local.buffer), nameBuf, data);
      central.push(new Uint8Array(head.buffer), nameBuf);
      offset += 30 + nameBuf.length + data.length;
    }
    const size = central.reduce((s, p) => s + p.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    [[0, 0x06054b50, 4], [4, 0, 2], [6, 0, 2], [8, files.length, 2], [10, files.length, 2],
     [12, size, 4], [16, offset, 4], [20, 0, 2]]
      .forEach(([o, v, n]) => (n === 4 ? end.setUint32(o, v, true) : end.setUint16(o, v, true)));
    return new Blob([...parts, ...central, new Uint8Array(end.buffer)],
      {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }

  const esc = (v) => String(v ?? "")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const col = (i) => String.fromCharCode(65 + i);

  async function build(sheetName, header, rows, widths) {
    const last = rows.length + 1, lastCol = col(header.length - 1);
    const cells = (vals, r, s) => vals.map((v, i) =>
      `<c r="${col(i)}${r}" t="inlineStr" s="${s}"><is><t xml:space="preserve">${esc(v)}</t></is></c>`).join("");
    const data = [`<row r="1">${cells(header, 1, 1)}</row>`]
      .concat(rows.map((row, i) => `<row r="${i + 2}">${cells(row, i + 2, 0)}</row>`)).join("");
    const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const rel = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    return await zip([
      ["[Content_Types].xml", xml + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        "</Types>"],
      ["_rels/.rels", xml + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        "</Relationships>"],
      ["xl/workbook.xml", xml + `<workbook ${ns} ${rel}><sheets><sheet name="${esc(sheetName)}" sheetId="1" r:id="rId1"/></sheets>` +
        `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'${esc(sheetName)}'!$A$1:$${lastCol}$${last}</definedName></definedNames>` +
        "</workbook>"],
      ["xl/_rels/workbook.xml.rels", xml + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        "</Relationships>"],
      ["xl/styles.xml", xml + `<styleSheet ${ns}>` +
        '<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        "</styleSheet>"],
      ["xl/worksheets/sheet1.xml", xml + `<worksheet ${ns} ${rel}><dimension ref="A1:${lastCol}${last}"/>` +
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        "<cols>" + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("") + "</cols>" +
        `<sheetData>${data}</sheetData><autoFilter ref="A1:${lastCol}${last}"/></worksheet>`],
    ]);
  }

  globalThis.FmsXlsx = {build};
})();
