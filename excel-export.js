"use strict";

async function downloadTravelWorkbook(result) {
  const spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
  const contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml";
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const escape = (value) => String(value).replace(/[&<>"]/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"
  })[char]);
  const zip = new window.JSZip();

  zip.file("[Content_Types].xml", xml + `
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="${contentType}.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="${contentType}.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="${contentType}.styles+xml"/>
    </Types>`);
  zip.file("_rels/.rels", xml + `
    <Relationships xmlns="${packageRelationships}">
      <Relationship Id="rId1" Type="${officeRelationships}/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`);
  zip.file("xl/workbook.xml", xml + `
    <workbook xmlns="${spreadsheet}" xmlns:r="${officeRelationships}">
      <bookViews><workbookView/></bookViews>
      <sheets><sheet name="출장여비 결과" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", xml + `
    <Relationships xmlns="${packageRelationships}">
      <Relationship Id="rId1" Type="${officeRelationships}/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="${officeRelationships}/styles" Target="styles.xml"/>
    </Relationships>`);
  zip.file("xl/styles.xml", xml + `
    <styleSheet xmlns="${spreadsheet}">
      <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
      <fonts count="3">
        <font><sz val="11"/><name val="Malgun Gothic"/></font>
        <font><b/><sz val="16"/><color rgb="FF102C36"/><name val="Malgun Gothic"/></font>
        <font><b/><sz val="11"/><color rgb="FF075B51"/><name val="Malgun Gothic"/></font>
      </fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFDBEEE9"/><bgColor indexed="64"/></patternFill></fill>
      </fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="7">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
        <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
        <xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="3" fontId="2" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>
      </cellXfs>
      <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    </styleSheet>`);

  const rows = [];
  const merges = [];
  function addRow(cells, { height, merge } = {}) {
    const row = rows.length + 1;
    const contents = cells.map(([value, style = 0], index) => {
      const address = "ABC"[index] + row;
      const opening = `<c r="${address}" s="${style}"`;
      return typeof value === "number"
        ? `${opening} t="n"><v>${value}</v></c>`
        : `${opening} t="inlineStr"><is><t>${escape(value)}</t></is></c>`;
    });
    if (merge) merges.push(merge[0] + row + ":" + merge[1] + row);
    const sizing = height ? ` ht="${height}" customHeight="1"` : "";
    rows.push(`<row r="${row}"${sizing}>${contents.join("")}</row>`);
  }

  function dateCell(date) {
    if (date.length !== 10 || date < "1900-01-01") return [date];
    const epochDays = Date.parse(date) / 86_400_000;
    const leapBugAdjustment = date < "1900-03-01" ? 1 : 0;
    return [epochDays + 25_569 - leapBugAdjustment, 4];
  }

  addRow([["국내 출장여비 계산 결과", 1]], { height: 28, merge: "AC" });
  addRow([["출장 유형 및 기간"], [result.context]], { merge: "BC" });
  addRow([["출장일"], dateCell(result.start)]);
  addRow([["복귀일"], dateCell(result.end)]);
  addRow([]);
  addRow([["지급 항목", 2], ["금액 (원)", 2], ["계산 기준", 2]]);
  for (const [label, amount, detail] of result.rows) {
    addRow([[label], [amount, 3], [detail]]);
  }
  addRow([["예상 지급액", 2], [result.total, 5], ["", 2]], { height: 22 });
  if (result.note) addRow([[result.note, 6]], { height: 42, merge: "AC" });

  zip.file("xl/worksheets/sheet1.xml", xml + `
    <worksheet xmlns="${spreadsheet}">
      <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
      <sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
      <sheetFormatPr defaultRowHeight="19"/>
      <cols>
        <col min="1" max="1" width="32" customWidth="1"/>
        <col min="2" max="2" width="17" customWidth="1"/>
        <col min="3" max="3" width="48" customWidth="1"/>
      </cols>
      <sheetData>${rows.join("")}</sheetData>
      <mergeCells count="${merges.length}">${merges.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>
      <pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.2" footer="0.2"/>
      <pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="1"/>
    </worksheet>`);
  const blob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    mimeType: contentType + ".sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "출장여비_" + result.start + "_" + result.end + ".xlsx";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
