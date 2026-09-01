#target "InDesign"

(function () {
  /* ================= CONFIG ================= */
  var ACTIVE_JOB_CONFIG_PATH = "D:\\Google Drive\\Publishing\\Code\\Programs\\Translate\\02 Translate Text\\Active Job.json";
  var OUTPUT_XLSX_PATH = "";
  var REPORT_DIR = "";
  var WORKSHEET_NAME = "content_export";
  var DIGITS = 5;
  var INCLUDE_HEADER = true;
  var MAX_ICML_FILES = 5000;
  var MAX_EXCEL_CELL_CHARS = 32767;
  var REPORT_FILENAME_PREFIX = "content_export_report_";
  /* ========================================= */

  function joinPath(folderPath, filename) {
    var tail = folderPath.charAt(folderPath.length - 1);
    return folderPath + (tail === "\\" || tail === "/" ? "" : "\\") + filename;
  }

  function desktopPath(filename) {
    return joinPath(Folder.desktop.fsName, filename);
  }

  function timestampForFilename() {
    var d = new Date();
    function pad(n) { return (n < 10 ? "0" : "") + n; }
    return d.getFullYear() +
      pad(d.getMonth() + 1) +
      pad(d.getDate()) + "_" +
      pad(d.getHours()) +
      pad(d.getMinutes()) +
      pad(d.getSeconds());
  }

  function makeTempFilePath(filename) {
    return joinPath(Folder.temp.fsName, filename);
  }

  function makeReportPath(prefix) {
    var folder = REPORT_DIR || Folder.desktop.fsName;
    return joinPath(folder, prefix + timestampForFilename() + ".txt");
  }

  function isoNow() {
    var d = new Date();
    try { return d.toISOString(); } catch (e) { return d.toUTCString(); }
  }

  function padN(n, digits) {
    var s = String(n);
    while (s.length < digits) s = "0" + s;
    return s;
  }

  function readFileUtf8(path) {
    var f = new File(path);
    if (!f.exists) throw new Error("Missing file:\n" + path);
    f.encoding = "UTF-8";
    if (!f.open("r")) throw new Error("Cannot open for read:\n" + path);
    var s = f.read();
    f.close();
    return s;
  }

  function writeFileUtf8(path, text) {
    var f = new File(path);
    f.encoding = "UTF-8";
    f.lineFeed = "Unix";
    if (!f.open("w")) throw new Error("Cannot open for write:\n" + path);
    f.write(text);
    f.close();
  }

  function ensureJsonSupport() {
    if (typeof JSON !== "undefined" && JSON.parse && JSON.stringify) return JSON;
    function quote(value) {
      return '"' + String(value).replace(/[\\"\u0000-\u001F]/g, function (character) {
        var escapes = { '"': '\\\"', '\\': '\\\\', '\b': '\\b', '\f': '\\f', '\n': '\\n', '\r': '\\r', '\t': '\\t' };
        if (escapes[character]) return escapes[character];
        var hex = character.charCodeAt(0).toString(16);
        while (hex.length < 4) hex = "0" + hex;
        return "\\u" + hex;
      }) + '"';
    }
    function encode(value, gap, level) {
      if (value === null) return "null";
      var valueType = typeof value;
      if (valueType === "string") return quote(value);
      if (valueType === "number") return isFinite(value) ? String(value) : "null";
      if (valueType === "boolean") return value ? "true" : "false";
      if (valueType !== "object") return undefined;
      var indent = "";
      for (var indentIndex = 0; indentIndex < level; indentIndex++) indent += gap;
      var childIndent = indent + gap;
      var parts = [];
      if (Object.prototype.toString.call(value) === "[object Array]") {
        for (var arrayIndex = 0; arrayIndex < value.length; arrayIndex++) {
          var arrayValue = encode(value[arrayIndex], gap, level + 1);
          parts.push(arrayValue === undefined ? "null" : arrayValue);
        }
        if (!parts.length) return "[]";
        return gap ? "[\n" + childIndent + parts.join(",\n" + childIndent) + "\n" + indent + "]" : "[" + parts.join(",") + "]";
      }
      for (var key in value) {
        if (!value.hasOwnProperty(key)) continue;
        var encoded = encode(value[key], gap, level + 1);
        if (encoded !== undefined) parts.push(quote(key) + (gap ? ": " : ":") + encoded);
      }
      if (!parts.length) return "{}";
      return gap ? "{\n" + childIndent + parts.join(",\n" + childIndent) + "\n" + indent + "}" : "{" + parts.join(",") + "}";
    }
    var implementation = {
      parse: function (text) { return eval("(" + String(text) + ")"); },
      stringify: function (value, replacer, space) {
        var gap = "";
        if (typeof space === "number") { for (var i = 0; i < Math.min(10, space); i++) gap += " "; }
        else if (typeof space === "string") gap = space.slice(0, 10);
        return encode(value, gap, 0);
      }
    };
    $.global.JSON = implementation;
    return implementation;
  }

  var JSON = ensureJsonSupport();

  function readJsonFile(path, label) {
    var raw = readFileUtf8(path).replace(/^\uFEFF/, "");
    try { return JSON.parse(raw); }
    catch (e) { throw new Error("Could not parse " + label + ":\n" + path + "\n" + e); }
  }

  function writeJsonFile(path, value) {
    writeFileUtf8(path, JSON.stringify(value, null, 2) + "\n");
  }

  function requireFolder(path, label) {
    var folder = new Folder(path);
    if (!folder.exists) throw new Error("Missing " + label + ":\n" + path);
    return folder.fsName;
  }

  function fnv1a32Hex(text) {
    text = String(text || "");
    var hash = 2166136261;
    for (var i = 0; i < text.length; i++) {
      hash ^= text.charCodeAt(i);
      hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
      hash >>>= 0;
    }
    var hex = (hash >>> 0).toString(16).toUpperCase();
    while (hex.length < 8) hex = "0" + hex;
    return hex;
  }

  function fileExists(path) {
    try { return File(path).exists; } catch (e) { return false; }
  }

  function deleteIfExists(path) {
    try {
      var f = new File(path);
      if (f.exists) f.remove();
    } catch (e) {}
  }

  function tsvEscape(v) {
    if (v === null || v === undefined) v = "";
    return String(v)
      .replace(/\\/g, "\\\\")
      .replace(/\t/g, "\\t")
      .replace(/\r/g, "\\r")
      .replace(/\n/g, "\\n");
  }

  function tsvUnescape(s) {
    s = String(s || "");
    return s
      .replace(/\\\\/g, "\uE000")
      .replace(/\\t/g, "\t")
      .replace(/\\r/g, "\r")
      .replace(/\\n/g, "\n")
      .replace(/\uE000/g, "\\");
  }

  function validateWorkbookRows(rows) {
    if (!rows.length) throw new Error("No workbook rows were produced.");

    for (var r = 0; r < rows.length; r++) {
      if (!rows[r] || rows[r].length !== 4) {
        throw new Error("Workbook row " + (r + 1) + " does not contain exactly four columns.");
      }
      for (var c = 0; c < 4; c++) {
        var value = String(rows[r][c] === null || rows[r][c] === undefined ? "" : rows[r][c]);
        if (value.length > MAX_EXCEL_CELL_CHARS) {
          throw new Error(
            "Workbook cell at row " + (r + 1) + ", column " + (c + 1) +
            " exceeds Excel's " + MAX_EXCEL_CELL_CHARS + "-character cell limit."
          );
        }
      }
    }
  }

  function buildXlsxViaExcel(rows, xlsxPath, worksheetName) {
    if ($.os.toLowerCase().indexOf("windows") === -1) {
      throw new Error("XLSX export requires Windows because it uses Excel COM.");
    }

    validateWorkbookRows(rows);

    var stamp = String(new Date().getTime());
    var tempInputPath = makeTempFilePath("indesign_content_export_" + stamp + ".tsv");
    var tempStatusPath = makeTempFilePath("indesign_content_export_" + stamp + ".status.tsv");
    var lines = [];

    for (var r = 0; r < rows.length; r++) {
      lines.push(
        tsvEscape(rows[r][0]) + "\t" +
        tsvEscape(rows[r][1]) + "\t" +
        tsvEscape(rows[r][2]) + "\t" +
        tsvEscape(rows[r][3])
      );
    }
    writeFileUtf8(tempInputPath, lines.join("\r\n"));

    var vbs = [
      'Option Explicit',
      '',
      'Sub Main()',
      '  Dim inputPath, xlsxPath, worksheetName, statusPath',
      '  Dim txt, lines, parts, data, rowCount, r, c',
      '  Dim excel, wb, ws, dataRange, fso, rs, errMsg',
      '',
      '  inputPath = CStr(arguments(0))',
      '  xlsxPath = CStr(arguments(1))',
      '  worksheetName = CStr(arguments(2))',
      '  statusPath = CStr(arguments(3))',
      '',
      '  On Error Resume Next',
      '  txt = ReadUtf8Text(inputPath)',
      '  If Err.Number <> 0 Then',
      '    errMsg = "Could not read temporary export data: " & Err.Description',
      '    Err.Clear',
      '    Fail statusPath, errMsg',
      '    Exit Sub',
      '  End If',
      '',
      '  If Len(txt) > 0 And Left(txt, 1) = ChrW(&HFEFF) Then txt = Mid(txt, 2)',
      '  txt = Replace(txt, vbCrLf, vbLf)',
      '  txt = Replace(txt, vbCr, vbLf)',
      '  Do While Len(txt) > 0 And Right(txt, 1) = vbLf',
      '    txt = Left(txt, Len(txt) - 1)',
      '  Loop',
      '  If txt = "" Then',
      '    Fail statusPath, "Temporary export data is empty."',
      '    Exit Sub',
      '  End If',
      '',
      '  lines = Split(txt, vbLf)',
      '  rowCount = UBound(lines) + 1',
      '  ReDim data(rowCount - 1, 3)',
      '',
      '  For r = 0 To rowCount - 1',
      '    parts = Split(lines(r), vbTab)',
      '    If UBound(parts) <> 3 Then',
      '      Fail statusPath, "Temporary row " & CStr(r + 1) & " does not contain four columns."',
      '      Exit Sub',
      '    End If',
      '    For c = 0 To 3',
      '      data(r, c) = Unesc(parts(c))',
      '    Next',
      '  Next',
      '',
      '  Set excel = CreateObject("Excel.Application")',
      '  If Err.Number <> 0 Then',
      '    errMsg = "Could not start Excel: " & Err.Description',
      '    Err.Clear',
      '    Fail statusPath, errMsg',
      '    Exit Sub',
      '  End If',
      '  excel.Visible = False',
      '  excel.DisplayAlerts = False',
      '',
      '  Set wb = excel.Workbooks.Add',
      '  If Err.Number <> 0 Then',
      '    errMsg = "Could not create workbook: " & Err.Description',
      '    Err.Clear',
      '    Cleanup excel, wb, ws',
      '    Fail statusPath, errMsg',
      '    Exit Sub',
      '  End If',
      '',
      '  Do While wb.Worksheets.Count > 1',
      '    wb.Worksheets(wb.Worksheets.Count).Delete',
      '  Loop',
      '  Set ws = wb.Worksheets(1)',
      '  ws.Name = worksheetName',
      '  If Err.Number <> 0 Then',
      '    errMsg = "Could not prepare worksheet: " & Err.Description',
      '    Err.Clear',
      '    Cleanup excel, wb, ws',
      '    Fail statusPath, errMsg',
      '    Exit Sub',
      '  End If',
      '',
      '  Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, 4))',
      '  dataRange.NumberFormat = "@"',
      '  Set rs = CreateObject("ADODB.Recordset")',
      '  rs.Fields.Append "Column1", 203, 32767',
      '  rs.Fields.Append "Column2", 203, 32767',
      '  rs.Fields.Append "Column3", 203, 32767',
      '  rs.Fields.Append "Column4", 203, 32767',
      '  rs.Open',
      '  For r = 0 To rowCount - 1',
      '    rs.AddNew',
      '    For c = 0 To 3',
      '      rs.Fields(c).Value = data(r, c)',
      '    Next',
      '    rs.Update',
      '  Next',
      '  rs.MoveFirst',
      '  ws.Range("A1").CopyFromRecordset rs',
      '  rs.Close',
      '  Set rs = Nothing',
      '  If Err.Number <> 0 Then',
      '    errMsg = "Could not populate workbook: " & Err.Description',
      '    Err.Clear',
      '    Cleanup excel, wb, ws',
      '    Fail statusPath, errMsg',
      '    Exit Sub',
      '  End If',
      '',
      '  ws.Range("A1:D1").Font.Bold = True',
      '  ws.Range("A1:D1").Interior.Color = RGB(31, 78, 121)',
      '  ws.Range("A1:D1").Font.Color = RGB(255, 255, 255)',
      '  ws.Range("A1:D1").AutoFilter',
      '  ws.Columns("A").ColumnWidth = 24',
      '  ws.Columns("B").ColumnWidth = 70',
      '  ws.Columns("C").ColumnWidth = 18',
      '  ws.Columns("D").ColumnWidth = 70',
      '  ws.Columns("A").HorizontalAlignment = -4131',
      '  ws.Columns("C").HorizontalAlignment = -4131',
      '  dataRange.VerticalAlignment = -4160',
      '  dataRange.WrapText = False',
      '',
      '  ws.Activate',
      '  excel.ActiveWindow.SplitRow = 1',
      '  excel.ActiveWindow.FreezePanes = True',
      '  Err.Clear',
      '',
      '  Set fso = CreateObject("Scripting.FileSystemObject")',
      '  If fso.FileExists(xlsxPath) Then fso.DeleteFile xlsxPath, True',
      '  If Err.Number <> 0 Then',
      '    errMsg = "Could not replace existing workbook: " & Err.Description',
      '    Err.Clear',
      '    Cleanup excel, wb, ws',
      '    Fail statusPath, errMsg',
      '    Exit Sub',
      '  End If',
      '',
      '  wb.SaveAs xlsxPath, 51',
      '  If Err.Number <> 0 Then',
      '    errMsg = "Could not save XLSX workbook: " & Err.Description',
      '    Err.Clear',
      '    Cleanup excel, wb, ws',
      '    Fail statusPath, errMsg',
      '    Exit Sub',
      '  End If',
      '',
      '  Cleanup excel, wb, ws',
      '  WriteUtf8Text statusPath, "OK" & vbTab & CStr(rowCount)',
      'End Sub',
      '',
      'Function ReadUtf8Text(path)',
      '  Dim stm',
      '  Set stm = CreateObject("ADODB.Stream")',
      '  stm.Type = 2',
      '  stm.Charset = "utf-8"',
      '  stm.Open',
      '  stm.LoadFromFile path',
      '  ReadUtf8Text = stm.ReadText',
      '  stm.Close',
      '  Set stm = Nothing',
      'End Function',
      '',
      'Function Unesc(s)',
      '  Dim marker',
      '  marker = ChrW(&HE000)',
      '  s = CStr(s)',
      '  s = Replace(s, "\\\\", marker)',
      '  s = Replace(s, "\\t", vbTab)',
      '  s = Replace(s, "\\r", vbCr)',
      '  s = Replace(s, "\\n", vbLf)',
      '  s = Replace(s, marker, "\\")',
      '  Unesc = s',
      'End Function',
      '',
      'Sub Cleanup(excel, wb, ws)',
      '  On Error Resume Next',
      '  If Not wb Is Nothing Then wb.Close False',
      '  If Not excel Is Nothing Then excel.Quit',
      '  Set ws = Nothing',
      '  Set wb = Nothing',
      '  Set excel = Nothing',
      'End Sub',
      '',
      'Sub Fail(statusPath, msg)',
      '  WriteUtf8Text statusPath, "ERROR" & vbTab & Esc(msg)',
      'End Sub',
      '',
      'Function Esc(s)',
      '  s = CStr(s)',
      '  s = Replace(s, "\\", "\\\\")',
      '  s = Replace(s, vbTab, "\\t")',
      '  s = Replace(s, vbCr, "\\r")',
      '  s = Replace(s, vbLf, "\\n")',
      '  Esc = s',
      'End Function',
      '',
      'Sub WriteUtf8Text(path, txt)',
      '  Dim stm',
      '  Set stm = CreateObject("ADODB.Stream")',
      '  stm.Type = 2',
      '  stm.Charset = "utf-8"',
      '  stm.Open',
      '  stm.WriteText txt',
      '  stm.SaveToFile path, 2',
      '  stm.Close',
      '  Set stm = Nothing',
      'End Sub',
      '',
      'Main'
    ].join("\r\n");

    try {
      app.doScript(vbs, ScriptLanguage.VISUAL_BASIC, [
        tempInputPath,
        xlsxPath,
        worksheetName,
        tempStatusPath
      ]);

      if (!fileExists(tempStatusPath)) {
        throw new Error("Excel did not return an XLSX export status.");
      }

      var status = readFileUtf8(tempStatusPath).replace(/^\uFEFF/, "");
      var statusParts = status.split("\t");
      if (statusParts[0] !== "OK") {
        throw new Error(tsvUnescape(statusParts.slice(1).join("\t")) || "Unknown Excel export error.");
      }
      if (!fileExists(xlsxPath)) {
        throw new Error("Excel reported success, but the XLSX file was not created.");
      }
    } finally {
      deleteIfExists(tempInputPath);
      deleteIfExists(tempStatusPath);
    }
  }

  function isIcml(name) {
    name = String(name || "").toLowerCase();
    return name.slice(-5) === ".icml" || name.slice(-4) === ".icm";
  }

  function pageFromFilename(filename) {
    var m = String(filename || "").match(/Page_(\d+)/i);
    return m ? m[1] : "";
  }

  function setIdAttr(startTag, idStr) {
    if (/(\s)id\s*=\s*"/i.test(startTag)) {
      return startTag.replace(/(\s)id\s*=\s*"[^"]*"/i, '$1id="' + idStr + '"');
    }
    return startTag.replace(/^<([A-Za-z0-9:_-]+)/, '<$1 id="' + idStr + '"');
  }

  function findMatchingPsrClose(text, startTagEnd) {
    var tokenRe = /<ParagraphStyleRange\b[^>]*>|<\/ParagraphStyleRange>/g;
    tokenRe.lastIndex = startTagEnd;
    var depth = 1;
    var match;

    while ((match = tokenRe.exec(text)) !== null) {
      if (/^<\/ParagraphStyleRange>/i.test(match[0])) {
        depth--;
        if (depth === 0) return match.index;
      } else {
        depth++;
      }
    }
    return -1;
  }

  function processIcmlText(icmlText, psrStartCounter, contentStartCounter) {
    var psrCounter = psrStartCounter;
    var contentCounter = contentStartCounter;
    var rows = [];
    var out = [];
    var last = 0;
    var outerPsrCount = 0;
    var rePSR = /<ParagraphStyleRange\b[^>]*>/g;
    var mPSR;

    while ((mPSR = rePSR.exec(icmlText)) !== null) {
      var psrTagStart = mPSR.index;
      var psrStartTag = mPSR[0];
      var psrStartTagEnd = psrTagStart + psrStartTag.length;

      // A self-closing ParagraphStyleRange is a valid empty ICML paragraph.
      // It has no content to export and no matching close tag to search for.
      if (/\/\>$/.test(psrStartTag)) {
        out.push(icmlText.substring(last, psrTagStart));
        out.push(psrStartTag);
        last = psrStartTagEnd;
        rePSR.lastIndex = psrStartTagEnd;
        continue;
      }

      var closeTag = "</ParagraphStyleRange>";
      var psrCloseIdx = findMatchingPsrClose(icmlText, psrStartTagEnd);
      if (psrCloseIdx < 0) {
        throw new Error("Malformed ICML: unmatched ParagraphStyleRange start tag.");
      }
      var psrCloseEnd = psrCloseIdx + closeTag.length;

      out.push(icmlText.substring(last, psrTagStart));
      var psrId = padN(psrCounter, DIGITS);
      out.push(setIdAttr(psrStartTag, psrId));

      var psrInner = icmlText.substring(psrStartTagEnd, psrCloseIdx);
      var innerOut = [];
      var innerLast = 0;
      var contentRows = [];
      var psrConcat = "";
      var reToken =
        /<ParagraphStyleRange\b[^>]*>|<\/ParagraphStyleRange>|<Content\b[^>]*\/?>|<\/Content>|<Br\b[^>]*\/>/gi;
      var mT;

      while ((mT = reToken.exec(psrInner)) !== null) {
        var tok = mT[0];
        var tokStart = mT.index;
        var tokEnd = tokStart + tok.length;
        innerOut.push(psrInner.substring(innerLast, tokStart));

        if (/^<ParagraphStyleRange\b/i.test(tok) || /^<\/ParagraphStyleRange>/i.test(tok) || /^<Br\b/i.test(tok)) {
          innerOut.push(tok);
          innerLast = tokEnd;
          continue;
        }

        if (/^<Content\b/i.test(tok)) {
          var selfClosing = /\/>$/.test(tok);
          var cId = padN(contentCounter, DIGITS);
          innerOut.push(setIdAttr(tok, cId));

          var cInner = "";
          if (!selfClosing) {
            var closeIdx = psrInner.indexOf("</Content>", tokEnd);
            if (closeIdx < 0) {
              throw new Error("Malformed ICML: unmatched Content start tag.");
            }
            cInner = psrInner.substring(tokEnd, closeIdx);
          }

          psrConcat += cInner;
          contentRows.push({ contentId: cId, contentInner: cInner });
          contentCounter++;
          innerLast = tokEnd;
          continue;
        }

        innerOut.push(tok);
        innerLast = tokEnd;
      }

      innerOut.push(psrInner.substring(innerLast));
      out.push(innerOut.join(""));
      out.push(closeTag);

      for (var r = 0; r < contentRows.length; r++) {
        rows.push([
          r === 0 ? psrId : "",
          r === 0 ? psrConcat : "",
          contentRows[r].contentId,
          contentRows[r].contentInner
        ]);
      }

      psrCounter++;
      outerPsrCount++;
      last = psrCloseEnd;
      rePSR.lastIndex = psrCloseEnd;
    }

    out.push(icmlText.substring(last));
    return {
      updatedText: out.join(""),
      rows: rows,
      outerPsrCount: outerPsrCount,
      nextPsrCounter: psrCounter,
      nextContentCounter: contentCounter
    };
  }

  function collectUniqueIcmlLinks(doc) {
    var files = [];
    var seen = {};
    var duplicatesSkipped = 0;

    for (var i = 0; i < doc.links.length; i++) {
      var link = doc.links[i];
      if (!isIcml(link.name)) continue;

      var fp = "";
      try { fp = link.filePath; } catch (e) { fp = ""; }
      if (!fp) continue;

      var key = String(fp).toLowerCase();
      if (seen[key]) {
        duplicatesSkipped++;
        continue;
      }
      seen[key] = true;
      files.push({
        name: String(link.name),
        path: String(fp),
        page: pageFromFilename(link.name)
      });
    }

    files.sort(function (a, b) {
      var ap = parseInt(a.page || "999999", 10);
      var bp = parseInt(b.page || "999999", 10);
      if (ap !== bp) return ap - bp;
      return a.name.localeCompare(b.name);
    });

    return { files: files, duplicatesSkipped: duplicatesSkipped };
  }

  function writePlansWithRollback(plans) {
    var written = [];
    try {
      for (var i = 0; i < plans.length; i++) {
        if (plans[i].updatedText === plans[i].originalText) continue;
        written.push(plans[i]);
        writeFileUtf8(plans[i].path, plans[i].updatedText);
      }
    } catch (writeError) {
      var rollbackFailures = [];
      for (var r = written.length - 1; r >= 0; r--) {
        try {
          writeFileUtf8(written[r].path, written[r].originalText);
        } catch (rollbackError) {
          rollbackFailures.push(written[r].path + ": " + rollbackError);
        }
      }
      throw new Error(
        "ICML write failed and rollback was attempted.\n" + writeError +
        (rollbackFailures.length ? "\nRollback failures:\n" + rollbackFailures.join("\n") : "")
      );
    }
    return written;
  }

  function restorePlans(plans) {
    var failures = [];
    for (var i = plans.length - 1; i >= 0; i--) {
      try {
        writeFileUtf8(plans[i].path, plans[i].originalText);
      } catch (e) {
        failures.push(plans[i].path + ": " + e);
      }
    }
    if (failures.length) {
      throw new Error("Rollback failures:\n" + failures.join("\n"));
    }
  }

  function writeRunReport(lines, prefix) {
    var path = makeReportPath(prefix || REPORT_FILENAME_PREFIX);
    writeFileUtf8(path, lines.join("\n"));
    return path;
  }

  try {
    var suppliedDocument = $.global.TRANSLATION_DOCUMENT;
    if ((!suppliedDocument || !suppliedDocument.isValid) && app.documents.length === 0) {
      throw new Error("No InDesign document is open.");
    }

    var activeJob = readJsonFile(ACTIVE_JOB_CONFIG_PATH, "active-job configuration");
    var jobPath = requireFolder(String(activeJob.jobPath || ""), "active translation job folder");
    var jobConfigPath = joinPath(jobPath, "job_config.json");
    var jobConfig = readJsonFile(jobConfigPath, "job configuration");
    var configuredJobPath = new Folder(String(jobConfig.jobPath || "")).fsName;
    if (configuredJobPath.toLowerCase() !== jobPath.toLowerCase()) {
      throw new Error("Active-job pointer and job configuration disagree.\n" + ACTIVE_JOB_CONFIG_PATH);
    }
    if (String(activeJob.jobId || "") !== String(jobConfig.jobId || "")) {
      throw new Error("Active-job ID and job configuration ID disagree.\n" + ACTIVE_JOB_CONFIG_PATH);
    }

    var inputDir = requireFolder(joinPath(jobPath, "input"), "job input folder");
    REPORT_DIR = requireFolder(joinPath(jobPath, "reports"), "job reports folder");
    OUTPUT_XLSX_PATH = joinPath(inputDir, "content_export.xlsx");
    var manifestPath = joinPath(jobPath, "job_manifest.json");
    if (fileExists(manifestPath)) {
      throw new Error("This job already has an export manifest. Create a new job instead of replacing its matched snapshot.\n" + manifestPath);
    }

    var doc = (suppliedDocument && suppliedDocument.isValid) ? suppliedDocument : app.activeDocument;
    var documentPath = "";
    try {
      if (!doc.saved) throw new Error("The document has never been saved.");
      if (doc.modified) throw new Error("The document has unsaved changes. Save a checkpoint before exporting.");
      documentPath = doc.fullName.fsName;
    } catch (documentError) {
      throw new Error("The active document must be saved and unmodified before export.\n" + documentError);
    }

    var linkResult = collectUniqueIcmlLinks(doc);
    var files = linkResult.files;

    if (!files.length) throw new Error("No readable .icml/.icm links with file paths were found.");
    if (files.length > MAX_ICML_FILES) {
      throw new Error("Refusing to process " + files.length + " ICML files; configured limit is " + MAX_ICML_FILES + ".");
    }

    var workbookRows = [];
    if (INCLUDE_HEADER) {
      workbookRows.push([
        "ParagraphStyleRange id",
        "ParagraphStyleRange content",
        "Content tag",
        "Content content"
      ]);
    }

    var plans = [];
    var psrCounter = 0;
    var contentCounter = 0;
    var totalPSR = 0;
    var totalContents = 0;

    for (var f = 0; f < files.length; f++) {
      var fileInfo = files[f];
      var original = readFileUtf8(fileInfo.path);
      var res = processIcmlText(original, psrCounter, contentCounter);

      plans.push({
        name: fileInfo.name,
        page: fileInfo.page,
        path: fileInfo.path,
        originalText: original,
        updatedText: res.updatedText,
        fingerprint: fnv1a32Hex(res.updatedText),
        contentIdCount: res.rows.length,
        outerPsrCount: res.outerPsrCount
      });

      for (var rr = 0; rr < res.rows.length; rr++) workbookRows.push(res.rows[rr]);
      totalPSR += res.outerPsrCount;
      totalContents += res.rows.length;
      psrCounter = res.nextPsrCounter;
      contentCounter = res.nextContentCounter;
    }

    var tempXlsxPath = makeTempFilePath("content_export_" + timestampForFilename() + ".xlsx");
    try {
      buildXlsxViaExcel(workbookRows, tempXlsxPath, WORKSHEET_NAME);
      var writtenPlans = writePlansWithRollback(plans);
      var modifiedFiles = writtenPlans.length;

      try {
        var tempXlsx = new File(tempXlsxPath);
        var finalXlsx = new File(OUTPUT_XLSX_PATH);
        if (finalXlsx.exists && !finalXlsx.remove()) {
          throw new Error("Cannot replace existing workbook:\n" + OUTPUT_XLSX_PATH);
        }
        if (!tempXlsx.copy(finalXlsx.fsName)) {
          throw new Error("Cannot copy completed workbook to:\n" + OUTPUT_XLSX_PATH);
        }
        tempXlsx.remove();
      } catch (promoteError) {
        try {
          restorePlans(writtenPlans);
        } catch (restoreError) {
          throw new Error(promoteError + "\n" + restoreError);
        }
        throw promoteError;
      }

      var manifestFiles = [];
      for (var mf = 0; mf < plans.length; mf++) {
        manifestFiles.push({
          name: plans[mf].name,
          page: plans[mf].page,
          path: plans[mf].path,
          fingerprint: plans[mf].fingerprint,
          fingerprintAlgorithm: "FNV1A32_UTF16",
          contentIdCount: plans[mf].contentIdCount,
          outerPsrCount: plans[mf].outerPsrCount
        });
      }

      var manifest = {
        schemaVersion: 1,
        jobId: String(jobConfig.jobId || ""),
        jobPath: jobPath,
        book: String(jobConfig.book || ""),
        targetLanguage: String(jobConfig.targetLanguage || ""),
        glossaryProfile: String(jobConfig.glossaryProfile || ""),
        glossaryTermKey: String(jobConfig.glossaryTermKey || ""),
        glossaryDefinitionKey: String(jobConfig.glossaryDefinitionKey || ""),
        status: "exported",
        createdAt: isoNow(),
        updatedAt: isoNow(),
        document: {
          name: String(doc.name),
          path: String(documentPath)
        },
        workbook: {
          worksheet: WORKSHEET_NAME,
          inputPath: OUTPUT_XLSX_PATH,
          rowCount: workbookRows.length,
          dataRowCount: workbookRows.length - (INCLUDE_HEADER ? 1 : 0),
          contentIdCount: totalContents,
          paragraphStyleRangeCount: totalPSR
        },
        icmlFingerprintAlgorithm: "FNV1A32_UTF16",
        icmlFiles: manifestFiles,
        export: {
          duplicateLinkedPathsSkipped: linkResult.duplicatesSkipped,
          filesProcessed: files.length,
          filesModified: modifiedFiles
        }
      };

      try {
        writeJsonFile(manifestPath, manifest);
      } catch (manifestError) {
        try { restorePlans(writtenPlans); } catch (manifestRestoreError) {
          throw new Error(manifestError + "\n" + manifestRestoreError);
        }
        deleteIfExists(OUTPUT_XLSX_PATH);
        throw manifestError;
      }

      var reportPath = writeRunReport([
        "CONTENT EXPORT REPORT",
        "=====================",
        "",
        "Run time: " + (new Date()).toString(),
        "Document: " + doc.name,
        "Job: " + String(jobConfig.jobId || ""),
        "Book: " + String(jobConfig.book || ""),
        "Target language: " + String(jobConfig.targetLanguage || ""),
        "Glossary profile: " + String(jobConfig.glossaryProfile || ""),
        "Workbook: " + OUTPUT_XLSX_PATH,
        "Manifest: " + manifestPath,
        "Worksheet: " + WORKSHEET_NAME,
        "ICML files processed: " + files.length,
        "Duplicate linked paths skipped: " + linkResult.duplicatesSkipped,
        "ICML files modified: " + modifiedFiles,
        "Outer ParagraphStyleRange tags exported: " + totalPSR,
        "Content tags exported: " + totalContents,
        "Workbook data rows: " + (workbookRows.length - (INCLUDE_HEADER ? 1 : 0)),
        "",
        "The workbook and the ID-tagged ICML files are one matched snapshot."
      ]);
      $.writeln("Content export completed. Report: " + reportPath);
    } finally {
      deleteIfExists(tempXlsxPath);
    }
  } catch (e) {
    var errorText = "CONTENT EXPORT ERROR\n====================\n\n" +
      "Run time: " + (new Date()).toString() + "\n" +
      "Error: " + e;
    try {
      var errorPath = writeRunReport(errorText.split("\n"), "content_export_error_");
      $.writeln("Content export failed. Report: " + errorPath);
    } catch (reportError) {
      $.writeln(errorText + "\nReport error: " + reportError);
    }
  }
})();
