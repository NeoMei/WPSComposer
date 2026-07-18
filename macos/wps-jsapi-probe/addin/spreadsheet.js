(function () {
  "use strict";

  function requirePath(params, name) {
    const value = params[name];
    if (typeof value !== "string" || value.length === 0) {
      throw new Error(`${name} must be a non-empty path`);
    }
    return value;
  }

  function requireOutputPath(params) {
    return requirePath(params, "outputPath");
  }

  function conversionError(error, code) {
    const allowedCodes = [
      "CONVERSION_COMMAND_FAILED",
      "INTERACTIVE_INPUT_REQUIRED",
      "NO_VISIBLE_WORKSHEETS"
    ];
    if (error && allowedCodes.indexOf(error.code) !== -1) {
      return error;
    }
    const message = String(error && error.message ? error.message : error);
    const typed = new Error(message);
    typed.code = code || (
      /password|passcode|protected|dialog|prompt|密码|对话框/i.test(message)
        ? "INTERACTIVE_INPUT_REQUIRED"
        : "CONVERSION_COMMAND_FAILED"
    );
    return typed;
  }

  function generationError(error, code) {
    const allowedCodes = [
      "GENERATION_COMMAND_FAILED",
      "OPERATION_PLAN_INVALID"
    ];
    if (error && allowedCodes.indexOf(error.code) !== -1) {
      return error;
    }
    const message = String(error && error.message ? error.message : error);
    const typed = new Error(message);
    typed.code = code || "GENERATION_COMMAND_FAILED";
    return typed;
  }

  function invalidSheetPlan(message) {
    throw generationError(new Error(message), "OPERATION_PLAN_INVALID");
  }

  function isObject(value) {
    return value !== null && typeof value === "object" && !Array.isArray(value);
  }

  function sheetCollectionItem(collection, index) {
    if (collection && typeof collection.Item === "function") {
      return collection.Item(index);
    }
    if (typeof collection === "function") {
      return collection(index);
    }
    invalidSheetPlan("Spreadsheet collection is unavailable");
  }

  function sheetColor(value) {
    if (typeof value !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(value)) {
      return value;
    }
    const red = parseInt(value.slice(1, 3), 16);
    const green = parseInt(value.slice(3, 5), 16);
    const blue = parseInt(value.slice(5, 7), 16);
    return red | (green << 8) | (blue << 16);
  }

  function sheetCharacters(value) {
    return Array.from(String(value));
  }

  function truncateSheetName(value, maximum) {
    return sheetCharacters(value).slice(0, maximum).join("");
  }

  function sanitizeSheetName(value) {
    let name = String(value).trim();
    name = name.replace(/^'+|'+$/g, "");
    name = name.replace(/[\[\]:*?/\\]/g, "_").trim();
    if (!name) {
      name = "Sheet";
    }
    return truncateSheetName(name, 31);
  }

  function uniqueSheetName(workbook, requested, excludedSheet) {
    const used = {};
    for (let index = 1; index <= workbook.Worksheets.Count; index += 1) {
      const sheet = sheetCollectionItem(workbook.Worksheets, index);
      if (sheet !== excludedSheet) {
        used[String(sheet.Name).toLowerCase()] = true;
      }
    }
    const base = sanitizeSheetName(requested);
    let candidate = base;
    let suffixIndex = 2;
    while (used[candidate.toLowerCase()]) {
      const suffix = ` (${suffixIndex})`;
      candidate = truncateSheetName(base, 31 - sheetCharacters(suffix).length) + suffix;
      suffixIndex += 1;
    }
    return candidate;
  }

  function resetSheet(workbook, args, context) {
    while (workbook.Worksheets.Count > 1) {
      sheetCollectionItem(workbook.Worksheets, workbook.Worksheets.Count).Delete();
    }
    context.sheet = sheetCollectionItem(workbook.Worksheets, 1);
    context.sheet.Cells.Clear();
  }

  function renameSheet(workbook, args, context) {
    const sheet = sheetCollectionItem(workbook.Worksheets, args.index);
    sheet.Name = uniqueSheetName(workbook, args.name, sheet);
    if (context.sheet === sheet) {
      context.sheet = sheet;
    }
  }

  function addSheet(workbook, args, context) {
    const last = sheetCollectionItem(workbook.Worksheets, workbook.Worksheets.Count);
    const sheet = workbook.Worksheets.Add(null, last);
    sheet.Name = uniqueSheetName(workbook, args.name, sheet);
    context.sheet = sheet;
  }

  function selectSheet(workbook, args, context) {
    const sheet = sheetCollectionItem(workbook.Worksheets, args.index);
    if (typeof sheet.Activate === "function") {
      sheet.Activate();
    } else if (typeof sheet.Select === "function") {
      sheet.Select();
    }
    context.sheet = sheet;
  }

  function writeSheetTable(workbook, args, context) {
    const sheet = context.sheet;
    const rowCount = args.values.length;
    const columnCount = args.values[0].length;
    const start = sheet.Cells.Item(args.startRow, args.startCol);
    const end = sheet.Cells.Item(
      args.startRow + rowCount - 1,
      args.startCol + columnCount - 1
    );
    const range = sheet.Range(start, end);
    range.Value2 = args.values;
    range.Font.Size = args.fontSize === undefined ? 11 : args.fontSize;

    const headerEnd = sheet.Cells.Item(args.startRow, args.startCol + columnCount - 1);
    const header = sheet.Range(start, headerEnd);
    header.Font.Bold = args.headerBold === undefined ? true : args.headerBold;
    header.Font.Color = sheetColor(
      args.headerFontColor === undefined ? "#FFFFFF" : args.headerFontColor
    );
    header.Interior.Color = sheetColor(
      args.headerShade === undefined ? "#4472C4" : args.headerShade
    );
  }

  function setSheetColumnWidth(workbook, args, context) {
    const columns = context.sheet.Columns;
    const column = columns && typeof columns.Item === "function"
      ? columns.Item(args.column)
      : columns(args.column);
    column.ColumnWidth = args.width;
  }

  function autofitSheet(workbook, args, context) {
    context.sheet.UsedRange.Columns.AutoFit();
  }

  const sheetOperations = {
    "sheet.reset": resetSheet,
    "sheet.rename": renameSheet,
    "sheet.add": addSheet,
    "sheet.select": selectSheet,
    "sheet.write_table": writeSheetTable,
    "sheet.set_column_width": setSheetColumnWidth,
    "sheet.autofit": autofitSheet
  };

  const sheetOperationNames = [
    "sheet.reset",
    "sheet.rename",
    "sheet.add",
    "sheet.select",
    "sheet.write_table",
    "sheet.set_column_width",
    "sheet.autofit"
  ];

  const sheetArgumentRules = {
    "sheet.reset": {required: [], allowed: []},
    "sheet.rename": {required: ["index", "name"], allowed: ["index", "name"]},
    "sheet.add": {required: ["name"], allowed: ["name"]},
    "sheet.select": {required: ["index"], allowed: ["index"]},
    "sheet.write_table": {required: ["startRow", "startCol", "values"], allowed: ["startRow", "startCol", "values", "headerBold", "headerShade", "headerFontColor", "fontSize"]},
    "sheet.set_column_width": {required: ["column", "width"], allowed: ["column", "width"]},
    "sheet.autofit": {required: [], allowed: []}
  };

  function validateSheetTable(values) {
    if (!Array.isArray(values) || values.length === 0 || !Array.isArray(values[0]) || values[0].length === 0) {
      invalidSheetPlan("sheet.write_table values are invalid");
    }
    const width = values[0].length;
    let cells = 0;
    values.forEach(function (row) {
      if (!Array.isArray(row) || row.length !== width) {
        invalidSheetPlan("sheet.write_table values must be rectangular");
      }
      cells += row.length;
      row.forEach(function (cell) {
        if (
          cell !== null && typeof cell !== "string" && typeof cell !== "boolean" &&
          (!Number.isFinite(cell) || Math.abs(cell) > Number.MAX_SAFE_INTEGER)
        ) {
          invalidSheetPlan("sheet.write_table cell is invalid");
        }
      });
    });
    if (cells > 10000) {
      invalidSheetPlan("sheet.write_table is too large");
    }
  }

  function requireSheetArguments(operation) {
    const args = operation.args;
    const rule = sheetArgumentRules[operation.op];
    const keys = Object.keys(args);
    if (keys.some(function (key) { return rule.allowed.indexOf(key) === -1; })) {
      invalidSheetPlan(`${operation.op} arguments are invalid`);
    }
    if (rule.required.some(function (key) { return !Object.prototype.hasOwnProperty.call(args, key); })) {
      invalidSheetPlan(`${operation.op} arguments are invalid`);
    }
    ["index", "startRow", "startCol"].forEach(function (name) {
      if (Object.prototype.hasOwnProperty.call(args, name) && (!Number.isSafeInteger(args[name]) || args[name] < 1)) {
        invalidSheetPlan(`${operation.op} ${name} is invalid`);
      }
    });
    ["fontSize", "width"].forEach(function (name) {
      if (Object.prototype.hasOwnProperty.call(args, name) && (!Number.isFinite(args[name]) || Math.abs(args[name]) > Number.MAX_SAFE_INTEGER)) {
        invalidSheetPlan(`${operation.op} ${name} is invalid`);
      }
    });
    ["name", "column", "headerShade", "headerFontColor"].forEach(function (name) {
      if (Object.prototype.hasOwnProperty.call(args, name) && typeof args[name] !== "string") {
        invalidSheetPlan(`${operation.op} ${name} is invalid`);
      }
    });
    if (Object.prototype.hasOwnProperty.call(args, "headerBold") && typeof args.headerBold !== "boolean") {
      invalidSheetPlan("sheet.write_table headerBold is invalid");
    }
    if (operation.op === "sheet.write_table") {
      validateSheetTable(args.values);
    }
  }

  function sheetUtf8ByteLength(value) {
    let bytes = 0;
    for (let index = 0; index < value.length; index += 1) {
      const code = value.charCodeAt(index);
      if (code < 0x80) {
        bytes += 1;
      } else if (code < 0x800) {
        bytes += 2;
      } else if (code >= 0xD800 && code <= 0xDBFF && index + 1 < value.length) {
        const next = value.charCodeAt(index + 1);
        if (next >= 0xDC00 && next <= 0xDFFF) {
          bytes += 4;
          index += 1;
        } else {
          bytes += 3;
        }
      } else {
        bytes += 3;
      }
    }
    return bytes;
  }

  function validateSheetJsonLimits(value) {
    const pending = [{value, depth: 0}];
    const seen = [];
    while (pending.length) {
      const current = pending.pop();
      if (current.depth > 64) {
        invalidSheetPlan("Spreadsheet generation plan is nested too deeply");
      }
      if (typeof current.value === "string") {
        if (sheetCharacters(current.value).length > 100000) {
          invalidSheetPlan("Spreadsheet generation plan string is too long");
        }
      } else if (current.value && typeof current.value === "object") {
        if (seen.indexOf(current.value) !== -1) {
          continue;
        }
        seen.push(current.value);
        Object.keys(current.value).forEach(function (key) {
          pending.push({value: key, depth: current.depth + 1});
          pending.push({value: current.value[key], depth: current.depth + 1});
        });
      }
    }
  }

  function validateSheetOperations(plan) {
    if (
      !isObject(plan) || Object.keys(plan).sort().join(",") !== "component,operations" ||
      plan.component !== "spreadsheet" || !Array.isArray(plan.operations) ||
      plan.operations.length === 0 || plan.operations.length > 10000
    ) {
      invalidSheetPlan("Spreadsheet generation plan is invalid");
    }
    let serialized;
    try {
      serialized = JSON.stringify(plan);
    } catch (error) {
      invalidSheetPlan("Spreadsheet generation plan is not JSON-compatible");
    }
    validateSheetJsonLimits(plan);
    if (serialized.length > 2000000 || sheetUtf8ByteLength(serialized) > 2000000) {
      invalidSheetPlan("Spreadsheet generation plan is too large");
    }
    const expected = sheetOperationNames.slice().sort().join(",");
    if (
      Object.keys(sheetOperations).sort().join(",") !== expected ||
      Object.keys(sheetArgumentRules).sort().join(",") !== expected ||
      sheetOperationNames.some(function (name) {
        return typeof sheetOperations[name] !== "function" || !sheetArgumentRules[name];
      })
    ) {
      invalidSheetPlan("Spreadsheet operation catalog is incomplete");
    }
    plan.operations.forEach(function (operation) {
      if (
        !isObject(operation) || Object.keys(operation).sort().join(",") !== "args,op" ||
        typeof operation.op !== "string" || !isObject(operation.args)
      ) {
        invalidSheetPlan("Spreadsheet operation is invalid");
      }
      if (typeof sheetOperations[operation.op] !== "function" || !sheetArgumentRules[operation.op]) {
        invalidSheetPlan(`Unsupported Spreadsheet operation: ${operation.op}`);
      }
      requireSheetArguments(operation);
    });
    return plan.operations;
  }

  function executeSheetOperations(workbook, plan) {
    const operations = validateSheetOperations(plan);
    const context = {sheet: sheetCollectionItem(workbook.Worksheets, 1)};
    operations.forEach(function (operation) {
      sheetOperations[operation.op](workbook, operation.args, context);
    });
    return operations.length;
  }

  function generateSpreadsheetWorkbook(params) {
    const stagedPath = requirePath(params, "stagedPath");
    validateSheetOperations(params.plan);
    const previousAlerts = Application.DisplayAlerts;
    let workbook = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      workbook = Application.Workbooks.Open(stagedPath, 0, false);
      const applied = executeSheetOperations(workbook, params.plan);
      workbook.Save();
      return {path: stagedPath, appliedOperations: applied};
    } catch (error) {
      failure = generationError(error);
      throw failure;
    } finally {
      try {
        if (workbook !== null) {
          workbook.Close(false);
        }
      } catch (closeError) {
        if (failure === null) {
          throw generationError(closeError);
        }
      } finally {
        Application.DisplayAlerts = previousAlerts;
      }
    }
  }

  function requireVisibleWorksheet(workbook) {
    for (let index = 1; index <= workbook.Worksheets.Count; index += 1) {
      if (workbook.Worksheets.Item(index).Visible !== 0) {
        return;
      }
    }
    throw conversionError(
      new Error("Workbook has no visible worksheets"),
      "NO_VISIBLE_WORKSHEETS"
    );
  }

  function saveXlsx(params) {
    const outputPath = requireOutputPath(params);
    const workbook = Application.Workbooks.Add();
    try {
      const sheet = workbook.Worksheets.Item(1);
      sheet.Name = "Phase0";
      const values = [
        ["项目", "数量", "单价", "金额"],
        ["Writer", 2, 120, null],
        ["Presentation", 3, 160, null],
        ["Spreadsheet", 4, 200, null]
      ];
      for (let row = 1; row <= values.length; row += 1) {
        for (let column = 1; column <= values[row - 1].length; column += 1) {
          const value = values[row - 1][column - 1];
          if (value !== null) {
            sheet.Cells.Item(row, column).Value2 = value;
          }
        }
      }
      sheet.Range("D2").Formula = "=B2*C2";
      sheet.Range("D2:D4").FillDown();
      sheet.Range("A1:D1").Font.Bold = true;
      sheet.Range("A1:D1").Interior.Color = 12611584;
      sheet.Range("A1:D4").Borders.LineStyle = 1;
      sheet.UsedRange.Columns.AutoFit();

      const chartObject = sheet.ChartObjects().Add(320, 40, 480, 260);
      chartObject.Chart.SetSourceData(sheet.Range("A1:D4"));
      chartObject.Chart.ChartType = 51;
      chartObject.Chart.HasTitle = true;
      chartObject.Chart.ChartTitle.Text = "WPSComposer Phase 0";

      workbook.SaveAs(outputPath, 51);
      return {
        path: outputPath,
        capabilities: {
          "spreadsheet.create": {
            classification: "native",
            detail: "Workbooks.Add"
          },
          "spreadsheet.values": {
            classification: "native",
            detail: "Cells.Item.Value2"
          },
          "spreadsheet.formulas": {
            classification: "native",
            detail: "Range.Formula and FillDown"
          },
          "spreadsheet.formatting": {
            classification: "native",
            detail: "Font, Interior, Borders, AutoFit"
          },
          "spreadsheet.chart": {
            classification: "native",
            detail: "ChartObjects.Add and SetSourceData"
          },
          "spreadsheet.save_xlsx": {
            classification: "native",
            detail: "Workbook.SaveAs format 51"
          },
          "spreadsheet.template_resolution": {
            classification: "mapped",
            detail: "Workbooks.Add uses the WPS-native default workbook"
          }
        }
      };
    } finally {
      workbook.Close(false);
    }
  }

  function probe() {
    return {
      component: "spreadsheet",
      applicationName: Application.Name,
      capabilities: {
        "spreadsheet.workbooks": {
          classification: "native",
          detail: String(Application.Workbooks.Count)
        },
        "spreadsheet.range": {
          classification: "native",
          detail: "Worksheet.Range"
        },
        "spreadsheet.charts": {
          classification: "native",
          detail: "Worksheet.ChartObjects"
        }
      }
    };
  }

  function convertWorkbookPdf(params) {
    const sourcePath = requirePath(params, "sourcePath");
    const outputPath = requireOutputPath(params);
    const previousAlerts = Application.DisplayAlerts;
    let workbook = null;
    let failure = null;
    try {
      Application.DisplayAlerts = false;
      workbook = Application.Workbooks.Open(sourcePath, 0, true);
      requireVisibleWorksheet(workbook);
      workbook.ExportAsFixedFormat(0, outputPath);
      return {path: outputPath};
    } catch (error) {
      failure = conversionError(error);
      throw failure;
    } finally {
      try {
        if (workbook !== null) {
          workbook.Close(false);
        }
      } catch (closeError) {
        if (failure === null) {
          throw conversionError(closeError);
        }
      } finally {
        Application.DisplayAlerts = previousAlerts;
      }
    }
  }

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_xlsx": saveXlsx,
    "convert_workbook_pdf": convertWorkbookPdf,
    "generate_spreadsheet_workbook": generateSpreadsheetWorkbook
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) {
        throw new Error(`Spreadsheet rejects method ${command.method}`);
      }
      return handler(command.params || {});
    }
  };
}());
