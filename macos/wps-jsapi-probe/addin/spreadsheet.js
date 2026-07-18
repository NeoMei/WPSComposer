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
    "convert_workbook_pdf": convertWorkbookPdf
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
