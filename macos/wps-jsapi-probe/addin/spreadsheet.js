(function () {
  "use strict";

  function requireOutputPath(params) {
    const outputPath = params.outputPath;
    if (typeof outputPath !== "string" || outputPath.length === 0) {
      throw new Error("outputPath must be a non-empty path");
    }
    return outputPath;
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

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_xlsx": saveXlsx
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
