(function () {
  "use strict";

  const capabilities = {};

  function record(name, classification, detail) {
    capabilities[name] = {classification, detail: detail || ""};
  }

  function attempt(name, action, classification) {
    try {
      const value = action();
      record(name, classification || "native", "");
      return value;
    } catch (error) {
      record(
        name,
        "unsupported",
        String(error && error.message ? error.message : error)
      );
      return null;
    }
  }

  function requirePath(params, name) {
    const value = params[name];
    if (typeof value !== "string" || value.length === 0) {
      throw new Error(`${name} must be a non-empty path`);
    }
    return value;
  }

  function typeParagraph(selection, text, size, bold) {
    selection.Font.Name = "PingFang SC";
    attempt("writer.font_ascii", function () {
      selection.Font.NameAscii = "Helvetica Neue";
    }, "mapped");
    selection.Font.Size = size;
    selection.Font.Bold = bold ? -1 : 0;
    selection.Font.Italic = 0;
    selection.ParagraphFormat.SpaceAfter = size >= 18 ? 8 : 5;
    selection.TypeText(text);
    selection.TypeParagraph();
  }

  function addTable(document, selection) {
    const table = document.Tables.Add(selection.Range, 3, 2);
    const values = [
      ["项目", "状态"],
      ["Writer JSAPI", "通过"],
      ["PDF 导出", "通过"]
    ];
    for (let row = 1; row <= values.length; row += 1) {
      for (let column = 1; column <= values[row - 1].length; column += 1) {
        table.Cell(row, column).Range.Text = values[row - 1][column - 1];
      }
    }
    table.Rows.Item(1).Range.Font.Bold = -1;
    selection.SetRange(document.Content.End, document.Content.End);
    selection.TypeParagraph();
    record("writer.tables", "native", "Document.Tables.Add");
  }

  function addImage(document, selection, imagePath) {
    selection.SetRange(document.Content.End, document.Content.End);
    const image = document.InlineShapes.AddPicture(
      imagePath,
      false,
      true,
      selection.Range
    );
    image.Width = 72;
    image.Height = 72;
    selection.SetRange(document.Content.End, document.Content.End);
    selection.TypeParagraph();
    record("writer.images", "native", "Document.InlineShapes.AddPicture");
  }

  function renderDocument(document, mode, imagePath) {
    document.PageSetup.TopMargin = 54;
    document.PageSetup.BottomMargin = 54;
    document.PageSetup.LeftMargin = 64;
    document.PageSetup.RightMargin = 64;
    const selection = Application.Selection;
    typeParagraph(
      selection,
      mode === "pdf"
        ? "WPSComposer Obsidian Reading"
        : "WPSComposer Writer Smoke",
      24,
      true
    );
    typeParagraph(selection, "macOS WPS JSAPI 可行性验证", 12, false);
    typeParagraph(selection, "能力概览", 18, true);
    typeParagraph(
      selection,
      "正文由 WPS Writer 在 macOS 上排版并序列化。",
      11,
      false
    );
    if (mode === "pdf") {
      typeParagraph(selection, "```python", 10, true);
      typeParagraph(selection, "print('Rendered by Mac WPS')", 10, false);
      typeParagraph(selection, "```", 10, true);
      typeParagraph(
        selection,
        "> [!TIP] 这是 Obsidian 风格提示块。",
        11,
        false
      );
      typeParagraph(selection, "☑ WPS 排版", 11, false);
      typeParagraph(selection, "☐ Phase 1 集成", 11, false);
    }
    addTable(document, selection);
    addImage(document, selection, imagePath);
  }

  function saveDocx(params) {
    const outputPath = requirePath(params, "outputPath");
    const imagePath = requirePath(params, "imagePath");
    const document = Application.Documents.Add();
    try {
      renderDocument(document, "docx", imagePath);
      document.SaveAs2(outputPath, 12);
      record("writer.save_docx", "native", "Document.SaveAs2 format 12");
      return {path: outputPath, capabilities};
    } finally {
      document.Close(0);
    }
  }

  function savePdf(params) {
    const outputPath = requirePath(params, "outputPath");
    const imagePath = requirePath(params, "imagePath");
    const document = Application.Documents.Add();
    try {
      renderDocument(document, "pdf", imagePath);
      document.ExportAsFixedFormat(outputPath, 17, false, 0, 0);
      record(
        "writer.export_pdf",
        "native",
        "Document.ExportAsFixedFormat format 17"
      );
      return {
        path: outputPath,
        preset: "obsidian-reading",
        capabilities
      };
    } finally {
      document.Close(0);
    }
  }

  function probe() {
    attempt("writer.documents", function () {
      return Application.Documents.Count;
    });
    attempt("writer.font_enumeration", function () {
      return Application.FontNames.Count;
    });
    attempt("writer.template_enumeration", function () {
      return Application.Templates.Count;
    });
    return {
      component: "writer",
      applicationName: Application.Name,
      capabilities
    };
  }

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_docx": saveDocx,
    "smoke_pdf": savePdf
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) {
        throw new Error(`Writer rejects method ${command.method}`);
      }
      return handler(command.params || {});
    }
  };
}());
