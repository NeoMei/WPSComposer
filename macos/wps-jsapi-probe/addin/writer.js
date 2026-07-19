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
      record(name, "unsupported", String(error && error.message ? error.message : error));
      return null;
    }
  }

  function typeParagraph(selection, text, size, bold) {
    selection.Font.Name = "PingFang SC";
    selection.Font.NameAscii = "Helvetica Neue";
    selection.Font.Size = size;
    selection.Font.Bold = bold;
    selection.TypeText(text);
    selection.TypeParagraph();
  }

  function addTable(document, selection) {
    const table = document.Tables.Add(selection.Range, 3, 2);
    const values = [["项目", "状态"], ["Writer JSAPI", "通过"], ["PDF 导出", "通过"]];
    for (let row = 1; row <= values.length; row += 1) {
      for (let column = 1; column <= values[row - 1].length; column += 1) {
        table.Cell(row, column).Range.Text = values[row - 1][column - 1];
      }
    }
    table.Rows.Item(1).Range.Font.Bold = true;
    selection.SetRange(document.Content.End, document.Content.End);
    selection.TypeParagraph();
  }

  function addImage(document, selection, imagePath) {
    selection.SetRange(document.Content.End, document.Content.End);
    var image = null;
    try {
      image = document.InlineShapes.AddPicture(imagePath, false, true, selection.Range);
    } catch (e) {
      record("writer.images", "unsupported", String(e && e.message ? e.message : e));
    }
    if (image) {
      image.Width = 72;
      image.Height = 72;
      record("writer.images", "native", "Document.InlineShapes.AddPicture");
    }
    selection.SetRange(document.Content.End, document.Content.End);
    selection.TypeParagraph();
  }

  function buildDocument(mode, imagePath) {
    const document = Application.Documents.Add();
    const selection = Application.Selection;
    typeParagraph(selection, mode === "pdf" ? "WPSComposer Obsidian Reading" : "WPSComposer Writer Smoke", 24, true);
    typeParagraph(selection, "macOS WPS JSAPI 可行性验证", 12, false);
    typeParagraph(selection, "## 能力概览", 18, true);
    typeParagraph(selection, "正文由 WPS Writer 在 macOS 上排版并序列化。", 11, false);
    if (mode === "pdf") {
      typeParagraph(selection, "```python", 10, true);
      typeParagraph(selection, "print('Rendered by Mac WPS')", 10, false);
      typeParagraph(selection, "```", 10, true);
      typeParagraph(selection, "> [!TIP] 这是 Obsidian 风格提示块。", 11, false);
      typeParagraph(selection, "☑ WPS 排版", 11, false);
      typeParagraph(selection, "☐ Phase 1 集成", 11, false);
    }
    addTable(document, selection);
    addImage(document, selection, imagePath);
    return document;
  }

  function saveDocx(params) {
    const outputPath = params.outputPath;
    const document = buildDocument("docx", params.imagePath);
    try {
      var containerDir = params.containerDir || "/tmp";
      var fileName = outputPath.split("/").pop();
      var containerPath = containerDir + "/" + fileName;
      document.SaveAs2(containerPath, 16);
      record("writer.save_docx", "mapped", "SaveAs2 to sandbox container, host copies out");
      return {path: outputPath, containerPath: containerPath, capabilities: capabilities};
    } finally {
      document.Close(0);
    }
  }

  function savePdf(params) {
    const outputPath = params.outputPath;
    const document = buildDocument("pdf", params.imagePath);
    try {
      document.ExportAsFixedFormat(outputPath, 17, false, 0, 0);
      record("writer.export_pdf", "native", "Document.ExportAsFixedFormat format 17");
      return {path: outputPath, preset: "obsidian-reading", capabilities};
    } finally {
      document.Close(0);
    }
  }

  function probe() {
    attempt("writer.documents_add", function () { return Application.Documents.Count; });
    attempt("writer.font_enumeration", function () { return Application.FontNames.Count; });
    attempt("writer.template_enumeration", function () { return Application.Templates.Count; });
    record("writer.tables", "native", "Document.Tables.Add");
    return {component: "writer", applicationName: Application.Name, capabilities};
  }

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_docx": saveDocx,
    "smoke_pdf": savePdf
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) throw new Error(`Writer rejects method ${command.method}`);
      return handler(command.params || {});
    }
  };
}());
