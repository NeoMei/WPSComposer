(function () {
  "use strict";
  const capabilities = {};

  function text(slide, value, left, top, width, height, size, bold) {
    const shape = slide.Shapes.AddTextbox(1, left, top, width, height);
    shape.TextFrame.TextRange.Text = value;
    shape.TextFrame.TextRange.Font.Name = "PingFang SC";
    shape.TextFrame.TextRange.Font.Size = size;
    shape.TextFrame.TextRange.Font.Bold = bold ? -1 : 0;
    return shape;
  }

  function tableSlide(presentation) {
    const slide = presentation.Slides.Add(presentation.Slides.Count + 1, 12);
    text(slide, "能力矩阵", 36, 24, 600, 40, 26, true);
    const shape = slide.Shapes.AddTable(3, 2, 72, 100, 576, 240);
    const values = [["能力", "结果"], ["文本与图片", "native"], ["表格与保存", "native"]];
    for (let row = 1; row <= 3; row += 1) {
      for (let column = 1; column <= 2; column += 1) {
        shape.Table.Cell(row, column).Shape.TextFrame.TextRange.Text = values[row - 1][column - 1];
      }
    }
  }

  function savePptx(params) {
    const outputPath = params.outputPath;
    var presentation = null;
    try {
      presentation = Application.Presentations.Add(-1);
    } catch (e) {
      throw new Error("Presentations.Add failed: " + String(e && e.message ? e.message : e));
    }
    if (!presentation) {
      throw new Error("Presentations.Add returned null");
    }
    try {
      let slide = presentation.Slides.Add(1, 12);
      text(slide, "WPSComposer macOS", 48, 80, 620, 80, 32, true);
      text(slide, "WPS Presentation JSAPI Phase 0", 48, 180, 620, 48, 18, false);

      slide = presentation.Slides.Add(2, 12);
      text(slide, "内容与图片", 36, 24, 600, 40, 26, true);
      text(slide, "• WPS 创建演示文稿\n• JSAPI 添加形状\n• WPS 保存 PPTX", 48, 100, 300, 240, 20, false);
      slide.Shapes.AddPicture(params.imagePath, 0, -1, 400, 120, 160, 160);
      tableSlide(presentation);

      presentation.SaveAs(outputPath, 24);
      capabilities["presentation.create"] = {classification: "native", detail: "Presentations.Add"};
      capabilities["presentation.text"] = {classification: "native", detail: "Shapes.AddTextbox"};
      capabilities["presentation.image"] = {classification: "native", detail: "Shapes.AddPicture"};
      capabilities["presentation.table"] = {classification: "native", detail: "Shapes.AddTable"};
      capabilities["presentation.save_pptx"] = {classification: "native", detail: "Presentation.SaveAs format 24"};
      return {path: outputPath, capabilities};
    } finally {
      presentation.Close();
    }
  }

  function probe() {
    return {
      component: "presentation",
      applicationName: Application.Name,
      capabilities: {
        "presentation.presentations": {classification: "native", detail: String(Application.Presentations.Count)},
        "presentation.slides": {classification: "native", detail: "Slides.Add"},
        "presentation.shapes": {classification: "native", detail: "Shapes collection"}
      }
    };
  }

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_pptx": savePptx
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) throw new Error(`Presentation rejects method ${command.method}`);
      return handler(command.params || {});
    }
  };
}());
