(function () {
  "use strict";

  const capabilities = {};

  function requirePath(params, name) {
    const value = params[name];
    if (typeof value !== "string" || value.length === 0) {
      throw new Error(`${name} must be a non-empty path`);
    }
    return value;
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

  function invalidSlidePlan(message) {
    throw generationError(new Error(message), "OPERATION_PLAN_INVALID");
  }

  function validateSlideOperations(plan) {
    if (!plan || plan.component !== "presentation" || !Array.isArray(plan.operations) || plan.operations.length === 0) {
      invalidSlidePlan("Presentation generation plan is invalid");
    }
    plan.operations.forEach(function (operation) {
      if (!operation || typeof operation !== "object" || !operation.args || typeof operation.args !== "object") {
        invalidSlidePlan("Presentation operation is invalid");
      }
      if (operation.op === "slide.reset") {
        if (Object.keys(operation.args).length !== 0) {
          invalidSlidePlan("slide.reset arguments are invalid");
        }
        return;
      }
      if (operation.op === "slide.add_title") {
        const keys = Object.keys(operation.args);
        if (
          typeof operation.args.title !== "string" ||
          keys.some(function (key) { return key !== "title"; })
        ) {
          invalidSlidePlan("slide.add_title arguments are invalid");
        }
        return;
      }
      invalidSlidePlan(`Unsupported Presentation operation: ${operation.op}`);
    });
    return plan.operations;
  }

  function executeSlideOperations(presentation, plan) {
    const operations = validateSlideOperations(plan);
    operations.forEach(function (operation) {
      if (operation.op === "slide.reset") {
        while (presentation.Slides.Count > 0) {
          presentation.Slides.Item(1).Delete();
        }
      } else {
        const slide = presentation.Slides.Add(
          presentation.Slides.Count + 1,
          12
        );
        const shape = slide.Shapes.AddTextbox(1, 48, 80, 620, 80);
        shape.TextFrame.TextRange.Text = operation.args.title;
      }
    });
    return operations.length;
  }

  function generatePresentationDeck(params) {
    const stagedPath = requirePath(params, "stagedPath");
    validateSlideOperations(params.plan);
    const previousAlerts = Application.DisplayAlerts;
    let presentation = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      presentation = Application.Presentations.Open(stagedPath, false, false, false);
      const applied = executeSlideOperations(presentation, params.plan);
      presentation.Save();
      return {path: stagedPath, appliedOperations: applied};
    } catch (error) {
      failure = generationError(error);
      throw failure;
    } finally {
      try {
        if (presentation !== null) {
          presentation.Close();
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

  function text(slide, value, left, top, width, height, size, bold) {
    const shape = slide.Shapes.AddTextbox(1, left, top, width, height);
    shape.TextFrame.TextRange.Text = value;
    shape.TextFrame.TextRange.Font.Name = "PingFang SC";
    shape.TextFrame.TextRange.Font.Size = size;
    shape.TextFrame.TextRange.Font.Bold = bold ? -1 : 0;
    return shape;
  }

  function addTableSlide(presentation) {
    const slide = presentation.Slides.Add(
      presentation.Slides.Count + 1,
      12
    );
    text(slide, "能力矩阵", 36, 24, 600, 40, 26, true);
    const shape = slide.Shapes.AddTable(3, 2, 72, 100, 576, 240);
    const values = [
      ["能力", "结果"],
      ["文本与图片", "native"],
      ["表格与保存", "native"]
    ];
    for (let row = 1; row <= 3; row += 1) {
      for (let column = 1; column <= 2; column += 1) {
        shape.Table.Cell(row, column).Shape.TextFrame.TextRange.Text =
          values[row - 1][column - 1];
      }
    }
    shape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Bold = -1;
    shape.Table.Cell(1, 2).Shape.TextFrame.TextRange.Font.Bold = -1;
  }

  function savePptx(params) {
    const outputPath = requirePath(params, "outputPath");
    const imagePath = requirePath(params, "imagePath");
    const presentation = Application.Presentations.Add(-1);
    try {
      let slide = presentation.Slides.Add(1, 12);
      text(slide, "WPSComposer macOS", 48, 80, 620, 80, 32, true);
      text(
        slide,
        "WPS Presentation JSAPI Phase 0",
        48,
        180,
        620,
        48,
        18,
        false
      );

      slide = presentation.Slides.Add(2, 12);
      text(slide, "内容与图片", 36, 24, 600, 40, 26, true);
      text(
        slide,
        "• WPS 创建演示文稿\n• JSAPI 添加形状\n• WPS 保存 PPTX",
        48,
        100,
        300,
        240,
        20,
        false
      );
      slide.Shapes.AddPicture(imagePath, 0, -1, 400, 120, 160, 160);
      addTableSlide(presentation);

      presentation.SaveAs(outputPath, 24);
      capabilities["presentation.create"] = {
        classification: "native",
        detail: "Presentations.Add"
      };
      capabilities["presentation.text"] = {
        classification: "native",
        detail: "Shapes.AddTextbox"
      };
      capabilities["presentation.image"] = {
        classification: "native",
        detail: "Shapes.AddPicture"
      };
      capabilities["presentation.table"] = {
        classification: "native",
        detail: "Shapes.AddTable"
      };
      capabilities["presentation.save_pptx"] = {
        classification: "native",
        detail: "Presentation.SaveAs format 24"
      };
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
        "presentation.presentations": {
          classification: "native",
          detail: String(Application.Presentations.Count)
        },
        "presentation.slides": {
          classification: "native",
          detail: "Slides.Add"
        },
        "presentation.shapes": {
          classification: "native",
          detail: "Shapes collection"
        },
        "presentation.template_resolution": {
          classification: "mapped",
          detail: "Presentations.Add uses the WPS-native default blank presentation"
        }
      }
    };
  }

  function convertPresentationPdf(params) {
    const sourcePath = requirePath(params, "sourcePath");
    const outputPath = requirePath(params, "outputPath");
    const previousAlerts = Application.DisplayAlerts;
    let presentation = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      presentation = Application.Presentations.Open(sourcePath, true, false, false);
      presentation.SaveAs(outputPath, 32);
      return {path: outputPath};
    } catch (error) {
      failure = conversionError(error);
      throw failure;
    } finally {
      try {
        if (presentation !== null) {
          presentation.Close();
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
    "smoke_pptx": savePptx,
    "convert_presentation_pdf": convertPresentationPdf,
    "generate_presentation_deck": generatePresentationDeck
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) {
        throw new Error(`Presentation rejects method ${command.method}`);
      }
      return handler(command.params || {});
    }
  };
}());
