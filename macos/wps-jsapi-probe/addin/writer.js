(function () {
  "use strict";

  const WRITER_IMAGE_URL = "http://127.0.0.1:3889/fixture.png";
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

  function invalidWriterPlan(message) {
    throw generationError(new Error(message), "OPERATION_PLAN_INVALID");
  }

  function validateWriterOperations(plan) {
    if (!plan || plan.component !== "writer" || !Array.isArray(plan.operations) || plan.operations.length === 0) {
      invalidWriterPlan("Writer generation plan is invalid");
    }
    plan.operations.forEach(function (operation) {
      if (!operation || typeof operation !== "object" || !operation.args || typeof operation.args !== "object") {
        invalidWriterPlan("Writer operation is invalid");
      }
      if (operation.op === "writer.reset") {
        if (Object.keys(operation.args).length !== 0) {
          invalidWriterPlan("writer.reset arguments are invalid");
        }
        return;
      }
      if (operation.op === "writer.add_paragraph") {
        const keys = Object.keys(operation.args);
        if (
          typeof operation.args.text !== "string" ||
          keys.some(function (key) { return key !== "text" && key !== "style"; }) ||
          (Object.prototype.hasOwnProperty.call(operation.args, "style") && typeof operation.args.style !== "string")
        ) {
          invalidWriterPlan("writer.add_paragraph arguments are invalid");
        }
        return;
      }
      invalidWriterPlan(`Unsupported Writer operation: ${operation.op}`);
    });
    return plan.operations;
  }

  function executeWriterOperations(document, plan) {
    const operations = validateWriterOperations(plan);
    operations.forEach(function (operation) {
      if (operation.op === "writer.reset") {
        document.Content.Text = "";
      } else {
        document.Content.InsertAfter(operation.args.text);
      }
    });
    return operations.length;
  }

  function generateWriterDocument(params) {
    const stagedPath = requirePath(params, "stagedPath");
    validateWriterOperations(params.plan);
    const previousAlerts = Application.DisplayAlerts;
    let document = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      document = Application.Documents.Open(stagedPath, false, false);
      const applied = executeWriterOperations(document, params.plan);
      document.Save();
      return {path: stagedPath, appliedOperations: applied};
    } catch (error) {
      failure = generationError(error);
      throw failure;
    } finally {
      try {
        if (document !== null) {
          document.Close(0);
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

  function requireImageUrl(params) {
    if (params.imageUrl !== WRITER_IMAGE_URL) {
      throw new Error("imageUrl must be the fixed Writer loopback asset");
    }
    record("writer.image_source", "mapped", "Served the allowed image through the loopback add-in origin");
    return params.imageUrl;
  }

  function waitForFileAfterSave(outputPath, action) {
    return new Promise(function (resolve, reject) {
      let settled = false;
      let timer = null;

      function finish(error) {
        if (settled) {
          return;
        }
        settled = true;
        if (timer !== null) {
          clearTimeout(timer);
        }
        try {
          Application.ApiEvent.RemoveApiEventListener("FileAfterSave");
        } catch (cleanupError) {
          if (!error) {
            error = cleanupError;
          }
        }
        if (error) {
          reject(error);
        } else {
          resolve();
        }
      }

      Application.ApiEvent.AddApiEventListener("FileAfterSave", function (
        document,
        filePath
      ) {
        if (typeof filePath === "string" && filePath !== outputPath) {
          return;
        }
        finish(null);
      });
      timer = setTimeout(function () {
        finish(new Error(`Timed out waiting for WPS to save ${outputPath}`));
      }, 15000);
      try {
        action();
      } catch (error) {
        finish(error);
      }
    });
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
    let image = document.InlineShapes.AddPicture(
      imagePath,
      false,
      true,
      selection.Range
    );
    if (image === null || typeof image === "undefined") {
      const imageCount = document.InlineShapes.Count;
      if (imageCount >= 1) {
        image = document.InlineShapes.Item(imageCount);
        record("writer.image_return", "mapped", "Recovered the inserted InlineShape from the collection");
        record("writer.image_api", "native", "Document.InlineShapes.AddPicture");
      } else {
        const previousShapeCount = document.Shapes.Count;
        image = document.Shapes.AddPicture(
          imagePath,
          false,
          true,
          0,
          0,
          72,
          72,
          selection.Range
        );
        const shapeCount = document.Shapes.Count;
        if ((image === null || typeof image === "undefined") && shapeCount > previousShapeCount) {
          image = document.Shapes.Item(shapeCount);
        }
        if (image === null || typeof image === "undefined") {
          throw new Error("Writer image insertion is unsupported by InlineShapes and Shapes");
        }
        record("writer.image_return", "mapped", "InlineShapes returned null; Shapes returned the inserted image");
        record("writer.image_api", "mapped", "Used Document.Shapes.AddPicture fallback");
      }
    } else {
      record("writer.image_return", "native", "InlineShapes.AddPicture returned InlineShape");
      record("writer.image_api", "native", "Document.InlineShapes.AddPicture");
    }
    image.Width = 72;
    image.Height = 72;
    selection.SetRange(document.Content.End, document.Content.End);
    selection.TypeParagraph();
    record(
      "writer.images",
      capabilities["writer.image_api"].classification,
      capabilities["writer.image_api"].detail
    );
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

  async function saveDocx(params) {
    const outputPath = requirePath(params, "outputPath");
    const imagePath = requireImageUrl(params);
    const document = Application.Documents.Add();
    try {
      renderDocument(document, "docx", imagePath);
      await waitForFileAfterSave(outputPath, function () {
        document.SaveAs2(outputPath, 12);
      });
      record("writer.save_docx", "native", "Document.SaveAs2 format 12");
      return {path: outputPath, capabilities};
    } finally {
      document.Close(0);
    }
  }

  async function savePdf(params) {
    const outputPath = requirePath(params, "outputPath");
    const sourcePath = requirePath(params, "sourcePath");
    const imagePath = requireImageUrl(params);
    const document = Application.Documents.Add();
    try {
      renderDocument(document, "pdf", imagePath);
      await waitForFileAfterSave(sourcePath, function () {
        document.SaveAs2(sourcePath, 12);
      });
      await waitForFileAfterSave(outputPath, function () {
        document.ExportAsFixedFormat(outputPath, 17, false, 0, 0);
      });
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

  async function convertWriterPdf(params) {
    const sourcePath = requirePath(params, "sourcePath");
    const outputPath = requirePath(params, "outputPath");
    const previousAlerts = Application.DisplayAlerts;
    let document = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      document = Application.Documents.Open(sourcePath, false, true);
      await waitForFileAfterSave(outputPath, function () {
        document.ExportAsFixedFormat(outputPath, 17, false, 0, 0);
      });
      return {path: outputPath};
    } catch (error) {
      failure = conversionError(error);
      throw failure;
    } finally {
      try {
        if (document !== null) {
          document.Close(0);
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
    "smoke_pdf": savePdf,
    "convert_writer_pdf": convertWriterPdf,
    "generate_writer_document": generateWriterDocument
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
