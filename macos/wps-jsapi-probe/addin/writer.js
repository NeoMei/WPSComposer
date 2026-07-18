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

  function isObject(value) {
    return value !== null && typeof value === "object" && !Array.isArray(value);
  }

  function writerStyle(document, name) {
    if (!document.Styles) {
      return null;
    }
    if (typeof document.Styles.Item === "function") {
      return document.Styles.Item(name);
    }
    if (typeof document.Styles === "function") {
      return document.Styles(name);
    }
    return null;
  }

  function writerCollectionItem(collection, index) {
    if (collection && typeof collection.Item === "function") {
      return collection.Item(index);
    }
    if (typeof collection === "function") {
      return collection(index);
    }
    return null;
  }

  function writerEndRange(document) {
    const content = document.Content;
    const end = content && typeof content.End === "number"
      ? Math.max(0, content.End - 1)
      : 0;
    if (typeof document.Range === "function") {
      return document.Range(end, end);
    }
    return content;
  }

  function setWriterValue(target, name, value) {
    if (target !== null && typeof target !== "undefined" && typeof value !== "undefined") {
      target[name] = value;
    }
  }

  function writerColor(value) {
    if (typeof value !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(value)) {
      return value;
    }
    const red = parseInt(value.slice(1, 3), 16);
    const green = parseInt(value.slice(3, 5), 16);
    const blue = parseInt(value.slice(5, 7), 16);
    return red | (green << 8) | (blue << 16);
  }

  function applyWriterFont(font, args) {
    if (!font) {
      return;
    }
    setWriterValue(font, "Name", args.fontName);
    setWriterValue(font, "NameFarEast", args.fontName);
    setWriterValue(font, "NameAscii", args.fontNameAscii || args.fontName);
    setWriterValue(font, "NameOther", args.fontNameAscii || args.fontName);
    setWriterValue(font, "Size", args.fontSize !== undefined ? args.fontSize : args.size);
    setWriterValue(font, "Bold", args.bold === undefined ? undefined : (args.bold ? -1 : 0));
    setWriterValue(font, "Italic", args.italic === undefined ? undefined : (args.italic ? -1 : 0));
    setWriterValue(font, "Underline", args.underline === undefined ? undefined : (args.underline ? 1 : 0));
    setWriterValue(font, "StrikeThrough", args.strikethrough);
    setWriterValue(font, "Color", writerColor(args.color));
  }

  function applyWriterParagraphFormat(format, args) {
    if (!format) {
      return;
    }
    setWriterValue(format, "Alignment", args.align);
    setWriterValue(format, "FirstLineIndent", args.indentFirst);
    setWriterValue(format, "LeftIndent", args.leftIndent);
    setWriterValue(format, "RightIndent", args.rightIndent);
    setWriterValue(format, "LineSpacing", args.lineSpacing);
    if (args.lineSpacingRule !== undefined) {
      const rules = {single: 0, one_and_half: 1, double: 2, at_least: 3, exact: 4, multiple: 5};
      setWriterValue(
        format,
        "LineSpacingRule",
        Object.prototype.hasOwnProperty.call(rules, args.lineSpacingRule)
          ? rules[args.lineSpacingRule]
          : args.lineSpacingRule
      );
    }
    setWriterValue(format, "SpaceBefore", args.spaceBefore);
    setWriterValue(format, "SpaceAfter", args.spaceAfter);
    setWriterValue(format, "KeepWithNext", args.keepWithNext);
    setWriterValue(format, "OutlineLevel", args.outlineLevel);
  }

  function insertWriterText(document, text, styleName, formatting) {
    const insertion = writerEndRange(document);
    const start = typeof insertion.Start === "number" ? insertion.Start : 0;
    if (typeof insertion.InsertAfter === "function") {
      insertion.InsertAfter(text);
      insertion.InsertAfter("\r");
    } else {
      insertion.Text = `${text}\r`;
    }
    let written = insertion;
    if (typeof document.Range === "function") {
      written = document.Range(start, start + String(text).length);
    }
    if (styleName) {
      const style = writerStyle(document, styleName);
      if (style) {
        written.Style = style;
      }
    }
    applyWriterFont(written.Font, formatting || {});
    applyWriterParagraphFormat(written.ParagraphFormat, formatting || {});
    return {range: written, start};
  }

  function resetWriter(document) {
    document.Content.Text = "";
  }

  function configureWriterPage(document, args) {
    const setup = document.PageSetup;
    setup.TopMargin = args.marginTop;
    setup.BottomMargin = args.marginBottom;
    setup.LeftMargin = args.marginLeft;
    setup.RightMargin = args.marginRight;
    setWriterValue(setup, "PageWidth", args.pageWidth);
    setWriterValue(setup, "PageHeight", args.pageHeight);
    setWriterValue(setup, "Orientation", args.landscape === undefined ? undefined : (args.landscape ? 1 : 0));
    if (args.columns !== undefined) {
      setup.TextColumns.SetCount(args.columns);
    }
    const section = writerCollectionItem(document.Sections, 1);
    if (section && args.header !== undefined) {
      writerCollectionItem(section.Headers, 1).Range.Text = args.header;
    }
    if (section && args.footer !== undefined) {
      writerCollectionItem(section.Footers, 1).Range.Text = args.footer;
    }
  }

  function ensureWriterStyles(document, args) {
    args.styles.forEach(function (definition) {
      let style = null;
      try {
        style = writerStyle(document, definition.name);
      } catch (error) {
        style = null;
      }
      if (!style && document.Styles && typeof document.Styles.Add === "function") {
        style = document.Styles.Add(definition.name, definition.type === "character" ? 2 : 1);
      }
      if (!style) {
        return;
      }
      if (definition.basedOn !== undefined) {
        const base = writerStyle(document, definition.basedOn);
        style.BaseStyle = base || definition.basedOn;
      }
      applyWriterFont(style.Font, definition);
      applyWriterParagraphFormat(style.ParagraphFormat, definition);
      if (definition.shading !== undefined && style.Shading) {
        style.Shading.BackgroundPatternColor = writerColor(definition.shading);
      }
      if (definition.leftBorder && style.ParagraphFormat && typeof style.ParagraphFormat.Borders === "function") {
        const border = style.ParagraphFormat.Borders(-2);
        border.LineStyle = 1;
        border.LineWidth = 18;
        border.Color = writerColor(definition.borderColor);
      }
    });
  }

  function addWriterParagraph(document, args) {
    if (!Array.isArray(args.spans) || args.spans.length === 0) {
      insertWriterText(document, args.text, args.style, args);
      return;
    }
    const paragraph = insertWriterText(document, args.text, args.style, args);
    let offset = paragraph.start;
    args.spans.forEach(function (span) {
      if (typeof document.Range === "function") {
        const spanRange = document.Range(offset, offset + span.text.length);
        applyWriterFont(spanRange.Font, span);
        if (span.code) {
          const codeStyle = writerStyle(document, "Verbatim Char");
          if (codeStyle) {
            spanRange.Style = codeStyle;
          }
        } else if (span.link) {
          const linkStyle = writerStyle(document, "Hyperlink");
          if (linkStyle) {
            spanRange.Style = linkStyle;
          }
        }
      }
      offset += span.text.length;
    });
  }

  function addWriterHeading(document, args) {
    insertWriterText(document, args.text, `Heading ${args.level}`, args);
  }

  function addWriterList(document, args) {
    args.items.forEach(function (item, index) {
      const prefix = args.ordered ? `${index + 1}.\t` : `${args.glyph || "\u2022"}\t`;
      const formatting = {
        leftIndent: args.indent === undefined ? 24 : args.indent,
        indentFirst: -(args.indent === undefined ? 24 : args.indent)
      };
      insertWriterText(document, prefix + item, "List Paragraph", formatting);
    });
  }

  function addWriterTable(document, args) {
    const insertion = writerEndRange(document);
    const table = document.Tables.Add(insertion, args.rows, args.cols);
    for (let row = 0; row < args.rows; row += 1) {
      for (let column = 0; column < args.cols; column += 1) {
        const cell = table.Cell(row + 1, column + 1);
        const value = args.data[row] && args.data[row][column] !== undefined
          ? args.data[row][column]
          : "";
        cell.Range.Text = String(value);
        applyWriterFont(cell.Range.Font, {fontSize: args.fontSize});
        const paragraph = cell.Range.ParagraphFormat;
        paragraph.FirstLineIndent = 0;
        paragraph.LeftIndent = 0;
        paragraph.RightIndent = 0;
        paragraph.SpaceBefore = 0;
        paragraph.SpaceAfter = 0;
        paragraph.LineSpacingRule = 0;
        cell.TopPadding = 1.5;
        cell.BottomPadding = 1.5;
        cell.LeftPadding = 4;
        cell.RightPadding = 4;
        cell.VerticalAlignment = 1;
        if (row === 0) {
          cell.Range.Font.Bold = -1;
          setWriterValue(cell.Range.Font, "Color", writerColor(args.headerColor));
          setWriterValue(cell.Shading, "BackgroundPatternColor", writerColor(args.shadeHeader));
        } else if (args.bandedRows && row % 2 === 0) {
          cell.Shading.BackgroundPatternColor = writerColor("#F2F2F2");
        }
        if (Array.isArray(args.alignments) && args.alignments[column]) {
          const alignments = {left: 0, center: 1, right: 2};
          cell.Range.ParagraphFormat.Alignment = alignments[args.alignments[column]];
        }
      }
    }
    if (args.repeatHeader && args.rows > 1) {
      writerCollectionItem(table.Rows, 1).HeadingFormat = -1;
    }
    if (Array.isArray(args.columnWidths)) {
      args.columnWidths.forEach(function (width, index) {
        const tableColumn = writerCollectionItem(table.Columns, index + 1);
        if (tableColumn) {
          tableColumn.Width = width;
        }
      });
    } else if (args.autoFit && typeof table.AutoFitBehavior === "function") {
      table.AutoFitBehavior(1);
    }
    if (typeof table.Borders === "function") {
      for (let borderIndex = -6; borderIndex < 0; borderIndex += 1) {
        const border = table.Borders(borderIndex);
        border.LineStyle = 1;
        border.LineWidth = 2;
        border.Color = writerColor(args.borderColor);
      }
    }
    if (table.Rows) {
      table.Rows.AllowBreakAcrossPages = false;
    }
    const afterTable = writerEndRange(document);
    if (typeof afterTable.InsertAfter === "function") {
      afterTable.InsertAfter("\r");
    }
  }

  function addWriterImage(document, args, resources) {
    const imagePath = resources[args.imageId];
    const insertion = writerEndRange(document);
    if (!args.inline && document.Shapes && typeof document.Shapes.AddPicture === "function") {
      const shape = document.Shapes.AddPicture(imagePath, false, true);
      setWriterValue(shape, "Width", args.width);
      setWriterValue(shape, "Height", args.height);
      setWriterValue(shape, "AlternativeText", args.alt);
      return;
    }
    const shape = document.InlineShapes.AddPicture(imagePath, false, true, insertion);
    if (!shape) {
      throw new Error("Writer image insertion returned no shape");
    }
    if (args.preserveAspect !== undefined) {
      shape.LockAspectRatio = args.preserveAspect ? -1 : 0;
    }
    const naturalWidth = Number(shape.Width) || 0;
    const naturalHeight = Number(shape.Height) || 0;
    if (args.width !== undefined && args.height !== undefined && args.preserveAspect && naturalWidth && naturalHeight) {
      const scale = Math.min(args.width / naturalWidth, args.height / naturalHeight);
      shape.Width = naturalWidth * scale;
      shape.Height = naturalHeight * scale;
    } else {
      setWriterValue(shape, "Width", args.width);
      setWriterValue(shape, "Height", args.height);
    }
    let scale = 1;
    if (args.maxWidth && shape.Width > args.maxWidth) {
      scale = Math.min(scale, args.maxWidth / shape.Width);
    }
    if (args.maxHeight && shape.Height > args.maxHeight) {
      scale = Math.min(scale, args.maxHeight / shape.Height);
    }
    if (scale < 1) {
      shape.Width *= scale;
      shape.Height *= scale;
    }
    setWriterValue(shape, "AlternativeText", args.alt);
    if (shape.Range && shape.Range.ParagraphFormat) {
      shape.Range.ParagraphFormat.Alignment = 1;
      shape.Range.ParagraphFormat.KeepWithNext = -1;
      shape.Range.ParagraphFormat.SpaceAfter = 0;
    }
    const afterImage = writerEndRange(document);
    if (typeof afterImage.InsertAfter === "function") {
      afterImage.InsertAfter("\r");
    }
  }

  function addWriterPageBreak(document) {
    writerEndRange(document).InsertBreak(7);
  }

  function addWriterSection(document) {
    writerEndRange(document).InsertBreak(2);
  }

  function addWriterHorizontalLine(document) {
    const range = writerEndRange(document);
    if (range.ParagraphFormat && typeof range.ParagraphFormat.Borders === "function") {
      const border = range.ParagraphFormat.Borders(-4);
      border.LineStyle = 1;
      border.LineWidth = 6;
    }
    if (typeof range.InsertAfter === "function") {
      range.InsertAfter(" \r");
    }
  }

  function insertWriterToc(document, args) {
    if (args.title) {
      insertWriterText(document, args.title, "Heading 1", {align: 1});
    }
    document.TablesOfContents.Add(writerEndRange(document), true, 1, 3);
  }

  function setWriterPageNumber(document) {
    const section = writerCollectionItem(document.Sections, 1);
    const footer = writerCollectionItem(section.Footers, 1).Range;
    footer.Text = "Page ";
    document.Fields.Add(footer, 33);
  }

  function updateWriterFields(document) {
    if (document.TablesOfContents) {
      for (let index = 1; index <= document.TablesOfContents.Count; index += 1) {
        writerCollectionItem(document.TablesOfContents, index).Update();
      }
    }
    document.Fields.Update();
  }

  const writerOperations = {
    "writer.reset": resetWriter,
    "writer.configure_page": configureWriterPage,
    "writer.ensure_styles": ensureWriterStyles,
    "writer.add_paragraph": addWriterParagraph,
    "writer.add_heading": addWriterHeading,
    "writer.add_list": addWriterList,
    "writer.add_table": addWriterTable,
    "writer.add_image": addWriterImage,
    "writer.add_page_break": addWriterPageBreak,
    "writer.add_section": addWriterSection,
    "writer.add_horizontal_line": addWriterHorizontalLine,
    "writer.insert_toc": insertWriterToc,
    "writer.set_page_number": setWriterPageNumber,
    "writer.update_fields": updateWriterFields
  };

  const writerArgumentRules = {
    "writer.reset": {required: [], allowed: []},
    "writer.configure_page": {required: ["marginTop", "marginBottom", "marginLeft", "marginRight"], allowed: ["marginTop", "marginBottom", "marginLeft", "marginRight", "pageWidth", "pageHeight", "landscape", "columns", "header", "footer"]},
    "writer.ensure_styles": {required: ["styles"], allowed: ["styles"]},
    "writer.add_paragraph": {required: ["text"], allowed: ["text", "style", "spans", "size", "bold", "italic", "color", "align", "indentFirst", "lineSpacing", "lineSpacingRule", "spaceBefore", "spaceAfter", "fontName", "fontNameAscii"]},
    "writer.add_heading": {required: ["text", "level"], allowed: ["text", "level", "size", "bold", "color", "lineSpacing", "lineSpacingRule", "spaceAfter"]},
    "writer.add_list": {required: ["items", "ordered"], allowed: ["items", "ordered", "glyph", "indent"]},
    "writer.add_table": {required: ["rows", "cols", "data"], allowed: ["rows", "cols", "data", "shadeHeader", "headerColor", "fontSize", "columnWidths", "alignments", "bandedRows", "autoFit", "repeatHeader", "borderColor"]},
    "writer.add_image": {required: ["imageId"], allowed: ["imageId", "width", "height", "maxWidth", "maxHeight", "wrap", "inline", "preserveAspect", "alt"]},
    "writer.add_page_break": {required: [], allowed: []},
    "writer.add_section": {required: [], allowed: []},
    "writer.add_horizontal_line": {required: [], allowed: []},
    "writer.insert_toc": {required: ["title"], allowed: ["title"]},
    "writer.set_page_number": {required: [], allowed: []},
    "writer.update_fields": {required: [], allowed: []}
  };

  function requireWriterArguments(operation) {
    const args = operation.args;
    const rule = writerArgumentRules[operation.op];
    const keys = Object.keys(args);
    if (keys.some(function (key) { return rule.allowed.indexOf(key) === -1; })) {
      invalidWriterPlan(`${operation.op} arguments are invalid`);
    }
    if (rule.required.some(function (key) { return !Object.prototype.hasOwnProperty.call(args, key); })) {
      invalidWriterPlan(`${operation.op} arguments are invalid`);
    }
    const numberFields = [
      "marginTop", "marginBottom", "marginLeft", "marginRight", "pageWidth",
      "pageHeight", "size", "indentFirst", "lineSpacing", "spaceBefore",
      "spaceAfter", "fontSize", "indent", "width", "height", "maxWidth",
      "maxHeight"
    ];
    const integerFields = ["columns", "align", "level", "rows", "cols", "wrap"];
    const booleanFields = [
      "landscape", "bold", "italic", "ordered", "bandedRows", "autoFit",
      "repeatHeader", "inline", "preserveAspect"
    ];
    const stringFields = [
      "header", "footer", "text", "style", "color", "lineSpacingRule",
      "fontName", "fontNameAscii", "glyph", "shadeHeader", "headerColor",
      "borderColor", "imageId", "title"
    ];
    numberFields.forEach(function (name) {
      if (Object.prototype.hasOwnProperty.call(args, name) && (!Number.isFinite(args[name]) || Math.abs(args[name]) > Number.MAX_SAFE_INTEGER)) {
        invalidWriterPlan(`${operation.op} ${name} is invalid`);
      }
    });
    integerFields.forEach(function (name) {
      if (Object.prototype.hasOwnProperty.call(args, name) && !Number.isSafeInteger(args[name])) {
        invalidWriterPlan(`${operation.op} ${name} is invalid`);
      }
    });
    booleanFields.forEach(function (name) {
      if (Object.prototype.hasOwnProperty.call(args, name) && typeof args[name] !== "boolean") {
        invalidWriterPlan(`${operation.op} ${name} is invalid`);
      }
    });
    stringFields.forEach(function (name) {
      if (Object.prototype.hasOwnProperty.call(args, name) && (typeof args[name] !== "string" || writerCharacterLength(args[name]) > 100000)) {
        invalidWriterPlan(`${operation.op} ${name} is invalid`);
      }
    });
    if (Object.prototype.hasOwnProperty.call(args, "alt") && args.alt !== null && typeof args.alt !== "string") {
      invalidWriterPlan("writer.add_image alt is invalid");
    }
    if (operation.op === "writer.add_heading" && (args.level < 1 || args.level > 6)) {
      invalidWriterPlan("writer.add_heading level is invalid");
    }
    if (operation.op === "writer.ensure_styles") {
      validateWriterStyles(args.styles);
    }
    if (operation.op === "writer.add_paragraph" && Object.prototype.hasOwnProperty.call(args, "spans")) {
      validateWriterSpans(args.spans);
    }
    if (operation.op === "writer.add_list" && (!Array.isArray(args.items) || args.items.some(function (item) { return typeof item !== "string"; }))) {
      invalidWriterPlan("writer.add_list items are invalid");
    }
    if (operation.op === "writer.add_table") {
      validateWriterTable(args);
    }
    if (Object.prototype.hasOwnProperty.call(args, "columnWidths") && (!Array.isArray(args.columnWidths) || args.columnWidths.some(function (value) { return !Number.isFinite(value); }))) {
      invalidWriterPlan("writer.add_table columnWidths are invalid");
    }
    if (Object.prototype.hasOwnProperty.call(args, "alignments") && (!Array.isArray(args.alignments) || args.alignments.some(function (value) { return typeof value !== "string"; }))) {
      invalidWriterPlan("writer.add_table alignments are invalid");
    }
  }

  function validateWriterNestedObject(value, required, fields, label) {
    if (!isObject(value)) {
      invalidWriterPlan(`${label} is invalid`);
    }
    const keys = Object.keys(value);
    if (keys.some(function (key) { return !Object.prototype.hasOwnProperty.call(fields, key); })) {
      invalidWriterPlan(`${label} is invalid`);
    }
    if (required.some(function (key) { return !Object.prototype.hasOwnProperty.call(value, key); })) {
      invalidWriterPlan(`${label} is invalid`);
    }
    keys.forEach(function (key) {
      const expected = fields[key];
      const item = value[key];
      const valid = expected === "string"
        ? typeof item === "string"
        : expected === "number"
          ? Number.isFinite(item) && Math.abs(item) <= Number.MAX_SAFE_INTEGER
          : expected === "integer"
            ? Number.isSafeInteger(item)
            : expected === "boolean"
              ? typeof item === "boolean"
              : expected === "nullable-string"
                ? item === null || typeof item === "string"
                : false;
      if (!valid) {
        invalidWriterPlan(`${label}.${key} is invalid`);
      }
    });
  }

  function validateWriterStyles(styles) {
    if (!Array.isArray(styles)) {
      invalidWriterPlan("writer.ensure_styles styles are invalid");
    }
    const fields = {
      name: "string", type: "string", basedOn: "string", fontName: "string",
      fontNameAscii: "string", fontSize: "number", bold: "boolean",
      italic: "boolean", underline: "boolean", strikethrough: "boolean",
      color: "string", align: "integer", indentFirst: "number",
      leftIndent: "number", rightIndent: "number", lineSpacing: "number",
      lineSpacingRule: "string", spaceBefore: "number", spaceAfter: "number",
      shading: "string", leftBorder: "boolean", borderColor: "string",
      keepWithNext: "boolean", outlineLevel: "integer"
    };
    styles.forEach(function (style) {
      validateWriterNestedObject(style, ["name"], fields, "writer style");
    });
  }

  function validateWriterSpans(spans) {
    if (!Array.isArray(spans)) {
      invalidWriterPlan("writer.add_paragraph spans are invalid");
    }
    const fields = {
      text: "string", bold: "boolean", italic: "boolean",
      strikethrough: "boolean", code: "boolean", link: "nullable-string",
      linkTitle: "nullable-string"
    };
    spans.forEach(function (span) {
      validateWriterNestedObject(span, ["text"], fields, "writer span");
    });
  }

  function validateWriterTable(args) {
    if (!Number.isSafeInteger(args.rows) || !Number.isSafeInteger(args.cols) || !Array.isArray(args.data)) {
      invalidWriterPlan("writer.add_table arguments are invalid");
    }
    if (args.rows <= 0 || args.cols <= 0 || args.data.length !== args.rows) {
      invalidWriterPlan("writer.add_table dimensions are invalid");
    }
    let width = null;
    let cells = 0;
    args.data.forEach(function (row) {
      if (!Array.isArray(row) || row.length !== args.cols || (width !== null && row.length !== width)) {
        invalidWriterPlan("writer.add_table data is invalid");
      }
      width = row.length;
      cells += row.length;
      row.forEach(function (cell) {
        if (cell !== null && typeof cell !== "string" && typeof cell !== "boolean" && (!Number.isFinite(cell) || Math.abs(cell) > Number.MAX_SAFE_INTEGER)) {
          invalidWriterPlan("writer.add_table cell is invalid");
        }
      });
    });
    if (Math.max(cells, Math.max(0, args.rows) * Math.max(0, args.cols)) > 10000) {
      invalidWriterPlan("writer.add_table is too large");
    }
  }

  function writerUtf8ByteLength(value) {
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

  function writerCharacterLength(value) {
    let characters = 0;
    for (let index = 0; index < value.length; index += 1) {
      const code = value.charCodeAt(index);
      if (code >= 0xD800 && code <= 0xDBFF && index + 1 < value.length) {
        const next = value.charCodeAt(index + 1);
        if (next >= 0xDC00 && next <= 0xDFFF) {
          index += 1;
        }
      }
      characters += 1;
    }
    return characters;
  }

  function validateWriterJsonLimits(value) {
    const pending = [{value, depth: 0}];
    const seen = [];
    while (pending.length) {
      const current = pending.pop();
      if (current.depth > 64) {
        invalidWriterPlan("Writer generation plan is nested too deeply");
      }
      if (typeof current.value === "string") {
        if (writerCharacterLength(current.value) > 100000) {
          invalidWriterPlan("Writer generation plan string is too long");
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

  function validateWriterResource(imageId, resources) {
    if (!/^[A-Za-z0-9][A-Za-z0-9_-]*$/.test(imageId)) {
      invalidWriterPlan("Writer imageId must be a logical identifier");
    }
    if (!isObject(resources) || !Object.prototype.hasOwnProperty.call(resources, imageId)) {
      invalidWriterPlan(`Missing Writer resource: ${imageId}`);
    }
    const resource = resources[imageId];
    if (typeof resource !== "string" || resource.indexOf("..") !== -1 || resource.indexOf("\\") !== -1 || resource.indexOf("%") !== -1) {
      invalidWriterPlan(`Unsafe Writer resource: ${imageId}`);
    }
    let parsed;
    try {
      parsed = new URL(resource);
    } catch (error) {
      invalidWriterPlan(`Unsafe Writer resource: ${imageId}`);
    }
    if (parsed.protocol !== "http:" || parsed.hostname !== "127.0.0.1" || parsed.port !== "3889" || parsed.username || parsed.password || parsed.search || parsed.hash) {
      invalidWriterPlan(`Unsafe Writer resource: ${imageId}`);
    }
    const segments = parsed.pathname.split("/").filter(Boolean);
    if (segments.length !== 1) {
      invalidWriterPlan(`Unsafe Writer resource: ${imageId}`);
    }
  }

  function validateWriterOperations(plan, resources) {
    if (!isObject(plan) || Object.keys(plan).sort().join(",") !== "component,operations" || plan.component !== "writer" || !Array.isArray(plan.operations) || plan.operations.length === 0 || plan.operations.length > 10000) {
      invalidWriterPlan("Writer generation plan is invalid");
    }
    let serialized;
    try {
      serialized = JSON.stringify(plan);
    } catch (error) {
      invalidWriterPlan("Writer generation plan is not JSON-compatible");
    }
    if (serialized.length > 2000000) {
      invalidWriterPlan("Writer generation plan is too large");
    }
    validateWriterJsonLimits(plan);
    if (writerUtf8ByteLength(serialized) > 2000000) {
      invalidWriterPlan("Writer generation plan is too large");
    }
    const usedResources = {};
    plan.operations.forEach(function (operation) {
      if (!isObject(operation) || Object.keys(operation).sort().join(",") !== "args,op" || typeof operation.op !== "string" || !isObject(operation.args)) {
        invalidWriterPlan("Writer operation is invalid");
      }
      if (typeof writerOperations[operation.op] !== "function" || !writerArgumentRules[operation.op]) {
        invalidWriterPlan(`Unsupported Writer operation: ${operation.op}`);
      }
      requireWriterArguments(operation);
      if (operation.op === "writer.add_image") {
        validateWriterResource(operation.args.imageId, resources);
        usedResources[operation.args.imageId] = true;
      }
    });
    Object.keys(resources || {}).forEach(function (resourceId) {
      if (!usedResources[resourceId]) {
        invalidWriterPlan(`Unused Writer resource: ${resourceId}`);
      }
    });
    return plan.operations;
  }

  function executeWriterOperations(document, plan, resources) {
    const operations = validateWriterOperations(plan, resources);
    operations.forEach(function (operation) {
      writerOperations[operation.op](document, operation.args, resources || {});
    });
    return operations.length;
  }

  function generateWriterDocument(params) {
    const stagedPath = requirePath(params, "stagedPath");
    validateWriterOperations(params.plan, params.resources || {});
    const previousAlerts = Application.DisplayAlerts;
    let document = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      document = Application.Documents.Open(stagedPath, false, false);
      const applied = executeWriterOperations(document, params.plan, params.resources || {});
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
