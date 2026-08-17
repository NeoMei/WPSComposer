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

  function isObject(value) {
    return value !== null && typeof value === "object" && !Array.isArray(value);
  }

  function hasOwn(object, key) {
    return Object.prototype.hasOwnProperty.call(object, key);
  }

  function slideCollectionItem(collection, index) {
    if (collection && typeof collection.Item === "function") {
      return collection.Item(index);
    }
    if (typeof collection === "function") {
      return collection(index);
    }
    invalidSlidePlan("Presentation slide collection is unavailable");
  }

  function slideColor(value) {
    const red = parseInt(value.slice(1, 3), 16);
    const green = parseInt(value.slice(3, 5), 16);
    const blue = parseInt(value.slice(5, 7), 16);
    return red | (green << 8) | (blue << 16);
  }

  function applySlideFill(fill, color) {
    if (!fill || color === undefined) {
      return;
    }
    if (!fill.ForeColor) {
      fill.ForeColor = {};
    }
    fill.ForeColor.RGB = slideColor(color);
    fill.Visible = 1;
    if (typeof fill.Solid === "function") {
      fill.Solid();
    }
  }

  function presetFont(context, role, fallback) {
    if (context.preset && context.preset.fonts[role]) {
      return context.preset.fonts[role];
    }
    return fallback;
  }

  function styleSlideText(range, font, size, color, bold, alignment) {
    if (!range.Font) {
      range.Font = {};
    }
    range.Font.Name = font.family;
    range.Font.Size = size;
    range.Font.Bold = bold ? -1 : 0;
    if (!range.Font.Color) {
      range.Font.Color = {};
    }
    range.Font.Color.RGB = slideColor(color);
    if (!range.ParagraphFormat) {
      range.ParagraphFormat = {};
    }
    if (alignment !== undefined) {
      range.ParagraphFormat.Alignment = alignment;
    }
  }

  function newSlide(presentation, context) {
    const slide = presentation.Slides.Add(presentation.Slides.Count + 1, 12);
    if (context.preset && slide.Background && slide.Background.Fill) {
      slide.FollowMasterBackground = false;
      applySlideFill(slide.Background.Fill, context.preset.colors.background);
    }
    return slide;
  }

  function addSlideText(slide, value, left, top, width, height, font, size, color, bold, alignment) {
    const shape = slide.Shapes.AddTextbox(1, left, top, width, height);
    const range = shape.TextFrame.TextRange;
    range.Text = value;
    styleSlideText(range, font, size, color, bold, alignment);
    return shape;
  }

  function resetSlides(presentation) {
    while (presentation.Slides.Count > 0) {
      slideCollectionItem(presentation.Slides, 1).Delete();
    }
  }

  function setSlideSize(presentation, args, context) {
    presentation.PageSetup.SlideWidth = args.width;
    presentation.PageSetup.SlideHeight = args.height;
    context.width = args.width;
    context.height = args.height;
  }

  function applySlidePreset(presentation, args, context) {
    context.preset = args.preset;
    if (
      presentation.SlideMaster && presentation.SlideMaster.Background &&
      presentation.SlideMaster.Background.Fill
    ) {
      applySlideFill(
        presentation.SlideMaster.Background.Fill,
        args.preset.colors.background
      );
    }
  }

  function addTitleSlide(presentation, args, context) {
    const slide = newSlide(presentation, context);
    const margin = context.preset && context.preset.spacing
      ? context.preset.spacing.margin : 80;
    const titleFont = presetFont(
      context,
      "title",
      {family: "Arial", size: 40, color: "#000000"}
    );
    const subtitleFont = presetFont(
      context,
      "subtitle",
      {family: "Arial", size: 20, color: "#646464"}
    );
    addSlideText(
      slide,
      args.title,
      margin,
      context.height * 0.28,
      context.width - 2 * margin,
      context.height * 0.18,
      titleFont,
      args.titleSize === undefined ? 40 : args.titleSize,
      args.titleColor === undefined ? titleFont.color : args.titleColor,
      true,
      2
    );
    if (args.subtitle) {
      addSlideText(
        slide,
        args.subtitle,
        margin,
        context.height * 0.5,
        context.width - 2 * margin,
        context.height * 0.12,
        subtitleFont,
        args.subtitleSize === undefined ? 20 : args.subtitleSize,
        subtitleFont.color,
        false,
        2
      );
    }
  }

  function addSectionSlide(presentation, args, context) {
    const slide = newSlide(presentation, context);
    const margin = context.preset && context.preset.spacing
      ? context.preset.spacing.margin : 80;
    const font = presetFont(
      context,
      "title",
      {family: "Arial", size: 36, color: "#000000"}
    );
    addSlideText(
      slide,
      args.title,
      margin,
      context.height * 0.36,
      context.width - 2 * margin,
      context.height * 0.2,
      font,
      font.size,
      font.color,
      true,
      2
    );
  }

  function addBulletsSlide(presentation, args, context) {
    const slide = newSlide(presentation, context);
    const margin = context.preset && context.preset.spacing
      ? context.preset.spacing.margin : 60;
    const titleFont = presetFont(
      context,
      "title",
      {family: "Arial", size: 32, color: "#000000"}
    );
    const bodyFont = presetFont(
      context,
      "body",
      {family: "Arial", size: 18, color: "#222222"}
    );
    addSlideText(
      slide,
      args.title,
      margin,
      36,
      context.width - 2 * margin,
      64,
      titleFont,
      args.titleSize === undefined ? 32 : args.titleSize,
      titleFont.color,
      true,
      1
    );
    const body = addSlideText(
      slide,
      args.items.join("\r"),
      margin,
      120,
      context.width - 2 * margin,
      context.height - 160,
      bodyFont,
      args.bodySize === undefined ? 18 : args.bodySize,
      bodyFont.color,
      false,
      1
    );
    const format = body.TextFrame.TextRange.ParagraphFormat;
    if (!format.Bullet) {
      format.Bullet = {};
    }
    format.Bullet.Visible = -1;
  }

  function addBlankSlide(presentation, args, context) {
    newSlide(presentation, context);
  }

  function validateRuntimePictureDimensions(width, height, context) {
    const horizontalLimit = context.width * 4;
    const verticalLimit = context.height * 4;
    if (
      typeof width !== "number" || typeof height !== "number" ||
      !Number.isFinite(width) || !Number.isFinite(height) ||
      width <= 0 || height <= 0 ||
      width > Number.MAX_SAFE_INTEGER || height > Number.MAX_SAFE_INTEGER ||
      width > horizontalLimit || height > verticalLimit
    ) {
      throw new Error("Presentation image dimensions are unsafe");
    }
  }

  function validateRuntimePictureBox(width, height, args, context) {
    validateRuntimePictureDimensions(width, height, context);
    const horizontalLimit = context.width * 4;
    const verticalLimit = context.height * 4;
    const right = args.left + width;
    const bottom = args.top + height;
    if (
      !Number.isFinite(right) || !Number.isFinite(bottom) ||
      right > Number.MAX_SAFE_INTEGER || bottom > Number.MAX_SAFE_INTEGER ||
      right > horizontalLimit || bottom > verticalLimit
    ) {
      throw new Error("Presentation image target box is unsafe");
    }
  }

  function addSlideImage(presentation, args, context, resources) {
    const slide = slideCollectionItem(presentation.Slides, args.slide);
    let shape = null;
    try {
      shape = slide.Shapes.AddPicture(
        resources[args.imageId],
        false,
        true,
        args.left,
        args.top
      );
      const naturalWidth = shape.Width;
      const naturalHeight = shape.Height;
      validateRuntimePictureDimensions(naturalWidth, naturalHeight, context);
      let scale = 1;
      if (args.width !== undefined && args.height !== undefined) {
        scale = Math.min(args.width / naturalWidth, args.height / naturalHeight);
      } else if (args.width !== undefined) {
        scale = args.width / naturalWidth;
      } else if (args.height !== undefined) {
        scale = args.height / naturalHeight;
      }
      const finalWidth = naturalWidth * scale;
      const finalHeight = naturalHeight * scale;
      validateRuntimePictureBox(finalWidth, finalHeight, args, context);
      shape.Width = finalWidth;
      shape.Height = finalHeight;
      validateRuntimePictureBox(shape.Width, shape.Height, args, context);
      shape.Left = args.left;
      shape.Top = args.top;
    } catch (error) {
      if (shape && typeof shape.Delete === "function") {
        try {
          shape.Delete();
        } catch (cleanupError) {
          // Closing the staged presentation without Save remains the fallback.
        }
      }
      throw error;
    }
  }

  function addSlideTable(presentation, args, context) {
    const slide = slideCollectionItem(presentation.Slides, args.slide);
    const shape = slide.Shapes.AddTable(
      args.rows,
      args.cols,
      args.left,
      args.top,
      args.width,
      args.height
    );
    const bodyFont = presetFont(
      context,
      "body",
      {family: "Arial", size: 11, color: "#222222"}
    );
    for (let row = 1; row <= args.rows; row += 1) {
      for (let column = 1; column <= args.cols; column += 1) {
        const cell = shape.Table.Cell(row, column).Shape;
        const range = cell.TextFrame.TextRange;
        const cellValue = args.data[row - 1][column - 1];
        range.Text = cellValue === null ? "" : String(cellValue);
        styleSlideText(
          range,
          bodyFont,
          args.fontSize === undefined ? 11 : args.fontSize,
          row === 1
            ? (args.headerFont === undefined ? "#FFFFFF" : args.headerFont)
            : bodyFont.color,
          row === 1,
          1
        );
        if (row === 1) {
          range.Font.Bold = -1; // msoTrue (JS true marshals to msoCTrue=1)
          applySlideFill(
            cell.Fill,
            args.headerShade === undefined ? "#4472C4" : args.headerShade
          );
        }
      }
    }
  }

  const slideOperations = {
    "slide.reset": resetSlides,
    "slide.set_size": setSlideSize,
    "slide.apply_preset": applySlidePreset,
    "slide.add_title": addTitleSlide,
    "slide.add_section": addSectionSlide,
    "slide.add_bullets": addBulletsSlide,
    "slide.add_blank": addBlankSlide,
    "slide.add_image": addSlideImage,
    "slide.add_table": addSlideTable
  };
  Object.setPrototypeOf(slideOperations, null);

  const slideOperationNames = [
    "slide.reset",
    "slide.set_size",
    "slide.apply_preset",
    "slide.add_title",
    "slide.add_section",
    "slide.add_bullets",
    "slide.add_blank",
    "slide.add_image",
    "slide.add_table"
  ];
  const slideOperationNameSet = new Set(slideOperationNames);

  const slideArgumentRules = {
    "slide.reset": {required: [], allowed: []},
    "slide.set_size": {required: ["width", "height"], allowed: ["width", "height"]},
    "slide.apply_preset": {required: ["preset"], allowed: ["preset"]},
    "slide.add_title": {required: ["title"], allowed: ["title", "subtitle", "titleSize", "subtitleSize", "titleColor"]},
    "slide.add_section": {required: ["title"], allowed: ["title"]},
    "slide.add_bullets": {required: ["title", "items"], allowed: ["title", "items", "titleSize", "bodySize"]},
    "slide.add_blank": {required: [], allowed: []},
    "slide.add_image": {required: ["slide", "imageId", "left", "top"], allowed: ["slide", "imageId", "left", "top", "width", "height"]},
    "slide.add_table": {required: ["slide", "rows", "cols", "left", "top", "width", "height", "data"], allowed: ["slide", "rows", "cols", "left", "top", "width", "height", "data", "headerShade", "headerFont", "fontSize"]}
  };
  Object.setPrototypeOf(slideArgumentRules, null);

  function exactObject(value, required, allowed, path) {
    if (!isObject(value)) {
      invalidSlidePlan(`${path} is invalid`);
    }
    const keys = Object.keys(value);
    if (
      keys.some(function (key) { return allowed.indexOf(key) === -1; }) ||
      required.some(function (key) { return !hasOwn(value, key); })
    ) {
      invalidSlidePlan(`${path} is invalid`);
    }
  }

  function requireString(value, path, nullable) {
    if ((nullable && value === null) || typeof value === "string") {
      return;
    }
    invalidSlidePlan(`${path} is invalid`);
  }

  function requireNumber(value, path, nonNegative) {
    if (
      !Number.isFinite(value) || Math.abs(value) > Number.MAX_SAFE_INTEGER ||
      (nonNegative && value < 0)
    ) {
      invalidSlidePlan(`${path} is invalid`);
    }
  }

  function requireColor(value, path) {
    if (typeof value !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(value)) {
      invalidSlidePlan(`${path} is invalid`);
    }
  }

  function validateSlideFont(value, path) {
    exactObject(value, ["family", "size", "color"], ["family", "size", "color"], path);
    requireString(value.family, `${path}.family`, false);
    requireNumber(value.size, `${path}.size`, false);
    requireColor(value.color, `${path}.color`);
  }

  function validateSlidePreset(value) {
    exactObject(value, ["name", "colors", "fonts"], ["name", "colors", "fonts", "spacing"], "slide.apply_preset preset");
    requireString(value.name, "slide.apply_preset preset.name", false);
    if (!isObject(value.colors) || !isObject(value.fonts)) {
      invalidSlidePlan("slide.apply_preset preset colors/fonts are invalid");
    }
    // extra custom roles are allowed (Python _role_map parity); every
    // present role must still be a valid color/font
    ["primary", "dark", "background"].forEach(function (name) {
      if (!hasOwn(value.colors, name)) {
        invalidSlidePlan(`slide.apply_preset preset.colors.${name} is required`);
      }
    });
    Object.keys(value.colors).forEach(function (name) {
      requireColor(value.colors[name], `slide.apply_preset preset.colors.${name}`);
    });
    ["title", "body"].forEach(function (name) {
      if (!hasOwn(value.fonts, name)) {
        invalidSlidePlan(`slide.apply_preset preset.fonts.${name} is required`);
      }
    });
    Object.keys(value.fonts).forEach(function (name) {
      validateSlideFont(value.fonts[name], `slide.apply_preset preset.fonts.${name}`);
    });
    if (hasOwn(value, "spacing")) {
      exactObject(
        value.spacing,
        ["margin", "gap", "cardPadding", "lineHeight"],
        ["margin", "gap", "cardPadding", "lineHeight"],
        "slide.apply_preset preset.spacing"
      );
      Object.keys(value.spacing).forEach(function (name) {
        requireNumber(value.spacing[name], `slide.apply_preset preset.spacing.${name}`, true);
      });
    }
  }

  function validateSlideTable(args) {
    if (
      !Number.isSafeInteger(args.rows) || args.rows < 1 ||
      !Number.isSafeInteger(args.cols) || args.cols < 1 ||
      !Array.isArray(args.data) || args.data.length !== args.rows
    ) {
      invalidSlidePlan("slide.add_table dimensions are invalid");
    }
    let cells = 0;
    args.data.forEach(function (row) {
      if (!Array.isArray(row) || row.length !== args.cols) {
        invalidSlidePlan("slide.add_table data must match rows and cols");
      }
      cells += row.length;
      row.forEach(function (cell) {
        if (
          cell !== null && typeof cell !== "string" && typeof cell !== "boolean" &&
          (!Number.isFinite(cell) || Math.abs(cell) > Number.MAX_SAFE_INTEGER)
        ) {
          invalidSlidePlan("slide.add_table cell is invalid");
        }
      });
    });
    if (cells > 10000) {
      invalidSlidePlan("slide.add_table is too large");
    }
  }

  function requireSlideArguments(operation) {
    const args = operation.args;
    const rule = slideArgumentRules[operation.op];
    exactObject(args, rule.required, rule.allowed, `${operation.op} arguments`);
    ["title", "imageId", "headerShade", "headerFont", "titleColor"].forEach(function (name) {
      if (hasOwn(args, name)) {
        requireString(args[name], `${operation.op} ${name}`, false);
      }
    });
    if (hasOwn(args, "subtitle")) {
      requireString(args.subtitle, "slide.add_title subtitle", true);
    }
    ["width", "height", "left", "top"].forEach(function (name) {
      if (hasOwn(args, name)) {
        requireNumber(args[name], `${operation.op} ${name}`, true);
      }
    });
    ["titleSize", "subtitleSize", "bodySize", "fontSize"].forEach(function (name) {
      if (hasOwn(args, name)) {
        requireNumber(args[name], `${operation.op} ${name}`, false);
      }
    });
    ["slide", "rows", "cols"].forEach(function (name) {
      if (hasOwn(args, name) && (!Number.isSafeInteger(args[name]) || args[name] < 1)) {
        invalidSlidePlan(`${operation.op} ${name} is invalid`);
      }
    });
    if (operation.op === "slide.apply_preset") {
      validateSlidePreset(args.preset);
    }
    if (operation.op === "slide.add_bullets") {
      if (!Array.isArray(args.items) || args.items.some(function (item) { return typeof item !== "string"; })) {
        invalidSlidePlan("slide.add_bullets items are invalid");
      }
    }
    ["headerShade", "headerFont", "titleColor"].forEach(function (name) {
      if (hasOwn(args, name)) {
        requireColor(args[name], `${operation.op} ${name}`);
      }
    });
    if (operation.op === "slide.add_table") {
      validateSlideTable(args);
    }
  }

  function slideUtf8ByteLength(value) {
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

  function validateSlideJsonLimits(value) {
    const pending = [{value, depth: 0}];
    const seen = [];
    while (pending.length) {
      const current = pending.pop();
      if (current.depth > 64) {
        invalidSlidePlan("Presentation generation plan is nested too deeply");
      }
      if (typeof current.value === "string") {
        if (Array.from(current.value).length > 100000) {
          invalidSlidePlan("Presentation generation plan string is too long");
        }
      } else if (typeof current.value === "number") {
        if (!Number.isFinite(current.value) || Math.abs(current.value) > Number.MAX_SAFE_INTEGER) {
          invalidSlidePlan("Presentation generation plan number is unsafe");
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
      } else if (
        current.value !== null && typeof current.value !== "boolean"
      ) {
        invalidSlidePlan("Presentation generation plan is not JSON-compatible");
      }
    }
  }

  function canonicalPresentationPath(value, label) {
    if (
      typeof value !== "string" || value.length === 0 || value[0] !== "/" ||
      value.indexOf("\\") !== -1 || value.indexOf("\0") !== -1 ||
      value.indexOf("://") !== -1 || value.indexOf("%") !== -1 ||
      value.indexOf("?") !== -1 || value.indexOf("#") !== -1
    ) {
      invalidSlidePlan(`${label} is unsafe`);
    }
    const normalized = [];
    value.split("/").forEach(function (segment) {
      if (!segment || segment === ".") {
        return;
      }
      if (segment === "..") {
        invalidSlidePlan(`${label} contains traversal`);
      }
      normalized.push(segment);
    });
    if (!normalized.length) {
      invalidSlidePlan(`${label} is unsafe`);
    }
    return `/${normalized.join("/")}`;
  }

  function validateSlideResource(imageId, resources, stagedPath) {
    if (
      !/^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(imageId) ||
      imageId === "__proto__" || imageId === "prototype" ||
      hasOwn(Object.prototype, imageId)
    ) {
      invalidSlidePlan("Presentation imageId must be a logical identifier");
    }
    if (!isObject(resources) || !hasOwn(resources, imageId)) {
      invalidSlidePlan(`Missing Presentation resource: ${imageId}`);
    }
    const canonicalStagedPath = canonicalPresentationPath(
      stagedPath,
      "Presentation stagedPath"
    );
    const lastSeparator = canonicalStagedPath.lastIndexOf("/");
    if (lastSeparator <= 0) {
      invalidSlidePlan("Presentation stagedPath has no session directory");
    }
    const resourceRoot = `${canonicalStagedPath.slice(0, lastSeparator)}/resources`;
    const resource = canonicalPresentationPath(
      resources[imageId],
      `Presentation resource ${imageId}`
    );
    if (resource.indexOf(`${resourceRoot}/`) !== 0) {
      invalidSlidePlan(`Unsafe Presentation resource: ${imageId}`);
    }
  }

  function validateGeometryBox(args, state, operationName) {
    const horizontalLimit = state.width * 4;
    const verticalLimit = state.height * 4;
    if (
      !Number.isFinite(horizontalLimit) || !Number.isFinite(verticalLimit) ||
      horizontalLimit > Number.MAX_SAFE_INTEGER ||
      verticalLimit > Number.MAX_SAFE_INTEGER
    ) {
      invalidSlidePlan("Presentation slide size is outside the safe range");
    }
    const width = hasOwn(args, "width") ? args.width : 0;
    const height = hasOwn(args, "height") ? args.height : 0;
    if (
      !Number.isFinite(args.left) || !Number.isFinite(args.top) ||
      !Number.isFinite(width) || !Number.isFinite(height) ||
      args.left < 0 || args.top < 0 || width < 0 || height < 0 ||
      args.left > horizontalLimit || width > horizontalLimit ||
      args.top > verticalLimit || height > verticalLimit
    ) {
      invalidSlidePlan(`${operationName} geometry exceeds the slide bounds`);
    }
    const right = args.left + width;
    const bottom = args.top + height;
    if (
      !Number.isFinite(right) || !Number.isFinite(bottom) ||
      Math.abs(right) > Number.MAX_SAFE_INTEGER ||
      Math.abs(bottom) > Number.MAX_SAFE_INTEGER ||
      right > horizontalLimit || bottom > verticalLimit
    ) {
      invalidSlidePlan(`${operationName} target box exceeds the slide bounds`);
    }
  }

  function validateDerivedSlideGeometry(operation, state) {
    const presetMargin = state.preset && state.preset.spacing
      ? state.preset.spacing.margin : null;
    if (operation.op === "slide.add_title") {
      const margin = presetMargin === null ? 80 : presetMargin;
      validateGeometryBox({
        left: margin,
        top: state.height * 0.28,
        width: state.width - 2 * margin,
        height: state.height * 0.18
      }, state, operation.op);
      if (operation.args.subtitle) {
        validateGeometryBox({
          left: margin,
          top: state.height * 0.5,
          width: state.width - 2 * margin,
          height: state.height * 0.12
        }, state, operation.op);
      }
    } else if (operation.op === "slide.add_section") {
      const margin = presetMargin === null ? 80 : presetMargin;
      validateGeometryBox({
        left: margin,
        top: state.height * 0.36,
        width: state.width - 2 * margin,
        height: state.height * 0.2
      }, state, operation.op);
    } else if (operation.op === "slide.add_bullets") {
      const margin = presetMargin === null ? 60 : presetMargin;
      validateGeometryBox({
        left: margin,
        top: 36,
        width: state.width - 2 * margin,
        height: 64
      }, state, operation.op);
      validateGeometryBox({
        left: margin,
        top: 120,
        width: state.width - 2 * margin,
        height: state.height - 160
      }, state, operation.op);
    }
  }

  function validatePresetGeometry(preset, state) {
    if (!preset.spacing) {
      return;
    }
    const margin = preset.spacing.margin;
    if (
      margin > state.width * 4 || margin > state.height * 4 ||
      state.width - 2 * margin < 0
    ) {
      invalidSlidePlan("Presentation preset geometry exceeds the slide bounds");
    }
  }

  function validateSlideOperations(plan, resources, stagedPath) {
    if (!isObject(resources)) {
      invalidSlidePlan("Presentation resources must be an object");
    }
    if (
      !isObject(plan) || Object.keys(plan).sort().join(",") !== "component,operations" ||
      plan.component !== "presentation" || !Array.isArray(plan.operations) ||
      plan.operations.length === 0 || plan.operations.length > 10000
    ) {
      invalidSlidePlan("Presentation generation plan is invalid");
    }
    validateSlideJsonLimits(plan);
    let serialized;
    try {
      serialized = JSON.stringify(plan);
    } catch (error) {
      invalidSlidePlan("Presentation generation plan is not JSON-compatible");
    }
    if (
      typeof serialized !== "string" || serialized.length > 2000000 ||
      slideUtf8ByteLength(serialized) > 2000000
    ) {
      invalidSlidePlan("Presentation generation plan is too large");
    }
    const expected = slideOperationNames.slice().sort().join(",");
    if (
      slideOperationNameSet.size !== slideOperationNames.length ||
      Object.keys(slideOperations).sort().join(",") !== expected ||
      Object.keys(slideArgumentRules).sort().join(",") !== expected ||
      slideOperationNames.some(function (name) {
        return !slideOperationNameSet.has(name) || !hasOwn(slideOperations, name) ||
          !hasOwn(slideArgumentRules, name) || typeof slideOperations[name] !== "function" ||
          !slideArgumentRules[name];
      })
    ) {
      invalidSlidePlan("Presentation operation catalog is incomplete");
    }
    if (!isObject(plan.operations[0]) || plan.operations[0].op !== "slide.reset") {
      invalidSlidePlan("Presentation generation plan must begin with slide.reset");
    }
    const state = {count: 0, width: 960, height: 540, preset: null};
    const usedResources = Object.create(null);
    let resetSeen = false;
    plan.operations.forEach(function (operation, operationIndex) {
      if (
        !isObject(operation) || Object.keys(operation).sort().join(",") !== "args,op" ||
        typeof operation.op !== "string" || !isObject(operation.args)
      ) {
        invalidSlidePlan("Presentation operation is invalid");
      }
      if (
        !slideOperationNameSet.has(operation.op) || !hasOwn(slideOperations, operation.op) ||
        !hasOwn(slideArgumentRules, operation.op) || typeof slideOperations[operation.op] !== "function" ||
        !slideArgumentRules[operation.op]
      ) {
        invalidSlidePlan(`Unsupported Presentation operation: ${operation.op}`);
      }
      requireSlideArguments(operation);
      if (operation.op === "slide.reset") {
        if (resetSeen || operationIndex !== 0) {
          invalidSlidePlan("Presentation generation plan must contain one leading slide.reset");
        }
        resetSeen = true;
        state.count = 0;
      } else if (operation.op === "slide.set_size") {
        if (
          operation.args.width <= 0 || operation.args.height <= 0 ||
          operation.args.width > Math.floor(Number.MAX_SAFE_INTEGER / 4) ||
          operation.args.height > Math.floor(Number.MAX_SAFE_INTEGER / 4)
        ) {
          invalidSlidePlan("slide.set_size dimensions are invalid");
        }
        state.width = operation.args.width;
        state.height = operation.args.height;
      } else if (operation.op === "slide.apply_preset") {
        validatePresetGeometry(operation.args.preset, state);
        state.preset = operation.args.preset;
      } else if (
        operation.op === "slide.add_title" || operation.op === "slide.add_section" ||
        operation.op === "slide.add_bullets" || operation.op === "slide.add_blank"
      ) {
        validateDerivedSlideGeometry(operation, state);
        state.count += 1;
      } else {
        if (operation.args.slide > state.count) {
          invalidSlidePlan(`${operation.op} index exceeds planned slide count`);
        }
        validateGeometryBox(operation.args, state, operation.op);
        if (operation.op === "slide.add_image") {
          validateSlideResource(operation.args.imageId, resources, stagedPath);
          usedResources[operation.args.imageId] = true;
        }
      }
    });
    Object.keys(resources).forEach(function (resourceId) {
      if (!usedResources[resourceId]) {
        invalidSlidePlan(`Unused Presentation resource: ${resourceId}`);
      }
    });
    return plan.operations;
  }

  function executeSlideOperations(presentation, plan, resources, stagedPath) {
    const operations = validateSlideOperations(plan, resources, stagedPath);
    const context = {width: 960, height: 540, preset: null};
    operations.forEach(function (operation) {
      slideOperations[operation.op](
        presentation,
        operation.args,
        context,
        resources
      );
    });
    return operations.length;
  }

  function generatePresentationDeck(params) {
    const stagedPath = requirePath(params, "stagedPath");
    const resources = hasOwn(params, "resources") ? params.resources : {};
    validateSlideOperations(params.plan, resources, stagedPath);
    const previousAlerts = Application.DisplayAlerts;
    let presentation = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      presentation = Application.Presentations.Open(stagedPath, false, false, false);
      const applied = executeSlideOperations(
        presentation,
        params.plan,
        resources,
        stagedPath
      );
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
    "generate_presentation_deck": generatePresentationDeck,
    "inspect_presentation": inspectPresentation,
    "edit_presentation": editPresentation
  };

  function safeGet(object, property, fallback) {
    try {
      var value = object[property];
      return value === undefined ? fallback : value;
    } catch (error) {
      return fallback;
    }
  }

  function safeCall(object, method, fallback) {
    try {
      var value = object[method]();
      return value === undefined ? fallback : value;
    } catch (error) {
      return fallback;
    }
  }

  function fontSnapshot(font) {
    if (!font) { return null; }
    return {
      name: safeGet(font, "Name", null),
      size: safeGet(font, "Size", null),
      bold: safeGet(font, "Bold", null),
      italic: safeGet(font, "Italic", null),
      underline: safeGet(font, "Underline", null),
      strikethrough: safeGet(font, "StrikeThrough", null),
      subscript: safeGet(font, "Subscript", null),
      superscript: safeGet(font, "Superscript", null),
      name_far_east: safeGet(font, "NameFarEast", null),
      color: safeGet(safeGet(font, "Color", null), "RGB", null)
    };
  }

  function paragraphSnapshot(paragraphFormat) {
    if (!paragraphFormat) { return null; }
    return {
      alignment: safeGet(paragraphFormat, "Alignment", null),
      left_indent: safeGet(paragraphFormat, "LeftIndent", null),
      first_line_indent: safeGet(paragraphFormat, "FirstLineIndent", null),
      line_spacing: safeGet(paragraphFormat, "LineSpacing", null),
      line_rule_within: safeGet(paragraphFormat, "LineRuleWithin", null),
      space_before: safeGet(paragraphFormat, "SpaceBefore", null),
      space_after: safeGet(paragraphFormat, "SpaceAfter", null)
    };
  }

  function fillSnapshot(fill) {
    if (!fill) { return null; }
    return {
      type: safeGet(fill, "Type", null),
      visible: safeGet(fill, "Visible", null),
      fore_color_rgb: safeGet(safeGet(fill, "ForeColor", null), "RGB", null)
    };
  }

  function lineSnapshot(line) {
    if (!line) { return null; }
    return {
      visible: safeGet(line, "Visible", null),
      weight: safeGet(line, "Weight", null),
      fore_color_rgb: safeGet(safeGet(line, "ForeColor", null), "RGB", null)
    };
  }

  function paragraphsSnapshot(textRange, elementId, includeText) {
    var paragraphRange = safeCall(textRange, "Paragraphs", null);
    var count = paragraphRange ? Number(safeGet(paragraphRange, "Count", 0) || 0) : 0;
    var paragraphs = [];
    for (var index = 1; index <= count; index += 1) {
      var para = textRange.Paragraphs(index, 1);
      var paragraphId = elementId + "/paragraph:" + index;
      var snap = {
        id: paragraphId,
        index: index,
        font: fontSnapshot(safeGet(para, "Font", null)),
        paragraph: paragraphSnapshot(safeGet(para, "ParagraphFormat", null)),
        runs: []
      };
      if (includeText) { snap.text = String(safeGet(para, "Text", "") || ""); }
      var runRange = safeCall(para, "Runs", null);
      var runCount = runRange ? Number(safeGet(runRange, "Count", 0) || 0) : 0;
      for (var runIndex = 1; runIndex <= runCount; runIndex += 1) {
        var run = para.Runs(runIndex, 1);
        var runSnap = {
          id: paragraphId + "/run:" + runIndex,
          index: runIndex,
          font: fontSnapshot(safeGet(run, "Font", null))
        };
        if (includeText) { runSnap.text = String(safeGet(run, "Text", "") || ""); }
        snap.runs.push(runSnap);
      }
      paragraphs.push(snap);
    }
    return paragraphs;
  }

  function tableSnapshot(shape, elementId, includeText) {
    var table = safeGet(shape, "Table", null);
    if (!table) { return null; }
    var rows = Number(safeGet(safeGet(table, "Rows", {}), "Count", 0) || 0);
    var cols = Number(safeGet(safeGet(table, "Columns", {}), "Count", 0) || 0);
    var cells = [];
    for (var row = 1; row <= rows; row += 1) {
      for (var col = 1; col <= cols; col += 1) {
        var cellShape = table.Cell(row, col).Shape;
        var cellTextRange = safeGet(safeGet(cellShape, "TextFrame", null), "TextRange", null);
        var cellSnap = {
          id: elementId + "/table/cell:" + row + "," + col,
          row: row,
          column: col,
          fill: fillSnapshot(safeGet(cellShape, "Fill", null))
        };
        if (cellTextRange) {
          if (includeText) { cellSnap.text = String(safeGet(cellTextRange, "Text", "") || ""); }
          cellSnap.font = fontSnapshot(safeGet(cellTextRange, "Font", null));
        }
        cells.push(cellSnap);
      }
    }
    return {rows: rows, columns: cols, cells: cells};
  }

  function shapeSnapshot(shape, slideIndex, shapeIndex, includeText, prefix) {
    var shapeId = safeGet(shape, "Id", null);
    var elementId = shapeId !== null
      ? prefix + "/shape:@id=" + shapeId
      : prefix + "/shape:" + shapeIndex;
    var hasTable = false;
    try { hasTable = Boolean(shape.HasTable); } catch (e) { hasTable = false; }
    var placeholderType = null;
    try { placeholderType = safeGet(shape.PlaceholderFormat, "Type", null); } catch(e) {}
    var snapshot = {
      id: elementId,
      index: shapeIndex,
      shape_id: shapeId,
      name: safeGet(shape, "Name", null),
      type: safeGet(shape, "Type", null),
      placeholder_type: placeholderType,
      geometry: {
        left: safeGet(shape, "Left", null),
        top: safeGet(shape, "Top", null),
        width: safeGet(shape, "Width", null),
        height: safeGet(shape, "Height", null),
        rotation: safeGet(shape, "Rotation", null)
      },
      fill: fillSnapshot(safeGet(shape, "Fill", null)),
      line: lineSnapshot(safeGet(shape, "Line", null))
    };
    var textFrame = safeGet(shape, "TextFrame", null);
    if (textFrame) {
      var range = safeGet(textFrame, "TextRange", null);
      var tf = {
        margin_left: safeGet(textFrame, "MarginLeft", null),
        margin_right: safeGet(textFrame, "MarginRight", null),
        margin_top: safeGet(textFrame, "MarginTop", null),
        margin_bottom: safeGet(textFrame, "MarginBottom", null),
        word_wrap: safeGet(textFrame, "WordWrap", null),
        auto_size: safeGet(textFrame, "AutoSize", null),
        vertical_anchor: safeGet(textFrame, "VerticalAnchor", null)
      };
      if (range) {
        tf.text = includeText ? String(safeGet(range, "Text", "") || "") : null;
        tf.font = fontSnapshot(safeGet(range, "Font", null));
        tf.paragraph = paragraphSnapshot(safeGet(range, "ParagraphFormat", null));
        tf.paragraphs = paragraphsSnapshot(range, elementId, includeText);
        // COM parity: also surface text and font at the shape top level.
        if (includeText) {
          snapshot.text = String(safeGet(range, "Text", "") || "");
        }
        snapshot.font = fontSnapshot(safeGet(range, "Font", null));
      }
      snapshot.text_frame = tf;
    }
    if (hasTable) {
      snapshot.table = tableSnapshot(shape, elementId, includeText);
    }
    return snapshot;
  }

  function collectionItem(collection, index) {
    if (collection && typeof collection.Item === "function") {
      return collection.Item(index);
    }
    if (typeof collection === "function") {
      return collection(index);
    }
    throw new Error("collection item " + index + " is unavailable");
  }

  function inspectPresentation(params) {
    var sourcePath = requirePath(params, "sourcePath");
    var includeText = hasOwn(params, "includeText") ? params.includeText !== false : true;
    var maxShapes = hasOwn(params, "maxShapes") ? Number(params.maxShapes) : null;
    var previousAlerts = Application.DisplayAlerts;
    var presentation = null;
    var failure = null;
    try {
      Application.DisplayAlerts = 0;
      presentation = Application.Presentations.Open(sourcePath, true, false, false);
      var slideCount = Number(safeGet(presentation.Slides, "Count", 0) || 0);
      var slides = [];
      var remaining = maxShapes !== null ? maxShapes : Infinity;
      var truncated = false;
      for (var slideIndex = 1; slideIndex <= slideCount; slideIndex += 1) {
        var slide = collectionItem(presentation.Slides, slideIndex);
        var shapeCount = Number(safeGet(slide.Shapes, "Count", 0) || 0);
        var shapes = [];
        for (var shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex += 1) {
          if (remaining <= 0) { truncated = true; break; }
          shapes.push(shapeSnapshot(
            collectionItem(slide.Shapes, shapeIndex), slideIndex, shapeIndex,
            includeText, "slide:" + slideIndex
          ));
          remaining -= 1;
        }
        var notes = null;
        try {
          notes = String(collectionItem(slide.NotesPage.Shapes.Placeholders, 2).TextFrame.TextRange.Text || "");
        } catch (e) { notes = null; }
        slides.push({
          id: "slide:" + slideIndex,
          index: slideIndex,
          name: safeGet(slide, "Name", null),
          layout: safeGet(slide, "Layout", null),
          follow_master_background: safeGet(slide, "FollowMasterBackground", null),
          shape_count: shapeCount,
          shapes: shapes,
          notes: notes
        });
        if (truncated) { break; }
      }
      return {
        kind: "slide",
        name: safeGet(presentation, "Name", null),
        path: safeGet(presentation, "FullName", null),
        saved: safeGet(presentation, "Saved", null),
        slide_count: slideCount,
        shapes_truncated: truncated,
        slides: slides
      };
    } catch (error) {
      failure = conversionError(error, "CONVERSION_COMMAND_FAILED");
      throw failure;
    } finally {
      try {
        if (presentation !== null) { presentation.Close(); }
      } catch (closeError) {
        if (failure === null) { throw conversionError(closeError, "CONVERSION_COMMAND_FAILED"); }
      } finally {
        Application.DisplayAlerts = previousAlerts;
      }
    }
  }


  // ---- patch application helpers (edit) ----

  function safeSet(object, property, value) {
    try { object[property] = value; return true; }
    catch (e) { return false; }
  }

  function colorLong(value) {
    if (typeof value === "number") { return value; }
    if (typeof value !== "string") { return null; }
    var c = value.replace(/^#/, "");
    if (!/^[0-9A-Fa-f]{6}$/.test(c)) { return null; }
    var r = parseInt(c.slice(0, 2), 16);
    var g = parseInt(c.slice(2, 4), 16);
    var b = parseInt(c.slice(4, 6), 16);
    return r + (g << 8) + (b << 16);
  }

  function applyFontEdit(font, patch) {
    var accepted = [], rejected = [];
    if (!font || typeof font !== "object") {
      return {accepted: [], rejected: Object.keys(patch || {})};
    }
    var mapping = {
      name: "Name", size: "Size", bold: "Bold",
      italic: "Italic", underline: "Underline",
      strikethrough: "StrikeThrough"
    };
    Object.keys(patch || {}).forEach(function (key) {
      var value = patch[key];
      var ok = false;
      if (key === "color") {
        var long = colorLong(value);
        if (long !== null) {
          try { font.Color.RGB = long; ok = true; } catch (e) {
            try { font.Color = long; ok = true; } catch (e2) { ok = false; }
          }
        }
      } else if (mapping[key]) {
        ok = safeSet(font, mapping[key], value);
      } else {
        ok = false;
      }
      (ok ? accepted : rejected).push(key);
    });
    return {accepted: accepted, rejected: rejected};
  }

  function applyParagraphEdit(fmt, patch) {
    var accepted = [], rejected = [];
    if (!fmt || typeof fmt !== "object") {
      return {accepted: [], rejected: Object.keys(patch || {})};
    }
    var mapping = {
      alignment: "Alignment",
      left_indent: "LeftIndent",
      first_line_indent: "FirstLineIndent",
      line_spacing: "LineSpacing",
      line_rule_within: "LineRuleWithin",
      space_before: "SpaceBefore",
      space_after: "SpaceAfter"
    };
    Object.keys(patch || {}).forEach(function (key) {
      var ok = mapping[key] ? safeSet(fmt, mapping[key], patch[key]) : false;
      (ok ? accepted : rejected).push(key);
    });
    return {accepted: accepted, rejected: rejected};
  }

  function applyGeometryEdit(shape, patch) {
    var accepted = [], rejected = [];
    if (!shape || typeof shape !== "object") {
      return {accepted: [], rejected: Object.keys(patch || {})};
    }
    var mapping = {
      left: "Left", top: "Top", width: "Width",
      height: "Height", rotation: "Rotation"
    };
    Object.keys(patch || {}).forEach(function (key) {
      var ok = mapping[key] ? safeSet(shape, mapping[key], patch[key]) : false;
      (ok ? accepted : rejected).push(key);
    });
    return {accepted: accepted, rejected: rejected};
  }

  function applyFillEdit(fill, patch) {
    var accepted = [], rejected = [];
    if (!fill || typeof fill !== "object") {
      return {accepted: [], rejected: Object.keys(patch || {})};
    }
    Object.keys(patch || {}).forEach(function (key) {
      var value = patch[key];
      var ok = false;
      if (key === "color") {
        var long = colorLong(value);
        if (long !== null) {
          try { fill.Visible = -1; fill.ForeColor.RGB = long; ok = true; }
          catch (e) { ok = false; }
        }
      } else if (key === "visible") {
        ok = safeSet(fill, "Visible", value);
      } else if (key === "transparency") {
        ok = safeSet(fill, "Transparency", value);
      }
      (ok ? accepted : rejected).push(key);
    });
    return {accepted: accepted, rejected: rejected};
  }

  function applyLineEdit(line, patch) {
    var accepted = [], rejected = [];
    if (!line || typeof line !== "object") {
      return {accepted: [], rejected: Object.keys(patch || {})};
    }
    var mapping = { visible: "Visible", weight: "Weight" };
    Object.keys(patch || {}).forEach(function (key) {
      var value = patch[key];
      var ok = false;
      if (key === "color") {
        var long = colorLong(value);
        if (long !== null) {
          try { line.Visible = -1; line.ForeColor.RGB = long; ok = true; }
          catch (e) { ok = false; }
        }
      } else if (mapping[key]) {
        ok = safeSet(line, mapping[key], value);
      }
      (ok ? accepted : rejected).push(key);
    });
    return {accepted: accepted, rejected: rejected};
  }

  function mergeResults() {
    var accepted = [], rejected = [];
    for (var i = 0; i < arguments.length; i += 1) {
      var r = arguments[i] || {};
      (r.accepted || []).forEach(function (k) { accepted.push(k); });
      (r.rejected || []).forEach(function (k) { rejected.push(k); });
    }
    return {accepted: accepted, rejected: rejected};
  }

  function patchTextRange(range, patch) {
    var font = patch.font || {};
    var paragraph = patch.paragraph || {};
    var results = [
      applyFontEdit(safeGet(range, "Font", null), font),
      applyParagraphEdit(safeGet(range, "ParagraphFormat", null), paragraph)
    ];
    if (patch.text !== undefined && patch.text !== null) {
      var ok = safeSet(range, "Text", String(patch.text));
      results.push({accepted: ok ? ["text"] : [], rejected: ok ? [] : ["text"]});
    }
    return mergeResults.apply(null, results);
  }

  function patchShape(shape, patch) {
    var results = [
      applyGeometryEdit(shape, patch.geometry || {}),
      applyFillEdit(safeGet(shape, "Fill", null), patch.fill || {}),
      applyLineEdit(safeGet(shape, "Line", null), patch.line || {})
    ];
    var textFrame = safeGet(shape, "TextFrame", null);
    var range = textFrame ? safeGet(textFrame, "TextRange", null) : null;
    if (range) {
      results.push(patchTextRange(range, patch));
      var tfPatch = patch.text_frame || {};
      var tfMap = {
        margin_left: "MarginLeft", margin_right: "MarginRight",
        margin_top: "MarginTop", margin_bottom: "MarginBottom",
        word_wrap: "WordWrap", auto_size: "AutoSize",
        vertical_anchor: "VerticalAnchor"
      };
      var tfAccepted = [], tfRejected = [];
      if (patch.vertical_alignment !== undefined && patch.vertical_alignment !== null) {
        tfPatch.vertical_anchor = patch.vertical_alignment;
      }
      Object.keys(tfPatch).forEach(function (key) {
        var ok = tfMap[key] ? safeSet(textFrame, tfMap[key], tfPatch[key]) : false;
        (ok ? tfAccepted : tfRejected).push("text_frame." + key);
      });
      results.push({accepted: tfAccepted, rejected: tfRejected});
    }
    return mergeResults.apply(null, results);
  }

  function findShapeByAttr(slide, attr, value) {
    var count = Number(safeGet(slide.Shapes, "Count", 0) || 0);
    for (var i = 1; i <= count; i += 1) {
      var shape = collectionItem(slide.Shapes, i);
      var v = safeGet(shape, attr, null);
      if (String(v) === String(value)) { return shape; }
    }
    return null;
  }

  function resolveTarget(presentation, target) {
    // Returns {slide, shape, paragraph, run, cell} or throws.
    var m;
    m = target.match(/^slide:(\d+)$/);
    if (m) {
      return {slide: collectionItem(presentation.Slides, Number(m[1]))};
    }
    m = target.match(/^slide:(\d+)\/shape:@id=(\d+)$/);
    if (m) {
      var slide = collectionItem(presentation.Slides, Number(m[1]));
      var shape = findShapeByAttr(slide, "Id", m[2]);
      if (!shape) { throw new Error("shape @id=" + m[2] + " not found"); }
      return {slide: slide, shape: shape};
    }
    m = target.match(/^slide:(\d+)\/shape:@name=(.+)$/);
    if (m) {
      slide = collectionItem(presentation.Slides, Number(m[1]));
      shape = findShapeByAttr(slide, "Name", m[2]);
      if (!shape) { throw new Error("shape @name=" + m[2] + " not found"); }
      return {slide: slide, shape: shape};
    }
    m = target.match(/^slide:(\d+)\/shape:(\d+)$/);
    if (m) {
      slide = collectionItem(presentation.Slides, Number(m[1]));
      return {slide: slide, shape: collectionItem(slide.Shapes, Number(m[2]))};
    }
    m = target.match(/^slide:(\d+)\/shape:(\d+)\/table\/cell:(\d+),(\d+)$/);
    if (m) {
      slide = collectionItem(presentation.Slides, Number(m[1]));
      shape = collectionItem(slide.Shapes, Number(m[2]));
      var cellShape = shape.Table.Cell(Number(m[3]), Number(m[4])).Shape;
      return {slide: slide, shape: cellShape};
    }
    m = target.match(/^slide:(\d+)\/shape:@id=(\d+)\/table\/cell:(\d+),(\d+)$/);
    if (m) {
      slide = collectionItem(presentation.Slides, Number(m[1]));
      shape = findShapeByAttr(slide, "Id", m[2]);
      if (!shape) { throw new Error("shape @id=" + m[2] + " not found"); }
      cellShape = shape.Table.Cell(Number(m[3]), Number(m[4])).Shape;
      return {slide: slide, shape: cellShape};
    }
    m = target.match(/^slide:(\d+)\/shape:(\d+)\/paragraph:(\d+)(?:\/run:(\d+))?$/);
    if (m) {
      slide = collectionItem(presentation.Slides, Number(m[1]));
      shape = collectionItem(slide.Shapes, Number(m[2]));
      var range = shape.TextFrame.TextRange;
      var para = range.Paragraphs(Number(m[3]), 1);
      var rng = m[4] ? para.Runs(Number(m[4]), 1) : para;
      return {slide: slide, shape: shape, range: rng};
    }
    m = target.match(/^slide:(\d+)\/shape:@id=(\d+)\/paragraph:(\d+)(?:\/run:(\d+))?$/);
    if (m) {
      slide = collectionItem(presentation.Slides, Number(m[1]));
      shape = findShapeByAttr(slide, "Id", m[2]);
      if (!shape) { throw new Error("shape @id=" + m[2] + " not found"); }
      range = shape.TextFrame.TextRange;
      para = range.Paragraphs(Number(m[3]), 1);
      rng = m[4] ? para.Runs(Number(m[4]), 1) : para;
      return {slide: slide, shape: shape, range: rng};
    }
    throw new Error("Unsupported edit target: " + target);
  }

  function editPresentation(params) {
    var sourcePath = requirePath(params, "sourcePath");
    var outputPath = requirePath(params, "outputPath");
    var patches = params.patches || [];
    var atomic = params.atomic !== false;
    var raiseOnError = params.raiseOnError === true;
    var previousAlerts = Application.DisplayAlerts;
    var presentation = null;
    var failure = null;
    try {
      Application.DisplayAlerts = 0;
      presentation = Application.Presentations.Open(sourcePath, false, false, false);
      var reports = patches.map(function (patch) {
        var target = patch.target;
        var rest = {};
        Object.keys(patch).forEach(function (k) {
          if (k !== "target") { rest[k] = patch[k]; }
        });
        try {
          var resolved = resolveTarget(presentation, target);
          var result;
          if (resolved.range) {
            result = patchTextRange(resolved.range, rest);
          } else if (resolved.shape) {
            result = patchShape(resolved.shape, rest);
          } else if (resolved.slide) {
            // slide-level patch (background, name, follow_master_background)
            var accepted = [], rejected = [];
            if (rest.name !== undefined) {
              (safeSet(resolved.slide, "Name", rest.name) ? accepted : rejected).push("name");
            }
            if (rest.follow_master_background !== undefined) {
              (safeSet(resolved.slide, "FollowMasterBackground", rest.follow_master_background) ? accepted : rejected).push("follow_master_background");
            }
            if (rest.background) {
              var fr = applyFillEdit(safeGet(safeGet(resolved.slide, "Background", null), "Fill", null), rest.background);
              result = mergeResults({accepted: accepted, rejected: rejected}, fr);
            } else {
              result = {accepted: accepted, rejected: rejected};
            }
          } else {
            result = {accepted: [], rejected: ["target"]};
          }
          return {target: target, accepted: result.accepted, rejected: result.rejected, ok: result.rejected.length === 0};
        } catch (e) {
          return {target: target, accepted: [], rejected: [], ok: false, error: String(e.message || e)};
        }
      });
      var hasFailures = reports.some(function (report) { return report.ok !== true; });
      if (atomic && hasFailures) {
        return {path: null, patches: reports, saved: false};
      }
      if (raiseOnError && hasFailures) {
        throw new Error("One or more presentation patches failed");
      }
      presentation.SaveAs(outputPath, 24);
      return {path: outputPath, patches: reports, saved: true};
    } catch (error) {
      failure = conversionError(error, "CONVERSION_COMMAND_FAILED");
      throw failure;
    } finally {
      try {
        if (presentation !== null) { presentation.Close(); }
      } catch (closeError) {
        if (failure === null) { throw conversionError(closeError, "CONVERSION_COMMAND_FAILED"); }
      } finally {
        Application.DisplayAlerts = previousAlerts;
      }
    }
  }

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
