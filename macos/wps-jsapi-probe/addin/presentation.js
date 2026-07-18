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
        range.Text = String(args.data[row - 1][column - 1]);
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
          range.Font.Bold = true;
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
    exactObject(
      value.colors,
      ["primary", "dark", "background"],
      ["primary", "secondary", "accent", "dark", "light", "background"],
      "slide.apply_preset preset.colors"
    );
    Object.keys(value.colors).forEach(function (name) {
      requireColor(value.colors[name], `slide.apply_preset preset.colors.${name}`);
    });
    exactObject(
      value.fonts,
      ["title", "body"],
      ["title", "subtitle", "body", "caption", "chinese"],
      "slide.apply_preset preset.fonts"
    );
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
      !/^[A-Za-z0-9][A-Za-z0-9_-]*$/.test(imageId) ||
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
