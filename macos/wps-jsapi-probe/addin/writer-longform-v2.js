(function () {
  "use strict";

  // BEGIN WPSCOMPOSER GENERATED RECOVERY MATRIX
  const WPSCOMPOSER_RECOVERY_MATRIX = Object.freeze({
    "BIBLIOGRAPHY_INSERT_FAILED": Object.freeze({
      "notice": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    }),
    "CROSS_REFERENCE_FAILED": Object.freeze({
      "inline-fallback": Object.freeze({"fallbackAttempts":1,"placement":"inline"})
    }),
    "DEGRADATION_INSERT_FAILED": Object.freeze({
      "inline": Object.freeze({"fallbackAttempts":1,"placement":"inline"}),
      "notice": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    }),
    "EQUATION_INSERT_FAILED": Object.freeze({
      "explicit-image-then-source-notice": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    }),
    "FIELD_REFRESH_UNSTABLE": Object.freeze({
      "document-quality-notice": Object.freeze({"fallbackAttempts":1,"placement":"document"})
    }),
    "IMAGE_INSERT_FAILED": Object.freeze({
      "figure-child-stack-then-notice": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    }),
    "TABLE_INSERT_FAILED": Object.freeze({
      "grid-then-text": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    }),
    "TABLE_MERGE_APPLY_FAILED": Object.freeze({
      "grid-then-text": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    }),
    "TABLE_ROW_FORCED_SPLIT": Object.freeze({
      "grid-then-text": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    }),
    "TABLE_STYLE_APPLY_FAILED": Object.freeze({
      "grid-then-text": Object.freeze({"fallbackAttempts":1,"placement":"block"})
    })
  });
  // END WPSCOMPOSER GENERATED RECOVERY MATRIX

  const LONGFORM_DEFERRED = {};

  const NATIVE_SEQUENCE_KIND = Object.freeze({
    WPSC_FIG: "SEQ_FIG",
    WPSC_TAB: "SEQ_TAB",
    WPSC_EQ: "SEQ_EQ"
  });
  const RECOVERABLE_TABLE_CODES = Object.freeze({
    TABLE_STYLE_APPLY_FAILED: true,
    TABLE_MERGE_APPLY_FAILED: true,
    TABLE_ROW_FORCED_SPLIT: true,
    TABLE_INSERT_FAILED: true
  });

  function hasOwn(object, key) {
    return Object.prototype.hasOwnProperty.call(object, key);
  }

  function safeNumber(value, fallback) {
    const n = Number(value);
    return Number.isFinite(n) ? n : fallback;
  }

  function safeString(value) {
    return value === null || value === undefined ? "" : String(value);
  }

  // BEGIN WPSCOMPOSER GENERATED PRIVACY FILTER
  const WPSCOMPOSER_PRIVATE_PATTERNS = Object.freeze([
    Object.freeze({"flags":"i","source":"(?:^|[^0-9a-f])[0-9a-f]{64}(?:$|[^0-9a-f])"}),
    Object.freeze({"flags":"","source":"(?:Traceback|[A-Za-z_][A-Za-z0-9_.]*(?:Error|Exception))\\s*(?:\\(|\\b)"}),
    Object.freeze({"flags":"i","source":"(?:^|[^A-Za-z0-9])[A-Za-z][A-Za-z0-9+.-]*://"}),
    Object.freeze({"flags":"i","source":"(?:^|[^A-Za-z0-9])(?:data|blob):"}),
    Object.freeze({"flags":"i","source":"[A-Za-z]:[\\\\/][^\\s|,;]+"}),
    Object.freeze({"flags":"","source":"(?:\\\\\\\\|//)[A-Za-z0-9_.-]+[\\\\/][^\\s|,;]+"}),
    Object.freeze({"flags":"","source":"(?:^|[\\s({=:\\\"'])~[\\\\/]"}),
    Object.freeze({"flags":"","source":"(?:^|[\\\\/])\\.\\.[\\\\/]"}),
    Object.freeze({"flags":"","source":"(?:^|[\\s({=:\\\"'])\\.\\.?[\\\\/]"}),
    Object.freeze({"flags":"","source":"(?:^|[^A-Za-z0-9_:-])(?:[A-Za-z0-9_.-]+[\\\\/])+[A-Za-z0-9_.-]+\\.[A-Za-z][A-Za-z0-9]{0,15}(?:$|[^A-Za-z0-9_.-])"}),
    Object.freeze({"flags":"i","source":"\\b(?:path|source|sourcePath|stagingPath|file)\\s*[:=]\\s*(?:\\.\\.?[\\\\/]|[^\\s|,;]+[\\\\/][^\\s|,;]+)"}),
    Object.freeze({"flags":"","source":"(?:^|[\\s({=:\\\"']|:(?!/))/(?!/)(?=\\S)"})
  ]);
  const WPSCOMPOSER_BASE64_PATTERN = "[A-Za-z0-9+/]{76,}={0,2}";

  function wpscHasBase64Payload(text) {
    const matches = text.match(new RegExp(WPSCOMPOSER_BASE64_PATTERN, "g")) || [];
    return matches.some(function (candidate) {
      return candidate.length % 4 === 0;
    });
  }

  function safePublicText(value) {
    const text = safeString(value);
    const privateValue = WPSCOMPOSER_PRIVATE_PATTERNS.some(function (item) {
      return new RegExp(item.source, item.flags).test(text);
    }) || wpscHasBase64Payload(text);
    return privateValue ? "<redacted>" : text;
  }
  // END WPSCOMPOSER GENERATED PRIVACY FILTER

  function setValue(target, name, value) {
    if (target !== null && typeof target !== "undefined" && typeof value !== "undefined") {
      target[name] = value;
    }
  }

  function colorFromHex(value) {
    if (typeof value !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(value)) {
      return value;
    }
    const red = parseInt(value.slice(1, 3), 16);
    const green = parseInt(value.slice(3, 5), 16);
    const blue = parseInt(value.slice(5, 7), 16);
    return red | (green << 8) | (blue << 16);
  }

  function endRange(document) {
    const content = document.Content;
    const end = content && typeof content.End === "number"
      ? Math.max(0, content.End - 1)
      : 0;
    if (typeof document.Range === "function") {
      return document.Range(end, end);
    }
    return content;
  }

  function collectionItem(collection, index) {
    if (collection && typeof collection.Item === "function") {
      return collection.Item(index);
    }
    if (typeof collection === "function") {
      return collection(index);
    }
    return null;
  }

  function getStyle(document, name) {
    if (!document.Styles) {
      return null;
    }
    try {
      if (typeof document.Styles.Item === "function") {
        return document.Styles.Item(name);
      }
      if (typeof document.Styles === "function") {
        return document.Styles(name);
      }
    } catch (error) {
      return null;
    }
    return null;
  }

  function insertText(document, text, styleName, formatting) {
    formatting = formatting || {};
    const insertion = endRange(document);
    const start = typeof insertion.Start === "number" ? insertion.Start : 0;
    if (typeof insertion.InsertAfter === "function") {
      insertion.InsertAfter(text + "\r");
    } else {
      insertion.Text = String(text) + "\r";
    }
    let written = insertion;
    if (typeof document.Range === "function") {
      written = document.Range(start, start + String(text).length);
    }
    if (styleName) {
      const style = getStyle(document, styleName);
      if (style) {
        written.Style = style;
      }
    }
    applyFont(written.Font, formatting);
    applyParagraphFormat(written.ParagraphFormat, formatting);
    return { range: written, start: start };
  }

  function applyFont(font, args) {
    if (!font) {
      return;
    }
    setValue(font, "Name", args.fontName);
    setValue(font, "NameFarEast", args.fontName);
    setValue(font, "NameAscii", args.fontNameAscii || args.fontName);
    setValue(font, "NameOther", args.fontNameAscii || args.fontName);
    setValue(font, "Size", args.fontSize !== undefined ? args.fontSize : args.size);
    setValue(font, "Bold", args.bold === undefined ? undefined : (args.bold ? -1 : 0));
    setValue(font, "Italic", args.italic === undefined ? undefined : (args.italic ? -1 : 0));
    setValue(font, "Underline", args.underline === undefined ? undefined : (args.underline ? 1 : 0));
    setValue(font, "Color", colorFromHex(args.color));
  }

  function applyParagraphFormat(format, args) {
    if (!format) {
      return;
    }
    const alignments = { left: 0, center: 1, right: 2, justified: 3 };
    const alignValue = args.align;
    if (alignValue !== undefined) {
      setValue(format, "Alignment", typeof alignValue === "string" ? alignments[alignValue] : alignValue);
    }
    setValue(format, "FirstLineIndent", args.indentFirst);
    setValue(format, "LeftIndent", args.leftIndent);
    setValue(format, "RightIndent", args.rightIndent);
    setValue(format, "LineSpacing", args.lineSpacing);
    if (args.lineSpacingRule !== undefined) {
      const rules = { single: 0, one_and_half: 1, double: 2, at_least: 3, exact: 4, multiple: 5 };
      setValue(format, "LineSpacingRule", hasOwn(rules, args.lineSpacingRule) ? rules[args.lineSpacingRule] : args.lineSpacingRule);
    }
    setValue(format, "SpaceBefore", args.spaceBefore);
    setValue(format, "SpaceAfter", args.spaceAfter);
    setValue(format, "KeepTogether", args.keepTogether);
    setValue(format, "KeepWithNext", args.keepWithNext);
    setValue(format, "OutlineLevel", args.outlineLevel);
  }

  function resetDocument(document) {
    document.Content.Text = "";
  }

  function configurePage(document, args) {
    const setup = document.PageSetup;
    setValue(setup, "TopMargin", args.marginTop);
    setValue(setup, "BottomMargin", args.marginBottom);
    setValue(setup, "LeftMargin", args.marginLeft);
    setValue(setup, "RightMargin", args.marginRight);
    setValue(setup, "PageWidth", args.pageWidth);
    setValue(setup, "PageHeight", args.pageHeight);
    setValue(setup, "Orientation", args.landscape === undefined ? undefined : (args.landscape ? 1 : 0));
  }

  function ensureStyles(document, args) {
    const styles = args.styles || [];
    styles.forEach(function (definition) {
      let style = null;
      try {
        style = getStyle(document, definition.name);
      } catch (error) {
        style = null;
      }
      if (!style && document.Styles && typeof document.Styles.Add === "function") {
        try {
          style = document.Styles.Add(definition.name, definition.type === "character" ? 2 : 1);
        } catch (error) {
          style = null;
        }
      }
      if (!style) {
        return;
      }
      if (definition.basedOn !== undefined) {
        const base = getStyle(document, definition.basedOn);
        try {
          style.BaseStyle = base || definition.basedOn;
        } catch (error) {
          // ignore
        }
      }
      applyFont(style.Font, definition);
      applyParagraphFormat(style.ParagraphFormat, definition);
    });
  }

  function addParagraph(document, args) {
    insertText(document, args.text, args.style, args);
  }

  function addHeading(document, args) {
    const styleName = args.style || ("Heading " + args.level);
    return insertText(document, args.text, styleName, args);
  }

  function unnumberedHeadingStyle(document, level, sequenceTransparent) {
    const name = (sequenceTransparent ? "WPSC Sequence Transparent Heading " : "WPSC Unnumbered Heading ") + level;
    let style = null;
    try { style = getStyle(document, name); } catch (error) { style = null; }
    if (!style && document.Styles && typeof document.Styles.Add === "function") {
      style = document.Styles.Add(name, 1);
      const source = getStyle(document, "Heading " + level);
      if (source && source.Font && style.Font) {
        style.Font.Name = source.Font.Name;
        style.Font.NameFarEast = source.Font.NameFarEast;
        style.Font.NameAscii = source.Font.NameAscii;
        style.Font.Size = source.Font.Size;
        style.Font.Bold = source.Font.Bold;
      }
      // SEQ \\s 1 resets on any outline-level-1 paragraph, even if that
      // paragraph has no list number. Explicit sequence-transparent headings
      // keep the H1 visual treatment but use body outline; ordinary globally
      // unnumbered headings remain outline-visible for TOC population.
      style.ParagraphFormat.OutlineLevel = sequenceTransparent ? 10 : level;
    }
    return name;
  }

  function headingListTemplate(document, scheme) {
    if (!document._wpscHeadingTemplates) document._wpscHeadingTemplates = {};
    if (document._wpscHeadingTemplates[scheme]) return document._wpscHeadingTemplates[scheme];
    const safeScheme = String(scheme || "decimal").replace(/[^a-z0-9_-]/gi, "_");
    const template = document.ListTemplates.Add(true, "wpsc_m3_" + safeScheme);
    const formats = scheme === "chinese-formal"
      ? ["第%1章", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4"]
      : ["%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4"];
    for (let level = 1; level <= 4; level += 1) {
      const listLevel = collectionItem(template.ListLevels, level);
      listLevel.NumberFormat = formats[level - 1];
      listLevel.NumberStyle = scheme === "chinese-formal" && level === 1 ? 37 : 0;
      listLevel.NumberPosition = (level - 1) * 18;
      listLevel.TextPosition = level * 18;
      listLevel.ResetOnHigher = level === 1 ? 0 : level - 1;
      listLevel.StartAt = 1;
    }
    document._wpscHeadingTemplates[scheme] = template;
    return template;
  }

  function addHeadingNative(document, args) {
    if (!args.numbering) {
      const styleName = unnumberedHeadingStyle(document, args.level, args.sequenceTransparent === true);
      insertText(document, args.text, styleName, args);
      return;
    }
    const levelIndex = safeNumber(args.level, 1);
    try {
      const template = headingListTemplate(document, args.numberingScheme || "decimal");
      const inserted = addHeading(document, args);
      const builtInStyle = collectionItem(document.Styles, -1 - levelIndex);
      if (builtInStyle) inserted.range.Style = builtInStyle;
      inserted.range.ListFormat.ApplyListTemplateWithLevel(template, false, 0, 0, levelIndex);
      inserted.range.ListFormat.ListLevelNumber = levelIndex;
      if (!String(inserted.range.ListFormat.ListString || "").replace(/\s+/g, "")) {
        throw nativeError("EXECUTION_ABORTED");
      }
    } catch (error) {
      throw nativeError("EXECUTION_ABORTED");
    }
  }

  function addList(document, args) {
    const items = args.items || [];
    const ordered = !!args.ordered;
    items.forEach(function (item, index) {
      const prefix = ordered ? (index + 1) + ".\t" : (args.glyph || "\u2022") + "\t";
      const formatting = {
        leftIndent: args.indent === undefined ? 24 : args.indent,
        indentFirst: -(args.indent === undefined ? 24 : args.indent)
      };
      insertText(document, prefix + item, "List Paragraph", formatting);
    });
  }

  function configureSection(document, args) {
    const first = !document._wpscFirstSectionConfigured;
    if (!first) {
      try {
        endRange(document).InsertBreak(2);
      } catch (error) {
        // ignore
      }
    }
    document._wpscFirstSectionConfigured = true;

    const setup = document.PageSetup;
    if (args.landscape !== undefined) {
      setValue(setup, "Orientation", args.landscape ? 1 : 0);
    }
    if (args.margins) {
      setValue(setup, "TopMargin", args.margins.top);
      setValue(setup, "BottomMargin", args.margins.bottom);
      setValue(setup, "LeftMargin", args.margins.left);
      setValue(setup, "RightMargin", args.margins.right);
    }

    setPageRole(document, args.role || "body");
    setHeaderFooter(document, {
      headerText: args.headerText,
      footerText: args.footerText,
      linkToPreviousHeader: args.linkToPreviousHeader,
      linkToPreviousFooter: args.linkToPreviousFooter
    });
    setPageNumbering(document, {
      format: args.pageNumberFormat || "continue",
      start: args.startPageNumber,
      restart: args.restartPageNumbering
    });
    const front = document._wpscFrontMatter || {};
    if (args.role === "cover" && front.titlePage) {
      insertText(document, front.title || "", "Title", {
        align: "center", size: 24, bold: true, spaceAfter: 24
      });
      if (front.author) insertText(document, front.author, "Body Text", {align: "center", size: 14});
      if (front.date) insertText(document, front.date, "Body Text", {align: "center", size: 12});
    }
  }

  function setPageRole(document, role) {
    try {
      const section = collectionItem(document.Sections, document.Sections.Count);
      if (section && section.Range && section.Range.DocumentVariables && typeof section.Range.DocumentVariables.Add === "function") {
        section.Range.DocumentVariables.Add("WpsComposerSectionRole_" + document.Sections.Count, String(role));
      }
    } catch (error) {
      // ignore
    }
  }

  function setPageNumbering(document, args) {
    try {
      const section = collectionItem(document.Sections, document.Sections.Count);
      const footer = collectionItem(section.Footers, 1);
      const pageNumbers = footer.PageNumbers;
      if (args.restart !== undefined) {
        pageNumbers.RestartNumberingAtSection = args.restart ? -1 : 0;
      }
      if (args.start !== undefined && args.start !== null) {
        pageNumbers.StartingNumber = safeNumber(args.start, 1);
      }
      const styleMap = { none: 0, roman: 2, arabic: 0, continue: 0 };
      if (args.format in styleMap) {
        pageNumbers.NumberStyle = styleMap[args.format];
      }
      if (args.format === "none") {
        footer.Range.Text = "";
      } else {
        try {
          footer.Range.Collapse(0);
          footer.Range.Fields.Add(footer.Range, 33);
        } catch (error) {
          throw nativeError("FIELD_REFRESH_FAILED");
        }
      }
    } catch (error) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
  }

  function setHeaderFooter(document, args) {
    try {
      const section = collectionItem(document.Sections, document.Sections.Count);
      const header = collectionItem(section.Headers, 1);
      const footer = collectionItem(section.Footers, 1);
      if (args.linkToPreviousHeader !== undefined) {
        try {
          header.LinkToPrevious = args.linkToPreviousHeader ? -1 : 0;
        } catch (error) {
          // ignore
        }
      }
      if (args.linkToPreviousFooter !== undefined) {
        try {
          footer.LinkToPrevious = args.linkToPreviousFooter ? -1 : 0;
        } catch (error) {
          // ignore
        }
      }
      if (args.headerText !== undefined && args.headerText !== null) {
        header.Range.Text = String(args.headerText);
        try {
          header.Range.ParagraphFormat.Alignment = 1;
          header.Range.ParagraphFormat.Borders(-3).LineStyle = 1;
          header.Range.ParagraphFormat.Borders(-3).LineWidth = 6;
          header.Range.ParagraphFormat.Borders(-3).Color = 0;
        } catch (error) {
          // ignore
        }
      }
      if (args.footerText !== undefined && args.footerText !== null) {
        if (String(args.footerText) !== "") {
          footer.Range.Text = String(args.footerText);
        }
      }
    } catch (error) {
      // ignore
    }
  }

  function insertTocWithStyles(document, args, resources, context) {
    void resources;
    const density = args.density || document._wpscTocDensity || {};
    let toc;
    try {
      toc = document.TablesOfContents.Add(endRange(document), true, 1, args.levels || 3);
    } catch (error) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    trackNativeField(
      document,
      context && context.ownerNodeId ? context.ownerNodeId : "doc:toc",
      "TOC",
      toc,
      "index"
    );
    ["toc1", "toc2", "toc3"].forEach(function (key, index) {
      try {
        const style = getStyle(document, "TOC " + (index + 1));
        if (!style) {
          return;
        }
        const minFont = density.minFontSizePt && density.minFontSizePt[key];
        if (minFont !== undefined && minFont !== null) {
          style.Font.Size = safeNumber(minFont, style.Font.Size);
        }
        const minBefore = density.minSpaceBeforePt && density.minSpaceBeforePt[key];
        if (minBefore !== undefined && minBefore !== null) {
          style.ParagraphFormat.SpaceBefore = safeNumber(minBefore, 0);
        }
        const minAfter = density.minSpaceAfterPt && density.minSpaceAfterPt[key];
        if (minAfter !== undefined && minAfter !== null) {
          style.ParagraphFormat.SpaceAfter = safeNumber(minAfter, 0);
        }
      } catch (error) {
        // ignore
      }
    });
  }

  function nativeError(code) {
    const error = new Error(code);
    error.code = code;
    return error;
  }

  function recoveryDecision(code, fallbackKind, placement) {
    const branches = WPSCOMPOSER_RECOVERY_MATRIX[code];
    const rule = branches && branches[fallbackKind];
    if (!rule || rule.placement !== placement) {
      throw nativeError(typeof code === "string" && code ? code : "EXECUTION_FAILED");
    }
    return {
      code: code,
      recoverable: true,
      placement: placement,
      fallbackKind: fallbackKind,
      fallbackAttempts: rule.fallbackAttempts
    };
  }

  function createLocalRecoveryController() {
    return {byIdentity: Object.create(null), decisions: []};
  }

  function runLocalRecovery(controller, spec) {
    if (!controller || !spec || !spec.descriptor) {
      throw nativeError("CONFIGURATION_INVALID");
    }
    const descriptor = spec.descriptor;
    const identity = safeString(spec.nodeId) + "\u0000" + safeString(descriptor.code);
    if (controller.byIdentity[identity]) return controller.byIdentity[identity];

    let checkpoint;
    try { checkpoint = spec.checkpoint(); }
    catch (error) { throw nativeError("LOCAL_MUTATION_CHECKPOINT_FAILED"); }

    let nativeCode = null;
    try {
      spec.nativeAttempt();
      return {
        code: descriptor.code,
        recoverable: false,
        placement: descriptor.placement,
        fallbackKind: descriptor.fallbackKind,
        fallbackAttempts: 0
      };
    } catch (error) {
      nativeCode = error && typeof error.code === "string" ? error.code : "EXECUTION_FAILED";
    }
    return recoverLocalFailure(controller, spec, checkpoint, nativeCode);
  }

  function recoverLocalFailure(controller, spec, checkpoint, nativeCode) {
    const descriptor = spec.descriptor;
    const identity = safeString(spec.nodeId) + "\u0000" + safeString(descriptor.code);
    if (controller.byIdentity[identity]) return controller.byIdentity[identity];
    if (nativeCode !== descriptor.code) throw nativeError(nativeCode);
    const decision = recoveryDecision(nativeCode, descriptor.fallbackKind, descriptor.placement);
    try { spec.rollback(checkpoint); }
    catch (error) { throw nativeError("LOCAL_MUTATION_ROLLBACK_FAILED"); }
    try { spec.fallbackAttempt(descriptor); }
    catch (error) {
      if (error && error._wpscPreserveFallbackFatal === true) throw error;
      throw nativeError("DEGRADATION_FALLBACK_FAILED");
    }
    try { spec.insertNotice(safeString(spec.nodeId), descriptor); }
    catch (error) { throw nativeError("DEGRADATION_INSERT_FAILED"); }

    controller.byIdentity[identity] = decision;
    controller.decisions.push(decision);
    return decision;
  }

  function preserveFallbackFatal(error) {
    const failure = error && typeof error === "object"
      ? error : nativeError("EXECUTION_ABORTED");
    failure._wpscPreserveFallbackFatal = true;
    return failure;
  }

  function currentPosition(document) {
    const range = endRange(document);
    return safeNumber(range && range.Start, safeNumber(range && range.End, 0));
  }

  function insertInlineText(document, value, targetRange) {
    const text = safeString(value);
    const range = targetRange || endRange(document);
    if (typeof range.InsertAfter === "function") range.InsertAfter(text);
    else range.Text = text;
    return range;
  }

  function nativeFields(document) {
    if (!Array.isArray(document._wpscNativeFields)) document._wpscNativeFields = [];
    return document._wpscNativeFields;
  }

  function trackNativeField(document, ownerNodeId, fieldKind, native, category) {
    nativeFields(document).push({
      ownerNodeId: ownerNodeId || "doc:native",
      fieldKind: fieldKind,
      native: native,
      category: category
    });
    return native;
  }

  function addNativeField(document, code, ownerNodeId, fieldKind, category, failureCode) {
    let field;
    try {
      field = document.Fields.Add(endRange(document), -1, code, true);
    } catch (error) {
      throw nativeError(failureCode || "FIELD_REFRESH_FAILED");
    }
    return trackNativeField(document, ownerNodeId, fieldKind, field, category);
  }

  function validateNumbering(numbering) {
    if (!numbering || !hasOwn(NATIVE_SEQUENCE_KIND, numbering.sequenceId)) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    if (numbering.mode !== "global" && numbering.mode !== "chapter") {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    if (numbering.mode === "chapter" &&
        (numbering.chapterStyleLevel !== 1 || numbering.resetLevel !== 1)) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
  }

  function nativeHeadingStyleName(document, level) {
    const style = collectionItem(document.Styles, -1 - level);
    const name = style && (style.NameLocal || style.Name);
    if (typeof name !== "string" || !name || name.indexOf('"') !== -1) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    return name;
  }

  function addNativeNumberShell(document, numbering, bookmarkName, ownerNodeId) {
    validateNumbering(numbering);
    insertInlineText(document, numbering.prefix);
    const numberStart = currentPosition(document);
    if (numbering.mode === "chapter") {
      // WPS for macOS interprets numeric STYLEREF style arguments as literal
      // style names. Translate the closed level-1 descriptor to the stable
      // built-in English style ID used by this executor.
      // WPS field lookup uses the localized UI name even though styles.xml
      // persists an English built-in name. Resolve the locale-independent
      // built-in style (-2 for H1), then emit its NameLocal at runtime.
      const headingStyle = nativeHeadingStyleName(document, 1);
      addNativeField(document, 'STYLEREF "' + headingStyle + '" \\s', ownerNodeId, "STYLEREF", "numbering");
      insertInlineText(document, "-");
    }
    const sequenceCode = "SEQ " + numbering.sequenceId + " \\* ARABIC" +
      (numbering.mode === "chapter" ? " \\s 1" : "");
    addNativeField(
      document,
      sequenceCode,
      ownerNodeId,
      NATIVE_SEQUENCE_KIND[numbering.sequenceId],
      "numbering"
    );
    const numberEnd = currentPosition(document);
    if (bookmarkName !== undefined && bookmarkName !== null) {
      if (!/^wpsc_(fig|tab|eq)_[a-z0-9]{24}$/.test(bookmarkName)) {
        throw nativeError("FIELD_REFRESH_FAILED");
      }
      try {
        document.Bookmarks.Add(bookmarkName, document.Range(numberStart, numberEnd));
      } catch (error) {
        throw nativeError("FIELD_REFRESH_FAILED");
      }
    }
    if (numbering.suffix) insertInlineText(document, numbering.suffix);
    return {start: numberStart, end: numberEnd};
  }

  function addNativeCaption(document, args, ownerNodeId, keepWithNext) {
    const start = currentPosition(document);
    addNativeNumberShell(document, args.numbering, args.bookmarkName, ownerNodeId);
    if (args.caption) insertInlineText(document, " " + safeString(args.caption));
    const paragraph = document.Range(start, currentPosition(document));
    if (paragraph.ParagraphFormat) {
      paragraph.ParagraphFormat.Alignment = 1;
      paragraph.ParagraphFormat.KeepTogether = -1;
      paragraph.ParagraphFormat.KeepWithNext = keepWithNext ? -1 : 0;
    }
    insertInlineText(document, "\r");
    return paragraph;
  }

  function rollbackMutation(document, start, end) {
    try {
      const stop = end === undefined ? currentPosition(document) : safeNumber(end, start);
      if (stop > start) document.Range(start, stop).Delete();
    } catch (error) {
      throw nativeError("LOCAL_MUTATION_ROLLBACK_FAILED");
    }
  }

  function checkpointRecoverableMutation(document) {
    try {
      if (!document || typeof document.Range !== "function") {
        throw nativeError("LOCAL_MUTATION_CHECKPOINT_FAILED");
      }
      const content = document.Content;
      const rawEnd = content && Number(content.End);
      if (!Number.isFinite(rawEnd)) {
        throw nativeError("LOCAL_MUTATION_CHECKPOINT_FAILED");
      }
      const position = Math.max(0, rawEnd - 1);
      const checkpointRange = document.Range(position, position);
      if (!checkpointRange || typeof checkpointRange.Delete !== "function") {
        throw nativeError("LOCAL_MUTATION_CHECKPOINT_FAILED");
      }
      const start = Number(checkpointRange.Start);
      if (!Number.isFinite(start) || start < 0) {
        throw nativeError("LOCAL_MUTATION_CHECKPOINT_FAILED");
      }
      return start;
    } catch (error) {
      throw nativeError("LOCAL_MUTATION_CHECKPOINT_FAILED");
    }
  }

  function addExplicitOrientationSection(document, landscape) {
    const range = endRange(document);
    if (typeof range.InsertBreak === "function") range.InsertBreak(2);
    const sections = document.Sections;
    const section = sections ? collectionItem(sections, sections.Count) : null;
    const setup = section && section.PageSetup ? section.PageSetup : document.PageSetup;
    if (!setup) throw nativeError("EXECUTION_ABORTED");
    setup.Orientation = landscape ? 1 : 0;
  }

  function addFigureChild(document, child, locator, ownerNodeId, targetRange) {
    const start = currentPosition(document);
    let shape;
    try {
      const rawBefore = document.InlineShapes && Number(document.InlineShapes.Count);
      const before = Number.isInteger(rawBefore) && rawBefore >= 0 ? rawBefore : -1;
      shape = document.InlineShapes.AddPicture(locator, false, true, targetRange || endRange(document));
      const rawAfter = document.InlineShapes && Number(document.InlineShapes.Count);
      const after = Number.isInteger(rawAfter) && rawAfter >= 0 ? rawAfter : -1;
      if ((shape === null || shape === undefined) && before >= 0 && after > before) {
        shape = collectionItem(document.InlineShapes, after);
      }
      if (shape === null || shape === undefined) throw nativeError("IMAGE_INSERT_FAILED");
    } catch (error) {
      rollbackMutation(document, start);
      throw nativeError("IMAGE_INSERT_FAILED");
    }
    try {
      shape.Width = safeNumber(child.displayWidthPt, shape.Width);
      shape.Height = safeNumber(child.displayHeightPt, shape.Height);
      if (hasOwn(shape, "AlternativeText")) shape.AlternativeText = ownerNodeId || "";
      if (shape.Range && shape.Range.ParagraphFormat) {
        shape.Range.ParagraphFormat.Alignment = 1;
        shape.Range.ParagraphFormat.KeepTogether = -1;
        shape.Range.ParagraphFormat.KeepWithNext = -1;
      }
      if (!targetRange) insertInlineText(document, "\r");
      return shape;
    } catch (error) {
      rollbackMutation(document, start);
      throw error;
    }
  }

  function addFigureNotice(document, code) {
    addDegradationNotice(document, {
      code: code,
      fallbackText: "[" + code + "]",
      placement: "block"
    });
  }

  function renderFigureStack(document, children, resources, context, retryOnce) {
    let degraded = false;
    children.forEach(function (child) {
      if (child.plannedDegradation) {
        const planned = child.plannedDegradation;
        addDegradationNotice(document, {
          code: planned.code,
          fallbackText: planned.fallback,
          placement: planned.placement || "block"
        });
        context.childResults.push({nodeId: child.nodeId, status: "degraded", issueCode: planned.code});
        appendIssueOnce(context.issues, {
          code: planned.code,
          message: "Figure child used its planned fallback",
          placement: planned.placement || "block",
          nodeId: context.ownerNodeId
        });
        return;
      }
      const locator = resources[child.resourceId];
      if (typeof locator !== "string" || locator.length === 0) {
        throw nativeError("RESOURCE_HASH_MISMATCH");
      }
      try {
        addFigureChild(document, child, locator, context.ownerNodeId, null);
        context.childResults.push({nodeId: child.nodeId, status: "applied"});
      } catch (error) {
        if (error.code !== "IMAGE_INSERT_FAILED" || !retryOnce) throw error;
        try {
          addFigureChild(document, child, locator, context.ownerNodeId, null);
          context.childResults.push({nodeId: child.nodeId, status: "applied"});
        } catch (retryError) {
          if (retryError.code !== "IMAGE_INSERT_FAILED") throw retryError;
          addFigureNotice(document, "IMAGE_INSERT_FAILED");
          context.childResults.push({nodeId: child.nodeId, status: "degraded", issueCode: "IMAGE_INSERT_FAILED"});
        }
        degraded = true;
      }
    });
    return degraded;
  }

  function createFigureColumns(document, children, resources, context) {
    let table;
    try {
      table = document.Tables.Add(endRange(document), 1, 3);
      try { table.AllowAutoFit = false; } catch (error) { /* ignore */ }
      [children[0].displayWidthPt, 12, children[1].displayWidthPt].forEach(function (width, index) {
        const column = collectionItem(table.Columns, index + 1) || table.Columns(index + 1);
        if (typeof column.SetWidth === "function") column.SetWidth(width, 0);
        else column.Width = width;
      });
      for (let borderId = -6; borderId <= -1; borderId += 1) {
        tableBorder(table, borderId).LineStyle = 0;
      }
      if (table.Range && table.Range.ParagraphFormat) {
        table.Range.ParagraphFormat.KeepTogether = -1;
        table.Range.ParagraphFormat.KeepWithNext = -1;
      }
    } catch (error) {
      try { if (table && typeof table.Delete === "function") table.Delete(); } catch (deleteError) { /* ignore */ }
      throw nativeError("IMAGE_INSERT_FAILED");
    }
    try {
      children.forEach(function (child, index) {
        const locator = resources[child.resourceId];
        if (typeof locator !== "string" || locator.length === 0) {
          throw nativeError("RESOURCE_HASH_MISMATCH");
        }
        const cell = table.Cell(1, index === 0 ? 1 : 3);
        const target = cell.Range.Duplicate;
        target.End = target.Start;
        if (target.ParagraphFormat) {
          target.ParagraphFormat.Alignment = 1;
          target.ParagraphFormat.KeepTogether = -1;
          target.ParagraphFormat.KeepWithNext = -1;
        }
        addFigureChild(document, child, locator, context.ownerNodeId, target);
      });
      return table;
    } catch (error) {
      try {
        if (table && typeof table.Delete === "function") table.Delete();
      } catch (deleteError) {
        throw nativeError("LOCAL_MUTATION_ROLLBACK_FAILED");
      }
      throw error;
    }
  }

  function addCaptionedFigureNative(document, args, resources, context) {
    if (args.keepWithCaption !== true) throw nativeError("EXECUTION_ABORTED");
    const landscape = args.orientation === "landscape";
    if (landscape) addExplicitOrientationSection(document, true);
    let degraded = false;
    try {
      if (args.layout === "columns") {
        const start = currentPosition(document);
        try {
          createFigureColumns(document, args.children, resources, context);
          args.children.forEach(function (child) {
            context.childResults.push({nodeId: child.nodeId, status: "applied"});
          });
          insertInlineText(document, "\r");
        } catch (error) {
          if (error.code !== "IMAGE_INSERT_FAILED" || context.controllerOwned) throw error;
          rollbackMutation(document, start);
          context.childResults.length = 0;
          renderFigureStack(document, args.children, resources, context, false);
          degraded = true;
        }
      } else {
        degraded = renderFigureStack(
          document, args.children, resources, context, !context.controllerOwned
        );
      }
      if (degraded) appendIssueOnce(context.issues, {
        code: "IMAGE_INSERT_FAILED", message: "Figure used deterministic stack recovery",
        placement: "block", nodeId: context.ownerNodeId
      });
      if (args.caption) addNativeCaption(document, args, context.ownerNodeId, false);
    } finally {
      if (landscape) addExplicitOrientationSection(document, false);
    }
  }

  function addCaptionedFigureFallback(document, args, resources, context, code) {
    const landscape = args.orientation === "landscape";
    if (landscape) addExplicitOrientationSection(document, true);
    try {
      (args.children || []).forEach(function (child) {
        if (child.plannedDegradation) {
          const planned = child.plannedDegradation;
          addDegradationNotice(document, {
            code: planned.code, fallbackText: planned.fallback,
            placement: planned.placement || "block"
          });
          context.childResults.push({nodeId: child.nodeId, status: "degraded", issueCode: planned.code});
          appendIssueOnce(context.issues, {
            code: planned.code, message: "Figure child used its planned fallback",
            placement: planned.placement || "block", nodeId: context.ownerNodeId
          });
          return;
        }
        const locator = resources[child.resourceId];
        if (typeof locator !== "string" || locator.length === 0) {
          throw nativeError("RESOURCE_HASH_MISMATCH");
        }
        try {
          addFigureChild(document, child, locator, context.ownerNodeId, null);
          context.childResults.push({nodeId: child.nodeId, status: "applied"});
        } catch (error) {
          if (error.code !== "IMAGE_INSERT_FAILED") throw error;
          addFigureNotice(document, code);
          context.childResults.push({nodeId: child.nodeId, status: "degraded", issueCode: code});
        }
      });
      if (args.caption) addNativeCaption(document, args, context.ownerNodeId, false);
    } finally {
      if (landscape) addExplicitOrientationSection(document, false);
    }
  }

  function tableBorder(table, id) {
    if (typeof table.Borders === "function") return table.Borders(id);
    return collectionItem(table.Borders, id);
  }

  function tableRows(table, index) {
    if (typeof table.Rows === "function") return table.Rows(index);
    return collectionItem(table.Rows, index);
  }

  function applyTableBorders(table, spec) {
    const mapping = {top: -1, left: -2, bottom: -3, right: -4, insideHorizontal: -5, insideVertical: -6};
    Object.keys(mapping).forEach(function (name) {
      const points = safeNumber(spec[name], 0);
      const border = tableBorder(table, mapping[name]);
      border.LineStyle = points === 0 ? 0 : 1;
      if (points !== 0) border.LineWidth = points === 1.5 ? 12 : (points === 0.75 ? 6 : 2);
    });
    const headerPoints = safeNumber(spec.headerBottom, 0);
    const headerBorder = tableBorder(tableRows(table, 1), -3);
    headerBorder.LineStyle = headerPoints === 0 ? 0 : 1;
    if (headerPoints !== 0) headerBorder.LineWidth = headerPoints === 1.5 ? 12 : (headerPoints === 0.75 ? 6 : 2);
  }

  function createNativeTable(document, args) {
    const data = [args.headers].concat(args.rows || []);
    let table;
    try {
      table = document.Tables.Add(endRange(document), data.length, args.headers.length);
    } catch (error) {
      throw nativeError("TABLE_INSERT_FAILED");
    }
    const alignmentCodes = {left: 0, center: 1, right: 2};
    try {
      data.forEach(function (row, rowIndex) {
        row.forEach(function (value, columnIndex) {
          const cell = table.Cell(rowIndex + 1, columnIndex + 1);
          cell.Range.Text = safeString(value);
          const format = cell.Range.ParagraphFormat;
          format.FirstLineIndent = safeNumber(args.cellIndentPt, 0);
          format.LeftIndent = 0;
          format.RightIndent = 0;
          format.Alignment = alignmentCodes[args.alignments[columnIndex]];
        });
      });
      table.Rows.AllowBreakAcrossPages = args.allowRowSplit ? -1 : 0;
      if (args.repeatHeader) tableRows(table, 1).HeadingFormat = -1;
      applyTableBorders(table, args.borderSpec);
      applyTableCellMetadata(table, args);
    } catch (error) {
      throw nativeError("TABLE_STYLE_APPLY_FAILED");
    }
    const mergePageAnchors = [];
    try {
      (args.merges || []).forEach(function (merge) {
        const firstRange = table.Cell(merge.top, merge.left).Range;
        const lastRange = table.Cell(merge.bottom, merge.right).Range;
        mergePageAnchors.push({
          merge: merge,
          first: firstRange.Duplicate || firstRange,
          last: lastRange.Duplicate || lastRange
        });
        table.Cell(merge.top, merge.left).Merge(table.Cell(merge.bottom, merge.right));
      });
    } catch (error) {
      throw nativeError("TABLE_MERGE_APPLY_FAILED");
    }
    table._wpscMergePageAnchors = mergePageAnchors;
    insertInlineText(document, "\r");
    return table;
  }

  function applyTableCellMetadata(table, args) {
    // Citation cells already contain their resolved static [n]. Metadata is
    // deliberately non-mutating so visible cell data cannot be duplicated or
    // lost. Only explicit unresolved degradations receive local styling.
    (args.cellCitations || []).forEach(function (citation) {
      const cell = table.Cell(citation.row, citation.column);
      if (!cell || !cell.Range) throw nativeError("TABLE_STYLE_APPLY_FAILED");
    });
    (args.cellDegradations || []).forEach(function (degradation) {
      const cell = table.Cell(degradation.row, degradation.column);
      const range = cell && cell.Range;
      if (!range) throw nativeError("TABLE_STYLE_APPLY_FAILED");
      try {
        if (range.Font) {
          range.Font.Italic = -1;
          range.Font.Color = colorFromHex("#9C0006");
        }
        if (range.Shading) {
          range.Shading.BackgroundPatternColor = colorFromHex("#FCE8E6");
        }
      } catch (error) {
        throw nativeError("TABLE_STYLE_APPLY_FAILED");
      }
    });
  }

  function gridTableArgs(args) {
    const clone = Object.assign({}, args);
    clone.borderSpec = {
      top: 0.75, bottom: 0.75, headerBottom: 0.75,
      left: 0.75, right: 0.75, insideHorizontal: 0.75, insideVertical: 0.75
    };
    clone.merges = [];
    clone.allowRowSplit = true;
    return clone;
  }

  function tableOverflowGroup(table, merges) {
    const anchors = Array.isArray(table._wpscMergePageAnchors)
      ? table._wpscMergePageAnchors : [];
    for (let index = 0; index < anchors.length; index += 1) {
      const merge = anchors[index].merge;
      if (merge.top < 2 || merge.bottom <= merge.top) continue;
      const first = anchors[index].first;
      const last = anchors[index].last;
      try {
        if (safeNumber(first.Information(3), 0) !== safeNumber(last.Information(3), 0)) return true;
      } catch (error) {
        throw nativeError("TABLE_ROW_FORCED_SPLIT");
      }
    }
    return false;
  }

  function addTableTextFallback(document, args) {
    [args.headers].concat(args.rows || []).forEach(function (row) {
      insertInlineText(document, row.map(safeString).join(" | ") + "\r");
    });
  }

  function addNativeTableNotice(document, code) {
    const start = currentPosition(document);
    // A table recovery cannot depend on another Tables.Add call. Preserve a
    // minimal styled block notice beside the fallback grid/text instead.
    insertText(document, "[" + safeString(code) + "]", "Body Text", {
      italic: true, color: "#9C0006", spaceAfter: 3, outlineLevel: 10
    });
    const range = document.Range(start, currentPosition(document));
    if (range.ParagraphFormat) {
      range.ParagraphFormat.KeepTogether = -1;
      range.ParagraphFormat.KeepWithNext = -1;
    }
  }

  function addSemanticTableNative(document, args, resources, context) {
    void resources;
    if (args.keepCaptionWithFirstRow !== true) throw nativeError("EXECUTION_ABORTED");
    const landscape = args.orientation === "landscape";
    if (landscape) addExplicitOrientationSection(document, true);
    try {
      if (args.caption) addNativeCaption(document, args, context.ownerNodeId, true);
      (args.plannedDegradation || []).forEach(function (planned) {
        addNativeTableNotice(document, planned.code);
        appendIssueOnce(context.issues, {
          code: planned.code,
          message: "Table used its planned fallback",
          placement: planned.placement || "block",
          nodeId: context.ownerNodeId
        });
      });
      const tableStart = currentPosition(document);
      let table;
      try {
        table = createNativeTable(document, args);
      } catch (error) {
        if (context.controllerOwned || !RECOVERABLE_TABLE_CODES[error.code]) throw error;
        rollbackMutation(document, tableStart);
        addNativeTableNotice(document, error.code);
        const gridStart = currentPosition(document);
        try {
          table = createNativeTable(document, gridTableArgs(args));
          appendIssueOnce(context.issues, {code: error.code, message: "Table used deterministic grid fallback", placement: "block", nodeId: context.ownerNodeId});
        } catch (gridError) {
          if (!RECOVERABLE_TABLE_CODES[gridError.code]) throw gridError;
          rollbackMutation(document, gridStart);
          addTableTextFallback(document, args);
          appendIssueOnce(context.issues, {code: "TABLE_INSERT_FAILED", message: "Table used deterministic text fallback", placement: "block", nodeId: context.ownerNodeId});
          return;
        }
      }
      if (tableOverflowGroup(table, args.merges || [])) {
        if (context.controllerOwned) throw nativeError("TABLE_ROW_FORCED_SPLIT");
        rollbackMutation(document, tableStart);
        addNativeTableNotice(document, "TABLE_ROW_FORCED_SPLIT");
        appendIssueOnce(context.issues, {code: "TABLE_ROW_FORCED_SPLIT", message: "Vertical merge group rendered as splittable grid", placement: "block", nodeId: context.ownerNodeId});
        const gridStart = currentPosition(document);
        try {
          createNativeTable(document, gridTableArgs(args));
        } catch (gridError) {
          if (!RECOVERABLE_TABLE_CODES[gridError.code]) throw gridError;
          rollbackMutation(document, gridStart);
          addTableTextFallback(document, args);
          appendIssueOnce(context.issues, {code: "TABLE_INSERT_FAILED", message: "Overflow grid used deterministic text fallback", placement: "block", nodeId: context.ownerNodeId});
        }
      }
    } finally {
      if (landscape) addExplicitOrientationSection(document, false);
    }
  }

  function addSemanticTableFallback(document, args, resources, context, code) {
    void resources;
    const landscape = args.orientation === "landscape";
    if (landscape) addExplicitOrientationSection(document, true);
    try {
      if (args.caption) addNativeCaption(document, args, context.ownerNodeId, true);
      (args.plannedDegradation || []).forEach(function (planned) {
        addNativeTableNotice(document, planned.code);
        appendIssueOnce(context.issues, {
          code: planned.code, message: "Table used its planned fallback",
          placement: planned.placement || "block", nodeId: context.ownerNodeId
        });
      });
      addNativeTableNotice(document, code);
      const gridStart = currentPosition(document);
      try {
        createNativeTable(document, gridTableArgs(args));
      } catch (error) {
        if (!RECOVERABLE_TABLE_CODES[error.code]) throw error;
        rollbackMutation(document, gridStart);
        addTableTextFallback(document, args);
      }
    } finally {
      if (landscape) addExplicitOrientationSection(document, false);
    }
  }

  function addEquationNumberNative(document, args, resources, context) {
    void resources;
    const start = currentPosition(document);
    insertInlineText(document, args.source || args.fallbackText || "");
    insertInlineText(document, "\t");
    addNativeNumberShell(document, args.numbering, args.bookmarkName, context.ownerNodeId);
    const paragraph = document.Range(start, currentPosition(document));
    if (paragraph.ParagraphFormat) {
      paragraph.ParagraphFormat.Alignment = 2;
      paragraph.ParagraphFormat.KeepTogether = -1;
    }
    insertInlineText(document, "\r");
  }

  function formulaLayoutSpec(document) {
    let setup = document && document.PageSetup;
    const sections = document && document.Sections;
    const sectionCount = sections && Number(sections.Count);
    if (Number.isInteger(sectionCount) && sectionCount > 0) {
      const section = collectionItem(sections, sectionCount);
      if (section && section.PageSetup) setup = section.PageSetup;
    }
    const pageWidth = Number(setup && setup.PageWidth);
    const leftMargin = Number(setup && setup.LeftMargin);
    const rightMargin = Number(setup && setup.RightMargin);
    const usableWidth = pageWidth - leftMargin - rightMargin;
    if (!Number.isFinite(pageWidth) || !Number.isFinite(leftMargin) ||
        !Number.isFinite(rightMargin) || pageWidth <= 0 ||
        leftMargin < 0 || rightMargin < 0 || usableWidth <= 72) {
      throw nativeError("CAPABILITY_MISMATCH");
    }
    return {center: usableWidth / 2, right: usableWidth};
  }

  function beginFormulaLayout(document) {
    const start = currentPosition(document);
    const layout = formulaLayoutSpec(document);
    const probe = document.Range(start, start);
    const format = probe && probe.ParagraphFormat;
    if (!format || !format.TabStops || typeof format.TabStops.Add !== "function") {
      throw nativeError("CAPABILITY_MISMATCH");
    }
    insertInlineText(document, "\t");
    return {start: start, center: layout.center, right: layout.right};
  }

  function finishFormulaLayout(document, layout) {
    const paragraph = document.Range(layout.start, currentPosition(document));
    const format = paragraph && paragraph.ParagraphFormat;
    if (!format || !format.TabStops || typeof format.TabStops.Add !== "function") {
      throw nativeError("CAPABILITY_MISMATCH");
    }
    format.Alignment = 0;
    format.LeftIndent = 0;
    format.RightIndent = 0;
    format.FirstLineIndent = 0;
    format.KeepTogether = -1;
    format.TabStops.Add(layout.center, 1, 0);
    format.TabStops.Add(layout.right, 2, 0);
  }

  function addFormulaSourceTerminalNotice(document, args, code, context) {
    const layout = beginFormulaLayout(document);
    addInlineDegradation(document, {
      code: code,
      fallbackText: args.fallbackText || ""
    });
    insertInlineText(document, "\t");
    addNativeNumberShell(document, args.numbering, args.bookmarkName, context.ownerNodeId);
    finishFormulaLayout(document, layout);
    insertInlineText(document, "\r");
  }

  function addEquationNativeM4(document, args, resources, context) {
    const content = args.content || {};
    if (content.plannedDegradation) {
      const planned = content.plannedDegradation;
      addFormulaNativeFallback(document, args, resources, context, planned.code);
      appendIssueOnce(context.issues, {
        code: planned.code,
        message: "Formula used its planned content fallback",
        placement: planned.placement,
        nodeId: context.ownerNodeId,
        stage: "preflight",
        fallback: planned.fallbackKind,
        recoverable: true
      });
      emitFormulaResourceDegradation(document, args, context);
      return;
    }
    const nativeMath = content.nativeMath || {};
    if (nativeMath.syntax !== "wps-linear-v1") {
      throw nativeError("CAPABILITY_MISMATCH");
    }
    if (!document.OMaths || typeof document.OMaths.Add !== "function") {
      throw nativeError("CAPABILITY_MISMATCH");
    }
    const layout = beginFormulaLayout(document);
    const start = currentPosition(document);
    const linearText = safeString(nativeMath.linearText);
    insertInlineText(document, linearText);
    const mathEnd = currentPosition(document);
    const before = Number(document.OMaths.Count);
    if (!Number.isInteger(before) || before < 0) {
      throw nativeError("CAPABILITY_MISMATCH");
    }
    const addedRange = document.OMaths.Add(document.Range(start, mathEnd));
    const after = Number(document.OMaths.Count);
    if (!addedRange || after !== before + 1) {
      throw nativeError("EQUATION_INSERT_FAILED");
    }
    const addedMaths = addedRange.OMaths;
    if (!addedMaths || Number(addedMaths.Count) !== 1) {
      throw nativeError("EQUATION_INSERT_FAILED");
    }
    const math = collectionItem(addedMaths, 1);
    const documentMath = collectionItem(document.OMaths, after);
    if (!math || !documentMath || typeof math.BuildUp !== "function") {
      throw nativeError("CAPABILITY_MISMATCH");
    }
    const addedStart = Number(addedRange.Start);
    const addedEnd = Number(addedRange.End);
    const mathStart = Number(math.Range.Start);
    const mathFinish = Number(math.Range.End);
    const documentStart = Number(documentMath.Range.Start);
    const documentEnd = Number(documentMath.Range.End);
    if (!Number.isFinite(addedStart) || !Number.isFinite(addedEnd) ||
        !Number.isFinite(mathStart) || !Number.isFinite(mathFinish) ||
        addedEnd <= addedStart || addedStart < start || addedEnd > mathEnd ||
        mathFinish <= mathStart || mathStart < addedStart || mathFinish > addedEnd ||
        documentStart !== mathStart || documentEnd !== mathFinish) {
      throw nativeError("EQUATION_INSERT_FAILED");
    }
    math.BuildUp();
    const builtAddedStart = Number(addedRange.Start);
    const builtAddedEnd = Number(addedRange.End);
    const builtLocalCount = Number(addedMaths.Count);
    const builtGlobalCount = Number(document.OMaths.Count);
    const builtLocalMath = collectionItem(addedMaths, 1);
    const builtGlobalMath = collectionItem(document.OMaths, after);
    const builtStart = Number(builtLocalMath.Range.Start);
    const builtEnd = Number(builtLocalMath.Range.End);
    const builtGlobalStart = Number(builtGlobalMath.Range.Start);
    const builtGlobalEnd = Number(builtGlobalMath.Range.End);
    if (builtLocalMath !== math || builtGlobalMath !== documentMath ||
        builtLocalMath !== builtGlobalMath || builtLocalCount !== 1 ||
        builtGlobalCount !== after ||
        builtAddedStart !== addedStart || builtAddedEnd !== addedEnd ||
        builtEnd <= builtStart || builtStart < builtAddedStart ||
        builtEnd > builtAddedEnd || builtGlobalStart !== builtStart ||
        builtGlobalEnd !== builtEnd) {
      throw nativeError("EQUATION_INSERT_FAILED");
    }
    insertInlineText(document, "\t");
    addNativeNumberShell(document, args.numbering, args.bookmarkName, context.ownerNodeId);
    finishFormulaLayout(document, layout);
    insertInlineText(document, "\r");

    emitFormulaResourceDegradation(document, args, context);
  }

  function emitFormulaResourceDegradation(document, args, context) {
    const planned = (args.fallbackResource || {}).fallbackResourcePlannedDegradation;
    if (planned) {
      addDegradationNotice(document, planned);
      appendIssueOnce(context.issues, {
        code: "FORMULA_FALLBACK_IMAGE_UNAVAILABLE",
        message: "Optional formula fallback image is unavailable",
        placement: planned.placement,
        nodeId: context.ownerNodeId,
        stage: "native",
        fallback: "none",
        recoverable: true
      });
    }
  }

  function addFormulaNativeFallback(document, args, resources, context, code) {
    const fallback = args.fallbackResource || {};
    const resourceId = fallback.fallbackResourceId;
    if (resourceId !== undefined) {
      const locator = resources[resourceId];
      if (typeof locator !== "string" || !locator) {
        throw nativeError("RESOURCE_HASH_MISMATCH");
      }
      if (!document.InlineShapes || typeof document.InlineShapes.AddPicture !== "function") {
        throw nativeError("CAPABILITY_MISMATCH");
      }
      const layout = beginFormulaLayout(document);
      const imageStart = layout.start;
      try {
        const shape = document.InlineShapes.AddPicture(
          locator, false, true, endRange(document)
        );
        const shapeRange = shape && shape.Range;
        const shapeStart = Number(shapeRange && shapeRange.Start);
        const shapeEnd = Number(shapeRange && shapeRange.End);
        if (!shapeRange || !Number.isFinite(shapeStart) ||
            !Number.isFinite(shapeEnd) || shapeEnd <= shapeStart) {
          throw nativeError("IMAGE_INSERT_FAILED");
        }
        if (!shapeRange.ParagraphFormat) throw nativeError("CAPABILITY_MISMATCH");
        shapeRange.ParagraphFormat.KeepTogether = -1;
        addInlineDegradation(document, {
          code: code, fallbackText: "formula image fallback"
        });
        insertInlineText(document, "\t");
        addNativeNumberShell(document, args.numbering, args.bookmarkName, context.ownerNodeId);
        finishFormulaLayout(document, layout);
        insertInlineText(document, "\r");
        return;
      } catch (error) {
        try {
          rollbackMutation(document, imageStart);
        } catch (rollbackError) {
          throw preserveFallbackFatal(rollbackError);
        }
        if (!error || error.code !== "IMAGE_INSERT_FAILED") {
          throw preserveFallbackFatal(error);
        }
      }
    }
    try {
      addFormulaSourceTerminalNotice(document, args, code, context);
    } catch (error) {
      throw preserveFallbackFatal(error);
    }
  }

  function addCitationParagraph(document, args, resources, context) {
    void resources;
    const paragraphStart = currentPosition(document);
    (args.runs || []).forEach(function (run) {
      if (run.type === "text") {
        insertInlineText(document, run.text);
      } else if (run.type === "citation") {
        insertInlineText(document, run.fallbackText);
        context.childResults.push({nodeId: run.nodeId, status: "applied"});
      } else if (run.type === "degradation") {
        addLiteralInlineDegradation(document, run.fallbackText);
        appendIssueOnce(context.issues, {
          code: run.code,
          message: "Citation used its planned fallback",
          placement: "inline",
          nodeId: run.nodeId,
          stage: "preflight",
          fallback: "inline",
          recoverable: true
        });
        context.childResults.push({
          nodeId: run.nodeId, status: "degraded", issueCode: run.code
        });
      } else if (run.type === "reference") {
        insertInlineText(document, run.prefix);
        addNativeField(
          document, "REF " + run.bookmarkName + " \\h",
          context.ownerNodeId, "REF", "reference", "CROSS_REFERENCE_FAILED"
        );
        if (run.suffix) insertInlineText(document, run.suffix);
      } else {
        throw nativeError("CAPABILITY_MISMATCH");
      }
    });
    if (args.listFormatting) {
      const paragraph = document.Range(paragraphStart, currentPosition(document));
      const format = paragraph && paragraph.ParagraphFormat;
      if (!format) throw nativeError("CAPABILITY_MISMATCH");
      const indent = safeNumber(args.listFormatting.indentPt, 24);
      format.LeftIndent = indent;
      format.FirstLineIndent = -indent;
      format.SpaceBefore = 0;
      format.SpaceAfter = 3;
    }
    insertInlineText(document, "\r");
  }

  function addBibliographyNative(document, args) {
    if (!args || !Array.isArray(args.entries)) {
      throw nativeError("CONFIGURATION_INVALID");
    }
    const structured = args.schemaVersion === 1;
    const entries = args.entries;
    if (structured && entries.length === 0) {
      throw nativeError("CONFIGURATION_INVALID");
    }
    entries.forEach(function (entry) {
      const start = currentPosition(document);
      const text = structured
        ? "[" + entry.number + "] " + safeString(entry.text)
        : safeString(entry);
      insertInlineText(document, text);
      const paragraph = document.Range(start, currentPosition(document));
      if (!paragraph || !paragraph.ParagraphFormat) {
        throw nativeError("BIBLIOGRAPHY_INSERT_FAILED");
      }
      const format = paragraph.ParagraphFormat;
      format.Alignment = 0;
      if (structured) {
        format.LeftIndent = safeNumber(args.leftIndentPt, 18);
        format.FirstLineIndent = -safeNumber(args.hangingIndentPt, 18);
        format.SpaceBefore = 0;
        format.SpaceAfter = safeNumber(args.spaceAfterPt, 6);
        format.KeepTogether = -1;
      }
      insertInlineText(document, "\r");
    });
  }

  function addCrossReferenceParagraph(document, args, resources, context) {
    void resources;
    const paragraphStart = currentPosition(document);
    let degraded = false;
    (args.runs || []).forEach(function (run) {
      if (run.type === "text") {
        insertInlineText(document, run.text);
        return;
      }
      insertInlineText(document, run.prefix);
      const start = currentPosition(document);
      try {
        addNativeField(
          document, "REF " + run.bookmarkName + " \\h",
          context.ownerNodeId, "REF", "reference", "CROSS_REFERENCE_FAILED"
        );
      } catch (error) {
        if (error.code !== "CROSS_REFERENCE_FAILED" || context.controllerOwned) throw error;
        rollbackMutation(document, start);
        insertInlineText(document, run.fallbackText);
        degraded = true;
      }
      if (run.suffix) insertInlineText(document, run.suffix);
    });
    if (args.listFormatting) {
      const paragraph = document.Range(paragraphStart, currentPosition(document));
      const style = getStyle(document, "List Paragraph");
      if (style) paragraph.Style = style;
      const format = paragraph.ParagraphFormat;
      if (format) {
        const indent = safeNumber(args.listFormatting.indentPt, 24);
        format.LeftIndent = indent;
        format.FirstLineIndent = -indent;
        format.SpaceBefore = 0;
        format.SpaceAfter = 3;
        if (format.TabStops && typeof format.TabStops.Add === "function") {
          format.TabStops.Add(indent);
        }
      }
    }
    insertInlineText(document, "\r");
    if (degraded) appendIssueOnce(context.issues, {
      code: "CROSS_REFERENCE_FAILED", message: "Cross-reference used inline fallback",
      placement: "inline", nodeId: context.ownerNodeId
    });
  }

  function addCrossReferenceFallback(document, args) {
    const paragraphStart = currentPosition(document);
    (args.runs || []).forEach(function (run) {
      if (run.type === "text") {
        insertInlineText(document, safeString(run.text));
        return;
      }
      insertInlineText(document, safeString(run.prefix));
      const start = currentPosition(document);
      const fallbackText = safeString(run.fallbackText);
      insertInlineText(document, fallbackText);
      const inserted = document.Range(start, start + fallbackText.length);
      if (inserted && inserted.Font) {
        inserted.Font.Italic = -1;
        inserted.Font.Color = colorFromHex("#9C0006");
      }
      if (inserted && inserted.Shading) {
        inserted.Shading.BackgroundPatternColor = colorFromHex("#FCE8E6");
      }
      insertInlineText(document, safeString(run.suffix));
    });
    if (args.listFormatting) {
      const paragraph = document.Range(paragraphStart, currentPosition(document));
      const format = paragraph && paragraph.ParagraphFormat;
      if (format) {
        const indent = safeNumber(args.listFormatting.indentPt, 24);
        format.LeftIndent = indent;
        format.FirstLineIndent = -indent;
        format.SpaceBefore = 0;
        format.SpaceAfter = 3;
      }
    }
    insertInlineText(document, "\r");
  }

  function insertCaptionIndexNative(document, args, resources, context) {
    void resources;
    if (args.sequenceId !== "WPSC_FIG" && args.sequenceId !== "WPSC_TAB") {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    if (args.title) {
      const title = insertText(document, args.title, args.titleStyleId, {});
      if (title.range && title.range.ParagraphFormat) title.range.ParagraphFormat.OutlineLevel = 10;
    }
    let index;
    try {
      index = document.TablesOfFigures.Add(endRange(document), args.sequenceId);
    } catch (error) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    trackNativeField(
      document,
      context.ownerNodeId,
      args.sequenceId === "WPSC_FIG" ? "TOF_FIG" : "TOF_TAB",
      index,
      "index"
    );
    insertInlineText(document, "\r");
  }

  function insertFigureIndex(document, args, resources, context) {
    if (args.sequenceId) return insertCaptionIndexNative(document, args, resources, context);
    document.TablesOfFigures.Add(endRange(document), "Figure");
  }

  function insertTableIndex(document, args, resources, context) {
    if (args.sequenceId) return insertCaptionIndexNative(document, args, resources, context);
    document.TablesOfFigures.Add(endRange(document), "Table");
  }

  function updateTrackedFields(document, kinds) {
    nativeFields(document).forEach(function (entry) {
      if (kinds[entry.fieldKind]) entry.native.Update();
    });
  }

  function requiredCollectionCount(collection) {
    if (!collection) throw nativeError("FIELD_REFRESH_FAILED");
    const count = Number(collection.Count);
    if (!Number.isInteger(count) || count < 0) throw nativeError("FIELD_REFRESH_FAILED");
    return count;
  }

  function repaginateAndUpdateNumbering(document) {
    if (typeof document.Repaginate !== "function") throw nativeError("FIELD_REFRESH_FAILED");
    document.Repaginate();
    const tracked = nativeFields(document);
    updateTrackedFields(document, {STYLEREF: true, SEQ_FIG: true, SEQ_TAB: true, SEQ_EQ: true});
    // Compatibility for M1/M2 plans that predate tracked native fields.
    if (tracked.length === 0) {
      if (!document.Fields || typeof document.Fields.Update !== "function") throw nativeError("FIELD_REFRESH_FAILED");
      document.Fields.Update();
    }
  }

  function refreshBookmarksAndReferences(document) {
    requiredCollectionCount(document.Bookmarks);
    updateTrackedFields(document, {REF: true});
  }

  function refreshIndexes(document) {
    const tocCount = requiredCollectionCount(document.TablesOfContents);
    const figureCount = requiredCollectionCount(document.TablesOfFigures);
    for (let index = 1; index <= tocCount; index += 1) {
      collectionItem(document.TablesOfContents, index).Update();
    }
    for (let index = 1; index <= figureCount; index += 1) {
      collectionItem(document.TablesOfFigures, index).Update();
    }
  }

  function sectionPageFields(document) {
    const entries = [];
    const sectionCount = requiredCollectionCount(document.Sections);
    for (let sectionIndex = 1; sectionIndex <= sectionCount; sectionIndex += 1) {
      const section = collectionItem(document.Sections, sectionIndex);
      ["Headers", "Footers"].forEach(function (storyName) {
        const stories = section && section[storyName];
        const storyCount = requiredCollectionCount(stories);
        for (let storyIndex = 1; storyIndex <= storyCount; storyIndex += 1) {
          const story = collectionItem(stories, storyIndex);
          if (story && (story.Exists === false || story.Exists === 0)) continue;
          const fields = story && story.Range && story.Range.Fields;
          const fieldCount = requiredCollectionCount(fields);
          for (let fieldIndex = 1; fieldIndex <= fieldCount; fieldIndex += 1) {
            const native = collectionItem(fields, fieldIndex);
            const code = safeString(native && native.Code && native.Code.Text)
              .trim().toUpperCase().split(/\s+/, 1)[0];
            if (code === "PAGE" || code === "NUMPAGES") {
              entries.push({
                ownerNodeId: "section:" + sectionIndex + "/" + storyName.toLowerCase() + ":" + storyIndex,
                fieldKind: code,
                native: native,
                category: "page"
              });
            }
          }
        }
      });
    }
    return entries;
  }

  function repaginateAndUpdatePageFields(document) {
    if (typeof document.Repaginate !== "function") throw nativeError("FIELD_REFRESH_FAILED");
    document.Repaginate();
    updateTrackedFields(document, {PAGE: true, NUMPAGES: true});
    const tracked = new Set(nativeFields(document).filter(function (entry) {
      return entry.fieldKind === "PAGE" || entry.fieldKind === "NUMPAGES";
    }).map(function (entry) { return entry.native; }));
    sectionPageFields(document).forEach(function (entry) {
      if (!tracked.has(entry.native)) entry.native.Update();
    });
  }

  const SHA256_CONSTANTS = Object.freeze([
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
  ]);

  function utf8Bytes(text) {
    const bytes = [];
    for (let index = 0; index < text.length; index += 1) {
      let codePoint = text.charCodeAt(index);
      if (codePoint >= 0xd800 && codePoint <= 0xdbff) {
        const next = text.charCodeAt(index + 1);
        if (next >= 0xdc00 && next <= 0xdfff) {
          codePoint = 0x10000 + ((codePoint - 0xd800) << 10) + (next - 0xdc00);
          index += 1;
        } else {
          codePoint = 0xfffd;
        }
      } else if (codePoint >= 0xdc00 && codePoint <= 0xdfff) {
        codePoint = 0xfffd;
      }
      if (codePoint <= 0x7f) bytes.push(codePoint);
      else if (codePoint <= 0x7ff) {
        bytes.push(0xc0 | (codePoint >>> 6), 0x80 | (codePoint & 0x3f));
      } else if (codePoint <= 0xffff) {
        bytes.push(0xe0 | (codePoint >>> 12), 0x80 | ((codePoint >>> 6) & 0x3f), 0x80 | (codePoint & 0x3f));
      } else {
        bytes.push(0xf0 | (codePoint >>> 18), 0x80 | ((codePoint >>> 12) & 0x3f),
          0x80 | ((codePoint >>> 6) & 0x3f), 0x80 | (codePoint & 0x3f));
      }
    }
    return bytes;
  }

  function rotateRight(value, bits) {
    return (value >>> bits) | (value << (32 - bits));
  }

  function hashVisible(value) {
    const text = safeString(value).replace(/\r\n?/g, "\n").normalize("NFC");
    const bytes = utf8Bytes(text);
    const bitLength = bytes.length * 8;
    bytes.push(0x80);
    while (bytes.length % 64 !== 56) bytes.push(0);
    const high = Math.floor(bitLength / 0x100000000);
    const low = bitLength >>> 0;
    for (let shift = 24; shift >= 0; shift -= 8) bytes.push((high >>> shift) & 0xff);
    for (let shift = 24; shift >= 0; shift -= 8) bytes.push((low >>> shift) & 0xff);

    const state = [
      0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
      0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19
    ];
    const words = new Array(64);
    for (let offset = 0; offset < bytes.length; offset += 64) {
      for (let index = 0; index < 16; index += 1) {
        const base = offset + index * 4;
        words[index] = ((bytes[base] << 24) | (bytes[base + 1] << 16) |
          (bytes[base + 2] << 8) | bytes[base + 3]) >>> 0;
      }
      for (let index = 16; index < 64; index += 1) {
        const left = words[index - 15];
        const right = words[index - 2];
        const sigma0 = rotateRight(left, 7) ^ rotateRight(left, 18) ^ (left >>> 3);
        const sigma1 = rotateRight(right, 17) ^ rotateRight(right, 19) ^ (right >>> 10);
        words[index] = (words[index - 16] + sigma0 + words[index - 7] + sigma1) >>> 0;
      }
      let a = state[0]; let b = state[1]; let c = state[2]; let d = state[3];
      let e = state[4]; let f = state[5]; let g = state[6]; let h = state[7];
      for (let index = 0; index < 64; index += 1) {
        const bigSigma1 = rotateRight(e, 6) ^ rotateRight(e, 11) ^ rotateRight(e, 25);
        const choice = (e & f) ^ (~e & g);
        const temp1 = (h + bigSigma1 + choice + SHA256_CONSTANTS[index] + words[index]) >>> 0;
        const bigSigma0 = rotateRight(a, 2) ^ rotateRight(a, 13) ^ rotateRight(a, 22);
        const majority = (a & b) ^ (a & c) ^ (b & c);
        const temp2 = (bigSigma0 + majority) >>> 0;
        h = g; g = f; f = e; e = (d + temp1) >>> 0;
        d = c; c = b; b = a; a = (temp1 + temp2) >>> 0;
      }
      state[0] = (state[0] + a) >>> 0; state[1] = (state[1] + b) >>> 0;
      state[2] = (state[2] + c) >>> 0; state[3] = (state[3] + d) >>> 0;
      state[4] = (state[4] + e) >>> 0; state[5] = (state[5] + f) >>> 0;
      state[6] = (state[6] + g) >>> 0; state[7] = (state[7] + h) >>> 0;
    }
    return state.map(function (word) { return word.toString(16).padStart(8, "0"); }).join("");
  }

  function nativeVisibleResult(native) {
    if (native && native.Result && native.Result.Text !== undefined) return native.Result.Text;
    if (native && native.Range && native.Range.Text !== undefined) return native.Range.Text;
    return "";
  }

  function nativeRangePageSpan(native) {
    const range = native && native.Range;
    const start = Number(range && range.Start);
    const end = Number(range && range.End);
    if (!Number.isInteger(start) || !Number.isInteger(end) || end < start) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    if (end === start) return 1;
    const first = typeof range.Duplicate === "function" ? range.Duplicate() : range.Duplicate;
    const last = typeof range.Duplicate === "function" ? range.Duplicate() : range.Duplicate;
    if (!first || !last || typeof first.SetRange !== "function" ||
        typeof last.SetRange !== "function" || typeof first.Information !== "function" ||
        typeof last.Information !== "function") {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    first.SetRange(start, start);
    last.SetRange(end - 1, end - 1);
    const firstPage = Number(first.Information(3));
    const lastPage = Number(last.Information(3));
    if (!Number.isInteger(firstPage) || !Number.isInteger(lastPage)) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    return Math.max(1, lastPage - firstPage + 1);
  }

  function snapshotFields(document) {
    if (typeof document.ComputeStatistics !== "function") throw nativeError("FIELD_REFRESH_FAILED");
    const totalPages = Number(document.ComputeStatistics(2));
    if (!Number.isFinite(totalPages) || totalPages < 0) throw nativeError("FIELD_REFRESH_FAILED");
    requiredCollectionCount(document.TablesOfContents);
    requiredCollectionCount(document.TablesOfFigures);
    const tocPageCount = nativeFields(document).filter(function (entry) {
      return entry.fieldKind === "TOC";
    }).reduce(function (total, entry) { return total + nativeRangePageSpan(entry.native); }, 0);
    const figureIndexPageCount = nativeFields(document).filter(function (entry) {
      return entry.fieldKind === "TOF_FIG";
    }).reduce(function (total, entry) { return total + nativeRangePageSpan(entry.native); }, 0);
    const tableIndexPageCount = nativeFields(document).filter(function (entry) {
      return entry.fieldKind === "TOF_TAB";
    }).reduce(function (total, entry) { return total + nativeRangePageSpan(entry.native); }, 0);
    const ordinals = {};
    const snapshotEntries = nativeFields(document).concat(sectionPageFields(document));
    const snapshots = snapshotEntries.map(function (entry) {
      const identity = entry.ownerNodeId + "\u0000" + entry.fieldKind;
      const ordinal = ordinals[identity] || 0;
      ordinals[identity] = ordinal + 1;
      return {
        stableKey: [entry.ownerNodeId, entry.fieldKind, ordinal],
        fieldCategory: entry.category,
        resultHash: hashVisible(nativeVisibleResult(entry.native)),
        tocPageCount: tocPageCount,
        figureIndexPageCount: figureIndexPageCount,
        tableIndexPageCount: tableIndexPageCount,
        totalPages: totalPages
      };
    });
    if (snapshots.length === 0) snapshots.push({
      stableKey: ["doc:finalize", "PAGE", 0],
      fieldCategory: "page",
      resultHash: hashVisible(String(totalPages) + "-" + String(tocPageCount)),
      tocPageCount: tocPageCount,
      figureIndexPageCount: figureIndexPageCount,
      tableIndexPageCount: tableIndexPageCount,
      totalPages: totalPages
    });
    snapshots.sort(function (left, right) {
      return JSON.stringify(left.stableKey).localeCompare(JSON.stringify(right.stableKey));
    });
    return snapshots;
  }

  function fieldSnapshotSignature(snapshot) {
    return JSON.stringify(snapshot);
  }

  function appendIssueOnce(issues, issue) {
    const duplicateIndex = issues.findIndex(function (existing) {
      return existing.code === issue.code &&
        (existing.placement || "document") === (issue.placement || "document") &&
        (existing.nodeId || null) === (issue.nodeId || null);
    });
    if (duplicateIndex < 0) issues.push(issue);
    else if (issue.recoverable === true && issues[duplicateIndex].recoverable !== true) {
      issues[duplicateIndex] = issue;
    }
  }

  function unstableFieldIssue(snapshot, rounds) {
    const representative = snapshot.length ? snapshot[snapshot.length - 1] : {
      tocPageCount: 0, figureIndexPageCount: 0, tableIndexPageCount: 0, totalPages: 0
    };
    return {
      code: "FIELD_REFRESH_UNSTABLE",
      message: "Field refresh did not converge after " + rounds + " rounds; " +
        "fields=" + snapshot.length + ", toc_pages=" + representative.tocPageCount +
        ", figure_index_pages=" + representative.figureIndexPageCount +
        ", table_index_pages=" + representative.tableIndexPageCount +
        ", total_pages=" + representative.totalPages,
      placement: "document",
      stage: "field-refresh",
      fallback: "document-quality-notice",
      recoverable: true
    };
  }

  function nativeFieldAdapter(document) {
    return {
      repaginateAndUpdateNumbering: function () { repaginateAndUpdateNumbering(document); },
      refreshBookmarksAndReferences: function () { refreshBookmarksAndReferences(document); },
      refreshIndexes: function () { refreshIndexes(document); },
      repaginateAndUpdatePageFields: function () { repaginateAndUpdatePageFields(document); },
      snapshotFields: function () { return snapshotFields(document); },
      upsertDocumentQualityNotice: function (issue) {
        if (!document._wpscQualityAnchor) {
          reserveDocumentQualityAnchor(document, {title: "生成质量提示", notices: []});
        }
        upsertDocumentQualityNotice(document, issue);
      }
    };
  }

  function runNativeFieldConvergence(adapter, rawMaxRounds, issues) {
    if (!Number.isInteger(rawMaxRounds) || rawMaxRounds < 1 || rawMaxRounds > 3) {
      const boundError = nativeError("FIELD_REFRESH_CONTRACT_INVALID");
      throw boundError;
    }
    const required = [
      "repaginateAndUpdateNumbering",
      "refreshBookmarksAndReferences",
      "refreshIndexes",
      "repaginateAndUpdatePageFields",
      "snapshotFields",
      "upsertDocumentQualityNotice"
    ];
    if (!adapter || required.some(function (name) {
      return typeof adapter[name] !== "function";
    })) {
      throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    }
    const fieldSnapshots = [];
    let previousSignature = null;
    try {
      for (let round = 0; round < rawMaxRounds; round += 1) {
        adapter.repaginateAndUpdateNumbering();
        adapter.refreshBookmarksAndReferences();
        adapter.refreshIndexes();
        adapter.repaginateAndUpdatePageFields();
        const snapshot = adapter.snapshotFields();
        fieldSnapshots.push(snapshot);
        const signature = fieldSnapshotSignature(snapshot);
        if (previousSignature !== null && signature === previousSignature) return fieldSnapshots;
        previousSignature = signature;
      }
      const issue = unstableFieldIssue(
        fieldSnapshots[fieldSnapshots.length - 1], rawMaxRounds
      );
      appendIssueOnce(issues, issue);
      adapter.upsertDocumentQualityNotice(issue);
      adapter.repaginateAndUpdateNumbering();
      adapter.refreshBookmarksAndReferences();
      adapter.refreshIndexes();
      adapter.repaginateAndUpdatePageFields();
      const frozen = adapter.snapshotFields();
      fieldSnapshots.push(frozen);
      return fieldSnapshots;
    } catch (error) {
      if (error && error.code === "FIELD_REFRESH_CONTRACT_INVALID") throw error;
      throw nativeError("FIELD_REFRESH_FAILED");
    }
  }

  function runFieldConvergence(document, operations, issues) {
    const finalizers = operations.filter(function (operation) {
      return operation.op === "writer.finalize_fields";
    });
    if (finalizers.length === 0) return [];
    if (finalizers.length !== 1) throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    const operation = finalizers[0];
    const maxRounds = operation.args && operation.args.maxRounds !== undefined
      ? operation.args.maxRounds : 3;
    return runNativeFieldConvergence(nativeFieldAdapter(document), maxRounds, issues);
  }

  function addInlineDegradation(document, args) {
    const code = /^[A-Z][A-Z0-9_]{0,63}$/.test(safeString(args.code))
      ? safeString(args.code) : "DEGRADATION";
    const text = "[" + code + ": " + safePublicText(args.fallbackText) + "]";
    const start = currentPosition(document);
    insertInlineText(document, text);
    const written = typeof document.Range === "function"
      ? document.Range(start, start + text.length) : endRange(document);
    if (written && written.Font) {
      written.Font.Italic = -1;
      written.Font.Color = colorFromHex("#9C0006");
    }
    if (written && written.Shading) {
      written.Shading.BackgroundPatternColor = colorFromHex("#FCE8E6");
    }
    return written;
  }

  function addLiteralInlineDegradation(document, fallbackText) {
    const text = safePublicText(fallbackText);
    const start = currentPosition(document);
    insertInlineText(document, text);
    const written = typeof document.Range === "function"
      ? document.Range(start, start + text.length) : endRange(document);
    if (written && written.Font) {
      written.Font.Italic = -1;
      written.Font.Color = colorFromHex("#9C0006");
    }
    if (written && written.Shading) {
      written.Shading.BackgroundPatternColor = colorFromHex("#FCE8E6");
    }
    return written;
  }

  function insertDegradationBox(document, code, fallbackText, targetRange, rawDisplay) {
    if (!document.Tables || typeof document.Tables.Add !== "function") {
      throw nativeError("DEGRADATION_INSERT_FAILED");
    }
    const safeCode = /^[A-Z][A-Z0-9_]{0,63}$/.test(safeString(code))
      ? safeString(code) : "DEGRADATION";
    const target = targetRange || endRange(document);
    let table;
    try {
      table = document.Tables.Add(target, 1, 1);
      const cell = table.Cell(1, 1);
      cell.Range.Text = rawDisplay === undefined
        ? "[" + safeCode + "] " + safePublicText(fallbackText)
        : safePublicText(rawDisplay);
      cell.Range.Font.Italic = -1;
      cell.Range.Font.Color = colorFromHex("#9C0006");
      cell.Range.Shading.BackgroundPatternColor = colorFromHex("#FCE8E6");
      applyParagraphFormat(cell.Range.ParagraphFormat, {
        spaceBefore: 0, spaceAfter: 3, keepTogether: true, outlineLevel: 10
      });
      if (table.Rows) table.Rows.AllowBreakAcrossPages = false;
      return table;
    } catch (error) {
      throw nativeError("DEGRADATION_INSERT_FAILED");
    }
  }

  function insertStyledDegradationAtRange(document, target, text) {
    if (!target || typeof document.Range !== "function") {
      throw nativeError("DEGRADATION_INSERT_FAILED");
    }
    const start = safeNumber(target.Start, -1);
    if (start < 0) throw nativeError("DEGRADATION_INSERT_FAILED");
    try {
      if (typeof target.InsertAfter === "function") target.InsertAfter(text);
      else target.Text = text;
      const written = document.Range(start, start + text.length);
      if (!written) throw nativeError("DEGRADATION_INSERT_FAILED");
      if (written.Font) {
        written.Font.Italic = -1;
        written.Font.Color = colorFromHex("#9C0006");
      }
      if (written.Shading) {
        written.Shading.BackgroundPatternColor = colorFromHex("#FCE8E6");
      }
      applyParagraphFormat(written.ParagraphFormat, {
        spaceBefore: 0, spaceAfter: 3, keepTogether: true, outlineLevel: 10
      });
      return {Range: written};
    } catch (error) {
      throw nativeError("DEGRADATION_INSERT_FAILED");
    }
  }

  function addDegradationNotice(document, args) {
    const placement = args.placement || "block";
    if (placement === "inline") {
      return addInlineDegradation(document, args);
    }
    if (placement === "document") {
      upsertDocumentQualityNotice(document, args);
      return document._wpscQualityAnchor;
    }
    if (placement !== "block") throw nativeError("DEGRADATION_INSERT_FAILED");
    const target = endRange(document);
    try {
      return insertDegradationBox(document, args.code, args.fallbackText, target);
    } catch (error) {
      // A table failure may remove the one-cell styling API itself. Preserve
      // the minimal visible notice at the same block anchor before declaring
      // insertion fatal.
      try {
        return insertStyledDegradationAtRange(
          document, target,
          "[" + safeString(args.code || "DEGRADATION") + "] " +
            safePublicText(args.fallbackText)
        );
      } catch (minimalError) {
        throw nativeError("DEGRADATION_INSERT_FAILED");
      }
    }
  }

  function reserveDocumentQualityAnchor(document, args) {
    args = args || {};
    if (!document._wpscQualityAnchor) {
      const title = safeString(args.title || "生成质量提示");
      const position = currentPosition(document);
      try {
        if (!document.Bookmarks || typeof document.Bookmarks.Add !== "function" ||
            typeof document.Range !== "function") {
          throw nativeError("DEGRADATION_INSERT_FAILED");
        }
        document.Bookmarks.Add(
          "wpsc_document_quality_anchor", document.Range(position, position)
        );
      } catch (error) {
        throw nativeError("DEGRADATION_INSERT_FAILED");
      }
      document._wpscQualityAnchor = {title: title, position: position, empty: true};
      document._wpscQualityNoticeSeen = Object.create(null);
    }
    (args.notices || []).forEach(function (notice) {
      upsertDocumentQualityNotice(document, notice);
    });
  }

  function upsertDocumentQualityNotice(document, issue) {
    if (!document._wpscQualityAnchor || !document._wpscQualityNoticeSeen) {
      throw nativeError("DEGRADATION_INSERT_FAILED");
    }
    const code = safeString(issue && issue.code) || "QUALITY_NOTICE";
    const placement = safeString(issue && issue.placement) || "document";
    const nodeId = safeString(issue && (issue.nodeId || issue.node_id));
    const identity = code + "\u0000" + placement + "\u0000" + nodeId;
    if (document._wpscQualityNoticeSeen[identity]) return;
    const position = document._wpscQualityAnchor.position;
    const fallbackText = safePublicText(issue && (issue.fallbackText || issue.message));
    const visibleText = document._wpscQualityAnchor.empty
      ? document._wpscQualityAnchor.title + "\r[" + code + "] " + fallbackText
      : fallbackText;
    let table = null;
    let target = null;
    try {
      target = document.Range(position, position);
      table = insertDegradationBox(
        document,
        code,
        fallbackText,
        target,
        document._wpscQualityAnchor.empty ? visibleText : undefined
      );
      document._wpscQualityAnchor.position = safeNumber(
        table && table.Range && table.Range.End, position + 1
      );
    } catch (error) {
      const minimal = document._wpscQualityAnchor.empty
        ? visibleText : "[" + code + "] " + fallbackText;
      try {
        const inserted = insertStyledDegradationAtRange(document, target, minimal);
        document._wpscQualityAnchor.position = safeNumber(
          inserted && inserted.Range && inserted.Range.End,
          position + minimal.length
        );
      } catch (minimalError) {
        throw nativeError("DEGRADATION_INSERT_FAILED");
      }
    }
    document._wpscQualityAnchor.empty = false;
    document._wpscQualityNoticeSeen[identity] = true;
  }

  function addDocumentQualityNotice(document, args) {
    const notices = args.notices || [];
    if (!document._wpscQualityAnchor) {
      reserveDocumentQualityAnchor(document, {notices: notices});
      return;
    }
    notices.forEach(function (notice) { upsertDocumentQualityNotice(document, notice); });
  }

  const OPERATIONS = {
    "writer.reset": resetDocument,
    "writer.configure_page": configurePage,
    "writer.ensure_styles": ensureStyles,
    "writer.add_paragraph": addParagraph,
    "writer.add_heading": addHeadingNative,
    "writer.add_heading_native": addHeadingNative,
    "writer.add_list": addList,
    "writer.add_page_break": function (document) { endRange(document).InsertBreak(7); },
    "writer.configure_section": configureSection,
    "writer.configure_front_matter": function (document, args) {
      document._wpscFrontMatter = args;
      if (args.role) setPageRole(document, args.role);
    },
    "writer.configure_toc_styles": function (document, args) { document._wpscTocDensity = args; },
    "writer.set_page_role": function (document, args) { setPageRole(document, args.role); },
    "writer.set_page_numbering": function (document, args) { setPageNumbering(document, args); },
    "writer.set_header_footer": function (document, args) { setHeaderFooter(document, args); },
    "writer.insert_toc": insertTocWithStyles,
    "writer.insert_toc_with_styles": insertTocWithStyles,
    "writer.insert_figure_index": insertFigureIndex,
    "writer.insert_table_index": insertTableIndex,
    "writer.add_captioned_figure": addCaptionedFigureNative,
    "writer.add_semantic_table": addSemanticTableNative,
    "writer.add_equation": function (document, args, resources, context) {
      if (args.renderMode === "native-m4") {
        return addEquationNativeM4(document, args, resources, context);
      }
      return addEquationNumberNative(document, args, resources, context);
    },
    "writer.add_cross_reference": function (document, args, resources, context) {
      if ((args.runs || []).some(function (run) {
        return run.type === "citation" || run.type === "degradation";
      })) return addCitationParagraph(document, args, resources, context);
      return addCrossReferenceParagraph(document, args, resources, context);
    },
    "writer.add_bibliography": addBibliographyNative,
    // runFieldConvergence is the sole owner; operation dispatch is a no-op.
    "writer.finalize_fields": function () {},
    "writer.add_inline_degradation": addInlineDegradation,
    "writer.add_degradation_notice": addDegradationNotice,
    "writer.add_document_quality_notice": addDocumentQualityNotice,
    "writer.reserve_document_quality_anchor": reserveDocumentQualityAnchor
  };

  function fallbackPlacement(fallbackKind) {
    if (fallbackKind === "inline" || fallbackKind === "inline-fallback") return "inline";
    if (fallbackKind === "document-quality-notice") return "document";
    return "block";
  }

  function operationFallbackText(operation) {
    const args = operation.args || {};
    if (args.fallbackText || args.source || args.text || args.caption) {
      return safePublicText(args.fallbackText || args.source || args.text || args.caption);
    }
    if (Array.isArray(args.runs)) {
      return args.runs.map(function (run) {
        return run.type === "text" ? safePublicText(run.text) : safePublicText(run.fallbackText);
      }).join("");
    }
    if (Array.isArray(args.headers)) {
      return [args.headers].concat(args.rows || []).map(function (row) {
        return row.map(safePublicText).join(" | ");
      }).join("\n");
    }
    if (args.schemaVersion === 1 && Array.isArray(args.entries)) {
      return args.entries.map(function (entry) {
        return "[" + entry.number + "] " + safePublicText(entry.text);
      }).join("\n");
    }
    if (Array.isArray(args.entries)) {
      return args.entries.map(safePublicText).join("\n");
    }
    return "";
  }

  function applyOperationFallback(document, operation, code, fallbackKind, resources, context) {
    const text = operationFallbackText(operation);
    if (fallbackKind === "figure-child-stack-then-notice" &&
        operation.op === "writer.add_captioned_figure") {
      addCaptionedFigureFallback(document, operation.args || {}, resources, context, code);
      return;
    }
    if (fallbackKind === "grid-then-text" &&
        operation.op === "writer.add_semantic_table") {
      addSemanticTableFallback(document, operation.args || {}, resources, context, code);
      return;
    }
    if (fallbackKind === "inline-fallback" &&
        operation.op === "writer.add_cross_reference") {
      addCrossReferenceFallback(document, operation.args || {});
      return;
    }
    if (fallbackKind === "explicit-image-then-source-notice" &&
        operation.op === "writer.add_equation" &&
        operation.args && operation.args.renderMode === "native-m4") {
      addFormulaNativeFallback(
        document, operation.args || {}, resources || {}, context, code
      );
      return;
    }
    if (fallbackKind === "inline" || fallbackKind === "inline-fallback") {
      addInlineDegradation(document, {code: code, fallbackText: text});
      return;
    }
    if (fallbackKind === "notice" ||
        fallbackKind === "explicit-image-then-source-notice") {
      addDegradationNotice(document, {
        code: code, fallbackText: text, placement: "block"
      });
      return;
    }
    throw nativeError("DEGRADATION_FALLBACK_FAILED");
  }

  function recoverOperation(
    document, operation, resources, issues, context, checkpoint, code, fallbackKind
  ) {
    const placement = fallbackPlacement(fallbackKind);
    const controller = document._wpscRecoveryController || createLocalRecoveryController();
    document._wpscRecoveryController = controller;
    recoverLocalFailure(controller, {
      nodeId: operation.nodeId || operation.op,
      descriptor: {
        code: code, placement: placement, fallbackKind: fallbackKind,
        fallbackText: operationFallbackText(operation)
      },
      rollback: function (token) { rollbackMutation(document, token); },
      fallbackAttempt: function () {
        applyOperationFallback(
          document, operation, code, fallbackKind, resources || {}, context
        );
      },
      // The local inline fallback or restrained block is itself the visible
      // same-node notice. Reaching this callback proves insertion completed.
      insertNotice: function () {}
    }, checkpoint, code);
    appendIssueOnce(issues, {
      code: code,
      message: operation.op + " used its declared fallback",
      placement: placement,
      nodeId: operation.nodeId,
      stage: "native",
      fallback: fallbackKind,
      recoverable: true
    });
  }

  function runOperation(document, operation, resources, issues, childResults) {
    const opName = operation.op;
    const issueCheckpoint = issues.length;
    const childCheckpoint = childResults.length;
    const context = {
      ownerNodeId: operation.nodeId || null,
      issues: issues,
      childResults: childResults,
      controllerOwned: true
    };
    const deferred = LONGFORM_DEFERRED[opName];
    const handler = OPERATIONS[opName];
    if (!deferred && !handler) {
      throw nativeError("UNKNOWN_OPERATION");
    }
    const policy = operation.failurePolicy || {};
    let recoverable = false;
    if (deferred) {
      recoveryDecision(deferred[0], deferred[1], fallbackPlacement(deferred[1]));
      recoverable = true;
    } else if (policy.mode === "degrade" &&
               Array.isArray(policy.recoverableCodes) &&
               policy.recoverableCodes.length > 0) {
      policy.recoverableCodes.forEach(function (code) {
        recoveryDecision(code, policy.fallback, fallbackPlacement(policy.fallback));
      });
      recoverable = true;
    }
    const checkpoint = recoverable ? checkpointRecoverableMutation(document) : null;
    if (deferred) {
      recoverOperation(
        document, operation, resources, issues, context,
        checkpoint, deferred[0], deferred[1]
      );
      return;
    }
    try {
      handler(document, operation.args || {}, resources || {}, context);
    } catch (error) {
      if (error && error.code === "LOCAL_MUTATION_ROLLBACK_FAILED") {
        throw nativeError("LOCAL_MUTATION_ROLLBACK_FAILED");
      }
      if (policy.mode === "fail") {
        throw nativeError("EXECUTION_ABORTED");
      }
      if (policy.mode === "degrade") {
        const code = error && error.code;
        const allowed = Array.isArray(policy.recoverableCodes) &&
          policy.recoverableCodes.indexOf(code) !== -1;
        if (allowed) {
          issues.length = issueCheckpoint;
          childResults.length = childCheckpoint;
          recoverOperation(
            document, operation, resources, issues, context,
            checkpoint, code, policy.fallback
          );
          return;
        }
      }
      throw nativeError(error && error.code ? error.code : "EXECUTION_ABORTED");
    }
  }

  function buildPaginationMap(operations) {
    const nodes = [];
    const seen = {};
    operations.forEach(function (operation) {
      if (operation.nodeId && !seen[operation.nodeId]) {
        seen[operation.nodeId] = true;
        nodes.push({
          nodeId: operation.nodeId,
          fragments: [{ page: 1 }]
        });
      }
    });
    return { version: "M2-stub", nodes: nodes };
  }

  function validatePrivateResourceMap(plan, resources) {
    if (!resources || typeof resources !== "object" || Array.isArray(resources)) {
      throw nativeError("RESOURCE_HASH_MISMATCH");
    }
    const expected = {};
    (plan.operations || []).forEach(function (operation) {
      const args = operation.args || {};
      if (operation.op === "writer.add_captioned_figure") {
        (args.children || []).forEach(function (child) {
          if (child.resourceId) expected[child.resourceId] = true;
        });
      }
      if (operation.op === "writer.add_equation" && args.renderMode === "native-m4") {
        const fallback = args.fallbackResource || {};
        if (fallback.fallbackResourceId) expected[fallback.fallbackResourceId] = true;
      }
    });
    Object.keys(expected).forEach(function (resourceId) {
      if (typeof resources[resourceId] !== "string" || resources[resourceId].length === 0) {
        throw nativeError("RESOURCE_HASH_MISMATCH");
      }
    });
    Object.keys(resources).forEach(function (resourceId) {
      if (!expected[resourceId]) throw nativeError("RESOURCE_HASH_MISMATCH");
    });
  }

  function validateLongformRequest(params) {
    if (!params || typeof params !== "object" || Array.isArray(params) ||
        !params.plan || typeof params.plan !== "object" ||
        params.plan.component !== "writer" ||
        typeof params.outputPath !== "string" || !params.outputPath) {
      throw nativeError("PROTOCOL_MISMATCH");
    }
    const operations = Array.isArray(params.plan.operations) ? params.plan.operations : [];
    const isM4 = operations.some(function (operation) {
      const args = operation && operation.args || {};
      return args.renderMode === "native-m4" || args.schemaVersion === 1 ||
        hasOwn(args, "cellCitations") || hasOwn(args, "cellDegradations") ||
        (args.runs || []).some(function (run) {
          return run.type === "citation" || run.type === "degradation";
        });
    });
    const keys = Object.keys(params).sort().join(",");
    if ((isM4 && keys !== "outputPath,plan,resources") ||
        (!isM4 && keys !== "outputPath,plan" && keys !== "outputPath,plan,resources") ||
        (params.plan.protocolVersion !== undefined && params.plan.protocolVersion !== 2) ||
        (isM4 && params.plan.protocolVersion !== 2)) {
      throw nativeError("PROTOCOL_MISMATCH");
    }
    if (params.resources === undefined) params.resources = {};
    if (!params.resources || typeof params.resources !== "object" ||
        Array.isArray(params.resources)) throw nativeError("PROTOCOL_MISMATCH");
    Object.keys(params.resources).forEach(function (resourceId) {
      if (!resourceId || typeof params.resources[resourceId] !== "string" ||
          !params.resources[resourceId]) throw nativeError("PROTOCOL_MISMATCH");
    });
    return params;
  }

  function requiredBookmark(document, name) {
    if (typeof name !== "string" || !/^wpsc_(fig|tab|eq)_[a-z0-9]{24}$/.test(name)) {
      throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    }
    if (!document.Bookmarks || typeof document.Bookmarks.Exists !== "function" || !document.Bookmarks.Exists(name)) {
      throw nativeError("FIELD_REFRESH_FAILED");
    }
    return collectionItem(document.Bookmarks, name).Range;
  }

  function paragraphRangeFor(range) {
    const paragraphs = range && range.Paragraphs;
    const paragraph = paragraphs ? collectionItem(paragraphs, 1) : null;
    if (!paragraph || !paragraph.Range) throw nativeError("EXECUTION_ABORTED");
    return paragraph.Range;
  }

  function nativeObjectRange(document, mutation) {
    const anchor = requiredBookmark(document, mutation.bookmarkName);
    const caption = paragraphRangeFor(anchor);
    let start = safeNumber(caption.Start, 0);
    let finish = safeNumber(caption.End, start);
    if (mutation.kind === "figure") {
      let located = false;
      const shapes = document.InlineShapes;
      for (let index = 1; index <= requiredCollectionCount(shapes); index += 1) {
        const shape = collectionItem(shapes, index);
        if (shape && shape.AlternativeText === mutation.ownerNodeId && shape.Range) {
          start = Math.min(start, safeNumber(shape.Range.Start, start));
          located = true;
        }
      }
      if (!located && start > 0) {
        const previous = document.Range(Math.max(0, start - 2), start);
        start = safeNumber(paragraphRangeFor(previous).Start, start);
      }
      const tables = document.Tables;
      for (let tableIndex = 1; tableIndex <= requiredCollectionCount(tables); tableIndex += 1) {
        const table = collectionItem(tables, tableIndex);
        if (table && table.Range && safeNumber(table.Range.End, 0) <= safeNumber(caption.Start, 0)) {
          const distance = safeNumber(caption.Start, 0) - safeNumber(table.Range.End, 0);
          if (distance <= 3) start = Math.min(start, safeNumber(table.Range.Start, start));
        }
      }
    } else if (mutation.kind === "table") {
      const tables = document.Tables;
      let nearest = null;
      for (let index = 1; index <= requiredCollectionCount(tables); index += 1) {
        const table = collectionItem(tables, index);
        if (!table || !table.Range) continue;
        const tableStart = safeNumber(table.Range.Start, -1);
        if (tableStart >= finish && (nearest === null || tableStart < safeNumber(nearest.Range.Start, tableStart + 1))) {
          nearest = table;
        }
      }
      if (!nearest) throw nativeError("EXECUTION_ABORTED");
      finish = safeNumber(nearest.Range.End, finish);
    } else if (mutation.kind !== "equation") {
      throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    }
    return document.Range(start, finish);
  }

  function validateMutation(mutation) {
    if (!mutation || typeof mutation !== "object" || Array.isArray(mutation)) {
      throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    }
    const allowed = {
      move: ["type", "kind", "bookmarkName", "ownerNodeId", "beforeBookmarkName", "beforeKind", "beforeOwnerNodeId"],
      insert: ["type", "operations"],
      delete: ["type", "kind", "bookmarkName", "ownerNodeId"]
    };
    if (!allowed[mutation.type] || Object.keys(mutation).some(function (key) { return allowed[mutation.type].indexOf(key) === -1; })) {
      throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    }
    if (mutation.type === "insert" && (!Array.isArray(mutation.operations) || mutation.operations.length === 0)) {
      throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    }
  }

  function applyLongformMutation(document, mutation, resources, issues, childResults) {
    validateMutation(mutation);
    if (mutation.type === "move") {
      const source = nativeObjectRange(document, mutation);
      const destination = nativeObjectRange(document, {
        kind: mutation.beforeKind,
        bookmarkName: mutation.beforeBookmarkName,
        ownerNodeId: mutation.beforeOwnerNodeId
      }).Duplicate;
      source.Cut();
      // Pasting at a table's Start on WPS nests the moved table in its first
      // cell. For a table destination, the preceding paragraph mark is the
      // external insertion anchor. Non-table destinations can safely create
      // a fresh paragraph before their range.
      let destinationIsTable = false;
      for (let index = 1; index <= requiredCollectionCount(document.Tables); index += 1) {
        const table = collectionItem(document.Tables, index);
        if (table && table.Range && safeNumber(table.Range.Start, -1) === safeNumber(destination.Start, -2)) {
          destinationIsTable = true;
          break;
        }
      }
      let pasteRange;
      if (destinationIsTable) {
        const anchor = Math.max(0, safeNumber(destination.Start, 0) - 1);
        pasteRange = document.Range(anchor, anchor);
      } else {
        if (typeof destination.InsertParagraphBefore !== "function") {
          throw nativeError("EXECUTION_ABORTED");
        }
        destination.InsertParagraphBefore();
        pasteRange = document.Range(destination.Start, destination.Start);
      }
      pasteRange.Paste();
      return;
    }
    if (mutation.type === "delete") {
      nativeObjectRange(document, mutation).Delete();
      return;
    }
    mutation.operations.forEach(function (operation) {
      if (!operation || typeof operation.op !== "string" || operation.op === "writer.finalize_fields") {
        throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
      }
      runOperation(document, operation, resources, issues, childResults);
    });
  }

  function mutate(params) {
    const sourcePath = params.sourcePath;
    const outputPath = params.outputPath;
    const mutations = params.mutations;
    const resources = params.resources || {};
    if (typeof sourcePath !== "string" || !sourcePath || typeof outputPath !== "string" || !outputPath || !Array.isArray(mutations) || mutations.length === 0) {
      throw nativeError("FIELD_REFRESH_CONTRACT_INVALID");
    }
    const previousAlerts = Application.DisplayAlerts;
    const previousScreenUpdating = Application.ScreenUpdating;
    let document = null;
    const issues = [];
    const childResults = [];
    const refreshRounds = [];
    try {
      Application.DisplayAlerts = 0;
      Application.ScreenUpdating = false;
      document = Application.Documents.Open(sourcePath, false, false);
      mutations.forEach(function (mutation) {
        applyLongformMutation(document, mutation, resources, issues, childResults);
        refreshRounds.push(runNativeFieldConvergence(nativeFieldAdapter(document), 3, issues).length);
      });
      document.SaveAs2(outputPath, 12);
      document.Close(0);
      document = Application.Documents.Open(outputPath, false, false);
      refreshRounds.push(runNativeFieldConvergence(nativeFieldAdapter(document), 3, issues).length);
      document.Save();
      document.Close(0);
      document = null;
      return {
        outputPath: outputPath,
        mutationCount: mutations.length,
        mutationKinds: mutations.map(function (mutation) { return mutation.type; }).filter(function (kind, index, values) {
          return values.indexOf(kind) === index;
        }),
        refreshRounds: refreshRounds,
        issueCodes: issues.map(function (issue) { return issue.code; })
      };
    } catch (error) {
      if (document !== null) {
        try { document.Close(0); } catch (closeError) { /* ignore */ }
      }
      if (!error.code) error.code = "GENERATION_COMMAND_FAILED";
      throw error;
    } finally {
      try { Application.ScreenUpdating = previousScreenUpdating; }
      finally { Application.DisplayAlerts = previousAlerts; }
    }
  }

  function run(params) {
    params = validateLongformRequest(params);
    const plan = params.plan;
    const outputPath = params.outputPath;
    const resources = params.resources || {};
    const previousAlerts = Application.DisplayAlerts;
    const previousScreenUpdating = Application.ScreenUpdating;
    let document = null;
    let appliedCount = 0;
    const issues = [];
    const childResults = [];

    try {
      validatePrivateResourceMap(plan, resources);
      Application.DisplayAlerts = 0;
      Application.ScreenUpdating = false;
      document = Application.Documents.Add();
      document._wpscFirstSectionConfigured = false;

      const operations = plan.operations || [];
      operations.forEach(function (operation) {
        runOperation(document, operation, resources, issues, childResults);
        appliedCount += 1;
      });

      const fieldSnapshots = runFieldConvergence(document, operations, issues);

      try {
        document.SaveAs2(outputPath, 12);
      } catch (error) {
        throw nativeError("SAVE_FAILED");
      }
      document.Close(0);
      document = null;
      return {
        outputPath: outputPath,
        appliedOperations: appliedCount,
        issueCodes: issues,
        paginationMap: buildPaginationMap(operations),
        fieldSnapshots: fieldSnapshots,
        childResults: childResults
      };
    } catch (error) {
      if (document !== null) {
        try {
          document.Close(0);
        } catch (closeError) {
          // ignore
        }
        document = null;
      }
      if (!error.code) {
        error.code = "GENERATION_COMMAND_FAILED";
      }
      throw error;
    } finally {
      try {
        Application.ScreenUpdating = previousScreenUpdating;
      } finally {
        Application.DisplayAlerts = previousAlerts;
      }
    }
  }

  window.WPSComposerLongformV2 = Object.freeze({
    run: run,
    mutate: mutate,
    OPERATIONS: Object.keys(OPERATIONS),
    __test: Object.freeze({
      addNativeNumberShell: addNativeNumberShell,
      createFigureColumns: createFigureColumns,
      addCaptionedFigureNative: addCaptionedFigureNative,
      addSemanticTableNative: addSemanticTableNative,
      addEquationNumberNative: addEquationNumberNative,
      addEquationNativeM4: addEquationNativeM4,
      addFormulaNativeFallback: addFormulaNativeFallback,
      addCitationParagraph: addCitationParagraph,
      addBibliographyNative: addBibliographyNative,
      applyTableCellMetadata: applyTableCellMetadata,
      validateLongformRequest: validateLongformRequest,
      addCrossReferenceParagraph: addCrossReferenceParagraph,
      insertCaptionIndexNative: insertCaptionIndexNative,
      createNativeTable: createNativeTable,
      nativeFieldAdapter: nativeFieldAdapter,
      runNativeFieldConvergence: runNativeFieldConvergence,
      buildPaginationMap: buildPaginationMap,
      hashVisible: hashVisible,
      rollbackMutation: rollbackMutation,
      recoveryDecision: recoveryDecision,
      createLocalRecoveryController: createLocalRecoveryController,
      runLocalRecovery: runLocalRecovery,
      appendIssueOnce: appendIssueOnce,
      safePublicText: safePublicText,
      addInlineDegradation: addInlineDegradation,
      addDegradationNotice: addDegradationNotice,
      reserveDocumentQualityAnchor: reserveDocumentQualityAnchor,
      upsertDocumentQualityNotice: upsertDocumentQualityNotice,
      applyLongformMutation: applyLongformMutation,
      runOperation: runOperation
    })
  });
}());
