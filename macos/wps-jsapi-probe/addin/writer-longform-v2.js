(function () {
  "use strict";

  const LONGFORM_DEFERRED = {
    "writer.add_captioned_figure": ["IMAGE_INSERT_FAILED", "notice"],
    "writer.add_semantic_table": ["TABLE_INSERT_FAILED", "notice"],
    "writer.add_equation": ["EQUATION_INSERT_FAILED", "inline"],
    "writer.add_bibliography": ["BIBLIOGRAPHY_INSERT_FAILED", "notice"],
    "writer.add_cross_reference": ["CROSS_REFERENCE_FAILED", "inline"]
  };

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
    insertText(document, args.text, styleName, args);
  }

  function addHeadingNative(document, args) {
    addHeading(document, args);
    if (!args.numbering) {
      return;
    }
    try {
      const style = getStyle(document, "Heading " + args.level);
      if (!style) {
        return;
      }
      const template = document.ListTemplates.Add(true);
      const schemeMap = { "chinese-formal": "%1", decimal: "%1.", "hybrid-bid": "%1." };
      const format = schemeMap[args.numberingScheme] || "%1.";
      const levelIndex = safeNumber(args.level, 1);
      template.ListLevels(levelIndex).NumberFormat = format;
      style.LinkToListTemplate(template, levelIndex);
    } catch (error) {
      // Native numbering is best-effort.
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
    setPageNumbering(document, {
      format: args.pageNumberFormat || "continue",
      start: args.startPageNumber,
      restart: args.restartPageNumbering
    });
    setHeaderFooter(document, {
      headerText: args.headerText,
      footerText: args.footerText,
      linkToPreviousHeader: args.linkToPreviousHeader,
      linkToPreviousFooter: args.linkToPreviousFooter
    });
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
          // ignore
        }
      }
    } catch (error) {
      // ignore
    }
  }

  function setHeaderFooter(document, args) {
    try {
      const section = collectionItem(document.Sections, document.Sections.Count);
      if (args.linkToPreviousHeader !== undefined) {
        try {
          section.Headers(1).LinkToPrevious = args.linkToPreviousHeader ? -1 : 0;
        } catch (error) {
          // ignore
        }
      }
      if (args.linkToPreviousFooter !== undefined) {
        try {
          section.Footers(1).LinkToPrevious = args.linkToPreviousFooter ? -1 : 0;
        } catch (error) {
          // ignore
        }
      }
      if (args.headerText !== undefined && args.headerText !== null) {
        const header = section.Headers(1);
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
          section.Footers(1).Range.Text = String(args.footerText);
        }
      }
    } catch (error) {
      // ignore
    }
  }

  function insertTocWithStyles(document, args) {
    const density = args.density || document._wpscTocDensity || {};
    try {
      document.TablesOfContents.Add(endRange(document), true, 1, args.levels || 3);
    } catch (error) {
      // ignore
    }
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

  function insertFigureIndex(document, args) {
    try {
      document.TablesOfFigures.Add(endRange(document), "Figure");
    } catch (error) {
      // ignore
    }
  }

  function insertTableIndex(document, args) {
    try {
      document.TablesOfFigures.Add(endRange(document), "Table");
    } catch (error) {
      // ignore
    }
  }

  function finalizeFields(document, args) {
    try {
      for (let i = 1; i <= document.TablesOfContents.Count; i += 1) {
        document.TablesOfContents.Item(i).Update();
      }
    } catch (error) {
      // ignore
    }
    try {
      for (let i = 1; i <= document.TablesOfFigures.Count; i += 1) {
        document.TablesOfFigures.Item(i).Update();
      }
    } catch (error) {
      // ignore
    }
    try {
      document.Fields.Update();
    } catch (error) {
      // ignore
    }
  }

  function buildFieldSnapshot(document, roundIndex) {
    let totalPages = 1;
    let tocPageCount = 0;
    try {
      totalPages = Number(document.ComputeStatistics(2)) || 1;
    } catch (error) {
      totalPages = 1;
    }
    try {
      tocPageCount = Number(document.TablesOfContents.Count) || 0;
    } catch (error) {
      tocPageCount = 0;
    }
    return {
      stableKey: ["doc:finalize", "PAGE", 0],
      fieldCategory: "page",
      resultHash: String(totalPages) + "-" + String(tocPageCount),
      tocPageCount: tocPageCount,
      figureIndexPageCount: 0,
      tableIndexPageCount: 0,
      totalPages: totalPages
    };
  }

  function runFieldConvergence(document, operations) {
    let fieldSnapshots = [];
    operations.forEach(function (operation) {
      if (operation.op === "writer.finalize_fields") {
        const maxRounds = operation.args && Number(operation.args.maxRounds) || 3;
        for (let round = 0; round <= maxRounds; round += 1) {
          finalizeFields(document, operation.args);
          fieldSnapshots.push(buildFieldSnapshot(document, round));
        }
      }
    });
    return fieldSnapshots;
  }

  function addInlineDegradation(document, args) {




    insertText(document, safeString(args.fallbackText), null, {});
  }

  function addDegradationNotice(document, args) {
    const placement = args.placement || "block";
    const text = "[" + args.code + "] " + safeString(args.fallbackText);
    if (placement === "inline") {
      insertText(document, text, null, {});
    } else {
      insertText(document, text, null, { italic: true });
    }
  }

  function addDocumentQualityNotice(document, args) {
    const notices = args.notices || [];
    notices.forEach(function (notice) {
      addDegradationNotice(document, notice);
    });
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
    "writer.configure_front_matter": function (document, args) { setPageRole(document, args.role || "front_matter"); },
    "writer.configure_toc_styles": function (document, args) { document._wpscTocDensity = args; },
    "writer.set_page_role": function (document, args) { setPageRole(document, args.role); },
    "writer.set_page_numbering": function (document, args) { setPageNumbering(document, args); },
    "writer.set_header_footer": function (document, args) { setHeaderFooter(document, args); },
    "writer.insert_toc": insertTocWithStyles,
    "writer.insert_toc_with_styles": insertTocWithStyles,
    "writer.insert_figure_index": insertFigureIndex,
    "writer.insert_table_index": insertTableIndex,
    "writer.finalize_fields": finalizeFields,
    "writer.add_inline_degradation": addInlineDegradation,
    "writer.add_degradation_notice": addDegradationNotice,
    "writer.add_document_quality_notice": addDocumentQualityNotice
  };

  function runOperation(document, operation, resources, issues) {
    const opName = operation.op;
    const deferred = LONGFORM_DEFERRED[opName];
    if (deferred) {
      issues.push({
        code: deferred[0],
        message: opName + " is deferred to fallback in M2",
        placement: "document",
        nodeId: operation.nodeId
      });
      if (deferred[1] === "inline") {
        addInlineDegradation(document, { fallbackText: operation.args.fallbackText || operation.args.source || operation.args.text || "" });
      } else {
        addDegradationNotice(document, {
          code: deferred[0],
          message: opName + " is deferred",
          fallbackText: operation.args.fallbackText || operation.args.source || operation.args.text || "",
          placement: operation.args.placement || "block"
        });
      }
      return;
    }
    const handler = OPERATIONS[opName];
    if (!handler) {
      issues.push({
        code: "UNKNOWN_OPERATION",
        message: "No handler for " + opName,
        placement: "document",
        nodeId: operation.nodeId
      });
      return;
    }
    try {
      handler(document, operation.args);
    } catch (error) {
      const policy = operation.failurePolicy || {};
      if (policy.mode === "fail") {
        error.code = "EXECUTION_ABORTED";
        throw error;
      }
      if (policy.mode === "degrade") {
        issues.push({
          code: policy.recoverableCodes && policy.recoverableCodes[0] || "EXECUTION_FAILED",
          message: opName + " degraded: " + error,
          placement: "document",
          nodeId: operation.nodeId
        });
        try {
          addDegradationNotice(document, {
            code: "EXECUTION_FAILED",
            message: opName + " degraded",
            fallbackText: operation.args.fallbackText || String(error),
            placement: operation.args.placement || "block"
          });
        } catch (fallbackError) {
          // ignore
        }
      } else {
        issues.push({
          code: "EXECUTION_FAILED",
          message: opName + " failed: " + error,
          placement: "document",
          nodeId: operation.nodeId
        });
      }
    }
  }

  function buildPaginationMap(operations) {
    const nodes = [];
    operations.forEach(function (operation) {
      if (operation.nodeId) {
        nodes.push({
          nodeId: operation.nodeId,
          fragments: [{ page: 1 }]
        });
      }
    });
    return { version: "M2-stub", nodes: nodes };
  }

  function run(params) {
    const plan = params.plan;
    const outputPath = params.outputPath;
    const resources = params.resources || {};
    const previousAlerts = Application.DisplayAlerts;
    const previousScreenUpdating = Application.ScreenUpdating;
    let document = null;
    let appliedCount = 0;
    const issues = [];

    try {
      Application.DisplayAlerts = 0;
      Application.ScreenUpdating = false;
      document = Application.Documents.Add();
      document._wpscFirstSectionConfigured = false;

      const operations = plan.operations || [];
      operations.forEach(function (operation) {
        runOperation(document, operation, resources, issues);
        appliedCount += 1;
      });

      const fieldSnapshots = runFieldConvergence(document, operations);

      document.SaveAs2(outputPath, 12);
      document.Close(0);
      document = null;
      return {
        outputPath: outputPath,
        appliedOperations: appliedCount,
        issueCodes: issues,
        paginationMap: buildPaginationMap(operations),
        fieldSnapshots: fieldSnapshots
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

  window.WPSComposerLongformV2 = Object.freeze({ run, OPERATIONS: Object.keys(OPERATIONS) });
}());
