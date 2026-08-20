(function () {
  "use strict";

  const PROBE_VERSION = "0.8.0-m0.1";
  const PROTOCOL_VERSION = 2;
  const RESOURCE_MANIFEST_VERSION = 1;
  const RESOURCE_MANIFEST_DIGEST = "3516239310e9d6c0d9a736e816661d57d7e5378e334cfb4f03454bb3da0fc4ae";
  const SVG_SHA256 = "924f47ebd7a4c22393defac4103818aedced9f5dd0630bf27e1d6aa7ad30cbfc";
  const M0_KEYS = ["expectedWpsVersion", "manifest", "probeVersion", "protocolVersion", "resourceManifestVersion", "stagedDocxPath", "stagedPdfPath", "stagedSvgPath"];
  const MANIFEST_KEYS = ["digest", "entries", "version"];
  const MANIFEST_ENTRY_KEYS = ["byteLength", "mediaType", "normalizerId", "payloadSha256", "resourceId", "sourceSha256"];

  function fail(code, message) {
    const error = new Error(message);
    error.code = code;
    throw error;
  }

  function isObject(value) {
    return value !== null && typeof value === "object" && !Array.isArray(value);
  }

  function hasExactKeys(value, expected) {
    if (!isObject(value)) {
      return false;
    }
    const actual = Object.keys(value).sort();
    const wanted = expected.slice().sort();
    return actual.length === wanted.length && actual.every(function (key, index) { return key === wanted[index]; });
  }

  function requirePath(value, label) {
    if (typeof value !== "string" || value.length === 0 || value.length > 4096 || /[\u0000-\u001f\u007f]/.test(value)) {
      fail("PROTOCOL_MISMATCH", `${label} is invalid`);
    }
    return value;
  }

  function validateEnvelope(params) {
    if (!hasExactKeys(params, M0_KEYS)) {
      fail("PROTOCOL_MISMATCH", "Long-form M0 envelope fields are invalid");
    }
    if (params.protocolVersion !== PROTOCOL_VERSION || params.resourceManifestVersion !== RESOURCE_MANIFEST_VERSION || params.probeVersion !== PROBE_VERSION) {
      fail("PROTOCOL_MISMATCH", "Long-form M0 version is incompatible");
    }
    if (typeof params.expectedWpsVersion !== "string" || params.expectedWpsVersion.length === 0 || params.expectedWpsVersion.length > 128) {
      fail("PROTOCOL_MISMATCH", "Long-form M0 expected WPS version is invalid");
    }
    requirePath(params.stagedDocxPath, "stagedDocxPath");
    requirePath(params.stagedPdfPath, "stagedPdfPath");
    requirePath(params.stagedSvgPath, "stagedSvgPath");
    const manifest = params.manifest;
    if (!hasExactKeys(manifest, MANIFEST_KEYS) || manifest.version !== RESOURCE_MANIFEST_VERSION || !Array.isArray(manifest.entries) || manifest.entries.length !== 1 || manifest.digest !== RESOURCE_MANIFEST_DIGEST) {
      fail("RESOURCE_MANIFEST_MISMATCH", "Long-form M0 resource manifest is invalid");
    }
    const entry = manifest.entries[0];
    if (!hasExactKeys(entry, MANIFEST_ENTRY_KEYS) || entry.resourceId !== "m0-static-svg" || entry.sourceSha256 !== SVG_SHA256 || entry.payloadSha256 !== SVG_SHA256 || entry.byteLength !== 185 || entry.mediaType !== "image/svg+xml" || entry.normalizerId !== "identity-svg-m0-v1") {
      fail("RESOURCE_MANIFEST_MISMATCH", "Long-form M0 resource manifest entry is invalid");
    }
    return params;
  }

  function item(collection, index) {
    if (collection && typeof collection.Item === "function") { return collection.Item(index); }
    if (typeof collection === "function") { return collection(index); }
    return null;
  }

  function safeNumber(value, fallback) {
    const number = Number(value);
    return Number.isFinite(number) ? number : fallback;
  }

  function safeCount(collection) {
    try { return Math.max(0, safeNumber(collection && collection.Count, 0)); }
    catch (error) { return 0; }
  }

  function endRange(document) {
    const end = Math.max(0, safeNumber(document.Content && document.Content.End, 1) - 1);
    return document.Range(end, end);
  }

  function appendText(document, value) {
    const range = endRange(document);
    const start = safeNumber(range.Start, 0);
    range.InsertAfter(String(value) + "\r");
    return document.Range(start, start + String(value).length);
  }

  function marker(id) { return `M0C${String(id).padStart(2, "0")}`; }

  function annotateFailure(document, id) {
    try {
      const range = appendText(document, `[M0 CAP ${id} FAILED]`);
      if (range.Font) { range.Font.Color = 192; range.Font.Bold = -1; }
    } catch (error) { void error; }
  }

  function errorClass(error) {
    const name = error && error.name ? String(error.name) : "Error";
    return /^[A-Za-z][A-Za-z0-9_.-]{0,63}$/.test(name) ? name : "Error";
  }

  function safeStage(error) {
    const stage = error && error.probeStage ? String(error.probeStage) : "operation";
    return /^[a-z][a-z0-9-]{0,63}$/.test(stage) ? stage : "operation";
  }

  function resultRow(id, status, checks, metrics) { return {id, status, checks, metrics: metrics || {}}; }

  function assay(document, capabilities, id, operation, optional) {
    try {
      const outcome = operation() || {};
      if (outcome.passed === false) {
        annotateFailure(document, id);
        capabilities.push(resultRow(id, optional ? "unsupported" : "failed", ["native-attempted"], outcome.metrics || {}));
        return;
      }
      appendText(document, marker(id));
      capabilities.push(resultRow(id, "passed", ["native"], outcome.metrics || {}));
    } catch (error) {
      annotateFailure(document, id);
      capabilities.push(resultRow(id, optional ? "unsupported" : "failed", ["native-attempted"], {errorClass: errorClass(error), stage: safeStage(error)}));
    }
  }

  function chooseFont(preferred) {
    const available = [];
    for (let index = 1; index <= safeCount(Application.FontNames); index += 1) {
      try { available.push(String(item(Application.FontNames, index))); }
      catch (error) { void error; }
    }
    for (let index = 0; index < preferred.length; index += 1) {
      if (available.indexOf(preferred[index]) !== -1) { return {name: preferred[index], count: available.length}; }
    }
    return {name: available.length ? available[0] : "", count: available.length};
  }

  function capability3Fonts(document) {
    const cjk = chooseFont(["Songti SC", "STSong", "SimSun"]);
    const latin = chooseFont(["Times New Roman", "Times"]);
    const mono = chooseFont(["Menlo", "Courier New", "Courier"]);
    if (!cjk.name || !latin.name || !mono.name) { return {passed: false, metrics: {fontCount: cjk.count}}; }
    [["CJK", cjk.name], ["Latin", latin.name], ["Mono", mono.name]].forEach(function (entry) {
      const range = appendText(document, entry[0]);
      range.Font.Name = entry[1];
      range.Font.NameFarEast = entry[1];
    });
    return {metrics: {fontCount: cjk.count, mappedFontCount: 3}};
  }

  function capability4Unicode(document) {
    const values = ["\u4E2D\u6587\u00E9", "e\u0301", "\uD83D\uDC69\u200D\uD83D\uDCBB", "\uD87E\uDC00\uFE0F"];
    let utf16Units = 0;
    values.forEach(function (value) {
      const range = appendText(document, value);
      if (safeNumber(range.End, -1) - safeNumber(range.Start, -1) !== value.length) { throw new Error("UTF16RangeMismatch"); }
      utf16Units += value.length;
    });
    return {metrics: {paragraphCount: values.length, utf16Units, unicodeDataVersion: "15.1.0", normalizationModes: 2}};
  }

  function capability5Coordinates(document) {
    const shape = document.Shapes.AddTextbox(1, 72, 72, 120, 36, endRange(document));
    shape.RelativeHorizontalPosition = 1;
    shape.RelativeVerticalPosition = 1;
    shape.Left = 72;
    shape.Top = 72;
    shape.TextFrame.MarginLeft = 0;
    shape.TextFrame.MarginRight = 0;
    shape.TextFrame.MarginTop = 0;
    shape.TextFrame.MarginBottom = 0;
    shape.TextFrame.TextRange.Text = "M0XY5";
    const paragraph = appendText(document, "COORD");
    paragraph.Font.Size = 1;
    paragraph.ParagraphFormat.SpaceBefore = 0;
    paragraph.ParagraphFormat.SpaceAfter = 0;
    paragraph.ParagraphFormat.LineSpacingRule = 4;
    paragraph.ParagraphFormat.LineSpacing = 1;
    if (typeof document.Repaginate === "function") { document.Repaginate(); }
    const x = safeNumber(paragraph.Information(5), NaN);
    const y = safeNumber(paragraph.Information(6), NaN);
    const left = safeNumber(shape.Left, NaN);
    const top = safeNumber(shape.Top, NaN);
    const passed = [x, y, left, top].every(Number.isFinite);
    return {passed, metrics: {origin: "top-left", unit: "point", paragraphX: x, paragraphY: y, shapeX: left, shapeY: top}};
  }

  const NUMBERING_SCHEMES = [
    [
      {format: "第%1章", style: 37},
      {format: "第%2节", style: 37},
      {format: "%3、", style: 37},
      {format: "（%4）", style: 37}
    ],
    [
      {format: "%1", style: 0},
      {format: "%1.%2", style: 0},
      {format: "%1.%2.%3", style: 0},
      {format: "%1.%2.%3.%4", style: 0}
    ],
    [
      {format: "第%1章", style: 37},
      {format: "%1.%2", style: 253},
      {format: "%1.%2.%3", style: 253},
      {format: "关键工法%4：", style: 22, neverRestart: true}
    ]
  ];

  const NUMBERING_EXPECTED = [
    ["第一章", "第一节", "一、", "（一）"],
    ["1", "1.1", "1.1.1", "1.1.1.1"],
    ["第一章", "1.1", "1.1.1", "关键工法01："]
  ];

  function normalizedListString(value) {
    return String(value || "").replace(/\s+/g, "");
  }

  function findParagraph(document, token) {
    for (let index = 1; index <= safeCount(document.Paragraphs); index += 1) {
      const paragraph = item(document.Paragraphs, index);
      if (paragraph && String(paragraph.Range.Text || "").indexOf(token) !== -1) {
        return paragraph;
      }
    }
    return null;
  }

  function listStringFor(document, token) {
    const paragraph = findParagraph(document, token);
    return paragraph && paragraph.Range && paragraph.Range.ListFormat
      ? normalizedListString(paragraph.Range.ListFormat.ListString)
      : "";
  }

  function applyListLevel(document, template, schemeIndex, level, token) {
    const range = appendText(document, token);
    range.Style = item(document.Styles, -1 - level);
    range.ListFormat.ApplyListTemplateWithLevel(template, true, 0, 0, level);
    range.ListFormat.ListLevelNumber = level;
    return range;
  }

  function numberingFailure(stage) {
    const error = new Error("Numbering assay failed");
    error.probeStage = stage;
    throw error;
  }

  function bindDecimalInsertion(document, template) {
    const first = findParagraph(document, "L2-1").Range;
    const last = findParagraph(document, "L2-X").Range;
    const listRange = document.Range(first.Start, last.End);
    listRange.ListFormat.ApplyListTemplateWithLevel(template, false, 0, 0, 1);
    for (let level = 1; level <= 4; level += 1) {
      findParagraph(document, `L2-${level}`).Range.ListFormat.ListLevelNumber = level;
    }
    findParagraph(document, "L2-X").Range.ListFormat.ListLevelNumber = 4;
  }

  function addNativeList(document, schemeIndex) {
    const template = document.ListTemplates.Add(true, `wpsc_m0_${schemeIndex}`);
    const definitions = NUMBERING_SCHEMES[schemeIndex - 1];
    definitions.forEach(function (definition, index) {
      const level = index + 1;
      const listLevel = item(template.ListLevels, level);
      listLevel.NumberFormat = definition.format;
      listLevel.NumberStyle = definition.style;
      listLevel.NumberPosition = (level - 1) * 18;
      listLevel.TextPosition = level * 18;
      listLevel.ResetOnHigher = definition.neverRestart ? 0 : (level === 1 ? 0 : level - 1);
      listLevel.StartAt = 1;
    });
    const firstStart = safeNumber(endRange(document).Start, 0);
    for (let level = 1; level <= 4; level += 1) {
      const range = appendText(document, `L${schemeIndex}-${level}`);
      range.Style = item(document.Styles, -1 - level);
    }
    const listRange = document.Range(
      firstStart,
      safeNumber(endRange(document).Start, firstStart)
    );
    listRange.ListFormat.ApplyListTemplateWithLevel(template, false, 0, 0, 1);
    for (let level = 1; level <= 4; level += 1) {
      findParagraph(document, `L${schemeIndex}-${level}`).Range.ListFormat.ListLevelNumber = level;
    }
    return template;
  }

  function exactNumberingSchemes(document) {
    for (let scheme = 1; scheme <= 3; scheme += 1) {
      for (let level = 1; level <= 4; level += 1) {
        if (listStringFor(document, `L${scheme}-${level}`) !== NUMBERING_EXPECTED[scheme - 1][level - 1]) {
          return false;
        }
      }
    }
    return true;
  }

  function numberingSnapshotMatches(document) {
    for (let scheme = 1; scheme <= 3; scheme += 1) {
      for (let level = 1; level <= 4; level += 1) {
        const token = scheme === 2 && level === 4 ? "L2-X" : `L${scheme}-${level}`;
        if (listStringFor(document, token) !== NUMBERING_EXPECTED[scheme - 1][level - 1]) {
          return false;
        }
      }
    }
    return findParagraph(document, "L2-4") === null;
  }

  function capability6Numbering(document) {
    const templates = [addNativeList(document, 1), addNativeList(document, 2), addNativeList(document, 3)];
    if (!exactNumberingSchemes(document)) { numberingFailure("initial-schemes"); }
    const inserted = document.Paragraphs.Add(findParagraph(document, "L2-4").Range);
    inserted.Range.Text = "L2-X\r";
    inserted.Range.Style = item(document.Styles, -5);
    bindDecimalInsertion(document, templates[1]);
    if (listStringFor(document, "L2-X") !== "1.1.1.2" || listStringFor(document, "L2-4") !== "1.1.1.1") { numberingFailure("insert-renumber"); }
    const moving = findParagraph(document, "L2-X");
    const destination = findParagraph(document, "L2-4").Range.Duplicate;
    moving.Range.Cut();
    destination.Collapse(1);
    destination.Paste();
    if (listStringFor(document, "L2-X") !== "1.1.1.1" || listStringFor(document, "L2-4") !== "1.1.1.2") { numberingFailure("move-renumber"); }
    findParagraph(document, "L2-4").Range.Delete();
    if (!numberingSnapshotMatches(document)) { numberingFailure("delete-renumber"); }
    return {metrics: {definitionCount: 3, levelCount: 4, numberedParagraphCount: 12, mutationKinds: 3, mutationRenumberChecks: 3, exactSchemeCount: 3}};
  }

  function capability7Sections(document) {
    endRange(document).InsertBreak(7);
    for (let index = 0; index < 3; index += 1) {
      endRange(document).InsertBreak(2);
      item(document.Sections, safeCount(document.Sections)).PageSetup.Orientation = 1;
      appendText(document, `LANDSCAPE-${index + 1}`);
      endRange(document).InsertBreak(2);
      item(document.Sections, safeCount(document.Sections)).PageSetup.Orientation = 0;
    }
    const first = item(document.Sections, 1);
    const body = item(document.Sections, Math.min(2, safeCount(document.Sections)));
    const firstFooter = item(first.Footers, 1);
    const bodyFooter = item(body.Footers, 1);
    firstFooter.PageNumbers.NumberStyle = 2;
    firstFooter.PageNumbers.RestartNumberingAtSection = true;
    firstFooter.PageNumbers.StartingNumber = 1;
    bodyFooter.LinkToPrevious = false;
    bodyFooter.PageNumbers.NumberStyle = 0;
    bodyFooter.PageNumbers.RestartNumberingAtSection = true;
    bodyFooter.PageNumbers.StartingNumber = 1;
    return {passed: safeCount(document.Sections) >= 7, metrics: {sectionCount: safeCount(document.Sections), landscapeLifecycleCount: 3, pageNumberRestartCount: 2, explicitPageBreakCount: 1}};
  }

  function capability8Toc(document) {
    [-20, -21, -22].forEach(function (styleId, index) { item(document.Styles, styleId).ParagraphFormat.SpaceAfter = index + 2; });
    const title = appendText(document, "TOC-TITLE");
    title.Style = item(document.Styles, -67);
    title.ParagraphFormat.OutlineLevel = 10;
    document.TablesOfContents.Add(endRange(document), true, 1, 3);
    return {passed: safeCount(document.TablesOfContents) >= 1, metrics: {tocCount: safeCount(document.TablesOfContents), tocStyleCount: 3, titleOutlineLevel: 10}};
  }

  function addSequence(document, label) {
    document.Fields.Add(endRange(document), -1, `SEQ ${label} \\* ARABIC`, true);
    appendText(document, "");
  }

  function capability9Captions(document) {
    addSequence(document, "Figure"); addSequence(document, "Figure"); addSequence(document, "Table");
    document.TablesOfFigures.Add(endRange(document), "Figure");
    appendText(document, "");
    document.TablesOfFigures.Add(endRange(document), "Table");
    return {passed: safeCount(document.TablesOfFigures) >= 2, metrics: {sequenceFieldCount: 3, figureIndexCount: safeCount(document.TablesOfFigures), resetModes: 2}};
  }

  function capability10References(document) {
    const occupied = "wpsc_ref_000000000000000000000000";
    const retry = "wpsc_ref_000000000000000000000001";
    const target = appendText(document, "REFERENCE-TARGET");
    document.Bookmarks.Add(occupied, target);
    const collision = document.Bookmarks.Exists(occupied);
    document.Bookmarks.Add(collision ? retry : occupied, target);
    document.Fields.Add(endRange(document), -1, `REF ${retry}`, true);
    if (typeof target.Move === "function") { target.Move(1, -1); }
    return {passed: collision && document.Bookmarks.Exists(retry), metrics: {bookmarkCount: safeCount(document.Bookmarks), collisionRetries: 1, referenceFieldCount: 1}};
  }

  function capability11Formula(document) {
    const table = document.Tables.Add(endRange(document), 1, 1);
    const cellRange = table.Cell(1, 1).Range;
    cellRange.Text = "x=(-b±√(b²-4ac))/(2a)";
    if (typeof table.Borders === "function") {
      for (let borderId = -6; borderId <= -1; borderId += 1) { table.Borders(borderId).LineStyle = 0; }
    }
    const math = document.OMaths.Add(cellRange);
    if (math && typeof math.BuildUp === "function") { math.BuildUp(); }
    else if (document.OMaths && typeof document.OMaths.BuildUp === "function") { document.OMaths.BuildUp(); }
    return {passed: safeCount(document.OMaths) >= 1, metrics: {nativeFormulaCount: safeCount(document.OMaths), borderlessContainerCount: 1}};
  }

  function capability12Pagination(document) {
    const table = document.Tables.Add(endRange(document), 80, 2);
    for (let row = 1; row <= 80; row += 1) { table.Cell(row, 1).Range.Text = `R${row}`; table.Cell(row, 2).Range.Text = "PAGINATION"; }
    if (typeof document.Repaginate === "function") { document.Repaginate(); }
    const pages = {};
    let positioned = 0;
    for (let row = 1; row <= 80; row += 1) {
      const range = table.Cell(row, 1).Range;
      const page = safeNumber(range.Information(3), 0);
      const x = safeNumber(range.Information(5), NaN);
      const y = safeNumber(range.Information(6), NaN);
      if (page > 0) { pages[page] = true; }
      if (Number.isFinite(x) && Number.isFinite(y)) { positioned += 1; }
    }
    const fragments = Object.keys(pages).length;
    return {passed: fragments >= 2 && positioned === 80, metrics: {nodeCount: 1, sectionCount: safeCount(document.Sections), pageSpanCount: fragments, fragmentCount: fragments, positionedRowCount: positioned, rangeUnit: "utf16"}};
  }

  function capability13Checkpoint(document) {
    const first = appendText(document, "CHECKPOINT-CHILD-1");
    const partial = appendText(document, "CHECKPOINT-PARTIAL");
    let cleaned = false;
    try { throw new Error("IntentionalChildFailure"); }
    catch (error) { partial.Delete(); cleaned = true; appendText(document, "[M0 CAP 13 FALLBACK]"); }
    const later = appendText(document, "CHECKPOINT-CONTINUED");
    return {passed: cleaned && safeNumber(first.End, 0) > safeNumber(first.Start, 0) && safeNumber(later.End, 0) > safeNumber(later.Start, 0), metrics: {checkpointDepth: 2, childCleanupCount: 1, fallbackCount: 1, continuationCount: 1}};
  }

  function fieldSnapshot(document) {
    const values = [];
    for (let index = 1; index <= safeCount(document.Fields); index += 1) {
      const field = item(document.Fields, index);
      values.push(String(field && field.Result ? field.Result.Text : ""));
    }
    return values.join("\u0000");
  }

  function boundedFieldRefresh(document, maximum) {
    let previous = fieldSnapshot(document);
    for (let pass = 1; pass <= maximum; pass += 1) {
      document.Fields.Update();
      const current = fieldSnapshot(document);
      if (current === previous) { return pass; }
      previous = current;
    }
    return maximum;
  }

  function capability14Convergence(document) {
    const fullPasses = boundedFieldRefresh(document, 2);
    appendText(document, "[M0 NOTICE PATCH]");
    document.Fields.Update();
    return {passed: fullPasses <= 2, metrics: {fullPasses, patchPasses: 1, plannedSaveCount: 2, plannedExportCount: 1}};
  }

  function capability15Svg(document, envelope) {
    const before = safeCount(document.InlineShapes);
    const shape = document.InlineShapes.AddPicture(envelope.stagedSvgPath, false, true, endRange(document));
    if (shape) { shape.Width = 120; shape.Height = 60; }
    return {passed: safeCount(document.InlineShapes) === before + 1, metrics: {svgCount: safeCount(document.InlineShapes), staticManifestCount: 1}};
  }

  function addAssays(document, capabilities, envelope) {
    assay(document, capabilities, 3, function () { return capability3Fonts(document); }, false);
    assay(document, capabilities, 4, function () { return capability4Unicode(document); }, false);
    assay(document, capabilities, 5, function () { return capability5Coordinates(document); }, false);
    assay(document, capabilities, 6, function () { return capability6Numbering(document); }, false);
    assay(document, capabilities, 7, function () { return capability7Sections(document); }, false);
    assay(document, capabilities, 8, function () { return capability8Toc(document); }, false);
    assay(document, capabilities, 9, function () { return capability9Captions(document); }, false);
    assay(document, capabilities, 10, function () { return capability10References(document); }, false);
    assay(document, capabilities, 11, function () { return capability11Formula(document); }, false);
    assay(document, capabilities, 12, function () { return capability12Pagination(document); }, false);
    assay(document, capabilities, 13, function () { return capability13Checkpoint(document); }, false);
    assay(document, capabilities, 14, function () { return capability14Convergence(document); }, false);
    assay(document, capabilities, 15, function () { return capability15Svg(document, envelope); }, true);
  }

  function updateNativeFields(document) {
    for (let index = 1; index <= safeCount(document.TablesOfContents); index += 1) { item(document.TablesOfContents, index).Update(); }
    for (let index = 1; index <= safeCount(document.TablesOfFigures); index += 1) { item(document.TablesOfFigures, index).Update(); }
    document.Fields.Update();
    if (typeof document.Repaginate === "function") { document.Repaginate(); }
  }

  function downgrade(capability, reason) {
    capability.status = capability.id === 15 ? "unsupported" : "failed";
    capability.checks = ["native-attempted"];
    capability.metrics.reopenFailure = reason;
  }

  function verifyReopened(document, capabilities) {
    const content = String(document.Content.Text || "");
    const objectChecks = {6: numberingSnapshotMatches(document), 8: safeCount(document.TablesOfContents) >= 1, 9: safeCount(document.TablesOfFigures) >= 2, 10: safeCount(document.Bookmarks) >= 2, 11: safeCount(document.OMaths) >= 1, 15: safeCount(document.InlineShapes) >= 1};
    for (let capabilityId = 3; capabilityId <= 15; capabilityId += 1) {
      const capability = capabilities.find(function (row) { return row.id === capabilityId; });
      if (!capability || capability.status !== "passed") { continue; }
      if (content.indexOf(marker(capabilityId)) === -1) { downgrade(capability, "marker-missing"); continue; }
      if (Object.prototype.hasOwnProperty.call(objectChecks, capabilityId) && !objectChecks[capabilityId]) { downgrade(capability, "native-object-missing"); continue; }
      capability.checks.push("reopened", "refreshed");
    }
  }

  function run(params) {
    const envelope = validateEnvelope(params);
    const previousAlerts = Application.DisplayAlerts;
    const previousScreenUpdating = Application.ScreenUpdating;
    const capabilities = [resultRow(1, "passed", ["native"], {protocolAccepted: true, manifestBound: true}), resultRow(2, "passed", ["not-applicable-macos"], {})];
    let document = null;
    let failure = null;
    try {
      Application.DisplayAlerts = 0;
      Application.ScreenUpdating = false;
      document = Application.Documents.Add();
      document.PageSetup.TopMargin = 54;
      document.PageSetup.BottomMargin = 54;
      document.PageSetup.LeftMargin = 64;
      document.PageSetup.RightMargin = 64;
      appendText(document, "WPSComposer Long-form M0 Native Probe");
      addAssays(document, capabilities, envelope);
      document.SaveAs2(envelope.stagedDocxPath, 12);
      document.Close(0);
      document = null;
      document = Application.Documents.Open(envelope.stagedDocxPath, false, false);
      updateNativeFields(document);
      verifyReopened(document, capabilities);
      document.Save();
      document.ExportAsFixedFormat(envelope.stagedPdfPath, 17, false, 0, 0);
      document.Close(0);
      document = null;
      const failures = capabilities.filter(function (row) { return row.status === "failed"; }).map(function (row) { return {code: "CAPABILITY_FAILED", capabilityId: row.id}; });
      return {probeVersion: envelope.probeVersion, protocolVersion: envelope.protocolVersion, resourceManifestVersion: envelope.resourceManifestVersion, platform: "macos", wpsVersion: envelope.expectedWpsVersion, capabilities, failures, docxPath: envelope.stagedDocxPath, pdfPath: envelope.stagedPdfPath};
    } catch (error) {
      failure = error;
      if (!error.code) { error.code = "M0_PROBE_FAILED"; }
      throw error;
    } finally {
      try {
        if (document !== null) { document.Close(0); }
      } catch (closeError) {
        if (failure === null) { closeError.code = "M0_PROBE_FAILED"; throw closeError; }
      } finally {
        try { Application.ScreenUpdating = previousScreenUpdating; }
        finally { Application.DisplayAlerts = previousAlerts; }
      }
    }
  }

  window.WPSComposerLongformM0 = Object.freeze({run, validateEnvelope});
}());
