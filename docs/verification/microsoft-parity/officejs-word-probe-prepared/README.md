# Word Office.js capability/binding probe — OFFLINE PREPARATION

**The manifest requests `ReadWriteDocument`, not a technically enforced read-only permission.** Microsoft requires that permission for application-specific APIs even when the add-in only reads. This package implements only the reads listed below. Enabling it therefore requires reviewing both the permission and these exact sources. [Microsoft permission model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins).

Nothing has been installed, trusted, served, launched or run inside Office. There is no bridge, registry, token pairing, document mutation, persistent document identity, snapshot collector or parity implementation. This is a reviewable candidate for a later controlled native probe.

## Files and allowed behavior

- `manifest.word.xml.template` is an add-in-only task-pane manifest for Word (`Host Name="Document"`), minimum WordApi 1.1. `manifest.word.xml` is its concrete rendering for **https://localhost:3443/taskpane.html**. Higher requirements are checked at runtime so missing optional sets are visible. No activation command or automatic document-open hook is declared. [Manifest overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview).
- `configure.py` accepts an exact HTTPS loopback origin (localhost, 127.0.0.1 or [::1]) and writes only `manifest.word.xml` in this directory, refusing to overwrite. It does not start anything. To change origin, move the existing generated manifest to a package-local backup, then run `python3 -B configure.py --origin https://localhost:3443` with the intended port. Review the generated URL before sideloading.
- `taskpane.html`, `taskpane.css`, `taskpane.js` show JSON for manual copy or an explicit user-triggered local Blob download. The only remote script is the [official production Office.js CDN](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn), loaded in the head. There are no package telemetry requests. The Office SDK/platform has its own network behavior; no SDK code is downloaded or vendored into this package.
- `probe.js` creates a cryptographically random, memory-only page instance ID. It is a **pane lifetime label, not proof of document identity**. It is not persisted in document settings, custom properties, bookmarks, storage or cookies. Closing/reloading retires it; real host document switching, duplicated windows and pane reuse still need native validation.
- Requirement reporting uses actual `Office.context.diagnostics` host/platform/version, `document.mode`, URL *availability only*, and `isSetSupported`. [Diagnostics](https://learn.microsoft.com/en-us/javascript/api/office/office.contextinformation?view=common-js) and [Document context](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=common-js) provide these values. Missing/error/nonboolean flags remain unknown, separate from false. No support is inferred from a version string. Method prototype exposure is labeled separately from native invocation.
- The explicit read button gates on reported desktop Word (Mac/PC), WordApi 1.1 and Word.run; it loads only `document.saved`. If WordApiDesktop 1.4 is reported, it also reads the current document's selection `start/end/storyType`. One sync is awaited, scalar types checked, errors recorded by code without document-bearing messages. It reads no document text, OOXML or raw path. [Saved flag](https://learn.microsoft.com/en-us/javascript/api/word/word.document?view=word-js-preview#word-word-document-saved-member) and [range coordinates](https://learn.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview) are read candidates, not established preservation guarantees.
- Only one read can be pending per pane. A 10-second local deadline retires the pane; it does **not** cancel an Office operation. Late results cannot become an ACK. Any read error also retires the pane. Reconcile before reloading after errors; there is no retry loop.

## Requirement reporting for the nine gaps and snapshot

The report includes WordApi 1.1–1.5, WordApiDesktop 1.2–1.4, File 1.1 and CompressedFile 1.1. These are the relevant documented sets from the existing gap design; they are not a declaration that every member is available or that the method contract is satisfied. [Word requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets), [Common sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/office-add-in-requirement-sets).

The nine entries are semantic-table native/fallback, captioned-figure native/fallback, equation native/fallback, heading-level native, horizontal line and WordArt. Every entry keeps `nativeAccepted:false`. Mapping notes identify OOXML/table/picture/field/font/drawing candidates and unresolved native type, placement, numbering, preset or recovery semantics. The snapshot entry reports File/CompressedFile flags and `getFileAsync` presence but always `executed:false`. **No getFileAsync call or snapshot success occurs in this package.**

## Reproduce the offline checks

From this package directory, using installed runtimes only:

```sh
node --test tests/*.test.cjs
python3 -B -m unittest discover -s tests -p 'test_*.py'
node --check probe.js
node --check taskpane.js
```

Tests use fake Office/Word objects and a fake task-pane DOM. They do not launch a browser, Office, network listener or document. The Python tests check manifest structure and closed loopback configuration, not Microsoft's full XSD/store validation. No npm packages are needed. `evidence/red-js.log` and `evidence/red-manifest.log` preserve initial failures; final logs are retained alongside them.

## Manual enable plan — future approval and native lease required

1. Approve the exact manifest ID **8abccf20-6807-4593-85f6-d7d039b9d68f**, origin, `ReadWriteDocument` permission, source hashes, fixture documents, and setup/removal impact. Establish an owned test document and an unrelated sentinel. Never restart Office while unrelated unsaved work remains open.
2. Separately authorize a static HTTPS server bound only to the loopback address, serving this directory. It must serve the four web assets with correct content types and no command endpoints. No server implementation or launch is included. Use a certificate already valid for the chosen loopback name; if trust installation is necessary, stop for explicit certificate/trust approval. Do not suppress certificate errors. HTTPS hosting and trusted development certificates are documented by [Microsoft](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).
3. **Mac:** with a recorded preimage of existing add-ins, copy this exact generated manifest into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`; create that directory only if approved. Open/restart Word only after the unrelated-document gate; open the owned fixture, then **Home → Add-ins → Word read-only capability probe**. These are Microsoft's [Mac sideload steps](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).
4. **Windows:** use an already authorized test shared-folder catalog, or separately approve a dedicated share/catalog. Put only the manifest in that catalog. Office **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs** adds the catalog URL and **Show in Menu**; this is a security-setting change requiring separate approval. Reopen only under the document gate, then **Home → Add-ins → Advanced → Shared Folder**, choose this probe and Add. No registry script is provided. [Microsoft Windows sideload procedure](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).
5. Record readiness, host/platform, version, exact generated source/manifest hashes and pane ID. Manually report flags first. Only on the approved fixture and lease, click the optional read button. Export/copy JSON locally. Compare saved/path/selection state independently; this package alone cannot prove nonmutation or identity. Repeat saved/unsaved, focus-switch, duplicated-window and pane-reload cases only as separately bounded native work. Older pane IDs must never be used as native document selectors.

## Manual removal plan — future coordinated approval

Close the probe pane, retire its recorded ID and stop its approved static server. Close/discard only owned fixtures after exact identity checks. Remove the dedicated Windows catalog/share registration via the same catalog UI only if it was added for this task. Certificate removal is separate and must target only an explicitly created task certificate; this package creates none.

Microsoft's removal guidance requires clearing Office caches and warns against deleting individual cached manifests. Cache cleanup can affect **all sideloaded add-ins**, so first inventory existing add-ins and obtain explicit coordinated cleanup approval. Do not present deleting just this XML from a Wef cache as a verified uninstall. [Office cache/removal guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache).

On an approved isolated Mac profile, after closing affected apps, Microsoft's manual procedure clears contents of `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` and Word's `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`. If OsfWebHost is absent, consult the same current guide for the host-specific cache alternative before changing paths. On an approved Windows test profile, follow the guide's manual web/Wef cache procedure; record the exact path and catalog removal before clearing. Restore/re-enable unrelated sideloaded add-ins from their inventoried manifests if required, then verify this probe is absent and unrelated add-ins still work. No broad cache-delete command is bundled or executed.

## Remaining gates

Native installation/removal, document-bound lifetime and multiwindow identity, read-only/protected fixtures, non-mutating native compressed snapshots, sliced-handle cleanup, OOXML insertion, semantic contract tests and Windows/Mac parity all remain open. Office.js batch success is not a transaction guarantee. This package must not be wired into public dispatch or treated as a production binding registry.
