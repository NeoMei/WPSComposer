(function () {
  "use strict";

  const PROBE_VERSION = "0.8.0-m0.1";
  const PROTOCOL_VERSION = 2;
  const RESOURCE_MANIFEST_VERSION = 1;
  const EMPTY_MANIFEST_DIGEST = "a6a20076da005b27c9afc3a5d5b2457798c0ac817d1abc38b2fee4398ac3f133";
  const M0_KEYS = [
    "manifest",
    "probeVersion",
    "protocolVersion",
    "resourceManifestVersion",
    "stagedDocxPath",
    "stagedPdfPath"
  ];
  const MANIFEST_KEYS = ["digest", "entries", "version"];

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
    return actual.length === wanted.length && actual.every(function (key, index) {
      return key === wanted[index];
    });
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
    if (params.protocolVersion !== PROTOCOL_VERSION) {
      fail("PROTOCOL_MISMATCH", "Long-form M0 protocol version is incompatible");
    }
    if (params.resourceManifestVersion !== RESOURCE_MANIFEST_VERSION) {
      fail("PROTOCOL_MISMATCH", "Long-form M0 resource manifest version is incompatible");
    }
    if (params.probeVersion !== PROBE_VERSION) {
      fail("PROTOCOL_MISMATCH", "Long-form M0 probe version is incompatible");
    }
    requirePath(params.stagedDocxPath, "stagedDocxPath");
    requirePath(params.stagedPdfPath, "stagedPdfPath");
    const manifest = params.manifest;
    if (!hasExactKeys(manifest, MANIFEST_KEYS) || manifest.version !== RESOURCE_MANIFEST_VERSION || !Array.isArray(manifest.entries) || manifest.entries.length !== 0 || typeof manifest.digest !== "string" || manifest.digest !== EMPTY_MANIFEST_DIGEST) {
      fail("RESOURCE_MANIFEST_MISMATCH", "Long-form M0 resource manifest is invalid");
    }
    return params;
  }

  function run(params) {
    const envelope = validateEnvelope(params);
    return {
      schemaVersion: 1,
      probeVersion: envelope.probeVersion,
      protocolVersion: envelope.protocolVersion,
      resourceManifestVersion: envelope.resourceManifestVersion,
      status: "protocol-accepted-native-probe-not-run"
    };
  }

  window.WPSComposerLongformM0 = Object.freeze({
    run,
    validateEnvelope
  });
}());
