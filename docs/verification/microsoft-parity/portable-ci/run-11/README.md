# Private Excel candidate hosted regression

Run 34488660705 at 70d3d86ff6b03945dacccffd1688922830ad6879 fails both Windows jobs with 79 failures each, all reporting `ValueError: Process executable must be absolute`. Both Linux jobs pass. Raw four-job artifacts, metadata and failed-job logs are retained unchanged. The identity model applied host-native path parsing to a Darwin executable identity, rejecting `/Applications/...` on Windows. A portable regression and repair are in progress; this failed run is not rewritten. This is portable CI evidence, not native Office acceptance.
