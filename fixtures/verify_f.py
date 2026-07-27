"""Checklist F: snapshot_to_patches round-trip fidelity."""
from pathlib import Path

from skills.WPSComposer import edit, inspect, snapshot_to_patches

FIX = Path(__file__).parent
OUT = FIX / "out"


def check(name, fn):
    try:
        fn()
        print(f"PASS {name}")
        return True
    except AssertionError as exc:
        print(f"FAIL {name}: {exc}")
        return False
    except Exception as exc:
        print(f"ERROR {name}: {type(exc).__name__}: {exc}")
        return False


def f_roundtrip():
    # Replay target: same structure as the source (generated from identical
    # markdown) — cross-document replay resolves positionally.
    from skills.WPSComposer import generate
    src_md = FIX / "_replay.md"
    src_md.write_text(
        "# Sample Title\n\n## Section One\n\n"
        "First paragraph of the sample document.\n\n"
        "Second paragraph with some more text content.\n\n"
        "Third paragraph completes the trio.\n\n"
        "| Name | Value |\n|------|-------|\n| alpha | 1 |\n| beta | 2 |\n\n"
        "## Section Two\n\nClosing paragraph for the sample.\n",
        encoding="utf-8",
    )
    replay_base = OUT / "replay_base.docx"
    replay_base.unlink(missing_ok=True)
    (OUT / "replay.docx").unlink(missing_ok=True)
    generate(str(src_md), format="docx", output=str(replay_base))
    src_md.unlink()
    from skills.WPSComposer.scripts.writer import WriterComposer
    with WriterComposer.open_document(str(replay_base)) as writer:
        writer._doc.Shapes.AddTextbox(1, 100, 100, 200, 50).TextFrame.TextRange.Text = "fixture textbox"
        writer.save(str(replay_base))

    snap = inspect(str(FIX / "sample.docx"))
    patches = snapshot_to_patches(snap, dimensions=("font",))
    assert patches, "no patches produced"
    out = OUT / "replay.docx"
    r = edit(str(replay_base), output=str(out), patches=patches)
    rejected = [
        (p["target"], p.get("rejected"))
        for p in r["patches"] if p.get("rejected")
    ]
    failed = [
        (p["target"], p.get("error"))
        for p in r["patches"] if p.get("error")
    ]
    assert not rejected, f"rejected keys (snapshot/apply asymmetry): {rejected[:5]}"
    assert not failed, f"failed patches: {failed[:5]}"
    assert r["ok"], r
    # fidelity spot-check: title font survives the round-trip
    snap_src = inspect(str(FIX / "sample.docx"))
    snap_out = inspect(str(out))
    src_title = next(p for p in snap_src["paragraphs"] if p["text"] == "Sample Title")
    out_title = next(p for p in snap_out["paragraphs"] if p["text"] == "Sample Title")
    assert out_title["font"] == src_title["font"], (
        f"title font drifted: {out_title['font']} != {src_title['font']}"
    )


def main():
    results = [check("F.round-trip font fidelity", f_roundtrip)]
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
