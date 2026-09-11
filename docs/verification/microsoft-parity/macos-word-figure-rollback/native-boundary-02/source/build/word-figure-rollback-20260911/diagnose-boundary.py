"""Retain exact guard failure facts; no rollback rule is relaxed."""
from pathlib import Path
import importlib.util
import sys
path=Path(__file__).with_name('macos_word_figure_rollback_probe.py')
spec=importlib.util.spec_from_file_location('figure_boundary_probe',path)
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
original=m._rollback_commands

def diagnostic(*args):
    lines=original(*args)
    old='if probeRollbackEnd is not probeMutationEnd then error "WPSC_FIGURE_PROBE_BOUND_CHANGED"'
    assert old in lines
    return [line if line!=old else 'if probeRollbackEnd is not probeMutationEnd then error ("WPSC_FIGURE_PROBE_BOUND_CHANGED expected=" & (probeMutationEnd as text) & " actual=" & (probeRollbackEnd as text) & " picture=" & (probePictureStart as text) & ":" & (probePictureEnd as text) & " suffix=" & (content of text object of probeSuffixBookmark as text))' for line in lines]
m._rollback_commands=diagnostic
m.SOURCES.append(Path(__file__))
raise SystemExit(m.main(sys.argv[1:]))
