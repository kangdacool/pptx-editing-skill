# -*- coding: utf-8 -*-
"""selftest.py — prove pptx_kit works without a real template.

Exercises the part that actually bit us: writing a speaker note to a notes slide whose master has
NO body placeholder (the case where python-pptx returns notes_text_frame is None). We reproduce that
by stripping the placeholder, then assert the note round-trips (write -> save -> reopen -> read).
    python selftest.py   # prints PASS / raises on failure
"""
import sys, io, os, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pptx import Presentation
from pptx_kit import new_deck, blank_slide_layout, speaker_note, fit_picture, overflows

def _strip_notes_placeholders(slide):
    ns = slide.notes_slide                       # creates the notes slide
    for ph in list(ns.placeholders):
        ph._element.getparent().remove(ph._element)   # force notes_text_frame -> None

def main():
    prs = new_deck()                             # default template, blank deck
    assert len(prs.slides) == 0, "new_deck did not clear slides"
    s = prs.slides.add_slide(blank_slide_layout(prs))

    # force the hard path: no notes body placeholder, then write a multi-line note
    _strip_notes_placeholders(s)
    assert s.notes_slide.notes_text_frame is None, "expected a placeholder-less notes master"
    NOTE = "line one\nline two with an em-dash — and 한글"
    speaker_note(s, NOTE)

    # overflow helper sanity: a shape pushed off the slide is reported
    from pptx.util import Inches
    s.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(3), Inches(2))  # runs off 13.33x7.5
    assert overflows(s), "overflows() failed to flag an off-slide shape"

    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "selftest.pptx")
        prs.save(out)
        r = Presentation(out)
        got = r.slides[0].notes_slide.notes_text_frame.text
        assert got == NOTE, f"note round-trip mismatch:\n  wrote={NOTE!r}\n  read ={got!r}"

    print("PASS  speaker_note round-trips on a placeholder-less notes master; overflows() works.")

if __name__ == "__main__":
    main()
