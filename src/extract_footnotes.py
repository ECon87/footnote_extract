"""
Extract footnotes (and endnotes) from a .docx together with the sentence
each one supports.

Usage:
    python extract_footnotes.py input.docx [output.csv]

Output CSV columns:
    kind, number, statement, paragraph, note_text

- "statement" is the full sentence containing the footnote mark. If the mark
  sits after the sentence's closing punctuation (the usual case), that is the
  sentence just before the mark.
- "paragraph" is the whole paragraph, for context when sentence splitting
  is imperfect.
No third-party packages needed: it reads the .docx XML directly.
"""
import csv
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SKIP_TYPES = {"separator", "continuationSeparator", "continuationNotice"}

# Abbreviations that end in a period but do not end a sentence
# (tuned for legal/academic writing).
ABBREV = {
    "v", "vs", "u.s", "inc", "corp", "co", "ltd", "llc", "no", "nos", "fed",
    "f", "supp", "cir", "e.g", "i.e", "etc", "id", "cf", "mr", "ms", "mrs",
    "dr", "prof", "al", "art", "sec", "ch", "p", "pp", "para", "fig", "vol",
    "ed", "eds", "rev", "cal", "app", "ct", "s.ct", "l.ed", "n.d", "s.d",
    "e.d", "w.d", "d", "jan", "feb", "mar", "apr", "jun", "jul", "aug",
    "sep", "sept", "oct", "nov", "dec", "st", "jr", "sr", "assn", "dept",
}

BOUNDARY = re.compile(r"[.!?][\"'”’)\]]*\s+(?=[\"'“‘(\[]?[A-Z0-9])")
ENDS_SENTENCE = re.compile(r"[.!?][\"'”’)\]]*$")


def is_real_boundary(text, m):
    """Reject boundaries that come right after a known abbreviation or an initial."""
    before = text[: m.start()]
    word = re.split(r"\s", before)[-1].lower().strip("(\"'“")
    if word in ABBREV or re.fullmatch(r"[a-z]", word):
        return False
    if re.fullmatch(r"([a-z]\.)+[a-z]", word):   # initialisms: S.D.N.Y., U.S.C.
        return False
    return True


def boundaries(text):
    return [m.end() for m in BOUNDARY.finditer(text) if is_real_boundary(text, m)]


def note_texts(z, part, tag):
    if part not in z.namelist():
        return {}
    root = ET.fromstring(z.read(part))
    out = {}
    for note in root.iter(W + tag):
        if note.get(W + "type") in SKIP_TYPES:
            continue
        paras = []
        for p in note.iter(W + "p"):
            paras.append("".join(t.text or "" for t in p.iter(W + "t")).strip())
        out[note.get(W + "id")] = " ".join(x for x in paras if x)
    return out


def paragraph_text_and_marks(p):
    """Return the paragraph's plain text and a list of (pos, kind, id) for note marks."""
    chunks, marks, length = [], [], 0
    for el in p.iter():
        if el.tag == W + "t":
            s = el.text or ""
        elif el.tag in (W + "tab", W + "br"):
            s = " "
        elif el.tag == W + "footnoteReference":
            marks.append((length, "footnote", el.get(W + "id")))
            continue
        elif el.tag == W + "endnoteReference":
            marks.append((length, "endnote", el.get(W + "id")))
            continue
        else:
            continue
        chunks.append(s)
        length += len(s)
    return "".join(chunks), marks


def sentence_for(text, pos):
    pre = text[:pos].rstrip()
    # start: last real boundary strictly before the closing punctuation of pre
    starts = [b for b in boundaries(text) if b <= len(pre) - 1]
    start = starts[-1] if starts else 0
    if ENDS_SENTENCE.search(pre):          # mark sits after the period
        end = len(pre)
    else:                                   # mark is mid-sentence
        later = [b for b in boundaries(text) if b > pos]
        end = later[0] if later else len(text)
    return text[start:end].strip()


def extract(path):
    with zipfile.ZipFile(path) as z:
        notes = {
            "footnote": note_texts(z, "word/footnotes.xml", "footnote"),
            "endnote": note_texts(z, "word/endnotes.xml", "endnote"),
        }
        body = ET.fromstring(z.read("word/document.xml"))

    rows, counters = [], {"footnote": 0, "endnote": 0}
    for p in body.iter(W + "p"):
        text, marks = paragraph_text_and_marks(p)
        for pos, kind, nid in marks:
            counters[kind] += 1
            rows.append({
                "kind": kind,
                "number": counters[kind],
                "statement": sentence_for(text, pos),
                "paragraph": text.strip(),
                "note_text": notes[kind].get(nid, ""),
            })
    return rows


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else re.sub(r"\.docx$", "", src, flags=re.I) + "_footnotes.csv"
    rows = extract(src)
    with open(dst, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig opens cleanly in Excel
        w = csv.DictWriter(f, fieldnames=["kind", "number", "statement", "paragraph", "note_text"])
        w.writeheader()
        w.writerows(rows)
    print(f"{len(rows)} notes written to {dst}")


if __name__ == "__main__":
    main()
