#!/usr/bin/env python3
"""Generate test-data/corpus/docx-comments.docx deterministically.

The fixture is the authoring counterpart to the read-path samples
(`apps/demo/public/samples/review-comments.docx`, `review-threads.docx`): three
body paragraphs, one comment thread with two replies, and one resolved thread
with one reply. Authoring commands (create/reply/resolve/delete) and their
round-trip tests load this file, then export it and reopen the result.

Only Python's standard library is used so the fixture can be regenerated in CI
without python-docx or other office tooling.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / "test-data" / "corpus" / "docx-comments.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"

COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
COMMENTS_EXTENDED_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml."
    "commentsExtended+xml"
)

CONTENT_TYPES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="{COMMENTS_CONTENT_TYPE}"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="{COMMENTS_EXTENDED_CONTENT_TYPE}"/>
</Types>
"""

PACKAGE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="{R_NS}/comments" Target="comments.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
</Relationships>
"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="character" w:styleId="CommentReference">
    <w:name w:val="annotation reference"/>
    <w:rPr><w:sz w:val="16"/></w:rPr>
  </w:style>
</w:styles>
"""

DATE = "2026-11-18T00:00:00Z"


def comment_markup(comment_id: int, text: str) -> tuple[str, str]:
    """Return (range start/text/end fragment, reference run) for a comment id."""
    body = (
        f'<w:commentRangeStart w:id="{comment_id}"/>'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
        f'<w:commentRangeEnd w:id="{comment_id}"/>'
    )
    reference = (
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{comment_id}"/></w:r>'
    )
    return body, reference


def paragraph(text: str, comment_id: int | None = None, anchor: str = "") -> str:
    if comment_id is None:
        return (
            '<w:p><w:r><w:t xml:space="preserve">'
            f"{escape(text)}</w:t></w:r></w:p>"
        )
    # Put the anchor in the middle of the paragraph so the offsets are not
    # trivially 0..len at parse time.
    before, _, after = text.partition(anchor)
    body, reference = comment_markup(comment_id, anchor)
    return (
        f"<w:p><w:r><w:t xml:space=\"preserve\">{escape(before)}</w:t></w:r>"
        f"{body}"
        f"<w:r><w:t xml:space=\"preserve\">{escape(after)}</w:t></w:r>"
        f"{reference}</w:p>"
    )


DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:w14="{W14_NS}" xmlns:w15="{W15_NS}">
  <w:body>
    {paragraph("The review thread anchors on this sentence.", 0, "review thread")}
    {paragraph("The resolved thread anchors on this sentence.", 3, "resolved thread")}
    {paragraph("The closing paragraph has no comment.")}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def comments_xml() -> str:
    comments = [
        (0, "Priya Reviewer", "PR", "1A2B3C01", "Root: confirm the review sentence."),
        (1, "Alex Reviewer", "AR", "1A2B3C02", "Reply one: confirmed this morning."),
        (2, "Morgan Editor", "ME", "1A2B3C03", "Reply two: closing the thread."),
        (3, "Alex Reviewer", "AR", "1A2B3C04", "Resolved root: wording accepted."),
        (4, "Priya Reviewer", "PR", "1A2B3C05", "Resolved reply: thanks, done."),
    ]
    body = "".join(
        f'<w:comment w:id="{comment_id}" w:author="{author}" w:initials="{initials}" '
        f'w:date="{DATE}">'
        f'<w:p w14:paraId="{para_id}">'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
        "</w:comment>"
        for comment_id, author, initials, para_id, text in comments
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:comments xmlns:w=\"{W_NS}\" xmlns:w14=\"{W14_NS}\">{body}</w:comments>"
    )


def comments_extended_xml() -> str:
    entries = [
        ("1A2B3C01", None, False),
        ("1A2B3C02", "1A2B3C01", False),
        ("1A2B3C03", "1A2B3C01", False),
        ("1A2B3C04", None, True),
        ("1A2B3C05", "1A2B3C04", False),
    ]
    body = []
    for para_id, parent, done in entries:
        attributes = [f'w15:paraId="{para_id}"']
        if parent is not None:
            attributes.append(f'w15:paraIdParent="{parent}"')
        if done:
            attributes.append('w15:done="1"')
        body.append(f"<w15:commentEx {' '.join(attributes)}/>")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w15:commentsEx xmlns:w15=\"{W15_NS}\">{''.join(body)}</w15:commentsEx>"
    )


def write_fixture() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    entries = (
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", PACKAGE_RELS),
        ("word/document.xml", DOCUMENT),
        ("word/_rels/document.xml.rels", DOCUMENT_RELS),
        ("word/styles.xml", STYLES),
        ("word/comments.xml", comments_xml()),
        ("word/commentsExtended.xml", comments_extended_xml()),
    )
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in entries:
            info = zipfile.ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o600 & 0xFFFF) << 16
            archive.writestr(info, text)


if __name__ == "__main__":
    write_fixture()
    print(f"wrote {OUTPUT}")
