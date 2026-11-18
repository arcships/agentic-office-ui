# OpenDocument text fixtures

`sample-html-features.odt` and `sample-html-features.fodt` are the paired
packaged/flat OpenDocument text fixtures used by the OpenDocument import tests.
They exercise headings, inline character formatting, hyperlinks, bullet and
numbered lists (including a nested level), a table with a horizontal span, a
vertical merge and covered cells, an inline PNG and page setup.

`frames-and-fields.fodt` exercises text frames (inline and floating), a nested
picture and a table inside a frame, a footer with page fields, and comment
anchors.

The files were produced by LibreOffice and copied from the public
`@extend-ai/react-docx` test fixtures (`tests/fixtures/opendocument/`, MIT) so
the Vue reader is checked against the same upstream documents the WebAssembly
parser uses. They contain no customer data.

SHA-256:

- `sample-html-features.odt`: `cfacb64f1904d0b0ebc8908e39d40125f3a3902e54eb5419356a20afb25e2841`
- `sample-html-features.fodt`: `adb52b44d19cca7cd1c6ac9407a6f4beb7225c2f5693c494d8c535ee4c7c5cd1`
- `frames-and-fields.fodt`: `315d344dd78f0efab877a0e1a882bad51cd5aff0d51b59c66f20f6fb02fec959`
