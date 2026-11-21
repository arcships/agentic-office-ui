# Third-Party Notices

Agentic Office UI is distributed under the Apache License 2.0 (see [LICENSE](LICENSE)). The root
license is not changed by this file and does not replace any third-party license.

This repository contains code, fonts, icons, and reference material that originate outside the
project. Every third-party project used by the repository is listed below with its license,
copyright holder, usage scope, and a verbatim copy of its license text. The same license texts are
stored under [docs/upstream/licenses/](docs/upstream/licenses/), and the Extend UI upstream file
index is [docs/upstream/extend-ui.md](docs/upstream/extend-ui.md).

If a file is adapted or copied from Extend UI, it carries a file-level header that names the pinned
commit and the MIT copyright. See the index for the exact format and for the status
(`reference-only` / `adapted` / `copied` / `not-used`) of every upstream file.

> Historical note: [docs/upstream-extend-ui.md](docs/upstream-extend-ui.md) is the earlier
> attribution note for the Extend AI React packages. This file and
> [docs/upstream/extend-ui.md](docs/upstream/extend-ui.md) supersede it for the 0.8 UI work.

## Summary

| Project | License | Copyright | Usage scope |
|---|---|---|---|
| Extend UI (extend-hq/ui) | MIT | Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn | Reference design, token, layout and behavior baseline for the 0.8 Office UI work. Adapted files carry the Extend header and are listed in docs/upstream/extend-ui.md. |
| @extend-ai/react-docx (DOCX engine port) | MIT | Copyright (c) 2026 CrowdView Inc, dba Extend | Source of the DOCX engine and editor port in packages/docx-core and packages/vue-docx (ported from upstream commit 6f70b92). File-level provenance stays in each ported module header. |
| @extend-ai/react-xlsx (XLSX chart and sheet port) | MIT | Copyright (c) 2026 CrowdView Inc, dba Extend | Source of the XLSX chart renderer and related sheet behavior adapted in packages/vue-xlsx. File-level provenance stays in each adapted module header. |
| shadcn (React primitives upstream) | MIT | Portions Copyright (c) 2023 shadcn (covered by the Extend UI notice) | Upstream primitive design that Extend UI vendors; attribution is carried by the Extend UI MIT notice that names the 2023 shadcn portions. |
| shadcn-vue (reference implementation, not a runtime dependency) | MIT | Copyright (c) 2023 unovue | Vue-side reference for primitive structure and API shape (2026-11-03 dev commit d1bf1b3). Not a runtime dependency; no file is copied yet. |
| reka-ui (runtime primitive dependency) | MIT | Copyright (c) 2023 radix-vue | Runtime dependency for the Vue primitive layer planned in #157 (reka-ui 2.11.x). Installed through pnpm; the MIT notice ships with the dependency. |
| Lucide (@lucide/vue) | ISC | Copyright (c) 2026 Lucide Icons and Contributors | Icon set for the planned OfficeIcon adapter (#156, @lucide/vue 1.76.x). Icons are consumed as a dependency, not vendored. |
| Geist and Geist Mono fonts | SIL Open Font License 1.1 | Copyright (c) 2023 Vercel, in collaboration with basement.studio | Default sans and mono font tokens for the Office theme (#156); consumed through @fontsource-variable/geist (OFL-1.1) or a host-provided Geist installation. |
| Hugeicons (@hugeicons/core-free-icons) | MIT | Copyright (c) 2025 Hugeicons | Icon data for the pixel-parity exceptions listed in docs/design/icon-policy.md (#159–#164, #170, #180; @hugeicons/core-free-icons 4.4.0); consumed as a runtime dependency, not vendored. |

## Full notices

### Extend UI (extend-hq/ui)

| Field | Value |
|---|---|
| License | MIT |
| Copyright | Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn |
| Usage scope | Reference design, token, layout and behavior baseline for the 0.8 Office UI work. Adapted files carry the Extend header and are listed in docs/upstream/extend-ui.md. |
| Source | https://github.com/extend-hq/ui |
| Pinned reference | https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e |
| License URL | https://github.com/extend-hq/ui/blob/70fe70c6f1e1eee53954c6923f3d9f583557a57e/LICENSE.md |
| License copy | [docs/upstream/licenses/extend-ui-MIT.txt](docs/upstream/licenses/extend-ui-MIT.txt) |

```text
MIT License

Copyright (c) 2026 CrowdView Inc, dba Extend
Portions Copyright (c) 2023 shadcn

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### @extend-ai/react-docx (DOCX engine port)

| Field | Value |
|---|---|
| License | MIT |
| Copyright | Copyright (c) 2026 CrowdView Inc, dba Extend |
| Usage scope | Source of the DOCX engine and editor port in packages/docx-core and packages/vue-docx (ported from upstream commit 6f70b92). File-level provenance stays in each ported module header. |
| Source | https://github.com/extend-hq/react-docx |
| Pinned reference | https://github.com/extend-hq/react-docx/tree/6f70b92 |
| License URL | https://github.com/extend-hq/react-docx/blob/main/LICENSE |
| License copy | [docs/upstream/licenses/extend-react-docx-MIT.txt](docs/upstream/licenses/extend-react-docx-MIT.txt) |

```text
MIT License

Copyright (c) 2026 CrowdView Inc, dba Extend

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### @extend-ai/react-xlsx (XLSX chart and sheet port)

| Field | Value |
|---|---|
| License | MIT |
| Copyright | Copyright (c) 2026 CrowdView Inc, dba Extend |
| Usage scope | Source of the XLSX chart renderer and related sheet behavior adapted in packages/vue-xlsx. File-level provenance stays in each adapted module header. |
| Source | https://github.com/extend-hq/react-xlsx |
| Pinned reference | https://github.com/extend-hq/react-xlsx |
| License URL | https://github.com/extend-hq/react-xlsx/blob/main/LICENSE |
| License copy | [docs/upstream/licenses/extend-react-xlsx-MIT.txt](docs/upstream/licenses/extend-react-xlsx-MIT.txt) |

```text
MIT License

Copyright (c) 2026 CrowdView Inc, dba Extend

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### shadcn (React primitives upstream)

| Field | Value |
|---|---|
| License | MIT |
| Copyright | Portions Copyright (c) 2023 shadcn (covered by the Extend UI notice) |
| Usage scope | Upstream primitive design that Extend UI vendors; attribution is carried by the Extend UI MIT notice that names the 2023 shadcn portions. |
| Source | https://github.com/shadcn-ui/ui |
| Pinned reference | https://github.com/shadcn-ui/ui |
| License URL | https://github.com/shadcn-ui/ui/blob/main/LICENSE.md |
| License copy | [docs/upstream/licenses/extend-ui-MIT.txt](docs/upstream/licenses/extend-ui-MIT.txt) |

```text
MIT License

Copyright (c) 2026 CrowdView Inc, dba Extend
Portions Copyright (c) 2023 shadcn

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### shadcn-vue (reference implementation, not a runtime dependency)

| Field | Value |
|---|---|
| License | MIT |
| Copyright | Copyright (c) 2023 unovue |
| Usage scope | Vue-side reference for primitive structure and API shape (2026-11-03 dev commit d1bf1b3). Not a runtime dependency; no file is copied yet. |
| Source | https://github.com/unovue/shadcn-vue |
| Pinned reference | https://github.com/unovue/shadcn-vue/tree/d1bf1b3a2d1a824b19f32d7253a28a43797b8253 |
| License URL | https://github.com/unovue/shadcn-vue/blob/dev/LICENSE |
| License copy | [docs/upstream/licenses/shadcn-vue-MIT.txt](docs/upstream/licenses/shadcn-vue-MIT.txt) |

```text
MIT License

Copyright (c) 2023 unovue

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### reka-ui (runtime primitive dependency)

| Field | Value |
|---|---|
| License | MIT |
| Copyright | Copyright (c) 2023 radix-vue |
| Usage scope | Runtime dependency for the Vue primitive layer planned in #157 (reka-ui 2.11.x). Installed through pnpm; the MIT notice ships with the dependency. |
| Source | https://github.com/unovue/reka-ui |
| Pinned reference | https://github.com/unovue/reka-ui/tree/v2.11.0 |
| License URL | https://github.com/unovue/reka-ui/blob/main/LICENSE |
| License copy | [docs/upstream/licenses/reka-ui-MIT.txt](docs/upstream/licenses/reka-ui-MIT.txt) |

```text
MIT License

Copyright (c) 2023 radix-vue

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### Lucide (@lucide/vue)

| Field | Value |
|---|---|
| License | ISC |
| Copyright | Copyright (c) 2026 Lucide Icons and Contributors |
| Usage scope | Icon set for the planned OfficeIcon adapter (#156, @lucide/vue 1.76.x). Icons are consumed as a dependency, not vendored. |
| Source | https://github.com/lucide-icons/lucide |
| Pinned reference | https://github.com/lucide-icons/lucide |
| License URL | https://github.com/lucide-icons/lucide/blob/main/LICENSE |
| License copy | [docs/upstream/licenses/lucide-ISC.txt](docs/upstream/licenses/lucide-ISC.txt) |

```text
ISC License

Copyright (c) 2026 Lucide Icons and Contributors

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted, provided that the above
copyright notice and this permission notice appear in all copies.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.

---

The following Lucide icons are derived from the Feather project:

airplay, alert-circle, alert-octagon, alert-triangle, aperture, arrow-down-circle, arrow-down-left, arrow-down-right, arrow-down, arrow-left-circle, arrow-left, arrow-right-circle, arrow-right, arrow-up-circle, arrow-up-left, arrow-up-right, arrow-up, at-sign, calendar, cast, check, chevron-down, chevron-left, chevron-right, chevron-up, chevrons-down, chevrons-left, chevrons-right, chevrons-up, circle, clipboard, clock, code, columns, command, compass, corner-down-left, corner-down-right, corner-left-down, corner-left-up, corner-right-down, corner-right-up, corner-up-left, corner-up-right, crosshair, database, divide-circle, divide-square, dollar-sign, download, external-link, feather, frown, hash, headphones, help-circle, info, italic, key, layout, life-buoy, link-2, link, loader, lock, log-in, log-out, maximize, meh, minimize, minimize-2, minus-circle, minus-square, minus, monitor, moon, more-horizontal, more-vertical, move, music, navigation-2, navigation, octagon, pause-circle, percent, plus-circle, plus-square, plus, power, radio, rss, search, server, share, shopping-bag, sidebar, smartphone, smile, square, table-2, tablet, target, terminal, trash-2, trash, triangle, tv, type, upload, x-circle, x-octagon, x-square, x, zoom-in, zoom-out

The MIT License (MIT) (for the icons listed above)

Copyright (c) 2013-present Cole Bemis

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### Geist and Geist Mono fonts

| Field | Value |
|---|---|
| License | SIL Open Font License 1.1 |
| Copyright | Copyright (c) 2023 Vercel, in collaboration with basement.studio |
| Usage scope | Default sans and mono font tokens for the Office theme (#156); consumed through @fontsource-variable/geist (OFL-1.1) or a host-provided Geist installation. |
| Source | https://github.com/vercel/geist-font |
| Pinned reference | https://github.com/vercel/geist-font |
| License URL | https://github.com/vercel/geist-font/blob/main/LICENSE.txt |
| License copy | [docs/upstream/licenses/geist-OFL-1.1.txt](docs/upstream/licenses/geist-OFL-1.1.txt) |

```text
Copyright (c) 2023 Vercel, in collaboration with basement.studio

This Font Software is licensed under the SIL Open Font License, Version 1.1.
This license is copied below, and is also available with a FAQ at:
http://scripts.sil.org/OFL

-----------------------------------------------------------
SIL OPEN FONT LICENSE Version 1.1 - 26 February 2007
-----------------------------------------------------------

PREAMBLE
The goals of the Open Font License (OFL) are to stimulate worldwide
development of collaborative font projects, to support the font creation
efforts of academic and linguistic communities, and to provide a free and
open framework in which fonts may be shared and improved in partnership
with others.

The OFL allows the licensed fonts to be used, studied, modified and
redistributed freely as long as they are not sold by themselves. The
fonts, including any derivative works, can be bundled, embedded,
redistributed and/or sold with any software provided that any reserved
names are not used by derivative works. The fonts and derivatives,
however, cannot be released under any other type of license. The
requirement for fonts to remain under this license does not apply
to any document created using the fonts or their derivatives.

DEFINITIONS
"Font Software" refers to the set of files released by the Copyright
Holder(s) under this license and clearly marked as such. This may
include source files, build scripts and documentation.

"Reserved Font Name" refers to any names specified as such after the
copyright statement(s).

"Original Version" refers to the collection of Font Software components as
distributed by the Copyright Holder(s).

"Modified Version" refers to any derivative made by adding to, deleting,
or substituting -- in part or in whole -- any of the components of the
Original Version, by changing formats or by porting the Font Software to a
new environment.

"Author" refers to any designer, engineer, programmer, technical
writer or other person who contributed to the Font Software.

PERMISSION AND CONDITIONS
Permission is hereby granted, free of charge, to any person obtaining
a copy of the Font Software, to use, study, copy, merge, embed, modify,
redistribute, and sell modified and unmodified copies of the Font
Software, subject to the following conditions:

1) Neither the Font Software nor any of its individual components,
in Original or Modified Versions, may be sold by itself.

2) Original or Modified Versions of the Font Software may be bundled,
redistributed and/or sold with any software, provided that each copy
contains the above copyright notice and this license. These can be
included either as stand-alone text files, human-readable headers or
in the appropriate machine-readable metadata fields within text or
binary files as long as those fields can be easily viewed by the user.

3) No Modified Version of the Font Software may use the Reserved Font
Name(s) unless explicit written permission is granted by the corresponding
Copyright Holder. This restriction only applies to the primary font name as
presented to the users.

4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font
Software shall not be used to promote, endorse or advertise any
Modified Version, except to acknowledge the contribution(s) of the
Copyright Holder(s) and the Author(s) or with their explicit written
permission.

5) The Font Software, modified or unmodified, in part or in whole,
must be distributed entirely under this license, and must not be
distributed under any other license. The requirement for fonts to
remain under this license does not apply to any document created
using the Font Software.

TERMINATION
This license becomes null and void if any of the above conditions are
not met.

DISCLAIMER
THE FONT SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT
OF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE
COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
INCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL
DAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM
OTHER DEALINGS IN THE FONT SOFTWARE.
```

### Hugeicons (@hugeicons/core-free-icons)

| Field | Value |
|---|---|
| License | MIT |
| Copyright | Copyright (c) 2025 Hugeicons |
| Usage scope | Icon data for the pixel-parity exceptions listed in docs/design/icon-policy.md (#159–#164, #170, #180; @hugeicons/core-free-icons 4.4.0); consumed as a runtime dependency, not vendored. |
| Source | https://github.com/hugeicons/hugeicons |
| Pinned reference | https://github.com/hugeicons/hugeicons |
| License URL | https://github.com/hugeicons/hugeicons/blob/main/LICENSE |
| License copy | [docs/upstream/licenses/hugeicons-MIT.txt](docs/upstream/licenses/hugeicons-MIT.txt) |

```text
MIT License

Copyright (c) 2025 Hugeicons

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Dependency-only notices

The following projects are installed as dependencies rather than vendored into this repository.
Their license texts are reproduced above and stored in [docs/upstream/licenses/](docs/upstream/licenses/):

- **reka-ui** 2.11.x (MIT) for the Vue primitive layer planned in #157.
- **@lucide/vue** 1.76.x (Lucide, ISC) for the Office icon adapter planned in #156.
- **@hugeicons/core-free-icons** 4.4.0 (Hugeicons, MIT) for the pixel-parity viewer chrome and Extend
  components listed in [docs/design/icon-policy.md](docs/design/icon-policy.md) (#159–#164, #170, #180).
- **@fontsource-variable/geist** 5.3.x (SIL OFL 1.1) or a host-provided Geist installation for the
  default sans and mono tokens planned in #156.

When a dependency is added or upgraded, update the version note in this file, re-run
`node scripts/ci/verify-attribution.mjs`, and keep the license copy in
[docs/upstream/licenses/](docs/upstream/licenses/) unchanged except for an upstream license
revision.
