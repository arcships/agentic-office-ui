#!/usr/bin/env python3
"""Build a visual-audit atlas from a reference/subject screenshot pair.

This script does not compare pixels or declare differences. It prepares the
inputs for a blind, vision-first audit:

  overview/           whole-page A/B overviews
  tiles/<tile-id>/    native-resolution overlapping A/B crops
  coverage.json       one row per tile; reviewers mark reviewedBy/candidates
  manifest.json       hashes, dimensions, tiling parameters, capture metadata

The atlas is intentionally lossless PNG. Do not downscale tiles before a
vision reviewer reads them.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def tile_starts(length: int, size: int, overlap: int) -> list[int]:
    # A tiny remainder would create a nearly duplicate tile. If the axis is
    # only one overlap beyond the tile, use one full-length tile instead.
    if length <= size + overlap:
        return [0]
    step = max(1, size - overlap)
    starts = list(range(0, max(1, length - size + 1), step))
    last = length - size
    if starts[-1] != last:
        starts.append(last)
    return starts


def label_strip(image: Image.Image, label: str, height: int = 24) -> Image.Image:
    from PIL import ImageDraw

    canvas = Image.new("RGB", (image.width, image.height + height), "white")
    canvas.paste(image, (0, height))
    draw = ImageDraw.Draw(canvas)
    draw.text((8, 5), label, fill="black")
    return canvas


def stack_pair(reference: Image.Image, subject: Image.Image, label: bool = True) -> Image.Image:
    left = label_strip(reference, "A") if label else reference
    right = label_strip(subject, "B") if label else subject
    gap = 12
    canvas = Image.new("RGB", (left.width + gap + right.width, max(left.height, right.height)), "#c8c8c8")
    canvas.paste(left, (0, 0))
    canvas.paste(right, (left.width + gap, 0))
    return canvas


def build_tile(
    reference: Image.Image,
    subject: Image.Image,
    box: tuple[int, int, int, int],
    out_dir: Path,
    tile_id: str,
) -> dict:
    target = out_dir / "tiles" / tile_id
    target.mkdir(parents=True, exist_ok=True)
    ref = reference.crop(box)
    subj = subject.crop(box)
    ref_path = target / "A.png"
    subj_path = target / "B.png"
    pair_path = target / "AB.png"
    ref.save(ref_path, optimize=True)
    subj.save(subj_path, optimize=True)
    stack_pair(ref, subj).save(pair_path, optimize=True)
    return {
        "id": tile_id,
        "box": {"x": box[0], "y": box[1], "w": box[2] - box[0], "h": box[3] - box[1]},
        "files": {"A": str(ref_path), "B": str(subj_path), "AB": str(pair_path)},
        "hashes": {"A": sha256(ref_path), "B": sha256(subj_path), "AB": sha256(pair_path)},
        "reviewedBy": {},
        "candidates": [],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--subject", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--route", required=True)
    parser.add_argument("--viewport", required=True)
    parser.add_argument("--theme", required=True, choices=["light", "dark"])
    parser.add_argument("--state", default="default")
    parser.add_argument("--tile-size", type=int, default=512)
    parser.add_argument("--overlap", type=int, default=64)
    parser.add_argument("--locale", default="en")
    parser.add_argument("--zoom", default="100")
    parser.add_argument("--fixture", default=None)
    parser.add_argument("--fixture-sha256", default=None)
    parser.add_argument("--capture-sha", default=None)
    parser.add_argument("--no-pair-labels", action="store_true")
    args = parser.parse_args()

    if args.tile_size < 64:
        raise SystemExit("--tile-size must be >= 64")
    if args.overlap < 0 or args.overlap >= args.tile_size:
        raise SystemExit("--overlap must be in [0, tile-size)")

    reference = Image.open(args.reference).convert("RGB")
    subject = Image.open(args.subject).convert("RGB")
    if reference.size != subject.size:
        raise SystemExit(
            f"size mismatch: reference {reference.size}, subject {subject.size}; "
            "fix the capture before auditing"
        )

    out: Path = args.out
    (out / "overview").mkdir(parents=True, exist_ok=True)
    (out / "tiles").mkdir(parents=True, exist_ok=True)

    reference.save(out / "overview" / "A.png", optimize=True)
    subject.save(out / "overview" / "B.png", optimize=True)
    stack_pair(reference, subject, label=not args.no_pair_labels).save(
        out / "overview" / "AB.png", optimize=True
    )

    width, height = reference.size
    x_starts = tile_starts(width, args.tile_size, args.overlap)
    y_starts = tile_starts(height, args.tile_size, args.overlap)
    tiles = []
    for row, y in enumerate(y_starts):
        for col, x in enumerate(x_starts):
            x_end = width if len(x_starts) == 1 else min(width, x + args.tile_size)
            y_end = height if len(y_starts) == 1 else min(height, y + args.tile_size)
            box = (x, y, x_end, y_end)
            tile_id = f"r{row}c{col}"
            tiles.append(build_tile(reference, subject, box, out, tile_id))

    manifest = {
        "schemaVersion": 1,
        "kind": "visual-audit-atlas",
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "route": args.route,
        "viewport": args.viewport,
        "theme": args.theme,
        "state": args.state,
        "locale": args.locale,
        "zoom": args.zoom,
        "fixture": args.fixture,
        "fixtureSha256": args.fixture_sha256,
        "captureSha": args.capture_sha,
        "reference": {"path": str(args.reference), "sha256": sha256(args.reference), "width": width, "height": height},
        "subject": {"path": str(args.subject), "sha256": sha256(args.subject), "width": width, "height": height},
        "tiling": {
            "tileSize": args.tile_size,
            "overlap": args.overlap,
            "columns": len(x_starts),
            "rows": len(y_starts),
            "xStarts": x_starts,
            "yStarts": y_starts,
        },
        "overview": {
            "A": str(out / "overview" / "A.png"),
            "B": str(out / "overview" / "B.png"),
            "AB": str(out / "overview" / "AB.png"),
        },
        "tiles": tiles,
    }
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")

    coverage = {
        "schemaVersion": 1,
        "kind": "visual-audit-coverage",
        "route": args.route,
        "viewport": args.viewport,
        "theme": args.theme,
        "state": args.state,
        "tiles": [
            {"id": tile["id"], "box": tile["box"], "reviewedBy": {}, "candidates": []}
            for tile in tiles
        ],
        "candidates": [],
        "reviewers": {},
    }
    (out / "coverage.json").write_text(json.dumps(coverage, indent=2) + "\n")
    print(
        f"atlas: {len(tiles)} tiles ({len(x_starts)}x{len(y_starts)}) "
        f"for {args.route} {args.viewport}/{args.theme}/{args.state}"
    )
    print(f"manifest: {out / 'manifest.json'}")
    print(f"coverage: {out / 'coverage.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
