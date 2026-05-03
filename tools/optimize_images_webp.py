#!/usr/bin/env python3
"""
Scan and convert oversized site images to WebP.

Default behavior is safe: candidates are converted into a separate
`image_optimized/` directory, preserving relative paths and filenames.

Examples:
  python tools/optimize_images_webp.py
  python tools/optimize_images_webp.py --src image --out image_optimized
  python tools/optimize_images_webp.py --dry-run
  python tools/optimize_images_webp.py --overwrite --out image
"""

from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

try:
    from PIL import Image, ImageOps
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Pillow is required. Install it with: python -m pip install pillow"
    ) from exc


IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".webp",
    ".avif",
    ".gif",
}


@dataclass
class Candidate:
    source: Path
    width: int
    height: int
    size_bytes: int
    reason: str


def iter_images(root: Path) -> Iterable[Path]:
    for path in sorted(root.rglob("*")):
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS:
            yield path


def inspect_image(path: Path) -> tuple[int, int] | None:
    try:
        with Image.open(path) as image:
            return image.size
    except Exception:
        return None


def find_candidates(root: Path, threshold_bytes: int, max_width: int) -> list[Candidate]:
    candidates: list[Candidate] = []
    for path in iter_images(root):
        size = path.stat().st_size
        dimensions = inspect_image(path)
        if not dimensions:
            continue

        width, height = dimensions
        reasons = []
        if size > threshold_bytes:
            reasons.append(f">{threshold_bytes // 1024}KB")
        if width > max_width:
            reasons.append(f">{max_width}px wide")
        if reasons:
            candidates.append(
                Candidate(
                    source=path,
                    width=width,
                    height=height,
                    size_bytes=size,
                    reason=", ".join(reasons),
                )
            )
    return candidates


def destination_for(source_root: Path, output_root: Path, source: Path) -> Path:
    rel = source.relative_to(source_root)
    return (output_root / rel).with_suffix(".webp")


def convert_to_webp(
    candidate: Candidate,
    source_root: Path,
    output_root: Path,
    max_width: int,
    quality: int,
    overwrite: bool,
) -> tuple[Path, int, int, int]:
    destination = destination_for(source_root, output_root, candidate.source)
    if destination.exists() and not overwrite:
        with Image.open(destination) as existing:
            return destination, existing.width, existing.height, destination.stat().st_size

    with Image.open(candidate.source) as image:
        image = ImageOps.exif_transpose(image)
        if image.mode not in ("RGB", "RGBA"):
            image = image.convert("RGBA" if "A" in image.getbands() else "RGB")

        if image.mode == "RGBA":
            background = Image.new("RGB", image.size, "#ffffff")
            background.paste(image, mask=image.getchannel("A"))
            image = background

        if image.width > max_width:
            new_height = round(image.height * max_width / image.width)
            image = image.resize((max_width, new_height), Image.Resampling.LANCZOS)

        destination.parent.mkdir(parents=True, exist_ok=True)
        image.save(destination, "WEBP", quality=quality, method=6)
        return destination, image.width, image.height, destination.stat().st_size


def write_report(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "source",
                "source_kb",
                "source_width",
                "source_height",
                "reason",
                "output",
                "output_kb",
                "output_width",
                "output_height",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--src", default="image", help="Source image directory.")
    parser.add_argument(
        "--out",
        default="image_optimized",
        help="Output directory. Use --out image with --overwrite to replace WebP outputs in place.",
    )
    parser.add_argument("--max-width", type=int, default=1200)
    parser.add_argument("--threshold-kb", type=int, default=200)
    parser.add_argument("--quality", type=int, default=82)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument(
        "--report",
        default="image_optimization_report.csv",
        help="CSV report path.",
    )
    args = parser.parse_args()

    source_root = Path(args.src).resolve()
    output_root = Path(args.out).resolve()
    threshold_bytes = args.threshold_kb * 1024

    if not source_root.exists():
        raise SystemExit(f"Source directory not found: {source_root}")

    candidates = find_candidates(source_root, threshold_bytes, args.max_width)
    rows: list[dict[str, object]] = []

    print(f"Source: {source_root}")
    print(f"Candidates: {len(candidates)}")

    for candidate in candidates:
        row: dict[str, object] = {
            "source": str(candidate.source),
            "source_kb": round(candidate.size_bytes / 1024, 1),
            "source_width": candidate.width,
            "source_height": candidate.height,
            "reason": candidate.reason,
            "output": "",
            "output_kb": "",
            "output_width": "",
            "output_height": "",
        }

        if args.dry_run:
            print(
                f"DRY {candidate.source} "
                f"{candidate.width}x{candidate.height} "
                f"{candidate.size_bytes / 1024:.1f}KB ({candidate.reason})"
            )
        else:
            destination, out_w, out_h, out_size = convert_to_webp(
                candidate,
                source_root,
                output_root,
                args.max_width,
                args.quality,
                args.overwrite,
            )
            row.update(
                {
                    "output": str(destination),
                    "output_kb": round(out_size / 1024, 1),
                    "output_width": out_w,
                    "output_height": out_h,
                }
            )
            print(
                f"OK {candidate.source.name} -> {destination} "
                f"{out_w}x{out_h} {out_size / 1024:.1f}KB"
            )

        rows.append(row)

    write_report(Path(args.report), rows)
    print(f"Report: {Path(args.report).resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
