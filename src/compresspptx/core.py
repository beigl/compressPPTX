#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# -----------------------------------------------------------------------------
# File:        core.py
# Project:     compressPPTX
# Purpose:     Compress PPTX files by shrinking embedded media and cleaning up
#              unused or unnecessary package parts.
# Author:      Michael Beigl
# Copyright:   CC-BY 4.0
# License:     CC-BY 4.0
# Created:     2026-03-30
# Python:      3.10+
# Usage:       python -m compresspptx --input-dir input --output-dir output
# Notes:       Intended for presentation copies where smaller file size matters
#              more than perfect media fidelity.
# -----------------------------------------------------------------------------

from __future__ import annotations

import argparse
import os
import subprocess
import tempfile
import zipfile
from collections import deque
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from PIL import Image, ImageOps, ImageSequence

__author__ = "Michael Beigl"
__copyright__ = "CC-BY 4.0"
__license__ = "CC-BY 4.0"

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DC_NS = "http://purl.org/dc/elements/1.1/"
CPROP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_TERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"

ET.register_namespace("", CP_NS)
ET.register_namespace("cp", CPROP_NS)
ET.register_namespace("dc", DC_NS)
ET.register_namespace("dcterms", DC_TERMS_NS)
ET.register_namespace("xsi", XSI_NS)

CONTENT_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".mp4": "video/mp4",
    ".mp3": "audio/mpeg",
    ".m4a": "audio/mp4",
    ".wav": "audio/wav",
}

PROFILE_LIMITS = {
    "aggressive": (1280, 720, 854),
    "balanced": (1600, 900, 960),
    "light": (1920, 1080, 1280),
}


@dataclass(slots=True)
class ProcessReport:
    source: Path
    destination: Path
    source_size: int
    dest_size: int
    ratio: float
    media_notes: list[str]
    removed_parts: list[str]


def quote_path(path: Path) -> str:
    """Return a plain string path for subprocess arguments."""
    return str(path)


def run_ffmpeg(args: list[str]) -> None:
    """Run ffmpeg in quiet mode and surface failures as exceptions."""
    subprocess.run(
        ["ffmpeg", "-y", "-loglevel", "error", *args],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def has_alpha(image: Image.Image) -> bool:
    """Detect whether an image effectively uses transparency."""
    if image.mode in {"RGBA", "LA"}:
        return True
    if image.mode == "P":
        return "transparency" in image.info
    return False


def resize_fit(size: tuple[int, int], max_long_edge: int) -> tuple[int, int]:
    """Resize while preserving aspect ratio, bounded by the longest edge."""
    width, height = size
    long_edge = max(width, height)
    if long_edge <= max_long_edge:
        return width, height

    scale = max_long_edge / long_edge
    return max(1, int(width * scale)), max(1, int(height * scale))


def save_as_jpeg(src: Path, dest: Path, max_long_edge: int, quality: int) -> None:
    """Store a resized image as progressive JPEG."""
    with Image.open(src) as image:
        image = ImageOps.exif_transpose(image)
        image.load()

        target_size = resize_fit(image.size, max_long_edge)
        if target_size != image.size:
            image = image.resize(target_size, Image.Resampling.LANCZOS)

        image = image.convert("RGB")
        image.save(dest, format="JPEG", quality=quality, optimize=True, progressive=True)


def save_as_png(src: Path, dest: Path, max_long_edge: int) -> None:
    """Store a resized image as optimized PNG with palette reduction."""
    with Image.open(src) as image:
        image = ImageOps.exif_transpose(image)
        image.load()

        target_size = resize_fit(image.size, max_long_edge)
        if target_size != image.size:
            image = image.resize(target_size, Image.Resampling.LANCZOS)

        if image.mode not in {"RGBA", "LA", "P"}:
            image = image.convert("RGBA" if has_alpha(image) else "RGB")

        if image.mode == "RGBA":
            image = image.quantize(colors=64, method=Image.Quantize.FASTOCTREE)
        elif image.mode == "RGB":
            image = image.quantize(colors=128, method=Image.Quantize.MEDIANCUT)

        image.save(dest, format="PNG", optimize=True, compress_level=9)


def save_as_gif(src: Path, dest: Path, max_long_edge: int) -> None:
    """Re-encode an animated GIF with smaller frames and adaptive palette."""
    with Image.open(src) as image:
        frames: list[Image.Image] = []
        durations: list[int] = []
        disposal: list[int] = []

        for frame in ImageSequence.Iterator(image):
            current = frame.convert("RGBA")
            target_size = resize_fit(current.size, max_long_edge)
            if target_size != current.size:
                current = current.resize(target_size, Image.Resampling.LANCZOS)

            current = current.convert("P", palette=Image.Palette.ADAPTIVE, colors=96)
            frames.append(current)
            durations.append(frame.info.get("duration", image.info.get("duration", 100)))
            disposal.append(frame.info.get("disposal", 2))

        if not frames:
            raise RuntimeError(f"No frames found in {src}")

        frames[0].save(
            dest,
            format="GIF",
            save_all=True,
            append_images=frames[1:],
            optimize=True,
            loop=image.info.get("loop", 0),
            duration=durations,
            disposal=disposal,
        )


def should_convert_png_to_jpeg(image: Image.Image) -> bool:
    """Use JPEG only for non-transparent, photo-like images."""
    if has_alpha(image):
        return False

    colors = image.getcolors(maxcolors=512)
    if colors is not None and len(colors) <= 192:
        return False

    return True


def relative_target(from_part: PurePosixPath, to_part: PurePosixPath) -> str:
    """Build a relationship target relative to a package part."""
    return os.path.relpath(str(to_part), start=str(from_part.parent)).replace("\\", "/")


def part_from_rel_path(rel_path: Path, root: Path) -> PurePosixPath | None:
    """Resolve a .rels file back to the source package part it belongs to."""
    rel_posix = rel_path.relative_to(root).as_posix()
    if rel_posix == "_rels/.rels":
        return None

    parent = PurePosixPath(rel_posix)
    if parent.parent.name != "_rels" or not parent.name.endswith(".rels"):
        raise ValueError(f"Unexpected relationships path: {rel_posix}")

    source_dir = parent.parent.parent
    source_name = parent.name[:-5]
    return source_dir / source_name


def update_relationship_targets(root: Path, rename_map: dict[str, str]) -> None:
    """Rewrite relationship targets after media files were renamed."""
    for rel_file in root.rglob("*.rels"):
        tree = ET.parse(rel_file)
        rel_root = tree.getroot()
        source_part = part_from_rel_path(rel_file, root)
        changed = False

        for rel in rel_root.findall(f"{{{REL_NS}}}Relationship"):
            target = rel.get("Target")
            mode = rel.get("TargetMode")
            if not target or mode == "External":
                continue

            if source_part is None:
                current = PurePosixPath(target.lstrip("/"))
            else:
                resolved = os.path.normpath(str(source_part.parent / target)).replace("\\", "/")
                current = PurePosixPath(resolved)

            current_key = current.as_posix()
            if current_key not in rename_map:
                continue

            renamed = PurePosixPath(rename_map[current_key])
            if source_part is None:
                rel.set("Target", "/" + renamed.as_posix())
            else:
                rel.set("Target", relative_target(source_part, renamed))
            changed = True

        if changed:
            tree.write(rel_file, encoding="utf-8", xml_declaration=True)


def update_content_types(root: Path, rename_map: dict[str, str]) -> None:
    """Update content type entries after media extensions changed."""
    content_types_path = root / "[Content_Types].xml"
    tree = ET.parse(content_types_path)
    xml_root = tree.getroot()

    overrides = xml_root.findall(f"{{{CP_NS}}}Override")
    existing_overrides = {override.get("PartName"): override for override in overrides}
    defaults = {
        default.get("Extension"): default
        for default in xml_root.findall(f"{{{CP_NS}}}Default")
    }

    for old_name, new_name in rename_map.items():
        old_part = "/" + old_name
        new_part = "/" + new_name
        new_suffix = PurePosixPath(new_name).suffix.lower()
        old_override = existing_overrides.get(old_part)

        if old_override is not None:
            content_type = CONTENT_TYPES.get(new_suffix, old_override.get("ContentType"))
            old_override.set("PartName", new_part)
            if content_type:
                old_override.set("ContentType", content_type)
            existing_overrides[new_part] = old_override
            existing_overrides.pop(old_part, None)
            continue

        new_ext = new_suffix.lstrip(".")
        if new_ext and new_ext not in defaults and new_suffix in CONTENT_TYPES:
            default = ET.SubElement(xml_root, f"{{{CP_NS}}}Default")
            default.set("Extension", new_ext)
            default.set("ContentType", CONTENT_TYPES[new_suffix])
            defaults[new_ext] = default

    tree.write(content_types_path, encoding="utf-8", xml_declaration=True)


def strip_metadata(root: Path) -> None:
    """Remove descriptive metadata from the PowerPoint package."""
    core_xml = root / "docProps" / "core.xml"
    if core_xml.exists():
        tree = ET.parse(core_xml)
        core_root = tree.getroot()

        for tag in [
            f"{{{DC_NS}}}creator",
            f"{{{CPROP_NS}}}lastModifiedBy",
            f"{{{CPROP_NS}}}keywords",
            f"{{{DC_NS}}}description",
            f"{{{CPROP_NS}}}category",
            f"{{{DC_NS}}}title",
            f"{{{CPROP_NS}}}subject",
        ]:
            for node in core_root.findall(tag):
                node.text = ""

        tree.write(core_xml, encoding="utf-8", xml_declaration=True)

    app_xml = root / "docProps" / "app.xml"
    if app_xml.exists():
        tree = ET.parse(app_xml)
        app_root = tree.getroot()

        for node in app_root:
            if node.tag.endswith(("Company", "Manager")):
                node.text = ""

        tree.write(app_xml, encoding="utf-8", xml_declaration=True)


def remove_custom_properties(root: Path) -> None:
    """Remove custom document properties and dangling references to them."""
    custom_xml = root / "docProps" / "custom.xml"
    if custom_xml.exists():
        custom_xml.unlink()

    rels_path = root / "docProps" / "_rels" / "app.xml.rels"
    if rels_path.exists():
        tree = ET.parse(rels_path)
        rel_root = tree.getroot()
        changed = False

        for rel in list(rel_root.findall(f"{{{REL_NS}}}Relationship")):
            if rel.get("Target") == "custom.xml":
                rel_root.remove(rel)
                changed = True

        if changed:
            tree.write(rels_path, encoding="utf-8", xml_declaration=True)

    root_rels = root / "_rels" / ".rels"
    if root_rels.exists():
        tree = ET.parse(root_rels)
        rel_root = tree.getroot()
        changed = False

        for rel in list(rel_root.findall(f"{{{REL_NS}}}Relationship")):
            if rel.get("Target") == "docProps/custom.xml":
                rel_root.remove(rel)
                changed = True

        if changed:
            tree.write(root_rels, encoding="utf-8", xml_declaration=True)

    content_types_path = root / "[Content_Types].xml"
    if content_types_path.exists():
        tree = ET.parse(content_types_path)
        xml_root = tree.getroot()
        changed = False

        for override in list(xml_root.findall(f"{{{CP_NS}}}Override")):
            if override.get("PartName") == "/docProps/custom.xml":
                xml_root.remove(override)
                changed = True

        if changed:
            tree.write(content_types_path, encoding="utf-8", xml_declaration=True)


def collect_reachable_parts(root: Path) -> set[str]:
    """Traverse package relationships and collect every reachable part."""
    reachable: set[str] = set()
    queue: deque[tuple[Path, PurePosixPath | None]] = deque()

    root_rels = root / "_rels" / ".rels"
    if root_rels.exists():
        queue.append((root_rels, None))

    while queue:
        rel_file, source_part = queue.popleft()
        tree = ET.parse(rel_file)
        rel_root = tree.getroot()

        for rel in rel_root.findall(f"{{{REL_NS}}}Relationship"):
            target = rel.get("Target")
            if not target or rel.get("TargetMode") == "External":
                continue

            if source_part is None:
                part = PurePosixPath(target.lstrip("/"))
            else:
                resolved = os.path.normpath(str(source_part.parent / target)).replace("\\", "/")
                part = PurePosixPath(resolved)

            part_name = part.as_posix()
            if part_name in reachable:
                continue

            reachable.add(part_name)
            part_path = root / Path(part_name)
            rel_candidate = part_path.parent / "_rels" / f"{part_path.name}.rels"
            if rel_candidate.exists():
                queue.append((rel_candidate, part))

    return reachable


def prune_unreachable(root: Path) -> list[str]:
    """Delete package files that are no longer referenced anywhere."""
    reachable = collect_reachable_parts(root)
    removed: list[str] = []

    for file_path in sorted(path for path in root.rglob("*") if path.is_file()):
        rel_path = file_path.relative_to(root).as_posix()

        if rel_path in {"[Content_Types].xml", "_rels/.rels"}:
            continue

        if rel_path.endswith(".rels"):
            source_part = part_from_rel_path(file_path, root)
            if source_part is None:
                continue
            if source_part.as_posix() not in reachable:
                removed.append(rel_path)
                file_path.unlink()
            continue

        if rel_path not in reachable:
            removed.append(rel_path)
            file_path.unlink()

    for directory in sorted((path for path in root.rglob("*") if path.is_dir()), reverse=True):
        if not any(directory.iterdir()):
            directory.rmdir()

    return removed


def maybe_replace_original(temp_file: Path, current_file: Path, original_size: int) -> Path:
    """Keep the converted file only if it is actually smaller."""
    if temp_file.stat().st_size < original_size:
        current_file.unlink()
        temp_file.replace(current_file)
        return current_file

    temp_file.unlink()
    return current_file


def convert_jpeg_file(file_path: Path, max_image_edge: int, original_size: int) -> Path:
    """Re-encode JPEG media in place."""
    temp_path = file_path.with_suffix(".tmp.jpg")
    save_as_jpeg(file_path, temp_path, max_long_edge=max_image_edge, quality=66)

    target_path = file_path if file_path.suffix.lower() == ".jpg" else file_path.with_suffix(".jpg")
    if temp_path.stat().st_size >= original_size:
        temp_path.unlink()
        return file_path

    file_path.unlink()
    temp_path.replace(target_path)
    return target_path


def convert_bitmap_file(file_path: Path, max_image_edge: int, original_size: int) -> Path:
    """Convert bitmap-like formats either to JPEG or optimized PNG."""
    with Image.open(file_path) as image:
        convert_to_jpeg = should_convert_png_to_jpeg(image)

    if convert_to_jpeg:
        target_path = file_path.with_suffix(".jpg")
        temp_path = file_path.with_suffix(".tmp.jpg")
        save_as_jpeg(file_path, temp_path, max_long_edge=max_image_edge, quality=62)
    else:
        target_path = file_path.with_suffix(".png")
        temp_path = file_path.with_suffix(".tmp.png")
        save_as_png(file_path, temp_path, max_long_edge=max_image_edge)

    if temp_path.stat().st_size >= original_size:
        temp_path.unlink()
        return file_path

    file_path.unlink()
    temp_path.replace(target_path)
    return target_path


def convert_gif_file(file_path: Path, max_gif_edge: int, original_size: int) -> Path:
    """Re-encode GIF animations in place."""
    temp_path = file_path.with_suffix(".tmp.gif")
    save_as_gif(file_path, temp_path, max_long_edge=max_gif_edge)
    return maybe_replace_original(temp_path, file_path, original_size)


def convert_mp4_file(file_path: Path, video_width: int, original_size: int) -> Path:
    """Re-encode MP4 videos using a smaller bitrate and width cap."""
    temp_path = file_path.with_suffix(".tmp.mp4")
    run_ffmpeg(
        [
            "-i",
            quote_path(file_path),
            "-vf",
            f"scale='min({video_width},iw)':-2",
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-crf",
            "36",
            "-movflags",
            "+faststart",
            "-c:a",
            "aac",
            "-b:a",
            "48k",
            quote_path(temp_path),
        ]
    )
    return maybe_replace_original(temp_path, file_path, original_size)


def convert_wav_file(file_path: Path, original_size: int) -> Path:
    """Convert WAV audio to MP3 when that reduces file size."""
    target_path = file_path.with_suffix(".mp3")
    temp_path = file_path.with_suffix(".tmp.mp3")
    run_ffmpeg(
        [
            "-i",
            quote_path(file_path),
            "-c:a",
            "libmp3lame",
            "-b:a",
            "64k",
            quote_path(temp_path),
        ]
    )

    if temp_path.stat().st_size >= original_size:
        temp_path.unlink()
        return file_path

    file_path.unlink()
    temp_path.replace(target_path)
    return target_path


def convert_media(
    root: Path,
    max_image_edge: int,
    max_gif_edge: int,
    video_width: int,
) -> tuple[dict[str, str], list[str]]:
    """Compress media assets inside ppt/media and track renamed parts."""
    media_dir = root / "ppt" / "media"
    rename_map: dict[str, str] = {}
    notes: list[str] = []

    if not media_dir.exists():
        return rename_map, notes

    for file_path in sorted(media_dir.iterdir()):
        if not file_path.is_file():
            continue

        suffix = file_path.suffix.lower()
        rel_name = file_path.relative_to(root).as_posix()
        original_size = file_path.stat().st_size
        new_path = file_path

        try:
            if suffix in {".jpg", ".jpeg"}:
                new_path = convert_jpeg_file(file_path, max_image_edge, original_size)
            elif suffix in {".png", ".bmp", ".tif", ".tiff", ".emf", ".wmf"}:
                new_path = convert_bitmap_file(file_path, max_image_edge, original_size)
            elif suffix == ".gif":
                new_path = convert_gif_file(file_path, max_gif_edge, original_size)
            elif suffix == ".mp4":
                new_path = convert_mp4_file(file_path, video_width, original_size)
            elif suffix == ".wav":
                new_path = convert_wav_file(file_path, original_size)
        except Exception as exc:
            notes.append(f"{rel_name}: skipped ({exc})")
            continue

        new_rel_name = new_path.relative_to(root).as_posix()
        if new_rel_name != rel_name:
            rename_map[rel_name] = new_rel_name

        final_size = new_path.stat().st_size
        if final_size < original_size:
            notes.append(
                f"{rel_name}: {original_size / 1_048_576:.1f} MB -> "
                f"{final_size / 1_048_576:.1f} MB"
            )

    return rename_map, notes


def zip_directory(source_dir: Path, destination: Path) -> None:
    """Rebuild the unpacked PPTX directory as a zip archive."""
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for file_path in sorted(path for path in source_dir.rglob("*") if path.is_file()):
            archive.write(file_path, arcname=file_path.relative_to(source_dir).as_posix())


def process_pptx(
    src: Path,
    dest: Path,
    max_image_edge: int,
    max_gif_edge: int,
    video_width: int,
) -> ProcessReport:
    """Process one PPTX file end-to-end."""
    with tempfile.TemporaryDirectory(prefix="pptx-compress-") as temp_dir:
        temp_root = Path(temp_dir)

        with zipfile.ZipFile(src) as archive:
            archive.extractall(temp_root)

        rename_map, notes = convert_media(
            temp_root,
            max_image_edge=max_image_edge,
            max_gif_edge=max_gif_edge,
            video_width=video_width,
        )

        if rename_map:
            update_relationship_targets(temp_root, rename_map)
            update_content_types(temp_root, rename_map)

        remove_custom_properties(temp_root)
        strip_metadata(temp_root)
        removed_parts = prune_unreachable(temp_root)
        zip_directory(temp_root, dest)

    dest_size = dest.stat().st_size
    source_size = src.stat().st_size
    ratio = dest_size / source_size if source_size else 0.0

    return ProcessReport(
        source=src,
        destination=dest,
        source_size=source_size,
        dest_size=dest_size,
        ratio=ratio,
        media_notes=notes,
        removed_parts=removed_parts,
    )


def apply_profile(profile: str) -> tuple[int, int, int]:
    """Resolve profile presets for images, GIFs, and videos."""
    try:
        return PROFILE_LIMITS[profile]
    except KeyError as exc:
        raise ValueError(f"Unknown profile: {profile}") from exc


def parse_args() -> argparse.Namespace:
    """Parse command-line options."""
    parser = argparse.ArgumentParser(
        description="Compress PPTX files for projector playback by shrinking embedded media."
    )
    parser.add_argument("--input-dir", type=Path, default=Path("input"))
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument(
        "--compression-strength",
        choices=sorted(PROFILE_LIMITS),
        default="aggressive",
        help="Preset compression profile. 'aggressive' targets much smaller files.",
    )
    parser.add_argument("--max-image-edge", type=int)
    parser.add_argument("--max-gif-edge", type=int)
    parser.add_argument("--video-width", type=int)
    return parser.parse_args()


def fill_profile_defaults(args: argparse.Namespace) -> None:
    """Populate omitted CLI values from the selected profile."""
    profile_image_edge, profile_gif_edge, profile_video_width = apply_profile(args.compression_strength)

    if args.max_image_edge is None:
        args.max_image_edge = profile_image_edge
    if args.max_gif_edge is None:
        args.max_gif_edge = profile_gif_edge
    if args.video_width is None:
        args.video_width = profile_video_width


def collect_input_files(input_dir: Path) -> list[Path]:
    """Collect PPTX files from the configured input directory."""
    if not input_dir.exists():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    files = sorted(path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() == ".pptx")
    if not files:
        raise SystemExit(f"No PPTX files found in {input_dir}")

    return files


def print_report(report: ProcessReport) -> None:
    """Print a compact summary for one processed PPTX file."""
    print(
        f"{report.source.name}: "
        f"{report.source_size / 1_048_576:.1f} MB -> "
        f"{report.dest_size / 1_048_576:.1f} MB "
        f"({report.ratio:.1%})"
    )

    if report.removed_parts:
        print(f"  Removed {len(report.removed_parts)} unreachable/custom parts")

    if report.media_notes:
        print("  Media changes:")
        for note in report.media_notes[:12]:
            print(f"    - {note}")
        if len(report.media_notes) > 12:
            print(f"    - ... {len(report.media_notes) - 12} more")


def main() -> None:
    """CLI entry point."""
    Image.MAX_IMAGE_PIXELS = None

    args = parse_args()
    fill_profile_defaults(args)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    input_files = collect_input_files(args.input_dir)

    reports: list[ProcessReport] = []
    for file_path in input_files:
        destination = args.output_dir / file_path.name
        reports.append(
            process_pptx(
                file_path,
                destination,
                args.max_image_edge,
                args.max_gif_edge,
                args.video_width,
            )
        )

    for report in reports:
        print_report(report)
