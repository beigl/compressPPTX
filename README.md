# compressPPTX

`compressPPTX` is a small Python utility for reducing the size of PowerPoint `.pptx` files by recompressing embedded media and removing unused package parts.

## Features

- recompresses embedded JPEG, PNG, GIF, WAV, and MP4 media
- updates PPTX relationship and content-type references after media conversion
- removes custom document properties
- strips selected document metadata
- prunes unreachable package parts before rebuilding the archive

## Project Layout

```text
compressPPTX/
|-- src/compresspptx/   # packaged application code
|-- tests/              # lightweight regression tests
|-- compress_pptx.py    # compatibility wrapper
`-- pyproject.toml      # project metadata
```

## Requirements

- Python 3.10 or newer
- [ffmpeg](https://ffmpeg.org/) available in `PATH`

Install the package locally with:

```bash
pip install -e .
```

This creates the command-line tool:

```bash
compress-pptx
```

## Usage

Place source presentations in `input/` and run one of these commands:

```bash
python compress_pptx.py --input-dir input --output-dir output
python -m compresspptx --input-dir input --output-dir output
compress-pptx --input-dir input --output-dir output
```

Compression profiles:

- `aggressive`
- `balanced`
- `light`

Example:

```bash
python -m compresspptx --compression-strength balanced
```

## Tests

Run the basic regression tests with:

```bash
python -m unittest discover -s tests
```

GitHub Actions runs the same test suite automatically on pushes and pull requests.

## Development Notes

- The repository is intentionally source-only; no sample presentations are tracked.
- `ffmpeg` is required for audio and video recompression paths.
- The current tests focus on deterministic helper functions and packaging integrity.

## Repository Notes

This repository intentionally excludes presentation files, generated outputs, private credentials, and local tooling files from version control.

## Author and License

Author: Michael Beigl
License: CC-BY 4.0
