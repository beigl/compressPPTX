# compressPPTX

`compressPPTX` is a small Python utility for reducing the size of PowerPoint `.pptx` files by recompressing embedded media and removing unused package parts.

Repository: https://github.com/beigl/compressPPTX

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

## Windows EXE

You can build a standalone Windows executable with:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

The build output is:

```text
dist/compressPPTX.exe
```

Notes:

- For JPEG/PNG/GIF-only presentations, the EXE works by itself.
- For WAV/MP4 recompression, `ffmpeg.exe` must be available in `PATH` or placed next to `compressPPTX.exe`.
- The repository can be shared as a GitHub link. For non-technical users, the better download path is a GitHub Release or Actions artifact containing the EXE.

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

For the Windows EXE:

```powershell
.\dist\compressPPTX.exe --input-dir input --output-dir output
```

## Tests

Run the basic regression tests with:

```bash
python -m unittest discover -s tests
```

GitHub Actions runs the same test suite automatically on pushes and pull requests.

## Downloading From GitHub

If the repository is public, you can already share this link:

- https://github.com/beigl/compressPPTX

Others can then:

- download the source as ZIP from GitHub
- or, if you publish a Release, download the ready-made `compressPPTX.exe`

Without a Release or uploaded artifact, users do not automatically get a downloadable EXE from the repository homepage alone.

## Development Notes

- The repository is intentionally source-only; no sample presentations are tracked.
- `ffmpeg` is required for audio and video recompression paths.
- The current tests focus on deterministic helper functions and packaging integrity.

## Repository Notes

This repository intentionally excludes presentation files, generated outputs, private credentials, and local tooling files from version control.

## Author and License

Author: Michael Beigl
License: CC-BY 4.0
