# compressPPTX

`compressPPTX` is a small Python utility for reducing the size of PowerPoint `.pptx` files by recompressing embedded media and removing unused package parts.

## Features

- recompresses embedded JPEG, PNG, GIF, WAV, and MP4 media
- updates PPTX relationship and content-type references after media conversion
- removes custom document properties
- strips selected document metadata
- prunes unreachable package parts before rebuilding the archive

## Requirements

- Python 3.10 or newer
- [ffmpeg](https://ffmpeg.org/) available in `PATH`

Install the Python dependency with:

```bash
pip install -r requirements.txt
```

## Usage

Place source presentations in `input/` and run:

```bash
python compress_pptx.py --input-dir input --output-dir output
```

Compression profiles:

- `aggressive`
- `balanced`
- `light`

Example:

```bash
python compress_pptx.py --compression-strength balanced
```

## Repository Notes

This repository intentionally excludes presentation files, generated outputs, and private credentials from version control.

## Author and License

Author: Michael Beigl  
License: CC-BY 4.0
