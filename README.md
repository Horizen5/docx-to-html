# docx-to-html

Convert `.docx` (Office Open XML) documents into **standalone, offline HTML**,
preserving vector shapes, fills, thin divider lines, rotations/flips and real
tab-stop positions — **no AI or cloud dependency**.

## Usage

```bash
python docx2html.py input.docx -o output.html
```

Pure Python 3, standard library only.

## Features

- Offline conversion (no network / AI needed)
- Preserves preset & custom geometry shapes with rotation and flip
- Preserves line widths / colors and thin horizontal divider lines
- Preserves real tab-stop positions from the document
