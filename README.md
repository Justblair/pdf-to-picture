# PDF to Picture

A small Windows app that turns PDF diagrams into picture files Word can insert at print quality, and tells you whether the text will actually be readable on the page.

It's for diagrams exported from tools such as **Miro**, **Kumu**, **Lucidchart**, **draw.io** and **PowerPoint**, the kind people end up putting into reports, dissertations and theses.

![Platform](https://img.shields.io/badge/platform-Windows-blue) ![Licence](https://img.shields.io/badge/licence-AGPL--3.0-green)

---

## The problem it solves

Word can't insert a PDF as a picture. The usual workaround, copying the diagram out of a PDF viewer and pasting it into Word, quietly turns it into a **screen-resolution image**. It looks fine on screen, but the text goes fuzzy when you zoom in and prints badly.

There's often a second problem hiding behind the first. Online whiteboard tools let you build a diagram on an enormous canvas, so when the whole thing is shrunk to fit a 16 cm wide page, the text can end up **2–4 points high**. That's unreadable no matter how sharp the image is.

PDF to Picture fixes the first problem and measures the second, so you know what you're dealing with.

## Download

Get **PDF-to-Picture.exe** from the [latest release](../../releases/latest). It's a single file. There's nothing to install, and Python isn't needed.

> **Windows SmartScreen** may warn you the first time, because the app isn't code-signed. Click **More info › Run anyway**.

## Using it

1. **Add PDFs.** Drag them into the window, drag them onto the .exe icon, or click **Add PDFs…**.
2. **Choose a format:**
   - **PNG** works everywhere.
   - **SVG** is vector, so it stays perfectly sharp at any size and when cropped. It's the best choice for vector diagrams in Microsoft 365 Word.
   - **Both** gives you both.
3. **Choose options.** Pick the PNG resolution (300 dpi is print quality), whether to trim blank margins, the text width of your page, and where to save.
4. Click **Convert**, then **Open output folder**.

### Inserting into Word

- Use **Insert › Pictures › This Device**. Don't copy and paste, which brings back the original problem.
- Stop Word downsampling your pictures: **File › Options › Advanced › Image Size and Quality**. Tick **Do not compress images in file** and set **Default resolution** to **High fidelity**.
- For a large diagram, insert the SVG (or the high-resolution PNG) and use **Picture Format › Crop** to show one section per figure. Insert it again for each section.
- Add captions with **right-click › Insert Caption**, so figures number themselves and appear in **References › Insert Table of Figures**.

## Reading the log

For every file the app reports:

| Message | Meaning |
|---|---|
| `2779 × 2694 px – about 441 ppi at 16 cm wide (sharp in print)` | The pixel size, and how sharp it will be when the figure fills your text width. 250+ ppi is sharp, 150–250 is acceptable, and under 150 may look soft. |
| `Enough detail to crop in to about 1/5 of the width…` | You can crop in that far in Word and still print sharply. |
| `This PDF holds a picture, so its original pixels were kept` | Some tools (Miro, for example) export a PDF that's really a picture inside a PDF. The app extracts it at its original resolution rather than enlarging it, which would only make the file bigger. No SVG is made for these, because it wouldn't be any sharper. |
| `⚠ Smaller text prints at about 2.4pt at 16 cm wide (aim for 8pt+)` | For vector PDFs the app measures the text. Below about 8pt, readers will struggle, so crop to a section, use a landscape page, or increase the font size in the original tool. |
| `Text size can't be measured in a picture` | For picture PDFs, check readability yourself at 100% zoom in Word. |

## Options

| Option | Default | Notes |
|---|---|---|
| Save as | PNG | PNG, SVG or both |
| PNG resolution | 300 dpi | 150 for drafts, 600 for very fine detail. Ignored for picture PDFs, which keep their native resolution. |
| Trim blank space | On | Crops empty margins so the diagram fills more of the page |
| Text width of page | 16 cm | Used for the sharpness and text-size checks. A4 with 2.5 cm margins is 16 cm; landscape is about 24.7 cm. |
| Save to | Same folder as each PDF | Or choose one folder for everything |

Existing files are never overwritten: a new copy gets `(1)`, `(2)` and so on. Multi-page PDFs produce one file per page. Password-protected PDFs are skipped.

## How it works

- **Vector PDFs** are rendered with [PyMuPDF](https://pymupdf.readthedocs.io/) at the chosen DPI, or converted to SVG with text turned into shapes, so fonts don't need to be installed on the reader's machine.
- **Picture PDFs** are detected (no text, almost no vector drawing, one or more embedded images), then rendered at the embedded image's own resolution, so no detail is lost or invented.
- **Trimming** renders a low-resolution preview, finds everything that differs from the background colour, and crops to it with a small margin.
- **The readability check** takes the smaller text on the page (the 10th percentile by character) and scales it by how much the figure must shrink to fit your text width.
- Very large outputs are capped at 16,000 px on the longest side to keep files manageable in Word.

## Building from source

You'll need Windows and Python 3.9 or later.

```bat
git clone https://github.com/Justblair/pdf-to-picture.git
cd pdf-to-picture
build.bat
```

`build.bat` creates a private virtual environment in `.venv`, installs the requirements and builds `dist\PDF to Picture.exe` with PyInstaller.

To run it without building:

```bat
py -m pip install pymupdf pillow tkinterdnd2
py pdf_to_picture.py
```

The conversion logic in `convert_pdf()` has no GUI code, so it can also be used from other scripts:

```python
from pdf_to_picture import convert_pdf
convert_pdf("diagram.pdf", "out", fmt="both", dpi=300, trim=True, width_cm=16)
```

## Licence

[GNU AGPL-3.0](LICENSE). The app bundles PyMuPDF, which is licensed under the AGPL.
