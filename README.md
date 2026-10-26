# PDF to Picture

Converts PDF diagrams (Miro, Kumu, PowerPoint and so on) into picture files that Word can insert with **Insert > Pictures**, at print quality.

## Using it

1. Drag PDFs into the window (or onto the .exe icon), or click **Add PDFs…**.
2. Choose **PNG** (works everywhere), **SVG** (vector: best for Kumu-style diagrams in Microsoft 365 Word), or both.
3. Click **Convert**. Pictures are saved next to each PDF unless you choose a folder.

The log reports for each file:

- the pixel size, and how sharp it will print at the given text width;
- for PDFs that are really a picture (e.g. Miro exports), that the original pixels were kept rather than enlarged;
- for vector PDFs, how big the smaller text will print when the figure fills the text width, with a warning below 8pt.

"Trim blank space" crops away empty margins so the diagram fills more of the page.

## Building the .exe

On Windows with Python 3.9+ installed, double-click **build.bat**. The result is `dist\PDF to Picture.exe`, a single file (about 40–50 MB) that runs without Python.

Windows SmartScreen may warn the first time an unsigned .exe is run: **More info > Run anyway**.

## Notes

- Uses PyMuPDF (AGPL licence), which is fine for free internal use and distribution with source.
- Password-protected PDFs are skipped.
- Multi-page PDFs produce one file per page.
