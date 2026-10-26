"""PDF to Picture

Turns PDF diagrams (Miro, Kumu, PowerPoint, etc.) into picture files that
Word can insert with Insert > Pictures, at print quality, and reports whether
the text will be readable once the figure is shrunk to fit the page.

Run:   python pdf_to_picture.py [optional PDF files...]
Build: see build.bat
"""

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

import pymupdf
from PIL import Image

Image.MAX_IMAGE_PIXELS = None

APP_NAME = "PDF to Picture"
MAX_SIDE_PX = 16000          # keep files a sensible size for Word
MIN_READABLE_PT = 8          # smallest comfortable printed text size
TRIM_MARGIN_PT = 12          # breathing space left around trimmed diagrams


# --------------------------------------------------------------------------
# Conversion logic (no GUI code here, so it can be tested on its own)
# --------------------------------------------------------------------------

def unique_path(path: Path) -> Path:
    """Never overwrite: add (1), (2)... if the file already exists."""
    if not path.exists():
        return path
    i = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def content_area(page) -> pymupdf.Rect:
    """The part of the page that isn't plain background, plus a small margin."""
    dpi = 36
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, alpha=False)
    img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    background = img.getpixel((1, 1))
    mask = img.point(lambda v: 255 if abs(v - background) > 12 else 0)
    box = mask.getbbox()
    rect = page.rect
    if not box:
        return rect
    s = 72 / dpi
    area = pymupdf.Rect(
        rect.x0 + box[0] * s - TRIM_MARGIN_PT,
        rect.y0 + box[1] * s - TRIM_MARGIN_PT,
        rect.x0 + box[2] * s + TRIM_MARGIN_PT,
        rect.y0 + box[3] * s + TRIM_MARGIN_PT,
    )
    return area & rect


def picture_only_ppi(page):
    """If the page is just a picture (e.g. Miro exports), return its native
    resolution in pixels per inch, otherwise None."""
    if page.get_text("words"):
        return None
    infos = page.get_image_info()
    if not infos or len(page.get_drawings()) > 20:
        return None
    best = 0.0
    for info in infos:
        bbox = pymupdf.Rect(info["bbox"])
        if bbox.width > 0:
            best = max(best, info["width"] / (bbox.width / 72))
    return best or None


def small_text_size(page, clip):
    """Typical small text size (10th percentile by character) inside clip."""
    sizes = []
    for block in page.get_text("dict", clip=clip)["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                n = len(span["text"].strip())
                if n:
                    sizes.extend([span["size"]] * n)
    if not sizes:
        return None
    sizes.sort()
    return sizes[len(sizes) // 10]


def convert_pdf(pdf_path, out_dir, fmt="png", dpi=300, trim=True,
                width_cm=16.0, log=print):
    """Convert every page of pdf_path. fmt is 'png', 'svg' or 'both'.
    Returns a list of files written."""
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir) if out_dir else pdf_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    column_pt = width_cm / 2.54 * 72

    with pymupdf.open(pdf_path) as doc:
        if doc.needs_pass:
            log(f"  ✗ {pdf_path.name} is password protected - skipped.")
            return written
        multi = doc.page_count > 1
        for page in doc:
            stem = pdf_path.stem + (f" - page {page.number + 1}" if multi else "")
            clip = content_area(page) if trim else page.rect
            shrink = min(1.0, column_pt / clip.width)
            native_ppi = picture_only_ppi(page)
            make_png = fmt in ("png", "both")
            make_svg = fmt in ("svg", "both")
            if native_ppi and make_svg:
                # A picture wrapped in a PDF: SVG adds nothing, so give a PNG.
                make_svg = False
                if not make_png:
                    make_png = True
                    log("    This PDF is a picture, so an SVG wouldn't be any sharper - "
                        "making a PNG instead.")
                else:
                    log("    This PDF is a picture, so no SVG was made - use the PNG.")

            if make_png:
                use = native_ppi if native_ppi else dpi
                longest_in = max(clip.width, clip.height) / 72
                if longest_in * use > MAX_SIDE_PX:
                    use = MAX_SIDE_PX / longest_in
                pix = page.get_pixmap(matrix=pymupdf.Matrix(use / 72, use / 72),
                                      clip=clip, alpha=False)
                pix.set_dpi(round(use), round(use))
                out = unique_path(out_dir / f"{stem}.png")
                pix.save(out)
                written.append(out)
                print_width_in = min(width_cm / 2.54, clip.width / 72)
                print_ppi = pix.width / print_width_in
                quality = ("sharp in print" if print_ppi >= 250 else
                           "acceptable" if print_ppi >= 150 else
                           "may look soft in print")
                log(f"  ✓ {out.name}: {pix.width} × {pix.height} px "
                    f"- about {print_ppi:.0f} ppi at {width_cm:g} cm wide ({quality})")
                if print_ppi >= 600:
                    log(f"    Enough detail to crop in to about 1/{int(print_ppi // 300)} "
                        "of the width in Word and still print sharply.")
                if native_ppi:
                    log("    This PDF holds a picture, so its original pixels were "
                        "kept (enlarging it further wouldn't add detail).")

            if make_svg:
                if trim and page.rotation == 0:
                    dx, dy = page.cropbox.x0, page.cropbox.y0
                    page.set_cropbox(clip + (dx, dy, dx, dy))
                svg = page.get_svg_image(text_as_path=True)
                out = unique_path(out_dir / f"{stem}.svg")
                out.write_text(svg, encoding="utf-8")
                written.append(out)
                log(f"  ✓ {out.name}: vector - stays sharp at any size and when "
                    "cropped in Word")

            # Readability check
            if native_ppi:
                log("    Text size can't be measured in a picture - check it at "
                    "100% zoom in Word.")
            else:
                small = small_text_size(page, clip)
                if small:
                    printed = small * shrink
                    if printed >= MIN_READABLE_PT:
                        log(f"    Text prints at about {printed:.1f}pt at "
                            f"{width_cm:g} cm wide - readable.")
                    else:
                        log(f"    ⚠ Smaller text prints at about {printed:.1f}pt at "
                            f"{width_cm:g} cm wide (aim for {MIN_READABLE_PT}pt+). "
                            "Crop to a section in Word, use a landscape page, "
                            "or enlarge the text in the original.")
    return written


# --------------------------------------------------------------------------
# Window
# --------------------------------------------------------------------------

def open_folder(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606 - opens Explorer
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        root = TkinterDnD.Tk()
        can_drop = True
    except Exception:
        root = tk.Tk()
        can_drop = False

    root.title(APP_NAME)
    root.geometry("720x600")
    root.minsize(600, 500)

    files = []
    messages = queue.Queue()
    last_out = {"dir": None}

    # ---- file list --------------------------------------------------------
    top = ttk.LabelFrame(root, text="PDF files", padding=8)
    top.pack(fill="both", expand=False, padx=10, pady=(10, 5))
    hint = "Drag PDFs here, or click Add PDFs…" if can_drop else "Click Add PDFs… to choose files"
    listbox = tk.Listbox(top, height=6, selectmode="extended")
    listbox.pack(side="left", fill="both", expand=True)
    listbox.insert("end", hint)
    listbox.configure(fg="grey")

    def add_files(paths):
        new = [p for p in paths if p.lower().endswith(".pdf") and p not in files]
        if not new:
            return
        if not files:
            listbox.delete(0, "end")
            listbox.configure(fg="black")
        for p in new:
            files.append(p)
            listbox.insert("end", Path(p).name)

    def choose_files():
        add_files(list(filedialog.askopenfilenames(
            title="Choose PDF files", filetypes=[("PDF files", "*.pdf")])))

    def remove_selected():
        for i in reversed(listbox.curselection()):
            if files:
                files.pop(i)
                listbox.delete(i)

    def clear_all():
        files.clear()
        listbox.delete(0, "end")
        listbox.insert("end", hint)
        listbox.configure(fg="grey")

    side = ttk.Frame(top)
    side.pack(side="right", fill="y", padx=(8, 0))
    ttk.Button(side, text="Add PDFs…", command=choose_files).pack(fill="x")
    ttk.Button(side, text="Remove", command=remove_selected).pack(fill="x", pady=4)
    ttk.Button(side, text="Clear", command=clear_all).pack(fill="x")

    if can_drop:
        listbox.drop_target_register(DND_FILES)
        listbox.dnd_bind("<<Drop>>", lambda e: add_files(root.tk.splitlist(e.data)))

    # ---- options ----------------------------------------------------------
    opts = ttk.LabelFrame(root, text="Options", padding=8)
    opts.pack(fill="x", padx=10, pady=5)

    fmt = tk.StringVar(value="png")
    ttk.Label(opts, text="Save as:").grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(opts, text="PNG picture (works everywhere)", variable=fmt,
                    value="png").grid(row=0, column=1, sticky="w")
    ttk.Radiobutton(opts, text="SVG (vector - best for Kumu/diagrams in Microsoft 365)",
                    variable=fmt, value="svg").grid(row=1, column=1, sticky="w")
    ttk.Radiobutton(opts, text="Both", variable=fmt, value="both").grid(row=2, column=1, sticky="w")

    ttk.Label(opts, text="PNG resolution:").grid(row=3, column=0, sticky="w", pady=(8, 0))
    dpi = tk.StringVar(value="300")
    ttk.Combobox(opts, textvariable=dpi, values=["150", "300", "600"], width=6,
                 state="readonly").grid(row=3, column=1, sticky="w", pady=(8, 0))

    trim = tk.BooleanVar(value=True)
    ttk.Checkbutton(opts, text="Trim blank space around the diagram",
                    variable=trim).grid(row=4, column=1, sticky="w", pady=(6, 0))

    ttk.Label(opts, text="Text width of page (cm):").grid(row=5, column=0, sticky="w", pady=(6, 0))
    width = tk.StringVar(value="16")
    ttk.Entry(opts, textvariable=width, width=6).grid(row=5, column=1, sticky="w", pady=(6, 0))

    out_mode = tk.StringVar(value="same")
    out_dir = tk.StringVar(value="")
    ttk.Label(opts, text="Save to:").grid(row=6, column=0, sticky="w", pady=(6, 0))
    out_row = ttk.Frame(opts)
    out_row.grid(row=6, column=1, sticky="w", pady=(6, 0))
    ttk.Radiobutton(out_row, text="Same folder as each PDF", variable=out_mode,
                    value="same").pack(side="left")

    def choose_out():
        d = filedialog.askdirectory(title="Choose a folder for the pictures")
        if d:
            out_dir.set(d)
            out_mode.set("folder")
            out_label.configure(text=d)

    ttk.Radiobutton(out_row, text="Choose folder…", variable=out_mode,
                    value="folder", command=choose_out).pack(side="left", padx=(10, 0))
    out_label = ttk.Label(opts, text="", foreground="grey")
    out_label.grid(row=7, column=1, sticky="w")

    # ---- run + log --------------------------------------------------------
    run_row = ttk.Frame(root)
    run_row.pack(fill="x", padx=10, pady=5)
    convert_btn = ttk.Button(run_row, text="Convert")
    convert_btn.pack(side="left")
    open_btn = ttk.Button(run_row, text="Open output folder", state="disabled",
                          command=lambda: open_folder(last_out["dir"]))
    open_btn.pack(side="left", padx=8)
    progress = ttk.Progressbar(run_row, mode="determinate")
    progress.pack(side="left", fill="x", expand=True)

    log_box = tk.Text(root, height=12, wrap="word", state="disabled")
    log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def write_log(text):
        log_box.configure(state="normal")
        log_box.insert("end", text + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")

    def pump():
        try:
            while True:
                kind, value = messages.get_nowait()
                if kind == "log":
                    write_log(value)
                elif kind == "progress":
                    progress["value"] = value
                elif kind == "done":
                    convert_btn.configure(state="normal")
                    if last_out["dir"]:
                        open_btn.configure(state="normal")
        except queue.Empty:
            pass
        root.after(100, pump)

    def worker(job_files, settings):
        for n, f in enumerate(job_files, 1):
            messages.put(("log", f"{Path(f).name}"))
            try:
                target = settings["out"] or str(Path(f).parent)
                done = convert_pdf(f, target, settings["fmt"], settings["dpi"],
                                   settings["trim"], settings["width"],
                                   log=lambda m: messages.put(("log", m)))
                if done:
                    last_out["dir"] = str(Path(done[-1]).parent)
            except Exception as e:  # keep going with the other files
                messages.put(("log", f"  ✗ Couldn't convert: {e}"))
            messages.put(("progress", n / len(job_files) * 100))
        messages.put(("log", "Finished.\n"))
        messages.put(("done", None))

    def start():
        if not files:
            messagebox.showinfo(APP_NAME, "Add one or more PDF files first.")
            return
        try:
            w = float(width.get())
            assert 2 <= w <= 60
        except Exception:
            messagebox.showerror(APP_NAME, "Text width should be a number of cm, e.g. 16.")
            return
        if out_mode.get() == "folder" and not out_dir.get():
            choose_out()
            if not out_dir.get():
                return
        settings = {
            "fmt": fmt.get(), "dpi": int(dpi.get()), "trim": trim.get(), "width": w,
            "out": out_dir.get() if out_mode.get() == "folder" else None,
        }
        convert_btn.configure(state="disabled")
        progress["value"] = 0
        threading.Thread(target=worker, args=(list(files), settings), daemon=True).start()

    convert_btn.configure(command=start)

    # PDFs dropped onto the .exe icon arrive as command-line arguments
    add_files([a for a in sys.argv[1:] if os.path.isfile(a)])

    write_log("Tip: insert the results in Word with Insert > Pictures > This Device "
              "(not copy and paste).\n")
    pump()
    root.mainloop()


if __name__ == "__main__":
    main()
