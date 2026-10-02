#!/usr/bin/env python3
"""Render representative lab Office artifacts into slide-ready preview images."""

import glob
import importlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from PIL import Image, ImageChops

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import course_data as C


def repo_root(start):
    d = start
    for _ in range(8):
        d = os.path.dirname(d)
        if os.path.isdir(os.path.join(d, "courseware")) and os.path.isdir(os.path.join(d, "labs")):
            return d
    raise RuntimeError("Course repository not found")


REPO = repo_root(HERE)
OUT = os.path.join(REPO, "courseware", "assets", "screenshots")
os.makedirs(OUT, exist_ok=True)


def trim(path):
    image = Image.open(path).convert("RGB")
    bg = Image.new("RGB", image.size, "white")
    box = ImageChops.difference(image, bg).getbbox()
    if box:
        margin = 18
        box = (max(0, box[0]-margin), max(0, box[1]-margin), min(image.width, box[2]+margin), min(image.height, box[3]+margin))
        image = image.crop(box)
    image.save(path, quality=94)


def render(source, dest, page=1):
    with tempfile.TemporaryDirectory(prefix=f"{C.COURSE_CODE.lower()}-preview-") as tmp:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, source], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdfs = glob.glob(os.path.join(tmp, "*.pdf"))
        if not pdfs:
            raise RuntimeError(f"No PDF rendered from {source}")
        prefix = os.path.join(tmp, "page")
        subprocess.run(["pdftoppm", "-png", "-r", "120", "-f", str(page), "-l", str(page), pdfs[0], prefix], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pngs = glob.glob(prefix + "-*.png")
        if not pngs:
            raise RuntimeError(f"No preview page rendered from {source}")
        shutil.copy2(pngs[0], dest)
        trim(dest)
        print("Saved", dest)


for num in range(1, 12):
    folder = glob.glob(os.path.join(REPO, "labs", f"lab-{num:02d}-*"))[0]
    if num == 1:
        source = os.path.join(folder, "templates", "Lab-01-Trainer-Demonstration-Guide.docx"); page = 1
    elif num == 7:
        source = glob.glob(os.path.join(folder, "*-Working-Workbook.xlsx"))[0]; page = 1
    elif num == 8:
        source = glob.glob(os.path.join(folder, "*-Executive-Starter.pptx"))[0]; page = 8
    else:
        source = glob.glob(os.path.join(folder, "*-Claude-Generated-Work-Sample.docx"))[0]; page = 1
    render(source, os.path.join(OUT, f"lab-{num:02d}-artifact.png"), page)

# Classroom-readable zooms from the Lab 7 dashboard.  The full dashboard remains
# available as one exhibit; these crops let learners inspect the KPI strip and
# the two management charts without shrinking labels to thumbnail size.
dash = Image.open(os.path.join(OUT, "lab-07-artifact.png")).convert("RGB")
w, h = dash.size
crops = {
    "lab-07-kpis.png": (0, 0, w, int(h * 0.23)),
    "lab-07-revenue-chart.png": (0, int(h * 0.22), int(w * 0.51), int(h * 0.68)),
    "lab-07-contribution-chart.png": (int(w * 0.49), int(h * 0.22), w, int(h * 0.68)),
}
for name, box in crops.items():
    path = os.path.join(OUT, name)
    dash.crop(box).save(path, quality=95)
    print("Saved", path)
