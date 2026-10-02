#!/usr/bin/env python3
"""Generate self-contained lab-folder READMEs from canonical course data."""

import glob
import importlib
import os
import re
import shutil
import sys

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


def load_activities():
    acts = []
    for path in sorted(glob.glob(os.path.join(HERE, "data_domain[0-9]*.py"))):
        name = os.path.basename(path)[:-3]
        if not re.fullmatch(r"data_domain\d+", name):
            continue
        n = re.search(r"\d+", name).group()
        acts.extend(getattr(importlib.import_module(name), f"DOMAIN{n}"))
    return sorted(acts, key=lambda a: a["num"])


REPO = repo_root(HERE)
LABS = os.path.join(REPO, "labs")
ACTS = load_activities()
TOPICS = {t["num"]: t for t in C.TOPICS}


def lab_dir(a):
    return os.path.join(LABS, f"lab-{a['num']:02d}-{C.LAB_SLUGS[a['num']]}")


def company_slug():
    """Lab artifact filenames are stemmed on the scenario company, e.g.
    "Lumina Living Pte Ltd" -> "Lumina-Living"."""
    name = re.sub(r"\b(Pte|Ltd|Limited|Inc|LLC)\b\.?", "", C.COMPANY)
    return "-".join(name.split())


def resource_names(a):
    stem = f"{company_slug()}-Lab-{a['num']:02d}"
    if a["num"] == 1:
        return [
            "README.md",
            "TRAINER-GUIDE.md",
            f"{stem}-Candidates.xlsx" if a["num"]==1 else f"{stem}-Staff-Information.xlsx" if a["num"]==2 else f"{stem}-People-Numbers.xlsx" if a["num"]==7 else f"{stem}-Staff-Questions.xlsx" if a["num"]==9 else f"{stem}-Working-Workbook.xlsx",
            "templates/Lab-01-Trainer-Demonstration-Guide.docx",
        ]
    workbook = (
        f"{stem}-Candidates.xlsx" if a["num"]==1 else f"{stem}-Staff-Information.xlsx" if a["num"]==2 else f"{stem}-People-Numbers.xlsx" if a["num"]==7 else f"{stem}-Staff-Questions.xlsx" if a["num"]==9 else f"{stem}-Working-Workbook.xlsx"
        if a["num"] == 1 else f"{stem}-Working-Workbook.xlsx"
    )
    return [
        f"{stem}-HR-Brief.docx",
        f"{stem}-Claude-Generated-Work-Sample.docx",
        workbook,
        f"{stem}-Executive-Starter.pptx",
        "templates/Prompt-and-Review-Template.docx",
        "templates/Decision-and-Approval-Log.xlsx",
    ]


def looks_like_shell(value):
    first = value.strip().splitlines()[0] if value.strip() else ""
    return bool(re.match(r"^(pwd|find |python|python3|source |claude(?:\s|$)|pip |/mcp|/setup-cowork)", first))


for a in ACTS:
    t = TOPICS[a["topic"]]
    folder = lab_dir(a)
    os.makedirs(os.path.join(folder, "templates"), exist_ok=True)
    files = resource_names(a)
    out = [
        f"# Lab {a['num']} — {a['title']}", "",
        f"**Topic {t['code']}:** {t['title']}  |  **Day 1**  |  **Approx. {C.LAB_DURATIONS[a['num']]} min**  |  **Course:** {C.TITLE}", "",
        "## Company scenario", "", C.COMPANY_CONTEXT, "", a["desc"], "",
        "## Goal", "", a["objective"], "",
        "## What you'll build", "", a["build"], "",
        f"**Tools and techniques:** {a['services']}", "",
        "## Company use case", "",
        f"- **Department:** {a['case']['department']}",
        f"- **Sponsor:** {a['case']['sponsor']}",
        f"- **Business challenge:** {a['case']['challenge']}",
        f"- **Decision:** {a['case']['decision']}",
        f"- **Evidence:** {'; '.join(a['case']['sources'])}",
        f"- **Measures:** {'; '.join(a['case']['metrics'])}",
        f"- **Controls:** {'; '.join(a['case']['controls'])}", "",
        "## Files in this lab folder", "",
    ]
    out.extend(f"- `{name}`" for name in files)
    if a["num"] == 11:
        out.extend([
            "- `automation/update_daily_control.py`",
            "- `automation/generate_daily_brief.py`",
            "- `inputs/daily-input.csv`",
            "- `inputs/outlook-findings.json`",
        ])
    out.extend(["", "## Prerequisites", ""])
    out.extend(f"- {x}" for x in a["prerequisites"])
    if a.get("trainer_plan"):
        out.extend(["", "## Trainer delivery plan", "", "**Lab 01 is a 20-minute demonstration and guided practice. It is not a prompt-contract exercise.**", ""])
        out.extend(["| Time | Trainer action | What to teach | Learner evidence |", "|---|---|---|---|"])
        for timing, action, teaching, evidence in a["trainer_plan"]:
            out.append(f"| {timing} | {action} | {teaching} | {evidence} |")
        out.extend(["", "### Before class", ""])
        out.extend(f"- {x}" for x in a.get("trainer_preclass", []))
        out.extend(["", "### Do not teach in Lab 01", ""])
        out.extend(f"- {x}" for x in a.get("trainer_exclusions", []))
    out.extend(["", "## Process map", "", " → ".join(a["deck_flow"]), "", "## Steps", ""])
    for i, (instruction, payload) in enumerate(a["steps"], 1):
        out.extend([f"### Step {i}", "", instruction, ""])
        if payload:
            if looks_like_shell(payload):
                label, lang = "Command or in-app command", "bash"
            else:
                label, lang = "Prompt to give Claude", "text"
            out.extend([f"**{label}:**", "", f"```{lang}", payload, "```", ""])
    out.extend(["## Test it", "", a["test"], "", "## Troubleshooting", ""])
    for label, fix in a["troubleshooting"]:
        out.append(f"- **{label}.** {fix}")
    out.extend([
        "", "## Challenge", "", a["challenge"], "",
        "## Reflection", "", a["reflection"], "",
        "## Deliverable", "", a["build"], "",
        "## Current product references", "",
    ])
    # Keep individual guides compact; authoritative references are relevant to all
    # activities and the complete supplied-source list remains in labs/README.md.
    for name, url in C.LG_REFERENCES:
        keys = ("microsoft 365", "connector", "cowork", "claude code")
        if a["num"] == 1:
            keys += ("chrome", "word", "outlook", "office add-ins")
        if any(key in name.lower() for key in keys):
            out.append(f"- [{name}]({url})")
    out.extend(["", "---", "", f"*{C.TITLE} · {C.COURSE_CODE} · Version {C.VERSION} · © 2026 Tertiary Infotech Academy Pte Ltd*", ""])
    path = os.path.join(folder, "README.md")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out))
    print("Saved", path)
    if a.get("trainer_plan"):
        trainer = [
            f"# Trainer Guide — Lab {a['num']}", "",
            f"**Teaching outcome:** {a['objective']}", "",
            "## What you teach", "",
            "1. The Office add-in works with the open Office item.",
            "2. The Microsoft 365 connector is a separate route in Claude Desktop for authorised Microsoft 365 context.",
            "3. Claude in Chrome is the hands-on Outlook web route for this lab.",
            "4. Learners review the live message and send only after trainer approval.", "",
            "## 20-minute run sheet", "",
            "| Time | Trainer action | What to teach | Learner evidence |", "|---|---|---|---|",
        ]
        for timing, action, teaching, evidence in a["trainer_plan"]:
            trainer.append(f"| {timing} | {action} | {teaching} | {evidence} |")
        trainer.extend(["", "## Before class", ""])
        trainer.extend(f"- {x}" for x in a.get("trainer_preclass", []))
        trainer.extend(["", "## Keep out of Lab 01", ""])
        trainer.extend(f"- {x}" for x in a.get("trainer_exclusions", []))
        trainer.extend(["", "## Completion standard", "", a["test"], ""])
        trainer_path = os.path.join(folder, "TRAINER-GUIDE.md")
        with open(trainer_path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(trainer))
        print("Saved", trainer_path)

# The requested v2 structure replaces generated flat lab Markdown files.  The
# previous version remains recoverable from Git history and courseware/archive.
for old in glob.glob(os.path.join(LABS, "lab-*.md")):
    os.remove(old)

rows = [
    f"# Labs — {C.TITLE}", "",
    f"**Course Code:** {C.COURSE_CODE}  |  **Version {C.VERSION} · {C.VERSION_DATE}**", "",
    f"All {len(ACTS)} activities use one fictional company, **{C.COMPANY}**, and build a connected FY2027 planning and management pack.", "",
    "The presentation explains concepts, decisions, process maps and realistic work samples. These lab folders and the Learner Guide contain the complete prompts, commands and verification checks.", "",
    "## Lab folder standard", "",
    "Every lab folder contains a detailed README, a realistic company Word brief, an Excel working file, an editable PowerPoint starter and reusable review/approval templates. Lab 11 also contains safe local automation starters.", "",
    "## Lab sequence", "",
    "| Topic | Lab | Activity | Company outcome |", "|---|---:|---|---|",
]
for a in ACTS:
    folder = f"lab-{a['num']:02d}-{C.LAB_SLUGS[a['num']]}"
    rows.append(f"| {TOPICS[a['topic']]['code']} | {a['num']:02d} | [{a['title']}]({folder}/README.md) | {a['build']} |")

rows.extend(["", "## Supplied research and further learning", ""])
rows.extend(f"- [{name}]({url})" for name, url in C.LAB_RESEARCH_SOURCES)
rows.extend(["", "## Authoritative product guidance", ""])
rows.extend(f"- [{name}]({url})" for name, url in C.LG_REFERENCES)
rows.extend([
    "", "See [tools.md](tools.md) for account, add-in, connector, Cowork and Claude Code requirements.", "",
    "---", "", f"*{C.TITLE} · {C.COURSE_CODE} · Version {C.VERSION} · © 2026 Tertiary Infotech Academy Pte Ltd*", "",
])
with open(os.path.join(LABS, "README.md"), "w", encoding="utf-8") as fh:
    fh.write("\n".join(rows))
print("Saved", os.path.join(LABS, "README.md"))
