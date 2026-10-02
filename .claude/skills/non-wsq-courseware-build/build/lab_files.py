#!/usr/bin/env python3
"""Build the "Files you will use" map for a lab from the files actually on disk.

Steps in the course data refer to artifacts in plain English — "the Source
Register workbook", "the company template".  Learners could not tell which file
or which sheet tab that meant.  This module inspects the real lab folder and
produces an explicit name -> file -> tab mapping, so the guidance can never
drift from what was shipped.
"""

import os

import openpyxl

# Plain-English names that appear in step text, mapped to the sheet tab that
# actually holds them.  Only tabs present in the workbook are emitted.
SHEET_ALIASES = {
    "Where_Info_Is_Kept": "Where staff information is kept",
    "Review_Log": "The review log — where you record who checked the work",
    "What_Claude_May_Do": "What Claude may do with each one",
    "Summary": "The summary view",
    "Assumptions": "The assumptions used by every formula",
    "Transactions": "The transaction data",
    "Budget": "The budget figures",
    "Analysis": "The analysis working area",
    "Audit_Log": "The audit log",
    "Dashboard": "The dashboard",
    "Read_Me": "How this workbook is organised — read this tab first",
    "Candidates": "The 24 applicants and the two columns you fill in",
    "Staff_List": "The staff data you analyse",
    "Staff_Messages": "The staff messages you sort",
}

DOC_ALIASES = [
    ("-HR-Brief.docx", "The HR brief — the source you read before asking Claude anything"),
    ("-Claude-Generated-Work-Sample.docx", "A finished example showing the standard to aim for"),
    ("-Executive-Starter.pptx", "The company PowerPoint template — keep its slide master"),
    ("-Candidates.xlsx", "The list of job applicants you work on"),
    ("-Staff-Information.xlsx", "The workbook you complete in this lab"),
    ("-People-Numbers.xlsx", "The workbook you complete in this lab"),
    ("-Staff-Questions.xlsx", "The workbook you complete in this lab"),
    ("-Working-Workbook.xlsx", "The workbook you complete in this lab"),
]

TEMPLATE_ALIASES = {
    "Prompt-and-Review-Template.docx": "The blank prompt contract you fill in",
    "Decision-and-Approval-Log.xlsx": "Where you record the decision and who approved it",
    "Lab-01-Trainer-Demonstration-Guide.docx": "Trainer demonstration guide",
    "Daily-Brief-Template.docx": "The daily brief template",
}


def sheet_names(path):
    try:
        return openpyxl.load_workbook(path, read_only=True).sheetnames
    except Exception:
        return []


APP_BY_EXT = {".docx": "Word", ".xlsx": "Excel", ".pptx": "PowerPoint",
              ".csv": "Excel", ".json": "Text editor", ".py": "Claude Code"}


def app_for(name):
    for ext, app in APP_BY_EXT.items():
        if name.lower().endswith(ext):
            return app
    return "—"


def describe(folder):
    """Return rows of (what the guide calls it, app, file to open, tab to click)."""
    rows = []
    if not os.path.isdir(folder):
        return rows

    names = sorted(os.listdir(folder))
    for suffix, label in DOC_ALIASES:
        for name in names:
            if not name.endswith(suffix):
                continue
            if name.lower().endswith(".xlsx"):
                tabs = sheet_names(os.path.join(folder, name))
                rows.append((label, app_for(name), name, "opens on the first tab"))
                for tab in tabs:
                    if tab in SHEET_ALIASES:
                        rows.append((SHEET_ALIASES[tab], app_for(name), name, f'the "{tab}" tab'))
            else:
                rows.append((label, app_for(name), name, "—"))

    tdir = os.path.join(folder, "templates")
    if os.path.isdir(tdir):
        for name in sorted(os.listdir(tdir)):
            label = TEMPLATE_ALIASES.get(name)
            if label:
                rows.append((label, app_for(name), f"templates/{name}", "—"))
    return rows
