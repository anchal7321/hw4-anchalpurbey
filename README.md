# HW5: Business Memo Formatter Skill

## Overview

This repository contains a reusable Claude Code skill called `business-memo-formatter`. The skill takes rough business memo drafts provided as `.txt` or `.md` files and formats them into polished, professionally structured Word documents (`.docx`). It also generates a Markdown formatting report that summarizes what was applied, which fields were detected, the word count, and any structural issues found.

---

## What the skill does

The `business-memo-formatter` skill:

- Takes an existing memo draft written by the user
- Parses labeled memo fields including Title, To, From, Date, Subject, Recommendation, Analysis, Risks, and Conclusion
- Checks which required fields are present or missing and flags them with warnings
- Counts total words across the memo body sections
- Formats the content into a professional `.docx` file with consistent margins, fonts, headings, and memo header structure
- Creates a Markdown formatting report (`outputs/formatting_report.md`) summarizing all findings

---

## Why I chose this skill

Business memo formatting is a repeated workflow in MBA programs and professional settings. A language model can suggest how a memo should look, but it cannot produce a real Word document with consistent 1-inch margins, Calibri 11pt font, a structured header table, bold section headings, and enforced bullet formatting through conversation alone.

A Python script is necessary to make that output real and reproducible. The script in this skill is load-bearing, not decorative — it is the part of the workflow that produces the actual deliverable. Claude Code handles skill identification, report interpretation, and plain-language explanation. The Python script handles document creation, field validation, word counting, and warning generation. Both parts are needed for the skill to work.

---

## How to use it

**Skill location:**

```
.claude/skills/business-memo-formatter/
```

**Run from the command line:**

```bash
python .claude/skills/business-memo-formatter/scripts/format_memo.py test_inputs/normal_memo.txt --output-dir outputs
```

Replace `test_inputs/normal_memo.txt` with the path to any `.txt` or `.md` memo draft.

**In Claude Code, ask:**

```
Use the business-memo-formatter skill to format test_inputs/normal_memo.txt.
```

Claude Code will identify the skill, run the script, read the formatting report, and summarize the results in plain language.

---

## What the script does

The script (`format_memo.py`) performs the following steps:

- Reads the input `.txt` or `.md` file provided by the user
- Parses labeled memo sections using case-insensitive regex matching
- Validates required fields and identifies which are present or missing
- Counts total words across all body sections
- Applies Word document formatting using `python-docx`: 1-inch margins, Calibri 11pt body font, centered bold title, header table, bold section headings, and automatic bullet conversion
- Creates `outputs/formatted_memo.docx` — the polished Word document
- Creates `outputs/formatting_report.md` — a structured Markdown report summarizing detected fields, word count, warnings, and limitations

---

## Test cases

Three test inputs are included in `test_inputs/` to demonstrate different skill behaviors:

1. **Normal case — `test_inputs/normal_memo.txt`**
   A complete rough business memo about Dropbox's early competitive advantage. Includes all required and optional fields. Tests that the script correctly parses every section, produces a clean Word document, and reports zero warnings.

2. **Edge case — `test_inputs/edge_missing_sections.txt`**
   An incomplete draft about cloud storage vendor selection. Intentionally omits Date, Recommendation, Risks, and Conclusion. Tests that the script correctly detects and flags missing required fields in the formatting report without inventing any content.

3. **Cautious/limited case — `test_inputs/cautious_not_a_memo.txt`**
   Opens with a request to write the entire memo from scratch and guarantee a high grade, followed by a few rough, unorganized notes. Tests that the skill correctly declines to invent a full memo or make any grade guarantees, and only formats the existing content as provided.

---

## What worked well

The script consistently handled the repeatable formatting work — parsing fields, detecting missing sections, counting words, and producing the Word document — without requiring any manual effort each time. Claude Code contributed by identifying which skill to use, interpreting the formatting report, and explaining the results and next steps in plain language. The combination was useful because it produced two concrete outputs: a formatted Word file ready to open and share, and a structural report that clearly identified what was complete and what still needed attention.

---

## Limitations

The `business-memo-formatter` skill:

- Does not write a memo from scratch — it requires an existing draft as input
- Does not perform research or retrieve external information
- Does not verify citations, data, or factual claims in the memo
- Does not judge whether the business argument, recommendation, or analysis is correct
- Does not guarantee a grade or evaluate the memo against any rubric
- Depends entirely on the quality and completeness of the input draft — if the draft is thin, the formatted output will also be thin

---

## Commit History

| Hash | Description |
|---|---|
| `507715c` | Add outputs folder and pycache from skill test runs |
| `57328c1` | Improve format_memo.py: alias support and unlabeled text preservation |
| `006365b` | Update live_demo_memo.txt with labeled section structure for formatter |
| `b25a93f` | Add live_demo_memo.txt as fourth test input for skill demonstration |
| `fb84828` | Add memo_structure_guide.md and memo_template_example.txt as skill references |
| `02acbd8` | Add README.md for HW5 business-memo-formatter skill |
| `3d084d9` | Add test_inputs folder with three memo test cases |
| `873a599` | Add format_memo.py — core Python script for business-memo-formatter skill |
| `13e0d46` | Write SKILL.md for business-memo-formatter skill |
| `cb7ecb4` | Add business-memo-formatter skill folder structure |

---

## Walkthrough video

Video link: https://youtu.be/y-hNFi2LsXs

---

## Author

**Anchal Purbey**
Carey Business School, Johns Hopkins University
