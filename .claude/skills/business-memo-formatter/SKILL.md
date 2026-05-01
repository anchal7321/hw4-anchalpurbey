---
name: business-memo-formatter
description: Formats rough business memo drafts from text or Markdown files into polished Word documents with consistent memo headers, margins, fonts, spacing, headings, and bullet formatting. Use when the user asks to format, clean up, or package an existing memo draft into a professional .docx file.
---

# business-memo-formatter

## 1. What this skill does

This skill takes an existing rough business memo draft — provided as a `.txt` or `.md` file — and formats it into a polished, professionally structured Word document (`.docx`). It applies consistent memo headers, margins, fonts, spacing, headings, and bullet formatting. It also generates a short formatting report summarizing what was applied and flagging any issues found.

This skill formats content. It does not write a memo from scratch, rewrite the user's argument, or make editorial judgments about the business content.

---

## 2. When to use this skill

Use this skill when the user:

- Has an existing memo draft in a `.txt` or `.md` file and wants it formatted into a `.docx`
- Asks to "clean up", "format", "package", or "polish" a memo
- Wants a professional Word document output from a rough draft
- Needs consistent structure, spacing, and visual formatting applied to a business memo

---

## 3. When not to use this skill

Do not use this skill when the user:

- Wants a memo written from scratch with no existing draft provided
- Asks for feedback on the business argument, strategy, or recommendation
- Wants the memo rewritten or substantially edited — handle that as a separate task before formatting
- Needs research, data analysis, or content generation

---

## 4. Expected inputs

The user must provide a `.txt` or `.md` file containing a rough business memo draft. The memo may include any combination of the following labeled fields:

```
Title:
To:
From:
Date:
Subject:
Recommendation:
Analysis:
Risks:
Conclusion:
```

Fields do not need to be perfectly formatted — the skill will parse and apply structure based on what is present.

---

## 5. Step-by-step instructions

1. **Confirm the input file exists.** Check that the user has provided a `.txt` or `.md` file path. If no file is provided, ask the user to share the memo draft before proceeding.

2. **Check for required memo fields.** Before running the script, scan the input file for the expected fields (`To:`, `From:`, `Date:`, `Subject:`). Flag any that are missing and inform the user.

3. **Count words.** Check the total word count of the memo draft. Warn the user if the memo appears unusually long (over 800 words) or unusually short (under 100 words).

4. **Run the formatting script** using the following command:

   ```bash
   python .claude/skills/business-memo-formatter/scripts/format_memo.py <input_file> --output-dir outputs
   ```

   Replace `<input_file>` with the actual path to the user's memo draft file.

5. **Confirm outputs were created.** Verify that both output files exist:
   - `outputs/formatted_memo.docx`
   - `outputs/formatting_report.md`

6. **Summarize the formatting report.** Read `outputs/formatting_report.md` and provide the user with a short summary of what was applied and any issues or warnings flagged.

7. **Preserve original meaning.** Do not alter the user's business content, argument, numbers, or conclusions during formatting.

---

## 6. Expected output format

Running this skill produces two files in the `outputs/` directory:

| File | Description |
|---|---|
| `outputs/formatted_memo.docx` | The polished Word document with consistent memo structure, headers, fonts, margins, spacing, and bullet formatting applied |
| `outputs/formatting_report.md` | A short Markdown report summarizing what formatting was applied, which fields were detected, word count, and any warnings or missing fields |

After the script runs, summarize the formatting report for the user in plain language.

---

## 7. Important limitations and checks

**The skill will:**
- Apply formatting and structure to the existing draft
- Detect and report missing memo fields
- Count words and warn if the memo length is outside a normal range
- Preserve the user's original wording, meaning, and content throughout

**The skill will not:**
- Invent or fill in missing content, fields, or sections
- Conduct research or retrieve external information
- Rewrite the memo or improve the business argument
- Judge whether the recommendation or analysis is correct
- Guarantee a grade or evaluate the memo against a rubric

**Important:** If required fields such as `To:`, `From:`, `Date:`, or `Subject:` are missing from the draft, the skill will flag them and ask the user whether to proceed or update the draft first.
