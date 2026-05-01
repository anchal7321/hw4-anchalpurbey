# Memo Structure Guide

This guide describes the labeled sections that the `business-memo-formatter` skill expects in a draft memo file. The script parses these fields case-insensitively, so `RECOMMENDATION:` and `Recommendation:` are both valid.

---

## Required Fields

### Title
A short descriptive name for the memo. Displayed as a centered bold heading at the top of the formatted Word document.
Example: `Title: Dropbox's Early Competitive Advantage: A Strategic Assessment`

### To
The recipient or audience of the memo — typically a committee, team, or individual.
Example: `To: Strategy & Innovation Committee`

### From
The author's name and role.
Example: `From: Anchal Purbey, MBA Candidate`

### Date
The date the memo was written or submitted. If omitted, the header table will show `[Not provided]` — the script will not invent a date.
Example: `Date: May 1, 2026`

### Subject
A concise one-line statement of what the memo addresses.
Example: `Subject: Evaluation of Dropbox's Product-Led Growth Strategy`

### Recommendation
The author's primary business recommendation — what action or decision the memo is advocating for. This is the most important body section. The script will issue a warning if it is missing.
Example: `Recommendation: Dropbox's referral program model should be studied as a template for low-cost user acquisition strategies.`

### Analysis
The supporting argument, evidence, and reasoning behind the recommendation. This is typically the longest section and may include bullet points. Lines starting with `- ` or `* ` are automatically converted to formatted bullet points in the Word document.

### Conclusion
A closing synthesis that ties the analysis back to the recommendation. Should summarize the key takeaway without introducing new information. The script will issue a warning if it is missing.

---

## Optional Field

### Risks
Potential downsides, challenges, or uncertainties associated with the recommendation. Including this section strengthens the memo's credibility. It is optional, but the script will note its absence in the formatting report.

---

## What the formatter checks

The script checks **structure and formatting only**. Specifically, it:

- Detects which labeled fields are present or missing
- Counts total words across the body sections and warns if the count exceeds 750
- Applies consistent Word formatting (margins, fonts, headings, bullet conversion)
- Produces a Markdown report listing detected fields, warnings, and limitations

The script does **not**:
- Evaluate whether the business argument is correct or convincing
- Judge the quality of the analysis or recommendation
- Compare the memo against any course rubric
- Guarantee a grade or academic outcome
- Fill in missing sections or rewrite any content

The quality and completeness of the formatted output depends entirely on the quality and completeness of the input draft.
