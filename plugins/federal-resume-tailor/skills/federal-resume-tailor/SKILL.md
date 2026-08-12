---
name: federal-resume-tailor
description: Tailors a federal (USAJOBS-format) resume to a specific target job by rewriting a master resume's bullet points to mirror the target job posting's language and reprioritizing the most relevant experience, then builds the result as a Word document (.docx) formatted to federal resume rules (2-page limit, Calibri font, required header/education fields, etc.). Use ONLY when the user explicitly invokes this skill by name (e.g. "use the resume tailor skill," "run the federal resume tailor") — do not trigger just because a resume and a job posting appear together in the same message, and do not trigger for general resume questions, proofreading, or advice. Wait for explicit invocation.
compatibility: Requires Node with the `docx` package (declared in the plugin's package.json) to construct the output file. No other dependencies and no network access needed.
---

# Federal Resume Tailor

## What this does
Takes (1) a master resume (the user's full bullet-point bank, usually a .docx upload) and (2) a target job's context (job posting text, URL, or announcement excerpt), and produces one tailored resume — rewritten and reprioritized to fit that specific job — formatted to federal (USAJOBS) resume standards. Output is a single .docx file.

## When to run this
Only when the user explicitly asks for this skill by name or unmistakably asks to run this exact pipeline. Do not auto-trigger just because a resume and job posting appear together in the same message — this is opt-in only, by design.

## Inputs needed
1. **Master resume** — if it's a .docx upload, read it with `pandoc -t markdown file.docx` (see the docx skill). Identify every job entry (employer, title, dates, hours/week, series & grade if federal), all bullet points, education, and relevant volunteer/community roles.
2. **Target job context** — job posting text, a URL to fetch, or at minimum the Duties / Requirements-Qualifications / How You Will Be Evaluated sections (use whichever the posting actually has — not every posting includes all three). If only a job title is given with no posting text, ask for the posting — tailoring quality depends on matching its exact language.

If either input is missing, ask for it rather than guessing.

## Process
1. **Parse the target job context**: pull exact terminology from its Duties/Requirements/Evaluation sections — specific tools, certifications, and skill terms verbatim (e.g., "MS Project," not a paraphrase). Note whichever 2-3 duties or focus areas get the most emphasis (repeated, or called out in a "why does this job exist"-style section) — these set the priority order for step 2.
2. **Select and reprioritize bullets**, per job entry in the master resume:
   - Lead each role with whichever bullet is the closest match to the posting's most-emphasized duties, not necessarily the order the bullets appear in the master resume.
   - Keep/prioritize bullets relevant to the target job; cut or deprioritize outdated/unrelated ones if space is tight.
   - Reword using the format: "Accomplished [X], as measured by [Y], by doing [Z]" — quantify with numbers/percentages/dollars wherever the master resume supports it.
   - Mirror the posting's exact terms only where the underlying experience genuinely matches — never fabricate a tool, certification, or accomplishment absent from the master resume.
   - Reprioritize the Skills section the same way: most job-relevant skills first.
3. **Assemble the resume**:
   - Header: full name, email, phone (and LinkedIn if present in the master resume) only — no photo, SSN, age/sex/religion, or other demographic info.
   - Optional short SUMMARY (2-3 lines), tailored to the target job but grounded only in facts already in the master resume — a good place to echo the posting's own key phrases.
   - Each job: employer, title, start/end dates (month + year), hours/week; federal roles also need series & grade (e.g., "Program Analyst GS-343-11") — only include series/grade where the role actually had one.
   - Education: school, completion date, degree type, GPA if relevant.
   - Relevant volunteer/community roles, if supportive of qualifications.
   - Plain language throughout — no unexplained acronyms/jargon.
   - **If the master resume is missing a required field (most commonly hours/week per job), do not guess or default to 40 — insert a visible placeholder like `Hours/week: [confirm]` and flag it to the user afterward.**
4. **Build the .docx** using docx-js (Node `docx` package, preinstalled):
   - Page: US Letter — `size: { width: 12240, height: 15840 }` (DXA).
   - Margins: 0.5" on all sides — `margin: { top: 720, bottom: 720, left: 720, right: 720 }` (720 DXA = 0.5").
   - Font: Calibri on every `TextRun` — `font: "Calibri"`.
   - Sizes: 14pt for the name and ALL-CAPS section headers (`size: 28`), 10pt for everything else (`size: 20`) — docx sizes are in half-points.
   - Bullets: a `numbering` config with `LevelFormat.BULLET`, never a literal "•" typed into a TextRun.
   - Hard limit: 2 pages. If content still runs long after reprioritizing, cut the lowest-relevance bullets first — never shrink font/margins below spec to force a fit.
5. **Verify before delivering**:
   ```bash
   python <docx-skill-dir>/scripts/office/soffice.py --headless --convert-to pdf output.docx
   pdftoppm -jpeg -r 100 output.pdf page && pdfinfo output.pdf | grep Pages
   ```
   Look at the rendered page image(s) and confirm the page count is ≤2. Spot check the XML (`unzip -p output.docx word/document.xml | grep -o 'w:sz w:val="[0-9]*"' | sort | uniq -c`) to confirm only sizes 20 and 28 appear, and that margins are 720 on all sides.
6. **Save as .docx** and deliver to the user with `present_files`.

## Guardrails
- Never invent accomplishments, metrics, tools, or qualifications absent from the master resume — tailoring means reprioritizing and rewording, not fabricating.
- Never include SSN, photo, age, sex, religion, or other demographic data, even if present in the master resume.
- If the master resume or job context is missing, ask rather than guess.
- If a required field (e.g., hours/week) isn't in the master resume, mark it `[confirm]` in the output and flag it in your reply — don't default to a guessed number.
