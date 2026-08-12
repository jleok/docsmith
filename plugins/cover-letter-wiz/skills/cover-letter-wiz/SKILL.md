---
name: cover-letter-wiz
description: Writes a tailored cover letter as a Word document (.docx) from a finished resume plus a job posting, automatically matching the resume's font, sizes, and margins so the two documents look like a set, and deliberately avoiding the punctuation and phrasing that make writing read as AI generated. Use this whenever the user wants a cover letter, letter of interest, letter of application, or "the letter that goes with this resume" — whether or not they name the skill, and including phrasings like "write a cover letter for this job," "I need a cover letter for this posting," or "can you draft a letter of interest." Also use when the user attaches a finished resume alongside a job posting and asks for a letter. Do NOT use this to write, tailor, or rewrite the resume itself; that is the federal-resume-tailor skill's job.
compatibility: Requires docx (npm, preinstalled), pandoc, poppler (pdftotext, pdftoppm, pdfinfo), pdfplumber, and LibreOffice via the docx skill's scripts/office/soffice.py. No network access needed.
---

# Cover-Letter-Wiz

## What this does

Takes a finished, submission-ready resume plus the job it is being submitted
to, and writes the cover letter that goes with it. The letter matches the
resume's visual formatting so the pair looks like one set of documents, reads
like a person wrote it rather than a language model, and carries at least one
detail that could not appear in anyone else's letter.

Output is one .docx file. Nothing else.

## Inputs

Three. Ask for whatever is missing before doing anything else.

1. **The finished resume**, attached as .docx, .doc, or .pdf. Identify it by
   filename from the attachments. This is the source of both the content and
   the formatting.
2. **Bonus human mentions**, meaning any notes, instructions, opinions, or
   stray thoughts the user wants worked in. Usually typed straight into the
   prompt, occasionally pasted as a text file. This is where the letter's
   human flair comes from, so it matters more than its casual name suggests.
3. **The job posting**, usually pasted as text.

If any are missing, ask for the missing ones as a short numbered list and
stop. Do not proceed on a guess, and do not substitute a job title for a
posting. If the user says they have nothing to add for input 2, accept that
and source the flair from the resume instead (see `references/voice.md`).

## Process

### 1. Read the resume

Content:

```bash
pandoc -t markdown resume.docx          # .docx
pdftotext -layout resume.pdf -          # .pdf
```

Legacy `.doc` must be converted first:

```bash
python3 /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to docx resume.doc
```

Pull out the full name, email, phone, LinkedIn if present, and every concrete
accomplishment with its numbers. Note which experiences are strongest, and
note anything the resume states flatly that has a story behind it.

### 2. Extract the formatting spec

This is what makes the letter match the resume rather than merely look tidy.

```bash
python3 scripts/inspect_format.py resume.docx    # or resume.pdf
```

It prints body font, body size, name size, page size, and margins, in the
units docx-js wants. Read the `notes` line, and treat anything printed as
`UNKNOWN` as a real gap rather than a rounding error.

Then render the resume's first page and look at it, because alignment and
spacing are not in the spec block:

```bash
python3 /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf resume.docx
pdftoppm -jpeg -r 100 resume.pdf rpage && ls rpage-*.jpg
```

Check how the name and contact line sit, centered or left, and roughly how
much air is between blocks. Match it.

If the spec comes back mostly `UNKNOWN`, say so in your reply and fall back to
Calibri, 10pt body, 14pt name, US Letter, 0.5in margins, which is the user's
standing resume format. Do not silently invent a spec.

### 3. Read the posting

Find what the job actually needs, which is usually two or three things
repeated across the duties and qualifications sections. Note the employer or
agency name, the job title, the announcement number if there is one, and any
named hiring contact.

Ignore the posting's section headings as a structure for the letter. Walking
its bullet list in order is exactly the robotic repeat this skill exists to
avoid.

### 4. Decide the shape

Default to prose, three or four paragraphs.

Use three short bullets in the middle only when the posting presents a
discrete checklist of required or specialized qualifications AND the evidence
for them comes from genuinely different experiences that do not flow as one
paragraph. Even then, cap it at three, write them as full sentences, and
introduce them with a sentence ending in a period, never a colon.

If the letter is already short, or the posting reads as prose, use prose.

### 5. Draft

Read `references/voice.md` now and follow it. It is the part of this skill
most likely to fail quietly. The short version, which does not replace
reading the file:

- No em dashes, en dashes, colons, or semicolons anywhere.
- No "I am writing to express my interest," no "proven track record," no
  "leverage," no "not just X but Y," no rule of three, no rhetorical
  questions.
- Vary sentence and paragraph length. Contractions are fine.
- The letter says why the work mattered. The resume already said what it was.
- At least one passage that could not appear in anyone else's letter, drawn
  from input 2 wherever possible.
- Everything traces to the resume, the notes, or the posting. Never invent an
  accomplishment, a motivation, a connection, or a story.

### 6. Layout

Build in this order:

```
[Full name]                      resume's name size, bold, resume's alignment
[email  |  phone  |  linkedin]   body size, same alignment, no scheme on URLs
                                 blank line
[Month D, YYYY]                  today's date, left
                                 blank line
[Employer or agency name]        left, only if the posting names one
[Job title, and announcement number if there is one]
                                 blank line
Dear [Agency or Company] Hiring Team,
                                 comma, never a colon
                                 blank line
[body]
                                 blank line
Sincerely,
                                 three empty paragraphs, the signature space
[Full name]
```

If the posting names a hiring contact, address them by name instead of the
team. Write LinkedIn and any other URL without `https://` so no colon reaches
the page.

### 7. Build the .docx

Use docx-js (Node `docx`, preinstalled, do not npm install).

- Font: the extracted body font on every single `TextRun`.
- Sizes: half-points, so 10pt is `size: 20` and 14pt is `size: 28`.
- Page: the extracted `page_w_dxa` and `page_h_dxa`.
- Margins: the extracted margins, all four.
- Paragraph spacing: `spacing: { after: 200 }` between body paragraphs unless
  the resume clearly uses more or less.
- Signature space: three empty `Paragraph` elements between the closing and
  the typed name. Never `\n`.
- Bullets, if used: a `numbering` config with `LevelFormat.BULLET`, never a
  literal `•`.
- Hard limit one page. If it runs over, cut the weakest body sentences.
  Never shrink the font or margins to force the fit, because that breaks the
  match with the resume, which is the whole point.

### 8. Verify

Run all three checks. The first one is not optional, because banned
punctuation is the failure the user will notice first and it is trivially
detectable.

```bash
# 1. Banned punctuation. Expect no output.
unzip -p out.docx word/document.xml | python3 -c "
import sys, re
body = ' '.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', sys.stdin.read()))
for ch, label in [('\u2014','em dash'), ('\u2013','en dash'), (':','colon'), (';','semicolon')]:
    if ch in body:
        print('FOUND', label, '->', [s for s in body.split('.') if ch in s][:2])
"

# 2. One page, and look at it.
python3 /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf out.docx
pdfinfo out.pdf | grep Pages
pdftoppm -jpeg -r 100 out.pdf page && ls page-*.jpg

# 3. Formatting matches the resume spec.
unzip -p out.docx word/document.xml | grep -o 'w:ascii="[^"]*"' | sort | uniq -c
unzip -p out.docx word/document.xml | grep -o '<w:sz w:val="[0-9]*"' | sort | uniq -c
```

Read the rendered page image and confirm the signature space is actually
visible as a gap, the header sits like the resume's header, and the page is
not bottom heavy. Confirm the fonts and sizes match the spec from step 2.

Then reread the draft against the final checklist at the end of
`references/voice.md`.

### 9. Deliver

Copy only the `.docx` to `/mnt/user-data/outputs/` and present it with
`present_files`. Name it `Cover_Letter_[LastName]_[Employer or Job Title].docx`.

The PDF from step 8 is a verification artifact. Keep it in the working
directory, do not copy it to outputs, and do not present it. Produce a PDF as
a deliverable only when the user specifically asks for one in their prompt.

In your reply, mention the format spec you matched, and flag anything you had
to fall back on or could not source.

## Guardrails

- Never invent an accomplishment, metric, credential, motivation, story, or
  personal connection. Tailoring means selecting and framing what is already
  in the resume, the notes, or the posting.
- Never claim a skill or tool the resume does not support, even when the
  posting asks for it. If a required qualification is genuinely absent,
  either address the gap honestly or leave it alone. Do not paper over it.
- Never restate the resume line by line. If a paragraph could be replaced by
  reading the resume, cut it.
- If input 2 is empty and the resume offers nothing specific enough for real
  flair, ask the user for one detail rather than manufacturing one.
- Deliver .docx only unless a PDF was specifically requested.
