---
name: post-consultation-presentation
description: "Build client-ready follow-up PDF presentations from a meeting summary and product URL. Use when creating sales/diagnostic materials: verify the product page, capture and cleanly integrate the product hero visual, write the client summary, calculate pricing, keep only the approved company slides, export to PDF, and remove temporary PPTX/brief files before delivery."
---

# Post-Consultation Presentation

Use this skill when the user needs a ready-to-send follow-up after a diagnostic or sales call.

The client-facing output is now always a `.pdf`. The `.pptx` is only an intermediate build artifact and should be deleted after validation unless the user explicitly asks to keep it.

## Configuration

Read `config.json` from the project root for:
- `company.consultant.name` - consultant name for the intro slide
- `company.consultant.photo` - path to consultant photo
- `company.brandText` - brand text for slide footer (e.g. "company.com")
- `paths.output_dir` - directory for output files
- `paths.reference_pdf` - path to the reference presentation PDF
- `defaults.presentation.slide1_title_line1` - first title line
- `defaults.presentation.slide1_title_line2` - second title line
- `defaults.presentation.slide1_ai_badge_text` - AI badge text
- `defaults.presentation.slide1_ai_badge_description` - AI badge description

## Current Standard

- Final file: `ClientName.pdf`
- Intermediate files: temporary brief JSON and `ClientName.pptx`; delete both after successful PDF validation.
- Builder script: [build_presentation.ps1](scripts/build_presentation.ps1)
- Brief example: [presentation_brief.example.json](references/presentation_brief.example.json)
- Run the builder through `powershell.exe -ExecutionPolicy Bypass -File ...`

## Required Workflow

1. Read the user-provided meeting summary.
2. Verify the exact product page from the supplied URL.
3. Build a structured brief JSON.
4. Run the builder with `-OutputPath ClientName.pptx` and `-PdfOutputPath ClientName.pdf`.
5. Validate the PDF/PPTX internals.
6. Delete the temporary brief and intermediate `.pptx`.
7. Report only the final PDF path.

Do not ask for routine approvals for local build, screenshot, crop, export, cleanup, or validation steps.

## Product Page Rules

- Always verify the live product page from the provided URL.
- Always use the real product title from the page.
- Slide 1 must contain the verified product title and use the verified URL for hyperlinks.
- Slide 1 must use a clean product hero visual from the verified product page, integrated into the right side of the purple band.
- Crop away sticky banners, text blocks, buttons, review strips, unrelated lower sections, and broken page fragments.
- Keep the product URL as a hyperlink on the product title and product visual.

## Slide Structure

The standard deck has 16 slides:

- Slide 1: direction/product page in the reference layout: large green two-line title, full purple product band, product name on the left, and a clean right-side product hero visual from the verified product page.
- Slide 2: client self-summary in first person, without a card container; the client name must be visible at the top as `Имя: <name>`.
- Slide 3: meeting summary in the reference layout, without a card container.
- Slide 4: personal offer/pricing as a bordered table, not cards.
- Slide 5: diagnostic/proof slide from the reference.
- Slide 6: "Давайте знакомиться" with consultant photo and name from config; the consultant name must be near the top, not at the bottom.
- Slides 7-14: company proof, expert, partners, teaching, programs, route/support slides from the reference.
- Slide 15: reference special-offer slide with discount.
- Slide 16: final offer and next step.

## Personalized Slides

Slide 1:

- Title from config (`slide1_title_line1` and `slide1_title_line2`).
- Product title and product visual are clickable and use the verified URL.
- Use the clean product hero image from the verified page, not a broken crop with partial text/buttons.

Slide 2:

- Must contain exactly 4 concise paragraphs.
- Must be written from the client's first-person perspective.
- Example: `Сейчас у меня есть действующий бизнес...`, not `Сейчас у вас есть...`.
- Replace the client name.
- Do not include the old prompt/question text like `Расскажите немного о себе...`; the slide should start directly with the client's self-summary.

Slide 3:

- Must contain exactly 7 concise paragraphs.
- Write from the consultant side: `На встрече мы зафиксировали...`, `для вас важно...`.
- Cover: current situation, risk of no action, product fit, expected result, pricing, next step.

Slide 4:

- Use the reference pricing table layout:
- `Полная цена`
- `Цена в рассрочку` (or config header)
- `Цена при единовременной оплате по карте` (or config header)
- `Размер платежа в рассрочку на N месяцев`

## Pricing Rules

- `fullPrice`: full tariff price.
- `specialPrice`: amount used for installment approval / "Цена в рассрочку".
- `oneTimePrice`: if explicitly provided, use it.
- If `oneTimePrice` is not provided, calculate card payment as `floor(specialPrice * (1 - discount/100))`.
- N-month installment is always `floor(specialPrice / months)`.

## Brief JSON Rules

Required fields:

- `clientName`
- `productUrl`
- `productTitle`
- `productTitlePrefix`
- `slide2Paragraphs` exactly 4 items, first-person client text.
- `slide3Paragraphs` exactly 7 items, consultant summary text.
- `pricing.productName`
- `pricing.fullPrice`
- `pricing.specialPrice`

Default fields for the new algorithm:

- `exportPdf: true`
- `pdfOutputPath: "...\\ClientName.pdf"`
- `pricing.oneTimeDiscountPercent` (from config, default 10)
- `pricing.specialPriceHeader` (from config)
- `pricing.oneTimePriceHeader` (from config)
- `pricing.installmentMonths` (from config, default 24)

## Validation

**CRITICAL: NEVER use the Read tool on PDF, PNG, JPG, or any image/binary files.** Use Bash commands and the validation script instead.

Validate before final response using **programmatic checks only** (no vision/model read of binary files):

1. **PDF exists and size > 10KB**: `ls -la <path>` via Bash.
2. **Page count**: Run `python scripts/validate_pdf.py <path> 16` or `python scripts/validate_pdf.py <path>` (default 16 pages). Do NOT use the Read tool on the PDF.
3. **Slide content**: Run the PowerShell builder with `-Verbose` or check the brief JSON for product title, URL, client text, and pricing values. Do NOT open the PDF/PPTX in the model.
4. **Cleanup**: Delete the temporary brief JSON and intermediate `.pptx` after validation. Confirm with `ls` that only the `.pdf` remains.

PowerShell COM is allowed only for PDF export. For validation, prefer the `validate_pdf.py` script and file checks over any visual/model inspection.

## Encoding Guardrail

The builder reads brief JSON as UTF-8 first and falls back to Windows-1251.

Do not reintroduce `Get-Content -Raw` as the primary JSON reader for Cyrillic briefs. When checking Russian text, inspect the Open XML content with UTF-8 output; console encoding can mislead.