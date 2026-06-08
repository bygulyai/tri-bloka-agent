---
name: tri-bloka
description: "Build the full post-consultation package: Block 1 CRM report, Block 2 personalized client message, and Block 3 a validated client-ready PDF presentation. Use when the user sends raw consultation notes and asks for три блока, CRM + сообщение + презентация, СРМ + сообщение + презентация, or expects the new combined workflow."
---

# Tri Bloka

## Purpose

Use this skill for the combined daily workflow after a sales consultation. The output must always contain three separated blocks:

1. CRM report.
2. Client message.
3. PDF presentation result.

Keep the existing skills separate. This skill orchestrates them and adds the final three-block contract.

## Configuration

Read the configuration from `config.json` in the project root. Required fields:

- `company.name` - company name for branding
- `company.site` - company website URL
- `company.consultant.name` - consultant name for slide 6
- `company.consultant.photo` - path to consultant photo
- `company.consultant.competencies` - list of competence strings
- `paths.output_dir` - directory for PDF output
- `paths.reference_pdf` - path to reference presentation PDF
- `defaults.pricing.oneTimeDiscountPercent` - default discount for card payment (usually 10)
- `defaults.pricing.installmentMonths` - default installment period (usually 24)
- `defaults.pricing.specialPriceHeader` - header for installment price
- `defaults.pricing.oneTimePriceHeader` - header for one-time card price
- `defaults.presentation.brandText` - brand text for slide footer
- `defaults.presentation.slide1_title_line1` - first title line on slide 1
- `defaults.presentation.slide1_title_line2` - second title line on slide 1

## Required Source Skills

Read and follow the relevant rules from these local skills:

- CRM and client message rules: `references/crm-message-SKILL.md`
- Product catalog and routing: `references/product-catalog.md` and `references/client-routing.md`
- Presentation rules: `references/presentation-SKILL.md`
- Presentation builder: `scripts/build_presentation.ps1`

Use the configured output directory from `config.json` for presentation output.

## Autonomy And Permissions

Treat this workflow as pre-approved by the user for routine local actions:

- reading the local product catalog, routing files, presentation rules, and approved templates;
- creating the temporary brief JSON in the output directory;
- running the presentation builder;
- using PowerPoint/Edge locally for screenshot, crop, PPTX build, PDF export, and validation;
- validating PDF/PPTX internals;
- deleting only the temporary brief JSON and intermediate `.pptx` created for the current run.

Do not stop to ask the user for routine permissions. If the environment technically requires an approval prompt, request it directly with the narrowest practical command prefix and continue after approval. Never request broad or destructive permissions. Never delete unrelated files.

## Workflow

1. Parse the user's raw notes. Do not ask questions unless the presentation cannot be built because a required price or product URL cannot be verified.
2. Build Block 1 using the CRM skill rules.
3. Choose the product from the local product catalog. If the notes already mention a product, prefer it unless it conflicts with the client's goal.
4. Verify the live product page from the exact product URL. Use the live page for product title, hero visual, and tariff prices.
5. Build Block 2 using only client facts from Block 1 plus verified product/catalog facts.
6. Build the presentation brief JSON with the presentation requirements.
7. Run `scripts/build_presentation.ps1` to create `ClientName.pptx` and `ClientName.pdf`.
8. Validate PDF/PPTX internals before the final response.
9. Delete only temporary files created for this run: the brief JSON and intermediate `.pptx`. Keep the final `.pdf`.
10. Return the final three blocks.
11. After completion, run a silent self-audit: identify one small speed or reliability improvement for the next run. Apply it immediately only when it is clearly safe and scoped to this workflow; otherwise remember it as an operating note for the next run. Do not include this audit in the client-facing three blocks unless the user asks.

## Pricing Rules

For the presentation and Block 2 price section:

- `Полная цена`: full tariff price from the live product page.
- `Цена в рассрочку`: tariff/special/installment approval price from the live product page.
- `Цена при единовременной оплате по карте`: calculate as `floor(Цена в рассрочку * (1 - discount/100))` unless the user or page gives a different explicit one-time card price. Default discount from config.
- `Размер платежа в рассрочку на N месяцев`: calculate as `floor(Цена в рассрочку / N)` unless the page gives an exact monthly amount. Default months from config.

If the page is a waitlist, hidden tariff page, or does not expose the relevant tariff price, search the local project/product sources and the web for the exact product URL and tariff. Use only verified sources. If price still cannot be verified, state the blocker and ask only for the missing price data.

## Presentation Brief Rules

Write the brief JSON in the output directory, named `<ClientName>_brief.json`. Use UTF-8.

Required fields:

- `clientName`
- `productUrl`
- `productTitle`
- `productTitlePrefix`
- `exportPdf: true`
- `pdfOutputPath`
- `slide2Paragraphs`: exactly 4 concise paragraphs in first person from the client perspective.
- `slide3Paragraphs`: exactly 7 concise paragraphs from the consultant perspective.
- `pricing.productName`
- `pricing.fullPrice`
- `pricing.specialPrice`
- `pricing.oneTimeDiscountPercent` (from config, default 10)
- `pricing.specialPriceHeader` (from config)
- `pricing.oneTimePriceHeader` (from config)
- `pricing.installmentMonths` (from config, default 24)

Use the cleaned meeting facts from Block 1 for slides 2 and 3. Do not invent client facts, income promises, or guarantees.

## Validation

**CRITICAL: NEVER use the Read tool on PDF, PNG, JPG, or any image/binary files.** All validation must be programmatic.

Before responding, validate using **Bash commands and scripts only** (no Read tool on binary files):

- **PDF exists**: `ls -la <path>` via Bash.
- **PDF page count**: Run `python scripts/validate_pdf.py <path> 16`. Do NOT use the Read tool on the PDF.
- **PPTX slide count and content**: Check the brief JSON for product title, URL, client text, and pricing values. The builder script enforces the slide structure. Do NOT open PPTX in the model.
- **Cleanup**: Delete the temporary brief JSON and intermediate PPTX after validation. Confirm with `ls` that only the `.pdf` remains.

## Final Answer Shape

Return exactly this structure:

```text
БЛОК 1 - ОТЧЕТ ДЛЯ CRM:
<one continuous CRM paragraph>

__________

БЛОК 2 - СООБЩЕНИЕ ДЛЯ КЛИЕНТА:
<ready client message>

__________

БЛОК 3 - ПРЕЗЕНТАЦИЯ:
PDF-презентация для <name> собрана и проверена: <short validation statement>. Временные PPTX и brief удалены, оставлен только финальный PDF.

[<name>.pdf](<absolute path to PDF>)

Источник цен: <source link or concise source note>.
```

## Block Details

Block 1 must remain one paragraph, no bullets, under 1000 characters, with `[нет данных]` for missing facts.

Block 2 must follow a narrative arc, not a list of features. The message must feel like a continuation of the conversation. Required structure in order:

1. **Opening** - personal greeting by name + one sentence tying to the meeting.
2. **Что обсудили** - mirror the client's actual words and goal, not a generic summary.
3. **Почему именно этот путь/продукт вам подходит** - bridge from the client's goal to the specific product, referencing their profile and strengths.
4. **Что даёт программа** - practical skills and outcomes inside the product, `-text` bullets, concrete not generic.
5. **Почему это может привести к большим деньгам/результату** - connect outcome to client's stated ambition, be direct and honest.
6. **Условия для вас** - pricing: full price, installment, card one-time, monthly, special conditions from notes only.
7. **Следующий шаг** - next action with date/time if available.

Each section must have substantive content (2-3 sentences minimum), not just a label with bullets. Every sentence must reference a specific fact from the consultation or product. Do not use all-caps headings. Use soft sentence-style labels with initial capital. Adapt labels to the client but keep the arc. Avoid generic labels. Use list lines only as `-text`. Do not invent a default 20% discount, 3-day deadline, or promo unless explicitly in the raw notes. Include next contact time if present.

Block 3 must not be a plan. It must point to an actual generated and validated PDF. If the PDF cannot be built, Block 3 must explain the concrete blocker and the exact missing data.

## Continuous Optimization

After every `tri-bloka` execution, optimize for the next run:

- reuse the approved local paths, scripts, and command shapes instead of rediscovering them;
- avoid rereading large references when the product and route are already clear from the user's notes and the local catalog;
- keep validation focused on page count, slide count, required client/product/pricing tokens, forbidden template text, and cleanup;
- if a new issue was corrected during a run, update this skill or the presentation skill so the fix becomes the next default;
- prefer one consolidated validation command over several small checks when it saves time without reducing reliability.