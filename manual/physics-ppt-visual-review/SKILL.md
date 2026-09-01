---
name: physics-ppt-visual-review
description: Read-only visual regression review for original and normalized junior-high physics PPT page images. Use when this repository produces before/after PNG pages and needs a Passed, Review, or Blocked gate without modifying the PPTX.
---

# Physics PPT visual review

Review rendered evidence only. Never edit a PPTX, invoke a normalization/fix script, or reinterpret teaching content.

## Inputs

Use `ai-visual-review-request.json` when present. It maps every source page image to its normalized counterpart. If the packet is missing, use `review-manifest.json` and the corresponding original/normalized page-image directories.

Stop with `ReviewUnavailable` if pages cannot be paired exactly. Do not silently skip a page.

## Review

Inspect every page pair at full size. A contact sheet may help navigation but is not sufficient evidence.

Only judge regressions introduced by normalization:

- new wrapping or changed line breaks;
- clipping, overlap, disappearance, or newly blank content;
- font fallback, tofu boxes, garbled symbols, or formula baseline damage;
- reduced readability from font size, weight, or color;
- unintended movement, resizing, cropping, reordering, or other structure change.

Do not judge or rewrite teaching content. Do not require layout modernization. Existing source issues are `KnownSourceIssue`, not normalization failures.

## Gate

- `Passed`: no introduced visual problem.
- `Review`: evidence is ambiguous or a source issue prevents a reliable comparison.
- `Blocked`: a new visual problem or unapproved structure change is visible.

Write one result per page plus a file-level aggregate. Conform to [references/review-result.schema.json](references/review-result.schema.json). A file is `Blocked` if any page is blocked and `Review` if none is blocked but any page needs review.

Record the model name, review protocol version, image paths, and concise visual evidence. Do not include speculative fixes.
