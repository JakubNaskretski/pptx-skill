# Describe a slide template

Paste this prompt into a vision-capable LLM (Claude, ChatGPT, Gemini,
…) along with an image of the slide (the `slide_NN.png` preview, or
a screenshot if no preview exists). Paste the returned YAML into the
matching `slide_NN.yaml` file in
`authoring/workspace/decks/<deck>/slides/`.

Only describe what makes the **layout** and **intent** of this slide
distinctive. The downstream agent picks this template when it needs
a slide of this *shape* — its specific text and images will be
replaced at compose time.

---

## Output schema

Return YAML in exactly this shape — nothing else, no commentary:

```yaml
intent: ""            # one sentence: what this slide is FOR (under 20 words)
feel: ""              # formal | punchy | data-dense | warm | clinical | celebratory
suitable_for: []      # 1-4 tags from the controlled vocab below
notes: ""             # optional human reviewer note
interpretation: ""    # speculative observations — see "interpretation" below
```

Controlled vocab for `suitable_for`:
`opener`, `section_divider`, `content`, `data`, `quote`, `closing`,
`product`, `team`.

The `layout`, `slots`, and structural fields are auto-filled by
`ingest`. You do **not** need to write them. If the inferred `layout`
string is wrong, you may edit it directly in the YAML file — but
do not include it in the LLM response.

---

## `interpretation` — your speculative observations

The strict fields above are deliberately constrained to keep template
selection machine-actionable. This means there's no place for soft
observations: design rationale you infer, audience guesses, things
you noticed beyond what `intent`/`feel` capture. Put those in
`interpretation`.

Examples of what belongs here:

- "Reads as a Q4 retrospective opener — the yellow callout is more
  celebratory than warning, and the layout matches an annual-review
  pattern more than a regular content slide."
- "Likely from a sales pitch deck — the half-slide product mockup +
  customer-logo strip is a common B2B template."
- "Visually similar to slides 3 and 5 in this deck; appears to be a
  recurring 'chapter intro' pattern."

Rules for `interpretation`:

- **Mark uncertainty.** "Reads as", "appears", "looks like", "may be".
- **Don't repeat the strict fields.** If it belongs in `intent` /
  `feel` / `suitable_for`, put it there.
- **One short paragraph of plain prose**, no markdown.
- **Empty is fine.** Leave as `""` if you have nothing speculative.

The agent reads `interpretation` as informational context — it never
filters or routes on this field. Distinct from `notes`, which is
human reviewer free-form.

---

## Worked examples

### Example 1 — Hero opener

> Big bold title on the left, large product photo filling the right
> half of the slide.

```yaml
intent: "Hero opener with single bold claim alongside a product image"
feel: punchy
suitable_for: [opener, product]
notes: ""
interpretation: ""
```

### Example 2 — Data dashboard

> Title at top, four-quadrant grid of small charts below.

```yaml
intent: "At-a-glance dashboard across four metrics"
feel: data-dense
suitable_for: [data, content]
notes: ""
interpretation: "Layout reads as a quarterly business review — KPI tiles in even quadrants suggests a status snapshot rather than a deep-dive."
```

### Example 3 — Pull quote

> Single large centered quotation, small attribution beneath.

```yaml
intent: "Standalone customer or leader quote"
feel: warm
suitable_for: [quote, section_divider]
notes: ""
interpretation: ""
```

### Example 4 — Section divider

> Solid color background, large section number on the left,
> short section title beside it. No other content.

```yaml
intent: "Major section transition with number and title"
feel: formal
suitable_for: [section_divider]
notes: ""
interpretation: ""
```

---

## Rules

- Describe **purpose**, not content. "Showing a product" not
  "showing the X-1000".
- `intent` is one sentence, under 20 words.
- Only use enum values listed. Extend the enum only by editing this
  file in the authoring repo — do not invent values inline.
- `interpretation` carries soft observations (see section above).
  Empty is acceptable. Never use it to assert facts.
- Output ONLY the YAML. No preamble, no closing remarks.
