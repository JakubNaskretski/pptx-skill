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
notes: ""             # optional
```

Controlled vocab for `suitable_for`:
`opener`, `section_divider`, `content`, `data`, `quote`, `closing`,
`product`, `team`.

The `layout`, `slots`, and structural fields are auto-filled by
`ingest`. You do **not** need to write them. If the inferred `layout`
string is wrong, you may edit it directly in the YAML file — but
do not include it in the LLM response.

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
```

### Example 2 — Data dashboard

> Title at top, four-quadrant grid of small charts below.

```yaml
intent: "At-a-glance dashboard across four metrics"
feel: data-dense
suitable_for: [data, content]
notes: ""
```

### Example 3 — Pull quote

> Single large centered quotation, small attribution beneath.

```yaml
intent: "Standalone customer or leader quote"
feel: warm
suitable_for: [quote, section_divider]
notes: ""
```

### Example 4 — Section divider

> Solid color background, large section number on the left,
> short section title beside it. No other content.

```yaml
intent: "Major section transition with number and title"
feel: formal
suitable_for: [section_divider]
notes: ""
```

---

## Rules

- Describe **purpose**, not content. "Showing a product" not
  "showing the X-1000".
- `intent` is one sentence, under 20 words.
- Only use enum values listed. Extend the enum only by editing this
  file in the authoring repo — do not invent values inline.
- Output ONLY the YAML. No preamble, no closing remarks.
