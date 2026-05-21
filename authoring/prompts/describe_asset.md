# Describe a visual asset

Paste this entire prompt into a vision-capable LLM (Claude, ChatGPT,
Gemini, …) along with the image. Ask the model to return YAML matching
the schema below. Paste the returned YAML into the matching
`<sha1>.yaml` file in `authoring/workspace/assets/`.

The description is consumed by a downstream agent **with no vision
capability**. It picks between assets purely by reading these fields.
Optimize for **retrieval**, not captioning.

---

## What the downstream agent compares

When choosing between visually similar assets, the agent reads:

- `subject` — what is literally pictured
- `feel` — mood / aesthetic
- `composition` — where the focal subject sits in the frame
- `colors` — dominant palette
- `suitable_for` — slide types this works on

Do not write marketing language. Do not interpret meaning. Describe
what is visible and quantifiable.

---

## Output schema

Return YAML in exactly this shape — nothing else, no commentary:

```yaml
kind: ""              # photo | icon | logo | illustration | screenshot
subject: ""           # one sentence, neutral, under 25 words
feel: ""              # formal | warm | clinical | punchy | playful | minimal | dramatic
composition: ""       # centered | left-weighted | right-weighted | full-bleed | top-heavy | scattered
colors: []            # 1-3 plain color words
suitable_for: []      # 1-4 tags from the controlled vocab below
notes: ""             # optional, only if something important doesn't fit
```

Controlled vocab for `suitable_for`:
`team`, `hero`, `product`, `data`, `culture`, `event`, `abstract`,
`decorative`, `closing`, `quote`.

---

## Worked examples

### Example 1 — Photo

> A wide shot of four people gathered around a whiteboard in a sunlit
> office. Two are gesturing; one holds a marker. Warm afternoon light
> through floor-to-ceiling windows.

```yaml
kind: photo
subject: "Four people collaborating at a whiteboard in a sunlit office"
feel: warm
composition: centered
colors: [warm white, navy, soft yellow]
suitable_for: [team, culture]
notes: ""
```

### Example 2 — Icon

> A simple line-art bar chart, three bars increasing in height, single
> color stroke, transparent background, 96 × 96 px.

```yaml
kind: icon
subject: "Three-bar ascending bar chart in line-art style"
feel: minimal
composition: centered
colors: [navy]
suitable_for: [data]
notes: ""
```

### Example 3 — Logo

> Stylized wordmark "ACME" in bold sans-serif, single-color black on
> transparent background.

```yaml
kind: logo
subject: "ACME company wordmark in bold sans-serif"
feel: formal
composition: centered
colors: [black]
suitable_for: [closing, hero]
notes: ""
```

### Example 4 — Illustration

> Flat-style illustration of a stylized rocket lifting off, geometric
> shapes, three-color palette, no shading, transparent background.

```yaml
kind: illustration
subject: "Flat-style rocket lifting off, geometric shapes"
feel: playful
composition: centered
colors: [coral, navy, cream]
suitable_for: [hero, abstract]
notes: ""
```

---

## Rules

- Only use enum values listed. If nothing fits a field, pick the
  closest enum value and explain the tension in `notes`.
- `subject` is one sentence, under 25 words, no interpretation.
- `suitable_for` describes *slide types*, not *topics*. "team" is
  fine; "Q4 earnings deck" is too specific.
- If the asset is purely decorative (a gradient, a divider line,
  a background texture), describe it neutrally and set
  `suitable_for: [decorative]`.
- Output ONLY the YAML. No preamble, no closing remarks, no
  surrounding code fence labels other than `yaml`.
