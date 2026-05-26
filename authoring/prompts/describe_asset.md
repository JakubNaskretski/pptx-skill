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

The agent reads only three things on each asset:

- `kind` — what kind of atom (already filled by ingest; you usually
  don't change it)
- `tags` — 1–4 short classification labels from a closed list
- `description` — one short, neutral sentence saying what's visible

Plus mechanical metadata (`width`, `height`, `aspect`, `colors_hex`)
that is computed from the binary at ingest — never your job.

**Do not invent fields.** Anything outside the schema below is ignored
or rejected by `validate`.

---

## Output schema

Return YAML in exactly this shape — nothing else, no commentary:

```yaml
kind: ""              # see "Kind" below
tags: []              # 1-4 labels from the tag list below
description: ""       # one neutral sentence, under 25 words
```

If you're refining an existing asset, **preserve the existing
description verbatim unless it is clearly inaccurate** — the user
explicitly does not want their previous prose rewritten on a
re-describe pass.

---

## Kind

The atom's structural type. Ingest seeds this; you usually keep it.

Raster pictures:

- `photo` — photographic content
- `icon` — small symbolic mark, low detail
- `logo` — brand mark or wordmark
- `illustration` — drawn / rendered artwork (not photographic, not iconic)
- `screenshot` — UI capture, diagram exported as raster, etc.

Vector:

- `vector` — SVG / EMF; DPI-independent, recolourable

Structured atoms (rare to redescribe — these come from ingest):

- `table` · `chart` · `callout` · `freeform` · `smartart`

---

## Tags

Pick **1 to 4** tags from the workspace tag vocabulary. Tags describe
what is **literally pictured**, not the rhetorical message of the slide
the asset was on. ("growth" is not a tag; "chart" is.)

Current vocabulary (workspace/tag_vocab.yaml):

- `people` — one or more humans visible
- `office` — indoor work environment with desks/chairs/etc.
- `laptop` — laptop computer visible
- `device` — phone, tablet, headset, or other consumer hardware
- `screen` — close-up of a display or monitor
- `hands` — hands visible without the rest of the body
- `document` — printed page, contract, form
- `chart` — bar/line/pie/etc. data visualization
- `logo` — brand mark or wordmark
- `abstract` — non-representational; texture, gradient, shapes
- `outdoor` — outdoors / natural light
- `nature` — plants, animals, landscapes
- `city` — urban environment, skylines, streets
- `workplace` — work setting other than a desk office (factory, lab,
  field site, kitchen, etc.)

If none fit, return an empty `tags: []` and the user can either add
a new tag to the vocab via `cli.py tag-vocab add <tag>` or leave it
untagged.

---

## Description

One short, neutral sentence saying what is literally visible.

- Bad: "A team celebrating a successful Q4 launch."
- Good: "Four people standing around a laptop in an open office."

Limits: under 25 words. No marketing language. No interpretation of
meaning. State what is there.

For structured atoms (table, chart, callout, freeform, smartart), the
description still describes the **visual content** — e.g. "Six-row
table comparing pricing tiers across three plans." The structural
data (rows/cols/series counts) is captured automatically; you only
write the human-readable summary.

---

## Example

For a photo of a person tending plants in a controlled enclosure:

```yaml
kind: photo
tags:
  - people
  - workplace
  - nature
description: Man standing among rows of small plants in a greenhouse-like enclosure.
```
