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

When choosing between assets, the agent reads:

- `subject` — what is literally pictured
- `depicts` — the *concept* this is about (the THING, not the appearance)
- `feel` — mood / aesthetic
- `composition` — where the focal subject sits in the frame
- `colors` — dominant palette
- `scope` — where the asset is reusable (client, industry, etc.)
- `suitable_for` — slide types this works on

Do not write marketing language. Do not interpret meaning beyond what
is visible. Describe what is there and quantifiable.

---

## Output schema

Return YAML in exactly this shape — nothing else, no commentary:

```yaml
kind: ""              # photo | icon | logo | illustration | screenshot
subject: ""           # one sentence, neutral, under 25 words
depicts: ""           # the THING this is about (concept). Optional if purely decorative.
feel: ""              # formal | warm | clinical | punchy | playful | minimal | dramatic
composition: ""       # centered | left-weighted | right-weighted | full-bleed | top-heavy | scattered
colors: []            # 1-3 plain color words
scope: []             # REQUIRED — see "Scope" below. Use [generic] if unsure.
suitable_for: []      # 1-4 tags from the controlled vocab below
notes: ""             # optional
```

Controlled vocab for `suitable_for`:
`team`, `hero`, `product`, `data`, `culture`, `event`, `abstract`,
`decorative`, `closing`, `quote`.

---

## `depicts` — what concept the asset is about

The most important field for retrieval beyond visual matching. The
`subject` field is the literal appearance; `depicts` is the meaning.

- A BPMN diagram of an internal sales-offer workflow
  → `depicts: "sales-offer preparation workflow"`
- A photo of two people shaking hands at a desk
  → `depicts: "business deal handshake"`
- A bank's official logo
  → `depicts: "ACME Bank brand mark"`
- An abstract gradient or decorative texture
  → `depicts: ""` (truly generic — empty is fine)

Keep it short: 1-5 words. Name the THING, not its visual treatment.

---

## `scope` — where this asset is reusable

A list of namespaced tokens telling the agent **when** this asset is
appropriate. **Required: at least one entry.** If you cannot tell, use
`[generic]` and add a note.

Recognized prefixes:

| Prefix | Meaning | Example |
|---|---|---|
| `generic` | Bare token — usable in any deck | `[generic]` |
| `client:<slug>` | Tied to a specific client/account | `[client:acme-bank]` |
| `industry:<slug>` | Tied to a sector | `[industry:finance]` |
| `product:<slug>` | Tied to a specific product/offering | `[product:salesforce]` |
| `program:<slug>` | Tied to an internal program/initiative | `[program:offer-process]` |
| `topic:<slug>` | Tied to a topic | `[topic:sustainability]` |

Use lowercase, dash-separated slugs. An asset can have multiple
scopes — e.g. a photo from an ACME Bank finance summit:
`[client:acme-bank, industry:finance]`.

**Rule of thumb**: if the image contains anything client-specific
(logo, branded color, named product, official document) it is NOT
generic. Mark with the appropriate `client:`, `product:`, or
`program:` scope. Only mark `generic` for assets that could appear
in a deck for any organization or topic.

---

## Worked examples

### Example 1 — Generic team photo

> A wide shot of four people gathered around a whiteboard in a sunlit
> office. Two are gesturing; one holds a marker.

```yaml
kind: photo
subject: "Four people collaborating at a whiteboard in a sunlit office"
depicts: "team collaboration"
feel: warm
composition: centered
colors: [warm white, navy, soft yellow]
scope: [generic]
suitable_for: [team, culture]
notes: ""
```

### Example 2 — Client-specific logo

> ACME Bank wordmark and stylized "A" mark, navy on white.

```yaml
kind: logo
subject: "ACME Bank wordmark with stylized 'A' mark, navy on white"
depicts: "ACME Bank brand mark"
feel: formal
composition: centered
colors: [navy, white]
scope: [client:acme-bank]
suitable_for: [closing, hero]
notes: ""
```

### Example 3 — Internal process diagram

> Swimlane BPMN activity diagram showing sales-offer preparation steps
> across five organizational lanes.

```yaml
kind: screenshot
subject: "Swimlane activity diagram with five lanes and labeled process boxes"
depicts: "sales-offer preparation workflow"
feel: formal
composition: full-bleed
colors: [white, gray, black]
scope: [program:offer-process]
suitable_for: [data, abstract]
notes: "Activity diagram from internal RFP tooling"
```

### Example 4 — Industry-fit stock photo

> Stock photo of a busy trading floor, multiple screens, people gesturing
> at monitors.

```yaml
kind: photo
subject: "Busy trading floor with screens and people gesturing at monitors"
depicts: "financial markets activity"
feel: dramatic
composition: scattered
colors: [navy, white, red]
scope: [industry:finance]
suitable_for: [hero, data]
notes: ""
```

### Example 5 — Product screenshot

> Salesforce dashboard UI showing pipeline metrics, opportunity tiles.

```yaml
kind: screenshot
subject: "Salesforce dashboard with pipeline metrics and opportunity tiles"
depicts: "Salesforce opportunity pipeline view"
feel: clinical
composition: full-bleed
colors: [blue, white, gray]
scope: [product:salesforce]
suitable_for: [product, data]
notes: ""
```

### Example 6 — Decorative

> A subtle warm-gradient background, no subject.

```yaml
kind: illustration
subject: "Warm orange-to-cream radial gradient, no subject"
depicts: ""
feel: minimal
composition: full-bleed
colors: [orange, cream]
scope: [generic]
suitable_for: [decorative]
notes: ""
```

---

## Rules

- Only use enum values listed for `kind`, `feel`, `composition`,
  `suitable_for`. If nothing fits a field, pick the closest enum
  value and explain the tension in `notes`.
- `subject` is one sentence, under 25 words, no interpretation.
- `depicts` is 1-5 words naming the *concept* the asset is about.
  Leave empty only for purely decorative assets.
- `scope` must have at least one entry. Default to `[generic]` only
  for assets that could appear in a deck for any organization or
  topic. Anything client-, product-, or program-specific gets the
  matching namespaced scope.
- `suitable_for` describes *slide types*, not *topics*. "team" is
  fine; "Q4 earnings deck" is too specific.
- Output ONLY the YAML. No preamble, no closing remarks, no
  surrounding code fence labels other than `yaml`.
