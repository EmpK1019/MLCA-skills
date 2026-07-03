# Citation Playbook

Use this guide when a deck includes factual claims, market numbers, benchmarks, standards, or any content the user asked to make traceable.

## When Citations Are Required

- the user explicitly asked for references or traceability
- the deck cites market size, growth rate, benchmark, policy, standard, or industry comparison
- the deck summarizes uploaded research material and should remain auditable

## Citation Placement

Do not dump all links only in the final reply. Prefer:

- relevant slide footer labels for the claims on that slide
- a final sources appendix for the full URLs

Good slide-level labels:

- `McKinsey 2024`
- `PMI Pulse 2023`
- `IEA 2025`
- `User interview notes`

Avoid putting raw long URLs directly into content slides unless necessary.

## Source Mapping

For each factual slide, keep a `sources` array in the internal plan.

Use simple objects like:

```json
[
  { "label": "McKinsey 2024", "url": "https://example.com/report" },
  { "label": "PMI Pulse 2023", "url": "https://example.com/pmi" }
]
```

If the source is user-provided rather than web-researched:

```json
[
  { "label": "User attachment: strategy-notes.pdf", "url": "" }
]
```

## Research Order

When you need web evidence, use:

1. Tavily
2. Brave
3. model-native search only if the first two are unavailable or insufficient

## Honesty Rules

- If a number is estimated or synthesized, mark it as an estimate.
- If a slide is conceptual and does not rely on external facts, you do not need forced citations.
- Never invent institutions, report names, or URLs.
