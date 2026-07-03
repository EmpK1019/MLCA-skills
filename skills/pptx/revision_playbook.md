# Revision Playbook

This file defines how to revise a previously generated deck without rebuilding the whole logic from scratch.

## Core Rule

When a deck was already generated, prefer revising the existing `deck-spec.json` and rerendering.

Do not invent a brand new slide sequence unless the user changed the objective.

## Preferred Revision Input Order

1. a previously generated `deck-spec.json`
2. a previously generated `.pptx`
3. the assistant's earlier deck description in message history

If both a spec and a pptx exist, trust the spec as the editable source of truth.

## Revision Types

### Visual-only revision

Examples:

- make it more premium
- reduce visual clutter
- use a darker tone
- switch to investor style

Actions:

- keep slide order
- keep main messages
- update theme tokens
- optionally switch some slide kinds for better fit

### Structural revision

Examples:

- shorten to 6 slides
- add one section on risks
- turn slide 4 into a comparison

Actions:

- update the slide list in `deck-spec.json`
- preserve surviving slides and sources
- rerender

### Content revision

Examples:

- add traceable references
- sharpen wording
- replace metrics

Actions:

- update only affected slide payloads
- keep theme and unaffected slide kinds stable

## Suggested Workflow

1. read the existing spec if available
2. summarize the requested delta
3. patch the spec
4. rerender `.pptx`
5. overwrite or regenerate the companion `deck-spec.json`

## Good Practice

- keep `specOutputFile` stable so the next round can reuse the same filename
- preserve source labels when possible
- do not drop citations accidentally during visual revisions
- when changing slide kinds, preserve the slide's core message
