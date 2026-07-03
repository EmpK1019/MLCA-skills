# Slide Patterns — 参考资源库

以下 kind 是渲染器支持的布局类型参考。选择 kind 时以内容表达效果为准，不必强制使用所有类型，也不必避免相邻页使用同一 kind。如果连续两页都适合用 `grid`，那就用 `grid`。

## Core Slide Kinds

### `cover`

Use for the opening slide.

- dark or high-contrast background
- one clear title
- short subtitle
- minimal secondary text
- large visual atmosphere, not dense content

### `agenda`

Use when the deck benefits from orientation.

- 3-6 sections
- number each section
- keep wording short

### `section`

Use to reset the rhythm between chapters.

- oversized section label
- 1-line takeaway or promise
- strong background treatment

### `bullets`

Use only when the message truly is a short list.

- 3-5 bullets
- short lines
- one supporting callout or side note
- avoid paragraphs

### `two-column`

Use for explanation plus evidence.

- left: argument / text
- right: supporting blocks, checklist, examples, or mini-stats

### `stats`

Use when the slide should quantify.

- 3-6 metric cards
- big numbers
- short labels
- one sentence of interpretation

### `grid`

Use for capability maps, tool sets, pillars, or frameworks.

- 2x2, 2x3, or 3x2 card layouts
- each card: title + short descriptor

### `comparison`

Use for tradeoffs or alternatives.

- side-by-side columns
- consistent row labels
- highlight the preferred option

### `timeline`

Use for lifecycle, roadmap, or phased process.

- 4-6 stages
- each stage needs a short label and 1-3 supporting lines

### `quote`

Use for a single strong idea, testimony, or principle.

- one hero quote
- short attribution
- generous whitespace

### `closing`

Use for summary and next steps.

- 1-line closing statement
- 2-3 takeaways or actions
- stronger visual treatment than mid-deck slides

### `career-journey`

Use for background-to-transition narratives.

- left: domain / motivation bullets
- right: 3-stage vertical journey
- best for interview or founder background decks

### `process-flow`

Use for methodology, workflow, or operating model.

- 4-6 horizontal steps
- each step should show both explanation and outputs
- works well for product process, service delivery, and governance flows

### `capability-matrix`

Use for skill models or capability breakdowns.

- left: capability summary / score view
- right: stacked capability rows with tools and outputs

### `market-driver-stack`

Use for business context slides.

- left: narrative context
- right: top-down pressure stack ending in solution framing

### `persona-cards`

Use for user segmentation.

- 3-4 persona cards
- responsibilities + pain points + jobs-to-be-done

### `problem-cards`

Use for structured problem definition.

- 3 columns
- each column: problem, impact, example

### `architecture-stack`

Use for layered platform or system architecture.

- 2-4 layers on the left
- shared services or deployment notes on the right

### `feature-grid`

Use for 2x2 or 2x3 functional module overviews.

- strong module cards
- value tag or business outcome per card

### `swimlane`

Use for actor-by-phase workflows.

- actors in rows
- phases in columns
- emphasize handoffs and checkpoints

### `hub-spoke`

Use for platform ecosystem or stakeholder network slides.

- one central hub
- up to 5 surrounding roles
- emphasize the relationship between the center and its participants

### `metrics-dashboard`

Use for KPI reporting.

- 4 KPI cards
- one growth panel
- one North Star block

### `roadmap-phases`

Use for phased delivery plans.

- 3 main phases
- milestone per phase
- separate risks/dependencies strip

### `strengths-sidebar`

Use for candidate / founder / team advantage slides.

- 2x2 strengths grid
- dark sidebar with certifications, stack, or profile summary

### `free-layout`

Use when the slide needs a custom composition that should not be forced into a preset archetype.

- define `elements[]` with absolute positions
- use for hero compositions, mixed media layouts, or one-off showcase pages
- disable footer chrome with `pageChrome: false` when the composition needs a clean canvas
- keep theme tokens consistent even when the layout is fully custom

### `value-closing`

Use for the final “why me / why this matters / what value I bring” slide.

- centered closing statement
- 2-4 value cards or pillars
- lighter background than the cover
- should feel conclusive, calm, and premium

### `sources`

Use when the user asked for citations or the deck contains factual claims.

- collect labels and URLs from prior slides
- group duplicates
- prefer short readable source labels

## Pattern Selection Heuristics

- If the slide answers “what is this?” -> `bullets` or `two-column`
- If it answers “how does it work?” -> `timeline`
- If it answers “how do options differ?” -> `comparison`
- If it answers “what matters most numerically?” -> `stats`
- If it answers “what are the pillars?” -> `grid`
- If it changes chapters -> `section`
- If preset slide kinds would distort the message -> `free-layout`

## Anti-Patterns

- too many bullet slides in a row
- dense paragraphs
- repeated identical card layout on every slide
- putting all links only at the very end when slide-level references are needed
- title text that repeats the section name without saying anything new

## New Slide Kinds

### `kpi-dashboard`

Use for quantitative performance dashboards with multiple KPI metrics.

- 3-6 KPI cards with big numbers and short labels
- each card has a category-color left accent border
- optional bottom row for aggregate/ecosystem metrics
- one sentence of interpretation per KPI group

### `layered-architecture`

Use for multi-layer platform or system architecture with distinct layer colors.

- 4-5 horizontal layers stacked vertically (e.g. End→Edge→Cloud→Data→Intelligence)
- each layer has a distinct accent color left border and tint background
- right sidebar for cross-cutting concerns (security, integration)
- bottom flow bar showing data/process flow steps

### `timeline-dots`

Use for policy evolution, milestone tracking, or phased history with dot markers.

- horizontal timeline line with progress fill
- 4-6 dot markers with date labels above and description below
- active/past dots use accent color fill, future dots use muted color
- bottom summary box with accent left border

### `capability-radar`

Use for skill/capability assessment with radar chart visualization.

- left panel: radar chart (Chart.js) showing current vs target scores
- right panel: stacked capability cards with left border in category color
- each card shows: icon, title, tools (tag pills), outputs
- chart uses dominant/accent fills, category-color point markers

## Pattern Selection Heuristics (Extended)

- If the slide answers "what are the KPIs?" -> `kpi-dashboard` or `stats`
- If the slide answers "what is the system architecture?" -> `layered-architecture` or `architecture-stack`
- If the slide answers "how did policy/standards evolve over time?" -> `timeline-dots` or `timeline`
- If the slide answers "what are the capability gaps?" -> `capability-radar` or `capability-matrix`

## Theme-Aware Rendering Notes

Each theme preset changes how slide kinds are rendered. Key differences:

### McKinsey Rendering

- **Cover**: Deep blue `#002060` background, accent line, centered title, footer bar
- **Content pages**: White bg, 80px header bar with `border-bottom: 2px solid #002060`
- **Process-flow**: Step cards with gradient top borders (`#002060 → #004b8d → #0077b6 → #00A3E0 → #4CAF50`)
- **Capability-radar**: Left radar chart (blue fill `rgba(0,32,96,0.2)`, cyan points), right stacked cards with left borders
- **Stats/KPI**: White metric cards, no accent bar, big numbers in dominant color
- **Timeline**: Vertical timeline with cyan dot markers and blue border-left cards
- **Bullets**: Cyan `#00A3E0` bullet dots, blue icon accents

### Dark-Tech Rendering

- **Cover**: Navy `#0A2463` bg, 24px green accent bar, category label pill, declaration info plate
- **Content pages**: Navy bg, left accent bar (24px green + 8px navy), title with green border-left
- **All cards**: `rgba(30,58,95,0.6)` background, 1px category-color border, 8px border-radius
- **Layered-architecture**: 5 layers with distinct accent color left borders (green→blue→amber→purple→red)
- **Timeline-dots**: Horizontal line with green progress fill, navy-fill + green-border dots
- **KPI-dashboard**: Stat boxes with category-color left border, green/blue/amber tint backgrounds
- **Capability-radar**: Dual-dataset radar (amber=current, green=target), navy bg
- **Process flow**: Bottom rounded pill bar with flow steps

### Genspark Rendering

- **Cover**: Gradient background (`#667eea → #764ba2`), centered title, tag pills
- **Content pages**: Light `#F8FAFC` bg, card grid with category-color top borders
- **Stats/KPI**: White metric cards with category-color left accent strip
- **Feature-grid**: 2x3 card grid, icon circles with category tint bg, tag pills per card
- **Timeline**: Horizontal dots with category-color fills, connecting line
- **All cards**: White bg, 8px border-radius, category-color top or left border
