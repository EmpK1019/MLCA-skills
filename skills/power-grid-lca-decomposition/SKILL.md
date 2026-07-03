---
name: Power Grid Asset LCA Material Decomposition
description: Methodology and principles for decomposing power-sector asset inventories
  (generation, transmission, and distribution) into LCA material flows and matching
  background datasets. Applies when the user provides an equipment register plus a
  knowledge-base file of arbitrary structure.
source: nudged
created_at: 2026-05-26 01:19:38.388000+00:00
updated_at: 2026-05-26 01:26:46.086000+00:00
tags:
- lca
- power-grid
- decomposition
- upr
- transmission
- generation
- materials
id: power-grid-lca-decomposition
category: contextual
distribution_source: local_import
order: 108
version: '1.0'
trigger_hints: []
allowed_tools: []
output_contract: {}
scripts: []
form: []
prompt_template: ''
icon: skill
show_in_template_bar: false
show_in_drawer: true
paths: []
model: ''
effort: ''
publisher_name: ''
---

# Power Grid Asset LCA Material Decomposition

## Scope

Decompose equipment registers from **power generation, transmission, and distribution** projects into LCA unit-process (UPR) material flows, then match each material to a background dataset (BG ID).

Typical inputs:
- **Equipment register** — Excel/CSV, field structure varies by project
- **Knowledge-base (KB) file** — may be Excel, Word, PDF, or a table screenshot; structure is not standardised

---

## Step 1 — Understand the Input Files

### 1.1 Equipment Register
Scan every column and identify:

| What to find | Typical column names |
|---|---|
| Equipment / material name | 设备名称, 物料名称, Description, Item |
| Specification / model | 规格型号, Specification, Model, Rating |
| Quantity | 数量, Qty, Amount |
| Unit | 单位, Unit |
| Work package / batch | 施工包, Package, Phase, Zone |

Watch for merged cells, multi-header rows, or multiple sheets — flatten before processing.

### 1.2 Knowledge-Base File
The KB structure cannot be assumed. It may list one device per row, one per section, or mix prose and tables. For each device entry, extract:

1. **Trigger keywords** — terms that match against register names
2. **Material components** — name + mass fraction (fractions must sum to 1.0)
3. **Measurement unit** — must align with the register unit
4. **Unit mass or conversion factor** — how to convert register quantity → kg

If the KB provides a range rather than a single value, use the midpoint and flag it as an estimate.

---

## Step 2 — Build Decomposition Rules

### 2.1 Keyword Matching
- A register row is matched when its name **contains** a KB keyword (substring match, case-insensitive)
- If multiple rules match, prefer the **longest / most specific** keyword
- Log all unmatched rows in a separate list for manual review

### 2.2 Unit Conversion
| Register unit | Approach |
|---|---|
| Same as KB (e.g. kg) | No conversion needed |
| t → kg | factor = 1000 |
| pieces, sets, units | Need unit mass (kg/piece) from KB |
| m (linear) | Need linear density (kg/m) from KB |

### 2.3 Spec-Dependent Unit Mass
When the same device type has significantly different masses depending on its rating or model code:

1. Collect reference masses per specification from the KB or industry standards
2. Build a branching lookup (if/elif chain) keyed on substrings in the specification field
3. Pre-process spec strings: `.upper().strip()` to avoid case/space mismatches
4. Always include a catch-all default at the end; note "specification not recognised, default applied" in the description field

### 2.4 Material Fraction Validation
- Fractions must sum to 1.0 — verify before use
- If the KB fractions do not sum to 100 %, check for unlisted materials or rounding; do not silently ignore the gap
- Each material's absolute mass = total device mass × fraction

### 2.5 Unit-Mass Attribution (mandatory distinction)
In every item description, state the source of the unit mass:

| Source | Wording to use |
|---|---|
| KB-provided value | "Unit mass from knowledge base: X kg/unit" |
| Derived from project totals | "Unit mass derived from project data (total mass ÷ count)" |

Never write "assumed unit mass" — it obscures the provenance.

---

## Step 3 — UPR Data-Item Naming

Use a hierarchical name that reflects the project structure. Typical format (adjust delimiter to match your LCA tool's import format):

```
{Activity category},{Device category},{Package/batch}-{Register name},{Material name}
```

Rules:
- Keep the work-package identifier to disambiguate same-named items from different batches
- Use the KB's standard material name, not the register's raw text
- Delimiter choice must match the target LCA software import convention

---

## Step 4 — Background Dataset Matching

### 4.1 Matching Strategy
- Match on **material name keywords** — do not hard-code dataset keys across projects (keys may change when the database is updated)
- Re-run `search_catalog` for each new project to verify keys are still valid
- Prefer datasets from the **same region** (e.g. CN) and **same production route** as the actual material

### 4.2 Source Priority
HiQLCD (domestic data) → Ecoinvent (international fallback), CUT_OFF system boundary.

### 4.3 No Dataset Available
Some physical quantities (e.g. earthwork volume in m³, construction labour hours) have no matching LCA dataset:
- **Keep the row** in the UPR; leave the BG ID column blank
- Add a note in the description: "No matching LCA background dataset; included for completeness"
- Do not delete the row; do not substitute an unrelated dataset

---

## Step 5 — Pipeline Design

Separate the pipeline into discrete, independently debuggable modules:

```
1. Register ingestion + rule matching  →  intermediate JSON
2. Intermediate JSON  →  UPR Excel  (via LCA-tool plugin script)
3. BG ID fill  (separate script reading the generated Excel)
4. Original-quantity column write  (see critical note below)
5. File sync  (copy plugin output back to project directory)
```

Run all steps in sequence from a single entry point.

### Critical — Writing the Original-Quantity Column
The row order in the UPR Excel is determined by the **category sort order** applied by the plugin script, which differs from the JSON order. If you use item name as a dict key to write column values, rows sharing the same name (e.g. two conductor entries in the same package) will overwrite each other.

**Correct approach**: replicate the plugin's sort logic, then write values by **positional index**, not by name lookup.

### Zero-Quantity Rows
Register rows with quantity = 0 produce child rows with zero mass. **Keep them** — they preserve a complete audit trail of the register. Do not delete or skip.

### File Sync
If the plugin writes to a different directory than the one the user opens, add an automatic copy step at the end of the pipeline. Otherwise the user will always see a stale file.

---

## Step 6 — Item Description Format

Each UPR item description should contain these elements in order:

```
Source: {original register name}, Specification: {spec}. Decomposed using KB template [{KB template name}].
Original quantity: {qty} {unit}; {unit-mass attribution} {value} kg/{unit}.
{Material 1}: {fraction × 100}%, {Material 2}: {fraction × 100}%, …
```

Omit the specification clause if the register has no spec field.

---

## Step 7 — Visualisation Report (Recommended)

After generating the UPR, produce a self-contained HTML report (all data inline, no external files) with:

- **KPI bar**: total rows, matched, unmatched, BG coverage rate
- **Matched list**: searchable/filterable table with original name, rule triggered, material breakdown
- **Unmatched list**: for manual review
- **KB rule cards**: overview of all rules used

Inline the data as JSON so the file can be emailed or shared without dependencies.

---

## Common Pitfalls

| Symptom | Root Cause | Fix |
|---|---|---|
| Same-named rows show identical G-column values (last one wins) | Name-keyed dict overwrites on duplicate names | Write by positional index after replicating plugin sort order |
| Unit-mass wording misleads stakeholders | "Assumed" used for KB values | Distinguish "KB estimate" vs "project-derived" explicitly |
| BG-fill script reads stale Excel | Script defaults to project-dir file, not plugin output | Prefer plugin output path; fall back to project dir only if missing |
| Spec-matching function always hits default | Space or case difference in spec string | Pre-process: `.upper().strip()` before matching |
| f-string prints literal `{var}` | Double braces `{{var}}` inside f-string | Use `print("label:", var)` outside f-string |
| User sees old file after pipeline run | Plugin writes to one path, user opens another | Add `shutil.copy2(src, dst)` at pipeline end |
