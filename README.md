# M-LCA Skills

Private skill marketplace source for M-LCA workflows.

This repository follows the same outer structure as `EmpK1019/cortex-skills-main` and publishes selected M-LCA skills through `marketplace.json`.

## Structure

```text
.
|-- AGENTS.md
|-- CONTRIBUTING.md
|-- README.md
|-- marketplace.json
|-- docs/
`-- skills/
```

## Marketplace Index

`marketplace.json` is the source index for skills published from this repository. Add a skill entry only after the matching folder exists under `skills/`.

Current imported skills:

- `docx`
- `pdf`
- `power-grid-lca-decomposition`
- `pptx`
- `skill-creator`
- `xlsx`

## Skill Layout

Each installable skill should live under `skills/<skill-id>/`:

```text
skills/<skill-id>/
|-- SKILL.md
|-- references/
`-- scripts/
```

Only `SKILL.md` is required. Use `references/` for larger guidance files and `scripts/` for helper code.

## Install A Skill

After a skill is committed and pushed, install it from Codex with:

```powershell
python "$env:CODEX_HOME\skills\.system\skill-installer\scripts\install-skill-from-github.py" --repo EmpK1019/M-LCA-skills --path skills/<skill-id>
```

Restart Codex after installing a skill.
