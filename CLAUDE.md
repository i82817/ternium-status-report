# CLAUDE.md — Project Implementation & Status Report

## Purpose

This project generates SAP S/4HANA implementation status reports and governance artifacts as HTML documents. It is aimed at project managers, solution architects, and delivery leads tracking implementation progress, open points, risks, and milestone completion.

## Tech Stack

- **Frontend:** Vanilla HTML + CSS + JavaScript (no build step)
- **Output:** Static HTML files, optionally printable/exportable

## Project Structure

```
project-implementation-status-report/
├── CLAUDE.md
├── index.html          # Main status dashboard
├── reports/            # Generated status report HTML files
├── assets/             # CSS, logos, icons
└── .claude/
    └── skills/         # All available skills (copied from SKILL/)
```

## Active Skills

The following skills are available in `.claude/skills/` and should be used proactively:

| Skill | When to use |
|---|---|
| `sap-finance-solution-reviewer` | Reviewing scope, effort, ROM, proposals, open points |
| `sap-strategy-advisor` | SAP strategy context, Autonomous Enterprise framing |
| `sap-architecture-designer` | BTP architecture diagrams for solution documentation |
| `joule-ekx` | Querying SAP enterprise knowledge, best practices |
| `jwd-settings-features-coach` | JWD onboarding and coaching |
| `outlook-mac-email-draft` | Drafting status report emails |
| `teach-me-ai` | AI education and capability explanations |
| `automations-toolkit` | LeanIX automation integration |
| `btp-cli` | SAP BTP CLI commands |
| `joule-a2a-agent` | Building custom Joule agents |
| `ui5-best-practices-integration-cards` | UI5 Integration Cards in dashboards |
| `ui5-best-practices-tables` | Tables in HTML/UI5 reports |
| `accessibility` | Accessibility for HTML output |

## Primary Skill: sap-finance-solution-reviewer

This is the core skill for this project. Activate it with:
- `review this proposal` · `analyze this ROM` · `scope review`
- `red team` · `AI scan` · `historical benchmark`
- `effort review` · `executive summary` · `create PPT`
- `meeting minutes` · `discovery summary`

Responds in English, Brazilian Portuguese, and Spanish.

## Conventions

- All reports are static HTML files — no frameworks, no build step
- Use SAP Fiori Horizon color palette where applicable
- Reports should be print-friendly (CSS `@media print`)
- Date format in filenames: `YYYY-MM-DD`
