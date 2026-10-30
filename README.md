# IREx Data Gathering

How IREx collects Rehab Staff research activity for its KPI Review.

- **Collection Workbook macros:** `src/IRExCollection.bas`, imported into the Main Database. How to install and run them, and how to set up the reminder emails: [docs/collection-workbook.md](docs/collection-workbook.md).
- **Tests:** [tests/README.md](tests/README.md).
- **Design:** [GitHub issue #1](https://github.com/zkaiiii04/IREx-Data-Gathering/issues/1), [docs/discovery-notes.md](docs/discovery-notes.md), [GLOSSARY.md](GLOSSARY.md).

## Claude Code skills

`.claude/skills/` holds the published skills from [mattpocock/skills](https://github.com/mattpocock/skills) (commit `5beb35b`, MIT, see `.claude/skills/LICENSE-mattpocock-skills`). Claude Code loads them automatically in this repo; invoke one with `/<name>`, e.g. `/grill-me` or `/tdd`.

- `ask-matt`
- `code-review`
- `codebase-design`
- `diagnosing-bugs`
- `domain-modeling`
- `grill-me`
- `grill-with-docs`
- `grilling`
- `handoff`
- `implement-spec`
- `implement`
- `improve-codebase-architecture`
- `pr`
- `prototype`
- `research`
- `retro`
- `setup-matt-pocock-skills`
- `tdd`
- `teach`
- `to-questionnaire`
- `to-spec`
- `to-tickets`
- `triage`
- `wait-what`
- `wayfinder`
- `wizard`
- `writing-for-agents`
