# IREx Data Gathering

Designing and building how IREx collects Rehab Staff research activity for its KPI Review. The Collection Workbook is built (VBA in `src/`, tests in `tests/`); the Form phase is still a design interview.

- Terms: [GLOSSARY.md](GLOSSARY.md). Use its terms in conversation and in docs.
- Build and test: `python3 tests/run_in_libreoffice.py` runs the VBA tests headless; Excel on Windows is the reference run. See [tests/README.md](tests/README.md). User guide: [docs/collection-workbook.md](docs/collection-workbook.md).
- Design interview state, decisions and open questions: [docs/discovery-notes.md](docs/discovery-notes.md).
- Resuming the interview in a new session: read [docs/handoff.md](docs/handoff.md) first.

## Agent skills

### Issue tracker

Issues and specs are GitHub issues on zkaiiii04/IREx-Data-Gathering. See `docs/agents/issue-tracker.md`.

### Triage labels

The five default labels (needs-triage, needs-info, ready-for-agent, ready-for-human, wontfix). See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: `GLOSSARY.md` and `docs/adr/` at the repo root. See `docs/agents/domain.md`.
