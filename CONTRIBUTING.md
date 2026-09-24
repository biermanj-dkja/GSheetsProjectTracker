# Contributing

Thanks for your interest! This is a small, single-file project, so the process
is light.

## Guiding principle

**Add automation, not human-facing complexity.** A change should make the
tracker do more on its own without asking the person using it to learn new
steps, columns, or buttons. If a feature needs a new input column, it probably
belongs behind a config option that is off by default.

## Reporting bugs / asking for features

Open an issue using the templates. For bugs, please include:

- what you did, what you expected, and what happened
- the exact wording of any error (Apps Script errors appear as a red banner or
  in **Extensions → Apps Script → Executions**)
- **no real project data**: use made-up project names in screenshots

## Making changes

1. Fork the repo and create a branch.
2. Edit `src/Code.gs`. Keep the heavy commenting style: every function should
   say what it does and why.
3. Test in a **copy** of a tracker spreadsheet:
   - run **Project Tools → One-time setup / repair**
   - add a project, log a few entries, log a "complete" entry
   - run **Close completed projects** and check STATUS, COMPLETED and Stats
4. Update `CHANGELOG.md` and, if behaviour changed, `README.md`.
5. Open a pull request.

## Using clasp (optional)

If you prefer editing locally, [clasp](https://github.com/google/clasp) works
with this layout. Copy `.clasp.json.example` to `.clasp.json`, fill in your
script ID, then `clasp push`. `.clasp.json` is git-ignored so your script ID
stays private.
