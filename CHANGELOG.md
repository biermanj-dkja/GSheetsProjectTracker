# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and the project uses [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-09-24

### Added
- **Project Tools** menu.
- **Close completed projects**: moves every project whose newest log entry is
  "complete" (or done / closed / cancelled) into the COMPLETED archive and
  removes it from STATUS, with no blank rows left behind.
- **Start new year**: copies the tracker for the next year, carrying over open
  projects with their history and starting with an empty archive.
- **One-time setup / repair**: builds missing tabs, self-expanding lookup
  formulas, date colours (green / yellow / red, grey for on-hold), an A–Z
  project dropdown, and the Stats tab.
- **Stats tab**: open/closed/stale counts, averages, update pace, and a
  per-project table. All live formulas.
- Automatic date on new log entries.
- Automatic priority sort of the STATUS list.
- Optional self-updating graduation-year helper tab.
