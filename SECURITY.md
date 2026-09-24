# Security Policy

## What this script can access

The script is **container-bound**: it only runs inside the spreadsheet you paste
it into. It reads and writes that spreadsheet, shows menus and dialogs, and
(only when you choose **Start new year**) makes a copy of that spreadsheet in
your Google Drive. It makes no network requests and sends no data anywhere.

## Keeping data out of the repo

- Never commit exported copies of a real tracker (`.xlsx`, `.csv`). The
  `.gitignore` blocks these by default.
- Never commit `.clasp.json` or `.clasprc.json`; they contain your script ID and
  login token.
- Use made-up project names in issues, screenshots and examples.

## Reporting a vulnerability

Please open a GitHub issue **without** sensitive details and ask for a private
contact, or use GitHub's "Report a vulnerability" button on the Security tab if
it is enabled for this repository.
