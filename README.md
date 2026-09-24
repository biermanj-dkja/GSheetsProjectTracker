# GSheetsProjectTracker

**Google Sheets Lightweight Project Tracking AppScript**

A single Apps Script file that turns a plain Google Sheet into a personal
project tracker. It does the housekeeping for you (archiving finished projects,
dating entries, flagging stale work, keeping stats) **without adding anything
new for you to learn**. You keep typing into the same few cells; the script and
formulas do the rest.

> 🤖 **Vibe coded with Claude.** This project was built conversationally with
> [Claude](https://claude.ai) by Anthropic, using the **Claude Opus 5.5** model
> at **medium effort**. Review the code before trusting it with important data.
> It's heavily commented to make that easy.

---

## Features

- **One front page of open projects.** Each shows its newest update, action and
  date, pulled from the log automatically.
- **Staleness colours.** The last-update date is green when recent, yellow after
  7 days, and red after 14. Projects on hold turn grey instead of going stale.
- **Priority sorting.** The open list re-sorts itself when you set or change a
  priority.
- **Automatic dates.** A new log entry gets today's date if you leave the date
  blank.
- **A–Z project dropdown** on the log. It warns you (without blocking you) if a
  project name doesn't match, so typos can't hide entries.
- **One-click close-out.** Log an entry with the action `complete`, then choose
  **Close completed projects**. The project's entries move to the archive, it
  comes off the front page, and no blank rows are left behind.
- **Stats tab.** It shows open, closed, stale and on-hold counts, average days
  to complete, average age of open projects, and updates per project per week.
  It also has a per-project table. It's all live formulas, so there's nothing to
  refresh.
- **New-year rollover.** This copies the tracker for next year with open
  projects carried over and an empty archive. The old file stays as the
  record.
- **Optional graduation-year tab** for schools: a class-year → grade table that
  updates itself every July.

## How the sheet is laid out

| Tab | What it's for | You type in | Automatic |
|---|---|---|---|
| **STATUS** | Open projects | A: project, B: priority (1 = highest) | C:F latest entry; H:J completed list |
| **In Process** | Running log of updates | D: project, E: item/note, F: action, H: link | G: date (if left blank) |
| **COMPLETED** | This year's archive | Nothing | A:B written by the script; C:F final entry; H:L archived log |
| **Stats** | Numbers | Nothing | Everything |
| **Grad year List** | *Optional* grade lookup | Nothing | A: class year, B: grade |

Columns A:C on **In Process** are left free for your own use.

The **Action** column is a short status word such as `begun`, `email`,
`testing`, `waiting`, `HOLD` or `complete`. Two words are special:

- anything containing **hold** marks the project as on hold (grey, never stale)
- an action that **starts with** `complete`, `completed`, `done`, `closed`,
  `canceled` or `cancelled` marks the project as ready to close

## Installation

1. Open (or create) the Google Sheet you want to use.
2. Go to **Extensions → Apps Script**.
3. Delete anything in `Code.gs`, paste in the contents of
   [`src/Code.gs`](src/Code.gs), and click **Save**.
4. Close the Apps Script tab and **reload the spreadsheet**. A **Project Tools**
   menu appears.
5. Choose **Project Tools → One-time setup / repair** and approve the
   permissions prompt. Google will warn that the app is unverified, because it
   is your own script. Click *Advanced → Go to project* to continue.

Setup creates any missing tabs with headers, so it works on a blank spreadsheet
as well as an existing tracker. It never changes the data you typed; it only
rebuilds formulas, colours, the dropdown and the Stats tab. It's safe to run
again any time something looks off.

> **Tip:** Try it on a copy first (**File → Make a copy**) if you're installing
> into a tracker you already rely on.

### Using clasp (optional)

If you manage Apps Script from the command line with
[clasp](https://github.com/google/clasp), copy `.clasp.json.example` to
`.clasp.json`, add your script ID, and run `clasp push`. `.clasp.json` is
git-ignored.

## Daily use

1. **New project:** type its name in STATUS column A and a priority in B.
2. **Update:** on In Process, pick the project from the dropdown and write what
   happened and a one-word action. The date fills itself in.
3. **Finish:** log an entry with the action `complete`, then choose
   **Project Tools → Close completed projects**. Confirm the list and it's done.

### Menu reference

| Menu item | What it does |
|---|---|
| Close completed projects | Archives every project whose newest entry is a closing action. Also warns about log entries that don't match any project. |
| Sort log by project & date | Groups the log by project, oldest entry first. |
| Sort open projects by priority | Re-sorts STATUS (normally automatic). |
| Start new year… | Makes a copy for next year. If the file name ends in a year pair like `26-27`, the copy is named `27-28`. |
| One-time setup / repair | Builds or rebuilds everything. Run after installing or updating the script. |

## Configuration

Everything adjustable is in the `CFG` object at the top of `Code.gs`:

| Setting | Default | Meaning |
|---|---|---|
| `STATUS`, `LOG`, `DONE`, `STATS`, `GRAD` | tab names | Rename to match your sheet |
| `YELLOW_DAYS` / `RED_DAYS` | 7 / 14 | Days without an update before a date turns yellow / red |
| `HOLD_WORD` | `hold` | Action text that means "on hold" |
| `AUTO_DATE` | `true` | Fill today's date on new log entries |
| `AUTO_SORT_STATUS` | `true` | Re-sort STATUS when a priority changes |
| `NEW_YEAR_MONTH` | 7 | Month the grad-year tab rolls forward |
| `CLOSE_WORDS` | regex | Actions that mean "finished" |

After changing a setting, run **One-time setup / repair** again.

## How it works (short version)

- **STATUS C:F** and **COMPLETED C:F** are each a single `ARRAYFORMULA`. It
  sorts the log newest-first and looks up each project name. Because it's one
  formula keyed on names rather than one formula per row, rows can be added,
  removed and re-sorted freely.
- **Close completed projects** sorts the log so each project is one block,
  copies the blocks to the archive (keeping links and formatting), deletes them
  from the log, and rewrites the STATUS list.
- **Stats** is built from one `LET`/`MAP` formula that stacks the open log and
  the archive together.
- `onOpen` and `onEdit` are simple triggers, so no trigger setup is needed.

Every function in `Code.gs` is commented in detail.

## Requirements and limitations

- Google Sheets with the current formula set (`LET`, `MAP`, `LAMBDA`, `HSTACK`,
  `VSTACK`). All Google accounts have these.
- **Automatic dates** use the spreadsheet's time zone (**File → Settings**). If
  entries are dated a day off, check that setting.
- **New year** copies are saved to the root of your My Drive; move them wherever
  you like.
- Project names are matched ignoring case and surrounding spaces, so `MUSIC` and
  `music` are the same project.
- Designed for one person's tracker. Several people editing at once will work,
  but two people running *Close completed projects* at the same moment is not
  guarded against.

## Privacy

The script runs only inside your spreadsheet. It makes no network requests and
sends no data anywhere. See [SECURITY.md](SECURITY.md).

## Contributing

Issues and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).
The guiding rule: **add automation, not human-facing complexity.**

## License

[MIT](LICENSE)
