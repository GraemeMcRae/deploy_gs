# deploy_gs.py

### What deploy_gs.py does

Works with Google Sheets to let you edit complex formulas offline, then quickly and
simply deploy updated formulas to your spreadsheet.

You can edit formulas using Visual Studio Code (VSCode), a robust Interactive
Development Environment (IDE) that — with a plug-in — understands Google Sheets formula
syntax, and provides all the bells and whistles you expect from a good IDE: syntax
highlighting, bracket matching, and structured formatting across as many lines as you like.

When you're ready to test a formula, a simple command copied directly from the formula
itself will deploy the updated formula directly into the correct cell of your spreadsheet —
comments stripped, deployment timestamp updated, pretty formatting preserved.

---

### Prerequisites

**Platform**
- Windows or Linux (e.g. Ubuntu)
- GitBash or Bash shell

**Git** — optional but recommended for version-controlling your `.gs` formula files.

**Python setup**
1. Establish a project directory. All files mentioned below go in this directory.
2. Create a Python Virtual Environment (`venv`) or use an existing one. I recommend using Python version 3.13 or later. (Use an AI assistant for detailed instructions if you've never done this before.)
3. Install Python packages.
   - If you will be using deploy_gs.py in an existing Python environment, merge the following into your existing `requirements.txt`
   - Otherwise, use the `requirements.txt` file I've bundled with this package.
   - then run `pip install -r requirements.txt`

```
python-dotenv>=1.0.0
gspread>=6.1.2
google-auth>=2.38.0
pytz>=2024.1
tzdata>=2024.1
```

4. Save `deploy_gs.py` in your project directory.

**Google credentials**
1. Create a Google service account and download its credentials JSON file, or use existing credentials if you have them. (If you've never done this before, I included a "HowTo" file in this package to help you. Or use an AI assistant for customized help.)
2. Create or modify a file called `.env` in your project directory to include:

```
GOOGLE_CREDENTIALS="google_credentials.json"
LOCAL_TIMEZONE="America/Los_Angeles"
```

Pick your timezone from this list:
https://en.wikipedia.org/wiki/List_of_tz_database_time_zones

3. Explicitly share your Google spreadsheet (with **Editor** access) to the service
   account email address identified in your credentials JSON file.

**Visual Studio Code**
1. Install VSCode from https://code.visualstudio.com/
2. Install the following extensions using Ctrl+Shift+X:
   - **vscode-google-sheets-syntax** by dunstontc — provides syntax highlighting for
     `.gs` files
   - **GitLens** by GitKraken — optional but recommended for Git integration

**Formula directory**

If desired, create a subdirectory of your project directory called `formulas/` (or
any name you like) to store your `.gs` source files.

**Spreadsheet**

Use an existing Google spreadsheet with complex formulas, or create a new one. Either
way, make sure you have shared it with the service account email address.

---

### Testing / validation

For this example, the spreadsheet is called `Spreadsheet Name`, and the target cell
is `SheetName!$C$14`.

**Step 1 — Get a formula to work with**

If you don't have a handy complex formula, put this one in `SheetName!$C$14`:

```
=TextJoin(", ",,Filter(Sequence(99),Map(Sequence(99),Lambda(x,Countif(Mod(x,Sequence(x-1)),0)=1))))
```

This formula computes all prime numbers from 1 to 99.

**Step 2 — Create a `.gs` source file**

Copy the formula to your clipboard. In VSCode, create a new file called `myformula.gs`
in your `formulas/` subdirectory (use whatever name you like, as long as it has the
`.gs` extension).

When prompted, associate files with the `.gs` extension with Google Sheets syntax. You
should see **Google Sheets** as the file type at the very lower right of your VSCode
window. If you see some other file type there, click it, then select
**Configure File Associations for .gs**.

Paste your formula into `myformula.gs`. Then wrap it in a `LET` statement with the
four metadata symbol/value pairs that `deploy_gs.py` relies on:

```
=LET(
  _Author,N("your name"),
  _Source,N("formulas/myformula.gs"),
  _Deployed_using,N("python deploy_gs.py 'Spreadsheet Name' 'SheetName!$C$14'"),
  _Date_deployed,N("deployment date"),
  TextJoin(", ",,Filter(Sequence(99),Map(Sequence(99),Lambda(x,Countif(Mod(x,Sequence(x-1)),0)=1))))
)
```

If your formula already begins with `=LET(`, simply add the four `_Author`,
`_Source`, `_Deployed_using`, and `_Date_deployed` symbol/value pairs to the
existing `LET` — no extra wrapping parenthesis needed.

Paste this updated formula into `SheetName!$C$14` as well, so the cell contains the
`_Source` marker that `deploy_gs.py` needs to find the source file.

**Step 3 — Run a test deployment**

Make a small change to `myformula.gs` — for example, change `99` to `199` in both
places in the formula.

Copy the deployment command from the `_Deployed_using` line in `myformula.gs`. In
your Bash shell, `cd` to your project directory, then run it:

```bash
python deploy_gs.py 'Spreadsheet Name' 'SheetName!$C$14'
```

If you see this error:

```
Error: Spreadsheet 'Spreadsheet Name' not found or not accessible.
```

it means the spreadsheet has not been shared with the service account email address
in your credentials JSON file. Share it with Editor access and try again.

If the deployment is successful, you will see output like this:

```
Opening spreadsheet: Spreadsheet Name
Fetching 1 formula cell(s) in one batch...

Processing: SheetName!$C$14
  Source file: formulas/myformula.gs
  Comments stripped. Formula length: 292 -> 292 chars.

Writing 1 formula(s) in one batch...
  Done.

Deployment complete.
```

Look at the formula in the spreadsheet. You should see something like:

```
_Date_deployed,N("2/25/2026 23:43"),
```

which confirms the deployment was successful and the timestamp was updated.

**Step 4 — Add comments and formatting**

Now go back to VSCode and restructure the formula with comments and pretty formatting:

```
=LET(
  _Author,N("your name"),
  _Source,N("formulas/myformula.gs"),
  _Deployed_using,N("python deploy_gs.py 'Spreadsheet Name' 'SheetName!$C$14'"),
  _Date_deployed,N("deployment date"),

  /* Prime Number Calculator */
  TextJoin(", ",,
    Filter(
      Sequence(199),          /* For each number from 1 to 199, filter it based on the True/False values returned by Map */
      Map(                    /*    Map returns 199 True/False values, where True indicates a prime number.              */
        Sequence(199),        /* For x=1 to 199, use the following test to determine if x is prime:                     */
        Lambda(x,
          1=Countif(          /*    return True if Mod(x,y)=0 exactly once, where                                       */
            Mod(x,
              Sequence(x-1)   /*       y ranges from 1 to x-1                                                           */
            ),
            0
          )
        )
      )
    )
  )
)
```

Deploy again, and the formula written to the sheet will have all comments removed
but the pretty indentation preserved, and the timestamp updated:

```
=LET(
  _Author,N("your name"),
  _Source,N("formulas/myformula.gs"),
  _Deployed_using,N("python deploy_gs.py 'Spreadsheet Name' 'SheetName!$C$14'"),
  _Date_deployed,N("2/25/2026 23:43"),

  TextJoin(", ",,
    Filter(
      Sequence(199),
      Map(
        Sequence(199),
        Lambda(x,
          1=Countif(
            Mod(x,
              Sequence(x-1)
            ),
            0
          )
        )
      )
    )
  )
)
```

---

### Reference manual

**Three ways to invoke deploy_gs.py**

*1. Full command line arguments*

```bash
python deploy_gs.py "Spreadsheet Name" "Sheet1!ColumnHeader" "Sheet1!AnotherCol" "Sheet2!$A$1"
```

The spreadsheet name is the first argument. All remaining arguments are cell
references. Execution begins immediately with no prompts.

*2. Spreadsheet name only — cell references entered interactively*

```bash
python deploy_gs.py "Spreadsheet Name"
```

When the spreadsheet name is given but no cell references follow it, the program
prompts you to enter cell references one per line and signals end-of-input with
Ctrl-Z (Windows) or Ctrl-D (Linux/Mac).

*3. No arguments — fully interactive*

```bash
python deploy_gs.py
```

The program first prompts for the spreadsheet name, then prompts for cell references
as above.

*4. Redirected input from a file*

```bash
python deploy_gs.py < full_deployment.txt
```

The input file has the spreadsheet name on the first line, followed by one cell
reference per line. If the spreadsheet name is already given on the command line,
the redirected input contains only the cell references.

**One spreadsheet per run**

Each invocation of `deploy_gs.py` works against exactly one spreadsheet. However,
any number of cells across any number of sheets within that spreadsheet can be
deployed in a single run. All reads and the final write are performed as batch
operations to minimize API calls.

**Cell reference formats**

Each cell reference takes one of these forms:

| Format | Meaning |
|---|---|
| `SheetName!ColumnHeader` | Row 2 of the column whose header in row 1 matches `ColumnHeader` |
| `SheetName!$C$14` | The specific cell `$C$14` on `SheetName` |
| `ColumnHeader` | Same sheet as the previous reference; column matched by header |
| `$C$14` | Same sheet as the previous reference; specific cell |

Sheet names with spaces must be quoted on the command line. In Bash/GitBash, use
single quotes to prevent `!` and `$` from being interpreted by the shell:

```bash
python deploy_gs.py 'My Spreadsheet' 'Sheet One!$A$1' 'Sheet One!$B$1'
```

**Named column references use row 2**

When you specify a column by header name rather than an absolute reference,
`deploy_gs.py` always reads from and writes to **row 2** of that column. The
assumption is that row 1 contains the header and row 2 contains the formula.
After deployment you manually copy the updated formula down the column as needed.
This keeps the program simple and gives you an easy revert path — copy row 3
back to row 2 to undo a deployment.

**How the metadata markers work**

`deploy_gs.py` looks for these markers inside the formula currently stored in the
target cell:

| Marker | Purpose |
|---|---|
| `_Source,N("formulas/myformula.gs"),` | Tells the program which `.gs` file to read |
| `_Date_deployed,N("deployment date"),` | Replaced with the current local date/time |
| `_Author,N("your name"),` | Informational only; passed through unchanged |
| `_Deployed_using,N("python deploy_gs.py ..."),` | Informational only; passed through unchanged |

The `N()` function in Google Sheets returns zero for any text argument, so these
markers have no effect on the formula's computed result. They serve purely as
self-documenting metadata visible to anyone inspecting the formula in the sheet.

**Verification with _Verify_ bookends**

Google Sheets automatically rewrites cell references in formulas when columns or rows
are inserted, deleted, or reordered. This is normally helpful, but it means that a
formula in the spreadsheet can silently drift away from its `.gs` source file. If you
then redeploy from the `.gs` file, you'll overwrite those automatic adjustments —
potentially breaking the formula without any warning.

The `_Verify_` bookend system guards against this. You wrap selected parts of your
formula in named bookend pairs. Before each deployment, `deploy_gs.py` compares those
sections between the live cell and the `.gs` file. If they don't match, deployment is
blocked and you are shown exactly what changed, so you can decide whether to update
the `.gs` file or investigate further.

*Bookend syntax*

A bookend is any symbol name of the form `_Verify_<n>_<suffix>` where `<n>` is the
bookend's name — one or more letters or digits — and `<suffix>` is anything you like
(e.g. `Begin`, `End`, `Start`). The underscore after `<n>` acts as the name
terminator, so `_Verify_1_Begin` and `_Verify_1_End` both have the name `1`, and
`_Verify_anchor_Begin` has the name `anchor`.

A bookshelf is a matched pair of bookends with the same name, appearing exactly twice
in the formula. The bookshelf content is everything from the start of the first
bookend up to (but not including) the start of the second bookend. This means the
first bookend is included in the bookshelf content and the second is not. A bookshelf
may span multiple lines.

Example formula with two bookend pairs:

```
=LET(
  _Author,N("your name"),
  _Source,N("formulas/myformula.gs"),
  _Deployed_using,N("python deploy_gs.py 'Spreadsheet Name' 'SheetName!$C$14'"),
  _Date_deployed,N("deployment date"),

  _Verify_refs_Begin,N("Cell references"),
  anchor,Indirect("Trips_Sheet!$K$13"),
  total,Sum(Trips_Sheet!$B$2:$B$100),
  _Verify_refs_End,N(""),

  anchor+total
)
```

In this example, the `refs` bookshelf content is everything from `_Verify_refs_Begin`
up to (not including) `_Verify_refs_End`:

```
_Verify_refs_Begin,N("Cell references"),
  anchor,Indirect("Trips_Sheet!$K$13"),
  total,Sum(Trips_Sheet!$B$2:$B$100),
```

If someone inserts a column before K, Google Sheets will rewrite `$K$13` to `$L$13`
in the live cell. The next time you run `deploy_gs.py`, the bookshelf in the cell
will no longer match the bookshelf in the `.gs` file, and deployment will be blocked:

```
  [Error] _Verify_refs mismatch:
    Before deployment:
      _Verify_refs_Begin,N("Cell references"),
      anchor,Indirect("Trips_Sheet!$L$13"),
      total,Sum(Trips_Sheet!$B$2:$B$100),
    After deployment:
      _Verify_refs_Begin,N("Cell references"),
      anchor,Indirect("Trips_Sheet!$K$13"),
      total,Sum(Trips_Sheet!$B$2:$B$100),
```

You can then decide: if the column insertion was intentional, update the `.gs` file
to use `$L$13`. If it was accidental, fix the spreadsheet first.

*What deploy_gs.py reports during verification*

| Message | Meaning |
|---|---|
| `[Verify] _Verify_<n>: OK` | Bookshelf matches; no action needed |
| `[Error] _Verify_<n> mismatch` | Bookshelf differs; deployment blocked for this cell |
| `[Info] _Verify_<n> exists in the cell but not in the .gs file` | Bookshelf will be removed by this deployment (informational) |
| `[Info] _Verify_<n> is new in the .gs file and not yet in the cell` | Bookshelf will be added by this deployment (informational) |
| `[Warning] _Verify_<n> appears N time(s)` | Bookend found a number of times other than 2; this bookshelf is skipped |

*Bookend naming rules*

The name `<n>` consists of letters and digits only — no underscores. The first
underscore after `_Verify_` terminates the name. So:

- `_Verify_1_Begin` → name is `1`
- `_Verify_anchor_End` → name is `anchor`
- `_Verify_myRef_Start` → name is `myRef`

This also means `_Verify_11` and `_Verify_1` are completely independent names.
A stray `_Verify_11` that appears only once will produce a warning for name `11`
and will have no effect on any bookshelf named `1`.

*Bookend placement tips*

Because the `_Verify_` symbols are just `N(...)` values inside a `LET`, they have
no effect on the formula's computed result. Place them around any expression that
contains cell references you want to protect — typically anchor cells, named ranges,
or table references that could shift if the spreadsheet structure changes.

You can add, remove, or rename bookshelves at any time. Adding a new bookshelf to
the `.gs` file produces an informational message on the first deployment (since it
doesn't yet exist in the cell) and is then tracked on all subsequent deployments.
Removing a bookshelf from the `.gs` file produces an informational message on the
next deployment and is then silently gone.

Overlapping bookshelves are permitted. For example, a large outer bookshelf covering
an entire section and a smaller inner bookshelf covering a critical cell reference
within it can coexist without interfering with each other.

The closing bookend does not need to be on its own line. It can appear at the end of
a data line, and any trailing whitespace after it is ignored.

*Bookend comparison and whitespace*

Before scanning for bookends and comparing bookshelves, `deploy_gs.py` applies a
normalization step to both the live cell formula and the comment-stripped `.gs`
content. This normalization: collapses any run of spaces or tabs within a line down
to a single space; strips leading and trailing whitespace from each line; removes
blank lines entirely; and normalises Windows-style CRLF line endings to LF. Error
messages show the normalized versions of mismatched bookshelves, which will appear
more condensed than the original source but are easier to diff visually.

Note that whitespace *inside string literals* is not affected by this normalization —
`N("hello  world")` and `N("hello world")` are considered different. Only whitespace
outside of quoted strings is collapsed.

---

### Technical information

**How comment removal works**

`deploy_gs.py` processes `.gs` source files through four steps:

*Step 1 — Block comment removal*

All `/* ... */` block comments are removed, including multi-line ones. Any
horizontal whitespace (spaces and tabs) immediately preceding the `/*` on the
same line is also consumed, so the code to the left of the comment is not left
with a ragged trailing edge of spaces.

The regex uses non-greedy matching (`.*?` with `DOTALL`) so that each `/*` is
paired with the very next `*/` encountered. This means a `*/` sequence inside a
block comment will prematurely close it. The workaround is to avoid `*/` inside
comments — which is not an unusual restriction; the C language has the same one.

*Step 2 — Line comment removal*

Line comments beginning with `//` are removed to the end of the line, along with
any horizontal whitespace immediately preceding the `//`.

To avoid clobbering URLs like `"http://google.com"`, the `//` must either appear
at the very start of a line or be preceded by whitespace. The `:` in `http:` is
not whitespace, so URLs are left alone. This is not a perfect heuristic — a `//`
preceded by whitespace inside a string literal would still be stripped — but in
practice Google Sheets formulas do not contain such constructs.

*Step 3 — Blank line collapsing*

After comment removal, runs of consecutive blank or whitespace-only lines are
collapsed to a single blank line. This prevents comments that occupied their own
lines from leaving large gaps in the deployed formula.

*Step 4 — Leading/trailing whitespace trimming*

The entire result is stripped of leading and trailing whitespace before being
written to the sheet.

**The comment delimiter edge case**

Because the comment delimiters are recognized everywhere in the source text —
not just outside of string literals — a formula that needs to contain the literal
strings `/*`, `*/`, or `//` as data values will have those sequences stripped as
if they were comments. The solution is to construct them using string
concatenation so the delimiter never appears as a literal sequence:

```
LeftCommentDelim,  "/"&"*",
RightCommentDelim, "*"&"/",
DoubleSlash,       "/"&"/",
```

This is analogous to the way you would escape special sequences in any other
template or preprocessing system.

**Batch API strategy**

To minimize the risk of hitting Google Sheets API rate limits (HTTP 429), all
reads and writes are consolidated into as few API calls as possible:

- All header rows (row 1) for sheets that use named column references are fetched
  in a single batch GET.
- All formula cells are fetched in a single batch GET.
- All updated formulas are written in a single batch update.

If a retryable error occurs (HTTP 408, 429, 500, 502, 503, or 504), the program
waits 10 seconds and retries, up to 59 times, printing a message each time. After
59 retries it gives up and reports the error. Ctrl-C is handled gracefully —
the first press requests a clean shutdown after the current operation; the second
press forces an immediate exit.
