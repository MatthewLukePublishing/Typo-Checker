# Typo Checker

Typo Checker exports text from the one saved document open in Adobe InDesign,
reviews it through the official Codex CLI and a ChatGPT subscription, and writes
an Excel corrections report. Its public launcher downloads a fresh copy of the
latest stable GitHub release for every run. It is designed to fail closed: it
never uses an OpenAI API key, never falls back to the paid API, and never silently
selects an older model or stale program release.

The primary workflow checks every text item twice. For every query it:

- resolves OpenAI's current official frontier model at runtime;
- records the exact resolved model in the job record;
- resolves the frontier again immediately before the query is sent;
- uses independent `xhigh` reasoning for the primary review and recheck;
- stops if model resolution fails, the configured model is outdated, or the
  official frontier changes during the job; and
- requires Codex CLI authentication through ChatGPT, with API credentials stripped
  from every child process.

## Requirements

- Windows 10 or Windows 11
- Node.js 20 or newer
- Adobe InDesign 2026 for the open-document workflow
- The official [Codex CLI](https://learn.chatgpt.com/docs/codex/cli), signed in with
  **Sign in with ChatGPT**
- A ChatGPT plan with access to Codex
- Microsoft Excel or another application that can open `.xlsx` files
- Internet access while a review is running

The standalone workbook tools do not require InDesign.

## Open InDesign and run the latest release

1. Open exactly one `.indd` document in InDesign 2026.
2. Save it and leave it open with no unsaved changes.
3. Ask Codex/ChatGPT to follow the copy-paste instructions in
   [`CHATGPT_PROMPT.md`](CHATGPT_PROMPT.md), specifying the language to check.

If `Run_Latest_Typo_Checker.ps1` is already available locally, the direct command
is:

```powershell
powershell -ExecutionPolicy Bypass -File .\Run_Latest_Typo_Checker.ps1 -Language German
```

The launcher does not trust or execute a previously downloaded checker. On every
run it:

- queries the canonical repository's latest stable release;
- resolves the release tag to an immutable 40-character commit SHA;
- downloads that exact commit into a new temporary folder;
- verifies the archive structure, package version, bundled exporter, and offline
  self-tests;
- rechecks the latest release immediately before starting the checker; and
- records the release tag, commit SHA, and downloaded archive SHA-256 in the job.

If the latest release cannot be resolved or changes during preparation, the run
stops. Only the temporary program download is removed afterward. The generated
job and report files remain under this folder beside the open document:

```text
typo_checks\<run-id>\
  input\content_export.xlsx
  reports\content_export_corrections.xlsx
  reports\content_export_corrections.xlsx.job.json
```

No Excel input file is stored in or downloaded from GitHub. The bundled, hash-
verified InDesign exporter creates `content_export.xlsx` locally from the open
document during each run. The controller does not save, close, or modify a
document that was already open.

No `.env` file, API key, or OpenAI API billing account is used by this program.
Run `codex` once and choose **Sign in with ChatGPT** if the CLI is not already
authenticated.

## Download for standalone workbook use

Download and extract the latest release ZIP. The executable checker files are
bundled; `npm install` is needed only when developing or rebuilding the project.

## Quick start: check one Excel column

Create an `.xlsx` workbook containing the text to review in column A. Do not add a
header row or populate any other column. Then run:

```powershell
$env:INPUT_XLSX = "C:\path\to\input.xlsx"
$env:OUTPUT_XLSX = "C:\path\to\typo_review.xlsx"
node .\Typo_Checker_Batch.js
Remove-Item Env:INPUT_XLSX, Env:OUTPUT_XLSX
```

The output has three headerless columns:

1. original text;
2. corrected text; and
3. the independent recheck note.

A JSON job record is written beside the output workbook. It contains hashes,
query evidence, the model policy, and the exact model used. Input, output, and job
record paths must be different.

Optional settings are supplied as PowerShell environment variables before the
command:

```powershell
$env:CHECK_LANGUAGE = "German"
$env:ROW_LIMIT = "20"
$env:CODEX_QUERY_TIMEOUT_MS = "1800000"
```

`OPENAI_MODEL` may be set only as a guardrail. If it is not the current official
frontier, the run stops instead of substituting another model.

## Other workbook workflows

- `Compare_5_Columns.js` compares exactly five translations per headerless row and
  writes a note for each target.
- `5_Columns Rechecker.js` independently rechecks the nontrivial comparison notes
  and verifies the source workbook and comparison-output hashes.
- `Style_Guide_Checker.js` checks headerless column-A text against the embedded
  style requirements.
- `Subscription_Harness.js` is the shared ChatGPT-subscription execution boundary
  for those three helper workflows.

Each script is configured with environment variables declared near the beginning
of the file. It rejects unexpected extra columns instead of silently discarding
them.

## Translation-job InDesign workflow

In addition to the public one-open-document launcher, `Typo_Checker.js` retains
the internal production workflow. It opens only the production INDD recorded in
an active-job JSON file, creates an isolated export, checks every exported
paragraph, performs the mandatory recheck, and writes a new report workbook and
JSON job record.

The distribution includes `Export All Content to Excel.jsx`, the exact exporter
whose SHA-256 is verified before each run. To use a non-default active-job file:

```powershell
$env:ACTIVE_JOB_CONFIG_PATH = "C:\path\to\Active Job.json"
node .\Typo_Checker.js
Remove-Item Env:ACTIVE_JOB_CONFIG_PATH
```

In unattended translation-job mode, InDesign must have no open documents. For a
deliberate test against the one configured document already open in InDesign:

```powershell
$env:ACTIVE_JOB_CONFIG_PATH = "C:\path\to\Active Job.json"
$env:USE_OPEN_INDESIGN_DOCUMENT = "true"
node .\Typo_Checker.js
Remove-Item Env:ACTIVE_JOB_CONFIG_PATH, Env:USE_OPEN_INDESIGN_DOCUMENT
```

The generic public launcher uses `OPEN_INDESIGN_DOCUMENT=true`; it discovers the
single open document without requiring the private production `Active Job.json`.
All InDesign discovery and export operations share one global mutex and run
serially. Existing translation manifests, source exports, and workbooks are never
overwritten.

If export succeeded but a later review stage failed, `REUSE_EXPORT_JOB_PATH` can
rerun the review stages without invoking InDesign or exporting again.

## Offline validation

Run all self-tests without opening InDesign or contacting Codex:

```powershell
npm test
```

The self-tests validate the bundled checker, adjacent exporter, public launcher,
subscription boundary, input contracts, release provenance, embedded PowerShell,
safe cleanup boundary, and Excel read/write support. They do not contact GitHub,
Codex, or InDesign.

## Repository permissions

This public repository is maintained by Matthew Luke Publishing. The public has
read and download access but no write access to the canonical repository. Anyone
may download or fork the project and modify their own copy under the MIT License.
A pull request is only a proposal and changes nothing here unless the repository
owner reviews and merges it.

## Maintainer notes

### Management contract

- Primary entry point: `Run_Latest_Typo_Checker.ps1`
- Bundled checker: `Typo_Checker.js`
- Archive boundary: `None`

The maintained source folder is
`D:\Google Drive\Publishing\Code\Programs\Typos`. After source, documentation,
or project-organization changes, run the publishing-code consolidation process.

## Authentication contract

This program must not store secrets or authentication state locally. Maintainer
credentials, tokens, cookies, browser profiles, and provider sessions belong only
under `D:\Google Drive\Publishing\Code\Admin\Access`. Never commit an `.env` file,
API key, token, job input, review output, or authentication artifact.

## License

[MIT](LICENSE)
