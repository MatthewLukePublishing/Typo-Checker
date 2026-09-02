# Prompt: run the latest Typo Checker on an open InDesign file

Replace `German` below if a different language should be checked, then paste the
entire prompt into Codex/ChatGPT on the Windows computer running InDesign.

```text
Run the Matthew Luke Publishing Typo Checker on the one document currently open
in Adobe InDesign 2026. Check the text as German.

Use only the canonical public repository:
https://github.com/MatthewLukePublishing/typo-checker

Do not use an existing local copy of the checker. Query the repository's latest
stable GitHub release, resolve its release tag to the exact immutable commit SHA,
and download that commit's source archive into a new temporary folder. Reject a
draft, prerelease, mutable branch archive, incomplete archive, or mismatched
package version. From that fresh archive, run:

Run_Latest_Typo_Checker.ps1 -Language German

The launcher must independently re-resolve and download the latest stable release
and must stop if the release changes during preparation. Do not run the JSX
exporter directly and do not substitute another local script.

Before running, require exactly one INDD document to already be open. It must be
saved, unmodified, and outside any Archive folder. Do not save, close, or alter
the open document, and do not run any other InDesign automation concurrently.

Use Codex only through Sign in with ChatGPT. Never use an API key, paid API, or
older-model fallback. The checker must resolve OpenAI's current official frontier
model at runtime, record it, recheck immediately before every query, and use xhigh
reasoning for both review passes.

When complete, report the GitHub release tag, commit SHA, downloaded archive
SHA-256, corrections workbook path, and JSON job-record path. Clean up only the
temporary downloaded program folder; preserve all generated job and report files.
```

The generated files are placed under `typo_checks` beside the open InDesign
document. No input Excel workbook needs to be downloaded from GitHub; the verified
InDesign exporter creates it during the run.
