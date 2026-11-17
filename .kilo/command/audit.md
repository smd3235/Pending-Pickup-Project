---
description: Audit the sibling project for portable scripts, procedures and rules
argument-hint: "[sections] [-report]"
---

Audit the sibling project (`Pending Pickup Project`) for parts worth bringing here,
then do the same in reverse. This is read-only on both sides.

The peer is discovered automatically: a sibling directory with a root `*.bas` and a
`scripts\` folder. `-Sections` narrows the audit (`scripts`, `documentation`,
`commands`, `vba`, `rules`); `-Report` writes a Markdown copy.

Full audit:

!`python "scripts\audit-peer.py"`

Sections that exist on both sides only (fast):

!`python "scripts\audit-peer.py" -Sections scripts,commands -SymbolLimit 40`

What the audit compares:

- files present on one side only, and files that differ (with per-file line counts)
- VBA procedure names: shared, peer-only (port candidates), and here-only
- Python function names and Markdown section headings, per matched file

Then report, concisely:

1. A shortlist of **port candidates**: for each, one line on what it does, whether
   it would actually help here, rough effort, and the risk of taking it. Say plainly
   when something looks portable but is not worth it (e.g. it duplicates something
   already here, or it assumes the peer's data).
2. A shortlist of what the peer might want from here.
3. Anything that looks like a rule or process difference (AGENTS.md sections,
   commit and changelog conventions) that is worth aligning.

Do not copy anything. Porting is a separate, approved task. When it is approved:
copy the file, adapt its data source and paths to this project, then run the
mandatory loop (`.\scripts\test-compile.py`, `.\scripts\sync_template.py`,
`.\scripts\run-report.py`) and update `documentation\changelog.md`.