# AGENTS.md

Repo guidance for work inside `msxl-tags`.

## Project focus

- This repository exists to design and build Excel tag or pill behavior using VBA and a future `.xlam` add-in.
- Current phase is docs-first. Start with written decisions before adding VBA or workbook artifacts.

## Current boundaries

- Prefer edits to docs, repo metadata, and project guidance until implementation is explicitly requested.
- Do not add binary Excel artifacts as placeholders.
- Do not treat an `.xlam` file as the review-friendly source of truth once coding begins; keep exported text assets in git.

## VBA workflow expectations

When implementation starts:

- Store exported VBA modules, classes, and forms as text files in `src/vba/`.
- Keep import/export steps documented so the workbook or add-in can be reconstructed reliably.
- Separate design docs from implementation notes so the reasoning stays easy to review.

## Git hygiene

- Keep commits focused and descriptive.
- Avoid committing Office lock files, autosave artifacts, or machine-specific editor settings.
- Prefer small docs-first changes while the architecture is still forming.

## Design priorities

- Optimize for reliable tag parsing and updates.
- Be explicit about Excel limitations and what is only visually emulated.
- Document tradeoffs between storage fidelity, usability, and rendering performance.

## Collaboration style

- Keep explanations practical and friendly.
- Record assumptions in docs when choosing between imperfect Excel workarounds.
- Leave clear notes for the future add-in/export workflow instead of relying on memory.
