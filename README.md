# msxl-tags

`msxl-tags` is a docs-first starter repository for exploring tag-like or pill-like behavior in Microsoft Excel with VBA and a future `.xlam` add-in.

## Current phase

We are in project setup and design mode.

- No VBA implementation lives here yet.
- No `.xlam` or workbook binaries are committed yet.
- Docs are the current source of truth for project direction and design decisions.

## What we are building

The project explores ways to create, process, manipulate, and display tag-like or pill-like items inside Excel cells.

Early work is focused on:

- representing tags in a way VBA can reliably read and update
- choosing storage strategies that work within normal worksheet constraints
- emulating pill-like display in or around cells with acceptable Excel tradeoffs
- defining the future add-in responsibilities before writing modules

## Planned deliverable

The intended runtime package is a `.xlam` Excel add-in.

When implementation starts, we expect to keep exported text-based VBA assets in git for reviewability and diff-friendly history, while the `.xlam` remains the runnable package assembled from those sources.

## Repository layout

- `docs/` - product framing and design notes
- `src/vba/` - future home for exported VBA modules, classes, forms, and import/export workflow notes
- `examples/` - future sample workbooks and usage scenarios
- `assets/` - future screenshots, diagrams, and supporting visuals

## Working conventions

- Keep the repo lean until the first design decisions are stable.
- Prefer documenting constraints and tradeoffs before adding implementation.
- Treat Excel binaries as generated or release artifacts unless we explicitly decide to version a sample workbook later.
- Capture design changes in docs so future VBA work has a stable baseline.

## Initial docs

- `docs/product-scope.md`
- `docs/data-rendering-model.md`
- `CONTRIBUTING.md`
- `AGENTS.md`

## Next milestones

1. Lock the initial tag data model and rendering assumptions.
2. Decide how worksheet state and metadata should be stored.
3. Define the first import/export workflow for VBA source files.
4. Create the first prototype `.xlam` once the doc baseline is stable.
