# msxl-tags

`msxl-tags` explores how to bring tag-like or pill-like behavior to Microsoft Excel using VBA and a future `.xlam` add-in.

The repository is currently docs-first: the goal is to lock the data model, rendering assumptions, and Excel-specific tradeoffs before building the first implementation.

## Current Status

This project is in planning and design mode.

- No VBA implementation is committed yet.
- No `.xlam` add-in or workbook binaries are committed yet.
- The docs in this repository are the current source of truth for project direction.

## Why This Exists

Excel works well for structured tables, formulas, and calculations, but it does not natively offer a clean tag or pill UI for lightweight metadata inside the worksheet grid.

This project exists to explore a practical way to:

- store one or more tags per target cell in a reliable form
- parse and update those tags from VBA without corrupting adjacent state
- present a result that feels recognizably tag-like within Excel's constraints
- document the tradeoffs clearly enough that future implementation can proceed with fewer re-decisions

## What We Are Deciding Now

The current design pass is focused on a few core questions:

- how tags should be represented so they can be created, parsed, updated, reordered, and removed reliably
- which storage strategy gives the best balance of inspectability, durability, and edit safety
- how far inline rendering can go before Excel limitations make visual emulation too fragile
- what responsibilities should belong to the future add-in versus workbook-level content

The current working direction in the design docs favors a hybrid model:

- a human-readable cell display
- a structured backing payload in a secondary Excel-supported location
- inline rendering as the baseline visual mode, with room to evaluate higher-fidelity options later

## Planned Deliverable

The intended runtime package is a `.xlam` Excel add-in.

When implementation begins, the review-friendly source of truth will remain exported text-based VBA assets in git, while the `.xlam` will be treated as the runnable package assembled from those sources.

That means future repository contents are expected to include:

- exported VBA modules, classes, and forms as text files
- workflow notes for import and export
- examples and assets that support testing, explanation, and evaluation

## Repository Layout

- `docs/` - product framing, data model notes, and design decisions
- `src/vba/` - future home for exported VBA modules, classes, forms, and workflow notes
- `examples/` - future sample workbooks and usage scenarios
- `assets/` - future screenshots, diagrams, and supporting visuals

## Next Milestones

1. Lock the initial tag data model and rendering assumptions.
2. Decide how worksheet-visible state and backing metadata should be stored.
3. Define the first import/export workflow for text-based VBA source files.
4. Build the first prototype `.xlam` once the documentation baseline is stable.

## Read Next

- [Product Scope](docs/product-scope.md)
- [Data and Rendering Model](docs/data-rendering-model.md)
- [Contributing](CONTRIBUTING.md)
- [Agent Guidance](AGENTS.md)

## License

This project is licensed under the GNU General Public License v3.0. See [LICENSE](LICENSE).
