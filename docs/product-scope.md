# Product Scope

## Project intent

`msxl-tags` explores how to emulate tag or pill behavior in Excel using VBA and a future `.xlam` add-in.

The goal is not to change Excel's core cell type system. The goal is to provide a practical authoring and display model that feels like tags while remaining compatible with ordinary worksheets.

## Problem statement

Excel cells are text, numbers, formulas, errors, blanks, or rich content overlays. They do not natively support structured multi-tag pill widgets inside a single cell.

To offer tag-like behavior, we need a model that can:

- preserve a reliable tag list for each target cell
- support create, parse, update, remove, and display operations from VBA
- present a pill-like visual treatment that is understandable to users
- degrade gracefully when Excel rendering limits get in the way

## In scope for the first design pass

- defining what a tag is conceptually in this project
- comparing storage strategies for one or many tags per cell
- defining manipulation operations the add-in will eventually support
- comparing visual emulation strategies inside or around worksheet cells
- documenting Excel constraints that shape the implementation

## Out of scope for now

- Ribbon customization and full add-in UX flow
- installer or deployment automation
- production-grade import or export tooling
- cross-platform parity beyond the Excel environment we actively test
- advanced collaboration, sync, or external database features

## Early success criteria

A first prototype direction should make it possible to:

- assign one or more tags to a cell in a deterministic format
- read those tags back without losing identity or order assumptions
- update or delete individual tags without corrupting the rest
- show a display that users can recognize as tag-like, even if the effect is partly simulated
- explain the tradeoffs well enough that future VBA implementation can proceed without re-deciding the basics

## Core conceptual interface

The project assumes a future tag model with fields along these lines:

- `id` - stable identifier when tag identity matters
- `label` - user-facing text
- `color_key` or `style_key` - rendering hint
- `value` or `payload` - optional machine-usable value separate from label
- `order` - explicit or implied position within the cell's tag collection

At the cell level, the future add-in will likely need operations such as:

- `create_tags`
- `parse_tags`
- `update_tag`
- `remove_tag`
- `render_tags`
- `clear_tags`

These names are conceptual only at this stage.
