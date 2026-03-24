# Data and Rendering Model

## Design goal

Represent tags in a way VBA can manipulate predictably, while presenting a pill-like result that feels useful inside Excel's grid-based UI.

## Tag entity shape

The starting conceptual tag shape is:

- `id` - optional stable identity for updates or deduplication
- `label` - required display text
- `style_key` - semantic style or color family
- `payload` - optional backing value or metadata token
- `position` - relative order within the cell

Not every storage mode needs every field, but the model should leave room for them.

## Per-cell collection model

A target cell conceptually owns a collection of zero or more tags.

Initial requirements for the collection model:

- preserve order when order affects meaning or display
- support repeated parsing without drift
- allow single-tag edits without forcing ambiguous text replacements
- remain inspectable enough for debugging in Excel/VBA workflows

## Candidate storage strategies

### 1. Delimited text in the visible cell

Store the visible cell value as a serialized tag list.

Pros:

- easy to inspect
- no hidden dependency for basic data presence
- compatible with plain worksheet copy and paste

Tradeoffs:

- parsing becomes fragile if labels can contain delimiters
- styling and metadata are hard to preserve cleanly
- visible text and render intent are tightly coupled

### 2. Structured payload in cell metadata or adjacent hidden storage

Keep display text separate from a structured backing payload stored in comments, notes, shapes, hidden sheets, or named ranges.

Pros:

- supports richer metadata and more reliable updates
- allows display text to differ from storage format
- reduces ambiguity for parsing operations

Tradeoffs:

- adds indirection and recovery complexity
- can be easier for users to break accidentally
- portability depends on which Excel feature stores the payload

### 3. Hybrid model

Keep a simple visible representation in the cell and richer backing state elsewhere.

Pros:

- balances inspectability with better edit safety
- allows graceful fallback when metadata is lost
- gives the add-in room to optimize rendering separately from storage

Tradeoffs:

- reconciliation rules must be defined clearly
- two sources of state can drift without careful ownership rules

## Rendering strategies

### Inline text styling

Use cell text, separators, unicode-adjacent glyph choices, and font or fill styling to approximate pills.

Pros:

- stays anchored to the cell value
- easier to sort, filter, and copy with the worksheet

Tradeoffs:

- true rounded pill visuals are limited
- per-tag styling inside one cell is constrained
- longer tag sets may become hard to read

### Overlay shapes anchored to cells

Draw pill-like shapes on top of or near cells while keeping underlying data elsewhere.

Pros:

- closest visual match to modern tag pills
- flexible per-tag color and layout

Tradeoffs:

- shape lifecycle and sheet movement are harder to manage
- filtering, scrolling, resizing, and copy/paste can desync visuals
- performance may degrade with many tagged cells

### Hybrid rendering

Use styled inline text as the baseline and optional overlays only where higher fidelity is worth the cost.

Pros:

- offers a resilient fallback path
- can preserve usability when shapes fail or are disabled

Tradeoffs:

- needs a clear precedence model between text and overlays
- increases implementation complexity

## Expected manipulation operations

The future VBA layer should be able to:

- create tags for a target cell from user input or structured data
- parse existing cell state into a stable collection
- append, reorder, update, or remove one tag without corrupting others
- regenerate the displayed representation after data changes
- clear tag data and any linked rendering artifacts together

## Recommended starting direction

For the first prototype, favor a hybrid data model with:

- a human-readable cell display
- a structured backing payload in a secondary location
- inline rendering as the baseline visual mode

This gives the project a safer path for tag updates while avoiding early dependence on shape-heavy rendering.

## Open questions for implementation

- which Excel-backed storage mechanism gives the best balance of durability and debuggability
- how much tag metadata must survive plain copy and paste
- whether formula cells can participate directly or need wrapper behavior
- what limits should apply to tag count, label length, and render density per cell
