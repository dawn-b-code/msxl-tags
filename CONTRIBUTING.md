# Contributing

Thanks for helping shape `msxl-tags`.

## Current contribution model

This repository is in a docs-first setup phase.

Good early contributions include:

- clarifying project goals and constraints
- refining storage and rendering proposals
- documenting Excel limitations and experiment ideas
- tightening repo conventions for future VBA work

## Branching and commits

- Use short-lived branches for focused changes.
- Keep one topic per commit when possible.
- Write commit messages that describe the outcome, not just the action.
- Reserve large restructuring for moments when docs clearly justify it.

## Working with future VBA code

When VBA implementation begins, we plan to:

- keep exported text assets in git for modules, classes, and forms
- treat the `.xlam` as the runnable package rather than the primary diff target
- document any import/export tooling alongside the source layout

## Recording experiments

Before or during prototype work:

- note the workbook context, Excel limitation, and expected outcome
- capture whether the behavior is native, simulated, or a compromise
- record any open questions that affect future interface or storage choices
- link experiment results back to the relevant design doc when they change assumptions

## Review expectations

- Prefer changes that make the next VBA step easier to understand.
- Call out tradeoffs directly when Excel forces a workaround.
- Keep docs in sync when the intended data model or rendering direction changes.
