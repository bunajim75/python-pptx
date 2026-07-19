# python-pptx fork context

## Purpose and upstream relationship

This repository is a disciplined, general-purpose fork of
[`scanny/python-pptx`](https://github.com/scanny/python-pptx), retaining the
`python-pptx` library's existing behavior and public import namespace. Its
current upstream base is release `v1.0.2`, commit
`278b47b1dedd5b46ee84c286e77cdfb0bf4594be`.

## Current state

- Fork capabilities beyond upstream: none.
- Public API additions: none.
- Lecturegen integration: not integrated. Lecturegen-specific concepts remain
  outside this repository.
- Verification: changes require proportionate unit, acceptance, integration,
  and round-trip coverage, plus desktop PowerPoint inspection where behavior
  can affect rendered or application-specific results.
- Graphify: use it only to navigate the codebase; confirm conclusions in the
  real files and keep generated output untracked.

## Initial roadmap

Implement and review each item independently, in this order:

1. Slide-layout cloning research and fixtures.
2. `SlideLayouts.clone()`.
3. Shape creation on layouts and masters.
4. Placeholder creation.
5. Theme editing.
6. Optional later slide cloning and advanced PowerPoint structures.
