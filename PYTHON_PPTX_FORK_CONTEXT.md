# python-pptx fork context

## Purpose and upstream relationship

This repository is a disciplined, general-purpose fork of
[`scanny/python-pptx`](https://github.com/scanny/python-pptx), retaining the
`python-pptx` library's existing behavior and public import namespace. Its
current upstream base is release `v1.0.2`, commit
`278b47b1dedd5b46ee84c286e77cdfb0bf4594be`.

## Current state

- Minimum supported Python version: Python 3.11.
- CI and tox cover Python 3.11 and 3.12.
- The inherited test suite is compatible with current pyparsing and pytest;
  pytest and Behave pass on the supported GitHub Actions matrix.
- Fork-specific PowerPoint capabilities and public API additions: none.
- Lecturegen-specific concepts remain outside this repository.
- Verification: changes require proportionate unit, acceptance, integration,
  and round-trip coverage, plus desktop PowerPoint inspection where behavior
  can affect rendered or application-specific results.
- Slide-layout research infrastructure includes independent package-level test
  helpers and an initially empty, SHA-256-pinned corpus convention with
  provenance and desktop-verification sidecars.
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
