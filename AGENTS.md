# Repository working rules

This is a general-purpose fork of `python-pptx`; it is not a home for
Lecturegen-specific code or concepts.

- Inspect the existing architecture, tests, fixtures, and upstream patterns before implementing a change. Preserve upstream naming, style, licensing, attribution, and repository structure.
- Keep changes focused. Avoid unrelated refactors, reformatting, generated output, and fixture or binary asset changes. Use focused branches and pull requests, and inspect diffs for accidental generated, binary, fixture, or asset changes.
- Design coherent public APIs; do not require consumers to manipulate private XML internals.
- Add proportionate unit, acceptance, integration, and round-trip tests. Use XML and PowerPoint fixtures when they are necessary to establish behavior.
- For master, layout, placeholder, theme, chart, media, transition, or animation behavior, inspect results in desktop PowerPoint. Automated tests are never proof of visual or desktop PowerPoint compatibility.
- Keep the MIT license and upstream attribution intact.
- Graphify is navigation assistance only: verify against real source and configuration files. All generated reports, graph files, visualizations, and output directories must remain untracked unless explicitly approved.

Consult the existing development documentation for detailed project practices rather than duplicating it here.
