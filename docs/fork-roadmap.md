# Fork roadmap

The fork will add capabilities in small, independently reviewable changes:

1. Research slide-layout cloning.
2. Add `SlideLayouts.clone()`.
3. Support shape creation on layouts and masters.
4. Add placeholder creation.
5. Add theme editing.
6. Consider slide cloning and other advanced PowerPoint structures only later.

For every capability, complete these distinct phases in order:

1. Open XML research and representative PowerPoint specimens or fixtures.
2. Public API design and architecture review.
3. Focused implementation.
4. Automated unit, integration, acceptance, and round-trip verification.
5. Desktop PowerPoint verification where application behavior or visual fidelity
   is involved.

Slide-layout cloning research must establish package relationships,
layout/master inheritance, fixture coverage, round trips, and desktop
PowerPoint results before implementation of `SlideLayouts.clone()` begins.
Placeholder creation must establish inheritance, indexing, and PowerPoint
behavior. Theme editing must establish theme/XML scope, inheritance, rendering,
and desktop results. Slide cloning must establish relationship and part-copying
semantics, fixtures, round trips, and desktop results. Transitions or animations
must first establish Office compatibility, relationship semantics, round trips,
and desktop PowerPoint inspection.

Automated tests support these gates but do not prove visual or desktop
compatibility. Lecturegen-specific concepts remain outside this repository.
