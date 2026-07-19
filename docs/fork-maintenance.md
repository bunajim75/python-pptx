# Fork maintenance

`origin` is this fork (`bunajim75/python-pptx`); `upstream` is the canonical
`scanny/python-pptx` repository. Keep `upstream` fetch-only:

```console
git remote set-url --push upstream DISABLED
git fetch upstream --tags
```

For normal work, branch from the current fork base, make one coherent change,
run the applicable commands from `tox.ini`, `Makefile`, and the development
documentation, then open a focused pull request to `master`. Typical local
commands are `pytest -qx`, `behave --format progress --stop --tags=-wip`,
`ruff check .`, and `make build`; use `tox` for the supported Python matrix
when practical. Review the complete diff before committing and pushing to
`origin`.

Synchronize upstream only on a dedicated branch and through its own pull
request:

```console
git switch master
git pull --ff-only origin master
git fetch upstream --tags
git switch -c chore/sync-upstream
git merge upstream/master
```

Resolve conflicts by preserving intentional fork changes and upstream behavior
elsewhere. Test the resulting combined fork-plus-upstream state, rather than
upstream changes in isolation. Re-run relevant unit and acceptance tests, lint,
import/build checks, and desktop PowerPoint inspection where affected. State any
local limitations and leave the full CI matrix to CI when it cannot be run
locally.

Before releases, review versioning, package metadata, dependencies, supported
Python versions, changelog/documentation, distributions, and the CI matrix.
Avoid unnecessary dependency changes and keep divergence from upstream small,
intentional, and easy to review.
