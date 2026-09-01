# Release process

The NUC Forgejo repository `BramAlkema/openxml-audit` is canonical. GitHub is a
release-only downstream because PyPI supports GitHub Actions as a Trusted Publisher but
does not support a self-hosted Forgejo issuer.

```text
Forgejo main + annotated vX.Y.Z tag
    -> release-mirror.yml validates the canonical ref
    -> fast-forward GitHub main and add the immutable tag
    -> GitHub release.yml builds once
    -> PyPI Trusted Publishing and GitHub Release
```

## One-time configuration

- GitHub has a write-enabled deploy key named `Forgejo NUC release mirror`. Its private
  key is stored only as the Forgejo repository Actions secret
  `RELEASE_MIRROR_DEPLOY_KEY`.
- Only `.github/workflows/release.yml` is enabled on GitHub. CI, security, and docs jobs
  run on Forgejo or remain disabled on the downstream mirror.
- PyPI's Trusted Publisher remains bound to GitHub repository
  `BramAlkema/openxml-audit` and workflow `release.yml`.

## Creating a release

1. Merge the release commit to Forgejo `main` and verify all required Forgejo checks.
2. Ensure `project.version` in `pyproject.toml` is the intended version.
3. Create an annotated tag named exactly `v<project.version>` on that `main` tip.
4. Push `main` and the tag to Forgejo `origin`; do not push either directly to GitHub.
5. Confirm `Mirror release to GitHub publisher` succeeds on Forgejo.
6. Confirm GitHub's `Release` workflow succeeds, then verify the version-specific PyPI
   JSON response, Simple index, clean installation, packaged oracle entry points, and
   PyPI provenance.

## Safety properties

`scripts/release_mirror.py` fails closed unless all of these are true:

- the event comes from the expected Forgejo repository;
- the tag matches `pyproject.toml` and is annotated;
- the tag resolves to the current Forgejo `main` tip;
- an existing GitHub tag is byte-for-byte the same tag object;
- GitHub `main` is an ancestor of the release commit; and
- the post-push GitHub branch and tag resolve to the expected objects.

The bridge pushes only `main` and the single release tag. It never uses a Git push
mirror, force-push, or a PyPI API token.

If GitHub has diverged, leave the Forgejo tag intact and reconcile the downstream copy
explicitly. Never bypass the ancestry or immutable-tag checks.
