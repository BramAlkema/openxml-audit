#!/usr/bin/env python3
"""Validate a Forgejo release tag and mirror it safely to GitHub.

The NUC Forgejo repository is authoritative. GitHub is a release-only
downstream used for GitHub OIDC trusted publishing to PyPI. This helper
deliberately pushes only ``main`` and one immutable annotated release tag;
it never uses ``--mirror`` or ``--force``.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

import tomllib

PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXPECTED_FORGEJO_REPOSITORY = "BramAlkema/openxml-audit"
GITHUB_RELEASE_REPOSITORY = "git@github.com:BramAlkema/openxml-audit.git"


class ReleaseMirrorError(RuntimeError):
    """Raised when a release cannot be mirrored without weakening a guard."""


@dataclass(frozen=True)
class ReleaseContext:
    """Validated immutable identities used by the mirror operation."""

    tag: str
    tag_object: str
    commit: str


def _git(
    arguments: Sequence[str],
    *,
    cwd: Path,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        ["git", *arguments],
        cwd=cwd,
        check=False,
        capture_output=True,
        text=True,
    )
    if check and result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip() or "git command failed"
        raise ReleaseMirrorError(f"git {' '.join(arguments)}: {detail}")
    return result


def _git_output(arguments: Sequence[str], *, cwd: Path) -> str:
    return _git(arguments, cwd=cwd).stdout.strip()


def _project_version(project_root: Path) -> str:
    pyproject_path = project_root / "pyproject.toml"
    with pyproject_path.open("rb") as handle:
        pyproject = tomllib.load(handle)
    try:
        return str(pyproject["project"]["version"])
    except KeyError as exc:
        raise ReleaseMirrorError("pyproject.toml has no project.version") from exc


def _remote_ref(remote: str, ref: str, *, cwd: Path) -> str | None:
    output = _git_output(["ls-remote", "--refs", remote, ref], cwd=cwd)
    if not output:
        return None
    lines = output.splitlines()
    if len(lines) != 1:
        raise ReleaseMirrorError(f"remote returned multiple values for {ref}")
    sha, returned_ref = lines[0].split(maxsplit=1)
    if returned_ref != ref:
        raise ReleaseMirrorError(f"remote returned unexpected ref {returned_ref!r} for {ref}")
    return sha


def validate_release(
    *,
    project_root: Path,
    tag: str,
    event_sha: str,
    forgejo_repository: str,
    origin: str = "origin",
) -> ReleaseContext:
    """Validate that an annotated version tag is exactly the Forgejo main tip."""
    if forgejo_repository != EXPECTED_FORGEJO_REPOSITORY:
        raise ReleaseMirrorError(
            "refusing release from unexpected repository "
            f"{forgejo_repository!r}; expected {EXPECTED_FORGEJO_REPOSITORY!r}"
        )

    expected_tag = f"v{_project_version(project_root)}"
    if tag != expected_tag:
        raise ReleaseMirrorError(
            f"release tag {tag!r} does not match pyproject version {expected_tag!r}"
        )

    tag_ref = f"refs/tags/{tag}"
    object_type = _git_output(["cat-file", "-t", tag_ref], cwd=project_root)
    if object_type != "tag":
        raise ReleaseMirrorError(f"{tag} must be an annotated tag, got {object_type!r}")

    tag_object = _git_output(["rev-parse", tag_ref], cwd=project_root)
    commit = _git_output(["rev-parse", f"{tag_ref}^{{}}"], cwd=project_root)
    if event_sha not in {tag_object, commit}:
        raise ReleaseMirrorError(
            f"event SHA {event_sha} is neither tag object {tag_object} nor commit {commit}"
        )

    _git(
        [
            "fetch",
            "--no-tags",
            origin,
            "+refs/heads/main:refs/remotes/origin/main",
        ],
        cwd=project_root,
    )
    origin_main = _git_output(["rev-parse", "refs/remotes/origin/main"], cwd=project_root)
    if commit != origin_main:
        raise ReleaseMirrorError(
            f"release commit {commit} is not the current Forgejo main tip {origin_main}"
        )

    checkout_commit = _git_output(["rev-parse", "HEAD"], cwd=project_root)
    if checkout_commit != commit:
        raise ReleaseMirrorError(
            f"release checkout is at {checkout_commit}, expected tagged commit {commit}"
        )

    if _git_output(["status", "--porcelain"], cwd=project_root):
        raise ReleaseMirrorError("release checkout is dirty")

    return ReleaseContext(tag=tag, tag_object=tag_object, commit=commit)


def mirror_release(
    context: ReleaseContext,
    *,
    project_root: Path,
    mirror_remote: str = GITHUB_RELEASE_REPOSITORY,
) -> None:
    """Fast-forward GitHub main, add the immutable tag, and verify both refs."""
    tag_ref = f"refs/tags/{context.tag}"
    remote_tag = _remote_ref(mirror_remote, tag_ref, cwd=project_root)
    if remote_tag is not None and remote_tag != context.tag_object:
        raise ReleaseMirrorError(
            f"GitHub tag {context.tag} already exists at {remote_tag}, "
            f"expected {context.tag_object}; tags are never rewritten"
        )

    _git(
        [
            "fetch",
            "--no-tags",
            mirror_remote,
            "+refs/heads/main:refs/remotes/github-release/main",
        ],
        cwd=project_root,
    )
    remote_main = _git_output(["rev-parse", "refs/remotes/github-release/main"], cwd=project_root)
    ancestry = _git(
        ["merge-base", "--is-ancestor", remote_main, context.commit],
        cwd=project_root,
        check=False,
    )
    if ancestry.returncode != 0:
        raise ReleaseMirrorError(
            f"GitHub main {remote_main} is not an ancestor of release {context.commit}; "
            "refusing a non-fast-forward mirror"
        )

    if remote_main != context.commit:
        _git(
            ["push", mirror_remote, f"{context.commit}:refs/heads/main"],
            cwd=project_root,
        )
    if remote_tag is None:
        _git(["push", mirror_remote, f"{tag_ref}:{tag_ref}"], cwd=project_root)

    verified_main = _remote_ref(mirror_remote, "refs/heads/main", cwd=project_root)
    verified_tag = _remote_ref(mirror_remote, tag_ref, cwd=project_root)
    if verified_main != context.commit or verified_tag != context.tag_object:
        raise ReleaseMirrorError(
            "post-push verification failed: "
            f"main={verified_main}, tag={verified_tag}, "
            f"expected main={context.commit}, tag={context.tag_object}"
        )


def _required(value: str | None, name: str) -> str:
    if value:
        return value
    raise ReleaseMirrorError(f"{name} is required")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Safely mirror a Forgejo release tag to the GitHub PyPI publisher."
    )
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    parser.add_argument("--tag", default=os.environ.get("FORGEJO_REF_NAME"))
    parser.add_argument("--sha", default=os.environ.get("FORGEJO_SHA"))
    parser.add_argument("--repository", default=os.environ.get("FORGEJO_REPOSITORY"))
    parser.add_argument("--origin", default="origin")
    parser.add_argument("--mirror-remote", default=GITHUB_RELEASE_REPOSITORY)
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Run all canonical-release checks without touching GitHub.",
    )
    args = parser.parse_args()

    try:
        context = validate_release(
            project_root=args.project_root.resolve(),
            tag=_required(args.tag, "--tag or FORGEJO_REF_NAME"),
            event_sha=_required(args.sha, "--sha or FORGEJO_SHA"),
            forgejo_repository=_required(args.repository, "--repository or FORGEJO_REPOSITORY"),
            origin=args.origin,
        )
        if not args.validate_only:
            mirror_release(
                context,
                project_root=args.project_root.resolve(),
                mirror_remote=args.mirror_remote,
            )
    except ReleaseMirrorError as exc:
        print(f"release mirror refused: {exc}", file=sys.stderr)
        return 1

    action = "validated" if args.validate_only else "mirrored and verified"
    print(f"Release {context.tag} ({context.commit[:12]}) {action}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
