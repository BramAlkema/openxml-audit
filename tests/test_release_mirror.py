from __future__ import annotations

import subprocess
from pathlib import Path

import pytest
from scripts.release_mirror import (
    EXPECTED_FORGEJO_REPOSITORY,
    ReleaseMirrorError,
    mirror_release,
    validate_release,
)


def _git(cwd: Path, *arguments: str) -> str:
    return subprocess.run(
        ["git", *arguments],
        cwd=cwd,
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()


def _write_version(project: Path, version: str) -> None:
    (project / "pyproject.toml").write_text(
        f'[project]\nname = "release-mirror-fixture"\nversion = "{version}"\n',
        encoding="utf-8",
    )


def _commit(project: Path, message: str) -> str:
    _git(project, "add", "pyproject.toml")
    _git(project, "commit", "-m", message)
    return _git(project, "rev-parse", "HEAD")


def _release_fixture(tmp_path: Path) -> tuple[Path, Path, str]:
    origin = tmp_path / "origin.git"
    github = tmp_path / "github.git"
    project = tmp_path / "project"
    _git(tmp_path, "init", "--bare", str(origin))
    _git(tmp_path, "init", "--bare", str(github))
    _git(tmp_path, "init", "-b", "main", str(project))
    _git(project, "config", "user.name", "Release Mirror Test")
    _git(project, "config", "user.email", "release-mirror@example.test")
    _git(project, "remote", "add", "origin", str(origin))

    _write_version(project, "1.2.2")
    _commit(project, "Initial release")
    _git(project, "push", "origin", "main")
    _git(project, "push", str(github), "main")

    _write_version(project, "1.2.3")
    release_commit = _commit(project, "Release 1.2.3")
    _git(project, "tag", "-a", "v1.2.3", "-m", "Release 1.2.3")
    _git(project, "push", "origin", "main", "refs/tags/v1.2.3")
    return project, github, release_commit


def test_release_mirror_fast_forwards_main_and_pushes_annotated_tag(tmp_path: Path) -> None:
    project, github, release_commit = _release_fixture(tmp_path)
    context = validate_release(
        project_root=project,
        tag="v1.2.3",
        event_sha=release_commit,
        forgejo_repository=EXPECTED_FORGEJO_REPOSITORY,
    )

    mirror_release(context, project_root=project, mirror_remote=str(github))
    mirror_release(context, project_root=project, mirror_remote=str(github))

    assert _git(project, "ls-remote", str(github), "refs/heads/main").split()[0] == release_commit
    assert (
        _git(project, "ls-remote", str(github), "refs/tags/v1.2.3").split()[0] == context.tag_object
    )


def test_release_validation_rejects_version_mismatch(tmp_path: Path) -> None:
    project, _, release_commit = _release_fixture(tmp_path)

    with pytest.raises(ReleaseMirrorError, match="does not match pyproject version"):
        validate_release(
            project_root=project,
            tag="v9.9.9",
            event_sha=release_commit,
            forgejo_repository=EXPECTED_FORGEJO_REPOSITORY,
        )


def test_release_validation_rejects_tag_behind_forgejo_main(tmp_path: Path) -> None:
    project, _, release_commit = _release_fixture(tmp_path)
    (project / "post-release.txt").write_text("new main commit\n", encoding="utf-8")
    _git(project, "add", "post-release.txt")
    _git(project, "commit", "-m", "Post-release change")
    _git(project, "push", "origin", "main")

    with pytest.raises(ReleaseMirrorError, match="not the current Forgejo main tip"):
        validate_release(
            project_root=project,
            tag="v1.2.3",
            event_sha=release_commit,
            forgejo_repository=EXPECTED_FORGEJO_REPOSITORY,
        )


def test_release_validation_rejects_checkout_other_than_tag(tmp_path: Path) -> None:
    project, _, release_commit = _release_fixture(tmp_path)
    (project / "local-only.txt").write_text("different checkout\n", encoding="utf-8")
    _git(project, "add", "local-only.txt")
    _git(project, "commit", "-m", "Local-only commit")

    with pytest.raises(ReleaseMirrorError, match="release checkout is at"):
        validate_release(
            project_root=project,
            tag="v1.2.3",
            event_sha=release_commit,
            forgejo_repository=EXPECTED_FORGEJO_REPOSITORY,
        )


def test_release_mirror_rejects_divergent_github_main(tmp_path: Path) -> None:
    project, github, release_commit = _release_fixture(tmp_path)
    context = validate_release(
        project_root=project,
        tag="v1.2.3",
        event_sha=release_commit,
        forgejo_repository=EXPECTED_FORGEJO_REPOSITORY,
    )

    github_work = tmp_path / "github-work"
    _git(tmp_path, "clone", "--branch", "main", str(github), str(github_work))
    _git(github_work, "config", "user.name", "GitHub-only Test")
    _git(github_work, "config", "user.email", "github-only@example.test")
    _write_version(github_work, "8.8.8")
    _commit(github_work, "Divergent GitHub commit")
    _git(github_work, "push", "origin", "main")

    with pytest.raises(ReleaseMirrorError, match="refusing a non-fast-forward mirror"):
        mirror_release(context, project_root=project, mirror_remote=str(github))


def test_forgejo_workflow_is_release_only_and_uses_scoped_ssh_secret() -> None:
    workflow = (
        Path(__file__).resolve().parents[1] / ".forgejo/workflows/release-mirror.yml"
    ).read_text(encoding="utf-8")

    assert '      - "v*"' in workflow
    assert "branches:" not in workflow
    assert "RELEASE_MIRROR_DEPLOY_KEY" in workflow
    assert "PYPI_API_TOKEN" not in workflow
    assert "--force" not in workflow
    assert "--mirror" not in workflow
    assert (
        "https://data.forgejo.org/actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1"
    ) in workflow
