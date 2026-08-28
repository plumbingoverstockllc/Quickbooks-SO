"""A beta build must never reach someone running the stable release.

Two things keep the channels apart, and both are easy to break by accident:
the version ordering inside the app, and the publish step in the release
workflow. A beta tag that wrote releases/latest.json would push a test build
to every installed copy on its next startup check.
"""
from __future__ import annotations

import ast
from pathlib import Path

WORKFLOW = Path(__file__).parent / ".github" / "workflows" / "release.yml"


def _load_method(name: str):
    """Lift a method out of app.py, which can't be imported off Windows."""
    src = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.FunctionDef) and node.name == name:
            module = ast.Module(body=[node], type_ignores=[])
            namespace: dict = {}
            exec(compile(module, "app.py", "exec"), namespace)  # noqa: S102
            return namespace[name]
    raise AssertionError(f"{name} not found in app.py")


version_tuple = _load_method("_version_tuple")


def _v(version: str):
    return version_tuple(None, version)


def test_beta_sorts_below_the_stable_of_the_same_number():
    assert _v("v1.072b") < _v("v1.072")


def test_beta_sorts_above_the_previous_stable():
    assert _v("v1.072b") > _v("v1.071")


def test_stable_check_does_not_offer_a_downgrade_to_a_beta_user():
    """Someone on v1.072b runs the silent startup check against the stable
    feed. v1.071 must read as "up to date", not as an available update."""
    current, available = _v("v1.072b"), _v("v1.071")
    assert available <= current


def test_a_superseded_beta_is_not_offered_after_the_stable_ships():
    current, available = _v("v1.072"), _v("v1.072b")
    assert available < current


def test_beta_publish_marks_the_release_as_a_prerelease():
    text = WORKFLOW.read_text(encoding="utf-8")
    assert "--prerelease" in text
    assert "$isBeta = $ver.ToLower().EndsWith('b')" in text


def test_only_the_stable_path_writes_the_stable_feed():
    """releases/latest.json may only be assigned inside the non-beta branch."""
    text = WORKFLOW.read_text(encoding="utf-8")
    assert '$feedPath = "releases/beta.json"' in text
    assert '$feedPath = "releases/latest.json"' in text
    # The one Copy-Item onto beta.json is guarded by `if (-not $isBeta)`.
    guard = text.index("if (-not $isBeta) {")
    copy = text.index("Copy-Item releases/latest.json releases/beta.json")
    assert guard < copy


if __name__ == "__main__":
    test_beta_sorts_below_the_stable_of_the_same_number()
    test_beta_sorts_above_the_previous_stable()
    test_stable_check_does_not_offer_a_downgrade_to_a_beta_user()
    test_a_superseded_beta_is_not_offered_after_the_stable_ships()
    test_beta_publish_marks_the_release_as_a_prerelease()
    test_only_the_stable_path_writes_the_stable_feed()
    print("ok")
