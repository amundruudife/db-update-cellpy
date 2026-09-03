import ast
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
RETIRED_MODULES = (
    "cleanup_old_files.py",
    "common_utils.py",
    "config.py",
    "copy_sharepoint_file.py",
    "data_processing.py",
    "database.py",
    "exceptions.py",
    "file_operations.py",
    "logging_utils.py",
)


def test_retired_modules_and_pandas_runtime_dependency_are_absent():
    assert all(not (REPOSITORY_ROOT / "src" / module).exists() for module in RETIRED_MODULES)
    requirements = (REPOSITORY_ROOT / "requirements.txt").read_text(encoding="utf-8")
    assert not any(line.strip().lower().startswith("pandas") for line in requirements.splitlines())
    assert 'pywin32>=312; sys_platform == "win32"' in requirements

    setup_tree = ast.parse((REPOSITORY_ROOT / "setup.py").read_text(encoding="utf-8"))
    setup_call = next(
        node
        for node in ast.walk(setup_tree)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id == "setup"
    )
    install_requires = next(
        keyword.value
        for keyword in setup_call.keywords
        if keyword.arg == "install_requires"
    )
    assert ast.literal_eval(install_requires) == [
        "openpyxl>=3.1.5",
        "keyring>=24.0",
        'pywin32>=312; sys_platform == "win32"',
    ]


def test_repository_contains_no_sharing_token_resolver_contract():
    forbidden = ("/" + "shares" + "/", "u" + "!", "share" + "_" + "token")
    for path in REPOSITORY_ROOT.rglob("*"):
        if not path.is_file() or any(part in {".git", "graphify-out", ".pytest_cache"} for part in path.parts):
            continue
        if path.suffix.lower() not in {".py", ".md", ".txt", ".yaml", ".yml", ".json"}:
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        assert not any(marker in text for marker in forbidden), path
