import toml
from pathlib import Path
from xlsql.version import VERSION

project = Path("pyproject.toml")
xlsql = toml.loads(project.read_text())

if not VERSION == xlsql["tool"]["poetry"]["version"]:
    import sys

    print("Detected version mismatch between pyproject.toml and xlsql/version.py!")
    print(" ? ?")
    print(f"  0/     pyproject.toml: {xlsql['tool']['poetry']['version']}")
    print(f" <Y    xlsql/version.py: {VERSION}")
    print(f" / \\")
    sys.exit(1)
