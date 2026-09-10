#!/usr/bin/env python3
"""Limited cross-platform consistency checks; NOT a C# compiler or runtime test.
Requires: Python 3, pygments, json5, Node.js on PATH.
"""
import argparse
import collections
import csv
import datetime
import hashlib
import json
import pathlib
import platform
import re
import subprocess
import sys
import json5
from pygments import lex
from pygments.lexers import CSharpLexer
from pygments.token import Comment, Error, String


def check(condition, message):
    if not condition:
        raise AssertionError(message)


def cs_tokens(text):
    # Pygments misclassifies valid unicode-escape C# char literals.
    text = re.sub(r"'\\u[0-9a-fA-F]{4}'", "'X'", text)
    return list(lex(text, CSharpLexer()))


def check_cs(path):
    tokens = cs_tokens(path.read_text(encoding="utf-8-sig"))
    check(not [(str(k), v) for k, v in tokens if k in Error], "unexpected lexical token in " + path.name)
    stack = []
    for kind, value in tokens:
        if kind in Comment or kind in String:
            continue
        for char in value:
            if char in "{([":
                stack.append(char)
            elif char in "})]":
                check(stack and "{([".index(stack.pop()) == "})]".index(char), "delimiter mismatch in " + path.name)
    check(not stack, "unclosed delimiter in " + path.name)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=pathlib.Path, default=pathlib.Path(__file__).resolve().parents[1])
    parser.add_argument("--original", type=pathlib.Path, help="Optional unmodified extracted original root")
    args = parser.parse_args()
    root = args.root.resolve()
    results = []
    details = {}

    def test(name, action):
        try:
            action()
            results.append({"name": name, "status": "PASS"})
        except Exception as error:
            results.append({"name": name, "status": "FAIL", "detail": str(error)})
        print(results[-1]["status"], name, results[-1].get("detail", ""))

    sources = sorted((root / "src").glob("*.cs")) + [root / "tests" / "RegressionTests.cs"]
    for source in sources:
        test("csharp-lexical-delimiters:" + str(source.relative_to(root)), lambda p=source: check_cs(p))

    code_only = "\n".join("".join(v for k, v in cs_tokens(p.read_text(encoding="utf-8-sig"))
                                        if k not in Comment and k not in String) for p in sources)

    def text_members():
        definition = (root / "src" / "Rdv3Text.cs").read_text(encoding="utf-8-sig")
        refs = set(re.findall(r"\bRdv3Text\.(\w+)\b", code_only))
        declared = set(re.findall(r"\b(?:const\s+string|static\s+(?:readonly\s+)?string(?:\[\])?)\s+(\w+)", definition))
        check(not refs - declared, "undeclared text names: " + str(sorted(refs - declared)))
        constants = re.findall(r"\bconst\s+string\s+(\w+)", definition)
        check(len(constants) == len(set(constants)), "duplicate text constants")
        details["text_reference_count"] = len(refs)
    test("text-reference-names", text_members)

    def helpers():
        for name in ("Rdv3Files", "Rdv3Csv"):
            definition = (root / "src" / (name + ".cs")).read_text()
            declared = set(re.findall(r"\b(?:public|private)\s+static\s+\S+\s+(\w+)\s*\(", definition))
            refs = set(re.findall(r"\b" + name + r"\.(\w+)\b", code_only))
            check(not refs - declared, name + " missing names: " + str(refs - declared))
        check(all("System.IO.Rdv3Files" not in p.read_text(encoding="utf-8-sig") for p in sources), "invalid namespace substitution")
    test("new-helper-reference-names", helpers)

    def javascript():
        run = subprocess.run(["node", "--check", str(root / "web" / "app.js")], capture_output=True, text=True, timeout=30)
        check(run.returncode == 0, run.stdout + run.stderr)
        details["node_version"] = subprocess.check_output(["node", "--version"], text=True).strip()
    test("javascript-node-syntax", javascript)

    # Public, isolated generic fixture; root settings may be site-specific.
    fixture = root / "tests" / "fixtures"
    cfg = json5.loads((fixture / "settings.json").read_text(encoding="utf-8-sig"))
    def sample_config():
        check(cfg["schema"] == 3 and cfg["screen"]["workState"]["trigger"] == "manual", "schema/trigger")
        check(cfg["data"]["ledger"]["identity"] in cfg["data"]["ledger"]["columns"]["source"], "identity not persisted")
        ids = [job["id"] for job in cfg["data"]["jobs"]]
        check(len(ids) == len(set(ids)), "duplicate job id")
        details["native_test_cases_provided_not_executed"] = len(re.findall(r'\bTest\("', (root / "tests" / "RegressionTests.cs").read_text()))
    test("sample-jsonc-structure-and-safe-trigger", sample_config)

    def sample_data():
        data = cfg["data"]
        directory = fixture / cfg["paths"]["dataDir"]
        tables = {}
        info = {}
        for name, spec in data["tables"].items():
            with (directory / spec["file"]).open(encoding=data["encoding"], newline="") as stream:
                reader = csv.DictReader(stream)
                rows = list(reader)
                heads = reader.fieldnames
            check(heads and len(heads) == len(set(heads)) and spec["key"] in heads, name + " invalid header")
            check(all(None not in row and all(v is not None for v in row.values()) for row in rows), name + " column count")
            validation = spec.get("keyValidation", {})
            keys = [row[spec["key"]] for row in rows if row[spec["key"]] != "" or validation.get("empty") != "skip"]
            check(all(keys), name + " empty key")
            if validation.get("characters", "ascii") == "ascii":
                check(all(key.isascii() for key in keys), name + " non-ASCII key")
            if validation.get("length", "fixed") == "fixed":
                check(len(set(map(len, keys))) <= 1, name + " variable key length")
            if validation.get("duplicates", "error") == "error":
                check(len(keys) == len(set(keys)), name + " duplicate key")
            tables[name] = (heads, rows)
            info[name] = {"source_rows": len(rows), "keyed_rows": len(keys), "columns": len(heads), "file": spec["file"]}
        for ref in data["ledger"]["columns"]["source"]:
            table, column = ref.split(".", 1)
            check(table in tables and column in tables[table][0], "unknown persisted source column: " + ref)
        for ref in data["types"]:
            table, column = ref.split(".", 1)
            check(table in tables and column in tables[table][0], "unknown typed column: " + ref)
        for job in data["jobs"]:
            for spec in job["inputs"]:
                if "file" in spec:
                    with (directory / spec["file"]).open(encoding=data["encoding"], newline="") as stream:
                        reader = csv.DictReader(stream); rows = list(reader)
                        check(spec["column"] in reader.fieldnames, "missing file-only input column")
                        info[spec["file"]] = {"source_rows": len(rows), "columns": len(reader.fieldnames)}
        # Independent shape/reference checks, not execution of the C# merge engine.
        details["sample_data"] = info
    test("supplied-csv-shape-keys-source-references", sample_data)

    if args.original:
        def preserved_files():
            original = args.original.resolve()
            old = [p for p in original.rglob("*") if p.is_file()]
            missing = [str(p.relative_to(original)) for p in old if not (root / p.relative_to(original)).is_file()]
            check(not missing, "missing original files: " + str(missing))
            protected = [p for p in old if p.relative_to(original).parts[0] in ("lib", "data") or p.name in ("LICENSE", "THIRD-PARTY-NOTICES.md")]
            changed = [str(p.relative_to(original)) for p in protected if p.read_bytes() != (root / p.relative_to(original)).read_bytes()]
            check(not changed, "original inputs/binaries/licenses changed: " + str(changed))
            details["original_file_count"] = len(old)
            details["preserved_input_binary_license_files"] = len(protected)
            details["original_zip_sha256"] = hashlib.sha256((original.parent.parent / "ReaderDataViewer.zip").read_bytes()).hexdigest() if (original.parent.parent / "ReaderDataViewer.zip").exists() else None
        test("original-file-completeness-and-input-binary-license-preservation", preserved_files)

    passed = sum(item["status"] == "PASS" for item in results)
    output = {"scope": "Lexical/static/sample-file checks only; NOT native compilation or C# execution",
              "utc": datetime.datetime.now(datetime.timezone.utc).isoformat(), "platform": platform.platform(),
              "passed": passed, "failed": len(results) - passed, "details": details, "tests": results}
    path = root / "tests" / "results" / "static-results.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
