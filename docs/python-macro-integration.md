# Optional workbook Python macros

> **Last reviewed:** 2026-10-17 · **Last updated:** 2026-10-17

Desktop implementation of the workbook-project portion of [#1566](https://github.com/storytold/gridcraft/issues/1566), built on the shared plugin host in [#1123](https://github.com/storytold/gridcraft/pull/1123), commit `3e325c1c2c2e35cf7acc6c01bda8b326739ec0a3`. It requires the companion optional role in [sirmmo/craft-python](https://github.com/sirmmo/craft-python), based on 0.3.0 (`f3a04b8`). It does not replace craft-python's Run Script, Console, Run File or `PY.EVAL`, or the Python-in-cells implementation in #469.

## Scope

A workbook holds a validated, versioned Python project: up to 128 modules, 1 MiB per module, 4 MiB total source and an 8 KiB PEP 723 comment block. The shared Developer editor lists, creates, edits and deletes modules; its Python Settings dialog edits interpreter/dependency metadata. Top-level public synchronous functions without arguments are macros. Discovery parses source and does not execute it.

The versioned JSON project has a separate XLSX package part. VBA's `vbaProject.bin` is preserved independently. Excel cannot execute these Python macros. Macro execution remains disabled until the user enables macros for the document.

## Install the development provider

Use a build of this host branch and the companion craft-python branch. Install Python and uv separately; neither CPython nor scientific packages are bundled in GridCraft.

Load craft-python's `plugin-macros.json` with **File > Options > Add-ins > Load** (select the manifest explicitly), or run `plugins.load {"path":"/absolute/path/to/craft-python/plugin-macros.json"}`. Its `run.command` defaults to `python3`; on Windows, set that field to the installed Python executable if necessary. Optional launcher arguments are `--uv-executable`, `--python-interpreter` and `--macro-cache`.

The ordinary `plugin.json` continues to load craft-python's scripts, console, functions and WordCraft/DeckCraft support. The macro role has a separate ID and registers no duplicate functions or ribbon commands. Load the ordinary role too when workbook formulas use `PY.*`: the host cannot re-enter a busy macro provider, but can call the separate ordinary provider during recalculation. Both roles reuse craft-python's App and Channel.

Open the shared Developer editor, choose the Python provider and create a module, for example:

```python
import gridcraft as gc

def double_sales():
    wb = gc.active_workbook()
    values = wb.read("A2:A101")
    wb.write("B2", [[row[0] * 2] for row in values])
```

Run `analysis.double_sales` from Macros. `wb.write` writes literal values: formula-looking and date-looking strings stay text. Use `wb.execute("cell.set", cell="B2", input="=A2*2")` for an explicit formula. The bridge also offers `wb.cell`, `wb.inspect` and a bounded set of engine editing commands; external file/app operations and recursive macro/plugin execution are excluded from this transaction bridge. Trusted Python itself still has the user's OS permissions.

## Additive host interface

Providers require `pythonProjects`, `engineCommands` and `literalValues`. An unmodified #1123 host refuses the manifest as unsupported. Existing protocol 2.2 capabilities and VBA/Office.js methods remain intact.

| Method | Params | Result |
| --- | --- | --- |
| `host/pythonProject/get` | `{}` | `{documentId, project, allowDownloads}`; document identity and validated project, or null project |
| `host/pythonProject/set` | `{project}` | Validates and replaces the project; identical source is a no-op |
| `host/engine.execute` | `{command, params}` | Result of an allowed engine command |

Macro manifests can opt out of recording or Office ribbon callbacks with `recording:false` and `ribbonCallbacks:false`. Missing flags preserve existing providers' behavior. Python opts out of both.

## Runtime, transactions and consent

Only running a macro starts its uv-managed worker. The provider prepares a PEP 723 environment and reuses its interpreter for the same document and exact source. A source/document change or failure discards the worker; after success, 60 idle seconds retire it. Dependencies remain cached. Windows kill-on-close jobs own worker descendants; POSIX uses a small pipe-lifetime helper that kills the worker group even after abrupt provider termination; the helper blocks without polling and retires with the worker. Native POSIX lifecycle tests are in the companion provider suite.

Startup and execution share a 30-second absolute budget. Pipe messages are bounded at 16 MiB and startup diagnostics at 8 KiB; continuing output cannot extend the deadline. Host requests run on the Session's owner thread through bounded queues. Stop/Escape cancels the transport without joining a worker thread in the UI.

One successful run is one undo entry. An error, Stop or timeout restores the workbook, bounded Undo/Redo history, selection and session download consent. Nested host edits do not add journal/history steps. Conflicting edits, saves and exports are rejected while running; plugin pushes, invalidation, file arrivals and autosave are deferred. Switching/closing a document cancels and restores first. Headless callers use synchronous `developer.runMacro`; UI/control owners use `developer.startMacro`, `developer.pollMacro` and `developer.cancelMacro`. The async path refuses live collaboration; the UI keeps VBA and other providers on their original synchronous, activity-based timeout path; only Python opts into the new UI lifecycle.

Python Settings grants uv download consent only for the exact metadata in the current document session. It is absent from serialized workbook/XLSX data and cannot be forged inside a project. Editing metadata revokes it. Without consent, preparation uses `--offline --no-python-downloads`; with consent those restrictions are lifted. `--no-build --no-config --no-project --no-env-file` remain, and explicit interpreter pins still apply. uv reports dependency/TOML errors; core validation bounds metadata and rejects executable text outside comments. This is dependency-preparation consent, not a network or OS sandbox.

## Verification and remaining scope

Native integration uses either the minimal protocol fixture or the real companion provider selected by `GRIDCRAFT_TEST_PYTHON_PLUGIN`. `GRIDCRAFT_TEST_PYTHON` and `GRIDCRAFT_TEST_UV` are required; missing tools fail rather than silently skip. Tests cover inert discovery, editing, literal writes, undo/rollback, no-op history, warm reuse, source/document isolation, deadlines, cancellation and ordinary `PY.TWICE` recalculation. Real-process transport regressions cover silence, output flooding, initialization, late replies, busy providers and repeated Stop. Offscreen UI tests exercise Run/Stop/Escape and Python Settings in both themes; rendered snapshots were inspected locally. GitHub CI prepares isolated Python and a hash-pinned uv wheel.

This is a desktop phase, not completion of #1566. Form controls, Pyodide, a language server, module-file import/export, rich output/traceback navigation and binary bulk transfer remain future work. Do not close #1566 or fold the unmerged #1123 implementation into a PR against main.
