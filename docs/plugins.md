# Plugins

> **Last reviewed:** 2026-10-16 · **Last updated:** 2026-10-17 · **Change:** minor (additive workbook Python capabilities and cancellable macro lifecycle) · **Target:** Microsoft Excel (Microsoft 365)

A plugin adds **functions** that formulas call and **commands** the user (or an agent) runs,
and can put its commands on the ribbon (a group on a built-in tab, or a tab of its own) and in
the right-click menus.
It's a separate program, written in any language, that GridCraft starts and talks to over its
standard input and output. This is GridCraft's equivalent of Excel's add-ins (custom functions,
XLL and COM add-ins). VBA stays out of scope.

Because a plugin runs in its own process, a crash, a hang or a bad answer costs a cell an error
value, never the workbook. The host's code is `crates/plugin`; the reference plugin is
`crates/plugin/src/bin/example_plugin.rs`.

**One plugin, several apps.** The manifest and the protocol are shared with WordCraft, so one
plugin can serve both: `apps` names the apps it's for, `uiFor` gives each app its own ribbon and
menus, and `initialize` says which app started it (`host.id`). WordCraft plugins have commands
only; functions are GridCraft's. A test plugin written in Python that works in both is
[sirmmo/gridcraft-python-plugin](https://github.com/sirmmo/gridcraft-python-plugin).

**Office add-ins** (Office.js web add-ins with an XML manifest) load as plugins too: their
commands open their task pane in a web view; see [officejs.md](officejs.md).

**Co-editing** (protocol 2): transport plugins carry live sessions between people; see
[Protocol 2](#protocol-2-the-push-channel-and-co-editing) and [collab.md](collab.md).
**Protocol 2.2** adds what add-ins that run code from workbooks need (R, Python): see
[Protocol 2.2](#protocol-22-functions-that-know-their-workbook). Live data
from sources and streams comes later; see [gaps.md](gaps.md) #16 and #21.

## Installing a plugin

**From GitHub.** In **File › Options › Add-ins** (or **Automate › Add-ins**), paste the
repository's address in **Install from GitHub** and confirm. GridCraft downloads it in the
background, unpacks it into the plugins folder of the settings (as `<plugin id>/`, replacing an
earlier install of the same plugin) and loads it. Agents and the CLI use
`plugins.install {"url": …, "dir": …}`. Addresses it understands:

| Address | Installs |
|---|---|
| `https://github.com/owner/repo` (or `github.com/owner/repo`, `owner/repo`) | The latest release; the default branch's source when there is no release |
| `…/releases/tag/v1.2.0` | That release |
| `…/tree/<branch, tag or commit>` | That source |

From a release, GridCraft takes the `.zip` asset built for this computer when there is one
(see [Publishing a plugin](#publishing-a-plugin-on-github)), else the release's source code. The
plugin is the folder in the zip that holds `plugin.json` (the shallowest one). Downloads are
limited to 256 MiB and 512 MiB unpacked; paths that would leave the plugin's folder and symbolic
links are skipped. GitHub allows 60 anonymous requests an hour per address; private repositories
can't be installed. `.craft-source.json` in the plugin's folder records where it came from,
and Options shows the repository.

**From a folder.** A plugin is a folder holding a `plugin.json` manifest and whatever the plugin
needs to run. GridCraft loads plugins from:

1. each folder in `GRIDCRAFT_PLUGIN_PATH` (separated like `PATH`);
2. `--plugin DIR` on the command line (`gridcraft` and every `gridcraft-cli` subcommand; repeatable);
3. (desktop app only) the `plugins` folder in the settings folder: `~/.config/gridcraft/plugins`
   on Linux and BSD, `~/Library/Application Support/GridCraft/plugins` on macOS,
   `%APPDATA%\GridCraft\plugins` on Windows.

A folder holding `plugin.json` is one plugin. Otherwise, each of its subfolders that holds one is
a plugin. Loading only reads the manifest. The plugin's process starts the first time one of its
functions or commands is used.

**File › Options › Add-ins** lists the installed plugins with their version, author and status.
**Documentation** opens a plugin's web page (its `homepage`), and **Deactivate** turns a plugin
off: it stays listed, but its functions (cells show `#NAME?`), commands, ribbon buttons and menu
items are gone and its process stops. **Activate** brings it back. The app remembers which
plugins are off (`disabledPlugins` in `prefs.json`), so they stay off at the next start and when
loaded again.

**Automate › Add-ins** (also **Manage Add-ins…** in Options) lists the loaded plugins with their
functions, runs their commands, turns them off and on, and loads or unloads a plugin. Commands do
the same:

| Command | Params | Does |
|---|---|---|
| `plugins.list` (also `home.addIns`) | `{}` | The loaded plugins with their status, `state` (`loaded`, `running`, `failed`, `blocked`, `disabled`), `blockedBy`, `requires`, `requiredBy`, functions and commands |
| `plugins.load` | `{path, force?}` | Loads the plugin in a folder (or a manifest file). A plugin with the same id is replaced; with a version outside the range a turned-on plugin requires, only with `force` |
| `plugins.install` | `{url (or github), dir?, version?, withRequirements?, force?}` | Installs the plugin in a GitHub repository into `dir` (default: the app's plugins folder) and loads it. The result adds `release` and `missing`; see [Requirements](#requirements-manifest-13) |
| `plugins.installMissing` | `{id, dir?, force?}` | Installs the missing required add-ins of `id` that name `github` |
| `plugins.requirements` | `{id}` | `{id, satisfied, blockedBy, tree}`: what `id` needs and what's installed |
| `plugins.unload` | `{id}` | Unloads a plugin and stops its process (until the next start); `dependentsBlocked` lists the plugins that needed it |
| `plugins.setEnabled` | `{id, enabled}` | Turns a plugin off or back on, and remembers it in the preferences; turning off reports `dependentsBlocked` |
| `plugins.openHomepage` | `{id}` | Opens the plugin's web page in the browser |
| `plugins.run` | `{plugin, command, params?}` | Runs a plugin command (one undoable step) |

Open workbooks that call a plugin's functions are recalculated when it loads or unloads.

```sh
gridcraft-example-plugin --manifest > ~/plugins/example/plugin.json
gridcraft-cli --plugin ~/plugins/example eval '=EXAMPLE.HELLO("world")'     # Hello, world!
gridcraft-cli --plugin ~/plugins/example run --sample sales --quiet \
  --cmd 'selection.set={"range":"K1"}' --cmd 'plugins.run={"plugin":"org.gridcraft.example","command":"countUp","params":{"count":5}}' --print K1:K5
```

**Trust.** A plugin is a program running with your permissions, like any app you install.
Install plugins only from people you trust. GridCraft doesn't sandbox them.

## The manifest

```json
{
  "id": "org.example.finance",
  "name": "Finance tools",
  "version": "1.2.0",
  "description": "Prices and rates",
  "author": "Example Corp",
  "homepage": "https://example.org/finance-plugin",
  "apps": ["gridcraft"],
  "run": {"command": "python3", "args": ["main.py"]},
  "functions": [
    {
      "name": "FIN.PRICE",
      "signature": "FIN.PRICE(ticker, [date])",
      "description": "Closing price",
      "minArgs": 1,
      "maxArgs": 2,
      "volatile": false,
      "timeoutMs": 5000
    }
  ],
  "commands": [
    {"id": "refreshAll", "label": "Refresh Prices", "description": "…", "params": "{sheet?}"}
  ],
  "ui": {
    "ribbon": [
      {"tab": "Finance", "group": "Prices", "buttons": [
        {"command": "refreshAll", "icon": "refresh"},
        {"command": "refreshAll", "label": "This Sheet", "size": "small", "params": {"sheet": "current"}}
      ]},
      {"tab": "Data", "group": "Finance", "buttons": [{"command": "refreshAll", "icon": "currency"}]}
    ],
    "menus": [
      {"menu": "cell", "command": "refreshAll", "label": "Refresh Prices Here"},
      {"menu": "sheetTab", "command": "refreshAll"}
    ]
  }
}
```

| Field | Meaning |
|---|---|
| `id` | Unique id: 1–128 letters, digits, `.`, `-`, `_`. Reverse-domain style is recommended |
| `apps` | The apps the plugin is for (`gridcraft`, `wordcraft`; any case). Left out: any app. A plugin made only for other apps isn't loaded |
| `author`, `homepage` | Shown in Options › Add-ins. `homepage` is an `http://` or `https://` link to the plugin's documentation |
| `run.command`, `run.args` | The program to start. A command containing `/` or `\` is relative to the plugin folder; a bare name is looked up on `PATH`. It runs in the plugin folder, with `GRIDCRAFT_PLUGIN_DIR`, `CRAFT_PLUGIN_DIR` (the folder), `CRAFT_HOST` (`gridcraft`) and `GRIDCRAFT_VERSION` set |
| `functions[].name` | The name formulas use (case-insensitive). It starts with a letter and holds letters, digits, `.` and `_`. It can't be a built-in function or another plugin's function. Prefix it with your own namespace (`FIN.`) |
| `minArgs`, `maxArgs` | Argument count (default 0 and 255). Other counts give `#VALUE!` without calling the plugin |
| `volatile` | Recalculate on every recalculation, like `RAND`. Otherwise a result is reused for the same arguments while the plugin is loaded (up to 10,000 results), and the cell is recalculated only when its precedents change |
| `timeoutMs` | How long a call may take: default 2,000, at most 30,000 |
| `commands[]` | `id` (unique in the plugin), `label`, `description`, `params` (documentation) |
| `uiFor.<app>` | Ribbon groups and menu items for one app: `uiFor.gridcraft` is used instead of `ui` when it's there |
| `ui.ribbon[]` | A group of buttons: `tab`, `group`, `buttons`. A built-in tab's English name (`Home`, `Insert`, `Draw`, `Page Layout`, `Formulas`, `Data`, `Review`, `View`, `Automate`) adds the group at the end of that tab; any other name makes a new tab after Automate, shared by the plugins that name it |
| `buttons[]` | `command` (one of the plugin's `commands[].id`), `label` and `tooltip` (default: the command's label and description), `icon`, `size` (`large`, the default, or `small`: stacked three high), `params` (passed to the command) |
| `events` | The document events the plugin wants (protocol 1.2): `documentOpened`, `documentSaved`, `documentClosed`. Other names are ignored |
| `ui.menus[]` | A right-click menu item: `menu` (`cell`, the default; `row` or `column` for the headers; `sheetTab`), `command`, `label`, `params`. Items go at the end of the menu, after a separator. A `sheetTab` item also gets `sheet` (the tab's index) in its params |
| `requires.apps` | Manifest 1.3. Per app id, the app versions the plugin works with (`{"gridcraft": ">=0.7.0"}`); an app missing from the map isn't constrained. See [Requirements](#requirements-manifest-13) |
| `requires.protocol` | The plugin protocol versions it needs (`">=1.2"`), checked against `host.protocol` (`"2.2"`) |
| `requires.features` | Names from `host.features` (`dialogs`, `open`, `events`, `document`, `push`, `collab`; protocol 2.2: `functionDocument`, `openedBeforeCalc`, `multiline`, `emptyResult`, `literalValues`, `invalidate`) that must all be present |
| `requires.addins[]` | Other plugins it needs: `id`, `version` (a range; default `*`), `github` (`owner/repo` to install it from), `optional` (default false), `why` (shown to the person, at most 200 characters). At most 32 |

Icons are named from GridCraft's own icon set: `script` (the default), `refresh`, `calc`, `sum`,
`function`, `fx`, `table`, `pivotTable`, `chart`, `chartBar`, `chartLine`, `chartPie`,
`chartScatter`, `sparkline`, `picture`, `link`, `comment`, `note`, `filter`, `sort`, `sortDesc`,
`find`, `search`, `lock`, `settings`, `share`, `sheet`, `workbook`, `folder`, `save`, `print`,
`calendar`, `clock`, `info`, `check`, `plus`, `minus`, `pen`, `currency`, `percent`,
`validation`, `name`, `zoom`, `more`, `copy`, `cut`, `paste`, `insert`, `delete`, `format`.
An unknown name gets the default. Plugins can't ship their own icon images yet. Labels are at
most 64 characters. Buttons and menu items run the command exactly as `plugins.run` does, as one
undoable step.

## Protocol (version 1)

JSON-RPC 2.0, **one JSON message per line**, UTF-8. GridCraft writes requests to the plugin's
standard input; the plugin writes one response per request to standard output. Lines on
standard output that aren't JSON are logged and ignored. Standard error goes to GridCraft's log.
The plugin exits when its standard input closes, or when it receives the `shutdown`
notification.

### `initialize`

Sent once, after the process starts. Answer within 10 s with any result, for example
`{"protocolVersion": 1}`.

```json
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":1,"host":{"id":"gridcraft","name":"GridCraft","version":"0.6.1"},"plugin":{"id":"org.example.finance","dir":"/home/me/plugins/finance"}}}
```

Since manifest 1.3, `host.protocol` is the protocol version as text (`"2.2"`, what
`requires.protocol` is checked against), and `addins` lists every installed plugin as
`{id, version, enabled}` (`enabled`: turned on, not blocked and not failed), so a plugin can adapt
to the optional ones it found.

### `functions/call`

```json
{"jsonrpc":"2.0","id":2,"method":"functions/call","params":{"name":"FIN.PRICE","args":["ACME",[[1,2],[3,null]]],
  "document":{"path":"/home/me/Prices.xlsx","name":"Prices.xlsx","id":"doc-1"}}}
{"jsonrpc":"2.0","id":2,"result":{"value":42.5}}
```

`document` (protocol 2.2) is the workbook whose formula calls the function; see
[Protocol 2.2](#protocol-22-functions-that-know-their-workbook).

Arguments arrive evaluated. A reference arrives as the array of its values. Values on the wire:

| Cell value | JSON |
|---|---|
| empty | `null`. As a result, it shows as 0, as Excel shows a VBA function that returns `Empty`; `{"empty": true}` (protocol 2.2) shows blank |
| number | number |
| text | string (up to 32,767 characters) |
| logical | `true` / `false` |
| error | `{"error": "#N/A"}` (any Excel error text) |
| array | rows of values: `[[1, 2], [3, 4]]`. A result may also be a flat list (one row). Ragged rows are padded with `#N/A`. An array result spills |

A call carries at most 1,000,000 cells, and a result may hold as many. Anything larger gives
`#VALUE!` (arguments) or `#NUM!` (result).

While a formula is typed, plugin functions are suggested with the built-in ones, with their
`signature` as the argument hint and their `description` beside the list; Insert Function and
`formulas.functions` list them under the category **Add-ins**. A function without
a `signature` shows as `NAME(…)`. Turning a plugin off takes its functions out of all of these.

### `commands/run`

```json
{"jsonrpc":"2.0","id":3,"method":"commands/run","params":{"command":"refreshAll","params":{},
  "context":{"workbook":"Book1","sheet":"Sheet1","active":"B2","selection":"B2:C3","values":[[1,2],[3,4]]}}}
{"jsonrpc":"2.0","id":3,"result":{"commands":[{"command":"range.setValues","params":{"range":"B2","values":[[5],[6]]}}],"message":"Updated 2 prices","result":{"any":"json"}}}
```

`context.values` and `context.formulas` (each cell's formula text such as `"=B2*1.1"`, or `null`
for a cell without a formula) are included for selections up to 100,000 cells. Every command also gets
`context.activeSheet`, `context.sheets` (`[{index, name, visibility, usedRange}]`, every sheet including
hidden ones) and `context.names` (`[{name, scope, hidden, refersTo}]`, at most 10,000, with
`namesTruncated: true` when there were more). The plugin answers with engine
commands to run. They are the same commands the CLI, MCP and control channel use (see
`gridcraft-cli commands`), except `plugins.*`. GridCraft runs them in order, as **one undoable
step**: if any fails, the workbook goes back to how it was and the error is shown. `message` is
shown to the user, and `result` is passed back to the caller of `plugins.run`. A command may
take up to 60 s.

### Errors and failures

- An answer `{"error": {"code": …, "message": "…"}}` gives the cell `#VALUE!`, or fails the
  command with that message. The process keeps running.
- A call that times out, a process that exits, or broken output gives `#VALUE!`. The process is
  stopped and started again on the next call.
- After 3 such failures the plugin is marked `failed` and its functions give `#VALUE!` until
  it's reloaded (`plugins.load`).

### Writing a plugin

Any language that reads and writes lines works. A complete Python plugin with one function:

```python
import json, sys

for line in sys.stdin:
    msg = json.loads(line)
    if "id" not in msg:          # notifications: shutdown
        break
    method, params = msg["method"], msg.get("params", {})
    if method == "initialize":
        result = {"protocolVersion": 1}
    elif method == "functions/call" and params["name"] == "PY.TWICE":
        x = params["args"][0]
        result = {"value": x * 2 if isinstance(x, (int, float)) else {"error": "#VALUE!"}}
    else:
        print(json.dumps({"jsonrpc": "2.0", "id": msg["id"], "error": {"code": -32601, "message": "unknown"}}), flush=True)
        continue
    print(json.dumps({"jsonrpc": "2.0", "id": msg["id"], "result": result}), flush=True)
```

```json
{"id": "org.example.py", "name": "Python example", "run": {"command": "python3", "args": ["plugin.py"]},
 "functions": [{"name": "PY.TWICE", "minArgs": 1, "maxArgs": 1}]}
```

## Protocol 1.1: asking the host, and macros

**The plugin asks the host while it works.** While GridCraft waits for its answer to a request
(`commands/run`, `macros/run`…), a plugin may send requests of its own on standard output — JSON-RPC
requests with its own ids — and read GridCraft's answers on standard input before it goes on. Each
answer starts the request's timeout again. Everything these requests change is part of the same
undoable step, and is taken back if the plugin's answer is an error.

| Request | Params | Answer |
|---|---|---|
| `host/sync` | `{actions}` | `{results}`, or an error whose `data` is `{code, message, action}`. The actions are the Office bridge's ([officejs.md](officejs.md#the-bridge-officesync)): paths from `workbook` to worksheets, ranges, formats; `load`, `set`, `call` |
| `host/vbaProject/get` | `{}` | `{base64}`: the workbook's `vbaProject.bin`, or `null` |
| `host/vbaProject/set` | `{base64}` | `{}`: replaces it (the workbook then saves as `.xlsm`); at most 64 MiB |
| `host/message` | `{text, title?}` | `{result: 1}`: shows a message after the step (it doesn't wait for the user yet) |
| `host/log` | `{text}` | `{}`: writes to GridCraft's log |

A plugin calling the host can make GridCraft recalculate a formula that calls one of the plugin's
own functions; that call fails (`#VALUE!`) instead of waiting for the busy plugin.

**Macros.** A plugin that declares `"macros": {"language": "VBA"}` runs macros, and brings the
**Developer** tab: Visual Basic (the workbook's code modules in an editor), Macros (run, edit,
create), Record Macro, Macro Security and Add-ins. GridCraft draws all of it; the plugin owns the
language. It answers:

| Request | Params | Answer |
|---|---|---|
| `macros/list` | `{}` | `{macros: [{name: "Module1.Hello", module, procedure}]}` |
| `macros/run` | `{name, control?: {id, tag}}` | `{}`, or an error (`Run-time error '13': Type mismatch`, with `data: {module, line}`) |
| `macros/modules` | `{}` | `{modules: [{name, kind: standard\|class\|document, source}]}` |
| `macros/saveModule` | `{name, source}` | `{}` |
| `macros/deleteModule` | `{name}` | `{}` |
| `macros/record` | `{steps: [{command, params}]}` | `{source}`: the recorded steps as a macro |

Commands: `developer.macros`, `developer.runMacro {name}`, `developer.modules`,
`developer.saveModule {name, source}`, `developer.deleteModule {name}`, `developer.recordMacro {stop?}`,
`developer.recordToCode {steps}`, `developer.enableMacros`, `developer.macroSecurity {level}`.

**Ribbon buttons** (Developer › Ribbon Buttons, or Add to Ribbon in the Macros window): buttons
that run the workbook's macros, on a built-in tab or one of the workbook's own. They are saved
with the workbook as Excel's ribbon part (`customUI/customUI14.xml`), so Excel shows them too, and
a workbook made in Excel brings its buttons along. A ribbon that holds more than tabs, groups and
buttons (pictures of its own, menus, `onLoad`) is kept unchanged: its buttons show and run but
can't be edited. A button runs its macro as Excel does, with the button: `macros/run` gets
`control: {id, tag}`, which a VBA add-in passes as the `IRibbonControl` argument
(`Sub Go(control As IRibbonControl)`); a macro without parameters runs as it is. Commands:
`developer.ribbon`, `developer.addRibbonButton {macro, label?, tab?, group?, image?, large?}`,
`developer.editRibbonButton {id, …}`, `developer.removeRibbonButton {id}`,
`developer.pressRibbonButton {id}`.

**Saving:** a workbook with macros is saved as `.xlsm` (a template `.xltx` as `.xltm`) when it is
saved under an `.xlsx` name, because an `.xlsx` can't hold macros; `file.save` reports the name it
used (`renamedFrom` holds the one asked for) and the message says so. The `.xlsx` isn't written.

**Macro security** (Developer › Macro Security, `macroSecurity` in the preferences): `prompt`, the
default, runs a workbook's macros once the user enabled them for it (the message bar's **Enable
Content** on a workbook with macros, `developer.enableMacros`); `disable` runs none; `enable` runs
all. A macro is a program: agents and the CLI are bound by the same setting.

## Protocol 1.2: dialogs, opening files, document events

Protocol 1.2 adds to 1.1 and changes nothing that exists; WordCraft and DeckCraft speak it too, so
one plugin works the same in all three. A plugin can show a dialog GridCraft draws without
blocking the app, open a local file in GridCraft, and hear when workbooks open, save and close.
[craft-nextcloud](https://github.com/storytold/craft-nextcloud) is the first plugin that uses it.

**Detection.** `initialize`'s `host` gains `features`, `["dialogs", "open", "events", "document"]`
(later versions may add more), and `language`, the interface language as a BCP 47 tag (`it`,
`pt-BR`), so the plugin can translate its labels. Check for a feature before relying on it.
`protocolVersion` stays `1`.

```json
"host": {"id": "gridcraft", "name": "GridCraft", "version": "0.7.0", "features": ["dialogs", "open", "events", "document"], "language": "it"}
```

**`context.document`.** `commands/run` and `macros/run` get the workbook's file in `context`:

```json
"document": {"path": "/home/me/Book1.xlsx", "name": "Book1.xlsx", "dirty": false, "id": "doc-1"}
```

`path` is the local file it was opened from or last saved to (absolute), or `null` for a new
workbook that was never saved; `dirty` says whether it has unsaved changes. `macros/run` gets
`context` with `workbook`, `sheet`, `active`, `selection` and `document` (not the selection's
values: macros read the workbook through `host/sync`).

### Opening a file: `open`

A `commands/run` result may hold `open`, besides `commands`, `message` and `result` (all may come
together):

```json
"open": {"path": "/home/me/.local/share/craft-nextcloud/files/ab12/Documents/Budget.xlsx",
         "label": "Nextcloud · Documents/Budget.xlsx", "readOnly": false}
```

The path must be absolute, an existing regular file, and have an extension GridCraft opens
(`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.xlsb`, `.ods`, `.csv`, `.tsv`, `.txt`, `.json`);
otherwise the command fails with a message and nothing changes. When a ribbon button, a menu item,
Automate › Add-ins or a plugin dialog ran the command, GridCraft opens the file after the step,
exactly as File › Open does (a workbook already open is brought to the front). `label` (at most 200
characters) shows in the status bar while that workbook is active. GridCraft has no read-only mode
yet, so `readOnly` is ignored. `plugins.run` from the CLI, MCP or the control channel opens
nothing: it returns `open` in its result for the caller to act on.

### Dialogs: `dialog`

```json
"dialog": {
  "title": "Open from Nextcloud",
  "kind": "list",
  "prompt": "Documents",
  "items": [
    {"id": "/Documents/Reports", "label": "Reports", "detail": "Folder", "icon": "folder"},
    {"id": "/Documents/Budget.xlsx", "label": "Budget.xlsx", "detail": "12 KB · 2026-10-14 18:02", "icon": "workbook"}
  ],
  "command": "browse",
  "params": {"account": "me@cloud.example"},
  "okLabel": "Open",
  "buttons": [{"label": "Up", "command": "browse", "params": {"path": "/"}}]
}
```

After the step, GridCraft shows a modal dialog it draws itself, with its own widgets and theme.
Nothing waits: the plugin has already answered. What the user does runs another of the plugin's
commands exactly as a ribbon button does (one step; its result may carry another `dialog`, an
`open`, a `message` or commands, so a plugin can walk folders one dialog at a time). A message
that comes with a dialog shows in the status bar, so it doesn't cover the dialog.

| `kind` | Shows | OK runs `command` with `params` plus |
|---|---|---|
| `list` | `items` (at most 10,000; `id` up to 4,096 characters, `label` and `detail` up to 256) with their icon (the ribbon's icon names; others get a page), label and detail. A filter box (by label) above more than 20 items. `"multi": true` allows several | `"id"` (double-click also runs it), or `"ids": [...]` with `multi` |
| `form` | `fields` (at most 16): `{id, label, value?, kind?}` with `kind` `text` (the default), `secret` (a password box), `number`, `checkbox` (`value` true or false) or `choice` (`options: [{id, label}]`, `value` an option id). Protocol 2.2: a `text` field with `multiline: true` is a multi-line editor (`code: true` for code: monospace) | `"fields": {"<id>": value, …}` (a number field that isn't a number gives `null`) |
| `message` | `prompt` (several lines, at most 4,000 characters) | nothing; without `command`, OK just closes |

Every kind has `title` (at most 200 characters), `prompt` (text above the content), `okLabel`
(default `OK`) and up to 4 `buttons`, `{label, command, params?}`, left of OK. A button runs its
command with its own `params` plus, for a list, the selected `id`/`ids` when something is selected
and, for a form, `fields`. Cancel and Escape close the dialog. Labels show as given: the plugin
translates them (see `host.language`). Anything over a limit is truncated; a malformed dialog, or
one whose OK or buttons name a command the plugin doesn't declare, fails the command with a
message. Programmatic runs don't show the dialog: `plugins.run` returns it (validated and
truncated) in its result.

### Document events

A plugin that declares `"events"` in its manifest gets them as JSON-RPC requests, after the event
happened, from a background thread:

| Request | Params | When |
|---|---|---|
| `events/documentOpened` | `{path, name, id}` | A workbook was opened from a local file (File › Open, Recent, drag and drop, a plugin's `open`, `file.open {path}`). Since protocol 2.2 it comes **before** the workbook's first calculation (see [Protocol 2.2](#protocol-22-functions-that-know-their-workbook)) |
| `events/documentSaved` | `{path, name, format, id}` | Save or Save As to a local file succeeded (`format`: the extension, lower case). Not AutoSave (`file.save {autosave: true}`), and not Save as PDF |
| `events/documentClosed` | `{path, name, id}` | A workbook that has a file was closed, or GridCraft quits with it open |

Paths are absolute. The answer is `{}` or `{"message": "Uploaded to Nextcloud"}`, shown in the
status bar for 10 seconds; an error answer shows in the info bar. `documentOpened` and
`documentSaved` time out after 30 s, `documentClosed` after 5 s; on quitting GridCraft waits at most
5 s for them (and the CLI waits for events still on their way before it exits). An event is never
part of an undo step, and a slow or failing plugin never stops the save, the open or the close
(since protocol 2.2 the opening calculation waits for `documentOpened`, at most 10 s per plugin);
timeouts count as failures as usual. During an event a plugin may send `host/log`,
`host/message` (shown in a message box) and `host/functions/invalidate`, but `host/sync` and `host/vbaProject/*` are refused with
an error whose `data.code` is `AccessDenied`: events are notifications, not edits.

The web app has no plugins, so none of this applies there. The dialog isn't a web view and can't
hold HTML.

## Protocol 2: the push channel and co-editing

[Craft plugin protocol 2](https://github.com/sirmmo/craft-collab/blob/main/docs/protocol-2.md)
(2.0.1, shared with WordCraft and DeckCraft) adds to 1, 1.1 and 1.2 and changes nothing in them:

- **The push channel.** GridCraft reads every running plugin's output all the time. A plugin may
  send notifications whenever it likes, and requests to the *anytime* methods (`collab/snapshot`,
  `host/log`); GridCraft answers them in its main loop, in any order, and may send the plugin
  notifications interleaved with its own requests, so a plugin must match answers by `id`. The
  protocol-1 rules hold for GridCraft's requests. A line may be 64 MiB; more than 1,000 messages
  a second for 5 s and the plugin is stopped. `initialize` gets `host.features` `"push"` and
  `"collab"` and `host.collab = {formats: ["yjs-v1"], app: "gridcraft", schema: 1,
  encryption: ["aes-256-gcm"]}`.
- **Transports.** A plugin whose manifest has `"collab": {"transport": "mqtt", "label": "Live
  (MQTT)"}` may join rooms. Its command answers `{"collab": {"join": {room, mode, user, role?,
  key?, path?, title?}}}` (alongside `message`, `dialog`, `open`); `plugins.run` returns it and
  the interface runs `collab.join` with it. Then the `collab/*` messages flow: `collab/joined`,
  `sync1`/`sync2`/`update`/`awareness` (Yjs, base64, encrypted when keyed), `collab/status`,
  `collab/leave`, `collab/snapshot`.

GridCraft's shared document (schema 1), commands and behaviour: [collab.md](collab.md). A test
transport that relays between two rooms is `crates/engine/tests/collab_push.rs`; craft-collab
has MQTT and loopback transports.

## Protocol 2.2: functions that know their workbook

Protocol 2.2 adds to 2.0 and 2.1 (the document chat, [collab.md](collab.md#document-chat)) and changes nothing that exists. It closes the gaps an add-in that runs
code from workbooks (R, Python) runs into: it can decide per workbook what its functions may do
(trust), and show code, blanks and text as people expect. Each addition has its name in
`host.features`; check for it, and keep working without it on older hosts. `host.protocol` is
`"2.2"` (manifests write `"requires": {"protocol": ">=2.2"}` or name the features).

| Feature | Adds |
|---|---|
| `functionDocument` | `functions/call` carries `document`: the calling workbook |
| `openedBeforeCalc` | `events/documentOpened` arrives before an opened workbook's first calculation |
| `multiline` | Multi-line text fields in form dialogs, for code too |
| `emptyResult` | `{"empty": true}`, a function result that shows blank |
| `literalValues` | `range.setValues {literal: true}`: text written as text, never read as a number, date or formula |
| `invalidate` | `host/functions/invalidate`: the plugin's kept results are forgotten and the formulas that call them recalculated |

### The calling workbook: `document`

```json
{"jsonrpc":"2.0","id":7,"method":"functions/call","params":{"name":"R.EVAL","args":["x*2",3],
  "document":{"path":"/home/me/Model.xlsx","name":"Model.xlsx","id":"doc-3"}}}
```

`path` is the local file the workbook was opened from or last saved to (absolute), `null` for a
workbook without one (a new Book1, an imported `.ods`, a sample); `name` is what the title bar
shows; `id` is an opaque text that stays the same for the workbook while GridCraft runs, saved or
not, and is never reused for another workbook in that run. So a plugin can key trust on `path`,
or on `id` for a workbook with no file. The same `id` is in `context.document` (`commands/run`,
`macros/run`) and in the document events, so a command (Trust This Workbook) and the formulas agree
on which workbook is meant. `document` is `null` in the rare evaluations that belong to no
workbook. A kept (non-volatile) result is kept per workbook and per file: the same arguments in
another workbook, or after Save As, call the plugin again.

### `documentOpened` before the first calculation

When a workbook is opened from a local file, GridCraft reads it, sends `events/documentOpened
{path, name, id}` to the plugins that declared it, **waits for each one's answer (at most 10 s)**,
and only then calculates the workbook. Guaranteed: no formula of a workbook opened from a file
calls a plugin function before every plugin that declared `documentOpened` answered that event or
timed out, and those calls carry the event's `id` (and `path`, unless the file was imported).
Events queued before it are delivered first. Not covered: a plugin whose own command opened the file
(`file.open` among its reply's commands) is busy, so it gets the event afterwards, as before; a
plugin that takes longer than 10 s (or fails) doesn't hold the workbook up, its timeout counts as a
failure, and the calculation goes ahead; and workbooks that aren't opened from a file (new, samples,
`file.open {base64}`) get no event. The plugin must not wait for the user while it answers.

### Multi-line fields

```json
{"id": "code", "label": "R code", "kind": "text", "value": "m <- lm(y ~ x)\nsummary(m)", "multiline": true, "code": true, "rows": 10}
```

`multiline: true` on a `text` field draws a multi-line editor (Enter starts a new line; it
scrolls past its height); `code: true` makes it monospace with Tab indenting, and implies
`multiline`. `rows` is its height in lines (2–30, default 8). Its value may hold up to 200,000
characters (a one-line field 32,767); the dialog gets wider for it. On other field kinds both are
ignored. Hosts without the feature ignore the keys and show a one-line field.

### Blank results: `{"empty": true}`

A function result `null` is an empty value, which the cell shows as **0**: that's what Excel does
with a VBA function that returns `Empty` (and with a reference to an empty cell), so it stays the
default. `{"empty": true}`, as the result or as values in an array result, shows **blank** when
the formula is the call itself (`=R.EVAL(…)`, not `=IFERROR(R.EVAL(…), 0)`). Other formulas read the
cell as empty (0 in arithmetic, `""` in text). In an array result that holds `{"empty": true}`,
the `null`s still show 0. Wrapped in another function, the blank is an empty value again (0).

### Literal text: `range.setValues {literal: true}`

```json
{"command": "range.setValues", "params": {"range": "B2", "values": [["00123", "=not a formula", "1/2"]], "literal": true}}
```

With `literal`, every string is stored as text exactly as given (cut at 32,767 characters): never
read as a number, a date, a percentage, a logical or a formula, and a leading `'` stays part of
the text. An empty string clears the cell. Numbers, logicals and `null` are written as without it.
This replaces the leading-apostrophe trick, which leaves the apostrophe in other hosts.

### Invalidation: `host/functions/invalidate`

```json
{"jsonrpc":"2.0","id":"p1","method":"host/functions/invalidate","params":{"names":["R.EVAL","R.CALL"]}}
{"jsonrpc":"2.0","id":"p1","result":{"functions":["R.EVAL","R.CALL"]}}
```

GridCraft forgets the kept results of the named functions (all of the plugin's functions without
`names`; names that aren't the plugin's are ignored, at most 1,000) and recalculates the open
workbooks whose formulas call them. A plugin may send it any time: on the push channel (answered in
the main loop), while it runs a command, a macro or an event. The recalculation never happens while
the plugin is busy answering: after the command's step (outside its undo step), or in the app's
next turn. A workbook without unsaved changes stays so. The answer lists the names invalidated.
Typical use: an add-in's trust changed, a data source refreshed, a cache in the plugin expired.

## Requirements (manifest 1.3)

A plugin can say which other plugins it needs, in which versions, and which app versions and
protocol features. GridCraft checks this before starting it, says what's missing, and offers to
install it. The contract is shared with WordCraft and DeckCraft:
[Craft plugin manifests 1.3](https://github.com/sirmmo/craft-collab/blob/main/docs/plugin-requirements.md).

```json
{
  "id": "org.example.reports",
  "version": "1.4.0",
  "requires": {
    "apps": {"gridcraft": ">=0.7.0", "wordcraft": ">=0.7.0"},
    "protocol": ">=1.2",
    "features": ["dialogs", "events"],
    "addins": [
      {"id": "org.craft.vba", "version": "^0.1.4", "github": "sirmmo/craft-vba", "why": "Runs the report macros"},
      {"id": "org.craft.mqtt", "version": ">=0.1, <0.3", "github": "sirmmo/craft-collab", "optional": true, "why": "Live editing"}
    ]
  }
}
```

**Versions** are `MAJOR.MINOR.PATCH`; missing parts are 0, a pre-release (`1.0.0-beta.1`) sorts
before its release, `+build` is ignored, and an Office add-in's `1.0.0.0` compares on its first
three parts. **Ranges** are Cargo's: comma-separated comparators that must all hold. `^1.2.3` (also
a bare `1.2.3`) is `>=1.2.3, <2.0.0`, and for 0.x `^0.1.4` is `>=0.1.4, <0.2.0`; `~1.2.3` is
`>=1.2.3, <1.3.0`; `>=`, `>`, `<=`, `<`, `=` take a full or partial version (`>=0.1` is `>=0.1.0`;
as in Cargo, `=1.2` and `<=1.2` mean `1.2.*`, `>1.2` means `>=1.3.0`); `1.*` and `1.2.*` are
wildcards; `*` or nothing is any version, pre-releases included. Otherwise a pre-release only
matches a comparator that names a pre-release of the same version (`>=1.0.0-rc.1`). For the
protocol write `>=1.2`: Cargo's `^1.2` leaves out 2.0. Ranges are at most 128 characters.

**Blocked.** Whenever the plugins change (start-up, install, `plugins.load`, turning one on or
off, unloading, one failing for good), GridCraft resolves them all together. A plugin is
**blocked**, with its reasons in `blockedBy` and in Options › Add-ins, when:

| `kind` | When | Message |
|---|---|---|
| `app` | `requires.apps.gridcraft` doesn't match this version | Needs GridCraft 0.8 or later (this is 0.7.0) |
| `protocol` | `requires.protocol` doesn't match `host.protocol` | Needs plugin protocol 3.0 or later (this GridCraft has 2.0) |
| `feature` | A name in `requires.features` isn't in `host.features`; an Office add-in's requirement set or method isn't supported | Needs `x`, which this GridCraft doesn't have |
| `addin` | A required (not optional) plugin isn't installed, is turned off, failed, is blocked itself, or its version is out of range | Needs craft-vba ^0.1.4 (installed: 0.1.3) |
| `cycle` | Required plugins need each other | Requirement cycle: Reports → craft-vba → Reports |
| `depth` | The chain of required plugins below it is more than 8 deep | Its required add-ins go 9 deep (at most 8) |
| `range` | A range in its manifest doesn't parse | Its manifest's requirement for craft-vba (`soon`) isn't a version range |

A blocked plugin isn't started; its functions (cells show `#NAME?`), commands, ribbon groups and
menu items are absent, as for a plugin that's turned off. Installing what it needs unblocks it at
once, without restarting, and the workbooks using its functions recalculate. Optional plugins
never block. Required plugins start before the plugins that need them.

**Turning off or unloading** a plugin others need asks first in Options and Manage Add-ins
("Reports needs craft-vba. Turn both off?"). Programmatically it just happens: the plugins that
needed it are blocked, and the result's `dependentsBlocked` lists them. **Installing or updating**
a plugin to a version a turned-on plugin's range doesn't accept asks first ("Reports needs
craft-vba ^0.1.4; 0.2.0 doesn't match. Install craft-vba 0.2.0 anyway?"); programmatically
(`plugins.install`, `plugins.load`) it fails unless `force: true`, and then that plugin is blocked.

**Installing what's missing.** After installing a plugin from GitHub, GridCraft offers its
missing or out-of-range required plugins that name `github` ("Reports needs craft-vba ^0.1.4 (Runs
the report macros). Install it too?"), ticked; optional ones are offered unticked. For each it
takes **the newest release whose tag is a version in range** (`v0.1.5` or `0.1.5`; drafts are
skipped, and so are pre-releases unless the range names one), not just the latest, and says which.
What they need in turn is offered next. A blocked plugin whose missing plugins name `github` shows
**Install Missing** in Options › Add-ins and Manage Add-ins.

| Command | Does |
|---|---|
| `plugins.install {url, withRequirements?, force?, version?}` | `withRequirements` (default false here; the interface always asks) also installs the missing required plugins that name `github`, with `installed` and `failed` in the result. The result always lists `missing` (`[{id, name, version, github?, optional, why, installed?}]`) for the caller to install. `version`, a range, picks the newest release in it |
| `plugins.installMissing {id, force?}` | Installs `id`'s missing required plugins that name `github`, then theirs (at most 8 deep): `{installed: [{id, name, version, release}], failed: [{id, error}], missing, satisfied}` |
| `plugins.requirements {id}` | `{id, satisfied, blockedBy, tree: [{id, name, need, have, state, optional, why, github, children}]}`; `state` is `missing` for a plugin that isn't installed |
| `plugins.list` | Per plugin: `state` (`blocked` among `loaded`, `running`, `failed`, `disabled`), `blockedBy: [{id?, kind, need, have?, message}]`, `requires` (the manifest's), `requiredBy` (the ids of the plugins that need it) |

Office add-ins' `<Requirements>` are checked the same way; see [officejs.md](officejs.md).

## Publishing a plugin on GitHub

Put `plugin.json` at the root of the repository (or in one folder of it). A plugin written in
a scripting language (`"run": {"command": "python3", "args": ["main.py"]}`) needs nothing more:
GridCraft installs the source.

A compiled plugin attaches one `.zip` per platform to each GitHub release, holding `plugin.json`
and the program (`"run": {"command": "./my-plugin"}`; on Windows `./my-plugin.exe`). Name each zip
after its platform and architecture, separated by `-`, `.` or `+`:

| Platform | Words GridCraft recognises |
|---|---|
| macOS | `macos`, `darwin`, `mac`, `osx`, `apple`; `universal` for both architectures |
| Windows | `windows`, `win`, `win64`, `win32` |
| Linux | `linux` |
| FreeBSD | `freebsd` |
| Apple silicon, ARM | `aarch64`, `arm64` |
| Intel, AMD | `x86_64`, `x64`, `amd64` |

For example `finance-1.2.0-macos-aarch64.zip`, `finance-1.2.0-windows-x64.zip`. A zip naming no
platform is used on every platform that has no zip of its own. Executable permissions in the zip
are kept.

## Limits and next steps

- Calls are synchronous. A slow function holds up recalculation for up to its timeout.
  Asynchronous functions (`#GETTING_DATA` until the answer arrives) come with protocol 2.
- Data sources and streams pushed by the plugin (`STREAM()`, `RTD`), with targeted
  recalculation of the cells that use them, will build on protocol 2's push channel.
- Installed plugins aren't updated automatically; install again to update. There's no
  Uninstall button yet: unload the plugin and delete its folder.
- Ribbon groups go at the end of a tab; plugins can't place them between built-in groups, add
  drop-downs or toggles, or ship their own icons yet. Their labels aren't translated.
- The web app can't start processes, so plugins need the desktop app or the CLI.
- Workbooks don't record which plugin a formula needs. Without the plugin, its cells show
  `#NAME?` after a recalculation. In XLSX, the formula text is kept as written.

## Optional workbook Python project capabilities

The proposed `pythonProjects` and `engineCommands` features extend protocol 2.2 without
removing existing capabilities. They expose validated project get/set and a restricted engine
command bridge through the shared macro transaction. Macros may declare `recording:false`
and `ribbonCallbacks:false`; omitted flags preserve existing VBA behavior.

Native `developer.startMacro/pollMacro/cancelMacro` adds bounded, cancellable background
transport while Session work stays on the owner thread. Ordinary request activity timeouts
remain unchanged; asynchronous macro runs have an absolute budget, reject conflicting edits,
and roll back on Stop/error/timeout. Live collaboration is excluded from the async path.
Headless callers and the existing non-Python UI Run path use synchronous `developer.runMacro`; only Python opts into the new UI lifecycle.

The optional provider extends craft-python 0.3.0 instead of replacing its scripting dispatcher.
See [python-macro-integration.md](python-macro-integration.md) for installation, APIs, limits,
download consent and the separate provider identity.

## Revision history

| Date | Change | Summary |
|---|---|---|
| 2026-10-17 | minor | additive workbook Python capabilities and cancellable macro lifecycle |
| 2026-10-16 | minor | Protocol 2.2: `functions/call` carries `document {path, name, id}` (results kept per workbook), `id` in `context.document` and the document events, `documentOpened` before an opened workbook's first calculation (at most 10 s per plugin), multi-line and code form fields, `{"empty": true}` blank results (`null` stays 0, as in Excel), `range.setValues {literal: true}`, `host/functions/invalidate`; `host.features` `functionDocument`, `openedBeforeCalc`, `multiline`, `emptyResult`, `literalValues`, `invalidate` |
| 2026-10-15 | minor | Manifest 1.3: `requires` (apps, protocol, features, other add-ins with Cargo-style version ranges), blocked add-ins (`state`, `blockedBy`, `requiredBy`), start order, `host.protocol` and `addins` in `initialize`, `plugins.requirements`, `plugins.installMissing`, `plugins.install` `withRequirements`/`force`/`version` (newest release in range), Install Missing, asking before turning off a needed add-in or updating out of range; Office add-ins' requirement sets |
| 2026-10-15 | minor | Protocol 2: the push channel (plugins speak any time; `host.features` `push`, `collab`; `host.collab`), manifest `collab` (transports), `collab.join` in command results; co-editing in [collab.md](collab.md) |
| 2026-10-15 | minor | Protocol 1.2: `host.features` and `host.language`, `context.document`, command results that open a file (`open`) or show a host-drawn dialog (`dialog`: list, form, message; chained), manifest `events` and `events/documentOpened`, `documentSaved`, `documentClosed` |
| 2026-10-14 | minor | `context.sheets` and `context.names`: commands see every sheet and the defined names |
| 2026-10-14 | minor | `context.formulas`: commands see the selection's formula text, so plugins can tell inputs from calculations |
| 2026-10-14 | minor | Plugin functions are suggested while typing (signature hint, description) and listed in Insert Function and `formulas.functions` (category Add-ins) |
| 2026-10-14 | minor | Ribbon buttons that run macros, saved as the workbook's `customUI14.xml`; `macros/run` gets `control`; a workbook with macros saves as `.xlsm` |
| 2026-10-14 | minor | Protocol 1.1: `host/*` requests while a plugin works (one undo step), macro plugins (`macros/*`), the Developer tab and Macro Security |
| 2026-10-14 | minor | Office add-ins (`manifest.xml`) load as plugins whose commands open task panes ([officejs.md](officejs.md)) |
| 2026-10-14 | minor | One plugin for several apps: `apps`, `uiFor`, `host.id` in `initialize`, `CRAFT_HOST`/`CRAFT_PLUGIN_DIR`; the shared protocol WordCraft speaks too; `.craft-source.json` |
| 2026-10-13 | minor | Install from a GitHub repository (`plugins.install`, Options and Automate › Add-ins); wider add-in list |
| 2026-10-13 | minor | Options › Add-ins: turn plugins off and on (remembered), `homepage` and `author`, `plugins.setEnabled`, `plugins.openHomepage` |
| 2026-10-13 | minor | `ui` in the manifest: ribbon groups on built-in or new tabs, items in the cell, row, column and sheet-tab menus |
| 2026-10-13 | major | Created: out-of-process plugin host, manifest, protocol 1 (functions, commands), Add-ins dialog |
