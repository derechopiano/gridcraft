//! Plugin protocol 1.2 end to end: `host.features` and `host.language` in `initialize`,
//! `context.document`, command results that open a file or show a dialog (validated, bounded,
//! returned to programmatic callers), and document events (declared or not, after the fact, with
//! `host/sync` refused). The plugin is this test binary started again with `--serve-plugin` (no
//! libtest harness, so standard output is the protocol's alone).
#![allow(clippy::disallowed_methods)] // the clock, off wasm: timing the save and the close

use std::io::{BufRead, Write};
use std::time::{Duration, Instant};

use gridcraft_engine::Session;
use gridcraft_engine::core::{CellRef, Value};
use gridcraft_engine::plugin::events::{self, OutcomeKind};
use serde_json::{Value as Json, json};

const PLUGIN: &str = "org.gridcraft.test-protocol";
const QUIET: &str = "org.gridcraft.test-quiet";

fn main() {
    if std::env::args().any(|a| a == "--serve-plugin") {
        serve();
        return;
    }
    gridcraft_engine::plugin::set_host_language("it");
    let dir = setup();
    let tests: [(&str, fn(&std::path::Path)); 6] = [
        ("initialize_announces_features_and_language", initialize_announces_features_and_language),
        ("commands_see_the_document", commands_see_the_document),
        ("open_results_are_checked_and_returned", open_results_are_checked_and_returned),
        ("dialog_results_are_validated_bounded_and_returned", dialog_results_are_validated_bounded_and_returned),
        ("events_go_to_the_plugins_that_declared_them", events_go_to_the_plugins_that_declared_them),
        ("a_slow_event_never_holds_up_the_save_or_close", a_slow_event_never_holds_up_the_save_or_close),
    ];
    for (name, t) in tests {
        t(&dir);
        println!("test {name} ... ok");
    }
    let _ = std::fs::remove_dir_all(&dir);
}

// ---------------------------------------------------------------- the plugin

/// The plugin side. It remembers its `initialize` params and the events it got (`hello` reports
/// them), and answers events with a message; `documentOpened` tries `host/sync` (refused) and
/// `host/message`; `documentClosed` of a file named `slow.xlsx` takes 6 s (past its 5 s timeout).
fn serve() {
    let stdin = std::io::stdin();
    let mut lines = stdin.lock().lines();
    let mut out = std::io::stdout();
    let mut init = Json::Null;
    let mut seen: Vec<String> = Vec::new();
    let mut next = 0u32;
    while let Some(Ok(line)) = lines.next() {
        let Ok(msg) = serde_json::from_str::<Json>(&line) else { continue };
        let Some(id) = msg.get("id").cloned() else { break };
        let p = msg["params"].clone();
        let method = msg["method"].as_str().unwrap_or("").to_string();
        // Sends a host request and waits for its answer.
        let mut host = |method: &str, params: Json| -> Json {
            next += 1;
            let hid = format!("h{next}");
            writeln!(out, "{}", json!({"jsonrpc": "2.0", "id": hid, "method": method, "params": params})).unwrap();
            out.flush().unwrap();
            for line in lines.by_ref() {
                let msg: Json = serde_json::from_str(&line.unwrap()).unwrap();
                if msg["id"] == json!(hid) {
                    return msg;
                }
            }
            Json::Null
        };
        let answer = if method == "initialize" {
            init = p.clone();
            json!({"result": {"protocolVersion": 1}})
        } else if let Some(event) = method.strip_prefix("events/") {
            seen.push(format!("{event} {}", p["name"].as_str().unwrap_or("")));
            match event {
                "documentOpened" => {
                    let sync = host("host/sync", json!({"actions": []}));
                    host("host/message", json!({"text": format!("opened {}", p["name"].as_str().unwrap_or(""))}));
                    json!({"result": {"message": format!("sync: {}", sync["error"]["data"]["code"].as_str().unwrap_or("allowed"))}})
                }
                "documentSaved" => {
                    json!({"result": {"message": format!("saved {} as {}", p["name"].as_str().unwrap_or(""), p["format"].as_str().unwrap_or(""))}})
                }
                "documentClosed" if p["name"] == "slow.xlsx" => {
                    std::thread::sleep(Duration::from_secs(6));
                    json!({"result": {}})
                }
                "documentClosed" => json!({"error": {"code": 1, "message": "upload failed"}}),
                _ => json!({"result": {}}),
            }
        } else if method == "commands/run" {
            let q = &p["params"];
            let write = json!({"command": "cell.set", "params": {"cell": "A1", "input": "written"}});
            match p["command"].as_str().unwrap_or("") {
                "hello" => json!({"result": {"result": {"init": init, "events": seen, "document": p["context"]["document"]}}}),
                "openFile" => {
                    json!({"result": {"commands": [write], "open": {"path": q["path"], "label": q["label"], "readOnly": true}, "message": "Opened from the cloud"}})
                }
                "dialog" => json!({"result": {"commands": [write], "dialog": q["dialog"], "message": "Pick a file", "result": {"listed": 2}}}),
                other => json!({"error": {"code": -32000, "message": format!("unknown command {other}")}}),
            }
        } else {
            json!({"result": {}})
        };
        let mut reply = json!({"jsonrpc": "2.0", "id": id});
        if let (Some(r), Json::Object(a)) = (reply.as_object_mut(), answer) {
            r.extend(a);
        }
        if writeln!(out, "{reply}").and_then(|_| out.flush()).is_err() {
            break;
        }
    }
}

// ---------------------------------------------------------------- helpers

/// Loads the plugin twice: `PLUGIN` declares the three events, `QUIET` none.
fn setup() -> std::path::PathBuf {
    let dir = std::env::temp_dir().join(format!("gridcraft-engine-protocol-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let exe = std::env::current_exe().unwrap();
    let commands = json!([{"id": "hello", "label": "Hello"}, {"id": "openFile", "label": "Open"}, {"id": "dialog", "label": "Dialog"}]);
    for (id, events) in [(PLUGIN, json!(["documentOpened", "documentSaved", "documentClosed"])), (QUIET, json!([]))] {
        let pdir = dir.join(id);
        std::fs::create_dir_all(&pdir).unwrap();
        let manifest = json!({"id": id, "name": id, "run": {"command": exe.display().to_string(), "args": ["--serve-plugin"]}, "commands": commands, "events": events});
        std::fs::write(pdir.join("plugin.json"), manifest.to_string()).unwrap();
        gridcraft_engine::plugin::load(&pdir).unwrap();
    }
    dir
}

fn session() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    s
}

fn run(s: &mut Session, plugin: &str, command: &str, params: Json) -> Result<Json, String> {
    s.run("plugins.run", json!({"plugin": plugin, "command": command, "params": params}))
}

fn a1(s: &Session) -> Value {
    s.active().unwrap().wb.active().unwrap().value(CellRef::new(0, 0))
}

/// Every event sent so far has been delivered; what the plugins said.
fn settle() -> Vec<events::Outcome> {
    assert!(events::wait_idle(Duration::from_secs(20)), "the events were delivered");
    events::take_outcomes()
}

// ---------------------------------------------------------------- tests

fn initialize_announces_features_and_language(_dir: &std::path::Path) {
    let mut s = session();
    let r = run(&mut s, PLUGIN, "hello", json!({})).unwrap();
    let host = &r["result"]["init"]["host"];
    assert_eq!(host["id"], "gridcraft");
    assert_eq!(
        host["features"],
        json!([
            "dialogs",
            "open",
            "events",
            "document",
            "push",
            "collab",
            "functionDocument",
            "openedBeforeCalc",
            "multiline",
            "emptyResult",
            "literalValues",
            "invalidate",
            "pythonProjects",
            "engineCommands"
        ])
    );
    assert_eq!(host["protocol"], "2.2");
    assert_eq!(host["collab"], json!({"formats": ["yjs-v1"], "app": "gridcraft", "schema": 1, "encryption": ["aes-256-gcm"]}));
    assert_eq!(host["language"], "it");
    assert_eq!(r["result"]["init"]["protocolVersion"], 1, "the protocol version stays 1");
}

fn commands_see_the_document(dir: &std::path::Path) {
    let mut s = session();
    let doc = |s: &mut Session| run(s, QUIET, "hello", json!({})).unwrap()["result"]["document"].clone();
    let id = s.active().unwrap().session_id();
    assert!(id.starts_with("doc-"));
    assert_eq!(doc(&mut s), json!({"path": null, "name": "Book1", "dirty": false, "id": id}));
    s.run("cell.set", json!({"cell": "B2", "input": "1"})).unwrap();
    assert_eq!(doc(&mut s)["dirty"], true);
    let path = dir.join("Document.xlsx");
    s.run("file.save", json!({"path": path.display().to_string()})).unwrap();
    assert_eq!(
        doc(&mut s),
        json!({"path": path.display().to_string(), "name": "Document.xlsx", "dirty": false, "id": id}),
        "the id stays when it's saved"
    );
    settle();
}

fn open_results_are_checked_and_returned(dir: &std::path::Path) {
    let mut s = session();
    let file = dir.join("Budget.xlsx");
    std::fs::write(&file, gridcraft_engine::io::save_bytes(&gridcraft_engine::model::Workbook::new(), "x.xlsx").unwrap()).unwrap();
    let docs = s.documents().len();

    // Programmatic: nothing opens; `open` comes back for the caller, after the step ran.
    let r = run(&mut s, QUIET, "openFile", json!({"path": file.display().to_string(), "label": "Cloud · Budget.xlsx"})).unwrap();
    assert_eq!(r["open"], json!({"path": file.display().to_string(), "label": "Cloud · Budget.xlsx", "readOnly": true}));
    // Fields combine: the message comes with the file to open.
    assert_eq!(r["message"], "Opened from the cloud");
    assert!(s.take_ui_requests().contains(&gridcraft_engine::UiRequest::Message("Opened from the cloud".into())));
    assert_eq!(s.documents().len(), docs, "a programmatic run opens nothing");
    assert_eq!(a1(&s), Value::Text("written".into()));
    s.run("edit.undo", json!({})).unwrap();
    assert_eq!(a1(&s), Value::Empty, "the step is one undo");

    // Bad paths fail the command, and nothing changes.
    std::fs::write(dir.join("tool.exe"), b"MZ").unwrap();
    for (path, why) in [
        ("Budget.xlsx".to_string(), "absolute"),
        (dir.join("missing.xlsx").display().to_string(), "missing.xlsx"),
        (dir.display().to_string(), "doesn't open"),
        (dir.join("tool.exe").display().to_string(), "doesn't open"),
    ] {
        let e = run(&mut s, QUIET, "openFile", json!({"path": path})).unwrap_err();
        assert!(e.contains(why), "{path}: {e}");
        assert_eq!(a1(&s), Value::Empty, "{path}: nothing ran");
    }
    let dir_xlsx = dir.join("folder.xlsx");
    std::fs::create_dir_all(&dir_xlsx).unwrap();
    assert!(run(&mut s, QUIET, "openFile", json!({"path": dir_xlsx.display().to_string()})).unwrap_err().contains("isn't a file"));
    // A malformed `open` is an error, not a crash.
    assert!(run(&mut s, QUIET, "openFile", json!({"path": 42})).is_err());
}

fn dialog_results_are_validated_bounded_and_returned(_dir: &std::path::Path) {
    let mut s = session();
    let items: Vec<Json> = (0..10_050).map(|i| json!({"id": format!("/D/{i}"), "label": "x".repeat(300), "icon": "folder"})).collect();
    let dialog = json!({
        "title": "T".repeat(250), "kind": "list", "items": items, "command": "hello", "params": {"account": "me"},
        "okLabel": "Open", "buttons": (0..6).map(|_| json!({"label": "Up", "command": "dialog"})).collect::<Vec<_>>(),
    });
    let r = run(&mut s, QUIET, "dialog", json!({"dialog": dialog})).unwrap();
    let d = &r["dialog"];
    assert_eq!(d["kind"], "list");
    assert_eq!(d["title"].as_str().unwrap().chars().count(), 200);
    assert_eq!(d["items"].as_array().unwrap().len(), 10_000);
    assert_eq!(d["items"][0]["label"].as_str().unwrap().chars().count(), 256);
    assert_eq!(d["buttons"].as_array().unwrap().len(), 4);
    assert_eq!((d["command"].clone(), d["okLabel"].clone(), d["params"].clone()), (json!("hello"), json!("Open"), json!({"account": "me"})));
    assert_eq!(a1(&s), Value::Text("written".into()), "the step ran");
    // Fields combine: the dialog comes with the result and the message.
    assert_eq!((r["result"].clone(), r["message"].clone()), (json!({"listed": 2}), json!("Pick a file")));

    // A message needs no command.
    let r = run(&mut s, QUIET, "dialog", json!({"dialog": {"kind": "message", "title": "Done", "prompt": "Uploaded"}})).unwrap();
    assert_eq!(r["dialog"]["kind"], "message");

    // Malformed dialogs, and dialogs running commands the plugin doesn't have, fail the command.
    s.run("cell.set", json!({"cell": "A1", "input": ""})).unwrap();
    for bad in [
        json!({"kind": "web", "command": "hello"}),
        json!({"kind": "list", "items": []}),
        json!({"kind": "list", "command": "hello", "items": {"a": 1}}),
        json!({"kind": "list", "command": "nope", "items": []}),
        json!({"kind": "message", "buttons": [{"label": "Go", "command": "plugins.run"}]}),
        json!({"kind": "form", "command": "hello", "fields": [{"id": "x", "kind": "html"}]}),
        json!(17),
    ] {
        let e = run(&mut s, QUIET, "dialog", json!({"dialog": bad})).unwrap_err();
        assert!(!e.is_empty(), "{bad}");
        assert_eq!(a1(&s), Value::Empty, "{bad}: nothing ran");
    }
}

fn events_go_to_the_plugins_that_declared_them(dir: &std::path::Path) {
    events::take_outcomes();
    let mut s = session();
    let path = dir.join("Events.xlsx");
    s.run("cell.set", json!({"cell": "A1", "input": "1"})).unwrap();
    s.run("file.save", json!({"path": path.display().to_string()})).unwrap();
    let out = settle();
    assert!(out.iter().any(|o| o.plugin == PLUGIN && o.kind == OutcomeKind::Message && o.text == "saved Events.xlsx as xlsx"), "{out:?}");
    assert!(out.iter().all(|o| o.plugin != QUIET), "{out:?}");

    // An automatic save isn't an event.
    s.run("cell.set", json!({"cell": "A1", "input": "2"})).unwrap();
    s.run("file.save", json!({"autosave": true})).unwrap();
    assert!(settle().is_empty());

    // Closing a document with a file: the error answer is reported, the close happened.
    s.run("file.close", json!({"force": true})).unwrap();
    let out = settle();
    assert!(out.iter().any(|o| o.plugin == PLUGIN && o.kind == OutcomeKind::Error && o.text.contains("upload failed")), "{out:?}");

    // Opening: the plugin hears of it; `host/sync` is refused during the event, `host/message` works.
    s.run("file.open", json!({"path": path.display().to_string()})).unwrap();
    let out = settle();
    assert!(out.iter().any(|o| o.kind == OutcomeKind::Message && o.text == "sync: AccessDenied"), "{out:?}");
    assert!(out.iter().any(|o| o.kind == OutcomeKind::HostMessage && o.text == "opened Events.xlsx"), "{out:?}");
    assert_eq!(s.active().unwrap().undo.len(), 0, "an event is never an undo step");

    let seen = run(&mut s, PLUGIN, "hello", json!({})).unwrap()["result"]["events"].clone();
    let names: Vec<&str> = seen.as_array().unwrap().iter().filter_map(Json::as_str).filter(|e| e.ends_with("Events.xlsx")).collect();
    assert_eq!(names, ["documentSaved Events.xlsx", "documentClosed Events.xlsx", "documentOpened Events.xlsx"]);
    assert_eq!(run(&mut s, QUIET, "hello", json!({})).unwrap()["result"]["events"], json!([]), "undeclared: nothing sent");
}

fn a_slow_event_never_holds_up_the_save_or_close(dir: &std::path::Path) {
    events::take_outcomes();
    let mut s = session();
    let path = dir.join("slow.xlsx");
    s.run("file.save", json!({"path": path.display().to_string()})).unwrap();
    settle();
    let t = Instant::now();
    s.run("file.close", json!({})).unwrap();
    assert!(t.elapsed() < Duration::from_secs(2), "the close didn't wait for the plugin");
    assert!(s.documents().is_empty());
    // The plugin takes 6 s over a 5 s event: it times out (reported), and the next save is fine.
    let out = settle();
    assert!(out.iter().any(|o| o.plugin == PLUGIN && o.kind == OutcomeKind::Error && o.text.contains("took longer")), "{out:?}");
    let mut s = session();
    s.run("file.save", json!({"path": dir.join("after.xlsx").display().to_string()})).unwrap();
    assert!(dir.join("after.xlsx").is_file());
    let out = settle();
    assert!(out.iter().any(|o| o.text == "saved after.xlsx as xlsx"), "the plugin was started again: {out:?}");
    // Quitting: `documentClosed` for the open workbooks with a file, waiting at most 5 s.
    let t = Instant::now();
    gridcraft_engine::cmd::plugins::quitting(&s);
    assert!(t.elapsed() < Duration::from_secs(6));
    let out = settle();
    assert!(out.iter().any(|o| o.text.contains("upload failed")), "{out:?}");
}
