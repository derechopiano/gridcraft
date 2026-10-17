//! Macros through a macro add-in (protocol 1.1): the add-in reads and writes the workbook while a
//! macro runs (`host/sync`), keeps its code in the workbook's `vbaProject.bin`, and a run is one
//! undo step. The add-in is this test binary started again with `--serve-macros` (no libtest
//! harness, so standard output is the protocol's alone).

use std::io::{BufRead, Write};

use gridcraft_engine::Session;
use gridcraft_engine::core::{CellRef, Value};
use serde_json::{Value as Json, json};

const PLUGIN: &str = "org.gridcraft.test-macros";

fn main() {
    if std::env::args().any(|a| a == "--serve-macros") {
        serve();
        return;
    }
    let tests: [(&str, fn()); 9] = [
        ("macros_run_as_one_undo_step_and_call_the_host", macros_run_as_one_undo_step_and_call_the_host),
        ("a_failing_macro_changes_nothing", a_failing_macro_changes_nothing),
        ("macro_security_gates_running", macro_security_gates_running),
        ("modules_live_in_the_workbooks_vba_project", modules_live_in_the_workbooks_vba_project),
        ("ribbon_buttons_run_their_macro_and_are_saved", ribbon_buttons_run_their_macro_and_are_saved),
        ("async_success_and_noop_history", async_success_and_noop_history),
        ("async_error_stop_and_timeout_rollback", async_error_stop_and_timeout_rollback),
        ("async_workbook_change_cancels", async_workbook_change_cancels),
        ("python_download_approval_is_session_and_metadata_bound", python_download_approval_is_session_and_metadata_bound),
    ];
    for (name, t) in tests {
        t();
        println!("test {name} ... ok");
    }
}

/// The add-in: `Fill` writes B1 = A1 * 2 (reading A1 through the host) and says so; `Broken`
/// writes then fails; modules are the vbaProject's bytes taken as UTF-8 text.
fn serve() {
    let stdin = std::io::stdin();
    let mut lines = stdin.lock().lines();
    let mut out = std::io::stdout();
    let mut next = 0u32;
    // Sends a host request and waits for its answer.
    let mut host = |out: &mut std::io::Stdout, lines: &mut std::io::Lines<std::io::StdinLock<'_>>, method: &str, params: Json| -> Json {
        next += 1;
        let id = format!("h{next}");
        writeln!(out, "{}", json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params})).unwrap();
        out.flush().unwrap();
        for line in lines.by_ref() {
            let msg: Json = serde_json::from_str(&line.unwrap()).unwrap();
            if msg["id"] == json!(id) {
                return msg;
            }
        }
        Json::Null
    };
    let a1 = json!([{"p": "workbook"}, {"p": "worksheets"}, {"m": "getActiveWorksheet"}, {"m": "getRange", "a": ["A1"]}]);
    let b1 = json!([{"p": "workbook"}, {"p": "worksheets"}, {"m": "getActiveWorksheet"}, {"m": "getRange", "a": ["B1"]}]);
    while let Some(Ok(line)) = lines.next() {
        let Ok(msg) = serde_json::from_str::<Json>(&line) else { continue };
        let Some(id) = msg.get("id").cloned() else { break };
        let p = msg["params"].clone();
        let answer = match msg["method"].as_str() {
            Some("macros/list") => {
                json!({"result": {"macros": [{"name": "Module1.Fill", "module": "Module1", "procedure": "Fill"}, {"name": "Module1.Broken", "module": "Module1", "procedure": "Broken"}]}})
            }
            Some("macros/run") if p["name"] == "Module1.Fill" || p["name"] == "Fill" => {
                let r = host(&mut out, &mut lines, "host/sync", json!({"actions": [{"op": "load", "path": a1, "props": ["values"]}]}));
                let x = r["result"]["results"][0]["values"][0][0].as_f64().unwrap_or(0.0);
                host(&mut out, &mut lines, "host/sync", json!({"actions": [{"op": "set", "path": b1, "prop": "values", "value": x * 2.0}]}));
                host(&mut out, &mut lines, "host/message", json!({"text": format!("doubled {x}")}));
                json!({"result": {}})
            }
            // A ribbon button's callback: writes which button called it.
            Some("macros/run") if p["name"] == "Module1.Pressed" => {
                let text = format!("{}/{}", p["control"]["id"].as_str().unwrap_or("none"), p["control"]["tag"].as_str().unwrap_or(""));
                host(&mut out, &mut lines, "host/sync", json!({"actions": [{"op": "set", "path": b1, "prop": "values", "value": text}]}));
                json!({"result": {}})
            }
            Some("macros/run") if p["name"] == "Module1.Wait" => {
                host(&mut out, &mut lines, "host/sync", json!({"actions": [{"op": "set", "path": b1, "prop": "values", "value": 99}]}));
                std::thread::sleep(std::time::Duration::from_secs(60));
                json!({"result": {}})
            }
            Some("macros/run") if p["name"] == "Module1.Broken" => {
                host(&mut out, &mut lines, "host/sync", json!({"actions": [{"op": "set", "path": b1, "prop": "values", "value": 99}]}));
                let bad = host(
                    &mut out,
                    &mut lines,
                    "host/sync",
                    json!({"actions": [{"op": "load", "path": [{"p": "workbook"}, {"p": "worksheets"}, {"m": "getItem", "a": ["Nope"]}], "props": ["name"]}]}),
                );
                json!({"error": {"code": 1, "message": format!("Run-time error '9': Subscript out of range ({})", bad["error"]["data"]["code"].as_str().unwrap_or(""))}})
            }
            Some("macros/modules") => {
                let r = host(&mut out, &mut lines, "host/vbaProject/get", json!({}));
                let text = r["result"]["base64"]
                    .as_str()
                    .map(|b| String::from_utf8(gridcraft_engine::io::base64_decode(b).unwrap()).unwrap())
                    .unwrap_or_default();
                json!({"result": {"modules": if text.is_empty() { json!([]) } else { json!([{"name": "Module1", "kind": "standard", "source": text}]) }}})
            }
            Some("macros/saveModule") => {
                let b64 = gridcraft_engine::io::base64_encode(p["source"].as_str().unwrap_or("").as_bytes());
                host(&mut out, &mut lines, "host/vbaProject/set", json!({"base64": b64}));
                json!({"result": {}})
            }
            Some("macros/record") => {
                json!({"result": {"source": format!("Sub Macro1()\r\n' {} steps\r\nEnd Sub", p["steps"].as_array().map(Vec::len).unwrap_or(0))}})
            }
            _ => json!({"result": {}}),
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

fn session() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    let dir = std::env::temp_dir().join(format!("gridcraft-engine-macros-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let exe = std::env::current_exe().unwrap();
    let manifest =
        json!({"id": PLUGIN, "name": "Test macros", "macros": {"language": "TestBasic"}, "run": {"command": exe, "args": ["--serve-macros"]}});
    std::fs::write(dir.join("plugin.json"), manifest.to_string()).unwrap();
    s.execute("plugins.load", json!({"path": dir})).unwrap();
    s
}

fn value(s: &Session, a: &str) -> Value {
    s.active().unwrap().wb.active().unwrap().value(CellRef::parse(a).unwrap())
}

fn macros_run_as_one_undo_step_and_call_the_host() {
    let mut s = session();
    s.execute("cell.set", json!({"cell": "A1", "input": "21"})).unwrap();
    let list = s.execute("developer.macros", json!({})).unwrap();
    assert_eq!(list["macros"][0]["name"], "Module1.Fill");
    assert_eq!(list["macros"][0]["language"], "TestBasic");
    assert_eq!(list["enabled"], false, "a workbook's macros are off until enabled");
    s.execute("developer.enableMacros", json!({})).unwrap();
    let undo = s.active().unwrap().undo.len();
    s.execute("developer.runMacro", json!({"name": "Module1.Fill"})).unwrap();
    assert_eq!(value(&s, "B1"), Value::Number(42.0));
    assert_eq!(s.active().unwrap().undo.len(), undo + 1, "one undo step for the whole macro");
    assert!(s.take_ui_requests().iter().any(|q| matches!(q, gridcraft_engine::UiRequest::Message(m) if m == "doubled 21")));
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(value(&s, "B1"), Value::Empty);
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
}

fn a_failing_macro_changes_nothing() {
    let mut s = session();
    s.execute("developer.enableMacros", json!({})).unwrap();
    let undo = s.active().unwrap().undo.len();
    let e = s.execute("developer.runMacro", json!({"name": "Module1.Broken"})).unwrap_err().to_string();
    assert!(e.contains("Subscript out of range") && e.contains("ItemNotFound"), "{e}");
    assert_eq!(value(&s, "B1"), Value::Empty, "the write before the failure was taken back");
    assert_eq!(s.active().unwrap().undo.len(), undo);
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
}

fn macro_security_gates_running() {
    let mut s = session();
    assert!(s.execute("developer.runMacro", json!({"name": "Fill"})).unwrap_err().to_string().contains("disabled"));
    s.execute("developer.macroSecurity", json!({"level": "enable"})).unwrap();
    assert!(s.execute("developer.runMacro", json!({"name": "Fill"})).is_ok());
    s.execute("developer.macroSecurity", json!({"level": "disable"})).unwrap();
    s.execute("developer.enableMacros", json!({})).unwrap();
    assert!(s.execute("developer.runMacro", json!({"name": "Fill"})).is_err(), "Disable wins over a workbook's Enable Content");
    assert!(s.execute("developer.macroSecurity", json!({"level": "maybe"})).is_err());
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
    assert!(s.execute("developer.macros", json!({})).unwrap()["macros"].as_array().unwrap().is_empty());
    assert!(s.execute("developer.runMacro", json!({"name": "Fill"})).is_err());
}

fn modules_live_in_the_workbooks_vba_project() {
    let mut s = session();
    assert_eq!(s.execute("developer.modules", json!({})).unwrap()["modules"], json!([]));
    s.execute("developer.saveModule", json!({"name": "Module1", "source": "Sub Hello()\r\nEnd Sub"})).unwrap();
    assert_eq!(s.active().unwrap().wb.vba_project.as_deref(), Some(&b"Sub Hello()\r\nEnd Sub"[..]));
    assert_eq!(s.execute("developer.modules", json!({})).unwrap()["modules"][0]["source"], "Sub Hello()\r\nEnd Sub");
    // Saving a module is undoable, like any change to the workbook.
    s.execute("edit.undo", json!({})).unwrap();
    assert!(s.active().unwrap().wb.vba_project.is_none());
    s.execute("edit.redo", json!({})).unwrap();
    // A workbook with macros saves as .xlsm, never as an .xlsx that would lose them.
    let dir = std::env::temp_dir().join(format!("gridcraft-engine-macros-save-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let asked = dir.join("Budget.xlsx");
    let r = s.execute("file.save", json!({"path": asked})).unwrap();
    let saved = dir.join("Budget.xlsm");
    assert_eq!(r["path"], json!(saved.to_string_lossy()), "{r}");
    assert_eq!(r["renamedFrom"], json!(asked.to_string_lossy()));
    assert!(saved.exists() && !asked.exists());
    assert!(s.take_ui_requests().iter().any(|q| matches!(q, gridcraft_engine::UiRequest::Message(m) if m.contains("Budget.xlsm"))));
    assert_eq!(s.active().unwrap().path.as_deref(), Some(saved.to_string_lossy().as_ref()), "the document is the .xlsm now");
    let (back, _) = gridcraft_xlsx::read_xlsx(&std::fs::read(&saved).unwrap()).unwrap();
    assert!(back.vba_project.is_some());
    let _ = std::fs::remove_dir_all(&dir);
    let code = s.execute("developer.recordToCode", json!({"steps": [{"command": "cell.set", "params": {"cell": "A1", "input": "5"}}]})).unwrap();
    assert!(code["source"].as_str().unwrap().contains("1 steps"));
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
}

fn ribbon_buttons_run_their_macro_and_are_saved() {
    // A Python provider loaded first must not take over Office-style ribbon callbacks.
    let other = "org.gridcraft.test-python-first";
    let manifest: gridcraft_engine::plugin::Manifest = serde_json::from_value(json!({
        "id":other,"name":"Python first","macros":{"language":"Python","recording":false,"ribbonCallbacks":false},
        "run":{"command":"deliberately-missing-python-provider"}
    }))
    .unwrap();
    gridcraft_engine::plugin::install(manifest, std::env::temp_dir()).unwrap();
    let mut s = session();
    assert!(s.execute("developer.recordMacro", json!({"plugin":other})).unwrap_err().to_string().contains("does not support recording"));
    assert!(s.recording.is_none(), "an unsupported recorder must not start recording");
    assert_eq!(s.execute("developer.ribbon", json!({})).unwrap(), json!({"tabs": [], "editable": true}));
    let id = s.execute("developer.addRibbonButton", json!({"macro": "Module1.Pressed", "label": "Go", "tab": "data", "tag": "north"})).unwrap()["id"]
        .as_str()
        .unwrap()
        .to_string();
    s.execute("developer.addRibbonButton", json!({"macro": "Module1.Fill", "large": false})).unwrap();
    let r = s.execute("developer.ribbon", json!({})).unwrap();
    assert_eq!((r["tabs"][0]["tab"].as_str(), r["tabs"][0]["builtin"].as_str()), (Some("Data"), Some("TabData")), "{r}");
    assert_eq!(r["tabs"][0]["groups"][0]["buttons"][0]["label"], "Go");
    assert_eq!(r["tabs"][1]["tab"], "Macros", "a tab of its own by default");
    assert_eq!(r["tabs"][1]["groups"][0]["buttons"][0]["label"], "Fill", "labelled after the macro");

    // Pressing it runs the macro with the button, under Macro Security like any macro.
    assert!(s.execute("developer.pressRibbonButton", json!({"id": id})).unwrap_err().to_string().contains("disabled"));
    s.execute("developer.enableMacros", json!({})).unwrap();
    s.execute("developer.pressRibbonButton", json!({"id": id})).unwrap();
    assert_eq!(value(&s, "B1"), Value::Text(format!("{id}/north").into()));
    assert!(s.execute("developer.pressRibbonButton", json!({"id": "nope"})).is_err());

    // Saved with the workbook as Excel's ribbon part, and read back.
    let wb = s.active().unwrap().wb.clone();
    let bytes = gridcraft_xlsx::write_xlsx(&wb).unwrap();
    let (back, _) = gridcraft_xlsx::read_xlsx(&bytes).unwrap();
    assert_eq!(back.ribbon, wb.ribbon);

    // Edits and removals are undoable; the last button gone, the workbook has no ribbon.
    s.execute("developer.editRibbonButton", json!({"id": id, "label": "Run it", "image": "HappyFace"})).unwrap();
    assert_eq!(s.active().unwrap().wb.ribbon.as_ref().unwrap().button(&id).unwrap().label, "Run it");
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(s.active().unwrap().wb.ribbon.as_ref().unwrap().button(&id).unwrap().label, "Go");
    assert!(s.execute("developer.editRibbonButton", json!({"id": id, "image": "<img>"})).is_err());
    assert!(s.execute("developer.addRibbonButton", json!({"macro": "x\"/><evil"})).is_err());
    let all: Vec<String> =
        s.active().unwrap().wb.ribbon.as_ref().unwrap().ids().map(str::to_string).filter(|i| i.starts_with("craftButton")).collect();
    for b in &all {
        s.execute("developer.removeRibbonButton", json!({"id": b})).unwrap();
    }
    assert!(s.active().unwrap().wb.ribbon.is_none());
    s.execute("edit.undo", json!({})).unwrap();
    assert!(s.active().unwrap().wb.ribbon.as_ref().is_some_and(|r| r.button(&all[all.len() - 1]).is_some()));
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
    assert!(gridcraft_engine::plugin::unload(other));
}

#[allow(clippy::disallowed_methods)] // Native process regressions only.
fn pump_until(s: &mut Session, written: bool) -> gridcraft_engine::Result<Json> {
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(5);
    loop {
        let result = s.execute("developer.pollMacro", json!({}))?;
        if result["running"] == false || (written && value(s, "B1") == Value::Number(99.0)) {
            return Ok(result);
        }
        assert!(std::time::Instant::now() < deadline, "macro did not make progress");
        std::thread::sleep(std::time::Duration::from_millis(1));
    }
}

fn async_ready() -> Session {
    let mut s = session();
    s.execute("cell.set", json!({"cell":"A1","input":"21"})).unwrap();
    s.execute("cell.set", json!({"cell":"C1","input":"17"})).unwrap();
    s.execute("edit.undo", json!({})).unwrap();
    s.execute("developer.enableMacros", json!({})).unwrap();
    s.journal = (0..10_000).map(|n| ("fixture".into(), json!({"n":n}))).collect();
    s
}

fn async_success_and_noop_history() {
    let mut s = async_ready();
    assert!(s.execute_headless("developer.startMacro", json!({"name":"Module1.Fill"})).unwrap_err().to_string().contains("developer.runMacro"));
    assert!(!s.macro_running());
    let before = s.active().unwrap().clone();
    let mut expected = s.journal.clone();
    s.execute("developer.startMacro", json!({"name":"Module1.Fill"})).unwrap();
    for id in ["cell.set", "file.save", "edit.undo", "plugins.unload", "developer.startMacro"] {
        assert!(s.execute(id, json!({})).unwrap_err().to_string().contains("macro"), "{id}");
    }
    assert_eq!(pump_until(&mut s, false).unwrap()["ran"], "Module1.Fill");
    assert_eq!(value(&s, "B1"), Value::Number(42.0));
    assert_eq!(s.active().unwrap().undo.len(), before.undo.len() + 1);
    assert_eq!(s.active().unwrap().revision, before.revision + 1);
    expected.push(("developer.runMacro".into(), json!({"name":"Module1.Fill","plugin":PLUGIN})));
    expected.drain(..5_000);
    assert_eq!(s.journal, expected, "only the final macro step may trim the bounded journal");
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(value(&s, "B1"), Value::Empty);
    let before = s.active().unwrap().clone();
    s.execute("developer.startMacro", json!({"name":"Module1.Nothing"})).unwrap();
    pump_until(&mut s, false).unwrap();
    assert!(std::sync::Arc::ptr_eq(&s.active().unwrap().wb, &before.wb));
    assert_eq!(s.active().unwrap().undo.len(), before.undo.len());
    assert_eq!(s.active().unwrap().redo.len(), before.redo.len());
    assert_eq!(s.active().unwrap().revision, before.revision);
    gridcraft_engine::plugin::unload(PLUGIN);
}

fn async_error_stop_and_timeout_rollback() {
    let mut s = async_ready();
    for mode in ["failure", "stop", "timeout"] {
        let before = s.active().unwrap().clone();
        let journal = s.journal.clone();
        s.execute(
            "developer.startMacro",
            json!({"name":if mode=="failure" {"Module1.Broken"} else {"Module1.Wait"},
            "timeoutMs":if mode=="timeout" {300} else {5000}}),
        )
        .unwrap();
        if mode == "stop" {
            pump_until(&mut s, true).unwrap();
            s.execute("developer.cancelMacro", json!({})).unwrap();
            assert!(s.macro_running(), "cleanup precedes re-enabling edits");
        }
        let error = pump_until(&mut s, false).unwrap_err().to_string();
        let expected = match mode {
            "failure" => "Subscript",
            "stop" => "cancelled",
            _ => "absolute time limit",
        };
        assert!(error.contains(expected), "{error}");
        assert!(std::sync::Arc::ptr_eq(&s.active().unwrap().wb, &before.wb));
        assert_eq!(s.active().unwrap().revision, before.revision);
        assert_eq!(s.active().unwrap().undo.len(), before.undo.len());
        assert_eq!(s.active().unwrap().redo.len(), before.redo.len());
        assert_eq!(s.active().unwrap().selection, before.selection);
        assert_eq!(s.journal, journal);
    }
    gridcraft_engine::plugin::unload(PLUGIN);
}

fn async_workbook_change_cancels() {
    let mut s = async_ready();
    let original = s.active_index();
    let before = s.active().unwrap().wb.clone();
    s.execute("developer.startMacro", json!({"name":"Module1.Wait"})).unwrap();
    pump_until(&mut s, true).unwrap();
    s.new_workbook();
    assert!(!s.macro_running());
    assert!(std::sync::Arc::ptr_eq(&s.documents()[original].wb, &before));
    assert_eq!(value(&s, "B1"), Value::Empty);
    gridcraft_engine::plugin::unload(PLUGIN);
}

fn python_download_approval_is_session_and_metadata_bound() {
    let mut s = Session::new();
    s.new_workbook();
    let metadata = s.execute("developer.pythonSettings", json!({})).unwrap()["metadata"].as_str().unwrap().to_string();
    s.execute("developer.setPythonSettings", json!({"metadata":metadata,"allowDownloads":true})).unwrap();
    assert_eq!(s.execute("developer.pythonSettings", json!({})).unwrap()["allowDownloads"], true);
    let wb = s.active().unwrap().wb.clone();
    let encoded = serde_json::to_string(&wb.python_macros).unwrap();
    assert!(!encoded.contains("allowDownloads"), "a workbook cannot carry download consent");
    let changed = metadata.replace(">=3.11", ">=3.12");
    gridcraft_engine::cmd::developer::host_call(&mut s, "fixture", "host/pythonProject/set", &json!({"project":{"metadata":changed,"modules":{}}}))
        .unwrap();
    assert_eq!(s.execute("developer.pythonSettings", json!({})).unwrap()["allowDownloads"], false);
    s.new_workbook();
    assert_eq!(s.execute("developer.pythonSettings", json!({})).unwrap()["allowDownloads"], false);
    let before = s.active().unwrap().wb.clone();
    assert!(s.execute("developer.setPythonSettings", json!({"metadata":"print('untrusted')","allowDownloads":true})).is_err());
    assert!(std::sync::Arc::ptr_eq(&before, &s.active().unwrap().wb));
}
