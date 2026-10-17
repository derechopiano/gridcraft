//! Real-process UI checks for Run, Stop and Escape; no external interpreter needed.
#![allow(clippy::disallowed_methods)] // Native-only process/UI regression harness.
use egui_kittest::kittest::Queryable as _;
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_ui_egui::SheetApp;
use serde_json::{Value as Json, json};
use std::io::{BufRead, Write};
use std::time::{Duration, Instant};
const PLUGIN: &str = "test.ui-macro-stop";

fn main() {
    if std::env::args().any(|a| a == "--serve") {
        serve();
        return;
    }
    for escape in [false, true] {
        stop_rolls_back(escape);
    }
    existing_vba_provider_keeps_its_synchronous_run();
    println!("Async macro UI: Run, enabled Stop, Escape, edit blocking and rollback passed");
}

fn stop_rolls_back(escape: bool) {
    let manifest = gridcraft_engine::plugin::Manifest::parse(
        &json!({
            "id":PLUGIN,"name":"UI macro fixture","macros":{"language":"Python"},
            "run":{"command":std::env::current_exe().unwrap(),"args":["--serve"]}
        })
        .to_string(),
    )
    .unwrap();
    gridcraft_engine::plugin::install(manifest, std::env::temp_dir()).unwrap();
    let mut s = Session::new();
    s.new_workbook();
    s.execute("developer.enableMacros", json!({})).unwrap();
    let mut app = SheetApp::new(s, Default::default());
    app.ui.dark = escape; // Exercise light and dark themes.
    app.open_dialog("macrosVba", json!({}));
    let screenshots = std::env::var_os("GRIDCRAFT_TEST_MACRO_SCREENSHOTS").map(std::path::PathBuf::from);
    let mut builder = egui_kittest::Harness::builder().with_size(egui::vec2(1200.0, 800.0));
    if screenshots.is_some() {
        builder = builder.wgpu();
    }
    let mut h = builder.build_ui_state(
        |ui, app: &mut SheetApp| {
            let ctx = ui.ctx().clone();
            app.logic(&ctx);
            app.ui(ui);
        },
        app,
    );
    h.run_steps(4);
    let before = h.state().session.active().unwrap().clone();
    h.get_by_label("Run").click();
    h.run_steps(1);
    assert!(h.state().session.macro_running());
    let deadline = Instant::now() + Duration::from_secs(5);
    while cell(&h.state().session) != Value::Number(99.0) {
        assert!(Instant::now() < deadline, "macro did not write");
        h.run_steps(1);
        std::thread::sleep(Duration::from_millis(1));
    }
    if let Some(folder) = &screenshots {
        std::fs::create_dir_all(folder).unwrap();
        h.render().unwrap().save(folder.join(if escape { "macro-progress-dark.png" } else { "macro-progress-light.png" })).unwrap();
    }
    h.input_mut().events.push(egui::Event::Text("Oops".into()));
    h.run_steps(1);
    assert!(h.state().editor.is_none(), "typing must not start an editor under the modal");
    if escape {
        h.key_press(egui::Key::Escape);
    } else {
        h.get_by_label("Stop").click();
    }
    let deadline = Instant::now() + Duration::from_secs(5);
    while h.state().session.macro_running() {
        assert!(Instant::now() < deadline, "Stop did not finish");
        h.run_steps(1);
        std::thread::sleep(Duration::from_millis(1));
    }
    let d = h.state().session.active().unwrap();
    assert!(std::sync::Arc::ptr_eq(&d.wb, &before.wb));
    assert_eq!(d.revision, before.revision);
    assert_eq!(d.undo.len(), before.undo.len());
    assert_eq!(d.redo.len(), before.redo.len());
    assert_eq!(cell(&h.state().session), Value::Empty);
    h.state_mut().open_dialog("visualBasic", json!({}));
    h.run_steps(2);
    h.get_by_label("Python Settings").click();
    h.run_steps(2);
    if let Some(folder) = screenshots {
        h.render().unwrap().save(folder.join(if escape { "python-settings-dark.png" } else { "python-settings-light.png" })).unwrap();
    }
    h.get_by_label("Allow uv to download packages and Python for this metadata in this session").click();
    h.run_steps(2);
    // The settings window is drawn after the editor; its Save is the final matching button.
    h.get_all_by_label("Save").last().unwrap().click();
    h.run_steps(2);
    assert_eq!(h.state_mut().session.execute("developer.pythonSettings", json!({})).unwrap()["allowDownloads"], true);
    drop(h);
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
}

fn cell(s: &Session) -> Value {
    s.active().unwrap().wb.active().unwrap().value(CellRef::parse("B1").unwrap())
}

fn existing_vba_provider_keeps_its_synchronous_run() {
    let manifest = gridcraft_engine::plugin::Manifest::parse(
        &json!({"id":PLUGIN,"name":"Existing VBA fixture","macros":{"language":"VBA"},
            "run":{"command":std::env::current_exe().unwrap(),"args":["--serve","--vba"]}})
        .to_string(),
    )
    .unwrap();
    gridcraft_engine::plugin::install(manifest, std::env::temp_dir()).unwrap();
    let mut s = Session::new();
    s.new_workbook();
    s.execute("developer.enableMacros", json!({})).unwrap();
    let mut app = SheetApp::new(s, Default::default());
    app.open_dialog("macrosVba", json!({}));
    let mut h = egui_kittest::Harness::builder().with_size(egui::vec2(1200.0, 800.0)).build_ui_state(
        |ui, app: &mut SheetApp| {
            app.logic(&ui.ctx().clone());
            app.ui(ui);
        },
        app,
    );
    h.run_steps(4);
    h.get_by_label("Run").click();
    h.run_steps(1);
    assert_eq!(cell(&h.state().session), Value::Number(99.0));
    assert!(!h.state().session.macro_running(), "VBA Run must not opt into Python's absolute deadline");
    assert!(h.state().dev.running_macro.is_none());
    drop(h);
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
}

fn serve() {
    let mut input = std::io::stdin().lock().lines();
    let mut output = std::io::stdout().lock();
    while let Some(Ok(line)) = input.next() {
        let Ok(msg) = serde_json::from_str::<Json>(&line) else { return };
        let id = msg.get("id").cloned().unwrap_or(Json::Null);
        let result = match msg.get("method").and_then(Json::as_str) {
            Some("macros/list") => json!({"macros":[{"name":"Module.Wait","module":"Module","procedure":"Wait"}]}),
            Some("macros/run") => {
                let request = json!({"jsonrpc":"2.0","id":"write","method":"host/engine.execute",
                    "params":{"command":"cell.set","params":{"cell":"B1","input":"99"}}});
                if writeln!(output, "{request}").and_then(|_| output.flush()).is_err() {
                    return;
                }
                if input.next().is_none() {
                    return;
                }
                if !std::env::args().any(|a| a == "--vba") {
                    // Wait until cancelled; a closed host pipe cannot leave a busy loop.
                    let _ = input.next();
                    return;
                }
                json!({})
            }
            _ => json!({"protocolVersion":1}),
        };
        let reply = json!({"jsonrpc":"2.0","id":id,"result":result});
        if writeln!(output, "{reply}").and_then(|_| output.flush()).is_err() {
            return;
        }
    }
}
