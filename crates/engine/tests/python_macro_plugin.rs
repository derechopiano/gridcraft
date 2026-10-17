//! Compatibility proof against the real plugin host from #1123, not a mock host.
#[path = "support/python_macro_provider.rs"]
mod provider;
#[path = "support/python_macro_worker.rs"]
mod worker;

use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use serde_json::json;
const PLUGIN: &str = "org.gridcraft.test-python-macros";

#[allow(clippy::disallowed_methods)] // Native integration benchmark; never compiled into the web app.
fn main() {
    let python = std::env::var("GRIDCRAFT_TEST_PYTHON").expect("Set GRIDCRAFT_TEST_PYTHON to Python 3.11+");
    let uv = std::env::var("GRIDCRAFT_TEST_UV").expect("Set GRIDCRAFT_TEST_UV to uv");
    let dir = std::env::temp_dir().join(format!("gridcraft-python-plugin-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::write(dir.join("provider.py"), provider::SOURCE).unwrap();
    std::fs::write(dir.join("worker.py"), worker::SOURCE).unwrap();
    let external = std::env::var_os("GRIDCRAFT_TEST_PYTHON_PLUGIN");
    let mut manifest = json!({"id":PLUGIN,"name":"Test Python macros","macros":{"language":"Python"},
        "requires":{"features":["pythonProjects","engineCommands"]},
        "run":{"command":python,"args":[dir.join("provider.py"),dir.join("worker.py"),uv,python,dir]}});
    if let Some(path) = &external {
        let root = std::path::PathBuf::from(path);
        manifest = serde_json::from_slice(&std::fs::read(root.join("plugin-macros.json")).unwrap()).unwrap();
        manifest["id"] = json!(PLUGIN);
        manifest["run"] = json!({"command":python,"args":[root.join("main.py"),"--workbook-macros",
            "--uv-executable",uv,"--python-interpreter",python,"--macro-cache",dir.join("macro-cache")]});
        println!("Testing extension of sirmmo's Python Tools: {}", root.display());
    }
    std::fs::write(dir.join("plugin.json"), manifest.to_string()).unwrap();
    let mut s = Session::new();
    s.new_workbook();
    let mut tools_id = None;
    if let Some(path) = &external {
        let root = std::path::PathBuf::from(path);
        let mut tools: serde_json::Value = serde_json::from_slice(&std::fs::read(root.join("plugin.json")).unwrap()).unwrap();
        tools_id = Some(tools["id"].as_str().unwrap().to_owned());
        tools["run"] = json!({"command":python,"args":[root.join("main.py")]});
        let tools_dir = dir.join("tools");
        std::fs::create_dir_all(&tools_dir).unwrap();
        std::fs::write(tools_dir.join("plugin.json"), tools.to_string()).unwrap();
        s.execute("plugins.load", json!({"path":tools_dir})).unwrap();
    }
    s.execute("plugins.load", json!({"path":dir})).unwrap();
    if external.is_some() {
        s.execute("cell.set", json!({"cell":"A2","input":"=PY.TWICE(4)"})).unwrap();
        assert_eq!(value(&s, "A2"), Value::Number(8.0), "existing Python Tools functions still work alongside macros");
        // Recalculation must use the ordinary Tools provider, not re-enter the busy macro provider.
        s.execute("cell.set", json!({"cell":"A3","input":"=PY.TWICE(B1)"})).unwrap();
    }
    let vba = b"opaque VBA project preserved by the Python adapter".to_vec();
    std::sync::Arc::make_mut(&mut s.active_mut().unwrap().wb).vba_project = Some(vba.clone());
    let source = "import gridcraft as gc\ndef fill():\n    w = gc.active_workbook()\n    value = w.read('A1')[0][0]\n    w.write('B1', [[value * 2]])\n    w.execute('home.bold', range='B1')\ndef broken():\n    gc.active_workbook().write('C1', [[99]])\n    raise ValueError('synthetic failure')\n";
    let source = format!(
        "{source}\ndef bulk():\n    w = gc.active_workbook()\n    data = w.read('D1:M1000')\n    w.write('D1002', [[v * 2 for v in row] for row in data])\n"
    );
    let source = format!(
        "{source}\ndef literal_text():\n    w = gc.active_workbook()\n    w.write('P1', [['00123', '=not a formula', '1/2']])\n    w.execute('cell.set', cell='P2', input='=1+2')\n"
    );
    let source = format!("{source}\ndef wait_for_stop():\n    import time\n    gc.active_workbook().write('C3', [[123]])\n    time.sleep(60)\n");
    s.execute("developer.saveModule", json!({"plugin":PLUGIN,"name":"analysis","source":source})).unwrap();
    assert_eq!(s.active().unwrap().wb.vba_project.as_ref(), Some(&vba), "Python must not overwrite VBA storage");
    s.execute(
        "developer.saveModule",
        json!({"plugin":PLUGIN,"name":"inert","source":"raise RuntimeError('discovery executed source')\ndef never_run():\n    pass\n"}),
    )
    .unwrap();
    assert_eq!(s.execute("developer.modules", json!({"plugin":PLUGIN})).unwrap()["modules"][0]["source"], source);
    let macros = s.execute("developer.macros", json!({"plugin":PLUGIN})).unwrap();
    assert_eq!(macros["macros"][0]["language"], "Python");
    assert_eq!(macros["macros"][0]["plugin"], PLUGIN);
    assert_eq!(macros["macros"].as_array().unwrap().len(), 6, "AST discovery never imports modules");
    assert!(s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.fill"})).is_err(), "shared security must gate Python");
    s.execute("cell.set", json!({"cell":"A1","input":"21"})).unwrap();
    s.execute("developer.enableMacros", json!({})).unwrap();
    let before = s.active().unwrap().undo.len();
    let revision = s.active().unwrap().revision;
    let start = std::time::Instant::now();
    s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.fill"})).unwrap();
    println!("Offline uv first run: {:.1} ms", start.elapsed().as_secs_f64() * 1000.0);
    assert_eq!(value(&s, "B1"), Value::Number(42.0));
    if external.is_some() {
        assert_eq!(value(&s, "A3"), Value::Number(84.0), "a Python macro must recalculate existing Python Tools formulas correctly");
    }
    assert_eq!(s.active().unwrap().undo.len(), before + 1);
    assert_eq!(s.active().unwrap().revision, revision + 1, "one public revision for the entire macro");
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(value(&s, "B1"), Value::Empty);
    let mut timings = Vec::new();
    for _ in 0..5 {
        let start = std::time::Instant::now();
        s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.fill"})).unwrap();
        timings.push(start.elapsed().as_secs_f64() * 1000.0);
        s.execute("edit.undo", json!({})).unwrap();
    }
    timings.sort_by(f64::total_cmp);
    println!("Offline uv repeat runs: median {:.1} ms, maximum {:.1} ms", timings[2], timings[4]);
    let before = s.active().unwrap().undo.len();
    s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.literal_text"})).unwrap();
    for (cell, expected) in [("P1", "00123"), ("Q1", "=not a formula"), ("R1", "1/2")] {
        assert_eq!(value(&s, cell), Value::Text(expected.into()), "Python writes must preserve literal text via protocol 2.2");
    }
    assert_eq!(value(&s, "P2"), Value::Number(3.0), "explicit commands still support formulas");
    assert_eq!(s.active().unwrap().undo.len(), before + 1);
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(value(&s, "P1"), Value::Empty);
    assert_eq!(value(&s, "P2"), Value::Empty);
    let before = s.active().unwrap().undo.len();
    let redo_before = s.active().unwrap().redo.len();
    let revision = s.active().unwrap().revision;
    let error = s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.broken"})).unwrap_err().to_string();
    assert!(error.contains("synthetic failure"), "{error}");
    assert_eq!(value(&s, "C1"), Value::Empty, "shared host rolls back failed Python runs");
    assert_eq!(s.active().unwrap().undo.len(), before);
    assert_eq!(s.active().unwrap().redo.len(), redo_before, "failed macro must preserve existing redo history");
    assert_eq!(s.active().unwrap().revision, revision, "failed macro must preserve the public revision");
    let data: Vec<Vec<_>> = (0..1000).map(|r| (0..10).map(|c| r * 10 + c).collect()).collect();
    s.execute("range.setValues", json!({"range":"D1","values":data})).unwrap();
    let start = std::time::Instant::now();
    s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.bulk"})).unwrap();
    println!("10,000-cell read/transform/write after worker reset: {:.1} ms", start.elapsed().as_secs_f64() * 1000.0);
    assert_bulk(&s);
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(value(&s, "D1002"), Value::Empty);
    let mut python_times = Vec::new();
    let mut native_times = Vec::new();
    for _ in 0..5 {
        let start = std::time::Instant::now();
        s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.bulk"})).unwrap();
        python_times.push(start.elapsed().as_secs_f64() * 1000.0);
        assert_bulk(&s);
        s.execute("edit.undo", json!({})).unwrap();
        let start = std::time::Instant::now();
        let read = s.execute("sheet.read", json!({"range":"D1:M1000"})).unwrap();
        let doubled: Vec<Vec<f64>> = read["values"]
            .as_array()
            .unwrap()
            .iter()
            .map(|row| row.as_array().unwrap().iter().map(|v| v.as_f64().unwrap() * 2.0).collect())
            .collect();
        s.execute("range.setValues", json!({"range":"D1002","values":doubled})).unwrap();
        native_times.push(start.elapsed().as_secs_f64() * 1000.0);
        assert_bulk(&s);
        s.execute("edit.undo", json!({})).unwrap();
    }
    python_times.sort_by(f64::total_cmp);
    native_times.sort_by(f64::total_cmp);
    println!("Warm 10,000-cell read/transform/write: median {:.1} ms, maximum {:.1} ms", python_times[2], python_times[4]);
    println!("Same-build native engine baseline (10,000 cells): median {:.1} ms, maximum {:.1} ms", native_times[2], native_times[4]);
    // The production provider must work through the new shared background transaction too.
    s.execute("developer.startMacro", json!({"plugin":PLUGIN,"name":"analysis.fill"})).unwrap();
    pump_async(&mut s, false).unwrap();
    assert_eq!(value(&s, "B1"), Value::Number(42.0));
    if external.is_some() {
        assert_eq!(value(&s, "A3"), Value::Number(84.0));
    }
    s.execute("edit.undo", json!({})).unwrap();
    // Process-tree ownership belongs to the production extension, not the minimal protocol fixture.
    if external.is_some() {
        let before_stop = s.active().unwrap().clone();
        s.execute("developer.startMacro", json!({"plugin":PLUGIN,"name":"analysis.wait_for_stop","timeoutMs":5000})).unwrap();
        pump_async(&mut s, true).unwrap();
        s.execute("developer.cancelMacro", json!({})).unwrap();
        assert!(pump_async(&mut s, false).unwrap_err().to_string().contains("cancelled"));
        assert!(std::sync::Arc::ptr_eq(&s.active().unwrap().wb, &before_stop.wb));
        assert_eq!(s.active().unwrap().revision, before_stop.revision);
        assert_eq!(s.active().unwrap().undo.len(), before_stop.undo.len());
        assert_eq!(s.active().unwrap().redo.len(), before_stop.redo.len());
        if external.is_some() {
            s.execute("cell.set", json!({"cell":"A4","input":"=PY.TWICE(5)"})).unwrap();
            assert_eq!(value(&s, "A4"), Value::Number(10.0), "Stop must leave ordinary Python Tools available");
        }
    }
    let before = s.active().unwrap().wb.clone();
    assert!(s.execute("developer.saveModule", json!({"plugin":PLUGIN,"name":"../escape","source":""})).is_err());
    assert!(std::sync::Arc::ptr_eq(&before, &s.active().unwrap().wb));
    let before = s.active().unwrap().undo.len();
    s.execute("developer.saveModule", json!({"plugin":PLUGIN,"name":"analysis","source":source})).unwrap();
    assert_eq!(s.active().unwrap().undo.len(), before, "unchanged source is a no-op");
    let state_source =
        format!("{source}\n_state = 0\ndef state():\n    global _state\n    _state += 1\n    gc.active_workbook().write('C2', [[_state]])\n");
    s.execute("developer.saveModule", json!({"plugin":PLUGIN,"name":"analysis","source":state_source})).unwrap();
    for _ in 0..2 {
        s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.state"})).unwrap();
    }
    assert_eq!(value(&s, "C2"), Value::Number(2.0), "unchanged project reuses its worker");
    let changed = format!("{state_source}\n# Edited project\n");
    s.execute("developer.saveModule", json!({"plugin":PLUGIN,"name":"analysis","source":changed})).unwrap();
    s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.state"})).unwrap();
    assert_eq!(value(&s, "C2"), Value::Number(1.0), "source changes invalidate the worker");
    let first_book = s.active_index();
    s.new_workbook();
    s.execute("developer.saveModule", json!({"plugin":PLUGIN,"name":"analysis","source":changed})).unwrap();
    s.execute("developer.enableMacros", json!({})).unwrap();
    s.execute("developer.runMacro", json!({"plugin":PLUGIN,"name":"analysis.state"})).unwrap();
    assert_eq!(value(&s, "C2"), Value::Number(1.0), "identical code in another workbook must not share Python globals");
    s.set_active(first_book);
    s.execute("developer.deleteModule", json!({"plugin":PLUGIN,"name":"analysis"})).unwrap();
    s.execute("developer.deleteModule", json!({"plugin":PLUGIN,"name":"inert"})).unwrap();
    assert!(s.execute("developer.modules", json!({"plugin":PLUGIN})).unwrap()["modules"].as_array().unwrap().is_empty());
    assert!(gridcraft_engine::plugin::unload(PLUGIN));
    if let Some(id) = tools_id {
        assert!(gridcraft_engine::plugin::unload(&id));
    }
    println!("Python provider compatibility: shared editing, discovery, security, real offline uv execution, undo and rollback passed");
}
fn value(s: &Session, cell: &str) -> Value {
    s.active().unwrap().wb.active().unwrap().value(CellRef::parse(cell).unwrap())
}
fn assert_bulk(s: &Session) {
    for row in 0..1000 {
        for col in 0..10 {
            assert_eq!(
                s.active().unwrap().wb.active().unwrap().value(CellRef::new(row + 1001, col + 3)),
                Value::Number(f64::from((row * 10 + col) * 2))
            );
        }
    }
}

#[allow(clippy::disallowed_methods)] // Native real-provider regression only.
fn pump_async(s: &mut Session, written: bool) -> gridcraft_engine::Result<serde_json::Value> {
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(8);
    loop {
        let result = s.execute("developer.pollMacro", json!({}))?;
        if result["running"] == false || (written && value(s, "C3") == Value::Number(123.0)) {
            return Ok(result);
        }
        assert!(std::time::Instant::now() < deadline, "Python async run did not make progress");
        std::thread::sleep(std::time::Duration::from_millis(1));
    }
}
