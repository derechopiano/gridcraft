//! Shared macro transaction bookkeeping; native async transport owns no Session.
use crate::{Clipboard, HistoryEntry, Selection, Session, UiRequest};
use gridcraft_model::Workbook;
#[cfg(not(target_arch = "wasm32"))]
use serde_json::Value as Json;
use std::sync::Arc;

pub(crate) struct Snapshot {
    uid: u64,
    wb: Arc<Workbook>,
    selection: Selection,
    undo: Vec<HistoryEntry>,
    redo: Vec<HistoryEntry>,
    revision: u64,
    sheet_selections: Vec<Selection>,
    struct_log: Vec<gridcraft_collab::StructOp>,
    ui_requests: Vec<UiRequest>,
    clipboard: Option<Clipboard>,
    download_metadata: Option<String>,
}

impl Snapshot {
    pub(crate) fn capture(s: &Session) -> Option<Self> {
        let d = s.active()?;
        Some(Self {
            uid: d.uid,
            wb: d.wb.clone(),
            selection: d.selection.clone(),
            undo: d.undo.clone(),
            redo: d.redo.clone(),
            revision: d.revision,
            sheet_selections: d.sheet_selections.clone(),
            struct_log: d.struct_log.clone(),
            ui_requests: s.ui_requests.clone(),
            clipboard: s.clipboard.clone(),
            download_metadata: d.python_download_metadata.clone(),
        })
    }

    /// Discard nested history on success; restore all workbook effects on failure.
    pub(crate) fn restore(&self, s: &mut Session, success: bool) {
        if !success {
            s.ui_requests.clone_from(&self.ui_requests);
            s.clipboard.clone_from(&self.clipboard);
        }
        if let Some(d) = s.docs.iter_mut().find(|d| d.uid == self.uid) {
            d.undo.clone_from(&self.undo);
            d.redo.clone_from(&self.redo);
            d.revision = self.revision;
            if !success {
                d.python_download_metadata.clone_from(&self.download_metadata);
                d.wb = self.wb.clone();
                d.selection = self.selection.clone();
                d.sheet_selections.clone_from(&self.sheet_selections);
                d.struct_log.clone_from(&self.struct_log);
                d.calc.rebuild(&d.wb);
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) struct MacroRun {
    pub(crate) call: gridcraft_plugin::macro_call::MacroCall,
    snapshot: Snapshot,
    name: String,
    plugin: String,
    error: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn start(s: &mut Session, p: &Json) -> crate::Result<Json> {
    use serde_json::json;
    const ID: &str = "developer.startMacro";
    if s.macro_running() {
        return Err(crate::cmd::bad(ID, "a macro is already running"));
    }
    if !s.collab.rooms.is_empty() {
        return Err(crate::cmd::bad(ID, "asynchronous macros are unavailable during live collaboration"));
    }
    let name =
        p.get("name").and_then(Json::as_str).filter(|n| !n.trim().is_empty()).ok_or_else(|| crate::cmd::bad(ID, "missing `name`"))?.to_string();
    let providers = gridcraft_plugin::macro_plugins();
    let plugin = match p.get("plugin").and_then(Json::as_str) {
        Some(id) => providers.iter().find(|(pid, _)| pid == id),
        None => providers.first(),
    }
    .map(|(id, _)| id.clone())
    .ok_or_else(|| crate::cmd::bad(ID, "no matching macro provider is loaded"))?;
    let snapshot = Snapshot::capture(s).ok_or(crate::EngineError::NoDocument)?;
    let mut params = json!({"name":name, "context":crate::cmd::plugins::brief_context(s)});
    if let Some(control) = p.get("control").filter(|v| v.is_object()) {
        params["control"] = control.clone();
    }
    let timeout = match p.get("timeoutMs") {
        None => 30_000,
        Some(value) => {
            value.as_u64().filter(|n| (1..=30_000).contains(n)).ok_or_else(|| crate::cmd::bad(ID, "`timeoutMs` must be between 1 and 30000"))?
        }
    };
    let call = gridcraft_plugin::macro_call::MacroCall::start(&plugin, params, std::time::Duration::from_millis(timeout))
        .map_err(crate::EngineError::Other)?;
    s.macro_run = Some(MacroRun { call, snapshot, name, plugin, error: None });
    Ok(json!({"running":true}))
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn abort(s: &mut Session) {
    if let Some(run) = s.macro_run.take() {
        run.call.cancel();
        run.snapshot.restore(s, false);
        s.ui_requests.push(UiRequest::Message("Macro cancelled before changing workbooks.".into()));
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn poll(s: &mut Session) -> crate::Result<Json> {
    use serde_json::json;
    let Some(mut run) = s.macro_run.take() else { return Ok(json!({"running":false})) };
    // The Session stays on its owner thread. Only this section may execute host edits.
    s.macro_host_call = true;
    let pumped = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        if s.active().is_none_or(|d| d.uid != run.snapshot.uid) {
            run.error = Some("the macro's workbook is no longer active".into());
            run.call.cancel();
        }
        // Bound host work per poll; a provider cannot monopolize a frame with tiny requests.
        for _ in 0..32 {
            let Some(request) = run.call.next_request() else { break };
            let answer = crate::cmd::developer::host_call(s, &run.plugin, &request.method, &request.params);
            if let Err(error) = request.answer(answer) {
                run.error = Some(error);
                run.call.cancel();
                break;
            }
        }
    }));
    s.macro_host_call = false;
    if pumped.is_err() {
        run.error = Some("the macro host failed while processing a workbook request".into());
        run.call.cancel();
    }
    let Some(result) = run.call.poll_result() else {
        s.macro_run = Some(run);
        return Ok(json!({"running":true}));
    };
    let result = run.error.take().map_or(result, Err);
    run.snapshot.restore(s, result.is_ok());
    result.map_err(crate::EngineError::Other)?;
    if let Some(d) = s.docs.iter_mut().find(|d| d.uid == run.snapshot.uid)
        && !Arc::ptr_eq(&d.wb, &run.snapshot.wb)
    {
        d.undo.push(HistoryEntry { label: "Run Macro".into(), wb: run.snapshot.wb, selection: run.snapshot.selection });
        if d.undo.len() > crate::UNDO_LIMIT {
            d.undo.remove(0);
        }
        d.redo.clear();
        d.revision = d.revision.saturating_add(1);
    }
    s.journal.push(("developer.runMacro".into(), json!({"name":run.name, "plugin":run.plugin})));
    if s.journal.len() > 10_000 {
        s.journal.drain(..5_000);
    }
    Ok(json!({"running":false,"ran":run.name}))
}
