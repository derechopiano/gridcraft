//! The GridCraft engine façade.
//!
//! Every user-visible action is a command with a stable id (`cell.set`, `home.bold`,
//! `data.sortAscending`…) and JSON parameters. The egui UI, the CLI, the control channel and MCP
//! all go through [`Session::execute`], so every action is scriptable and journaled.
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
#![forbid(unsafe_code)]

pub mod catalog;
pub mod cf;
mod clip_html;
pub mod cmd;
pub mod collab;
pub mod display;
pub mod fill;
pub mod io;
pub mod links;
pub mod live;
mod macro_transaction;
pub mod outline;
pub mod paper;
pub mod pivot;
pub mod pivot_calc;
pub mod sample;
pub mod selection;
pub mod tables;

use std::sync::Arc;

use gridcraft_calc::Calc;
use gridcraft_core::{CellRef, RangeRef};
use gridcraft_model::Workbook;
use serde::Serialize;
use serde_json::Value as Json;

pub use cmd::{CommandInfo, CommandSpec, command_specs, find_command, register_test_command};
pub use gridcraft_calc as calc;
pub use gridcraft_collab as collab_core;
pub use gridcraft_core as core;
pub use gridcraft_formula as formula;
pub use gridcraft_model as model;
pub use gridcraft_numfmt as numfmt;
pub use gridcraft_plugin as plugin;
pub use selection::Selection;

#[derive(Debug, thiserror::Error)]
pub enum EngineError {
    #[error("unknown command `{0}`")]
    UnknownCommand(String),
    #[error("command `{0}` is not available right now: {1}")]
    Disabled(String, String),
    #[error("invalid parameters for `{cmd}`: {msg}")]
    BadParams { cmd: String, msg: String },
    #[error("no workbook open")]
    NoDocument,
    #[error("{0}")]
    Other(String),
    #[error("internal error in `{0}` (the workbook was kept as it was): {1}")]
    Internal(String, String),
}

pub type Result<T> = std::result::Result<T, EngineError>;

#[derive(Clone, Debug)]
pub struct HistoryEntry {
    pub label: String,
    pub wb: Arc<Workbook>,
    pub selection: Selection,
}

/// The clipboard: a block of cells (with formulas and styles) or marching-ants reference.
#[derive(Clone, Debug)]
pub struct Clipboard {
    /// Source workbook snapshot, sheet and range.
    pub wb: Arc<Workbook>,
    pub sheet: usize,
    pub range: RangeRef,
    /// The range as copied, before whole rows/columns were trimmed to the used range
    /// (Insert Copied Cells inserts whole lines for a row or column copy).
    pub selected: RangeRef,
    pub cut: bool,
    /// Document uid of the source (a cut only moves within the same workbook).
    pub doc_uid: u64,
    /// Plain text (tab-separated) for the system clipboard.
    pub text: String,
    /// A multi-area copy: the areas as selected (for the marquee). `wb` then holds them slid
    /// together into one block at `range`. Empty for a single-range copy.
    pub areas: Vec<RangeRef>,
}

/// In-cell edit mode as shown in the status bar.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize)]
pub enum Mode {
    #[default]
    Ready,
    Enter,
    Edit,
    Point,
}

#[derive(Clone, Debug)]
pub struct DocState {
    pub wb: Arc<Workbook>,
    pub calc: Calc,
    pub selection: Selection,
    pub undo: Vec<HistoryEntry>,
    pub redo: Vec<HistoryEntry>,
    pub path: Option<String>,
    pub saved: Arc<Workbook>,
    pub revision: u64,
    pub uid: u64,
    pub title: String,
    /// Per-sheet selection memory.
    pub sheet_selections: Vec<Selection>,
    /// Allow Edit Ranges unlocked with their password in this session: (sheet name, title).
    pub unlocked_ranges: Vec<(String, String)>,
    /// The format `path`'s content was recognised as when its extension doesn't name one (an
    /// XLSX called `budget.dat`): Save writes it back in that format.
    pub content_kind: Option<io::FileKind>,
    /// The user enabled this workbook's macros (the message bar's Enable Content, or Macro
    /// Security set to enable all).
    pub macros_enabled: bool,
    /// User approval for uv downloads, bound to exact metadata in this session; never serialized.
    pub python_download_metadata: Option<String>,
    /// Row and column inserts and deletes of the command running now, for a live session (the
    /// ids of the lines move with them). Cleared after every command.
    pub struct_log: Vec<gridcraft_collab::StructOp>,
}

static NEXT_UID: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(1);
const UNDO_LIMIT: usize = 100;

impl DocState {
    pub fn new(wb: Workbook, path: Option<String>, title: String) -> DocState {
        let mut d = DocState::unsettled(wb, path, title);
        d.settle();
        d
    }

    /// A document whose formulas aren't calculated yet: [`DocState::settle`] calculates them.
    /// Opening a file sends `documentOpened` in between (plugin protocol 2.2).
    pub fn unsettled(wb: Workbook, path: Option<String>, title: String) -> DocState {
        let wb = Arc::new(wb);
        let sel = Selection::at(wb.active().map(|s| s.view_active).unwrap_or_default());
        DocState {
            saved: wb.clone(),
            wb,
            calc: Calc::new(),
            selection: sel,
            undo: vec![],
            redo: vec![],
            path,
            revision: 1,
            uid: NEXT_UID.fetch_add(1, std::sync::atomic::Ordering::Relaxed),
            title,
            sheet_selections: vec![],
            unlocked_ranges: vec![],
            content_kind: None,
            macros_enabled: false,
            python_download_metadata: None,
            struct_log: vec![],
        }
    }
    /// The first full calculation of a new or opened document; it stays unmodified.
    pub fn settle(&mut self) {
        let dirty = self.is_dirty();
        let doc = self.calling_document();
        // Calculated in place: the saved copy, the same workbook until now, isn't kept meanwhile.
        let saved = std::mem::replace(&mut self.saved, Arc::new(Workbook::new()));
        let saved = if dirty {
            Some(saved)
        } else {
            drop(saved);
            None
        };
        let (calc, wb) = (&mut self.calc, Arc::make_mut(&mut self.wb));
        gridcraft_calc::external::with_document(Some(doc), || calc.recalc_all(wb));
        self.saved = saved.unwrap_or_else(|| self.wb.clone());
    }

    /// The document as plugin functions are told (protocol 2.2): its file, name and session id.
    pub fn calling_document(&self) -> gridcraft_calc::external::CallingDocument {
        gridcraft_calc::external::CallingDocument {
            path: self.path.as_deref().map(cmd::plugins::absolute_path),
            name: self.display_title(),
            id: self.session_id(),
        }
    }

    /// An id for the document that stays the same while the app runs, saved or not (`doc-3`).
    pub fn session_id(&self) -> String {
        format!("doc-{}", self.uid)
    }

    /// Runs `f` with this document as the workbook plugin functions are calculated for.
    pub fn with_calling<R>(&self, f: impl FnOnce() -> R) -> R {
        gridcraft_calc::external::with_document(Some(self.calling_document()), f)
    }

    pub fn is_dirty(&self) -> bool {
        !Arc::ptr_eq(&self.wb, &self.saved)
    }
    pub fn sheet_index(&self) -> usize {
        self.wb.active_sheet
    }
    /// Records the view state that is saved with the file (`sheetView topLeftCell`, the active
    /// cell): the first visible cell of the scrolling area per sheet index in `tops`, and the
    /// active sheet's active cell. View state never makes the workbook dirty.
    pub fn store_view(&mut self, tops: &[(usize, CellRef)]) {
        let active = self.selection.active;
        let si = self.wb.active_sheet;
        let changed = tops.iter().any(|(i, c)| self.wb.sheets.get(*i).is_some_and(|sh| sh.view_top_left != *c))
            || self.wb.sheets.get(si).is_some_and(|sh| sh.view_active != active);
        if !changed {
            return;
        }
        let dirty = self.is_dirty();
        let wb = Arc::make_mut(&mut self.wb);
        for (i, c) in tops {
            if let Some(sh) = wb.sheets.get_mut(*i)
                && sh.view_top_left != *c
            {
                Arc::make_mut(sh).view_top_left = *c;
            }
        }
        if let Some(sh) = wb.sheets.get_mut(si)
            && sh.view_active != active
        {
            Arc::make_mut(sh).view_active = active;
        }
        if !dirty {
            self.saved = self.wb.clone();
        }
    }
    pub fn display_title(&self) -> String {
        self.path
            .as_deref()
            .and_then(|p| std::path::Path::new(p).file_name())
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| self.title.clone())
    }
}

/// Engine-side preferences.
#[derive(Clone, Debug, Serialize, serde::Deserialize)]
#[serde(default, rename_all = "camelCase")]
pub struct Prefs {
    /// After pressing Enter, move the selection: "down", "right", "up", "left", "none".
    pub enter_direction: String,
    pub autocomplete: bool,
    pub user_name: String,
    pub default_font: String,
    pub default_font_size: f32,
    pub sheets_in_new_workbook: usize,
    pub r1c1: bool,
    /// Words added with Add to Dictionary.
    pub user_dictionary: Vec<String>,
    /// View › Show › Formula Bar (an application option in Excel, not saved in the file).
    pub formula_bar: bool,
    /// Ids of the plugins turned off in File › Options › Add-ins.
    pub disabled_plugins: Vec<String>,
    /// Developer › Macro Security: `disable` (no macros run), `prompt` (a workbook's macros run
    /// once the user enables them for it; the default) or `enable` (all run).
    pub macro_security: String,
    /// Live sessions: this person's stable id (made on first use) and colour (`#RRGGBB`).
    pub collab_user_id: String,
    pub collab_color: String,
}

impl Default for Prefs {
    fn default() -> Self {
        Prefs {
            enter_direction: "down".into(),
            autocomplete: true,
            user_name: "GridCraft User".into(),
            default_font: gridcraft_model::DEFAULT_FONT.into(),
            default_font_size: gridcraft_model::DEFAULT_FONT_SIZE,
            sheets_in_new_workbook: 1,
            r1c1: false,
            user_dictionary: vec![],
            formula_bar: true,
            disabled_plugins: vec![],
            macro_security: "prompt".into(),
            collab_user_id: String::new(),
            collab_color: String::new(),
        }
    }
}

/// Requests from commands to the UI (open a dialog, show a message…).
#[derive(Clone, Debug, PartialEq, Serialize)]
pub enum UiRequest {
    Dialog(String, Json),
    Message(String),
    /// Start editing the active cell (F2) with optional initial text.
    EditCell(Option<String>),
    OpenUrl(String),
    /// Open an Office add-in's task pane (its `plugin` id, the pane's title and page). Panes with
    /// the same `pane` id are one pane.
    OfficePane {
        plugin: String,
        title: String,
        url: String,
        pane: Option<String>,
    },
    /// Call a function of an Office add-in's function file (a page with no pane); `control` is
    /// the manifest's control id (the function's `event.source.id`).
    OfficeFunction {
        plugin: String,
        url: String,
        function: String,
        control: String,
    },
}

#[derive(Default)]
pub struct Session {
    docs: Vec<DocState>,
    active: usize,
    pub clipboard: Option<Clipboard>,
    pub prefs: Prefs,
    pub mode: Mode,
    /// Commands run so far, without their passwords ([`cmd::journal_params`]).
    pub journal: Vec<(String, Json)>,
    pub ui_requests: Vec<UiRequest>,
    /// Format painter source style (sticky when double-clicked).
    pub format_painter: Option<(Arc<Workbook>, usize, RangeRef, bool)>,
    /// Find state (last search).
    pub last_find: Option<Json>,
    /// View › Show › Ruler is off (rulers in Page Layout view; an application option).
    pub hide_ruler: bool,
    /// Journal position where action recording started.
    pub recording: Option<usize>,
    /// Windows hidden with View › Hide.
    pub hidden_windows: Vec<usize>,
    /// Draw tab tool: "" (select), "pen" or "eraser"; pen colour and width.
    pub draw_tool: String,
    pub draw_color: String,
    pub draw_width: f32,
    /// Review › Show Comments: the Comments pane is open.
    pub comments_pane: bool,
    /// Review › Notes › Show All Notes: every note is displayed. An application display setting
    /// (like Excel's comment indicator option), so it's not saved in the file.
    pub show_all_notes: bool,
    /// Live sessions (plugin protocol 2).
    pub collab: collab::Collab,
    /// Commands running inside commands (a plugin command's steps): a live session shares the
    /// outermost one's effect as one change.
    depth: u32,
    #[cfg(not(target_arch = "wasm32"))]
    macro_run: Option<macro_transaction::MacroRun>,
    macro_host_call: bool,
}

impl Session {
    pub fn new() -> Session {
        Session::default()
    }
    /// True while a native macro transaction is in flight, including its host edits.
    pub fn macro_running(&self) -> bool {
        #[cfg(not(target_arch = "wasm32"))]
        {
            self.macro_run.is_some() || self.macro_host_call
        }
        #[cfg(target_arch = "wasm32")]
        {
            false
        }
    }

    fn abort_macro_before_document_change(&mut self) {
        #[cfg(not(target_arch = "wasm32"))]
        if !self.macro_host_call {
            macro_transaction::abort(self);
        }
    }

    pub fn documents(&self) -> &[DocState] {
        &self.docs
    }
    pub fn active_index(&self) -> usize {
        self.active
    }
    pub fn set_active(&mut self, i: usize) {
        if i < self.docs.len() && i != self.active {
            self.abort_macro_before_document_change();
            self.active = i;
        }
    }
    pub fn active(&self) -> Option<&DocState> {
        self.docs.get(self.active)
    }
    pub fn active_mut(&mut self) -> Option<&mut DocState> {
        self.docs.get_mut(self.active)
    }
    pub fn doc(&self) -> Result<&DocState> {
        self.active().ok_or(EngineError::NoDocument)
    }
    pub fn doc_mut(&mut self) -> Result<&mut DocState> {
        self.active_mut().ok_or(EngineError::NoDocument)
    }
    pub fn add_document(&mut self, d: DocState) -> usize {
        self.abort_macro_before_document_change();
        self.docs.push(d);
        self.active = self.docs.len() - 1;
        self.active
    }
    pub fn close_document(&mut self, i: usize) {
        if i < self.docs.len() {
            self.abort_macro_before_document_change();
            self.docs.remove(i);
            if self.active >= self.docs.len() {
                self.active = self.docs.len().saturating_sub(1);
            }
        }
    }
    /// A new blank workbook titled `BookN`.
    pub fn new_workbook(&mut self) -> usize {
        let n = (1..).find(|n| !self.docs.iter().any(|d| d.title == format!("Book{n}") && d.path.is_none())).unwrap_or(1);
        let mut wb = Workbook::new();
        for _ in 1..self.prefs.sheets_in_new_workbook.clamp(1, 255) {
            let name = wb.next_sheet_name();
            wb.sheets.push(Arc::new(gridcraft_model::Sheet::new(name)));
        }
        self.add_document(DocState::new(wb, None, format!("Book{n}")))
    }

    pub fn commands(&self) -> Vec<CommandInfo> {
        command_specs().iter().map(|c| c.info(self)).collect()
    }

    /// Runs a command by id. Undoable commands record a history step when the workbook changed.
    /// Panics inside commands are caught and reported; the workbook is restored. In a live
    /// session, the outermost command's change is shared as one update, and Undo and Redo are
    /// the person's own.
    pub fn execute(&mut self, id: &str, params: Json) -> Result<Json> {
        if self.macro_running() && !self.macro_host_call && !matches!(id, "developer.pollMacro" | "developer.cancelMacro") {
            return Err(EngineError::Disabled(id.into(), "a macro is running: stop it before editing or saving".into()));
        }
        if self.depth == 0
            && !self.collab.rooms.is_empty()
            && let Some(r) = collab::intercept(self, id, &params)
        {
            return r;
        }
        if !self.collab.rooms.is_empty()
            && let Some(spec) = find_command(id)
            && spec.undoable
            && let Some(why) = collab::read_only(self)
        {
            return Err(EngineError::Disabled(id.to_string(), why));
        }
        self.depth += 1;
        let r = self.execute_inner(id, params);
        self.depth = self.depth.saturating_sub(1);
        if self.depth == 0 && !self.macro_running() {
            collab::after_command(self);
            self.recalc_invalidated();
        }
        r
    }

    /// Recalculates the workbooks whose formulas call plugin functions invalidated since the last
    /// time (`host/functions/invalidate`, protocol 2.2). Runs after each command and between
    /// frames; a workbook that had no unsaved changes still has none. True when there were any.
    pub fn recalc_invalidated(&mut self) -> bool {
        if self.macro_running() {
            return false;
        }
        let names = gridcraft_plugin::take_invalidated();
        if names.is_empty() {
            return false;
        }
        cmd::plugins::recalc_users_keeping_saved(self, &names);
        true
    }

    fn execute_inner(&mut self, id: &str, params: Json) -> Result<Json> {
        let spec = find_command(id).ok_or_else(|| EngineError::UnknownCommand(id.to_string()))?;
        if let Err(why) = (spec.enabled)(self) {
            return Err(EngineError::Disabled(id.to_string(), why));
        }
        let before = self.active().map(|d| (d.uid, d.wb.clone(), d.selection.clone()));
        let params = if params.is_null() { Json::Object(Default::default()) } else { params };
        // Plugin functions calculated by the command are told which workbook calls them: the
        // active one (commands that calculate another document say so themselves).
        let calling = self.active().map(DocState::calling_document);
        let result =
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| gridcraft_calc::external::with_document(calling, || (spec.run)(self, &params))));
        let result = match result {
            Ok(r) => r,
            Err(p) => {
                let msg =
                    p.downcast_ref::<&str>().map(|s| s.to_string()).or_else(|| p.downcast_ref::<String>().cloned()).unwrap_or_else(|| "panic".into());
                // Restore the workbook as it was.
                if let Some((uid, wb, sel)) = &before
                    && let Some(d) = self.docs.iter_mut().find(|d| d.uid == *uid)
                {
                    d.wb = wb.clone();
                    d.selection = sel.clone();
                    d.calc.rebuild(&d.wb);
                }
                return Err(EngineError::Internal(id.to_string(), msg));
            }
        };
        if spec.journal && result.is_ok() && !self.macro_host_call {
            self.journal.push((id.to_string(), cmd::journal_params(&params)));
            if self.journal.len() > 10_000 {
                self.journal.drain(..5_000);
            }
        }
        if spec.undoable
            && !self.macro_host_call
            && let Some((uid, wb, sel)) = before
            && let Some(d) = self.docs.iter_mut().find(|d| d.uid == uid)
            && !Arc::ptr_eq(&d.wb, &wb)
        {
            d.undo.push(HistoryEntry { label: spec.label.trim_end_matches('…').to_string(), wb, selection: sel });
            if d.undo.len() > UNDO_LIMIT {
                d.undo.remove(0);
            }
            d.redo.clear();
            d.revision += 1;
        }
        result
    }

    pub fn run(&mut self, id: &str, params: Json) -> std::result::Result<Json, String> {
        self.execute(id, params).map_err(|e| e.to_string())
    }

    pub fn take_ui_requests(&mut self) -> Vec<UiRequest> {
        if self.macro_running() {
            return Vec::new();
        }
        std::mem::take(&mut self.ui_requests)
    }

    /// Runs a command for a caller with no UI (MCP, CLI): queued UI requests are dropped, and a command that only
    /// queued a file dialog because it got no file (`file.open`, `file.saveAs`, `file.save` on a workbook with no
    /// file yet, `data.getData`) is an error instead of a silent success.
    pub fn execute_headless(&mut self, id: &str, params: Json) -> Result<Json> {
        if id == "developer.startMacro" {
            return Err(EngineError::BadParams {
                cmd: id.into(),
                msg: "headless callers must use developer.runMacro, which waits for completion".into(),
            });
        }
        let r = self.execute(id, params);
        let requests = self.take_ui_requests();
        let asked_for_file = requests.iter().any(|q| matches!(q, UiRequest::Dialog(name, _) if name == "open" || name == "saveAs"));
        if r.is_ok() && asked_for_file {
            let what = match id {
                "file.open" => "`path` or `base64`",
                "data.getData" | "data.fromTextCsv" => "`path` or `text`",
                _ => "`path`",
            };
            let msg = format!("missing {what}: programmatic calls never open dialogs");
            return Err(EngineError::BadParams { cmd: id.to_string(), msg });
        }
        r
    }
}

/// Convenience: the active workbook's active sheet.
pub fn active_sheet(s: &Session) -> Option<&gridcraft_model::Sheet> {
    s.active().and_then(|d| d.wb.active())
}

pub fn cell(r: u32, c: u32) -> CellRef {
    CellRef::new(r, c)
}

#[cfg(test)]
mod tests;
