//! The Developer tab: macros run by a macro add-in. It shows once an add-in that runs
//! macros is loaded (`"macros": {"language": "VBA"}` in its manifest), with Visual Basic (the
//! workbook's code modules in an editor), Macros (list, run, edit, create, delete), Record Macro,
//! Macro Security and Add-ins. Everything goes through the `developer.*` commands; the add-in
//! owns the language, GridCraft the interface. A workbook with macros opens with them disabled
//! and a message bar to enable them (Macro Security: prompt).
//!
//! The workbook's own ribbon buttons (Developer › Ribbon Buttons, `developer.*RibbonButton`) run
//! its macros. They show on the tab they name, a built-in one or their own, whether or not a
//! macro add-in is loaded (without one, a click says to install one), and are saved with the
//! workbook so Excel shows them too.

use egui::{RichText, Ui, vec2};
use serde_json::{Value as Json, json};

use gridcraft_engine::model::RibbonGroup;

use crate::SheetApp;
use crate::icons::Icon;
use crate::widgets::{big_button, small_button};

/// The Developer tab's windows.
#[derive(Default)]
pub struct DevState {
    pub macros: Option<MacrosWindow>,
    pub editor: Option<Editor>,
    pub security: bool,
    pub running_macro: Option<String>,
    pub stopping_macro: bool,
    python_settings: Option<PythonSettings>,
    pub buttons: Option<ButtonsWindow>,
    /// The workbook's ribbon as last read: (workbook uid, workbook pointer, tabs, editable).
    book: Option<(u64, usize, std::rc::Rc<Vec<(String, Vec<RibbonGroup>)>>, bool)>,
    /// Workbooks whose "macros disabled" bar was shown (once per workbook).
    notice_seen: std::collections::HashSet<u64>,
}

#[derive(Default)]
pub struct MacrosWindow {
    pub provider: String,
    pub list: Vec<Json>,
    pub selected: usize,
    pub new_name: String,
    pub error: Option<String>,
}

struct PythonSettings {
    uid: u64,
    metadata: String,
    allow_downloads: bool,
    error: Option<String>,
}

/// Developer › Ribbon Buttons: the workbook's buttons, and the form adding one.
#[derive(Default)]
pub struct ButtonsWindow {
    pub macros: Vec<String>,
    pub macro_name: String,
    pub label: String,
    pub tab: String,
    pub group: String,
    pub image: String,
    pub small: bool,
    pub error: Option<String>,
}

#[derive(Default)]
pub struct Editor {
    pub provider: String,
    pub language: String,
    /// (name, kind, source)
    pub modules: Vec<(String, String, String)>,
    pub current: usize,
    pub text: String,
    pub dirty: bool,
    pub error: Option<String>,
    pub new_module: String,
}

/// Whether the Developer tab shows: an add-in that runs macros is loaded.
pub fn visible() -> bool {
    !gridcraft_engine::plugin::macro_plugins().is_empty()
}

fn language() -> String {
    gridcraft_engine::plugin::macro_plugins().into_iter().next().map(|(_, l)| l).unwrap_or_else(|| "VBA".into())
}

fn provider_language(provider: &str) -> String {
    gridcraft_engine::plugin::macro_plugins().into_iter().find(|(id, _)| id == provider).map(|(_, l)| l).unwrap_or_default()
}

fn provider_picker(ui: &mut Ui, salt: &str, provider: &mut String) -> bool {
    let before = provider.clone();
    egui::ComboBox::from_id_salt(salt).selected_text(provider_language(provider)).show_ui(ui, |ui| {
        for (id, language) in gridcraft_engine::plugin::macro_plugins() {
            ui.selectable_value(provider, id.clone(), format!("{language} ({id})"));
        }
    });
    *provider != before
}

/// Names a painted button for screen readers (and tests).
fn named(resp: egui::Response, label: &str) -> egui::Response {
    resp.widget_info(|| egui::WidgetInfo::labeled(egui::WidgetType::Button, true, label));
    resp
}

pub fn ribbon(app: &mut SheetApp, ui: &mut Ui) {
    let editor_label = if language() == "VBA" { "Visual\nBasic" } else { "Code\nEditor" };
    let editor_name = if language() == "VBA" { "Visual Basic" } else { "Code Editor" };
    if named(big_button(ui, Icon::Script, editor_label, "The workbook's code modules", false), editor_name).clicked() {
        open_editor(app, None);
    }
    if named(big_button(ui, Icon::Script, "Macros", "Run, edit or delete a macro", false), "Macros").clicked() {
        open_macros(app);
    }
    ui.vertical(|ui| {
        let recording = app.session.recording.is_some();
        let recorder = gridcraft_engine::plugin::macro_recording_plugins().into_iter().next().map(|(id, _)| id);
        let label = if recording { "Stop Recording" } else { "Record Macro" };
        if ui
            .add_enabled_ui(recording || recorder.is_some(), |ui| {
                named(
                    small_button(ui, if recording { Icon::Close } else { Icon::Plus }, label, "Record what you do as a macro in Module1", false),
                    label,
                )
            })
            .inner
            .clicked()
        {
            match app.session.run("developer.recordMacro", json!({"stop": recording, "plugin": recorder})) {
                Ok(r) if recording => {
                    let module = r.get("module").and_then(Json::as_str).unwrap_or("Module1").to_string();
                    app.after_engine();
                    open_editor_for(app, Some(&module), recorder.as_deref());
                }
                Ok(_) => app.after_engine(),
                Err(e) => app.message = Some(("Record Macro".into(), e)),
            }
        }
        if named(small_button(ui, Icon::Check, "Macro Security", "Which workbooks' macros may run", false), "Macro Security").clicked() {
            app.dev.security = true;
        }
    });
    if named(big_button(ui, Icon::Plus, "Ribbon\nButtons", "Buttons on the ribbon that run this workbook's macros", false), "Ribbon Buttons")
        .clicked()
    {
        open_buttons(app, None);
    }
    crate::ribbon::sep(ui);
    if named(big_button(ui, Icon::Script, "Add-ins", "Plugins and Office add-ins", false), "Add-ins").clicked() {
        app.open_dialog("addIns", json!({}));
    }
}

/// The workbook's ribbon tabs (GridCraft tab name, groups), read again when the workbook changed.
fn book(app: &mut SheetApp) -> (std::rc::Rc<Vec<(String, Vec<RibbonGroup>)>>, bool) {
    let Some(d) = app.session.active() else { return Default::default() };
    let key = (d.uid, std::sync::Arc::as_ptr(&d.wb) as usize);
    if let Some((uid, ptr, tabs, editable)) = &app.dev.book
        && (*uid, *ptr) == key
    {
        return (tabs.clone(), *editable);
    }
    let (r, editable) = gridcraft_engine::cmd::developer::current_ribbon(&app.session);
    let mut tabs: Vec<(String, Vec<RibbonGroup>)> = Vec::new();
    for t in r.tabs {
        let name = gridcraft_engine::cmd::developer::tab_name(&t);
        match tabs.iter_mut().find(|(n, _)| n.eq_ignore_ascii_case(&name)) {
            Some((_, groups)) => groups.extend(t.groups),
            None => tabs.push((name, t.groups)),
        }
    }
    let tabs = std::rc::Rc::new(tabs);
    app.dev.book = Some((key.0, key.1, tabs.clone(), editable));
    (tabs, editable)
}

/// The workbook's own tabs (not GridCraft's or a plugin's), for the tab row.
pub fn book_tabs(app: &mut SheetApp, shown: &[&str]) -> Vec<String> {
    book(app).0.iter().map(|(n, _)| n.clone()).filter(|n| !shown.iter().any(|s| s.eq_ignore_ascii_case(n)) && n != "Developer").collect()
}

/// Whether the workbook has groups on `tab`.
pub fn has_book_groups(app: &mut SheetApp, tab: &str) -> bool {
    book(app).0.iter().any(|(n, g)| n.eq_ignore_ascii_case(tab) && !g.is_empty())
}

/// The picture names offered for a button (Excel's `imageMso` names) and the icon GridCraft draws.
const PICTURES: &[(&str, Icon)] = &[
    ("MacroPlay", Icon::Script),
    ("HappyFace", Icon::Check),
    ("FileSave", Icon::Save),
    ("FilePrint", Icon::Print),
    ("Copy", Icon::Copy),
    ("Paste", Icon::Paste),
    ("Cut", Icon::Cut),
    ("Undo", Icon::Undo),
    ("Redo", Icon::Redo),
];

fn picture(name: Option<&str>) -> Icon {
    let Some(name) = name else { return Icon::Script };
    PICTURES
        .iter()
        .find(|(n, _)| n.eq_ignore_ascii_case(name))
        .map(|(_, i)| *i)
        .or_else(|| crate::plugin_ui::icon_named(name))
        .unwrap_or(Icon::Script)
}

/// Draws the workbook's groups on `tab` (after its own groups, separated when there are any).
pub fn book_ribbon_groups(app: &mut SheetApp, ui: &mut Ui, tab: &str, separate_first: bool) {
    let (tabs, _) = book(app);
    let groups: Vec<&RibbonGroup> = tabs.iter().filter(|(n, _)| n.eq_ignore_ascii_case(tab)).flat_map(|(_, g)| g).collect();
    let mut press = None;
    for (i, g) in groups.iter().enumerate() {
        if separate_first || i > 0 {
            crate::ribbon::sep(ui);
        }
        let mut k = 0;
        while let Some(b) = g.buttons.get(k) {
            let tip = b.screentip.clone().unwrap_or_else(|| format!("Runs {}", b.on_action));
            if b.large {
                if named(big_button(ui, picture(b.image.as_deref()), &b.label, &tip, false), &b.label).clicked() {
                    press = Some(b.id.clone());
                }
                k += 1;
                continue;
            }
            let column: Vec<_> = g.buttons.iter().skip(k).take_while(|b| !b.large).take(3).collect();
            k += column.len();
            ui.vertical(|ui| {
                for b in column {
                    let tip = b.screentip.clone().unwrap_or_else(|| format!("Runs {}", b.on_action));
                    if named(small_button(ui, picture(b.image.as_deref()), &b.label, &tip, false), &b.label).clicked() {
                        press = Some(b.id.clone());
                    }
                }
            });
        }
    }
    if let Some(id) = press {
        if let Err(e) = app.session.run("developer.pressRibbonButton", json!({"id": id})) {
            app.message = Some(("Run Macro".into(), e));
        }
        app.after_engine();
    }
}

/// Opens Developer › Ribbon Buttons, with `macro_name` filled in.
pub fn open_buttons(app: &mut SheetApp, macro_name: Option<&str>) {
    let mut w = ButtonsWindow { tab: "Macros".into(), group: "Macros".into(), image: "MacroPlay".into(), ..Default::default() };
    let providers: std::collections::HashSet<_> = gridcraft_engine::plugin::macro_ribbon_plugins().into_iter().map(|(id, _)| id).collect();
    if let Ok(r) = app.session.run("developer.macros", json!({})) {
        w.macros = r
            .get("macros")
            .and_then(Json::as_array)
            .into_iter()
            .flatten()
            .filter(|m| m.get("plugin").and_then(Json::as_str).is_some_and(|id| providers.contains(id)))
            .filter_map(|m| m.get("name").and_then(Json::as_str))
            .map(str::to_string)
            .collect();
    }
    app.after_engine();
    w.macro_name = macro_name.map(str::to_string).or_else(|| w.macros.first().cloned()).unwrap_or_default();
    app.dev.buttons = Some(w);
}

fn buttons_window(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(mut w) = app.dev.buttons.take() else { return };
    let mut open = true;
    let mut close = false;
    let (tabs, editable) = book(app);
    let mut remove = None;
    egui::Window::new(app.ui.language.tr("Ribbon Buttons"))
        .id(egui::Id::new("dev_buttons"))
        .collapsible(false)
        .resizable(false)
        .default_width(480.0)
        .open(&mut open)
        .show(ctx, |ui| {
            ui.weak("Buttons that run this workbook's macros. They are saved with the workbook, and Excel shows them too.");
            egui::Frame::group(ui.style()).show(ui, |ui| {
                ui.set_min_width(440.0);
                egui::ScrollArea::vertical().max_height(160.0).show(ui, |ui| {
                    if tabs.iter().all(|(_, g)| g.is_empty()) {
                        ui.weak("No buttons yet.");
                    }
                    for (tab, groups) in tabs.iter() {
                        for g in groups {
                            for b in &g.buttons {
                                ui.horizontal(|ui| {
                                    ui.label(format!("{tab} › {} › {}", g.label, b.label));
                                    ui.weak(format!("runs {}", b.on_action));
                                    if editable && ui.small_button("Remove").clicked() {
                                        remove = Some(b.id.clone());
                                    }
                                });
                            }
                        }
                    }
                });
            });
            if !editable {
                ui.label(RichText::new("This workbook's ribbon came from another program with more than buttons: its buttons run, but GridCraft doesn't change it.").weak());
                if ui.button("Close").clicked() {
                    close = true;
                }
                return;
            }
            ui.add_space(4.0);
            ui.strong("New button");
            egui::Grid::new("dev_button_form").num_columns(2).spacing(vec2(8.0, 4.0)).show(ui, |ui| {
                ui.label("Macro:");
                egui::ComboBox::from_id_salt("dev_button_macro").selected_text(w.macro_name.clone()).width(260.0).show_ui(ui, |ui| {
                    for m in &w.macros {
                        ui.selectable_value(&mut w.macro_name, m.clone(), m);
                    }
                });
                ui.end_row();
                ui.label("");
                ui.add(egui::TextEdit::singleline(&mut w.macro_name).hint_text("or type a name: Module1.Go").desired_width(260.0));
                ui.end_row();
                ui.label("Label:");
                ui.add(egui::TextEdit::singleline(&mut w.label).hint_text("the macro's name").desired_width(260.0));
                ui.end_row();
                ui.label("Tab:");
                egui::ComboBox::from_id_salt("dev_button_tab").selected_text(w.tab.clone()).width(260.0).show_ui(ui, |ui| {
                    for t in ["Macros", "Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View", "Developer"] {
                        ui.selectable_value(&mut w.tab, t.to_string(), app.ui.language.tr(t));
                    }
                });
                ui.end_row();
                ui.label("");
                ui.add(egui::TextEdit::singleline(&mut w.tab).hint_text("or a new tab's name").desired_width(260.0));
                ui.end_row();
                ui.label("Group:");
                ui.add(egui::TextEdit::singleline(&mut w.group).desired_width(260.0));
                ui.end_row();
                ui.label("Picture:");
                ui.horizontal(|ui| {
                    for (name, icon) in PICTURES {
                        let on = w.image == *name;
                        if crate::widgets::toggle_button(ui, *icon, on, name).clicked() {
                            w.image = name.to_string();
                        }
                    }
                });
                ui.end_row();
                ui.label("");
                ui.checkbox(&mut w.small, "Small button (three to a column)");
                ui.end_row();
            });
            ui.horizontal(|ui| {
                if ui.add_enabled(!w.macro_name.trim().is_empty(), egui::Button::new("Add")).clicked() {
                    let params = json!({"macro": w.macro_name.trim(), "label": w.label.trim(), "tab": w.tab.trim(), "group": w.group.trim(), "image": w.image, "large": !w.small});
                    match app.session.run("developer.addRibbonButton", params) {
                        Ok(_) => {
                            w.error = None;
                            w.label.clear();
                        }
                        Err(e) => w.error = Some(e),
                    }
                }
                if ui.button("Close").clicked() {
                    close = true;
                }
            });
            if let Some(e) = &w.error {
                ui.label(RichText::new(e).color(ui.visuals().error_fg_color));
            }
        });
    if let Some(id) = remove
        && let Err(e) = app.session.run("developer.removeRibbonButton", json!({"id": id}))
    {
        w.error = Some(e);
    }
    app.after_engine();
    if open && !close {
        app.dev.buttons = Some(w);
    }
}

pub fn open_macros(app: &mut SheetApp) {
    let mut w = MacrosWindow {
        provider: gridcraft_engine::plugin::macro_plugins().into_iter().next().map(|(id, _)| id).unwrap_or_default(),
        ..Default::default()
    };
    match app.session.run("developer.macros", json!({})) {
        Ok(r) => w.list = r.get("macros").and_then(Json::as_array).cloned().unwrap_or_default(),
        Err(e) => w.error = Some(e),
    }
    app.after_engine();
    app.dev.macros = Some(w);
}

pub fn open_editor(app: &mut SheetApp, module: Option<&str>) {
    open_editor_for(app, module, None);
}

fn open_editor_for(app: &mut SheetApp, module: Option<&str>, provider: Option<&str>) {
    let provider =
        provider.map(str::to_owned).or_else(|| gridcraft_engine::plugin::macro_plugins().into_iter().next().map(|(id, _)| id)).unwrap_or_default();
    let mut e = Editor { language: provider_language(&provider), provider, ..Default::default() };
    match app.session.run("developer.modules", json!({"plugin": e.provider})) {
        Ok(r) => {
            e.modules = r
                .get("modules")
                .and_then(Json::as_array)
                .into_iter()
                .flatten()
                .map(|m| {
                    let s = |k: &str| m.get(k).and_then(Json::as_str).unwrap_or("").to_string();
                    (s("name"), s("kind"), s("source"))
                })
                .collect();
        }
        Err(err) => e.error = Some(err),
    }
    app.after_engine();
    if let Some(m) = module {
        e.current = e.modules.iter().position(|(n, _, _)| n == m).unwrap_or(0);
    }
    e.text = e.modules.get(e.current).map(|(_, _, s)| s.clone()).unwrap_or_default();
    app.dev.editor = Some(e);
}

/// The message bar for a workbook whose macros are disabled (once per workbook).
pub fn notice(app: &mut SheetApp) {
    if app.notice.is_some() || !visible() || app.session.prefs.macro_security != "prompt" {
        return;
    }
    let Some(d) = app.session.active() else { return };
    if (d.wb.vba_project.is_none() && d.wb.python_macros.is_none()) || d.macros_enabled || app.dev.notice_seen.contains(&d.uid) {
        return;
    }
    app.dev.notice_seen.insert(d.uid);
    app.notice = Some(crate::Notice {
        text: "SECURITY WARNING  Macros have been disabled.".into(),
        kind: crate::NoticeKind::Warning,
        action: Some(("Enable Content".into(), "enableMacros".into())),
    });
}

/// Draws the open Developer windows.
pub fn windows(app: &mut SheetApp, ctx: &egui::Context) {
    macros_window(app, ctx);
    editor_window(app, ctx);
    security_window(app, ctx);
    buttons_window(app, ctx);
    python_settings_window(app, ctx);
}

fn macros_window(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(mut w) = app.dev.macros.take() else { return };
    let mut open = true;
    let mut close = false;
    let mut edit: Option<(String, String)> = None;
    let mut to_ribbon: Option<String> = None;
    let enabled = gridcraft_engine::cmd::developer::macros_allowed(&app.session);
    egui::Window::new(app.ui.language.tr("Macros"))
        .id(egui::Id::new("dev_macros"))
        .collapsible(false)
        .resizable(false)
        .default_width(460.0)
        .open(&mut open)
        .show(ctx, |ui| {
            if !enabled {
                ui.label(RichText::new("Macros are disabled for this workbook. Enable Content in the message bar, or change Macro Security.").weak());
            }
            ui.horizontal(|ui| {
                provider_picker(ui, "macro_create_provider", &mut w.provider);
                ui.label("Macro name:");
                ui.add(egui::TextEdit::singleline(&mut w.new_name).desired_width(260.0));
            });
            egui::Frame::group(ui.style()).show(ui, |ui| {
                ui.set_min_size(vec2(420.0, 180.0));
                egui::ScrollArea::vertical().max_height(200.0).show(ui, |ui| {
                    if w.list.is_empty() {
                        ui.weak("This workbook has no macros yet: type a name and Create.");
                    }
                    for (i, m) in w.list.iter().enumerate() {
                        let name = m.get("name").and_then(Json::as_str).unwrap_or("");
                        let language = m.get("language").and_then(Json::as_str).unwrap_or("");
                        if ui.selectable_label(i == w.selected, format!("{name} ({language})")).clicked() {
                            w.selected = i;
                            w.new_name = name.to_string();
                        }
                    }
                });
            });
            let chosen = w.list.get(w.selected).and_then(|m| m.get("name")).and_then(Json::as_str).map(str::to_string);
            ui.horizontal(|ui| {
                if ui.add_enabled(enabled && chosen.is_some(), egui::Button::new("Run")).clicked()
                    && let Some(name) = &chosen
                {
                    let plugin = w.list.get(w.selected).and_then(|m| m.get("plugin")).and_then(Json::as_str).unwrap_or("");
                    // Only Python opts into the new absolute deadline. Existing providers keep
                    // their synchronous, activity-based timeout and collaboration behaviour.
                    let command = if provider_language(plugin) == "Python" { "developer.startMacro" } else { "developer.runMacro" };
                    match app.session.run(command, json!({"name": name, "plugin": plugin})) {
                        Ok(result) => {
                            if result["running"] == true {
                                app.dev.running_macro = Some(name.clone());
                                app.dev.stopping_macro = false;
                            }
                            close = true;
                        }
                        Err(e) => w.error = Some(e),
                    }
                }
                if ui.add_enabled(chosen.is_some(), egui::Button::new("Edit")).clicked()
                    && let Some(m) = w.list.get(w.selected)
                {
                    edit = m
                        .get("module")
                        .and_then(Json::as_str)
                        .zip(m.get("plugin").and_then(Json::as_str))
                        .map(|(module, plugin)| (module.to_string(), plugin.to_string()));
                }
                let selected_provider = w.list.get(w.selected).and_then(|m| m.get("plugin")).and_then(Json::as_str);
                let ribbon_supported = gridcraft_engine::plugin::macro_ribbon_plugins().iter().any(|(id, _)| Some(id.as_str()) == selected_provider);
                if ui.add_enabled(chosen.is_some() && ribbon_supported, egui::Button::new("Add to Ribbon…")).clicked() {
                    to_ribbon = chosen.clone();
                }
                let name = w.new_name.trim().to_string();
                let valid = !name.is_empty()
                    && name.chars().all(|c| c.is_alphanumeric() || c == '_')
                    && !name.chars().next().is_some_and(|c| c.is_ascii_digit());
                if ui
                    .add_enabled(
                        valid
                            && !w.list.iter().any(|m| {
                                m.get("plugin").and_then(Json::as_str) == Some(w.provider.as_str())
                                    && m.get("procedure").and_then(Json::as_str) == Some(name.as_str())
                            }),
                        egui::Button::new("Create"),
                    )
                    .clicked()
                {
                    // A new procedure in the selected provider's Module1.
                    match create_macro(app, &name, &w.provider) {
                        Ok(module) => edit = Some((module, w.provider.clone())),
                        Err(error) => w.error = Some(error),
                    }
                }
                if ui.button("Close").clicked() {
                    close = true;
                }
            });
            if let Some(e) = &w.error {
                ui.label(RichText::new(e).color(ui.visuals().error_fg_color));
            }
        });
    app.after_engine();
    if let Some((module, provider)) = edit {
        open_editor_for(app, Some(&module), Some(&provider));
        return;
    }
    if let Some(name) = to_ribbon {
        open_buttons(app, Some(&name));
        return;
    }
    if open && !close {
        app.dev.macros = Some(w);
    }
}

/// Adds a language-appropriate procedure to Module1 (made when missing).
fn create_macro(app: &mut SheetApp, name: &str, provider: &str) -> Result<String, String> {
    let module = "Module1".to_string();
    let language = provider_language(provider);
    if language == "Python" {
        gridcraft_engine::model::python_macros::validate_module_name(name)?;
        if name.starts_with('_') {
            return Err("Macro names must be public Python functions".into());
        }
    }
    let modules = app.session.run("developer.modules", json!({"plugin": provider}))?;
    let existing = modules
        .get("modules")
        .and_then(Json::as_array)
        .cloned()
        .and_then(|m| m.into_iter().find(|x| x.get("name").and_then(Json::as_str) == Some("Module1")))
        .and_then(|m| m.get("source").and_then(Json::as_str).map(str::to_string))
        .map(|t| format!("{}\r\n\r\n", t.trim_end()))
        .unwrap_or_default();
    let source = procedure_source(&language, &existing, name);
    app.session.run("developer.saveModule", json!({"plugin": provider, "name": module, "source": source}))?;
    Ok(module)
}

fn procedure_source(language: &str, existing: &str, name: &str) -> String {
    if language == "Python" { format!("{existing}def {name}():\n    pass\n") } else { format!("{existing}Sub {name}()\r\n\r\nEnd Sub\r\n") }
}

fn module_source(language: &str, name: &str) -> String {
    if language == "Python" { String::new() } else { format!("Attribute VB_Name = \"{name}\"\r\n") }
}

fn editor_window(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(mut e) = app.dev.editor.take() else { return };
    let mut open = true;
    let editor_title = if e.language == "VBA" { app.ui.language.tr("Visual Basic").to_owned() } else { format!("{} Editor", e.language) };
    let title = format!("{} — {}", editor_title, app.session.active().map(|d| d.display_title()).unwrap_or_default());
    let mut switch_provider = None;
    egui::Window::new(title)
        .id(egui::Id::new("dev_editor"))
        .collapsible(false)
        .resizable(true)
        .default_size(vec2(820.0, 520.0))
        .open(&mut open)
        .show(ctx, |ui| {
            ui.horizontal(|ui| {
                let mut provider = e.provider.clone();
                if ui.add_enabled_ui(!e.dirty, |ui| provider_picker(ui, "editor_provider", &mut provider)).inner {
                    switch_provider = Some(provider);
                }
                if ui
                    .add_enabled(e.dirty, egui::Button::new("Save"))
                    .on_hover_text("Saves the module into the workbook (save the workbook as .xlsm to keep it)")
                    .clicked()
                {
                    let name = e.modules.get(e.current).map(|(n, _, _)| n.clone()).unwrap_or_else(|| "Module1".into());
                    match app.session.run("developer.saveModule", json!({"plugin": e.provider, "name": name, "source": e.text})) {
                        Ok(_) => {
                            if let Some(m) = e.modules.get_mut(e.current) {
                                m.2 = e.text.clone();
                            }
                            e.dirty = false;
                            e.error = None;
                        }
                        Err(err) => e.error = Some(err),
                    }
                }
                if ui.button("Run Macro…").clicked() {
                    open_macros(app);
                }
                ui.separator();
                ui.add(egui::TextEdit::singleline(&mut e.new_module).hint_text("Module2").desired_width(120.0));
                let valid = !e.new_module.trim().is_empty() && !e.modules.iter().any(|(n, _, _)| n.eq_ignore_ascii_case(e.new_module.trim()));
                if ui.add_enabled(valid, egui::Button::new("Insert Module")).clicked() {
                    let name = e.new_module.trim().to_string();
                    let source = module_source(&e.language, &name);
                    match app.session.run("developer.saveModule", json!({"plugin": e.provider, "name": name, "source": source})) {
                        Ok(_) => {
                            e.modules.push((name, "standard".into(), source.clone()));
                            e.current = e.modules.len() - 1;
                            e.text = source;
                            e.dirty = false;
                            e.new_module.clear();
                        }
                        Err(err) => e.error = Some(err),
                    }
                }
                if e.language == "Python" && ui.button(app.ui.language.tr("Python Settings")).clicked() {
                    match app.session.run("developer.pythonSettings", json!({})) {
                        Ok(settings) => {
                            if let Some(doc) = app.session.active() {
                                app.dev.python_settings = Some(PythonSettings {
                                    uid: doc.uid,
                                    metadata: settings["metadata"].as_str().unwrap_or_default().into(),
                                    allow_downloads: settings["allowDownloads"] == true,
                                    error: None,
                                });
                            }
                        }
                        Err(error) => e.error = Some(error),
                    }
                }
                let standard = e.modules.get(e.current).is_some_and(|(_, k, _)| k == "standard");
                if ui.add_enabled(standard, egui::Button::new("Remove Module")).clicked()
                    && let Some((name, _, _)) = e.modules.get(e.current).cloned()
                {
                    match app.session.run("developer.deleteModule", json!({"plugin": e.provider, "name": name})) {
                        Ok(_) => {
                            e.modules.remove(e.current);
                            e.current = 0;
                            e.text = e.modules.first().map(|(_, _, s)| s.clone()).unwrap_or_default();
                            e.dirty = false;
                        }
                        Err(err) => e.error = Some(err),
                    }
                }
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    ui.weak(&e.language);
                });
            });
            if let Some(err) = &e.error {
                ui.label(RichText::new(err).color(ui.visuals().error_fg_color));
            }
            ui.separator();
            ui.horizontal_top(|ui| {
                ui.vertical(|ui| {
                    ui.set_width(170.0);
                    ui.weak("Modules");
                    if e.modules.is_empty() {
                        ui.weak("No code yet: Insert Module.");
                    }
                    let mut pick = None;
                    for (i, (name, kind, _)) in e.modules.iter().enumerate() {
                        let label = if kind == "standard" { name.clone() } else { format!("{name} ({kind})") };
                        if ui.selectable_label(i == e.current, label).clicked() {
                            pick = Some(i);
                        }
                    }
                    if let Some(i) = pick
                        && i != e.current
                    {
                        if let Some(m) = e.modules.get_mut(e.current)
                            && e.dirty
                        {
                            // Unsaved text stays with its module until saved or closed.
                            m.2 = e.text.clone();
                        }
                        e.current = i;
                        e.text = e.modules.get(i).map(|(_, _, s)| s.clone()).unwrap_or_default();
                    }
                });
                ui.separator();
                egui::ScrollArea::both().id_salt("dev_code").show(ui, |ui| {
                    let resp = ui.add(
                        egui::TextEdit::multiline(&mut e.text)
                            .code_editor()
                            .desired_width(f32::INFINITY)
                            .desired_rows(24)
                            .lock_focus(true)
                            .interactive(!e.modules.is_empty()),
                    );
                    if resp.changed() {
                        e.dirty = true;
                    }
                });
            });
        });
    if let Some(provider) = switch_provider {
        open_editor_for(app, None, Some(&provider));
    } else if open {
        app.dev.editor = Some(e);
    }
}

fn security_window(app: &mut SheetApp, ctx: &egui::Context) {
    if !app.dev.security {
        return;
    }
    let mut open = true;
    let level = app.session.prefs.macro_security.clone();
    egui::Window::new(app.ui.language.tr("Macro Security"))
        .id(egui::Id::new("dev_security"))
        .collapsible(false)
        .resizable(false)
        .open(&mut open)
        .show(ctx, |ui| {
            for (value, text) in [
                ("disable", "Disable all macros"),
                ("prompt", "Disable macros with notification (enable them per workbook)"),
                ("enable", "Enable all macros (not recommended: a macro is a program)"),
            ] {
                if ui.radio(level == value, text).clicked() {
                    app.run_or_alert("developer.macroSecurity", json!({"level": value}));
                }
            }
            ui.add_space(6.0);
            if ui.button("Close").clicked() {
                app.dev.security = false;
            }
        });
    if !open {
        app.dev.security = false;
    }
}

/// Pumps the shared engine transaction before painting or accepting any other edits.
pub fn poll_macro(app: &mut SheetApp, ctx: &egui::Context) {
    if !app.session.macro_running() {
        return;
    }
    match app.session.run("developer.pollMacro", json!({})) {
        Ok(result) if result["running"] == true => ctx.request_repaint(),
        result => {
            app.dev.running_macro = None;
            app.dev.stopping_macro = false;
            app.after_engine();
            if let Err(error) = result {
                if error.contains("macro cancelled") {
                    app.toast = Some((app.ui.language.tr("Macro stopped.").into(), crate::now_ms()));
                } else {
                    app.message = Some((app.ui.language.tr("Macro").into(), error));
                }
            }
        }
    }
}

/// Edits stay disabled until the engine has completed cancellation and rollback.
pub fn macro_progress(app: &mut SheetApp, ctx: &egui::Context) {
    let lang = app.ui.language;
    let t = crate::theme::Tokens::get(ctx);
    let mut stop = ctx.input(|i| i.key_pressed(egui::Key::Escape));
    egui::Modal::new(egui::Id::new("macro_progress")).show(ctx, |ui| {
        ui.set_width(360.0);
        ui.heading(lang.tr(if app.dev.stopping_macro { "Stopping macro" } else { "Running macro" }));
        if let Some(name) = &app.dev.running_macro {
            ui.label(RichText::new(name).color(t.text));
        }
        ui.horizontal(|ui| {
            ui.spinner();
            ui.label(lang.tr("Workbook edits are paused until the macro finishes."));
        });
        stop |= ui.add_enabled(!app.dev.stopping_macro, egui::Button::new(lang.tr("Stop"))).clicked();
    });
    if stop && !app.dev.stopping_macro {
        match app.session.run("developer.cancelMacro", json!({})) {
            Ok(_) => app.dev.stopping_macro = true,
            Err(error) => app.message = Some((lang.tr("Macro").into(), error)),
        }
    }
    ctx.request_repaint();
}

fn python_settings_window(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(mut settings) = app.dev.python_settings.take() else { return };
    if app.session.active().is_none_or(|d| d.uid != settings.uid) {
        return;
    }
    let lang = app.ui.language;
    let mut open = true;
    let mut saved = false;
    egui::Window::new(lang.tr("Python Settings")).id(egui::Id::new("python_project_settings"))
        .default_width(580.0).collapsible(false).open(&mut open).show(ctx,|ui| {
            ui.label(lang.tr("PEP 723 interpreter and dependency metadata"));
            if ui.add(egui::TextEdit::multiline(&mut settings.metadata).code_editor().desired_rows(10).desired_width(f32::INFINITY)).changed() {
                settings.allow_downloads = false; // New metadata requires fresh approval.
            }
            ui.checkbox(&mut settings.allow_downloads,lang.tr("Allow uv to download packages and Python for this metadata in this session"));
            ui.label(lang.tr("Approval is not saved in the workbook. Without approval, only installed Python and cached dependencies are used. Source builds are disabled."));
            if let Some(error)=&settings.error { ui.colored_label(ui.visuals().error_fg_color,error); }
            if ui.button(lang.tr("Save")).clicked() {
                match app.session.run("developer.setPythonSettings",json!({"metadata":settings.metadata,"allowDownloads":settings.allow_downloads})) {
                    Ok(_)=>saved=true,
                    Err(error)=>settings.error=Some(error),
                }
            }
        });
    if open && !saved {
        app.dev.python_settings = Some(settings);
    }
}
