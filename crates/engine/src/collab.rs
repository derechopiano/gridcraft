//! Co-editing (plugin protocol 2): rooms, the `collab.*` commands, and the pump the app's main
//! loop runs between frames ([`Session::collab_poll`]).
//!
//! A room ties one open workbook to a transport plugin. The workbook is mirrored in a Yjs
//! document ([`gridcraft_collab::Replica`]): every committed command becomes one transaction
//! sent as `collab/update`; remote updates are applied between commands and mapped back into the
//! model, then recalculated. While a workbook is in a room, Undo and Redo are the person's own
//! (they revert only their changes). Transports only move bytes; with a room key every payload
//! is encrypted end to end (AES-256-GCM).

use std::sync::Arc;

use gridcraft_collab::awareness::{self, Awareness};
use gridcraft_collab::crypto::{self, Key};
use gridcraft_collab::{CollabError, Replica, StructOp};
use gridcraft_core::{CellRef, RangeRef};
use gridcraft_model::Workbook;
use gridcraft_plugin::push::{Incoming, Pushed};
use serde::{Deserialize, Serialize};
use serde_json::{Value as Json, json};

use crate::cmd::{CommandSpec, always, bad, cmd, has_doc, str_param};
use crate::{DocState, EngineError, Result, Session, UiRequest};

/// The transport of rooms joined with [`Session::collab_join_in_memory`]: messages wait in the
/// room for the caller ([`Session::collab_take_outgoing`]) instead of going to a plugin. Tests
/// connect sessions with it; a build without plugins can bridge it to a network itself.
pub const IN_MEMORY: &str = "(in-memory)";

/// Longest room id (protocol 2, section 2).
const MAX_ROOM: usize = 256;
/// After this many messages that don't decrypt, the status bar asks about the key.
const WRONG_KEY_AFTER: u32 = 10;
/// Awareness at most this often (20 a second).
const AWARENESS_MS: u64 = 50;
/// A transport that exited is started again after this long.
const RESTART_MS: u64 = 3_000;
/// Highlights of others' edits kept for the interface.
const MAX_REMOTE_EDITS: usize = 2_000;
/// Rooms per transport (protocol 2, section 7).
const MAX_ROOMS_PER_TRANSPORT: usize = 16;

/// A person in a room.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct User {
    pub id: String,
    pub name: String,
    pub color: String,
}

impl User {
    fn parse(v: &Json) -> Option<User> {
        let s = |k: &str| v.get(k).and_then(Json::as_str).map(str::to_string);
        let u = User { id: s("id")?, name: s("name")?, color: s("color").unwrap_or_else(|| "#808080".into()) };
        u.valid().then_some(u)
    }
    fn valid(&self) -> bool {
        let len = |t: &str| (1..=128).contains(&t.chars().count());
        len(&self.id) && len(&self.name) && valid_color(&self.color)
    }
}

/// `#RRGGBB`.
pub fn valid_color(c: &str) -> bool {
    c.len() == 7 && c.starts_with('#') && c[1..].chars().all(|ch| ch.is_ascii_hexdigit())
}

/// What the transport says about the room.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum RoomState {
    Connecting,
    Online,
    Offline,
    Ended,
}

impl RoomState {
    fn parse(s: &str) -> Option<RoomState> {
        Some(match s {
            "connecting" => RoomState::Connecting,
            "online" => RoomState::Online,
            "offline" => RoomState::Offline,
            "ended" => RoomState::Ended,
            _ => return None,
        })
    }
    pub fn as_str(self) -> &'static str {
        match self {
            RoomState::Connecting => "connecting",
            RoomState::Online => "online",
            RoomState::Offline => "offline",
            RoomState::Ended => "ended",
        }
    }
}

/// Someone else in the room, as their awareness state says.
#[derive(Clone, Debug, PartialEq)]
pub struct Peer {
    pub client_id: u64,
    pub user: User,
    /// Their selection: sheet index in our workbook, and ranges.
    pub selection: Option<(usize, Vec<RangeRef>)>,
    /// Their active cell.
    pub cursor: Option<(usize, CellRef)>,
}

/// One shared document.
pub struct Room {
    pub room: String,
    pub plugin: String,
    /// The workbook's `DocState::uid`.
    pub doc: u64,
    pub state: RoomState,
    /// Waiting for the room's document (`mode: "join"`): the workbook is read-only.
    pub joining: bool,
    /// `view`: the workbook is read-only.
    pub view_only: bool,
    pub user: User,
    /// Peers the transport listed (`collab/status`).
    pub listed: Vec<User>,
    /// The transport's or the app's message for the status bar.
    pub message: Option<String>,
    /// Another app's or schema's room: nothing syncs.
    pub refused: Option<String>,
    key: Option<Key>,
    replica: Replica,
    awareness: Awareness,
    decrypt_failures: u32,
    last_awareness: Option<Json>,
    last_awareness_at: u64,
    /// The transport exited at this time (it's started again after [`RESTART_MS`]).
    exited_at: Option<u64>,
    /// Messages for a transport served in memory ([`Session::collab_serve_in_memory`]).
    outbox: std::cell::RefCell<Vec<Json>>,
    in_memory: bool,
    /// Status-bar messages not shown yet ([`Session::collab_take_notices`]).
    notices: Vec<String>,
    /// This person is typing in the chat (awareness `typing`).
    pub typing: bool,
}

impl Room {
    pub fn client_id(&self) -> u64 {
        self.replica.client_id()
    }
    /// The room's title (`meta.title`).
    pub fn title(&self) -> Option<String> {
        self.replica.title()
    }
    pub fn encrypted(&self) -> bool {
        self.key.is_some()
    }
    /// Others present, from their awareness states.
    pub fn peers(&self) -> Vec<Peer> {
        self.awareness
            .present()
            .filter_map(|(id, st)| {
                let user = User::parse(st.get("user")?)?;
                let selection = st.get("selection").and_then(|s| self.replica.selection_from_ids(s));
                let cursor = st.get("cursor").and_then(|c| {
                    let cell = c.get("cell")?;
                    let (row, col) = (cell.get(0)?, cell.get(1)?);
                    let sel = json!({"sheet": c.get("sheet")?, "ranges": [[row, col, row, col]]});
                    let (si, r) = self.replica.selection_from_ids(&sel)?;
                    Some((si, r.first()?.start))
                });
                Some(Peer { client_id: id, user, selection, cursor })
            })
            .collect()
    }
    /// The others typing in the chat (awareness `typing: true`), by name.
    pub fn typing_names(&self) -> Vec<String> {
        let mut names: Vec<String> = self
            .awareness
            .present()
            .filter(|(_, st)| st.get("typing").and_then(Json::as_bool) == Some(true))
            .filter_map(|(_, st)| User::parse(st.get("user")?).map(|u| u.name))
            .collect();
        names.sort();
        names
    }
    fn info(&self, s: &Session) -> Json {
        let people: Vec<Json> = self.peers().iter().map(|p| json!({"clientId": p.client_id, "user": p.user})).collect();
        json!({
            "room": self.room,
            "plugin": self.plugin,
            "state": if self.joining { "joining" } else { self.state.as_str() },
            "peers": people,
            "listed": self.listed,
            "role": if self.view_only { "view" } else { "edit" },
            "encrypted": self.encrypted(),
            "clientId": self.client_id(),
            "user": self.user,
            "title": self.replica.title(),
            "document": s.docs.iter().find(|d| d.uid == self.doc).map(|d| d.display_title()),
            "message": self.refused.clone().or_else(|| self.message.clone()),
            "canUndo": self.replica.can_undo(),
            "canRedo": self.replica.can_redo(),
        })
    }
}

/// Co-editing state of a session.
#[derive(Default)]
pub struct Collab {
    pub rooms: Vec<Room>,
    /// Cells others changed, for the interface to highlight briefly: (document uid, sheet, cell).
    pub remote_edits: Vec<(u64, usize, CellRef)>,
    /// Where room states are kept for recovery (the app's data folder); `None`: not kept.
    pub recovery_dir: Option<std::path::PathBuf>,
    /// The main loop's clock (milliseconds), as [`Session::collab_poll`] got it.
    now: u64,
    /// Transport ids the caller serves itself, in memory ([`Session::collab_serve_in_memory`]).
    in_process: Vec<String>,
    /// Answers to their requests (`collab/snapshot`), for [`Session::collab_take_outgoing`].
    answers: Vec<Json>,
}

impl Collab {
    pub fn room_of(&self, doc: u64) -> Option<&Room> {
        self.rooms.iter().find(|r| r.doc == doc)
    }
    fn room_mut(&mut self, room: &str) -> Option<&mut Room> {
        self.rooms.iter_mut().find(|r| r.room == room)
    }
    /// Others' edits since the last call.
    pub fn take_remote_edits(&mut self) -> Vec<(u64, usize, CellRef)> {
        std::mem::take(&mut self.remote_edits)
    }
}

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            query "collab.share",
            "Start Live Session",
            ["Review", "Share"],
            None,
            "{plugin, room, user?, role?, key?, title?} → the room. Shares the active workbook in a live session through a transport plugin (protocol 2); with no params, the interface asks",
            has_doc,
            share
        ),
        cmd!(
            query "collab.join",
            "Join Live Session",
            ["Review", "Share"],
            None,
            "{plugin, room, mode: share|join, user?: {id, name, color}, role?: edit|view, key?: base64 of 32 bytes, path?, title?} → {room, clientId, state}. `share` starts a room from the active workbook (or `path`); `join` opens a new workbook (or `path`) filled from the room's document",
            always,
            join
        ),
        cmd!(query "collab.leave", "Leave Live Session", ["Review", "Share"], None, "{room?} → {left}. The workbook stays open with every edit, as an ordinary workbook", always, leave),
        cmd!(query "collab.status", "Live Sessions", [], None, "{} → {rooms: [{room, plugin, state, peers, role, encrypted, …}], transports: [{id, transport, label}]}", always, status),
        cmd!(query "collab.setUser", "Live Session Name and Colour", [], None, "{name?, color?: #RRGGBB} → the user. How others see you in live sessions", always, set_user),
    ]
}

/// Who this person is in chat messages outside a live session (and, before one, in it): the
/// name and colour of File › Options, without assigning an id if there isn't one yet.
pub fn local_user(s: &Session) -> User {
    let name: String = s.prefs.user_name.trim().chars().take(128).collect();
    User {
        id: if s.prefs.collab_user_id.is_empty() { "u-local".into() } else { s.prefs.collab_user_id.clone() },
        name: if name.is_empty() { "GridCraft User".into() } else { name },
        color: if valid_color(&s.prefs.collab_color) { s.prefs.collab_color.clone() } else { "#3B7DD8".into() },
    }
}

pub(crate) fn user_from_prefs(s: &mut Session) -> User {
    if s.prefs.collab_user_id.is_empty() {
        let mut b = [0u8; 8];
        let v = if gridcraft_collab::random_bytes(&mut b).is_ok() { u64::from_le_bytes(b) } else { gridcraft_collab::new_client_id() };
        s.prefs.collab_user_id = format!("u-{v:016x}");
    }
    let palette = ["#3B7DD8", "#D8553B", "#2E9E5B", "#9B4FD0", "#D89B1F", "#1F9BB0", "#C23B7A", "#6B7F2A"];
    if !valid_color(&s.prefs.collab_color) {
        let i = s.prefs.collab_user_id.bytes().fold(0usize, |a, b| a.wrapping_mul(31).wrapping_add(b as usize)) % palette.len();
        s.prefs.collab_color = palette.get(i).copied().unwrap_or("#3B7DD8").to_string();
    }
    let name: String = s.prefs.user_name.trim().chars().take(128).collect();
    User {
        id: s.prefs.collab_user_id.clone(),
        name: if name.is_empty() { "GridCraft User".into() } else { name },
        color: s.prefs.collab_color.clone(),
    }
}

fn share(s: &mut Session, p: &Json) -> Result<Json> {
    if p.as_object().is_none_or(|o| o.is_empty()) {
        s.ui_requests.push(UiRequest::Dialog("collabShare".into(), json!({"mode": "share"})));
        return Ok(json!({}));
    }
    let mut p = p.clone();
    p["mode"] = json!("share");
    join(s, &p)
}

fn join(s: &mut Session, p: &Json) -> Result<Json> {
    if p.as_object().is_none_or(|o| o.is_empty()) {
        s.ui_requests.push(UiRequest::Dialog("collabShare".into(), json!({"mode": "join"})));
        return Ok(json!({}));
    }
    join_with(s, p, false)
}

fn join_with(s: &mut Session, p: &Json, in_memory: bool) -> Result<Json> {
    const ID: &str = "collab.join";
    let plugin_id = match str_param(p, "plugin") {
        Some(id) => id.to_string(),
        None if in_memory => IN_MEMORY.to_string(),
        None => return Err(bad(ID, "missing `plugin` (the transport)")),
    };
    let in_memory = in_memory || plugin_id == IN_MEMORY || s.collab.in_process.contains(&plugin_id);
    let room = str_param(p, "room").ok_or_else(|| bad(ID, "missing `room`"))?.to_string();
    if room.is_empty() || room.chars().count() > MAX_ROOM || room.chars().any(char::is_control) {
        return Err(bad(ID, "`room` needs 1–256 characters"));
    }
    let mode = str_param(p, "mode").unwrap_or("join");
    if mode != "share" && mode != "join" {
        return Err(bad(ID, "`mode` is `share` or `join`"));
    }
    let view_only = match str_param(p, "role").unwrap_or("edit") {
        "edit" => false,
        "view" => true,
        _ => return Err(bad(ID, "`role` is `edit` or `view`")),
    };
    let key = match str_param(p, "key").filter(|k| !k.trim().is_empty()) {
        Some(k) => Some(Key::from_base64(k).map_err(|e| bad(ID, format!("`key`: {e}")))?),
        None => None,
    };
    let mut user = user_from_prefs(s);
    if let Some(u) = p.get("user").filter(|u| !u.is_null()) {
        user = User::parse(u).ok_or_else(|| bad(ID, "`user` needs `id` and `name` (1–128 characters) and `color` (#RRGGBB)"))?;
    }
    let title = str_param(p, "title").map(|t| t.chars().take(256).collect::<String>());
    let path = str_param(p, "path").filter(|x| !x.trim().is_empty()).map(str::to_string);

    let plugin = if in_memory {
        None
    } else {
        let plugin = gridcraft_plugin::get(&plugin_id).ok_or_else(|| bad(ID, format!("no plugin `{plugin_id}` is loaded")))?;
        if !plugin.enabled() {
            return Err(EngineError::Other(format!("plugin {plugin_id} is turned off (File › Options › Add-ins)")));
        }
        if let Some(b) = plugin.blocked_by().first() {
            return Err(EngineError::Other(format!("plugin {plugin_id} is blocked: {}", b.message)));
        }
        if !plugin.is_transport() {
            return Err(EngineError::Other(format!("plugin {plugin_id} isn't a co-editing transport (its manifest has no `collab`)")));
        }
        Some(plugin)
    };
    if s.collab.rooms.iter().any(|r| r.room == room) {
        return Err(EngineError::Other(format!("already in room {room}")));
    }
    if s.collab.rooms.iter().filter(|r| r.plugin == plugin_id).count() >= MAX_ROOMS_PER_TRANSPORT {
        return Err(EngineError::Other(format!("{plugin_id} already has {MAX_ROOMS_PER_TRANSPORT} rooms")));
    }
    if key.is_some() {
        // Encryption needs random nonces: refuse now rather than at the first edit.
        gridcraft_collab::random_bytes(&mut [0u8; 4]).map_err(|e| EngineError::Other(e.to_string()))?;
    }
    // The document: the active one (or `path`) to share; a new one (or `path`) to join.
    match (&path, mode) {
        (Some(path), _) => {
            s.execute("file.open", json!({"path": path}))?;
        }
        (None, "join") => {
            s.new_workbook();
        }
        (None, _) => {}
    }
    let doc = s.doc()?.uid;
    if s.collab.rooms.iter().any(|r| r.doc == doc) {
        return Err(EngineError::Other("this workbook is already in a live session".into()));
    }
    if let Some(plugin) = &plugin {
        plugin.ensure_running().map_err(|e| EngineError::Other(format!("can't start {plugin_id}: {e}")))?;
    }

    let client_id = gridcraft_collab::new_client_id();
    let mut replica = Replica::new(client_id);
    let mut joining = mode == "join";
    if mode == "share" {
        let d = s.doc()?;
        // The title: as given, else the file's name; an unsaved workbook has none.
        let title = title.clone().unwrap_or_else(|| if d.path.is_some() { d.display_title() } else { String::new() });
        let wb = d.wb.clone();
        replica.seed(wb, &title).map_err(|e| EngineError::Other(e.to_string()))?;
    } else if let Some(state) = read_recovery(s, &room) {
        // Our own edits from before a crash: merged with the room's when it answers.
        if replica.apply_update(&state).is_ok() {
            let wb = s.doc()?.wb.clone();
            if let Ok(Some(r)) = replica.refresh(&wb) {
                apply_refreshed(s, doc, r, None);
                joining = false;
            }
        }
    }
    let encrypted = key.is_some();
    let mut room_obj = Room {
        room: room.clone(),
        plugin: plugin_id.clone(),
        doc,
        state: RoomState::Connecting,
        joining,
        view_only,
        user: user.clone(),
        listed: vec![],
        message: None,
        refused: None,
        key,
        awareness: Awareness::new(client_id),
        replica,
        decrypt_failures: 0,
        last_awareness: None,
        last_awareness_at: 0,
        exited_at: None,
        outbox: Default::default(),
        in_memory,
        notices: vec![],
        typing: false,
    };
    // A room's undo is the person's own: the workbook's history starts here.
    if let Some(d) = s.docs.iter_mut().find(|d| d.uid == doc) {
        d.undo.clear();
        d.redo.clear();
        d.struct_log.clear();
    }
    announce(&mut room_obj).map_err(EngineError::Other)?;
    let out = json!({"room": room, "clientId": client_id, "mode": mode, "state": if room_obj.joining { "joining" } else { "connecting" }, "encrypted": encrypted});
    s.collab.rooms.push(room_obj);
    send_awareness(s, true);
    Ok(out)
}

/// `collab/joined`, then sync step 1 (and, for a shared document, our whole state).
fn announce(r: &mut Room) -> std::result::Result<(), String> {
    let joined = json!({"room": r.room, "clientId": r.client_id(), "user": r.user, "app": gridcraft_collab::APP, "schema": gridcraft_collab::SCHEMA, "encrypted": r.encrypted()});
    let sync1 = json!({"room": r.room, "stateVector": gridcraft_collab::b64_encode(&r.replica.state_vector())});
    if r.in_memory {
        notify(r, "collab/joined", joined);
        notify(r, "collab/sync1", sync1);
    } else {
        let plugin = gridcraft_plugin::get(&r.plugin).ok_or("the transport was unloaded")?;
        plugin.notify("collab/joined", joined)?;
        plugin.notify("collab/sync1", sync1)?;
    }
    Ok(())
}

fn seal(r: &Room, method: &str, bytes: &[u8]) -> std::result::Result<String, CollabError> {
    let data = match &r.key {
        Some(k) => crypto::encrypt(k, &r.room, method, bytes)?,
        None => bytes.to_vec(),
    };
    Ok(gridcraft_collab::b64_encode(&data))
}

fn unseal(r: &mut Room, method: &str, text: &str) -> std::result::Result<Vec<u8>, CollabError> {
    let data = gridcraft_collab::b64_decode(text)?;
    match &r.key {
        Some(k) => crypto::decrypt(k, &r.room, method, &data).inspect_err(|_| {
            r.decrypt_failures += 1;
            if r.decrypt_failures >= WRONG_KEY_AFTER {
                r.message = Some("Wrong room key?".into());
                if r.decrypt_failures == WRONG_KEY_AFTER {
                    r.notices.push("Wrong room key?".into());
                }
            }
        }),
        None => Ok(data),
    }
}

fn notify(r: &Room, method: &str, params: Json) {
    if r.in_memory {
        r.outbox.borrow_mut().push(json!({"jsonrpc": "2.0", "method": method, "params": params}));
        return;
    }
    if let Some(p) = gridcraft_plugin::get(&r.plugin)
        && let Err(e) = p.notify(method, params)
    {
        log::warn!("collab {}: {method}: {e}", r.room);
    }
}

fn leave(s: &mut Session, p: &Json) -> Result<Json> {
    let room = match str_param(p, "room") {
        Some(r) => r.to_string(),
        None => {
            let uid = s.doc()?.uid;
            s.collab.room_of(uid).map(|r| r.room.clone()).ok_or_else(|| EngineError::Other("this workbook isn't in a live session".into()))?
        }
    };
    if !s.collab.rooms.iter().any(|r| r.room == room) {
        return Err(EngineError::Other(format!("not in room {room}")));
    }
    leave_room(s, &room, None, true);
    Ok(json!({"left": room}))
}

/// Leaves `room`: tells the transport (unless it asked), keeps the workbook as it is.
fn leave_room(s: &mut Session, room: &str, reason: Option<&str>, tell: bool) {
    let Some(i) = s.collab.rooms.iter().position(|r| r.room == room) else { return };
    let mut r = s.collab.rooms.remove(i);
    flush(s, &mut r);
    if tell {
        let bye = r.awareness.set_local(None);
        if let Ok(u) = seal(&r, "collab/awareness", &bye) {
            notify(&r, "collab/awareness", json!({"room": r.room, "update": u}));
        }
        notify(&r, "collab/leave", json!({"room": r.room, "reason": reason.unwrap_or("left")}));
    }
    s.collab.answers.extend(std::mem::take(&mut *r.outbox.borrow_mut()));
    remove_recovery(s, room);
    if let Some(d) = s.docs.iter_mut().find(|d| d.uid == r.doc) {
        // Back to an ordinary workbook: its own undo starts afresh.
        d.undo.clear();
        d.redo.clear();
        d.struct_log.clear();
    }
    if let Some(why) = reason
        && !tell
    {
        s.ui_requests.push(UiRequest::Message(format!("The live session ended: {why}")));
    }
}

fn status(s: &mut Session, _p: &Json) -> Result<Json> {
    let rooms: Vec<Json> = s.collab.rooms.iter().map(|r| r.info(s)).collect();
    let transports: Vec<Json> = gridcraft_plugin::transports()
        .iter()
        .map(|p| {
            let c = p.manifest.collab.clone().unwrap_or_default();
            json!({"id": p.manifest.id, "transport": c.transport, "label": if c.label.is_empty() { p.manifest.name.clone() } else { c.label }})
        })
        .collect();
    Ok(json!({"rooms": rooms, "transports": transports, "user": user_from_prefs(s)}))
}

fn set_user(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "collab.setUser";
    if let Some(n) = str_param(p, "name") {
        let n = n.trim();
        if n.is_empty() || n.chars().count() > 128 {
            return Err(bad(ID, "`name` needs 1–128 characters"));
        }
        s.prefs.user_name = n.to_string();
    }
    if let Some(c) = str_param(p, "color") {
        if !valid_color(c) {
            return Err(bad(ID, "`color` is #RRGGBB"));
        }
        s.prefs.collab_color = c.to_ascii_uppercase();
    }
    let user = user_from_prefs(s);
    for r in s.collab.rooms.iter_mut() {
        r.user = User { id: r.user.id.clone(), ..user.clone() };
        r.last_awareness = None;
    }
    send_awareness(s, true);
    Ok(serde_json::to_value(&user).unwrap_or(Json::Null))
}

// ---------------------------------------------------------------- the engine's side

/// Why the active workbook can't be edited now (in a room while joining, as a viewer, or in a
/// room for another app).
pub fn read_only(s: &Session) -> Option<String> {
    let uid = s.active()?.uid;
    let r = s.collab.room_of(uid)?;
    if let Some(why) = &r.refused {
        return Some(why.clone());
    }
    if r.joining {
        return Some("Joining…".into());
    }
    r.view_only.then(|| "You can view this workbook (live session, view only)".into())
}

/// `edit.undo` and `edit.redo` in a room: the person's own changes. `None`: not in a room.
pub(crate) fn intercept(s: &mut Session, id: &str, p: &Json) -> Option<Result<Json>> {
    let undo = match id {
        "edit.undo" => true,
        "edit.redo" => false,
        _ => return None,
    };
    let uid = s.active()?.uid;
    let i = s.collab.rooms.iter().position(|r| r.doc == uid)?;
    if let Some(why) = read_only(s) {
        return Some(Err(EngineError::Disabled(id.to_string(), why)));
    }
    let steps = p.get("steps").and_then(Json::as_u64).unwrap_or(1).clamp(1, 100);
    // Anything not yet sent goes first (it becomes the step being undone).
    let mut r = s.collab.rooms.remove(i);
    flush(s, &mut r);
    let mut did = 0;
    let mut result = Ok(());
    for _ in 0..steps {
        let step = if undo { r.replica.undo() } else { r.replica.redo() };
        match step {
            Ok(Some(update)) => {
                did += 1;
                send_update(&r, "collab/update", &update);
            }
            Ok(None) => break,
            Err(e) => {
                result = Err(e);
                break;
            }
        }
    }
    let refreshed = r.replica.refresh(&s.docs.iter().find(|d| d.uid == uid).map(|d| d.wb.clone()).unwrap_or_default());
    match refreshed {
        Ok(Some(out)) => apply_refreshed(s, uid, out, Some(&mut r)),
        Ok(None) => {}
        Err(e) => result = Err(e),
    }
    s.collab.rooms.insert(i, r);
    Some(match result {
        Err(e) => Err(EngineError::Other(e.to_string())),
        Ok(()) if did == 0 => Err(EngineError::Disabled(id.to_string(), if undo { "nothing to undo".into() } else { "nothing to redo".into() })),
        Ok(()) => {
            let mut out = json!({"perUser": true});
            out[if undo { "undone" } else { "redone" }] = json!(did);
            Ok(out)
        }
    })
}

/// Whether Undo / Redo have something to do in the active workbook's room. `None`: not in one.
pub fn room_undo(s: &Session, redo: bool) -> Option<bool> {
    let uid = s.active()?.uid;
    let r = s.collab.room_of(uid)?;
    Some(r.replica.is_ready() && if redo { r.replica.can_redo() } else { r.replica.can_undo() })
}

fn send_update(r: &Room, method: &str, update: &[u8]) {
    if update.len() > gridcraft_collab::MAX_UPDATE {
        log::warn!("collab {}: an update of {} bytes is larger than 16 MiB; peers will fetch it with a sync", r.room, update.len());
        notify(r, "collab/sync1", json!({"room": r.room, "stateVector": gridcraft_collab::b64_encode(&r.replica.state_vector())}));
        return;
    }
    match seal(r, method, update) {
        Ok(u) => notify(r, method, json!({"room": r.room, "update": u})),
        Err(e) => log::warn!("collab {}: {e}", r.room),
    }
}

/// Sends the change since the last sync of `r`'s workbook, if any.
fn flush(s: &mut Session, r: &mut Room) {
    if !r.replica.is_ready() || r.refused.is_some() {
        return;
    }
    let Some(d) = s.docs.iter_mut().find(|d| d.uid == r.doc) else { return };
    let ops: Vec<StructOp> = std::mem::take(&mut d.struct_log);
    if r.replica.base().is_some_and(|b| Arc::ptr_eq(b, &d.wb)) {
        return;
    }
    if r.view_only || r.joining {
        // A viewer's messages go to the chat; their other edits never leave: back to the room's
        // state (with the chat).
        if r.view_only && !r.joining {
            let chat = d.wb.chat.clone();
            match r.replica.commit_chat(&chat) {
                Ok(Some(update)) => send_update(r, "collab/update", &update),
                Ok(None) => {}
                Err(e) => log::warn!("collab {}: {e}", r.room),
            }
        }
        if let Some(b) = r.replica.base() {
            d.wb = b.clone();
            d.calc.rebuild(&d.wb);
        }
        return;
    }
    let wb = d.wb.clone();
    let rep = &mut r.replica;
    let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| rep.commit(&wb, &ops)));
    match res {
        Ok(Ok(Some(update))) => send_update(r, "collab/update", &update),
        Ok(Ok(None)) => {}
        Ok(Err(e)) => {
            log::warn!("collab {}: {e}", r.room);
            r.message = Some(e.to_string());
        }
        Err(_) => {
            log::warn!("collab {}: a change couldn't be shared", r.room);
            r.message = Some("A change couldn't be shared".into());
        }
    }
}

/// After every top-level command: each room's workbook change becomes one update.
pub(crate) fn after_command(s: &mut Session) {
    if s.collab.rooms.is_empty() {
        for d in s.docs.iter_mut() {
            d.struct_log.clear();
        }
        return;
    }
    // Rooms whose workbook was closed end.
    let gone: Vec<String> = s.collab.rooms.iter().filter(|r| !s.docs.iter().any(|d| d.uid == r.doc)).map(|r| r.room.clone()).collect();
    for room in gone {
        leave_room(s, &room, Some("the workbook was closed"), true);
    }
    let mut rooms = std::mem::take(&mut s.collab.rooms);
    for r in rooms.iter_mut() {
        flush(s, r);
    }
    rooms.append(&mut s.collab.rooms);
    s.collab.rooms = rooms;
    for d in s.docs.iter_mut() {
        d.struct_log.clear();
    }
}

/// Puts a refreshed workbook (remote changes, an undo) into its document, recalculated.
fn apply_refreshed(s: &mut Session, uid: u64, out: gridcraft_collab::Refreshed, room: Option<&mut Room>) {
    let Some(d) = s.docs.iter_mut().find(|d| d.uid == uid) else { return };
    let mut wb = out.wb;
    let doc = d.calling_document();
    let calc = &mut d.calc;
    gridcraft_calc::external::with_document(Some(doc), || {
        if out.structural {
            calc.recalc_all(&mut wb);
        } else if !out.changed.is_empty() {
            calc.cells_changed(&mut wb, &out.changed);
        }
    });
    let wb = Arc::new(wb);
    let dirty = d.is_dirty();
    d.wb = wb.clone();
    // Others' changes make the workbook modified, their chat messages too (the chat is kept
    // only when the file is saved).
    if !dirty && out.changed.is_empty() && !out.structural && !out.chat {
        d.saved = wb.clone();
    }
    d.revision += 1;
    fix_selection(d);
    if let Some(r) = room {
        r.replica.rebase(wb);
    } else if let Some(r) = s.collab.rooms.iter_mut().find(|r| r.doc == uid) {
        r.replica.rebase(wb);
    }
    for (si, c) in out.changed {
        if s.collab.remote_edits.len() >= MAX_REMOTE_EDITS {
            s.collab.remote_edits.remove(0);
        }
        s.collab.remote_edits.push((uid, si, c));
    }
}

/// Keeps the selection on the sheet after others changed its shape.
fn fix_selection(d: &mut DocState) {
    let n = d.wb.sheets.len();
    if d.wb.active_sheet >= n {
        Arc::make_mut(&mut d.wb).active_sheet = n.saturating_sub(1);
    }
    d.sheet_selections.truncate(n);
}

impl Session {
    /// Runs co-editing between frames: takes what transports pushed, applies remote changes,
    /// answers snapshots and sends this person's selection. `now`: the caller's clock in
    /// milliseconds (the engine reads no clock: it would trap in the browser). Returns whether
    /// anything changed that the interface should show.
    pub fn collab_poll(&mut self, now: u64) -> bool {
        if self.macro_running() {
            return false;
        } // Leave pushes queued until the transaction ends.
        let pushed = gridcraft_plugin::push::take();
        self.collab_poll_pushed(now, pushed)
    }

    /// [`Session::collab_poll`] with the messages given (several sessions in one process share
    /// the plugins' queue: the caller splits it; [`IN_MEMORY`] rooms get theirs this way).
    pub fn collab_poll_pushed(&mut self, now: u64, pushed: Vec<Pushed>) -> bool {
        // A caller that already took messages cannot put them back: cancel before applying,
        // so remote/plugin changes are never overwritten by rollback or silently dropped.
        if !pushed.is_empty() {
            self.abort_macro_before_document_change();
        }
        self.collab.now = now;
        let mut changed = false;
        for p in pushed {
            changed |= handle(self, p);
        }
        // Remote changes into the workbooks.
        let mut rooms = std::mem::take(&mut self.collab.rooms);
        for r in rooms.iter_mut() {
            if r.refused.is_some() || !r.replica.has_changes() {
                continue;
            }
            flush(self, r);
            let Some(cur) = self.docs.iter().find(|d| d.uid == r.doc).map(|d| d.wb.clone()) else { continue };
            let rep = &mut r.replica;
            let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| rep.refresh(&cur)));
            match res {
                Ok(Ok(Some(out))) => {
                    let was_joining = r.joining;
                    r.joining = false;
                    let uid = r.doc;
                    apply_refreshed(self, uid, out, Some(r));
                    if was_joining {
                        if let Some(d) = self.docs.iter_mut().find(|d| d.uid == uid) {
                            d.saved = d.wb.clone();
                            if let Some(t) = r.replica.title() {
                                d.title = t;
                            }
                        }
                        r.last_awareness = None;
                    }
                    changed = true;
                }
                Ok(Ok(None)) => {}
                Ok(Err(e @ CollabError::Mismatch { .. })) => {
                    r.notices.push(e.to_string());
                    r.refused = Some(e.to_string());
                    changed = true;
                }
                Ok(Err(e)) => {
                    r.message = Some(e.to_string());
                    changed = true;
                }
                Err(_) => {
                    r.message = Some("A change from the room couldn't be applied".into());
                    changed = true;
                }
            }
            if r.replica.has_gaps() && r.state == RoomState::Online {
                notify(r, "collab/sync1", json!({"room": r.room, "stateVector": gridcraft_collab::b64_encode(&r.replica.state_vector())}));
            }
        }
        rooms.append(&mut self.collab.rooms);
        self.collab.rooms = rooms;
        // Transports that exited: started again, then the rooms rejoin.
        for i in 0..self.collab.rooms.len() {
            let Some(r) = self.collab.rooms.get_mut(i) else { continue };
            let Some(at) = r.exited_at else { continue };
            if now.saturating_sub(at) < RESTART_MS {
                continue;
            }
            r.exited_at = None;
            let Some(plugin) = gridcraft_plugin::get(&r.plugin) else { continue };
            match plugin.ensure_running() {
                Ok(()) => {
                    r.state = RoomState::Connecting;
                    r.message = None;
                    if let Err(e) = announce(r) {
                        r.message = Some(e);
                    }
                }
                Err(e) => {
                    r.message = Some(format!("{}: {e}", r.plugin));
                    if !e.starts_with("stopped after") {
                        r.exited_at = Some(now);
                    }
                }
            }
            changed = true;
        }
        let pruned = self.collab.rooms.iter_mut().fold(false, |acc, r| r.awareness.prune(now) | acc);
        send_awareness(self, false);
        // Plugin functions invalidated by a plugin (protocol 2.2), from any channel.
        let recalculated = self.recalc_invalidated();
        changed | pruned | recalculated
    }

    /// Serves transport `id` in memory: rooms joined through it (`collab.join` with `plugin: id`)
    /// work like [`Session::collab_join_in_memory`]'s. A conformance replayer stands in for a
    /// transport this way.
    pub fn collab_serve_in_memory(&mut self, id: &str) {
        if !self.collab.in_process.iter().any(|x| x == id) {
            self.collab.in_process.push(id.to_string());
        }
    }

    /// Joins a room whose transport is the caller ([`IN_MEMORY`]): `collab.join`'s params
    /// without `plugin`. Its messages are taken with [`Session::collab_take_outgoing`]; what
    /// arrives is handed to [`Session::collab_poll_pushed`] as coming from [`IN_MEMORY`].
    pub fn collab_join_in_memory(&mut self, params: Json) -> Result<Json> {
        self.depth += 1;
        let r = join_with(self, &params, true);
        self.depth = self.depth.saturating_sub(1);
        r
    }

    /// Messages for the status bar since the last call ("This room is for wordcraft 1",
    /// "Wrong room key?").
    pub fn collab_take_notices(&mut self) -> Vec<String> {
        self.collab.rooms.iter_mut().flat_map(|r| std::mem::take(&mut r.notices)).collect()
    }

    /// Messages [`IN_MEMORY`] rooms sent, oldest first (JSON-RPC notifications).
    pub fn collab_take_outgoing(&mut self) -> Vec<Json> {
        let mut out: Vec<Json> = std::mem::take(&mut self.collab.answers);
        out.extend(self.collab.rooms.iter().flat_map(|r| std::mem::take(&mut *r.outbox.borrow_mut())));
        out
    }

    /// The co-editing room of the active workbook.
    pub fn active_room(&self) -> Option<&Room> {
        self.collab.room_of(self.active()?.uid)
    }

    /// Every room's whole state, for recovery: `(room, plugin, update)`.
    pub fn collab_recovery(&self) -> Vec<(String, String, Vec<u8>)> {
        self.collab.rooms.iter().filter(|r| r.replica.is_ready()).map(|r| (r.room.clone(), r.plugin.clone(), r.replica.full_state())).collect()
    }

    /// Writes [`Session::collab_recovery`] into the recovery folder (the app does this with its
    /// AutoRecover timer).
    pub fn save_collab_recovery(&self) {
        let Some(dir) = &self.collab.recovery_dir else { return };
        for (room, plugin, state) in self.collab_recovery() {
            let body = json!({"room": room, "plugin": plugin, "update": gridcraft_collab::b64_encode(&state)});
            let path = dir.join(recovery_name(&room));
            if let Err(e) = std::fs::create_dir_all(dir).and_then(|_| std::fs::write(&path, body.to_string())) {
                log::warn!("collab: can't keep the room's state for recovery: {e}");
            }
        }
    }
}

fn recovery_name(room: &str) -> String {
    use sha2::{Digest, Sha256};
    let d = Sha256::digest(room.as_bytes());
    let hex: String = d.iter().take(12).map(|b| format!("{b:02x}")).collect();
    format!("room-{hex}.json")
}

fn read_recovery(s: &Session, room: &str) -> Option<Vec<u8>> {
    let dir = s.collab.recovery_dir.as_ref()?;
    let text = std::fs::read_to_string(dir.join(recovery_name(room))).ok()?;
    let v: Json = serde_json::from_str(&text).ok()?;
    if v.get("room").and_then(Json::as_str) != Some(room) {
        return None;
    }
    gridcraft_collab::b64_decode(v.get("update")?.as_str()?).ok()
}

fn remove_recovery(s: &Session, room: &str) {
    if let Some(dir) = &s.collab.recovery_dir {
        let _ = std::fs::remove_file(dir.join(recovery_name(room)));
    }
}

/// Sends this person's awareness now (after `typing` changed).
pub(crate) fn awareness_changed(s: &mut Session) {
    send_awareness(s, true);
}

/// This person's selection to the rooms (at most 20 a second, renewed every 15 s).
fn send_awareness(s: &mut Session, force: bool) {
    let now = s.collab.now;
    let mut sends = Vec::new();
    for (i, r) in s.collab.rooms.iter().enumerate() {
        if !r.replica.is_ready() || r.refused.is_some() {
            continue;
        }
        let Some(d) = s.docs.iter().find(|d| d.uid == r.doc) else { continue };
        let sheet = d.wb.active_sheet;
        let selection = r.replica.selection_to_ids(sheet, &d.selection.ranges);
        let cursor = r.replica.selection_to_ids(sheet, &[RangeRef::cell(d.selection.active)]).and_then(|c| {
            let first = c.get("ranges")?.get(0)?;
            Some(json!({"sheet": c.get("sheet")?, "cell": [first.get(0)?, first.get(1)?]}))
        });
        let mut state = json!({"user": r.user});
        if r.typing {
            state["typing"] = json!(true);
        }
        if let Some(sel) = selection {
            state["selection"] = sel;
        }
        if let Some(c) = cursor {
            state["cursor"] = c;
        }
        let same = r.last_awareness.as_ref() == Some(&state);
        let due = now.saturating_sub(r.last_awareness_at) >= awareness::RENEW_MS;
        let allowed = force || now.saturating_sub(r.last_awareness_at) >= AWARENESS_MS || now < r.last_awareness_at;
        if (!same || due) && allowed {
            sends.push((i, state));
        }
    }
    for (i, state) in sends {
        let Some(r) = s.collab.rooms.get_mut(i) else { continue };
        let update = r.awareness.set_local(Some(state.clone()));
        r.last_awareness = Some(state);
        r.last_awareness_at = now;
        if let Ok(u) = seal(r, "collab/awareness", &update) {
            notify(r, "collab/awareness", json!({"room": r.room, "update": u}));
        }
    }
}

/// Handles one pushed message. Returns whether the interface should repaint.
fn handle(s: &mut Session, p: Pushed) -> bool {
    let msg = match p.incoming {
        Incoming::Message(m) => m,
        Incoming::Exited(why) => {
            let mut any = false;
            let now = s.collab.now;
            for r in s.collab.rooms.iter_mut().filter(|r| r.plugin == p.plugin) {
                r.state = RoomState::Offline;
                r.message = Some(format!("Offline: {why}"));
                r.exited_at = Some(now);
                any = true;
            }
            if let Some(plugin) = gridcraft_plugin::get(&p.plugin) {
                plugin.exited();
                if any {
                    plugin.count_failure(&why);
                }
            }
            return any;
        }
        Incoming::Overlong => {
            if let Some(plugin) = gridcraft_plugin::get(&p.plugin) {
                plugin.count_failure("a message larger than 16 MiB was dropped");
            }
            return false;
        }
        Incoming::Flood => {
            if let Some(plugin) = gridcraft_plugin::get(&p.plugin) {
                plugin.fail("stopped: it sent more than 1,000 messages a second");
            }
            for r in s.collab.rooms.iter_mut().filter(|r| r.plugin == p.plugin) {
                r.state = RoomState::Offline;
                r.message = Some(format!("{} was stopped: it flooded the app", p.plugin));
            }
            return true;
        }
    };
    let method = msg.get("method").and_then(Json::as_str).unwrap_or("").to_string();
    let params = msg.get("params").cloned().unwrap_or(Json::Null);
    let id = msg.get("id").cloned();
    // Requests: answered now, in any order.
    if let Some(id) = id {
        let answer = match method.as_str() {
            "host/log" => {
                let text: String = params.get("text").and_then(Json::as_str).unwrap_or("").chars().take(4000).collect();
                log::info!("plugin {}: {text}", p.plugin);
                Ok(json!({}))
            }
            "collab/snapshot" => snapshot(s, &p.plugin, &params),
            // Protocol 2.2: recalculated at the end of this poll.
            "host/functions/invalidate" => gridcraft_plugin::invalidate_request(&p.plugin, &params),
            m if m.starts_with("host/") => Err(gridcraft_plugin::HostError {
                code: -32000,
                message: format!("{m} is only available while the plugin runs a command"),
                data: json!({"code": "NotNow"}),
            }),
            _ => Err(gridcraft_plugin::HostError::unknown()),
        };
        if p.plugin == IN_MEMORY || s.collab.in_process.contains(&p.plugin) {
            let msg = match answer {
                Ok(result) => json!({"jsonrpc": "2.0", "id": id, "result": result}),
                Err(e) => json!({"jsonrpc": "2.0", "id": id, "error": {"code": e.code, "message": e.message, "data": e.data}}),
            };
            s.collab.answers.push(msg);
        } else if let Some(plugin) = gridcraft_plugin::get(&p.plugin)
            && let Err(e) = plugin.respond(&id, answer)
        {
            log::warn!("plugin {}: answering {method}: {e}", p.plugin);
        }
        return false;
    }
    let Some(room) = params.get("room").and_then(Json::as_str).map(str::to_string) else {
        if !method.is_empty() && !method.starts_with("host/") {
            log::debug!("plugin {}: ignored notification {method}", p.plugin);
        }
        return false;
    };
    let now = s.collab.now;
    let Some(r) = s.collab.room_mut(&room) else {
        log::debug!("plugin {}: {method} for a room we're not in", p.plugin);
        return false;
    };
    if r.plugin != p.plugin {
        log::warn!("plugin {}: {method} for room {room}, which belongs to {}", p.plugin, r.plugin);
        return false;
    }
    match method.as_str() {
        "collab/sync1" => {
            let Some(sv) = params.get("stateVector").and_then(Json::as_str).and_then(|t| gridcraft_collab::b64_decode(t).ok()) else {
                return false;
            };
            if r.refused.is_some() {
                return false;
            }
            match r.replica.diff_for(&sv) {
                Ok(diff) => match seal(r, "collab/sync2", &diff) {
                    Ok(u) => {
                        let mut out = json!({"room": r.room, "update": u});
                        if let Some(to) = params.get("from").or_else(|| params.get("clientId")).filter(|c| c.is_u64()) {
                            out["to"] = to.clone();
                        }
                        notify(r, "collab/sync2", out);
                        // The asker has changes we lack: ask back.
                        if r.replica.lacks(&sv) {
                            notify(
                                r,
                                "collab/sync1",
                                json!({"room": r.room, "stateVector": gridcraft_collab::b64_encode(&r.replica.state_vector())}),
                            );
                        }
                    }
                    Err(e) => log::warn!("collab {room}: {e}"),
                },
                Err(e) => log::warn!("collab {room}: sync1: {e}"),
            }
            false
        }
        "collab/sync2" | "collab/update" => {
            if let Some(to) = params.get("to").and_then(Json::as_u64)
                && to != r.client_id()
            {
                return false;
            }
            if r.refused.is_some() {
                return false;
            }
            let Some(text) = params.get("update").and_then(Json::as_str) else { return false };
            // A stored snapshot (the room's whole state) is sealed as `collab/snapshot`.
            let sealed_as = if method == "collab/sync2" && params.get("snapshot").and_then(Json::as_bool) == Some(true) {
                "collab/snapshot"
            } else {
                method.as_str()
            };
            match unseal(r, sealed_as, text).and_then(|u| r.replica.apply_update(&u)) {
                Ok(()) => true,
                Err(e) => {
                    log::warn!("collab {room}: {method}: {e}");
                    match e {
                        CollabError::Decrypt => {}
                        CollabError::Mismatch { .. } => {
                            r.message = Some(e.to_string());
                            r.notices.push(e.to_string());
                        }
                        e => r.message = Some(e.to_string()),
                    }
                    true
                }
            }
        }
        "collab/awareness" => {
            let Some(text) = params.get("update").and_then(Json::as_str) else { return false };
            match unseal(r, "collab/awareness", text).and_then(|u| r.awareness.apply(&u, now)) {
                Ok(changed) => changed,
                Err(e) => {
                    log::debug!("collab {room}: awareness: {e}");
                    false
                }
            }
        }
        "collab/status" => {
            if let Some(st) = params.get("state").and_then(Json::as_str).and_then(RoomState::parse) {
                let back_online = st == RoomState::Online && r.state != RoomState::Online;
                r.state = st;
                if back_online {
                    notify(r, "collab/sync1", json!({"room": r.room, "stateVector": gridcraft_collab::b64_encode(&r.replica.state_vector())}));
                    r.last_awareness = None;
                    r.last_awareness_at = 0;
                }
            }
            if let Some(peers) = params.get("peers").and_then(Json::as_array) {
                r.listed = peers.iter().take(1000).filter_map(|p| p.get("user").and_then(User::parse)).collect();
            }
            r.message = params.get("message").and_then(Json::as_str).map(|m| m.chars().take(300).collect());
            if r.state == RoomState::Ended {
                let why = r.message.clone().unwrap_or_else(|| "the room was closed".into());
                leave_room(s, &room, Some(&why), false);
            }
            true
        }
        "collab/leave" => {
            let why = params.get("reason").and_then(Json::as_str).unwrap_or("the transport closed the room").to_string();
            leave_room(s, &room, Some(&why), false);
            true
        }
        _ => {
            log::debug!("plugin {}: ignored {method}", p.plugin);
            false
        }
    }
}

/// `collab/snapshot`: the room's whole state, and the workbook as `.xlsx` when asked (not in
/// an encrypted room: the transport would see the content).
fn snapshot(s: &mut Session, plugin: &str, params: &Json) -> std::result::Result<Json, gridcraft_plugin::HostError> {
    let err = |m: String| gridcraft_plugin::HostError { code: -32000, message: m, data: Json::Null };
    let room = params.get("room").and_then(Json::as_str).ok_or_else(|| err("missing `room`".into()))?.to_string();
    let i = s.collab.rooms.iter().position(|r| r.room == room && r.plugin == plugin).ok_or_else(|| err(format!("not in room {room}")))?;
    let mut r = s.collab.rooms.remove(i);
    flush(s, &mut r);
    let out = (|| {
        if !r.replica.is_ready() {
            return Err(err("the room's document hasn't arrived yet".into()));
        }
        let update = seal(&r, "collab/snapshot", &r.replica.full_state()).map_err(|e| err(e.to_string()))?;
        let mut out = json!({"update": update});
        if params.get("file").and_then(Json::as_bool) == Some(true) {
            if r.key.is_some() {
                out["fileRefused"] = json!("the room is encrypted: the document stays with the people in it");
            } else if let Some(d) = s.docs.iter().find(|d| d.uid == r.doc) {
                // `format`: the file's own format (craft-live asks for it), written when GridCraft
                // can; else an .xlsx, and `format` says so.
                let title = d.display_title();
                let asked = params.get("format").and_then(Json::as_str).map(str::to_string).unwrap_or_else(|| {
                    let ext = std::path::Path::new(&title).extension().and_then(|e| e.to_str()).unwrap_or("xlsx").to_ascii_lowercase();
                    if matches!(ext.as_str(), "xlsx" | "xlsm" | "xltx" | "xltm") { ext } else { "xlsx".into() }
                });
                let (name, format, bytes) = crate::io::live_file_bytes(&d.wb, &title, &asked).map_err(|e| err(e.to_string()))?;
                out["file"] = json!(gridcraft_collab::b64_encode(&bytes));
                out["name"] = json!(name);
                out["format"] = json!(format);
            }
        }
        Ok(out)
    })();
    s.collab.rooms.insert(i, r);
    out
}

/// A workbook's shared content, for comparing peers (tests, diagnostics).
pub fn fingerprint(wb: &Workbook) -> String {
    gridcraft_collab::fingerprint(wb)
}
