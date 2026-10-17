//! Native background macro transport. The engine/UI retains ownership of the Session.
//! This is transport only: callers must bracket the run with the shared macro transaction,
//! pin the original document and reject/defer conflicting edits before applying host requests.

use std::sync::Arc;
use std::sync::atomic::Ordering;
use std::sync::mpsc::{Receiver, RecvTimeoutError, SyncSender, TryRecvError, sync_channel};
use std::time::Duration;

use serde_json::Value as Json;

use crate::{COMMAND_TIMEOUT, HostError, Plugin, RequestControl};

/// One host request. Answer it on the Session's owning thread, never the worker thread.
pub struct HostRequest {
    pub method: String,
    pub params: Json,
    reply: SyncSender<Result<Json, HostError>>,
}

impl HostRequest {
    /// Nonblocking: the macro may have been cancelled while the host was working.
    pub fn answer(self, answer: Result<Json, HostError>) -> Result<(), String> {
        self.reply.try_send(answer).map_err(|_| "the macro stopped waiting for its host reply".into())
    }
}

/// One in-flight request, with bounded queues. Dropping it requests cancellation; it never
/// joins a thread on the UI thread. A Stop does not increment the provider's failure count.
#[must_use = "poll the macro and answer its host requests, or drop it to cancel"]
pub struct MacroCall {
    requests: Receiver<HostRequest>,
    result: Receiver<Result<Json, String>>,
    control: RequestControl,
    completed: bool,
}

struct RunLease(Arc<Plugin>);
impl Drop for RunLease {
    fn drop(&mut self) {
        self.0.macro_in_flight.store(false, Ordering::SeqCst);
    }
}

impl MacroCall {
    /// Starts only macros/run through the existing host connection. One run per provider.
    pub fn start(plugin_id: &str, params: Json, budget: Duration) -> Result<Self, String> {
        let plugin = crate::get(plugin_id).ok_or_else(|| format!("no plugin `{plugin_id}` is loaded"))?;
        if !plugin.enabled() || plugin.manifest.macros.is_none() {
            return Err(format!("plugin {plugin_id} is not an enabled macro provider"));
        }
        plugin.check_blocked()?;
        plugin
            .macro_in_flight
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .map_err(|_| format!("plugin {plugin_id} is busy"))?;
        let lease = RunLease(plugin);
        let control = RequestControl::new(budget);
        let worker_control = control.clone();
        let (to_host, requests) = sync_channel(1);
        let (finished, result) = sync_channel(1);
        std::thread::Builder::new()
            .name(format!("macro-{plugin_id}"))
            .spawn(move || {
                let answer = lease.0.request_controlled(
                    "macros/run",
                    params,
                    COMMAND_TIMEOUT,
                    &mut |method, params| {
                        let failure = |message| HostError { code: -32000, message, data: Json::Null };
                        worker_control.remaining().map_err(|e| failure(e.message))?;
                        let (reply, wait) = sync_channel(1);
                        to_host
                            .try_send(HostRequest { method: method.into(), params: params.clone(), reply })
                            .map_err(|_| failure("the macro host is unavailable".into()))?;
                        loop {
                            let left = worker_control.remaining().map_err(|e| failure(e.message))?;
                            match wait.recv_timeout(left.min(Duration::from_millis(20))) {
                                Ok(answer) => return answer,
                                Err(RecvTimeoutError::Timeout) => continue,
                                Err(RecvTimeoutError::Disconnected) => return Err(failure("the macro host dropped its request".into())),
                            }
                        }
                    },
                    Some(&worker_control),
                );
                // Release before publishing completion, so the next run cannot spuriously be busy.
                drop(lease);
                let _ = finished.try_send(answer);
            })
            .map_err(|e| format!("could not start the macro request thread: {e}"))?;
        Ok(Self { requests, result, control, completed: false })
    }

    pub fn cancel(&self) {
        self.control.cancel();
    }

    pub fn next_request(&self) -> Option<HostRequest> {
        if self.control.is_cancelled() || self.completed { None } else { self.requests.try_recv().ok() }
    }

    /// One final result. A terminated thread is a failure, never a successful run.
    pub fn poll_result(&mut self) -> Option<Result<Json, String>> {
        if self.completed {
            return None;
        }
        let result = match self.result.try_recv() {
            Ok(result) => result,
            Err(TryRecvError::Empty) => return None,
            Err(TryRecvError::Disconnected) => Err("the macro request thread stopped unexpectedly".into()),
        };
        self.completed = true;
        Some(result)
    }
}

impl Drop for MacroCall {
    fn drop(&mut self) {
        self.control.cancel();
    }
}
