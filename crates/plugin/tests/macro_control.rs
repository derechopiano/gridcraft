//! Transport regressions against this executable as a real child process.
#![allow(clippy::disallowed_methods)] // Native process tests; never run in the web app.
use std::io::{BufRead, Write};
use std::time::{Duration, Instant};

use gridcraft_plugin::{Connection, HostError, Manifest, RequestControl};
use serde_json::{Value as Json, json};

fn main() {
    if let Some(mode) = std::env::args().nth(1) {
        serve(&mode);
        return;
    }
    silent_run_is_cancellable();
    initialization_is_cancellable();
    host_traffic_cannot_extend_an_absolute_budget();
    ordinary_requests_keep_their_activity_timeout();
    cancellation_during_a_host_call_refuses_its_reply();
    background_host_requests_are_nonblocking_and_stop_does_not_disable_the_plugin();
    dropping_a_background_job_cancels_its_request();
    background_initialization_obeys_the_budget();
    println!("Macro transport: eight real-process cancellation/deadline/background checks passed");
}

fn connection(mode: &str) -> Connection {
    let manifest = Manifest::parse(
        &json!({
            "id":"test.macro-control", "name":"Macro control fixture",
            "run":{"command":std::env::current_exe().unwrap(), "args":[mode]}
        })
        .to_string(),
    )
    .unwrap();
    Connection::start(&manifest, &std::env::temp_dir()).unwrap()
}

fn ready(mode: &str) -> Connection {
    let mut conn = connection(mode);
    conn.request("initialize", json!({}), Duration::from_secs(5)).unwrap();
    conn
}

fn cancel_soon(control: &RequestControl) -> std::thread::JoinHandle<()> {
    let control = control.clone();
    std::thread::spawn(move || {
        std::thread::sleep(Duration::from_millis(50));
        control.cancel();
    })
}

fn silent_run_is_cancellable() {
    let mut conn = ready("silent");
    let control = RequestControl::new(Duration::from_secs(10));
    let stop = cancel_soon(&control);
    let start = Instant::now();
    let error =
        conn.request_with_control("macros/run", json!({}), Duration::from_secs(10), &mut |_, _| Err(HostError::unknown()), &control).unwrap_err();
    assert!(error.fatal && error.message.contains("cancelled"), "{error:?}");
    assert!(start.elapsed() < Duration::from_secs(2));
    stop.join().unwrap();
    drop(conn); // Fatal replies must never reuse a connection with an answer in flight.
    let mut restarted = ready("done");
    assert_eq!(restarted.request("macros/run", json!({}), Duration::from_secs(5)).unwrap(), json!({"ran":true}));
}

fn initialization_is_cancellable() {
    let mut conn = connection("silent-init");
    let control = RequestControl::new(Duration::from_secs(10));
    let stop = cancel_soon(&control);
    let start = Instant::now();
    let error =
        conn.request_with_control("initialize", json!({}), Duration::from_secs(10), &mut |_, _| Err(HostError::unknown()), &control).unwrap_err();
    assert!(error.fatal && error.message.contains("cancelled"));
    assert!(start.elapsed() < Duration::from_secs(2));
    stop.join().unwrap();
}

fn host_traffic_cannot_extend_an_absolute_budget() {
    let mut conn = ready("chatty");
    let control = RequestControl::new(Duration::from_millis(120));
    let start = Instant::now();
    let mut calls = 0;
    let error = conn
        .request_with_control(
            "macros/run",
            json!({}),
            Duration::from_secs(5),
            &mut |_, _| {
                calls += 1;
                Ok(Json::Null)
            },
            &control,
        )
        .unwrap_err();
    assert!(error.fatal && error.message.contains("absolute time limit"), "{error:?}");
    assert!(calls > 0);
    assert!(start.elapsed() < Duration::from_secs(2));
}

fn ordinary_requests_keep_their_activity_timeout() {
    let mut conn = ready("chatty");
    let mut calls = 0;
    let result = conn
        .request_with("macros/run", json!({}), Duration::from_millis(200), &mut |_, _| {
            calls += 1;
            Ok(Json::Null)
        })
        .unwrap();
    assert_eq!(result, json!({"ran":true}));
    assert_eq!(calls, 20); // About 300 ms overall; each notification renews the ordinary timeout.
}

fn cancellation_during_a_host_call_refuses_its_reply() {
    let mut conn = ready("host-call");
    let control = RequestControl::new(Duration::from_secs(5));
    let error = conn
        .request_with_control(
            "macros/run",
            json!({}),
            Duration::from_secs(5),
            &mut |method, _| {
                assert_eq!(method, "host/probe");
                control.cancel();
                Ok(Json::Null)
            },
            &control,
        )
        .unwrap_err();
    assert!(error.fatal && error.message.contains("cancelled"));
}

fn install_provider(mode: &str) -> std::sync::Arc<gridcraft_plugin::Plugin> {
    let manifest = Manifest::parse(
        &json!({
            "id":"test.background-macro", "name":"Background macro fixture",
            "macros":{"language":"Test"},
            "run":{"command":std::env::current_exe().unwrap(), "args":[mode]}
        })
        .to_string(),
    )
    .unwrap();
    gridcraft_plugin::install(manifest, std::env::temp_dir()).unwrap()
}

fn host_request(call: &mut gridcraft_plugin::macro_call::MacroCall) -> gridcraft_plugin::macro_call::HostRequest {
    let deadline = Instant::now() + Duration::from_secs(5);
    loop {
        if let Some(request) = call.next_request() {
            return request;
        }
        if let Some(result) = call.poll_result() {
            panic!("macro ended before its host request: {result:?}");
        }
        assert!(Instant::now() < deadline, "host request did not arrive");
        std::thread::sleep(Duration::from_millis(5));
    }
}

fn result(call: &mut gridcraft_plugin::macro_call::MacroCall) -> Result<Json, String> {
    let deadline = Instant::now() + Duration::from_secs(5);
    loop {
        if let Some(result) = call.poll_result() {
            return result;
        }
        assert!(Instant::now() < deadline, "background macro did not finish");
        std::thread::sleep(Duration::from_millis(5));
    }
}

fn background_host_requests_are_nonblocking_and_stop_does_not_disable_the_plugin() {
    use gridcraft_plugin::macro_call::MacroCall;
    let provider = install_provider("host-call");
    for _ in 0..4 {
        let mut call = MacroCall::start("test.background-macro", json!({}), Duration::from_secs(5)).unwrap();
        let request = host_request(&mut call);
        assert_eq!(request.method, "host/probe");
        let start = Instant::now();
        assert!(call.poll_result().is_none());
        assert_eq!(provider.status(), "running", "status must not take a busy request's state lock");
        assert_eq!(provider.state(), "running");
        assert!(MacroCall::start("test.background-macro", json!({}), Duration::from_secs(5)).is_err());
        let error = gridcraft_plugin::call_macros("test.background-macro", "macros/list", json!({}), &mut |_, _| Ok(Json::Null)).unwrap_err();
        assert!(error.contains("busy"));
        assert!(start.elapsed() < Duration::from_secs(2), "owner thread waited for the provider mutex");
        call.cancel();
        assert!(result(&mut call).unwrap_err().contains("cancelled"));
        assert!(request.answer(Ok(Json::Null)).is_err(), "late answer must not reach a later run");
        assert!(!provider.failed(), "ordinary Stop must not consume the provider failure budget");
    }
    let mut call = MacroCall::start("test.background-macro", json!({}), Duration::from_secs(5)).unwrap();
    host_request(&mut call).answer(Ok(json!({"ok":true}))).unwrap();
    assert_eq!(result(&mut call).unwrap(), json!({"ran":true}));
    assert!(call.poll_result().is_none(), "completion is delivered once");
    drop(call);
    assert!(gridcraft_plugin::unload("test.background-macro"));
}

fn dropping_a_background_job_cancels_its_request() {
    use gridcraft_plugin::macro_call::MacroCall;
    let _provider = install_provider("host-call");
    let mut call = MacroCall::start("test.background-macro", json!({}), Duration::from_secs(5)).unwrap();
    let request = host_request(&mut call);
    let start = Instant::now();
    drop(call);
    assert!(start.elapsed() < Duration::from_secs(2), "drop must not join the worker thread");
    let deadline = Instant::now() + Duration::from_secs(5);
    let mut next = loop {
        match MacroCall::start("test.background-macro", json!({}), Duration::from_secs(5)) {
            Ok(next) => break next,
            Err(error) => assert!(error.contains("busy"), "{error}"),
        }
        assert!(Instant::now() < deadline, "dropped job retained its provider lease");
        std::thread::sleep(Duration::from_millis(5));
    };
    assert!(request.answer(Ok(Json::Null)).is_err());
    host_request(&mut next).answer(Ok(Json::Null)).unwrap();
    assert!(result(&mut next).is_ok());
    drop(next);
    assert!(gridcraft_plugin::unload("test.background-macro"));
}

fn background_initialization_obeys_the_budget() {
    use gridcraft_plugin::macro_call::MacroCall;
    let _provider = install_provider("silent-init");
    let mut call = MacroCall::start("test.background-macro", json!({}), Duration::from_millis(100)).unwrap();
    let start = Instant::now();
    assert!(result(&mut call).unwrap_err().contains("absolute time limit"));
    assert!(start.elapsed() < Duration::from_secs(2));
    drop(call);
    assert!(gridcraft_plugin::unload("test.background-macro"));
}

fn serve(mode: &str) {
    let mut stdin = std::io::stdin().lock().lines();
    let mut stdout = std::io::stdout().lock();
    while let Some(Ok(line)) = stdin.next() {
        let Ok(msg) = serde_json::from_str::<Json>(&line) else { return };
        let method = msg.get("method").and_then(Json::as_str).unwrap_or("");
        let id = msg.get("id").cloned().unwrap_or(Json::Null);
        if (method == "initialize" && mode == "silent-init") || (method == "macros/run" && mode == "silent") {
            std::thread::sleep(Duration::from_secs(60));
            return;
        }
        if method == "macros/run" && mode == "chatty" {
            for _ in 0..20 {
                let notice = json!({"jsonrpc":"2.0", "method":"host/log", "params":{"text":"working"}});
                if writeln!(stdout, "{notice}").and_then(|_| stdout.flush()).is_err() {
                    return;
                }
                std::thread::sleep(Duration::from_millis(15));
            }
        }
        if method == "macros/run" && mode == "host-call" {
            let call = json!({"jsonrpc":"2.0", "id":"probe", "method":"host/probe", "params":{}});
            if writeln!(stdout, "{call}").and_then(|_| stdout.flush()).is_err() {
                return;
            }
            if stdin.next().is_none() {
                return;
            }
        }
        let result = if method == "initialize" { json!({"protocolVersion":1}) } else { json!({"ran":true}) };
        let reply = json!({"jsonrpc":"2.0", "id":id, "result":result});
        if writeln!(stdout, "{reply}").and_then(|_| stdout.flush()).is_err() {
            return;
        }
    }
}
