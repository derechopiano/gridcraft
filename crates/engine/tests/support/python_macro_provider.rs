//! Test-only Python provider for the existing #1123 macro protocol.
//! This proves the adapter boundary, not a production runner or editor.
pub const SOURCE: &str = r#"
import ast, hashlib, json, os, pathlib, subprocess, sys, tempfile, traceback

wire_in, wire_out = sys.stdin, sys.stdout
counter = 0
MAX = 16 * 1024 * 1024
worker, worker_key, module_dir, document_id = None, None, None, None
environment_key, environment_python = None, None

def receive(stream=wire_in):
    line = stream.readline(MAX + 1)
    if not line or len(line.encode('utf-8')) > MAX:
        raise RuntimeError('Missing or oversized protocol message')
    return json.loads(line)

def send(message):
    text = json.dumps(message, allow_nan=False)
    if len(text.encode('utf-8')) > MAX:
        raise ValueError('Oversized protocol message')
    wire_out.write(text + '\n')
    wire_out.flush()

def host(method, params):
    global counter
    counter += 1
    key = 'python:' + str(counter)
    send({'jsonrpc': '2.0', 'id': key, 'method': method, 'params': params})
    reply = receive()
    if reply.get('id') != key:
        raise RuntimeError('Unexpected host response identity')
    if 'error' in reply:
        raise RuntimeError(reply['error']['message'])
    return reply['result']

def project():
    global document_id
    result = host('host/pythonProject/get', {})
    document_id = result['documentId']
    return result['project'] or {
        'metadata': '# /// script\n# requires-python = ">=3.11"\n# dependencies = []\n# ///\n',
        'modules': {},
    }

def close_worker():
    global worker, worker_key, module_dir
    if worker is not None:
        try:
            worker.stdin.write('{"action":"shutdown"}\n')
            worker.stdin.flush()
            worker.wait(timeout=1)
        except Exception:
            worker.kill()
            worker.wait()
        worker.stdin.close()
        worker.stdout.close()
        worker = None
    if module_dir is not None:
        module_dir.cleanup()
        module_dir = None
    worker_key = None

def get_worker(p):
    global worker, worker_key, module_dir, environment_key, environment_python
    key = (document_id, hashlib.sha256(json.dumps(p,sort_keys=True).encode()).hexdigest())
    if worker is not None and worker_key == key and worker.poll() is None:
        return worker, pathlib.Path(module_dir.name)
    close_worker()
    folder = pathlib.Path(sys.argv[4])
    module_dir = tempfile.TemporaryDirectory(prefix='modules-', dir=folder)
    modules = pathlib.Path(module_dir.name)
    for name, source in p['modules'].items():
        (modules / (name+'.py')).write_text(source, encoding='utf-8')
    script = folder / 'entry.py'
    bootstrap = pathlib.Path(sys.argv[1]).read_text(encoding='utf-8')
    script.write_text(p['metadata']+'\n'+bootstrap, encoding='utf-8')
    env_key = hashlib.sha256((p['metadata']+bootstrap).encode()).hexdigest()
    # Resolve with uv once per environment; source changes need a fresh interpreter,
    # not another environment-resolution subprocess. Never reuse across metadata changes.
    if environment_key == env_key and environment_python and pathlib.Path(environment_python).is_file():
        command = [environment_python, str(script)]
    else:
        command = [sys.argv[2], 'run', '--offline', '--no-python-downloads', '--no-config',
            '--no-project', '--no-env-file', '--no-build', '--python', sys.argv[3], '--script', str(script), '--prepare']
    env = os.environ.copy()
    env.pop('PYTHONPATH', None)
    env.update(PYTHONUTF8='1', PYTHONUNBUFFERED='1')
    worker = subprocess.Popen(command,
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
        text=True, encoding='utf-8', cwd=folder, env=env)
    ready = receive(worker.stdout)
    if ready.get('kind') != 'ready' or not isinstance(ready.get('interpreter'),str):
        close_worker()
        raise RuntimeError('Worker did not complete its startup handshake')
    environment_key, environment_python = env_key, ready['interpreter']
    if command[-1] == '--prepare':
        # uv prepares the environment, then exits. Keep only the interpreter alive.
        if worker.wait(timeout=5) != 0:
            close_worker()
            raise RuntimeError('uv environment preparation failed')
        worker.stdin.close()
        worker.stdout.close()
        worker = subprocess.Popen([environment_python, str(script)], stdin=subprocess.PIPE,
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, encoding='utf-8', cwd=folder, env=env)
        ready = receive(worker.stdout)
        if ready.get('kind') != 'ready':
            close_worker()
            raise RuntimeError('Prepared interpreter failed its startup handshake')
    worker_key = key
    return worker, modules

def handle(method, params):
    if method == 'initialize':
        features = params['host'].get('features', [])
        if not all(f in features for f in ('pythonProjects', 'engineCommands')):
            raise RuntimeError('This Python provider requires the Python-project host extension')
        return {'protocolVersion': 1}
    p = project()
    if method == 'macros/modules':
        return {'modules': [{'name': n, 'kind': 'standard', 'source': s} for n,s in p['modules'].items()]}
    if method == 'macros/list':
        macros = []
        for name, source in p['modules'].items():
            for node in ast.parse(source, filename=name+'.py').body:
                if isinstance(node, ast.FunctionDef) and not node.name.startswith('_'):
                    a = node.args
                    if not (a.posonlyargs or a.args or a.kwonlyargs or a.vararg or a.kwarg):
                        macros.append({'name': name+'.'+node.name, 'module': name, 'procedure': node.name})
        return {'macros': macros}
    if method == 'macros/saveModule':
        p['modules'][params['name']] = params['source']
        return host('host/pythonProject/set', {'project': p})
    if method == 'macros/deleteModule':
        p['modules'].pop(params['name'], None)
        return host('host/pythonProject/set', {'project': p})
    if method == 'macros/run':
        # Keep a warm worker for this workbook and exact project; never share module state between books.
        child, modules = get_worker(p)
        try:
            child.stdin.write(json.dumps({'action':'run', 'macro':params['name'], 'modulePath':str(modules)})+'\n')
            child.stdin.flush()
            while True:
                message = receive(child.stdout)
                if message['kind'] == 'command':
                    result = host('host/engine.execute', {'command':message['command'], 'params':message['params']})
                    child.stdin.write(json.dumps({'result':result})+'\n')
                    child.stdin.flush()
                elif message['kind'] == 'output':
                    host('host/log', {'text':message.get('text','')})
                elif message['kind'] == 'failed':
                    raise RuntimeError(message['error'])
                elif message['kind'] == 'done':
                    return {}
                else:
                    raise RuntimeError('Unexpected worker message')
        except BaseException:
            close_worker()
            raise
    raise ValueError('Unsupported method: '+method)

while True:
    try:
        request = receive()
    except Exception:
        break
    if request.get('method') == 'shutdown':
        close_worker()
        break
    reply = {'jsonrpc':'2.0', 'id':request['id']}
    try:
        reply['result'] = handle(request['method'], request.get('params', {}))
    except Exception:
        reply['error'] = {'code':-32000, 'message':traceback.format_exc(limit=8)[-8192:]}
    send(reply)
close_worker()
"#;
