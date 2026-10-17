//! Original Python bridge data for the user's explicitly trusted interpreter.
//! Private pipes carry commands to a transaction session, never a network listener.
pub(super) const SOURCE: &str = r#"
import ast as _ast
import contextlib as _contextlib
import importlib as _importlib
import inspect as _inspect
import json as _json
import sys as _sys
import traceback as _traceback
import types as _types

_wire_in, _wire_out = _sys.stdin, _sys.stdout
_MAX = 16 * 1024 * 1024

def _send(message):
    data = _json.dumps(message, allow_nan=False, separators=(',', ':'))
    if len(data.encode('utf-8')) > _MAX:
        raise ValueError('Macro message exceeds 16 MiB')
    _wire_out.write(data + '\n')
    _wire_out.flush()

def _receive():
    data = _wire_in.buffer.readline(_MAX + 1)
    if not data or len(data) > _MAX:
        raise ValueError('Missing or oversized host response')
    return _json.loads(data)

def _call(command, params):
    _send({'kind': 'command', 'command': command, 'params': params})
    result = _receive()
    if 'error' in result:
        raise RuntimeError(result['error'])
    return result['result']

class _Workbook:
    def read(self, reference, formulas=False):
        return _call('sheet.read', {'range': reference, 'formulas': formulas})['values']
    def cell(self, reference):
        return _call('cell.get', {'cell': reference})
    def inspect(self):
        return _call('document.inspect', {})
    def write(self, reference, values):
        return _call('range.setValues', {'range': reference, 'values': values, 'literal': True})
    def execute(self, command, **params):
        return _call(command, params)

_bridge = _types.ModuleType('gridcraft')
_bridge.active_workbook = lambda: _Workbook()
_sys.modules['gridcraft'] = _bridge

class _Output:
    def __init__(self, name):
        self.name, self.remaining = name, 65536
    def write(self, text):
        if self.remaining:
            part = text[:self.remaining]
            self.remaining -= len(part)
            _send({'kind': 'output', 'stream': self.name, 'text': part})
        return len(text)
    def flush(self):
        pass

_send({'kind': 'ready', 'interpreter': _sys.executable})
if _sys.argv[-1] == '--prepare':
    _sys.exit(0)
while True:
    try:
        _request = _receive()
    except Exception:
        break
    if _request.get('action') == 'shutdown':
        break
    try:
        if _request['modulePath'] not in _sys.path:
            _sys.path.insert(0, _request['modulePath'])
        with _contextlib.redirect_stdout(_Output('stdout')), _contextlib.redirect_stderr(_Output('stderr')):
            _module, _function = _request['macro'].split('.')
            _loaded = _importlib.import_module(_module)
            _entry = getattr(_loaded, _function)
            if not _inspect.isfunction(_entry) or _entry.__module__ != _loaded.__name__ or _function.startswith('_'):
                raise ValueError('Select a public top-level function defined in the module')
            if _inspect.signature(_entry).parameters or _inspect.isgeneratorfunction(_entry):
                raise TypeError('Select a synchronous function with no arguments and no yield')
            _return = _entry()
            if _inspect.isawaitable(_return):
                if _inspect.iscoroutine(_return):
                    _return.close()
                raise TypeError('Async macros are not supported')
        _send({'kind': 'done'})
    except BaseException:
        _send({'kind': 'failed', 'error': _traceback.format_exc(limit=12)[-8192:]})
"#;
