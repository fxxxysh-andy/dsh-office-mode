#!/usr/bin/env python3
"""本地 embedding 服务（第一百零三轮）。

插件按需拉起这个进程，把 /v1/embeddings 当成 OpenAI 兼容端点调用；空闲超时自己
退出（退出前不需要父进程再做什么，父进程另有 taskkill 兜底）。

只依赖 torch + transformers（用户机器上已有；不新增插件运行时依赖）。模型默认是
BAAI/bge-small-zh-v1.5：512 维、CLS pooling、L2 归一化，batch8/maxlen256/2 线程
实测峰值 RSS 约 480–515MB。

就绪约定：模型加载完、端口绑定后向 stdout 打一行
    DSH_LOCAL_EMBED_READY {"port": 12345, "dim": 512, "model": "bge-small-zh-v1.5"}
插件等这一行；等不到就判启动失败。
"""
import argparse
import json
import os
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

READY_PREFIX = 'DSH_LOCAL_EMBED_READY '
MAX_BODY_BYTES = 8 * 1024 * 1024


def log(message):
    sys.stderr.write('[local-embed] %s\n' % message)
    sys.stderr.flush()


class Embedder:
    def __init__(self, model_dir, threads, max_length, batch_size):
        import torch
        from transformers import AutoModel, AutoTokenizer

        torch.set_num_threads(max(1, int(threads)))
        torch.set_grad_enabled(False)
        self._torch = torch
        self.tokenizer = AutoTokenizer.from_pretrained(model_dir)
        self.model = AutoModel.from_pretrained(model_dir)
        self.model.eval()
        self.max_length = max(16, int(max_length))
        self.batch_size = max(1, int(batch_size))
        self.dim = int(getattr(self.model.config, 'hidden_size', 0) or 0)
        self._lock = threading.Lock()

    def embed(self, texts):
        import torch

        vectors = []
        with self._lock:
            for start in range(0, len(texts), self.batch_size):
                batch = texts[start:start + self.batch_size]
                encoded = self.tokenizer(
                    batch,
                    padding=True,
                    truncation=True,
                    max_length=self.max_length,
                    return_tensors='pt',
                )
                with torch.no_grad():
                    hidden = self.model(**encoded).last_hidden_state[:, 0]
                normalized = torch.nn.functional.normalize(hidden.float(), p=2, dim=1)
                vectors.extend(normalized.tolist())
        return vectors


class State:
    def __init__(self):
        self.last_activity = time.time()
        self.started = time.time()
        self.stopping = False

    def touch(self):
        self.last_activity = time.time()


def make_handler(embedder, state, model_name):
    class Handler(BaseHTTPRequestHandler):
        protocol_version = 'HTTP/1.1'

        def log_message(self, fmt, *args):  # 安静：错误走 stderr 就够了
            return

        def _send(self, status, payload):
            body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
            self.send_response(status)
            self.send_header('content-type', 'application/json; charset=utf-8')
            self.send_header('content-length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            if self.path in ('/health', '/v1/health'):
                state.touch()
                self._send(200, {'ok': True, 'model': model_name, 'dim': embedder.dim})
                return
            self._send(404, {'error': 'not found'})

        def do_POST(self):
            if self.path not in ('/v1/embeddings', '/embeddings'):
                self._send(404, {'error': 'not found'})
                return
            try:
                length = int(self.headers.get('content-length') or '0')
            except ValueError:
                self._send(400, {'error': 'bad content-length'})
                return
            if length <= 0 or length > MAX_BODY_BYTES:
                self._send(400, {'error': 'body too large or empty'})
                return
            try:
                payload = json.loads(self.rfile.read(length).decode('utf-8'))
            except Exception as error:  # noqa: BLE001
                self._send(400, {'error': 'bad json: %s' % error})
                return
            raw = payload.get('input')
            if isinstance(raw, str):
                texts = [raw]
            elif isinstance(raw, list) and all(isinstance(item, str) for item in raw):
                texts = raw
            else:
                self._send(400, {'error': 'input must be a string or an array of strings'})
                return
            if len(texts) == 0:
                self._send(400, {'error': 'input is empty'})
                return
            state.touch()
            try:
                started = time.time()
                vectors = embedder.embed(texts)
            except Exception as error:  # noqa: BLE001
                log('embed failed: %s' % error)
                self._send(500, {'error': 'embed failed: %s' % error})
                return
            state.touch()
            self._send(200, {
                'object': 'list',
                'data': [
                    {'object': 'embedding', 'index': index, 'embedding': vector}
                    for index, vector in enumerate(vectors)
                ],
                'model': model_name,
                'usage': {'prompt_tokens': 0, 'total_tokens': 0},
                'elapsed_ms': int((time.time() - started) * 1000),
            })

    return Handler


def idle_watcher(state, idle_seconds):
    while not state.stopping:
        time.sleep(1.0)
        if idle_seconds > 0 and time.time() - state.last_activity > idle_seconds:
            log('idle for %ss, exiting' % idle_seconds)
            sys.stdout.flush()
            os._exit(0)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--model', required=True)
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=0)
    parser.add_argument('--idle', type=int, default=120)
    parser.add_argument('--threads', type=int, default=2)
    parser.add_argument('--max-length', type=int, default=256)
    parser.add_argument('--batch-size', type=int, default=8)
    parser.add_argument('--name', default='bge-small-zh-v1.5')
    args = parser.parse_args()

    log('loading model from %s (threads=%s, max_length=%s)' % (args.model, args.threads, args.max_length))
    started = time.time()
    embedder = Embedder(args.model, args.threads, args.max_length, args.batch_size)
    log('model loaded in %.1fs, dim=%s' % (time.time() - started, embedder.dim))

    state = State()
    threading.Thread(target=idle_watcher, args=(state, args.idle), daemon=True).start()
    server = ThreadingHTTPServer((args.host, args.port), make_handler(embedder, state, args.name))
    port = server.server_address[1]
    sys.stdout.write(READY_PREFIX + json.dumps({
        'port': port,
        # 真正吃内存的是这个进程。Windows 上父进程可能是 py.exe 启动器，
        # 插件要用这里的 pid 采 RSS / 杀进程，不能用 spawn 返回的 pid。
        'pid': os.getpid(),
        'dim': embedder.dim,
        'model': args.name,
        'maxLength': args.max_length,
    }) + '\n')
    sys.stdout.flush()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        state.stopping = True
        server.shutdown()


if __name__ == '__main__':
    main()
