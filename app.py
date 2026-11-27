"""考试宝题库爬取 - Web服务端（Chrome CDP模式）"""

import asyncio
import json
import re
import sys
import io
import threading
import time
import uuid
import os
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file

from scraper_core import (
    OUTPUT_DIR, DATA_DIR, CDP_URL, collect_questions,
    save_to_excel, load_progress, save_progress,
    validate_paper_id, ensure_chrome_debug, get_chrome_status, open_cdp_page,
    reopen_kaoshibao, load_segment_progress, save_segment_progress, merge_segments,
    save_to_csv, save_to_json, export_questions
)

from logging_config import init_default_logger, get_logger

if not isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# 初始化日志
init_default_logger(level="INFO")
logger = get_logger("kaoshibao.app")

app = Flask(__name__, template_folder=str(Path(__file__).parent / 'templates'))

# 安全配置：从环境变量读取 SECRET_KEY，未设置则生成临时值并警告
app.secret_key = os.environ.get('FLASK_SECRET_KEY', '')
if not app.secret_key:
    app.secret_key = os.urandom(32).hex()
    logger.warning("FLASK_SECRET_KEY 未设置，使用临时随机密钥（重启后 session 失效）")

# 请求体大小上限（1MB，足够 JSON 请求，防止超大请求耗尽内存）
app.config['MAX_CONTENT_LENGTH'] = 1 * 1024 * 1024

# 可选认证token（设置AUTH_TOKEN环境变量后启用）
AUTH_TOKEN = os.environ.get('AUTH_TOKEN', '')


@app.after_request
def set_security_headers(resp):
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    resp.headers['X-Frame-Options'] = 'SAMEORIGIN'
    resp.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    resp.headers['X-XSS-Protection'] = '1; mode=block'
    resp.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline'; "
        "style-src 'self' 'unsafe-inline'; "
        "img-src 'self' data:; "
        "font-src 'self' data:; "
        "connect-src 'self'; "
        "frame-ancestors 'self'"
    )
    return resp


@app.before_request
def auth_check():
    """可选token认证：设置AUTH_TOKEN环境变量后启用"""
    if not AUTH_TOKEN:
        return None
    if request.path in ('/', '/health', '/manifest.json', '/api/docs'):
        return None
    token = request.headers.get('X-Auth-Token', '') or request.args.get('token', '')
    if token != AUTH_TOKEN:
        return jsonify({'error': '未授权'}), 401


@app.before_request
def csrf_origin_check():
    """CSRF防护：验证POST请求的Origin头（OWASP推荐方案）"""
    if request.method == 'POST':
        origin = request.headers.get('Origin')
        if origin:
            from urllib.parse import urlparse
            origin_host = urlparse(origin).netloc
            request_host = request.headers.get('Host', '')
            if origin_host and request_host and origin_host != request_host:
                if ':' not in request_host:
                    origin_port = urlparse(origin).port
                    default_port = 443 if urlparse(origin).scheme == 'https' else 80
                    if origin_port and origin_port != default_port:
                        return jsonify({'error': '跨站请求被拒绝'}), 403
                else:
                    return jsonify({'error': '跨站请求被拒绝'}), 403


def _is_api_request():
    return request.path.startswith('/api/') or request.path.startswith('/health')


@app.errorhandler(404)
def not_found(e):
    if _is_api_request():
        return jsonify({'error': '资源不存在', 'code': 404}), 404
    return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>404 - 页面不存在</title><style>body{font-family:sans-serif;text-align:center;padding:50px;color:#666}h1{font-size:48px;color:#4a90d9}a{color:#4a90d9;text-decoration:none}</style></head><body><h1>404</h1><p>页面不存在</p><a href="/">返回首页</a></body></html>', 404


@app.errorhandler(413)
def too_large(e):
    if _is_api_request():
        return jsonify({'error': '请求体过大，超过 1MB 限制', 'code': 413}), 413
    return jsonify({'error': '请求体过大'}), 413


@app.errorhandler(500)
def server_error(e):
    logger.exception(f"服务器内部错误: {e}")
    if _is_api_request():
        return jsonify({'error': '服务器内部错误', 'code': 500}), 500
    return jsonify({'error': '服务器内部错误'}), 500


@app.errorhandler(405)
def method_not_allowed(e):
    if _is_api_request():
        return jsonify({'error': '请求方法不允许', 'code': 405}), 405
    return jsonify({'error': '请求方法不允许'}), 405

MAX_TASKS = 5
_start_time = time.time()

tasks = {}
tasks_lock = threading.Lock()

# 全局Chrome操作锁：串行化所有任务的Chrome CDP操作，避免多任务并发连接/翻页互相干扰
_chrome_lock = threading.Lock()

# 速率限制
_rate_limit = {}
_rate_limit_lock = threading.Lock()
_rate_limit_file = DATA_DIR / 'rate_limit.json'


def _load_rate_limit():
    """启动时从文件加载速率限制数据"""
    global _rate_limit
    try:
        if _rate_limit_file.exists():
            with open(_rate_limit_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            now = time.time()
            with _rate_limit_lock:
                _rate_limit = {k: v for k, v in data.items() if now - v < 3600}
    except Exception:
        pass


def _save_rate_limit():
    """退出时保存速率限制到文件"""
    try:
        with _rate_limit_lock:
            data = dict(_rate_limit)
        with open(_rate_limit_file, 'w', encoding='utf-8') as f:
            json.dump(data, f)
    except Exception:
        pass


_load_rate_limit()
import atexit
atexit.register(_save_rate_limit)


def new_task(paper_id, segment_id=-1, start_sequence=0, max_questions=0, mode='ocr'):
    return {
        'task_id': '',
        'paper_id': paper_id,
        'segment_id': segment_id,
        'start_sequence': start_sequence,
        'max_questions': max_questions,
        'mode': mode,
        'running': True,
        'phase': 'init',
        'progress': 0,
        'total': 0,
        'has_answer': 0,
        'message': '启动中...',
        'log': [],
        'stop_requested': False,
        'created_at': time.time(),
        'output_file': '',
        'started_at': 0,
        'timeout': 3600,
        'timed_out': False,
    }


def task_log(task, msg):
    timestamp = time.strftime('%H:%M:%S')
    entry = f"[{timestamp}] {msg}"
    task['log'].append(entry)
    if len(task['log']) > 200:
        task['log'] = task['log'][-200:]
    logger.info(f"[{task['task_id'][:8]}] {msg}")


def check_rate_limit(ip, min_interval=1.0):
    now = time.time()
    with _rate_limit_lock:
        if len(_rate_limit) > 1000:
            cutoff = now - 3600
            expired = [k for k, v in _rate_limit.items() if v < cutoff]
            for k in expired:
                del _rate_limit[k]
        last = _rate_limit.get(ip, 0)
        if now - last < min_interval:
            return False
        _rate_limit[ip] = now
    return True


def get_request_json():
    if not request.is_json:
        return None, '请求格式错误，请使用JSON'
    data = request.json
    if data is None:
        return None, '无效的JSON数据'
    return data, None


def _merge_questions(result: list, questions: list) -> int:
    """合并新收集的题目到已有列表，返回新增数量"""
    existing_keys = {q.get('题目', '')[:80] for q in questions}
    added = 0
    for q in result:
        key = q.get('题目', '')[:80]
        if key and key not in existing_keys:
            q['序号'] = len(questions) + 1
            questions.append(q)
            existing_keys.add(key)
            added += 1
    return added


def _finalize_task_success(task, questions, is_segment, paper_id, segment_id, mode_suffix, output_file):
    """任务成功完成后的保存和状态更新"""
    if is_segment:
        save_segment_progress(paper_id, segment_id, questions)
    elif questions:
        save_progress(paper_id, questions, suffix=mode_suffix)
        save_to_excel(questions, output_file)
    try:
        from db_store import save_questions_to_db, is_db_configured
        if is_db_configured():
            save_questions_to_db(paper_id, questions, task.get('mode', 'ocr'))
    except Exception as e:
        logger.warning(f"数据库保存失败(不影响文件存储): {e}")
    task['phase'] = 'done'
    task['has_answer'] = sum(1 for q in questions if q.get('正确答案'))
    task['total'] = len(questions)
    task['progress'] = len(questions)
    task_log(task, f"完成! 总 {len(questions)} 道, 有答案 {task['has_answer']} 道")
    task['message'] = f"完成! {len(questions)}题, {task['has_answer']}道有答案"


async def _do_scrape(task, paper_id, segment_id, is_segment, start_sequence, max_questions, mode, questions):
    """执行异步爬取，返回 (questions, status)，status: ok/auth_error/error"""
    from playwright.async_api import async_playwright
    label = f"分段{segment_id}" if is_segment else ""
    task_log(task, f"连接Chrome{'('+label+')' if is_segment else ''}...")
    async with async_playwright() as p:
        browser, page = await open_cdp_page(p)
        if not browser:
            task_log(task, "无法连接Chrome，请确保Chrome已启动并登录考试宝")
            task['phase'] = 'error'
            task['message'] = '无法连接Chrome'
            return questions, 'error'
        try:
            task_log(task, "已连接Chrome，打开新标签页")
            task['phase'] = 'collect'
            task['started_at'] = time.time()

            def on_progress(collected, with_answer):
                task['total'] = collected
                task['progress'] = collected
                task['has_answer'] = with_answer

            def stop_with_timeout():
                if task.get('stop_requested'):
                    return True
                if task.get('started_at') and time.time() - task['started_at'] > task.get('timeout', 3600):
                    task['stop_requested'] = True
                    task['timed_out'] = True
                    task_log(task, "任务超时，正在停止...")
                    return True
                return False

            result, status = await collect_questions(
                page, paper_id,
                log_fn=lambda msg: task_log(task, msg),
                stop_check=stop_with_timeout,
                progress_fn=on_progress,
                start_sequence=start_sequence,
                max_questions=max_questions,
                mode=mode
            )

            if status == 'auth_error':
                task['phase'] = 'error'
                task['message'] = '登录超时或已取消，请在Chrome中登录考试宝后重试'
                return questions, 'auth_error'

            if result:
                added = _merge_questions(result, questions)
                task_log(task, f"新增 {added} 题，总计 {len(questions)} 题")
            else:
                task_log(task, "未收集到新题目，保留已有数据")
            return questions, status
        except Exception as e:
            logger.exception(f"爬取出错: {e}")
            task_log(task, f"爬取出错: {e}")
            return questions, 'error'
        finally:
            try:
                await page.close()
            except Exception:
                pass


def run_task(task_id):
    task = tasks[task_id]
    paper_id = task['paper_id']
    segment_id = task.get('segment_id', -1)
    start_sequence = task.get('start_sequence', 0)
    max_questions = task.get('max_questions', 0)
    mode = task.get('mode', 'ocr')
    is_segment = segment_id >= 0

    mode_suffix = f"_{mode}" if mode and mode != 'ocr' else ''
    output_file = OUTPUT_DIR / f"考试宝题库_{paper_id}{mode_suffix}.xlsx"
    task['output_file'] = str(output_file)

    questions = []
    if not is_segment:
        questions = load_progress(paper_id, suffix=mode_suffix)
        if questions:
            task_log(task, f"已有进度: {len(questions)} 道题目，从第{len(questions)+1}题继续")
            task['has_answer'] = sum(1 for q in questions if q.get('正确答案'))
            task['total'] = len(questions)
            task['progress'] = len(questions)
            start_sequence = len(questions)

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    try:
        from contextlib import nullcontext
        with (nullcontext() if is_segment else _chrome_lock):
            max_retries = 3
            status = 'error'
            for attempt in range(max_retries):
                if task.get('stop_requested'):
                    break
                if attempt > 0:
                    task_log(task, f"第 {attempt+1} 次尝试（重连Chrome）...")
                    task['started_at'] = 0
                    time.sleep(3)
                _result, status = loop.run_until_complete(
                    _do_scrape(task, paper_id, segment_id, is_segment, start_sequence, max_questions, mode, questions))
                if status == 'ok' and questions:
                    break
                if status == 'auth_error':
                    break
                if questions:
                    break
                if attempt < max_retries - 1:
                    task_log(task, f"第 {attempt+1} 次未成功，稍后重试...")
        if status != 'ok' or not questions:
            task['running'] = False
            task['phase'] = 'error'
            if status == 'auth_error':
                task['message'] = '登录超时或已取消，请在Chrome中登录考试宝后重试'
            else:
                task['message'] = f'爬取失败（{status}），已收集 {len(questions)} 题'
            task_log(task, f"爬取未完成（{status}），已收集 {len(questions)} 题")
            return
        if not questions:
            task['running'] = False
            task['phase'] = 'error'
            task['message'] = '未收集到任何题目，请检查登录状态后重试'
            task_log(task, "未收集到任何题目")
            return
        _finalize_task_success(task, questions, is_segment, paper_id, segment_id, mode_suffix, output_file)
    except Exception as e:
        logger.exception(f"任务执行错误: {e}")
        task_log(task, f"错误: {e}")
        task['phase'] = 'error'
        task['message'] = f"错误: {e}"
    finally:
        task['running'] = False
        loop.close()


# ---- Flask 路由 ----

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/chrome-status')
def chrome_status():
    ok = get_chrome_status()
    return jsonify({'online': ok})


@app.route('/api/chrome-start', methods=['POST'])
def chrome_start():
    if not check_rate_limit(request.remote_addr, 5.0):
        return jsonify({'error': '操作太频繁'}), 429
    logger.info(f"Chrome启动请求 from {request.remote_addr}")
    ok, msg, restarted = ensure_chrome_debug()

    login_status = None
    if ok and restarted:
        # Chrome重启后，自动打开本工具页面 + 考试宝并检查登录态
        try:
            from playwright.async_api import async_playwright
            async def reopen_pages():
                async with async_playwright() as p:
                    browser = await p.chromium.connect_over_cdp(CDP_URL)
                    if not browser.contexts:
                        return None
                    context = browser.contexts[0]
                    # 先打开本工具页面（用户能立刻看到界面）
                    tool_page = await context.new_page()
                    await tool_page.goto('http://localhost:5678')
                    # 再打开考试宝检查登录态
                    ks_page, logged_in = await reopen_kaoshibao(p)
                    if ks_page:
                        try:
                            await ks_page.close()
                        except Exception:
                            pass
                    return logged_in
            loop = asyncio.new_event_loop()
            login_status = loop.run_until_complete(reopen_pages())
            loop.close()
        except Exception as e:
            logger.error(f"Chrome重启后打开页面失败: {e}")

    logger.info(f"Chrome启动结果: ok={ok}, restarted={restarted}, logged_in={login_status}")
    return jsonify({'ok': ok, 'message': msg, 'restarted': restarted, 'logged_in': login_status})


@app.route('/api/tasks')
def list_tasks():
    with tasks_lock:
        result = []
        for tid, t in tasks.items():
            elapsed = 0
            if t.get('started_at') and (t['running'] or t['phase'] == 'done'):
                elapsed = int(time.time() - t['started_at'])
            result.append({
                'task_id': t['task_id'],
                'paper_id': t['paper_id'],
                'segment_id': t.get('segment_id', -1),
                'mode': t.get('mode', 'ocr'),
                'running': t['running'],
                'phase': t['phase'],
                'progress': t['progress'],
                'total': t['total'],
                'has_answer': t['has_answer'],
                'message': t['message'],
                'log': t['log'],
                'output_file': t['output_file'],
                'created_at': t['created_at'],
                'elapsed': elapsed,
                'timeout': t.get('timeout', 3600),
                'timed_out': t.get('timed_out', False),
            })
    return jsonify(result)


@app.route('/api/start', methods=['POST'])
def start_task():
    if not check_rate_limit(request.remote_addr, 2.0):
        return jsonify({'error': '操作太频繁，请稍后再试'}), 429

    data, err = get_request_json()
    if err:
        return jsonify({'error': err}), 400

    raw_id = data.get('paper_id', '')
    paper_id, err = validate_paper_id(raw_id)
    if err:
        return jsonify({'error': err}), 400

    mode = data.get('mode', 'ocr')
    VALID_MODES = {'ocr', 'dom', 'hybrid', 'race'}
    if mode not in VALID_MODES:
        return jsonify({'error': f'无效的读取模式: {mode}'}), 400

    # 检查Chrome是否可用
    if not get_chrome_status():
        return jsonify({'error': 'Chrome未启动，请先点击"启动Chrome"'}), 400

    # 竞速模式：同时启动OCR和DOM两个任务
    if mode == 'race':
        with tasks_lock:
            running_count = sum(1 for t in tasks.values() if t['running'])
            if running_count + 2 > MAX_TASKS * 2:
                return jsonify({'error': '当前运行任务过多'}), 400

            task_ids = []
            for m in ['hybrid', 'ocr', 'dom']:
                task_id = str(uuid.uuid4())[:8]
                task = new_task(paper_id, mode=m)
                task['task_id'] = task_id
                task['message'] = f'竞速模式({m.upper()})'
                tasks[task_id] = task
                task_ids.append(task_id)

        for tid in task_ids:
            thread = threading.Thread(target=run_task, args=(tid,), daemon=True)
            thread.start()
        return jsonify({'ok': True, 'task_ids': task_ids, 'mode': 'race'})

    with tasks_lock:
        running_count = sum(1 for t in tasks.values() if t['running'])
        if running_count >= MAX_TASKS:
            return jsonify({'error': f'最多同时运行{MAX_TASKS}个任务'}), 400

        for t in tasks.values():
            if t['paper_id'] == paper_id and t['running']:
                return jsonify({'error': f'题库 {paper_id} 正在爬取中'}), 400

        segment_id = data.get('segment_id', -1)
        start_sequence = data.get('start_sequence', 0)
        max_questions = data.get('max_questions', 0)
        if segment_id >= 0:
            start_sequence = data.get('start_sequence', 0)
            max_questions = data.get('max_questions', 0)

        task_id = str(uuid.uuid4())[:8]
        task = new_task(
            paper_id, segment_id=segment_id,
            start_sequence=start_sequence, max_questions=max_questions,
            mode=mode
        )
        task['task_id'] = task_id
        tasks[task_id] = task

    thread = threading.Thread(target=run_task, args=(task_id,), daemon=True)
    thread.start()
    return jsonify({'ok': True, 'task_id': task_id})


@app.route('/api/stop/<task_id>', methods=['POST'])
def stop_task(task_id):
    if not re.match(r'^[a-f0-9]{8}$', task_id):
        return jsonify({'error': '无效的任务ID'}), 400
    with tasks_lock:
        task = tasks.get(task_id)
        if not task:
            return jsonify({'error': '任务不存在'}), 404
        if not task['running']:
            return jsonify({'error': '任务已停止'}), 400
        task['stop_requested'] = True
        task_log(task, "正在停止...")
    return jsonify({'ok': True})


@app.route('/api/remove/<task_id>', methods=['POST'])
def remove_task(task_id):
    if not re.match(r'^[a-f0-9]{8}$', task_id):
        return jsonify({'error': '无效的任务ID'}), 400
    with tasks_lock:
        task = tasks.get(task_id)
        if not task:
            return jsonify({'error': '任务不存在'}), 404
        if task['running']:
            return jsonify({'error': '请先停止任务'}), 400
        del tasks[task_id]
    return jsonify({'ok': True})


@app.route('/api/files')
def list_files():
    files = []
    for f in OUTPUT_DIR.glob('*.xlsx'):
        files.append({
            'name': f.name,
            'size': f.stat().st_size,
            'time': f.stat().st_mtime
        })
    files.sort(key=lambda x: x['time'], reverse=True)
    return jsonify(files)


@app.route('/api/download/<filename>')
def download_file(filename):
    if not filename or '..' in filename or '/' in filename or '\\' in filename:
        return jsonify({'error': '无效文件名'}), 400
    if not re.match(r'^[\w\-\.一-鿿]+$', filename):
        return jsonify({'error': '无效文件名'}), 400
    filepath = OUTPUT_DIR / filename
    if not filepath.exists() or filepath.suffix != '.xlsx':
        return jsonify({'error': '文件不存在'}), 404
    return send_file(filepath, as_attachment=True, download_name=filename)


@app.route('/api/delete-file', methods=['POST'])
def delete_file():
    data, err = get_request_json()
    if err:
        return jsonify({'error': err}), 400
    name = data.get('name', '')
    if not name or '..' in name or '/' in name or '\\' in name:
        return jsonify({'error': '无效文件名'}), 400
    filepath = OUTPUT_DIR / name
    if filepath.exists() and filepath.suffix == '.xlsx':
        filepath.unlink()
        return jsonify({'ok': True})
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/batch-start', methods=['POST'])
def batch_start():
    if not check_rate_limit(request.remote_addr, 3.0):
        return jsonify({'error': '操作太频繁'}), 429

    data, err = get_request_json()
    if err:
        return jsonify({'error': err}), 400

    ids_raw = data.get('ids', [])
    if not ids_raw or not isinstance(ids_raw, list):
        return jsonify({'error': '请输入题库ID列表'}), 400
    if len(ids_raw) > 10:
        return jsonify({'error': '单次最多10个'}), 400

    if not get_chrome_status():
        return jsonify({'error': 'Chrome未启动，请先启动Chrome'}), 400

    results = []
    for raw_id in ids_raw:
        raw_str = str(raw_id)
        if len(raw_str) > 200:
            results.append({'id': raw_str[:50], 'error': 'ID 过长'})
            continue
        paper_id, err = validate_paper_id(raw_str)
        if err:
            results.append({'id': raw_str, 'error': err})
            continue

        with tasks_lock:
            running_count = sum(1 for t in tasks.values() if t['running'])
            if running_count >= MAX_TASKS:
                results.append({'id': paper_id, 'error': f'已达最大并发数{MAX_TASKS}'})
                continue

            duplicate = any(t['paper_id'] == paper_id and t['running'] for t in tasks.values())
            if duplicate:
                results.append({'id': paper_id, 'error': '已在运行中'})
                continue

            task_id = str(uuid.uuid4())[:8]
            task = new_task(paper_id)
            task['task_id'] = task_id
            tasks[task_id] = task

        thread = threading.Thread(target=run_task, args=(task_id,), daemon=True)
        thread.start()
        results.append({'id': paper_id, 'ok': True, 'task_id': task_id})

    return jsonify({'results': results})


@app.route('/api/start-parallel', methods=['POST'])
def start_parallel():
    """分段并行爬取：把一个题库拆成多段同时爬"""
    if not check_rate_limit(request.remote_addr, 5.0):
        return jsonify({'error': '操作太频繁'}), 429

    data, err = get_request_json()
    if err:
        return jsonify({'error': err}), 400

    raw_id = data.get('paper_id', '')
    paper_id, err = validate_paper_id(raw_id)
    if err:
        return jsonify({'error': err}), 400

    workers = data.get('workers', 4)
    if not isinstance(workers, int) or workers < 2 or workers > 5:
        return jsonify({'error': '并行数需在2-5之间'}), 400

    mode = data.get('mode', 'hybrid')
    if mode not in ('ocr', 'dom', 'hybrid'):
        mode = 'hybrid'

    total_questions = data.get('total_questions', 702)
    if not isinstance(total_questions, int) or total_questions < 10:
        return jsonify({'error': '总题数需>=10'}), 400

    if not get_chrome_status():
        return jsonify({'error': 'Chrome未启动，请先启动Chrome'}), 400

    # 计算每段的题数和起始位置
    from scraper_core import load_segment_progress
    seg_size = total_questions // workers
    segments = []
    for i in range(workers):
        start = i * seg_size
        count = seg_size if i < workers - 1 else total_questions - start
        segments.append({'segment_id': i, 'start_sequence': start, 'max_questions': count})

    # 跳过已爬完整的分段（断点续爬，避免重复爬已完成的段）
    active_segments = []
    skipped = []
    for seg in segments:
        existing = load_segment_progress(paper_id, seg['segment_id'])
        if len(existing) >= seg['max_questions']:
            skipped.append(seg['segment_id'])
            continue
        active_segments.append(seg)
    if not active_segments:
        return jsonify({'ok': False, 'error': '所有分段均已爬完，无需重复爬取'}), 400

    with tasks_lock:
        running_count = sum(1 for t in tasks.values() if t['running'])
        if running_count + len(active_segments) > MAX_TASKS * 2:
            return jsonify({'error': f'当前运行任务过多，请稍后再试'}), 400

        task_ids = []
        for seg in active_segments:
            task_id = str(uuid.uuid4())[:8]
            task = new_task(
                paper_id,
                segment_id=seg['segment_id'],
                start_sequence=seg['start_sequence'],
                max_questions=seg['max_questions'],
                mode=mode
            )
            task['task_id'] = task_id
            task['message'] = f"分段{seg['segment_id']+1}/{workers}: 第{seg['start_sequence']+1}-{seg['start_sequence']+seg['max_questions']}题"
            tasks[task_id] = task
            task_ids.append(task_id)

    # 启动分段任务
    for tid in task_ids:
        thread = threading.Thread(target=run_task, args=(tid,), daemon=True)
        thread.start()

    return jsonify({
        'ok': True,
        'task_ids': task_ids,
        'paper_id': paper_id,
        'workers': workers,
        'skipped': skipped,
        'segments': active_segments
    })


@app.route('/api/merge-segments', methods=['POST'])
def merge_segments_api():
    """合并分段结果"""
    data, err = get_request_json()
    if err:
        return jsonify({'error': err}), 400

    raw_id = data.get('paper_id', '')
    paper_id, err = validate_paper_id(raw_id)
    if err:
        return jsonify({'error': err}), 400

    workers = data.get('workers', 4)

    # 检查所有分段是否完成
    with tasks_lock:
        for t in tasks.values():
            if t['paper_id'] == paper_id and t.get('segment_id', -1) >= 0 and t['running']:
                return jsonify({'error': '还有分段任务在运行中'}), 400

    result = merge_segments(paper_id, workers)
    has_ans = sum(1 for q in result if q.get('正确答案'))
    return jsonify({
        'ok': True,
        'total': len(result),
        'has_answer': has_ans
    })


@app.route('/api/delete-files', methods=['POST'])
def delete_files_batch():
    """批量删除文件"""
    data, err = get_request_json()
    if err:
        return jsonify({'error': err}), 400
    names = data.get('names', [])
    if not isinstance(names, list) or len(names) == 0:
        return jsonify({'error': '请提供文件名列表'}), 400
    if len(names) > 100:
        return jsonify({'error': '单次最多删除 100 个文件'}), 400
    removed = 0
    failed = 0
    for name in names:
        if not name or '..' in name or '/' in name or '\\' in name or len(str(name)) > 255:
            failed += 1
            continue
        filepath = OUTPUT_DIR / name
        if filepath.exists() and filepath.suffix == '.xlsx':
            try:
                filepath.unlink()
                removed += 1
            except Exception:
                failed += 1
        else:
            failed += 1
    return jsonify({'ok': True, 'removed': removed, 'failed': failed})


@app.route('/api/cleanup', methods=['POST'])
def cleanup_old_progress():
    """清理超过指定天数未修改的进度文件"""
    data, err = get_request_json()
    if err:
        data = {}
    days = data.get('days', 30) if data else 30
    try:
        days = int(days)
    except (TypeError, ValueError):
        return jsonify({'error': 'days 必须为整数'}), 400
    if days < 1 or days > 365:
        return jsonify({'error': 'days 需在 1-365 之间'}), 400

    cutoff = time.time() - days * 86400
    removed = []
    for f in DATA_DIR.glob('progress_*.json'):
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
                removed.append(f.name)
        except Exception:
            pass
    logger.info(f"清理 {len(removed)} 个超过 {days} 天的进度文件")
    return jsonify({'ok': True, 'removed': len(removed), 'files': removed[:50], 'days': days})


@app.route('/api/clear-progress/<paper_id>', methods=['POST'])
def clear_progress(paper_id):
    """清除指定题库的进度文件（用于重试）"""
    paper_id, err = validate_paper_id(paper_id)
    if err:
        return jsonify({'error': err}), 400
    removed = 0
    for f in DATA_DIR.glob(f'progress_{paper_id}*.json'):
        try:
            f.unlink()
            removed += 1
        except Exception:
            pass
    logger.info(f"清除题库 {paper_id} 进度: 删除 {removed} 个文件")
    return jsonify({'ok': True, 'removed': removed})


@app.route('/api/history')
def history():
    """获取已爬取过的题库历史（基于 progress 文件）"""
    history_list = []
    for f in DATA_DIR.glob('progress_*.json'):
        name = f.name
        if '_seg' in name:
            continue
        m = re.match(r'progress_(\d+)(.*)\.json', name)
        if not m:
            continue
        paper_id = m.group(1)
        suffix = m.group(2)
        try:
            with open(f, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            count = len(data)
            has_answer = sum(1 for q in data if q.get('正确答案'))
        except Exception:
            count = 0
            has_answer = 0
        history_list.append({
            'paper_id': paper_id,
            'mode_suffix': suffix,
            'count': count,
            'has_answer': has_answer,
            'mtime': f.stat().st_mtime,
        })
    history_list.sort(key=lambda x: x['mtime'], reverse=True)
    return jsonify(history_list)


@app.route('/api/export/<paper_id>', methods=['GET'])
def export_paper(paper_id):
    """从已保存进度导出为指定格式（csv/json/xlsx/docx）"""
    paper_id, err = validate_paper_id(paper_id)
    if err:
        return jsonify({'error': err}), 400
    fmt = (request.args.get('format') or 'xlsx').lower()
    if fmt in ('word',):
        fmt = 'docx'
    if fmt not in ('xlsx', 'csv', 'json', 'docx'):
        return jsonify({'error': '不支持的格式，可选 xlsx/csv/json/docx'}), 400

    # 进度文件可能带模式后缀（_hybrid / _dom 等），按优先级查找
    questions = load_progress(paper_id)
    suffix_used = ''
    if not questions:
        for sfx in ('_hybrid', '_dom', '_ocr'):
            q = load_progress(paper_id, suffix=sfx)
            if q:
                questions = q
                suffix_used = sfx
                break
    if not questions:
        return jsonify({'error': f'题库 {paper_id} 无进度数据，请先爬取'}), 404
    logger.info(f"导出来源: progress_{paper_id}{suffix_used}.json（{len(questions)} 题）")

    ext = fmt
    filename = f"考试宝题库_{paper_id}.{ext}"
    filepath = OUTPUT_DIR / filename
    try:
        if fmt == 'docx':
            from scraper_core import save_to_docx

            def _log(m):
                logger.info(m)

            save_to_docx(questions, str(filepath),
                         title=f"考试宝题库 {paper_id}", log_fn=_log)
        else:
            export_questions(questions, str(filepath), fmt=fmt)
    except Exception as e:
        logger.exception(f"导出失败: {e}")
        return jsonify({'error': f'导出失败: {e}'}), 500
    return send_file(filepath, as_attachment=True, download_name=filename)


@app.route('/api/audit/<paper_id>', methods=['GET'])
def audit_paper(paper_id):
    """核对题库完整性：返回总题数、缺选项/答案的题清单"""
    paper_id, err = validate_paper_id(paper_id)
    if err:
        return jsonify({'error': err}), 400

    questions = load_progress(paper_id)
    suffix_used = ''
    if not questions:
        for sfx in ('_hybrid', '_dom', '_ocr'):
            q = load_progress(paper_id, suffix=sfx)
            if q:
                questions = q
                suffix_used = sfx
                break
    if not questions:
        return jsonify({'error': f'题库 {paper_id} 无进度数据'}), 404

    from scraper_core import audit_questions
    report = audit_questions(questions)
    report['source'] = f'progress_{paper_id}{suffix_used}.json'
    return jsonify(report)


@app.route('/api/fix/<paper_id>', methods=['POST'])
def fix_paper(paper_id):
    """核对并回抓缺失选项/答案的题目"""
    paper_id, err = validate_paper_id(paper_id)
    if err:
        return jsonify({'error': err}), 400

    suffix_used = ''
    questions = load_progress(paper_id)
    if not questions:
        for sfx in ('_hybrid', '_dom', '_ocr'):
            q = load_progress(paper_id, suffix=sfx)
            if q:
                questions = q
                suffix_used = sfx
                break
    if not questions:
        return jsonify({'error': f'题库 {paper_id} 无进度数据'}), 404

    if not get_chrome_status():
        return jsonify({'error': 'Chrome 未连接，请先启动 Chrome'}), 400

    from scraper_core import audit_questions, rescrape_questions, open_cdp_page, save_progress

    report = audit_questions(questions)
    if report['bad'] == 0:
        return jsonify({'ok': True, 'message': '全部题目选项/答案齐全，无需补全',
                        'total': report['total'], 'fixed': 0})

    async def _run():
        from playwright.async_api import async_playwright
        async with async_playwright() as p:
            browser, page = await open_cdp_page(p)
            if not page:
                return {'ok': False, 'error': '无法连接 Chrome'}
            try:
                targets = []
                for item in report['bad_list']:
                    orig = next((x for x in questions if x.get('序号') == item['序号']), None)
                    if orig is not None:
                        targets.append({'序号': item['序号'], '缺失': item['缺失'], '_orig': orig})
                res = await rescrape_questions(page, paper_id, targets, len(questions),
                                               mode='hybrid')
                save_progress(paper_id, questions,
                              suffix=suffix_used)
                after = audit_questions(questions)
                res['after'] = {'total': after['total'], 'ok': after['ok'], 'bad': after['bad']}
                res['ok'] = True
                return res
            finally:
                try:
                    await page.close()
                except Exception:
                    pass
                try:
                    await browser.close()
                except Exception:
                    pass

    try:
        loop = asyncio.new_event_loop()
        try:
            result = loop.run_until_complete(_run())
        finally:
            loop.close()
        result['before'] = {'total': report['total'], 'bad': report['bad']}
        return jsonify(result)
    except Exception as e:
        logger.exception(f"补全失败: {e}")
        return jsonify({'error': f'补全失败: {e}'}), 500


def _build_openapi_spec() -> dict:
    """构建 OpenAPI 3.0 规范（替代手动维护的文档列表）"""
    return {
        'openapi': '3.0.0',
        'info': {
            'title': '考试宝题库爬取API',
            'version': '1.0.0',
            'description': '通过Chrome已登录会话爬取考试宝题库的Web API'
        },
        'paths': {
            '/': {'get': {'summary': 'Web界面', 'responses': {'200': {'description': 'HTML页面'}}}},
            '/health': {'get': {'summary': '健康检查', 'responses': {'200': {'description': '服务健康'}, '503': {'description': '服务不健康'}}}},
            '/manifest.json': {'get': {'summary': 'PWA manifest', 'responses': {'200': {'description': 'manifest JSON'}}}},
            '/api/docs': {'get': {'summary': 'API文档(OpenAPI规范)', 'responses': {'200': {'description': 'OpenAPI spec'}}}},
            '/api/stats': {'get': {'summary': '监控指标', 'responses': {'200': {'description': '统计信息'}}}},
            '/api/history': {'get': {'summary': '已爬取题库历史', 'responses': {'200': {'description': '历史列表'}}}},
            '/api/chrome-status': {'get': {'summary': 'Chrome调试端口状态', 'responses': {'200': {'description': '状态信息'}}}},
            '/api/chrome-start': {'post': {'summary': '启动Chrome调试模式', 'responses': {'200': {'description': '启动结果'}, '429': {'description': '操作太频繁'}}}},
            '/api/tasks': {'get': {'summary': '获取所有任务状态', 'responses': {'200': {'description': '任务列表'}}}},
            '/api/start': {'post': {'summary': '启动单个爬取任务', 'requestBody': {'content': {'application/json': {'schema': {'type': 'object', 'properties': {'paper_id': {'type': 'string'}, 'mode': {'type': 'string', 'enum': ['ocr', 'dom', 'hybrid', 'race']}}}}}}, 'responses': {'200': {'description': '任务已启动'}, '400': {'description': '参数错误'}}}},
            '/api/stop/{task_id}': {'post': {'summary': '停止任务', 'parameters': [{'name': 'task_id', 'in': 'path', 'required': True, 'schema': {'type': 'string', 'pattern': '^[a-f0-9]{8}$'}}], 'responses': {'200': {'description': '已停止'}, '400': {'description': '无效task_id'}}}},
            '/api/remove/{task_id}': {'post': {'summary': '移除已完成任务', 'parameters': [{'name': 'task_id', 'in': 'path', 'required': True, 'schema': {'type': 'string', 'pattern': '^[a-f0-9]{8}$'}}], 'responses': {'200': {'description': '已移除'}, '400': {'description': '无效task_id'}}}},
            '/api/files': {'get': {'summary': '列出Excel文件', 'responses': {'200': {'description': '文件列表'}}}},
            '/api/download/{filename}': {'get': {'summary': '下载Excel文件', 'parameters': [{'name': 'filename', 'in': 'path', 'required': True, 'schema': {'type': 'string'}}], 'responses': {'200': {'description': '文件内容'}, '404': {'description': '文件不存在'}}}},
            '/api/export/{paper_id}': {'get': {'summary': '导出为xlsx/csv/json', 'parameters': [{'name': 'paper_id', 'in': 'path', 'required': True, 'schema': {'type': 'string'}}, {'name': 'format', 'in': 'query', 'schema': {'type': 'string', 'enum': ['xlsx', 'csv', 'json']}}], 'responses': {'200': {'description': '导出文件'}, '404': {'description': '无数据'}}}},
            '/api/delete-file': {'post': {'summary': '删除单个文件', 'requestBody': {'content': {'application/json': {'schema': {'type': 'object', 'properties': {'name': {'type': 'string'}}}}}}, 'responses': {'200': {'description': '已删除'}}}},
            '/api/delete-files': {'post': {'summary': '批量删除文件', 'requestBody': {'content': {'application/json': {'schema': {'type': 'object', 'properties': {'names': {'type': 'array', 'items': {'type': 'string'}}}}}}}, 'responses': {'200': {'description': '删除结果'}}}},
            '/api/batch-start': {'post': {'summary': '批量启动任务', 'requestBody': {'content': {'application/json': {'schema': {'type': 'object', 'properties': {'ids': {'type': 'array', 'items': {'type': 'string'}}}}}}}, 'responses': {'200': {'description': '批量启动结果'}}}},
            '/api/start-parallel': {'post': {'summary': '分段并行爬取', 'requestBody': {'content': {'application/json': {'schema': {'type': 'object', 'properties': {'paper_id': {'type': 'string'}, 'total_questions': {'type': 'integer'}, 'workers': {'type': 'integer'}}}}}}, 'responses': {'200': {'description': '已启动'}, '400': {'description': '参数错误'}}}},
            '/api/merge-segments': {'post': {'summary': '合并分段结果', 'requestBody': {'content': {'application/json': {'schema': {'type': 'object', 'properties': {'paper_id': {'type': 'string'}, 'workers': {'type': 'integer'}}}}}}, 'responses': {'200': {'description': '合并结果'}, '400': {'description': '参数错误'}}}},
            '/api/clear-progress/{paper_id}': {'post': {'summary': '清除题库进度', 'parameters': [{'name': 'paper_id', 'in': 'path', 'required': True, 'schema': {'type': 'string'}}], 'responses': {'200': {'description': '已清除'}}}},
            '/api/cleanup': {'post': {'summary': '清理旧进度文件', 'requestBody': {'content': {'application/json': {'schema': {'type': 'object', 'properties': {'days': {'type': 'integer'}}}}}}, 'responses': {'200': {'description': '清理结果'}}}},
        }
    }


@app.route('/api/docs')
def api_docs():
    """API 文档端点：返回 OpenAPI 3.0 规范"""
    return jsonify(_build_openapi_spec())


@app.route('/manifest.json')
def manifest():
    return jsonify({
        'name': '考试宝题库爬取工具',
        'short_name': '考试宝爬取',
        'description': '通过Chrome已登录会话爬取考试宝题库',
        'start_url': '/',
        'display': 'standalone',
        'background_color': '#f0f2f5',
        'theme_color': '#4a90d9',
        'icons': [
            {'src': 'data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 64 64%22%3E%3Crect width=%2264%22120%22 height=%2264%22 rx=%2212%22 fill=%22%234a90d9%22/%3E%3C/svg%3E', 'sizes': '64x64', 'type': 'image/svg+xml'}
        ]
    })


@app.route('/health')
def health_check():
    """健康检查端点（用于 Docker / 负载均衡 / 监控）"""
    chrome_ok = get_chrome_status()
    with tasks_lock:
        running = sum(1 for t in tasks.values() if t['running'])
        total = len(tasks)
        failed = sum(1 for t in tasks.values() if t.get('error') and not t['running'])
    return jsonify({
        'status': 'healthy',
        'ready': chrome_ok,
        'chrome': 'online' if chrome_ok else 'offline',
        'tasks_running': running,
        'tasks_total': total,
        'tasks_failed': failed,
        'uptime': int(time.time() - _start_time)
    }), 200


@app.route('/api/stats')
def stats():
    """监控指标端点"""
    with tasks_lock:
        running = sum(1 for t in tasks.values() if t['running'])
        done = sum(1 for t in tasks.values() if t['phase'] == 'done')
        error = sum(1 for t in tasks.values() if t['phase'] == 'error')
        total = len(tasks)
    try:
        file_count = sum(1 for _ in OUTPUT_DIR.glob('*.xlsx'))
    except Exception:
        file_count = 0
    return jsonify({
        'tasks': {'running': running, 'done': done, 'error': error, 'total': total, 'max': MAX_TASKS},
        'files': file_count,
        'chrome_online': get_chrome_status(),
        'uptime': int(time.time() - _start_time)
    })

try:
    from db_store import init_db, is_db_configured
    if is_db_configured():
        init_db()
        logger.info("PostgreSQL存储层已启用")
    else:
        logger.info("PostgreSQL未配置，使用文件存储")
except Exception as e:
    logger.warning(f"数据库模块加载失败: {e}")


if __name__ == '__main__':
    logger.info("正在检查Chrome...")
    ok, msg, restarted = ensure_chrome_debug()
    logger.info(msg)

    import webbrowser
    # Chrome就绪后再打开页面（如果Chrome重启了，新窗口里打开）
    if ok:
        time.sleep(1)  # 等Chrome完全就绪
        webbrowser.open('http://localhost:5678')
    else:
        logger.warning("Chrome未就绪，请手动打开 http://localhost:5678")

    app.run(host='0.0.0.0', port=5678, debug=False)
