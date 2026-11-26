"""考试宝爬取核心逻辑 - 共享模块"""

from __future__ import annotations

import asyncio
import json
import re
import html
import subprocess
import sys
import io
import time
import sys
import socket
import subprocess
import time
import re
import json
import os
import asyncio
import logging
from pathlib import Path
from typing import Any
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# 日志配置
from logging_config import init_default_logger, get_logger
init_default_logger(level="INFO")
logger = get_logger("kaoshibao.scraper_core")

# OCR引擎（延迟加载）
_ocr_engine = None
def _get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine

# Windows控制台GBK编码兼容：强制stdout/stderr为UTF-8
if sys.platform == 'win32' and hasattr(sys.stdout, 'buffer'):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except Exception:
        pass

# 项目根目录
PROJECT_ROOT = Path(__file__).parent
OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# 数据目录（进度文件、cookies、字体映射）
DATA_DIR = PROJECT_ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

# ---- Playwright 临时目录修正 ----
# Playwright 会在系统 %TEMP% 下 mkdtemp 存放 artifacts，某些 Windows 环境下
# 那个目录不可写，导致 connect_over_cdp 直接报错：
#     EPERM: operation not permitted, mkdtemp '...\Temp\playwright-artifacts-XXXXXX'
# 这里把 TEMP/TMP 指向项目内可写目录，避免依赖系统临时目录权限。
# 设置 KSB_KEEP_TEMP=1 可跳过（用于排查是否为本项目引起）。
if os.environ.get('KSB_KEEP_TEMP') != '1':
    try:
        _LOCAL_TMP = PROJECT_ROOT / "tmp"
        _LOCAL_TMP.mkdir(exist_ok=True)
        # 先验证真的可写，避免把 TEMP 指到一个同样不可写的地方
        _probe = _LOCAL_TMP / ".write_probe"
        _probe.write_text("ok", encoding="utf-8")
        _probe.unlink()
        os.environ['TEMP'] = str(_LOCAL_TMP)
        os.environ['TMP'] = str(_LOCAL_TMP)
    except Exception:
        pass

CDP_PORT = int(os.environ.get("CDP_PORT", "9223"))
CDP_URL = f"http://127.0.0.1:{CDP_PORT}"

# 支持的浏览器列表（按优先级排序，根据平台自动选择）
#
# ⚠️ 关于 --user-data-dir：本项目一律使用「专用 profile」，绝不碰用户的默认 profile。
#
#   原因有两层：
#
#   1. 技术上：Chrome 136 起（实测 158 更严格）直接拒绝用默认目录开调试端口，
#      报错 DevTools remote debugging requires a non-default data directory.
#      这是 Chrome 防止调试端口窃取 cookie 的加固措施，不是本项目的 bug。
#
#   2. 安全上：调试端口一旦接上，等于能以完整登录态操作那个浏览器。
#      用户的默认 profile 里有 ta 所有网站的 cookie（网银、邮箱、社交……）。
#      爬个题库不该把这个面暴露出来。用专用 profile 就只暴露考试宝这一个站。
#
#   代价：专用 profile 是全新的，需要你在里面登录一次考试宝。
#   这是一次性成本，之后登录态会保存在这个目录里。
#
#   profile 目录固定在项目的 data/chrome-profile（已在 .gitignore 中排除）。
_CDP_PROFILE_DIR = str(PROJECT_ROOT / "data" / "chrome-profile")

if sys.platform == 'win32':
    _BROWSER_PROFILES = [
        {
            'name': '360浏览器',
            'process': '360ChromeX.exe',
            'exe_path': str(Path.home() / "AppData/Local/360ChromeX/Chrome/Application/360ChromeX.exe"),
            'user_data': _CDP_PROFILE_DIR,
        },
        {
            'name': 'Google Chrome',
            'process': 'chrome.exe',
            'exe_path': r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            'user_data': _CDP_PROFILE_DIR,
        },
        {
            'name': 'Edge',
            'process': 'msedge.exe',
            'exe_path': r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            'user_data': _CDP_PROFILE_DIR,
        },
    ]
else:
    _pw_chrome = str(Path.home() / ".cache/ms-playwright")
    _pw_chrome_path = ''
    for p in sorted(Path(_pw_chrome).glob('chromium-*/chrome-linux64/chrome'), reverse=True):
        _pw_chrome_path = str(p)
        break
    _BROWSER_PROFILES = [
        {
            'name': 'Playwright Chromium',
            'process': 'chrome',
            'exe_path': _pw_chrome_path or '/usr/bin/chromium',
            'user_data': str(Path.home() / ".config/chromium"),
        },
        {
            'name': 'Google Chrome',
            'process': 'chrome',
            'exe_path': '/usr/bin/google-chrome',
            'user_data': str(Path.home() / ".config/google-chrome"),
        },
        {
            'name': 'Chromium',
            'process': 'chromium',
            'exe_path': '/usr/bin/chromium',
            'user_data': str(Path.home() / ".config/chromium"),
        },
    ]


def is_port_open(port: int, host: str = '127.0.0.1') -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex((host, port)) == 0


def _detect_browser():
    """检测当前正在运行的Chromium浏览器，返回浏览器配置或None"""
    for browser in _BROWSER_PROFILES:
        try:
            if sys.platform == 'win32':
                r = subprocess.run(
                    ["tasklist", "/FI", f"IMAGENAME eq {browser['process']}", "/NH"],
                    capture_output=True, text=True, timeout=5
                )
                if browser['process'].lower() not in r.stdout.lower():
                    continue
            else:
                r = subprocess.run(
                    ["pgrep", "-x", browser['process']],
                    capture_output=True, text=True, timeout=5
                )
                if r.returncode != 0:
                    continue
            if Path(browser['exe_path']).exists():
                return browser
        except Exception:
            continue
    return None


def _detect_any_installed_browser():
    """检测任意已安装的Chromium浏览器（不管是否在运行）"""
    for browser in _BROWSER_PROFILES:
        if Path(browser['exe_path']).exists():
            return browser
    return None


def _is_default_user_data(user_data: str) -> bool:
    """判断某个 user_data 目录是否是浏览器的默认 profile 目录。

    这是安全兜底：带调试端口的浏览器不该使用用户日常 profile，
    否则调试端口就等于能以完整登录态操作 ta 所有网站的账号。
    """
    try:
        target = Path(user_data).resolve()
    except Exception:
        return False
    home = Path.home()
    defaults = [
        home / "AppData/Local/Google/Chrome/User Data",
        home / "AppData/Local/360ChromeX/Chrome/User Data",
        home / "AppData/Local/Microsoft/Edge/User Data",
        home / ".config/google-chrome",
        home / ".config/chromium",
    ]
    for d in defaults:
        try:
            if target == d.resolve():
                return True
        except Exception:
            continue
    return False


def _spawn_detached(exe: str, args: list) -> bool:
    """启动一个独立进程（不阻塞、不继承句柄）。

    历史包袱说明：曾经以为 Popen 拉不起 Chrome（当时看到 0x80000003 误判成
    0xC0000003），试过 DETACHED_PROCESS / cmd start / ShellExecuteW /
    PowerShell .NET Process 等多种绕法。

    真正的原因是 Chrome 沙箱初始化失败，解决办法是启动时带 --no-sandbox
    （见 _start_browser_with_debug 里的说明）。带上之后，普通 Popen 就够了，
    不需要那些绕路方案。
    """
    kwargs = dict(
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    try:
        if sys.platform == 'win32':
            # 不让子进程继承控制台，避免关掉本服务时把浏览器一起带走
            kwargs['creationflags'] = 0x00000008  # DETACHED_PROCESS
        else:
            kwargs['start_new_session'] = True
        subprocess.Popen([exe] + args, **kwargs)
        return True
    except Exception as e:
        logger.error(f"启动浏览器失败: {e}")
        return False


def _start_browser_with_debug(browser):
    """以调试模式启动指定浏览器"""
    user_data = browser['user_data']
    if _is_default_user_data(user_data):
        logger.error(
            f"拒绝用浏览器默认 profile 启动调试端口: {user_data}\n"
            f"调试端口能以完整登录态操作该浏览器，不应指向用户日常 profile。"
        )
        return False

    args = [
        f"--remote-debugging-port={CDP_PORT}",
        "--no-first-run",
        "--no-default-browser-check",
        f"--user-data-dir={user_data}",
        # Windows 上必须带 --no-sandbox
        #
        # 否则 Chrome 会在沙箱初始化阶段直接退出，退出码 0x80000003
        # (STATUS_BREAKPOINT)，进程数归零、端口不开，而且 stderr 往往什么都不打。
        # 打开 Chrome 日志能看到真正的错误：
        #     ERROR:crashpad_util\file\file_io_win.cc:181]
        #     CreateFile ...\User Data\Crashpad\settings.dat: 拒绝访问。 (0x5)
        #
        # 实测（profile 换空目录也无法避免，说明与 profile 内容无关）：
        #     空 profile + CDP       -> 0x80000003
        #     仅 profile，无 CDP     -> 0x80000003
        #     无任何参数             -> 0x80000003
        #     加 --no-sandbox        -> 正常启动，端口打开
        "--no-sandbox",
        "--disable-gpu",
    ]
    if sys.platform != 'win32':
        args.append("--disable-dev-shm-usage")

    if not _spawn_detached(browser['exe_path'], args):
        return False

    for _ in range(25):
        time.sleep(1)
        if is_port_open(CDP_PORT):
            return True
    return False


def _chrome_proc_count():
    """当前 chrome.exe 进程数（Windows）；非 Windows 返回 -1"""
    if sys.platform != 'win32':
        return -1
    try:
        r = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
            capture_output=True, text=True, timeout=8,
            encoding='utf-8', errors='ignore',
        )
        return r.stdout.lower().count('chrome.exe')
    except Exception:
        return -1


def _kill_chrome(process_name='chrome.exe'):
    """结束 Chrome 进程。

    ⚠️ 注意：这个函数不会 100% 成功，因为 Chrome 有进程守护/崩溃恢复机制，
    主进程被杀后可能被重新拉起。因此不要依赖它来「重启用户的浏览器」。

    ensure_chrome_debug() 已经改成不主动杀用户的浏览器了，
    这个函数只在用户明确要求重启时使用。
    """
    if sys.platform == 'win32':
        for attempt in range(3):
            subprocess.run(["taskkill", "/F", "/IM", process_name, "/T"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(1.5)
            left = _chrome_proc_count()
            if left <= 0:
                logger.info(f"Chrome 已退出（第 {attempt + 1} 轮）")
                return True
            logger.debug(f"Chrome 仍有 {left} 个进程（第 {attempt + 1} 轮）")
        return _chrome_proc_count() <= 0
    else:
        subprocess.run(["pkill", "-f", process_name],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(2)
        return True


def ensure_chrome_debug():
    """
    确保有可用的调试端口。返回 (ok, message, chrome_restarted)

    设计原则：不去动用户正在使用的浏览器。

    如果 Chrome 已经在运行但没有调试端口，杀掉它再重启是行不通的：
      - 杀掉用户浏览器 = 丢掉 ta 所有标签页和会话
      - Chrome 有进程守护，杀了可能被重新拉起，杀不干净
      - 即使杀干净，重启后调试端口也未必能开（Chrome 136+ 对默认目录的限制）

    所以这里改成：
      1. 端口已开 -> 直接用
      2. Chrome 在跑但没端口 -> 不杀，告诉用户怎么办
      3. Chrome 没在跑 -> 用专用 profile 启动（不碰用户默认 profile）
    """
    if is_port_open(CDP_PORT):
        return True, f"调试端口 {CDP_PORT} 已就绪，复用现有连接", False

    # 有浏览器在运行，但不带调试端口
    running_browser = _detect_browser()
    if running_browser:
        # 不杀。用专用 profile 另起一个实例（Chrome 支持多实例并行）
        installed = _detect_any_installed_browser()
        if installed:
            logger.info(f"{running_browser['name']} 正在运行但无调试端口，"
                        f"尝试用专用 profile 另起一个实例（不影响你当前的浏览器）")
            if _start_browser_with_debug(installed):
                return (True,
                        f"已用专用 profile 启动调试实例（profile: {installed['user_data']}）。\n"
                        f"你原来的浏览器窗口未受影响。\n"
                        f"如需登录考试宝，请在弹出的新窗口里操作。",
                        True)
        return (False,
                f"{running_browser['name']} 正在运行，但没有调试端口。\n"
                f"本项目不会关掉你正在用的浏览器。\n\n"
                f"请选一种方式：\n"
                f"  1. 手动关闭所有 {running_browser['process']} 窗口，然后重启本项目\n"
                f"  2. 或先用专用 profile 手动开一个调试实例：\n"
                f"     \"{running_browser['exe_path']}\" "
                f"--remote-debugging-port={CDP_PORT} "
                f"--user-data-dir=\"{running_browser['user_data']}\"\n"
                f"     然后刷新本页面",
                False)

    # 没有浏览器在运行，正常启动
    installed = _detect_any_installed_browser()
    if not installed:
        return False, "未找到Chrome/360浏览器/Edge，请先安装其中一个", False

    ok = _start_browser_with_debug(installed)
    if ok:
        return (True,
                f"{installed['name']}已启动，调试端口就绪\n"
                f"profile: {installed['user_data']}", True)
    return (False,
            f"{installed['name']}启动超时。\n"
            f"当前 user-data-dir: {installed['user_data']}\n"
            f"可手动执行：\n"
            f"  \"{installed['exe_path']}\" --remote-debugging-port={CDP_PORT} "
            f"--user-data-dir=\"{installed['user_data']}\"",
            False)


async def reopen_kaoshibao(playwright):
    """Chrome重启后重新打开考试宝页面，检查登录态。返回 (page, logged_in)"""
    try:
        browser = await playwright.chromium.connect_over_cdp(CDP_URL)
        if not browser.contexts:
            return None, False
        context = browser.contexts[0]
        page = await context.new_page()
        await navigate_with_retry(page, 'https://www.kaoshibao.com/')
        await asyncio.sleep(2)
        logged_in = not await check_login_required(page)
        return page, logged_in
    except Exception as e:
        logger.error(f"reopen_kaoshibao失败: {e}")
        return None, False


def get_chrome_status():
    """检查Chrome调试端口状态"""
    return is_port_open(CDP_PORT)


async def open_cdp_page(playwright):
    """通过CDP连接Chrome并打开新标签页，返回 (browser, page)，失败返回 (None, None)"""
    try:
        browser = await playwright.chromium.connect_over_cdp(CDP_URL, timeout=15000)
        if not browser.contexts:
            return None, None
        context = browser.contexts[0]
        page = await context.new_page()
        return browser, page
    except Exception as e:
        logger.error(f"连接Chrome失败: {e}")
        return None, None


async def navigate_with_retry(page, url, retries=3, timeout=30000) -> bool:
    """带重试的页面导航"""
    for attempt in range(retries):
        try:
            resp = await page.goto(url, wait_until='domcontentloaded', timeout=timeout)
            try:
                await page.wait_for_load_state('networkidle', timeout=10000)
            except Exception:
                pass
            return resp
        except Exception as e:
            if attempt < retries - 1:
                await asyncio.sleep(2)
            else:
                raise
    return None


async def close_popups(page) -> None:
    """关闭弹窗（覆盖多种常见弹窗类型）"""
    selectors = [
        '.el-dialog__headerbtn',
        '.membership-dialog .close',
        '.el-dialog__close',
        '.el-message-box__headerbtn',
        '.el-notification__closeBtn',
        '.el-alert__closebtn',
        '[class*="modal"] .close',
        '[class*="popup"] .close',
        '.v-modal',
        '.el-overlay .el-dialog__headerbtn',
    ]
    for _ in range(3):
        try:
            closed = await page.evaluate('''(sels) => {
                let c = false;
                sels.forEach(sel => {
                    document.querySelectorAll(sel).forEach(b => {
                        if (b.offsetParent !== null) { b.click(); c = true; }
                    });
                });
                // 也尝试按ESC关闭
                if (!c) {
                    const overlays = document.querySelectorAll('.el-overlay, .v-modal, [class*="modal-mask"]');
                    if (overlays.length > 0) { document.dispatchEvent(new KeyboardEvent('keydown', {key: 'Escape'})); c = true; }
                }
                return c;
            }''', selectors)
            if not closed:
                break
            await asyncio.sleep(0.5)
        except Exception:
            break


async def ocr_element_text(page, selector) -> str:
    """截取DOM元素并用OCR识别真实文本（绕过字体反爬）"""
    try:
        el = await page.query_selector(selector)
        if not el:
            return ''
        box = await el.bounding_box()
        if not box or box['width'] < 5:
            return ''
        img_path = os.path.join(os.environ.get('TEMP', '/tmp'), 'ocr_temp.png')
        await page.screenshot(path=img_path, clip={
            'x': box['x'] - 2, 'y': box['y'] - 2,
            'width': box['width'] + 4, 'height': box['height'] + 4,
        })
        return await _ocr_file(img_path)
    except Exception:
        return ''


async def ocr_region_text(page, x, y, w, h) -> str:
    """截取指定区域并用OCR识别"""
    try:
        img_path = os.path.join(os.environ.get('TEMP', '/tmp'), 'ocr_temp.png')
        await page.screenshot(path=img_path, clip={
            'x': x - 2, 'y': y - 2, 'width': w + 4, 'height': h + 4,
        })
        return await _ocr_file(img_path)
    except Exception:
        return ''


async def _ocr_file(img_path) -> str:
    """对图片文件执行OCR，自动加padding+放大提升识别率"""
    from PIL import Image, ImageOps, ImageEnhance
    ocr = _get_ocr()
    img = Image.open(img_path)
    # 放大3倍 + 白色边框 + 增强对比度
    scaled = img.resize((img.width * 3, img.height * 3), Image.LANCZOS)
    enhanced = ImageEnhance.Contrast(scaled).enhance(1.5)
    padded = ImageOps.expand(enhanced, border=20, fill='white')
    padded_path = img_path.replace('.png', '_padded.png')
    padded.save(padded_path)
    result, _ = ocr(padded_path)
    if not result:
        # 兜底：灰度+二值化
        gray = padded.convert('L')
        bw = gray.point(lambda x: 0 if x < 140 else 255)
        bw_path = img_path.replace('.png', '_bw.png')
        bw.save(bw_path)
        result, _ = ocr(bw_path)
    if not result:
        return ''
    return ''.join(line[1] for line in result).strip()


async def click_by_text(page, text) -> bool:
    """按文本点击元素，返回是否成功点击"""
    # 方法1: Playwright text selector
    try:
        btn = await page.query_selector(f'text={text}')
        if btn and await btn.is_visible():
            await btn.click(timeout=5000)
            return True
    except Exception:
        pass
    # 方法2: JS精确匹配叶子节点
    clicked = await page.evaluate(f'''() => {{
        for (const el of document.querySelectorAll('*')) {{
            if (el.innerText?.trim() === '{text}' && el.offsetParent !== null && el.children.length === 0) {{
                el.click(); return true;
            }}
        }}
        return false;
    }}''')
    if clicked:
        return True
    # 方法3: JS模糊匹配（包含目标文本的可点击元素）
    return await page.evaluate(f'''() => {{
        const targets = document.querySelectorAll('a, button, [onclick], [role="button"]');
        for (const el of targets) {{
            if (el.innerText?.includes('{text}') && el.offsetParent !== null) {{
                el.click(); return true;
            }}
        }}
        // 最后尝试：任何包含文本的可见元素
        for (const el of document.querySelectorAll('*')) {{
            if (el.innerText?.includes('{text}') && el.offsetParent !== null && el.offsetHeight > 0) {{
                el.click(); return true;
            }}
        }}
        return false;
    }}''')


async def enable_recite_mode(page) -> bool:
    """开启背题模式，返回是否成功开启"""
    enabled = await page.evaluate('''() => {
        // 考试宝的背题模式开关可能在多个位置
        // 方法1: .set p 中的 el-switch
        const ps = document.querySelectorAll('.set p, .set-item, .setting-item');
        for (const p of ps) {
            if (p.innerText?.includes('背题模式')) {
                const sw = p.querySelector('.el-switch');
                if (sw) {
                    if (!sw.classList.contains('is-checked')) sw.click();
                    return true;
                }
            }
        }
        // 方法2: 直接查找包含"背题"的可点击元素
        const allEls = document.querySelectorAll('.el-switch, [role="switch"]');
        for (const el of allEls) {
            const parent = el.closest('p, div, label');
            if (parent && parent.innerText?.includes('背题')) {
                if (!el.classList.contains('is-checked')) el.click();
                return true;
            }
        }
        return false;
    }''')
    await asyncio.sleep(2)
    return enabled


async def wait_for_question(page, timeout=10) -> bool:
    """等待题目元素出现（包含modal中的题目）"""
    selectors = [
        '.qusetion-title',
        '.topic-top .question',
        '.question-content',
        '.topic-tit',
        '[class*="exam_modal"]',
    ]
    for sel in selectors:
        try:
            await page.wait_for_selector(sel, timeout=timeout * 1000, state='visible')
            return True
        except Exception:
            continue
    return False


async def get_current_question_dom(page) -> dict | None:
    """获取当前题目数据（DOM直接读取，速度快但有乱码）

    同时收集题干和选项里的 <img> 图片 URL（考试宝题目配图挂在
    .qusetion-title / .option 内部，src 指向 up.zaixiankaoshi.com）。
    """
    data = await page.evaluate('''() => {
        const r = { question: '', options: [], answer: '', analysis: '', type: '',
                    images: [], option_images: [] };
        // 过滤站点图标（icon_vip / nandu_ed / arrow 等），只留真正的题目配图
        const NOISE = ['icon','logo','nandu','star','arrow','vip','avatar',
                       'sprite','loading','placeholder','qrcode','erweima',
                       '/pc/','/common/','/static/'];
        const okImg = (src, w, h) => {
            if (!src || src.startsWith('data:')) return false;
            const low = src.toLowerCase();
            if (NOISE.some(x => low.includes(x))) return false;
            if (!(low.includes('up.zaixiankaoshi.com') || low.includes('image') || low.includes('/upload'))) return false;
            if (w && h && (w < 50 || h < 50)) return false;
            return true;
        };
        const typeEl = document.querySelector('.topic-type');
        if (typeEl) r.type = typeEl.innerText.trim();
        const qEl = document.querySelector('.qusetion-title');
        if (qEl) {
            r.question = qEl.innerText.trim();
            qEl.querySelectorAll('img').forEach(im => {
                const src = im.getAttribute('src') || im.src || '';
                const w = im.naturalWidth || 0, h = im.naturalHeight || 0;
                if (okImg(src, w, h)) r.images.push({src: src, w: w, h: h});
            });
        }
        document.querySelectorAll('.option').forEach((opt, idx) => {
            const t = opt.innerText.trim();
            if (t) r.options.push(t);
            const oimgs = [];
            opt.querySelectorAll('img').forEach(im => {
                const src = im.getAttribute('src') || im.src || '';
                const w = im.naturalWidth || 0, h = im.naturalHeight || 0;
                if (okImg(src, w, h)) oimgs.push({src: src, w: w, h: h});
            });
            if (oimgs.length) r.option_images.push({index: idx, images: oimgs});
        });
        const rights = document.querySelectorAll('.option.right');
        if (rights.length > 0) {
            const answers = [];
            rights.forEach(opt => {
                const icon = opt.querySelector('.before-icon');
                if (icon) answers.push(icon.innerText.trim());
            });
            r.answer = answers.join(',');
        }
        const analEl = document.querySelector('.answer-box-detail');
        if (analEl) {
            r.analysis = analEl.innerText.trim();
            const aimgs = [];
            analEl.querySelectorAll('img').forEach(im => {
                const src = im.getAttribute('src') || im.src || '';
                const w = im.naturalWidth || 0, h = im.naturalHeight || 0;
                if (okImg(src, w, h)) aimgs.push({src: src, w: w, h: h});
            });
            if (aimgs.length) r.analysis_images = aimgs;
        }
        return r;
    }''')
    return data


async def get_current_question_hybrid(page) -> dict | None:
    """混合模式：DOM读结构+题目，OCR仅当DOM题目为空时兜底（速度快，质量好）"""
    dom_data = await get_current_question_dom(page)

    # 题目优先用DOM读取（快、无OCR误差），DOM为空时再用OCR兜底
    question = dom_data.get('question', '').strip()
    if not question or len(question) < 5:
        from PIL import Image, ImageOps, ImageEnhance
        try:
            qEl = await page.query_selector('.qusetion-title')
            if qEl:
                box = await qEl.bounding_box()
                if box and box['width'] > 10:
                    img_path = os.path.join(os.environ.get('TEMP', '/tmp'), 'hybrid_q.png')
                    await page.screenshot(path=img_path, clip={
                        'x': box['x'] - 5, 'y': box['y'] - 5,
                        'width': box['width'] + 10, 'height': box['height'] + 10,
                    })
                    ocr_q = await _ocr_file(img_path)
                    if ocr_q and len(ocr_q) >= 5:
                        question = ocr_q
        except Exception:
            pass
        if not question or len(question) < 5:
            question = dom_data.get('question', '')

    dom_data['question'] = question
    return dom_data


async def _extract_answer_and_type(page) -> tuple[str, str]:
    """从DOM提取正确答案和题型"""
    answer = await page.evaluate('''() => {
        const rights = document.querySelectorAll('.option.right');
        if (rights.length > 0) {
            const answers = [];
            rights.forEach(opt => {
                const icon = opt.querySelector('.before-icon');
                if (icon) answers.push(icon.innerText.trim());
            });
            return answers.join(',');
        }
        const ansEl = document.querySelector('.right-ans');
        if (ansEl) return ansEl.innerText.replace('正确答案', '').trim();
        return '';
    }''')
    qtype = await page.evaluate('''() => {
        const el = document.querySelector('.topic-type');
        return el ? el.innerText.trim() : '';
    }''')
    return answer, qtype


def _enhance_and_ocr(img_path: str, scale: int = 3, contrast: float = 1.5, border: int = 25) -> str:
    """图片放大+增强对比度+padding后OCR，返回OCR路径"""
    from PIL import Image, ImageOps, ImageEnhance
    img = Image.open(img_path)
    scaled = img.resize((img.width * scale, img.height * scale), Image.LANCZOS)
    enhanced = ImageEnhance.Contrast(scaled).enhance(contrast)
    padded = ImageOps.expand(enhanced, border=border, fill='white')
    ocr_path = img_path.replace('.png', '_final.png')
    padded.save(ocr_path)
    return ocr_path


async def _ocr_analysis_section(page) -> str:
    """OCR解析区域文本"""
    try:
        anal_el = await page.query_selector('.answer-box-detail')
        if anal_el:
            anal_box = await anal_el.bounding_box()
            if anal_box and anal_box['width'] > 20:
                img_path = os.path.join(os.environ.get('TEMP', '/tmp'), 'anal_ocr.png')
                await page.screenshot(path=img_path, clip={
                    'x': anal_box['x'] - 5, 'y': anal_box['y'] - 5,
                    'width': anal_box['width'] + 10, 'height': anal_box['height'] + 10,
                })
                ocr_path = _enhance_and_ocr(img_path, border=20)
                ocr = _get_ocr()
                result, _ = ocr(ocr_path)
                if result:
                    return ' '.join(line[1].strip() for line in result if line[1].strip())
    except Exception:
        pass
    return ''


def _is_question_image(src: str, w: int = 0, h: int = 0) -> bool:
    """判断一个 <img> 是不是真正的题目配图。

    排除：站点图标(icon/logo)、难度星级、箭头、VIP 标记、svg-data、
    以及过小的装饰元素（题目配图一般 ≥ 60px）。
    """
    if not src or src.startswith('data:'):
        return False
    low = src.lower()
    noise_words = (
        'icon', 'logo', 'nandu', 'star', 'arrow', 'vip', 'avatar',
        'sprite', 'loading', 'placeholder', 'qrcode', 'erweima',
        '/pc/', '/common/', '/static/',
    )
    if any(x in low for x in noise_words):
        return False
    # 题目配图域名
    if 'up.zaixiankaoshi.com' not in low and 'image' not in low and '/upload' not in low:
        return False
    # 尺寸过小的一律排除（装饰性）
    if w and h and (w < 50 or h < 50):
        return False
    return True


async def _extract_question_images(page) -> dict:
    """只抓图片 URL，不关心文字。OCR 模式也必须调用，否则配图全丢。

    关键：必须过滤站点图标。考试宝解析区里塞了大量 icon_vip / nandu_ed 之类的
    装饰图，不过滤会导致每道题都"带 8 张图"。
    """
    try:
        return await page.evaluate('''() => {
            const out = {images: [], option_images: [], analysis_images: []};
            const NOISE = ['icon','logo','nandu','star','arrow','vip','avatar',
                           'sprite','loading','placeholder','qrcode','erweima',
                           '/pc/','/common/','/static/'];
            const ok = (src, w, h) => {
                if (!src || src.startsWith('data:')) return false;
                const low = src.toLowerCase();
                if (NOISE.some(x => low.includes(x))) return false;
                if (!(low.includes('up.zaixiankaoshi.com') || low.includes('image') || low.includes('/upload'))) return false;
                if (w && h && (w < 50 || h < 50)) return false;
                return true;
            };
            const push = (arr, im) => {
                const src = im.getAttribute('src') || im.src || '';
                const w = im.naturalWidth || 0, h = im.naturalHeight || 0;
                if (ok(src, w, h)) arr.push({src: src, w: w, h: h});
            };
            const qEl = document.querySelector('.qusetion-title');
            if (qEl) qEl.querySelectorAll('img').forEach(im => push(out.images, im));
            document.querySelectorAll('.option').forEach((opt, idx) => {
                const buf = [];
                opt.querySelectorAll('img').forEach(im => push(buf, im));
                if (buf.length) out.option_images.push({index: idx, images: buf});
            });
            const analEl = document.querySelector('.answer-box-detail');
            if (analEl) analEl.querySelectorAll('img').forEach(im => push(out.analysis_images, im));
            return out;
        }''')
    except Exception as e:
        logger.warning(f"提取题目图片失败: {e}")
        return {'images': [], 'option_images': [], 'analysis_images': []}


async def get_current_question(page, mode='ocr') -> dict | None:
    """获取当前题目数据"""
    if mode == 'dom':
        return await get_current_question_dom(page)
    elif mode == 'hybrid':
        return await get_current_question_hybrid(page)
    from PIL import Image, ImageOps

    # 正确答案和题型（DOM结构读取，不受字体影响）
    answer, qtype = await _extract_answer_and_type(page)

    # 截取整个modal区域 + OCR
    question = ''
    options = []
    try:
        modal = await page.query_selector('[class*="exam_modal"]')
        if not modal:
            modal = await page.query_selector('.topic-top')
        if modal:
            box = await modal.bounding_box()
            if box and box['width'] > 50:
                img_path = os.path.join(os.environ.get('TEMP', '/tmp'), 'modal_ocr.png')
                await page.screenshot(path=img_path, clip={
                    'x': box['x'] - 10, 'y': box['y'] - 10,
                    'width': box['width'] + 20, 'height': box['height'] + 20,
                })
                # 放大3倍+增强对比度+加padding
                ocr_path = _enhance_and_ocr(img_path)

                ocr = _get_ocr()
                result, _ = ocr(ocr_path)
                if result:
                    all_lines = [line[1].strip() for line in result if line[1].strip()]
                    # 清理噪音行
                    clean_lines = []
                    for line in all_lines:
                        # 跳过URL、域名、按钮、噪音
                        if any(skip in line.lower() for skip in [
                            'kaoshibao.com', 'www.', 'http', '下一题', '上一题',
                            '收藏', '笔记', '举报', '斩题', 'aost', '请输入',
                            '保存答案', '原解析', 'AI解析', 'VIP', '深度解题',
                            '本解析由AI', '获取深度', '提供当日', 'hibao.', 'nibao.',
                            'ww.kaos', 'ww.k', '.kaos', 'ao.com', 'no.com',
                            'com', '≤', '☆', '但', '↑',
                        ]):
                            continue
                        # 跳过纯噪音短行
                        if len(line) < 3:
                            continue
                        # 跳过纯域名片段
                        if re.match(r'^[a-z.]+$', line.lower()):
                            continue
                        clean_lines.append(line)

                    # 解析题目和选项
                    # 方法：从DOM获取选项数量，OCR文本中题目之后的内容按顺序分配给选项
                    dom_info = await page.evaluate('''() => {
                        const opts = document.querySelectorAll('.option');
                        return {
                            optCount: opts.length,
                            hasQuestion: !!document.querySelector('.qusetion-title'),
                        };
                    }''')
                    num_opts = dom_info.get('optCount', 4)

                    # 分离题目和选项文本
                    # 找题型行之后的内容
                    start = 0
                    for i, line in enumerate(clean_lines):
                        if re.match(r'^(单选题|多选题|判断题|填空题|简答题)$', line):
                            start = i + 1
                            break

                    # 从start开始，找题号行或题目行
                    q_text = ''
                    opt_start = start

                    # 找题号行
                    q_num_idx = -1
                    for i in range(start, len(clean_lines)):
                        line = clean_lines[i]
                        m = re.search(r'\d+/\d+[、,.。]\s*', line)
                        if m:
                            q_text = line[m.end():].strip()
                            prefix = line[:m.start()].strip()
                            if prefix and not re.match(r'^[A-F]\s*[.、]', prefix):
                                q_text = prefix + ' ' + q_text if q_text else prefix
                            q_num_idx = i
                            break

                    # 题号之后，继续读取题目文本（跳过题型行），直到遇到选项或问号
                    if q_num_idx >= 0:
                        for i in range(q_num_idx + 1, len(clean_lines)):
                            line = clean_lines[i]
                            # 跳过题型行
                            if re.match(r'^(单选题|多选题|判断题|填空题|简答题)$', line):
                                continue
                            # 遇到选项行停止
                            if re.match(r'^[A-F]\s*[.、]', line):
                                opt_start = i
                                break
                            # 拼接到题目
                            q_text += (' ' if q_text else '') + line
                            opt_start = i + 1
                            # 如果题目已有问号，可能是完整题目
                            if q_text.endswith('?') or q_text.endswith('？'):
                                opt_start = i + 1
                                break

                    # 没找到题号行：第一个非选项行是题目
                    if not q_text:
                        for i in range(start, len(clean_lines)):
                            line = clean_lines[i]
                            if re.match(r'^[A-F]\s*[.、]', line):
                                opt_start = i
                                break
                            q_text += (' ' if q_text else '') + line
                            opt_start = i + 1

                    # 去掉残留域名
                    q_text = re.sub(r'\s*[a-z]{2,8}\.(com|cn|c)\b.*$', '', q_text, flags=re.IGNORECASE).strip()

                    # 如果题目为空或太短，用DOM兜底
                    if len(q_text) < 5:
                        dom_q = await page.evaluate('() => document.querySelector(".qusetion-title")?.innerText?.trim() || ""')
                        if dom_q:
                            q_text = re.sub(r'[a-z]{2,8}\.(com|cn|c)\b.*', '', dom_q, flags=re.IGNORECASE).strip()

                    # 用问号分割题目和选项（题目通常以?或？结尾）
                    question = q_text
                    if '?' in q_text or '？' in q_text:
                        # 找最后一个问号
                        last_q = max(q_text.rfind('?'), q_text.rfind('？'))
                        # 只有当问号后面有足够内容时才分割（至少10字符）
                        if last_q > 10 and last_q < len(q_text) - 10:
                            question = q_text[:last_q + 1].strip()
                            # 问号后面的内容作为选项
                            extra_opts = q_text[last_q + 1:].strip()
                            if extra_opts:
                                for opt in re.split(r'\s+', extra_opts):
                                    opt = opt.strip()
                                    if len(opt) > 1 and opt not in options:
                                        options.append(opt)

                    # 选项：从opt_start开始，过滤噪音，取num_opts个有效选项
                    content_lines = clean_lines[opt_start:]
                    skip_words = ['下一题', '上一题', '收藏', '笔记', '举报', '斩题', '保存', '请输入',
                                  '下一', '上一', '→', '☆', '≤', '但', '↑', '答案', '正确答案']
                    for line in content_lines[:num_opts * 5]:
                        cleaned = line.strip()
                        # 去掉域名噪音
                        cleaned = re.sub(r'\s*[a-z]{2,8}\.(com|cn|c)\b.*$', '', cleaned, flags=re.IGNORECASE)
                        # 去掉按钮/导航文字
                        if any(sw in cleaned for sw in skip_words):
                            continue
                        # 去掉纯噪音（太短、纯符号、纯数字）
                        if len(cleaned) < 3:
                            continue
                        if re.match(r'^[a-z0-9.]+$', cleaned.lower()):
                            continue
                        cleaned = cleaned.strip()
                        if cleaned:
                            options.append(cleaned)
                        if len(options) >= num_opts:
                            break

                    # 如果第一个"选项"看起来像题目（以?或"哪"结尾），移到题目
                    if options and (not question or len(question) < 10):
                        first = options[0]
                        if first.endswith('?') or first.endswith('？') or '哪' in first[-5:] or '？' in first:
                            question = first
                            options = options[1:]
    except Exception:
        pass

    # 解析（从分析区域OCR或DOM）
    analysis = await _ocr_analysis_section(page)

    # 题目配图：OCR 只认文字，图片必须单独从 DOM 抓 URL
    # （考试宝把图挂在 .qusetion-title / .option / .answer-box-detail 的 <img> 里）
    fig = await _extract_question_images(page)

    return {
        'type': qtype,
        'question': question,
        'options': options,
        'answer': answer,
        'analysis': analysis,
        'images': fig.get('images', []),
        'option_images': fig.get('option_images', []),
        'analysis_images': fig.get('analysis_images', []),
    }


async def trigger_answer(page, question_type) -> bool:
    """根据题型触发答案显示，带重试"""
    qtype = question_type or ''
    for attempt in range(2):
        # 检查答案是否已经显示（考试宝在背题模式下可能直接显示）
        has_answer = await page.evaluate('''() => {
            if (document.querySelector('.option.right')) return true;
            if (document.querySelector('.right-ans, .correct-answer')) return true;
            return false;
        }''')
        if has_answer:
            return

        if '填空' in qtype or '简答' in qtype or '问答' in qtype:
            clicked = await page.evaluate('''() => {
                const btns = document.querySelectorAll('button, .btn, a, [role="button"]');
                for (const el of btns) {
                    const txt = el.innerText?.trim();
                    if ((txt === '查看答案' || txt === '确定' || txt === '提交' || txt === '查看解析') && el.offsetParent !== null) {
                        el.click();
                        return true;
                    }
                }
                return false;
            }''')
            if clicked:
                await asyncio.sleep(2)
        else:
            # 点击第一个选项触发答案显示
            await page.evaluate('''() => {
                const opts = document.querySelectorAll('.option, .option-item');
                for (const opt of opts) {
                    if (opt.offsetParent !== null) { opt.click(); break; }
                }
            }''')
            await asyncio.sleep(1.5)

        # 再次检查
        has_answer = await page.evaluate('''() => {
            if (document.querySelector('.option.right')) return true;
            if (document.querySelector('.right-ans, .correct-answer')) return true;
            return false;
        }''')
        if has_answer:
            return
        await asyncio.sleep(1)


async def click_next_question(page) -> bool:
    """点击下一题，多种方式尝试，返回是否成功"""
    # 方法1: 考试宝专用 — .next-preve 中的"下一题"按钮
    clicked = await page.evaluate('''() => {
        // 考试宝的下一题按钮在 .next-preve 容器中
        const btns = document.querySelectorAll('.next-preve button, .next-preve .el-button');
        for (const b of btns) {
            if (b.innerText?.includes('下一题') && b.offsetParent !== null) {
                b.click();
                return true;
            }
        }
        // 通用：查找包含"下一题"的按钮
        const allBtns = document.querySelectorAll('button, a, [role="button"]');
        for (const b of allBtns) {
            const txt = b.innerText?.trim();
            if (txt && txt.includes('下一题') && b.offsetParent !== null) {
                b.click();
                return true;
            }
        }
        return false;
    }''')
    if clicked:
        await asyncio.sleep(2)
        return True

    # 方法2: Playwright text selector
    next_selectors = [
        'text=下一题',
        'button:has-text("下一题")',
        '.next-question',
    ]
    for sel in next_selectors:
        try:
            btn = await page.query_selector(sel)
            if btn and await btn.is_visible():
                await btn.click(timeout=3000)
                await asyncio.sleep(2)
                return True
        except Exception:
            continue

    # 方法3: 键盘右箭头
    try:
        await page.keyboard.press('ArrowRight')
        await asyncio.sleep(2)
        return True
    except Exception:
        pass

    return False


async def inject_cookies(context, cookies) -> None:
    """将cookies注入浏览器上下文"""
    if not cookies:
        return False
    try:
        playwright_cookies = []
        for c in cookies:
            if not c.get('name'):
                continue
            cookie = {
                'name': c['name'],
                'value': c.get('value', ''),
                'domain': c.get('domain', '.kaoshibao.com'),
                'path': c.get('path', '/'),
            }
            if c.get('expires'):
                cookie['expires'] = c['expires']
            if c.get('secure'):
                cookie['secure'] = True
            if c.get('httpOnly'):
                cookie['httpOnly'] = True
            if c.get('sameSite'):
                ss = c['sameSite']
                if ss in ('Strict', 'Lax', 'None'):
                    cookie['sameSite'] = ss
            playwright_cookies.append(cookie)
        if not playwright_cookies:
            return False
        await context.add_cookies(playwright_cookies)
        return True
    except Exception as e:
        logger.error(f"注入cookies失败: {e}")
        return False


async def dump_page_info(page, log_fn=None) -> None:
    """输出页面诊断信息"""
    def log(msg):
        if log_fn:
            log_fn(msg)
        try:
            logger.info(msg)
        except UnicodeEncodeError:
            logger.info(str(msg).encode('utf-8', errors='replace').decode('utf-8', errors='replace'))

    info = await page.evaluate('''() => {
        const result = {};
        // 页面标题
        result.title = document.title;
        // 所有可见链接
        result.links = [];
        document.querySelectorAll('a[href]').forEach(a => {
            if (a.offsetParent !== null || a.offsetHeight > 0) {
                const text = a.innerText?.trim().substring(0, 50);
                const href = a.href;
                if (text) result.links.push({text, href});
            }
        });
        // 所有可见按钮
        result.buttons = [];
        document.querySelectorAll('button, [role="button"], .btn, [onclick]').forEach(b => {
            if (b.offsetParent !== null || b.offsetHeight > 0) {
                const text = b.innerText?.trim().substring(0, 50);
                if (text) result.buttons.push(text);
            }
        });
        // 页面主要文本内容（前500字符）
        const body = document.body?.innerText || '';
        result.bodyText = body.substring(0, 500);
        // Vue/Element组件
        result.vueComponents = document.querySelectorAll('[class*="el-"]').length;
        return result;
    }''')

    log(f"  页面标题: {info.get('title', 'N/A')}")
    log(f"  Vue组件数: {info.get('vueComponents', 0)}")
    links = info.get('links', [])
    if links:
        log(f"  可见链接 ({len(links)}个):")
        for l in links[:10]:
            log(f"    [{l['text']}] -> {l['href']}")
    buttons = info.get('buttons', [])
    if buttons:
        log(f"  可见按钮 ({len(buttons)}个):")
        for b in buttons[:10]:
            log(f"    [{b}]")
    body_text = info.get('bodyText', '')
    if body_text:
        log(f"  页面文本(前200字): {body_text[:200]}")
    return info


async def check_login_required(page) -> bool:
    """检查是否需要登录（URL或DOM双重检测）"""
    if 'login' in page.url.lower():
        return True
    return await page.evaluate('''() => {
        // 检查登录表单、登录弹窗、二维码登录等
        if (document.querySelector('.login-form, .login-box, .qrcode-login, #loginForm')) return true;
        if (document.querySelector('[class*="login"]')) {
            const el = document.querySelector('[class*="login"]');
            if (el && el.offsetHeight > 100) return true;
        }
        return false;
    }''')


async def wait_for_login(page, log_fn=None, stop_check=None, timeout=300) -> bool:
    """等待用户在Chrome中完成登录，返回 True=已登录, False=超时/取消"""
    def log(msg):
        if log_fn:
            log_fn(msg)
        logger.info(msg)

    log("需要登录考试宝，请在Chrome中扫码登录...")
    start = time.time()
    while time.time() - start < timeout:
        if stop_check and stop_check():
            log("用户取消")
            return False
        await asyncio.sleep(3)
        try:
            await page.reload(wait_until='domcontentloaded')
            try:
                await page.wait_for_load_state('networkidle', timeout=10000)
            except Exception:
                pass
            await asyncio.sleep(1)
        except Exception:
            await asyncio.sleep(2)
            continue
        if not await check_login_required(page):
            log("登录成功!")
            return True
    log(f"等待登录超时（{timeout}秒）")
    return False


async def find_practice_url(page, paper_id, log_fn=None) -> str | None:
    """尝试找到"顺序练习"的实际链接"""
    def log(msg):
        if log_fn:
            log_fn(msg)
        logger.info(msg)

    # 方法1: 从页面链接中提取（过滤掉javascript:伪链接）
    practice_href = await page.evaluate('''() => {
        const links = document.querySelectorAll('a[href]');
        for (const a of links) {
            const text = a.innerText?.trim();
            const href = a.href || '';
            if (text && text.includes('顺序练习') && href.startsWith('http')) {
                return href;
            }
        }
        for (const a of links) {
            const text = a.innerText?.trim();
            const href = a.href || '';
            if (text && (text.includes('练习') || text.includes('开始')) && href.includes('/ti/') && href.startsWith('http')) {
                return href;
            }
        }
        return '';
    }''')
    if practice_href:
        log(f"  找到练习链接: {practice_href}")
        return practice_href

    # 方法2: 尝试常见URL模式
    candidates = [
        f'https://www.kaoshibao.com/ti/{paper_id}/practice/',
        f'https://www.kaoshibao.com/ti/{paper_id}/do/',
        f'https://www.kaoshibao.com/ti/{paper_id}/exam/',
        f'https://www.kaoshibao.com/ti/{paper_id}/?mode=practice',
    ]
    for url in candidates:
        log(f"  尝试直接访问: {url}")
        try:
            resp = await navigate_with_retry(page, url, retries=1, timeout=15000)
            await asyncio.sleep(2)
            if resp and resp.status == 200 and not await check_login_required(page):
                # 检查页面是否有题目相关元素
                has_content = await page.evaluate('''() => {
                    return document.querySelector('.topic-top, .question, .option-item, .set') !== null;
                }''')
                if has_content:
                    log(f"  成功! URL: {page.url}")
                    return page.url
        except Exception:
            continue

    return None


async def save_screenshot(page, name, log_fn=None) -> None:
    """保存截图用于调试"""
    try:
        path = OUTPUT_DIR / f"debug_{name}_{int(time.time())}.png"
        await page.screenshot(path=str(path), full_page=False)
        if log_fn:
            log_fn(f"截图已保存: {path.name}")
    except Exception:
        pass


async def _init_recite_mode(page, log_fn=None) -> bool:
    """开启背题模式并等待题目元素出现，返回是否成功"""
    def log(msg):
        if log_fn:
            log_fn(msg)
        logger.info(msg)
    log("开启背题模式...")
    recite_ok = await enable_recite_mode(page)
    if recite_ok:
        log("背题模式已开启")
    else:
        log("未找到背题模式开关，继续尝试...")
    await close_popups(page)
    found = await wait_for_question(page, timeout=10)
    if not found:
        log("题目元素未找到，等待5秒后重试...")
        await asyncio.sleep(5)
        found = await wait_for_question(page, timeout=10)
        if not found:
            log("题目元素仍未找到，输出诊断:")
            await dump_page_info(page, log_fn)
            await save_screenshot(page, 'no_question', log_fn)
            return False
    return True


async def _navigate_to_practice(page, paper_id, start_sequence, log_fn, stop_check, new_page_holder=None) -> str:
    """导航到练习页面，返回状态：ok/auth_error
    new_page_holder: dict，若练习在新标签页打开，将 {'page': new_page} 放入其中"""

    def log(msg):
        if log_fn:
            log_fn(msg)
        logger.info(msg)

    log(f"打开题库页面 {paper_id}...")
    await navigate_with_retry(page, f'https://www.kaoshibao.com/ti/{paper_id}/')
    await asyncio.sleep(2)
    await close_popups(page)

    if await check_login_required(page):
        logged_in = await wait_for_login(page, log_fn, stop_check)
        if not logged_in:
            return 'auth_error'
        await navigate_with_retry(page, f'https://www.kaoshibao.com/ti/{paper_id}/')
        await asyncio.sleep(2)
        await close_popups(page)
        if await check_login_required(page):
            log("登录后仍无法访问题库")
            return 'auth_error'

    if start_sequence > 0:
        practice_url = (
            f'https://www.kaoshibao.com/online/?paperId={paper_id}'
            f'&practice=&modal=1&type=qtype_random&is_recite=&qtype='
            f'&text=顺序练习&sequence={start_sequence}'
            f'&is_collect=1&is_vip_paper=0&gid=&is_buy=1'
            f'&qtype_random_index={start_sequence}'
        )
        log(f"分段模式：跳到第 {start_sequence + 1} 题...")
        await navigate_with_retry(page, practice_url)
        await asyncio.sleep(3)
        await close_popups(page)
    else:
        practice_url = (
            f'https://www.kaoshibao.com/online/?paperId={paper_id}'
            f'&practice=&modal=1&type=qtype_random&is_recite=&qtype='
            f'&text=顺序练习&sequence=0'
            f'&is_collect=1&is_vip_paper=0&gid=&is_buy=1'
            f'&qtype_random_index=0'
        )
        log("直接访问练习页面...")
        await navigate_with_retry(page, practice_url)
        await asyncio.sleep(3)
        await close_popups(page)
        if not await wait_for_question(page, timeout=8):
            # 直接访问未加载题目，回退到点击顺序练习（含新标签页捕获）
            log("直接访问未加载题目，点击顺序练习...")
            paper_url = page.url
            captured_pages = []
            if new_page_holder is not None:
                async def _on_new(pg):
                    try:
                        captured_pages.append(pg)
                    except Exception:
                        pass
                page.context.on('page', _on_new)
            clicked = await page.evaluate('''() => {
                const allItems = document.querySelectorAll('a.item, a[href="javascript:;"]');
                for (const a of allItems) {
                    if (a.innerText?.includes('顺序练习')) {
                        ['mousedown', 'mouseup', 'click'].forEach(type => {
                            a.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
                        });
                        return true;
                    }
                }
                return false;
            }''')
            log(f"  点击结果: {clicked}")
            new_page = None
            for wait_i in range(10):
                await asyncio.sleep(1)
                if captured_pages:
                    new_page = captured_pages[0]
                    break
                if page.url != paper_url:
                    break
            if new_page is not None:
                log(f"  练习在新标签页打开: {new_page.url}")
                if new_page_holder is not None:
                    new_page_holder['page'] = new_page
                try:
                    await page.close()
                except Exception:
                    pass
                return 'ok'
            await close_popups(page)
            if page.url == paper_url:
                log("  URL未变化，尝试直接访问练习页面...")
                practice_url2 = await find_practice_url(page, paper_id, log_fn)
                if practice_url2:
                    await navigate_with_retry(page, practice_url2)
                    await asyncio.sleep(2)
                    await close_popups(page)
        if await check_login_required(page):
            log("进入练习页后需要登录...")
            logged_in = await wait_for_login(page, log_fn, stop_check)
            if not logged_in:
                return 'auth_error'
            await navigate_with_retry(page, f'https://www.kaoshibao.com/ti/{paper_id}/')
            await asyncio.sleep(2)
            await close_popups(page)
            await click_by_text(page, '顺序练习')
            await asyncio.sleep(5)
            await close_popups(page)
            if await check_login_required(page):
                log("登录后仍无法进入练习页面")
                return 'auth_error'

    log(f"当前URL: {page.url}")
    return 'ok'


def _is_valid_question_text(q_text: str) -> bool:
    """判断题目文本是否有效：至少10字符，包含中文，不全是噪音，且非页面提示/垃圾内容"""
    has_chinese = bool(re.search(r'[一-鿿]', q_text))
    if not (len(q_text) >= 10 and has_chinese):
        return False
    if re.match(r'^[a-z0-9.\s]+$', q_text.lower()):
        return False
    # 过滤页面提示、AI解析说明等非题目内容
    noise_markers = [
        '请先点击上方设置按钮', '关闭AI解析', '随时生成的文字',
        '本解析由AI', '原解析', 'AI解析', '获取深度解题', '深度解题',
        '上方设置', '点击上方',
    ]
    if any(m in q_text for m in noise_markers):
        return False
    return True


def _add_question_if_new(data: dict[str, Any], questions: list, existing_keys: set,
                         paper_id: str, mode: str, progress_fn) -> bool:
    """去重并添加题目，返回是否为新题"""
    q_key = data['question'][:80]
    if q_key in existing_keys:
        return False
    questions.append({
        '序号': len(questions) + 1,
        '题目类型': data.get('type', ''),
        '题目': data.get('question', ''),
        '选项': data.get('options', []),
        '正确答案': data.get('answer', ''),
        '解析': data.get('analysis', ''),
        '图片': data.get('images', []),
        '选项图片': data.get('option_images', []),
    })
    if data.get('analysis_images'):
        questions[-1]['解析图片'] = data['analysis_images']
    existing_keys.add(q_key)
    if progress_fn:
        has_ans = sum(1 for q in questions if q.get('正确答案'))
        progress_fn(len(questions), has_ans)
    try:
        save_progress(paper_id, questions, suffix=f'_{mode}' if mode and mode != 'ocr' else '')
    except Exception:
        pass
    return True


async def _recover_page(page, paper_id, questions: list, log_fn=None) -> bool:
    """页面恢复：重新导航到练习页面"""
    def log(msg):
        if log_fn:
            log_fn(msg)
        logger.info(msg)
    log("尝试恢复页面...")
    try:
        await navigate_with_retry(page, f'https://www.kaoshibao.com/ti/{paper_id}/')
        await asyncio.sleep(3)
        await close_popups(page)
        practice_url = (
            f'https://www.kaoshibao.com/online/?paperId={paper_id}'
            f'&practice=&modal=1&type=qtype_random&is_recite=&qtype='
            f'&text=顺序练习&sequence={len(questions)}'
            f'&is_collect=1&is_vip_paper=0&gid=&is_buy=1'
            f'&qtype_random_index={len(questions)}'
        )
        await navigate_with_retry(page, practice_url)
        await asyncio.sleep(3)
        await close_popups(page)
        await enable_recite_mode(page)
        await asyncio.sleep(2)
        return True
    except Exception as e:
        log(f"恢复失败: {e}")
        return False


async def collect_questions(page, paper_id, log_fn=None, stop_check=None, progress_fn=None,
                           start_sequence=0, max_questions=0, mode='ocr',
                           existing_questions=None, repeat_limit=10) -> tuple[list[dict[str, Any]], str]:
    """
    背题模式下收集所有题目+答案+解析

    Args:
        page: Playwright page
        paper_id: 题库ID
        log_fn: 日志回调 fn(msg)
        stop_check: 停止检查回调 fn() -> bool
        progress_fn: 进度回调 fn(collected, with_answer)
        start_sequence: 起始题目序号（用于分段并行）
        max_questions: 最多收集题目数（0=不限）
        existing_questions: 已有进度（用于去重，仅新增缺失题）
    """
    questions = []
    existing_keys = set()
    if existing_questions:
        questions = list(existing_questions)
        for q in questions:
            existing_keys.add(str(q.get('题目', ''))[:80])

    def log(msg):
        if log_fn:
            log_fn(msg)
        logger.info(msg)

    try:
        np_holder = {}
        status = await _navigate_to_practice(page, paper_id, start_sequence, log_fn, stop_check, new_page_holder=np_holder)
        if np_holder.get('page'):
            page = np_holder['page']
        if status != 'ok':
            return questions, status

        if not await _init_recite_mode(page, log_fn):
            return questions, 'error'

        log("开始收集...")
        consecutive_failures = 0
        no_new_count = 0
        last_success_i = 0  # 上次成功收集的迭代编号
        normal_end = False  # 是否正常结束（达上限/用户停止）

        for i in range(10000):
            if stop_check and stop_check():
                log("用户请求停止")
                normal_end = True
                break

            # 分段模式：达到最大题数后停止
            if max_questions > 0 and len(questions) >= max_questions:
                log(f"已达到分段上限 {max_questions} 题，停止收集")
                normal_end = True
                break

            # 检查登录是否过期
            if await check_login_required(page):
                log("登录已过期，尝试等待重新登录...")
                logged_in = await wait_for_login(page, log_fn, stop_check, timeout=120)
                if not logged_in:
                    log("登录过期，停止收集")
                    break
                await _recover_page(page, paper_id, questions, log_fn)

            await close_popups(page)

            # 卡住检测：如果连续30次迭代没有新题目，尝试恢复页面
            if i - last_success_i >= 30 and last_success_i > 0:
                log(f"已连续{i - last_success_i}次无新题目，尝试恢复页面...")
                if await _recover_page(page, paper_id, questions, log_fn):
                    last_success_i = i
                    consecutive_failures = 0
                else:
                    break

            # 等待题目加载
            loaded = await wait_for_question(page, timeout=10)
            if not loaded:
                consecutive_failures += 1
                if consecutive_failures >= 20:
                    log("连续20次未加载到题目，尝试恢复...")
                    if await _recover_page(page, paper_id, questions, log_fn):
                        consecutive_failures = 0
                        continue
                    else:
                        break
                await asyncio.sleep(2)
                continue

            data = await get_current_question(page, mode=mode)
            if not _is_valid_question_text(data.get('question', '')):
                consecutive_failures += 1
                if consecutive_failures % 3 == 0:
                    await asyncio.sleep(2)
                    data = await get_current_question(page, mode=mode)
                    if _is_valid_question_text(data.get('question', '')):
                        consecutive_failures = 0
                # 无效题目可能是考试宝插入的提示/广告题，点击"下一题"跳过
                if consecutive_failures >= 4:
                    log(f"题目文本无效已连续{consecutive_failures}次，点击下一题跳过...")
                    skipped = await click_next_question(page)
                    consecutive_failures = 0
                    if not skipped:
                        await asyncio.sleep(2)
                        skipped = await click_next_question(page)
                    if not skipped:
                        break
                    continue
                await asyncio.sleep(2)
                if consecutive_failures >= 20:
                    log("连续20次无法读取有效题目，尝试恢复...")
                    if await _recover_page(page, paper_id, questions, log_fn):
                        consecutive_failures = 0
                        continue
                    else:
                        break
                continue
            consecutive_failures = 0

            # 触发答案显示
            if not data.get('answer'):
                await trigger_answer(page, data.get('type', ''))
                data = await get_current_question(page, mode=mode)

            # 去重并保存
            if _add_question_if_new(data, questions, existing_keys, paper_id, mode, progress_fn):
                no_new_count = 0
                last_success_i = i
            else:
                no_new_count += 1
                if no_new_count >= repeat_limit:
                    log(f"连续{repeat_limit}题重复，可能已到末尾")
                    break

            # 定期日志
            if (i + 1) % 10 == 0:
                has_ans = sum(1 for q in questions if q.get('正确答案'))
                log(f"已收集 {len(questions)} 题, 有答案 {has_ans} 题")

            # 点击下一题
            moved = await click_next_question(page)
            if not moved:
                await asyncio.sleep(2)
                moved = await click_next_question(page)
            if not moved:
                consecutive_failures += 1
                if consecutive_failures >= 15:
                    log("连续15次无法翻到下一题，尝试恢复...")
                    if await _recover_page(page, paper_id, questions, log_fn):
                        consecutive_failures = 0
                        continue
                    else:
                        break
                await asyncio.sleep(2)

        has_ans = sum(1 for q in questions if q.get('正确答案'))
        log(f"完成: {len(questions)} 题, 有答案 {has_ans} 题")

        # 分段模式：未达上限且非正常结束（翻题失败/恢复失败等），标记 error 触发重试
        if max_questions > 0 and not normal_end and len(questions) < max_questions:
            log(f"分段未达上限: 已收集 {len(questions)}/{max_questions} 题，标记为需重试")
            return questions, 'error'

    except Exception as e:
        log(f"出错: {e}")
        await save_screenshot(page, 'error', log_fn)
        return questions, 'error'

    return questions, 'ok'


def save_to_excel(questions: list[dict[str, Any]], filepath: str) -> None:
    """保存题目到Excel（选项分列显示）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "题库"
    headers = ['序号', '题目类型', '题目', '选项A', '选项B', '选项C', '选项D', '选项E', '选项F', '正确答案', '解析']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center')
    for row, q in enumerate(questions, 2):
        ws.cell(row=row, column=1, value=q.get('序号', ''))
        ws.cell(row=row, column=2, value=q.get('题目类型', ''))
        ws.cell(row=row, column=3, value=q.get('题目', ''))

        # 选项分列
        opts = _normalize_options(q.get('选项', []))
        for i, opt in enumerate(opts[:6]):
            ws.cell(row=row, column=4 + i, value=opt)

        ws.cell(row=row, column=10, value=q.get('正确答案', ''))
        ws.cell(row=row, column=11, value=q.get('解析', ''))

        # 格式
        for col in [3, 4, 5, 6, 7, 8, 9, 11]:
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical='top')

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 20
    ws.column_dimensions['J'].width = 12
    ws.column_dimensions['K'].width = 40
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:K{len(questions) + 1}"
    wb.save(filepath)


def _normalize_options(opts) -> list[str]:
    """统一选项为列表形式"""
    if isinstance(opts, str):
        opts = [o.strip() for o in opts.split('\n') if o.strip()]
    cleaned_list = []
    for opt in (opts or []):
        cleaned = opt
        if len(cleaned) > 2 and cleaned[0] in 'ABCDEF' and cleaned[1] in '.、 ':
            cleaned = cleaned[2:].strip()
        elif len(cleaned) > 3 and cleaned[0] in 'ABCDEF' and cleaned[1:3] in ['. ', '、 ']:
            cleaned = cleaned[3:].strip()
        cleaned_list.append(cleaned)
    return cleaned_list


def save_to_csv(questions: list[dict[str, Any]], filepath: str) -> None:
    """保存题目到CSV（UTF-8 BOM，Excel 友好）"""
    import csv
    headers = ['序号', '题目类型', '题目', '选项A', '选项B', '选项C', '选项D', '选项E', '选项F', '正确答案', '解析']
    with open(filepath, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for q in questions:
            opts = _normalize_options(q.get('选项', []))
            row = [
                q.get('序号', ''),
                q.get('题目类型', ''),
                q.get('题目', ''),
            ]
            for i in range(6):
                row.append(opts[i] if i < len(opts) else '')
            row.append(q.get('正确答案', ''))
            row.append(q.get('解析', ''))
            writer.writerow(row)


def save_to_json(questions: list[dict[str, Any]], filepath: str) -> None:
    """保存题目到JSON（结构化，含选项数组）"""
    structured = []
    for q in questions:
        opts = _normalize_options(q.get('选项', []))
        structured.append({
            '序号': q.get('序号', ''),
            '题目类型': q.get('题目类型', ''),
            '题目': q.get('题目', ''),
            '选项': opts,
            '正确答案': q.get('正确答案', ''),
            '解析': q.get('解析', ''),
        })
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(structured, f, ensure_ascii=False, indent=2)


def export_questions(questions: list[dict[str, Any]], filepath: str, fmt: str = 'xlsx') -> None:
    """统一导出入口：根据格式调用对应保存函数"""
    fmt = (fmt or 'xlsx').lower()
    if fmt == 'csv':
        save_to_csv(questions, filepath)
    elif fmt == 'json':
        save_to_json(questions, filepath)
    else:
        save_to_excel(questions, filepath)


def load_progress(paper_id: str, suffix: str = '') -> list[dict[str, Any]]:
    """加载已有进度"""
    progress_file = DATA_DIR / f"progress_{paper_id}{suffix}.json"
    if progress_file.exists():
        try:
            with open(progress_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_progress(paper_id: str, questions: list[dict[str, Any]], suffix: str = '') -> None:
    """保存进度（原子写入，防止中断损坏）"""
    progress_file = DATA_DIR / f"progress_{paper_id}{suffix}.json"
    tmp_file = progress_file.with_suffix('.json.tmp')
    with open(tmp_file, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)
    os.replace(tmp_file, progress_file)


def load_segment_progress(paper_id: str, segment_id: int) -> list[dict[str, Any]]:
    """加载分段进度"""
    progress_file = DATA_DIR / f"progress_{paper_id}_seg{segment_id}.json"
    if progress_file.exists():
        try:
            with open(progress_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_segment_progress(paper_id: str, segment_id: int, questions: list[dict[str, Any]]) -> None:
    """保存分段进度（原子写入）"""
    progress_file = DATA_DIR / f"progress_{paper_id}_seg{segment_id}.json"
    tmp_file = progress_file.with_suffix('.json.tmp')
    with open(tmp_file, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)
    os.replace(tmp_file, progress_file)


def merge_segments(paper_id: str, num_segments: int) -> list[dict[str, Any]]:
    """合并所有分段的结果，去重后保存为最终文件"""
    all_questions = []
    seen = set()
    for seg_id in range(num_segments):
        seg_questions = load_segment_progress(paper_id, seg_id)
        for q in seg_questions:
            key = q.get('题目', '')[:80]
            if key and key not in seen:
                seen.add(key)
                all_questions.append(q)
    # 重新编号
    for i, q in enumerate(all_questions):
        q['序号'] = i + 1
    if all_questions:
        save_progress(paper_id, all_questions)
        output_file = OUTPUT_DIR / f"考试宝题库_{paper_id}.xlsx"
        save_to_excel(all_questions, output_file)
    # 清理分段文件
    for seg_id in range(num_segments):
        seg_file = DATA_DIR / f"progress_{paper_id}_seg{seg_id}.json"
        if seg_file.exists():
            seg_file.unlink()
    return all_questions


def validate_paper_id(raw: str) -> tuple[str | None, str | None]:
    """验证并提取题库ID，返回 (paper_id, error_msg)"""
    raw = raw.strip()
    if not raw:
        return None, '请输入题库ID'
    if 'kaoshibao.com' in raw:
        m = re.search(r'/ti/(\d+)', raw)
        if m:
            return m.group(1), None
        return None, '无法从URL中提取题库ID'
    if not raw.isdigit():
        return None, '题库ID必须是数字'
    if len(raw) > 20:
        return None, '题库ID过长'
    return raw, None


def sanitize_text(text: Any) -> str:
    """转义HTML特殊字符，用于前端安全渲染"""
    if not text:
        return ''
    return html.escape(str(text))
