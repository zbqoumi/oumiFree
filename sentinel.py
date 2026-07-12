"""Sentinel (Cloudflare) token extraction — shared module."""
import json, time, os, sys, uuid, secrets
from urllib.parse import quote
from pathlib import Path

CACHE_TTL = 600


def _app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def _load_config():
    p = _app_dir() / "config.json"
    if p.exists():
        return json.load(open(p, encoding="utf-8"))
    return {
        "chatgpt": {
            "auth_base_url": "https://auth.openai.com",
            "chat_base_url": "https://chatgpt.com",
            "chat_web_client_id": "app_X8zY6vW2pQ9tR3dE7nK1jL5gH",
            "user_agent_chrome": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        },
        "http": {
            "proxy": {"default": "http://127.0.0.1:7897"},
        },
    }


def get_cached(force_fresh=False):
    cf = _app_dir() / "sentinel_cache.json"
    if force_fresh:
        return None
    if cf.exists():
        try:
            cache = json.load(open(cf))
            if time.time() - cache.get("ts", 0) < CACHE_TTL and cache.get("sentinel_token"):
                return cache
        except Exception:
            pass
    return None


def save_cache(data):
    data["ts"] = time.time()
    json.dump(data, open(_app_dir() / "sentinel_cache.json", "w"), ensure_ascii=False)


def clear_cache():
    cf = _app_dir() / "sentinel_cache.json"
    if cf.exists():
        cf.unlink()


def get_proxy_url():
    try:
        return _load_config().get("http", {}).get("proxy", {}).get("default")
    except Exception:
        return None


def get_playwright_proxy():
    url = get_proxy_url()
    if not url:
        return None
    import re
    # 带认证的代理: http://user:pass@host:port
    m = re.match(r"(https?)://([^:]+):([^@]+)@([^:]+):(\d+)", url)
    if m:
        return {
            "server": f"{m.group(1)}://{m.group(4)}:{m.group(5)}",
            "username": m.group(2),
            "password": m.group(3),
        }
    # 无认证的代理: http://host:port 或 socks5://host:port
    m = re.match(r"^(https?|socks5)://([^:]+):(\d+)$", url)
    if m:
        return {"server": url}
    return None


_CHROME_CACHE = None


def _find_chrome():
    global _CHROME_CACHE
    if _CHROME_CACHE:
        return _CHROME_CACHE

    # Windows paths
    for p in (
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ):
        if os.path.exists(p):
            _CHROME_CACHE = p
            return p

    # macOS paths
    for p in (
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        os.path.expanduser("~/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
    ):
        if os.path.exists(p):
            _CHROME_CACHE = p
            return p

    # Linux paths
    for p in ("/usr/bin/google-chrome", "/usr/bin/chromium-browser", "/usr/bin/chromium"):
        if os.path.exists(p):
            _CHROME_CACHE = p
            return p

    # Fallback: search PATH
    try:
        import subprocess, shutil
        for name in ("google-chrome", "chrome", "chromium-browser", "chromium", "Google Chrome"):
            found = shutil.which(name)
            if found:
                _CHROME_CACHE = found
                return found
    except Exception:
        pass

    return None


def extract_sentinel(force_fresh=False, use_proxy=True, stop_event=None):
    if not force_fresh:
        cached = get_cached()
        if cached:
            return cached

    cfg = _load_config()
    from playwright.sync_api import sync_playwright

    chrome = _find_chrome()
    if not chrome:
        raise RuntimeError("未找到 Google Chrome，请安装: https://www.google.com/chrome/")

    if stop_event and stop_event.is_set():
        return None

    proxy_info = get_playwright_proxy()
    print(f"[Sentinel] Chrome: {chrome}")
    print(f"[Sentinel] Proxy: {proxy_info}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            executable_path=chrome,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )

        ctx_kwargs = {
            "user_agent": cfg.get("chatgpt", {}).get("user_agent_chrome") or cfg.get("http", {}).get("user_agent_chrome") or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            "viewport": {"width": 1280, "height": 800},
            "locale": "en-US",
            "timezone_id": "America/New_York",
        }

        if use_proxy and proxy_info:
            ctx_kwargs["proxy"] = proxy_info

        ctx = browser.new_context(**ctx_kwargs)
        page = ctx.new_page()

        device_id = str(uuid.uuid4())
        state = secrets.token_urlsafe(32)
        scope = "openid email profile offline_access model.request model.read organization.read organization.write"

        auth_url = (
            f"{cfg['chatgpt']['auth_base_url']}"
            f"/api/accounts/authorize?client_id={cfg['chatgpt']['chat_web_client_id']}"
            f"&scope={quote(scope)}"
            f"&response_type=code&redirect_uri={quote('https://chatgpt.com/api/auth/callback/openai')}"
            f"&audience={quote('https://api.openai.com/v1')}"
            f"&device_id={device_id}"
            f"&prompt=login&screen_hint=signup&state={state}"
        )

        print(f"[Sentinel] Opening: {auth_url[:80]}...")
        try:
            resp = page.goto(auth_url, wait_until="domcontentloaded", timeout=30000)
            print(f"[Sentinel] Page loaded: status={resp.status if resp else 'None'}, url={page.url[:80]}")
        except Exception as e:
            print(f"[Sentinel] goto domcontentloaded failed: {e}")
            try:
                resp = page.goto(auth_url, wait_until="commit", timeout=30000)
                print(f"[Sentinel] Page loaded (commit): status={resp.status if resp else 'None'}, url={page.url[:80]}")
            except Exception as e2:
                print(f"[Sentinel] goto commit also failed: {e2}")
                browser.close()
                return None

        found = False
        for i in range(45):
            if stop_event and stop_event.is_set():
                print("[Sentinel] Cancelled by stop_event")
                browser.close()
                return None
            time.sleep(2)
            try:
                ready = page.evaluate("typeof window.SentinelSDK !== 'undefined'")
                if i % 5 == 0:
                    print(f"[Sentinel] Waiting... ({i*2}s) SentinelSDK={ready} url={page.url[:60]}")
                if ready:
                    found = True
                    break
            except Exception:
                pass

        if not found:
            browser.close()
            return None

        if stop_event and stop_event.is_set():
            browser.close()
            return None

        page.evaluate("SentinelSDK.init()")
        time.sleep(2)
        did = page.evaluate("document.cookie.match(/oai-did=([^;]+)/)?.[1] || ''")

        sentinel_token = page.evaluate(
            "(d) => SentinelSDK.token().then(r => { let p=JSON.parse(r); p.id=d; p.flow='username_password_create'; return JSON.stringify(p); })",
            did,
        )
        sentinel_so = page.evaluate(
            "(d) => SentinelSDK.token().then(r => { let p=JSON.parse(r); return JSON.stringify({so:r, c:p.c, id:d, flow:'oauth_create_account'}); })",
            did,
        )

        cookie_str = "; ".join(f"{c['name']}={c['value']}" for c in ctx.cookies())
        browser.close()

    result = {
        "sentinel_token": sentinel_token,
        "sentinel_so_token": sentinel_so,
        "cookie_str": cookie_str,
        "oai_did": did,
    }
    save_cache(result)
    return result


_extract_sentinel = extract_sentinel
_get_cached_sentinel = get_cached
