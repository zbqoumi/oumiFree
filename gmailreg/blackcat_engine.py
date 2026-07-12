"""欧米freeGPT — ChatGPT 自动注册引擎（Gmail / Outlook OAuth2 双模式，不加入工作区）"""
import json, time, re, random, string, sys, uuid, secrets, base64, hashlib
import imaplib
import email as eml
import urllib.request as urllib
from urllib.parse import quote, urlencode, urlparse, parse_qs
from pathlib import Path
from datetime import datetime, timezone, timedelta

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from curl_cffi import requests as curl_requests
from sentinel import extract_sentinel, clear_cache, get_proxy_url

BASE_DIR = Path(__file__).resolve().parent


def _get_app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


_PARENT_CFG = _get_app_dir() / "config.json"
if _PARENT_CFG.exists():
    CFG = json.load(open(_PARENT_CFG))
else:
    CFG = {
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

UA = CFG.get("http", {}).get("user_agent_chrome", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

UA_LIST = [
    UA,
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
]

PROXY_URL = get_proxy_url()

AUTH_BASE = CFG.get("chatgpt", {}).get("auth_base_url", "https://auth.openai.com")
CHAT_BASE = CFG.get("chatgpt", {}).get("chat_base_url", "https://chatgpt.com")

IMAP_SERVERS = {"gmail": "imap.gmail.com", "outlook": "outlook.office365.com", "hotmail": "outlook.office365.com"}

# —— Graph API 配置（common 端点 + .default scope）——
OUTLOOK_GRAPH_TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
OUTLOOK_GRAPH_MESSAGES_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
OUTLOOK_GRAPH_SCOPE = "https://graph.microsoft.com/.default"

# —— IMAP 配置（consumers 端点 + outlook.live.com）——
OUTLOOK_IMAP_TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
OUTLOOK_IMAP_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
OUTLOOK_IMAP_HOST = "outlook.live.com"

# 所有"账号列表"模式
ACCOUNT_LIST_MODES = ("api", "outlook_token")

OAUTH_TOKEN_FILE = BASE_DIR / "outlook_oauth2.json"

NAMES_FIRST = ["James", "John", "Robert", "Michael", "David", "William", "Mary", "Linda", "Barbara", "Jennifer", "Elizabeth", "Susan"]
NAMES_LAST = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Wilson", "Anderson", "Taylor", "Thomas"]


def _random_name():
    return (random.choice(NAMES_FIRST), random.choice(NAMES_LAST))


def _random_birthdate():
    return f"{random.randint(1985, 2004)}-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}"


def _gen_pw():
    cs = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    return "".join(random.choices(cs, k=8)) + "!A1"


def _gen_alias_email(base_email):
    l, d = base_email.split("@", 1)
    suffix = "".join(random.choices(string.ascii_lowercase, k=random.randint(4, 6)))
    return f"{l}+{suffix}@{d}"


def _parse_csrf(h):
    m = re.search(r"__Host-next-auth\.csrf-token=([^;]+)", h or "")
    if not m:
        return ""
    return m.group(1).split("%7C")[0]


def _random_ua():
    return random.choice(UA_LIST)


def _random_imp():
    return "chrome"


class BlackCatEngine:
    def __init__(self, mode="gmail", imap_user="", imap_pass="", workspace_ids=None,
                 codexproxy_url="", codexproxy_key="", log_cb=None, save_dir="codex",
                 outlook_client_id="14d82eec-204b-4c2f-b7e8-296a70dab67e",
                 manual_otp_cb=None, accounts=None, stop_event=None):
        self.mode = mode
        self.imap_user = (imap_user or "").replace(" ", "")
        self.imap_pass = imap_pass or ""
        self.imap_host = IMAP_SERVERS.get(mode, "imap.gmail.com")
        self.workspace_ids = workspace_ids or []
        self.codexproxy_url = codexproxy_url
        self.codexproxy_key = codexproxy_key
        self.save_dir = Path(save_dir or "codex")
        if not self.save_dir.is_absolute():
            self.save_dir = BASE_DIR / self.save_dir
        self.outlook_client_id = outlook_client_id
        self.manual_otp_cb = manual_otp_cb
        self.accounts = accounts or []
        self._stop_event = stop_event
        self._current_api_url = ""
        self.outlook_refresh_token = ""
        self.outlook_mail_mode = "auto"  # graph / imap / auto — 自动探测哪种 scope 可用
        self._token_cache = {}  # scope → (token, expiry_monotonic)
        self.log = log_cb or print

    def _load_oauth2_token(self):
        if OAUTH_TOKEN_FILE.exists():
            return json.load(open(OAUTH_TOKEN_FILE, encoding="utf-8"))
        return {}

    def _save_oauth2_token(self, data):
        OAUTH_TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        json.dump(data, open(OAUTH_TOKEN_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    def _get_outlook_access_token(self):
        token_data = self._load_oauth2_token()
        access_token = token_data.get("access_token", "")
        refresh_token = token_data.get("refresh_token", "")
        if not refresh_token:
            return ""
        if time.time() < token_data.get("expires_at", 0) - 60 and access_token:
            return access_token
        self.log("  [OAuth2] 刷新 Outlook token...")
        try:
            from msal import PublicClientApplication
            actual_client_id = self.outlook_client_id or "d3590ed6-52b3-4102-aeff-aad2292ab01c"
            app = PublicClientApplication(
                client_id=actual_client_id,
                authority="https://login.microsoftonline.com/consumers",
            )
            result = app.acquire_token_by_refresh_token(
                refresh_token,
                scopes=["https://graph.microsoft.com/Mail.Read", "https://outlook.office.com/IMAP.AccessAsUser.All"],
            )
            if "access_token" in result:
                token_data["access_token"] = result["access_token"]
                token_data["refresh_token"] = result.get("refresh_token", refresh_token)
                token_data["expires_at"] = time.time() + result.get("expires_in", 3600)
                self._save_oauth2_token(token_data)
                return result["access_token"]
            else:
                self.log("  [OAuth2] 刷新失败: " + str(result.get("error_description", result.get("error", ""))))
        except ImportError:
            self.log("  [OAuth2] msal 未安装，请运行: pip install msal")
        except Exception as e:
            self.log("  [OAuth2] 刷新失败: " + str(e))
        return ""

    def outlook_login(self):
        try:
            from msal import PublicClientApplication
        except ImportError:
            return {"error": "请先安装 msal: pip install msal"}
        app = PublicClientApplication(
            client_id=self.outlook_client_id,
            authority="https://login.microsoftonline.com/consumers",
        )
        flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/Mail.Read", "https://outlook.office.com/IMAP.AccessAsUser.All"])
        if "error" in flow:
            return flow
        user_code = flow.get("user_code", "")
        verification_uri = flow.get("verification_uri", "https://microsoft.com/devicelogin")
        message = flow.get("message", "")
        expires_in = flow.get("expires_in", 900)
        return {"device_flow": flow, "verification_uri": verification_uri, "user_code": user_code, "message": message, "expires_in": expires_in}

    def outlook_login_finish(self, device_flow_result):
        try:
            from msal import PublicClientApplication
        except ImportError:
            return {"error": "msal 未安装"}
        app = PublicClientApplication(
            client_id=self.outlook_client_id,
            authority="https://login.microsoftonline.com/consumers",
        )
        flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/Mail.Read", "https://outlook.office.com/IMAP.AccessAsUser.All"])
        flow["user_code"] = str(device_flow_result)
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            return {"error": result.get("error_description", result.get("error", "Unknown error"))}
        token_data = {
            "access_token": result["access_token"],
            "refresh_token": result.get("refresh_token", ""),
            "expires_at": time.time() + result.get("expires_in", 3600) - 60,
        }
        self._save_oauth2_token(token_data)
        return {"success": True}

    def _poll_otp(self, target, timeout=120):
        if self.manual_otp_cb:
            return self.manual_otp_cb(target, timeout)
        if self.mode == "api":
            return self._poll_otp_api(target, timeout)
        elif self.mode == "outlook":
            return self._poll_otp_outlook(target, timeout)
        elif self.mode == "outlook_token":
            return self._poll_otp_outlook_token(target, timeout)
        else:
            return self._poll_otp_gmail(target, timeout)

    def _poll_otp_gmail(self, target, timeout=120):
        dead = time.time() + timeout
        last = 0
        while time.time() < dead:
            if self._stop_event and self._stop_event.is_set():
                break
            time.sleep(3)
            try:
                m = imaplib.IMAP4_SSL(self.imap_host, 993)
                m.login(self.imap_user, self.imap_pass)
                m.select("inbox")
                _, msgs = m.search(None, '(TO "%s")' % target)
                if msgs[0]:
                    for mid in reversed(msgs[0].split()):
                        _, d = m.fetch(mid, "(RFC822)")
                        msg = eml.message_from_bytes(d[0][1])
                        body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() in ("text/plain", "text/html"):
                                    body += part.get_payload(decode=True).decode("utf-8", "replace")
                        else:
                            body = msg.get_payload(decode=True).decode("utf-8", "replace")
                        body = re.sub(r"<style[^>]*>.*?</style>", "", body, flags=re.DOTALL)
                        body = re.sub(r"<[^>]+>", " ", body)
                        body = re.sub(r"\s+", " ", body).strip()
                        p = re.search(r"(?<!\d)(\d{6})(?!\d)", body)
                        if p and time.time() - last > 5:
                            m.logout()
                            return p.group(1)
                m.logout()
            except Exception:
                pass
            time.sleep(0.5)
        return None

    def _poll_otp_outlook(self, target, timeout=120):
        dead = time.time() + timeout
        last = 0
        token = self._get_outlook_access_token()
        if not token:
            self.log("  [OAuth2] 无有效 token，请在界面中先登录 Outlook")
            return None
        while time.time() < dead:
            time.sleep(3)
            try:
                m = imaplib.IMAP4_SSL(self.imap_host, 993)
                auth = f"user={self.imap_user}\x01auth=Bearer {token}\x01\x01"
                m.authenticate("XOAUTH2", lambda x: auth.encode())
                m.select("inbox")
                _, msgs = m.search(None, '(TO "%s")' % target)
                if msgs[0]:
                    for mid in reversed(msgs[0].split()):
                        _, d = m.fetch(mid, "(RFC822)")
                        msg = eml.message_from_bytes(d[0][1])
                        body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() in ("text/plain", "text/html"):
                                    body += part.get_payload(decode=True).decode("utf-8", "replace")
                        else:
                            body = msg.get_payload(decode=True).decode("utf-8", "replace")
                        body = re.sub(r"<style[^>]*>.*?</style>", "", body, flags=re.DOTALL)
                        body = re.sub(r"<[^>]+>", " ", body)
                        body = re.sub(r"\s+", " ", body).strip()
                        p = re.search(r"(?<!\d)(\d{6})(?!\d)", body)
                        if p and time.time() - last > 5:
                            m.logout()
                            return p.group(1)
                m.logout()
            except Exception as e:
                self.log("  [Outlook] " + str(e))
                time.sleep(5)
            time.sleep(0.5)
        return None

    def _poll_otp_api(self, target, timeout=120):
        api_url = getattr(self, "_current_api_url", "")
        if not api_url:
            self.log("  [API] 未设置 API 链接")
            return None

        def find_code(value):
            if isinstance(value, int) and 100000 <= value <= 999999:
                return str(value)
            if isinstance(value, str):
                m = re.search(r"(?<!\d)(\d{6})(?!\d)", value)
                if m:
                    return m.group(1)
            if isinstance(value, dict):
                for key in ("code", "otp", "verify_code", "verification_code", "data", "msg", "message"):
                    code = find_code(value.get(key))
                    if code:
                        return code
                for v in list(value.values()):
                    code = find_code(v)
                    if code:
                        return code
            return None

        self.log("  [API] 等 10 秒后开始轮询验证码...")
        time.sleep(10)
        dead = time.time() + timeout
        last_report = ""
        while time.time() < dead:
            if self._stop_event and self._stop_event.is_set():
                break
            try:
                req = urllib.Request(api_url, headers={"User-Agent": UA}, method="GET")
                with urllib.urlopen(req, timeout=15) as resp:
                    body = resp.read().decode("utf-8", "replace")
                data = json.loads(body)
                body_preview = body[:300]
                if body_preview != last_report:
                    self.log(f"  [API] 返回: {body_preview}")
                    last_report = body_preview
                code = find_code(data)
                if code:
                    self.log(f"  [API] 提取到验证码: {code}")
                    return code
            except Exception as e:
                self.log("  [API] 轮询失败: " + str(e))
            time.sleep(3)
        return None

    # ==================== Outlook Token 接码方法 (refresh_token 方案) ====================

    def _build_url_opener(self):
        """构建带代理的 urllib opener"""
        if PROXY_URL:
            handler = urllib.ProxyHandler({"http": PROXY_URL, "https": PROXY_URL})
            return urllib.build_opener(handler)
        return urllib.build_opener()

    def _exchange_refresh_token(self, client_id, refresh_token, scope, token_url):
        """用 curl_cffi 刷新 access_token（对齐 mail_provider.py 开源实现，浏览器 TLS 指纹）"""
        kwargs = {"impersonate": "chrome", "verify": False}
        if PROXY_URL:
            kwargs["proxy"] = PROXY_URL
        resp = curl_requests.post(
            token_url,
            data={
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": scope,
            },
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "User-Agent": UA,
            },
            timeout=15,
            **kwargs,
        )
        try:
            result = resp.json()
        except Exception:
            result = {}
        if resp.status_code != 200:
            detail = result.get("error_description") or result.get("error") or resp.text[:300]
            raise RuntimeError(f"refresh_token 刷新 HTTP {resp.status_code}: {str(detail)[:200]}")
        access_token = result.get("access_token", "").strip()
        if not access_token:
            raise RuntimeError("refresh_token 响应缺少 access_token")
        return access_token

    def _cached_access_token(self, client_id, refresh_token, scope, token_url):
        """缓存 access_token 10 分钟，避免轮询时频繁刷新"""
        cached = self._token_cache.get(scope)
        if cached and time.monotonic() < cached[1]:
            return cached[0]
        token = self._exchange_refresh_token(client_id, refresh_token, scope, token_url)
        self._token_cache[scope] = (token, time.monotonic() + 600)
        return token

    @staticmethod
    def _extract_outlook_code(subject, body):
        """三层策略提取 6 位验证码（对齐外层 mail_provider.py._extract_code）"""
        content = f"{subject}\n{body}".strip()
        if not content:
            return None
        # 第一层: OpenAI 邮件 HTML 样式 background-color: #F3F3F3
        match = re.search(
            r"background-color:\s*#F3F3F3[^>]*>[\s\S]*?(\d{6})[\s\S]*?</p>",
            content, re.I,
        )
        if match:
            return match.group(1)
        # 第二层: 关键词后的 6 位数字
        match = re.search(
            r"(?:Verification code|code is|代码为|验证码)[:\s]*(\d{6})",
            content, re.I,
        )
        if match and match.group(1) != "177010":
            return match.group(1)
        # 第三层: 通用 6 位数字（排除已知误报 177010）
        for code in re.findall(r">\s*(\d{6})\s*<|(?<![#&])\b(\d{6})\b", content):
            value = code[0] or code[1]
            if value and value != "177010":
                return value
        return None

    def _read_graph_messages(self, access_token, boundary_str):
        """通过 Graph API 读取 boundary 之后的邮件（用 curl_cffi 浏览器模拟）"""
        kwargs = {"impersonate": "chrome", "verify": False}
        if PROXY_URL:
            kwargs["proxy"] = PROXY_URL
        resp = curl_requests.get(
            OUTLOOK_GRAPH_MESSAGES_URL,
            params={
                "$top": 10,
                "$orderby": "receivedDateTime desc",
                "$filter": f"receivedDateTime ge {boundary_str}",
                "$select": "subject,body,from,toRecipients,receivedDateTime",
            },
            headers={
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json",
                "User-Agent": UA,
                "Prefer": "outlook.body-content-type='text'",
            },
            timeout=15,
            **kwargs,
        )
        try:
            data = resp.json()
        except Exception:
            data = {}
        if resp.status_code != 200:
            err = data.get("error", {})
            detail = err.get("message", "") if isinstance(err, dict) else resp.text[:300]
            raise RuntimeError(f"Graph API HTTP {resp.status_code}: {str(detail)[:200]}")
        messages = []
        for msg in data.get("value", []):
            from_addr = msg.get("from", {}).get("emailAddress", {}).get("address", "")
            subject = msg.get("subject", "")
            body_content = msg.get("body", {}).get("content", "")
            combined = subject + " " + body_content
            if "openai" not in from_addr.lower() and "openai" not in combined.lower():
                continue
            messages.append({"subject": subject, "body": body_content})
        return messages

    def _read_imap_messages(self, access_token):
        """通过 IMAP XOAUTH2 (OAuth2 令牌) 读取最近邮件"""
        messages = []
        auth_string = f"user={self.imap_user}\x01auth=Bearer {access_token}\x01\x01"
        imap = imaplib.IMAP4_SSL(OUTLOOK_IMAP_HOST)
        try:
            imap.authenticate("XOAUTH2", lambda _: auth_string.encode("utf-8"))
            imap.select("inbox", readonly=True)
            _, data = imap.uid("search", None, "ALL")
            if not data or not data[0]:
                return []
            uids = data[0].split()[-10:]
            for uid in reversed(uids):
                _, fetched = imap.uid("fetch", uid, "(RFC822)")
                raw = next(
                    (p[1] for p in fetched if isinstance(p, tuple) and isinstance(p[1], bytes)),
                    b"",
                )
                if not raw:
                    continue
                msg_obj = eml.message_from_bytes(raw)
                subject = msg_obj.get("Subject", "")
                body = ""
                if msg_obj.is_multipart():
                    for part in msg_obj.walk():
                        if part.get_content_type() in ("text/plain", "text/html"):
                            body += part.get_payload(decode=True).decode("utf-8", "replace")
                else:
                    body = msg_obj.get_payload(decode=True).decode("utf-8", "replace")
                from_addr = msg_obj.get("From", "")
                if "openai" not in from_addr.lower() and "openai" not in body.lower() and "openai" not in subject.lower():
                    continue
                messages.append({"subject": subject, "body": body})
        finally:
            try:
                imap.logout()
            except Exception:
                pass
        return messages

    def _poll_otp_outlook_token(self, target, timeout=120):
        """Outlook refresh_token 方案 — 自动探测 Graph / IMAP 哪种 scope 可用"""
        client_id = self.outlook_client_id or ""
        refresh_token = self.outlook_refresh_token or ""
        if not client_id or not refresh_token:
            self.log("  [Outlook] 缺少 client_id 或 refresh_token，无法接码")
            return None

        self.log("  [Outlook] 开始 refresh_token 接码...")
        boundary_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        dead = time.time() + timeout
        first_poll = True

        while time.time() < dead:
            if self._stop_event and self._stop_event.is_set():
                break
            time.sleep(5 if not first_poll else 2)
            first_poll = False

            messages = []
            errors = []

            # Graph API 模式
            if self.outlook_mail_mode in ("graph", "auto"):
                try:
                    token = self._cached_access_token(client_id, refresh_token, OUTLOOK_GRAPH_SCOPE, OUTLOOK_GRAPH_TOKEN_URL)
                    messages = self._read_graph_messages(token, boundary_str)
                    if messages:
                        self.log(f"  [Graph] 读取到 {len(messages)} 封相关邮件")
                        if self.outlook_mail_mode == "auto":
                            self.outlook_mail_mode = "graph"
                except Exception as e:
                    self._token_cache.pop(OUTLOOK_GRAPH_SCOPE, None)
                    if self.outlook_mail_mode == "graph":
                        self.log(f"  [Graph] 失败: {str(e)[:120]}")
                        return None
                    errors.append(f"graph: {str(e)[:80]}")

            # IMAP XOAUTH2 回退
            if not messages and self.outlook_mail_mode in ("imap", "auto"):
                try:
                    token = self._cached_access_token(client_id, refresh_token, OUTLOOK_IMAP_SCOPE, OUTLOOK_IMAP_TOKEN_URL)
                    messages = self._read_imap_messages(token)
                    if messages:
                        self.log(f"  [IMAP] 读取到 {len(messages)} 封相关邮件")
                        if self.outlook_mail_mode == "auto":
                            self.outlook_mail_mode = "imap"
                except Exception as e:
                    self._token_cache.pop(OUTLOOK_IMAP_SCOPE, None)
                    if self.outlook_mail_mode == "imap":
                        self.log(f"  [IMAP] 失败: {str(e)[:120]}")
                        return None
                    errors.append(f"imap: {str(e)[:80]}")

            if errors and not messages:
                self.log(f"  [Outlook] 轮询失败: {'; '.join(errors)[:200]}")

            # 三层策略提取验证码
            for msg in messages:
                code = self._extract_outlook_code(msg["subject"], msg["body"])
                if code:
                    self.log(f"  [Outlook] 提取到验证码: {code}")
                    return code

        self.log("  [Outlook] 超时未获取到验证码")
        return None

    def _oauth_get_tokens(self, session, email="", oai_did="", log_cb=None):
        """利用已有 session cookies 发起 Platform OAuth PKCE 流程，换取 refresh_token

        对齐 login_flow.py 的做法：
        - client_id: app_2SKx67EdpoN0G6j64rFvigXD (Platform)
        - redirect_uri: https://platform.openai.com/auth/callback
        - authorize: /api/accounts/authorize, allow_redirects=True
        - token: /api/accounts/oauth/token
        """
        dbg = log_cb or (lambda msg: None)

        code_verifier = secrets.token_urlsafe(64)
        code_challenge = base64.urlsafe_b64encode(
            hashlib.sha256(code_verifier.encode()).digest()
        ).rstrip(b"=").decode()

        client_id = "app_2SKx67EdpoN0G6j64rFvigXD"
        redirect_uri = "https://platform.openai.com/auth/callback"
        audience = "https://api.openai.com/v1"
        device_id = str(uuid.uuid4())
        auth0_client = "eyJuYW1lIjoiYXV0aDAtc3BhLWpzIiwidmVyc2lvbiI6IjEuMjEuMCJ9"

        params = {
            "issuer": AUTH_BASE,
            "client_id": client_id,
            "audience": audience,
            "redirect_uri": redirect_uri,
            "device_id": device_id,
            "screen_hint": "login",
            "max_age": "0",
            "login_hint": email,
            "scope": "openid profile email offline_access",
            "response_type": "code",
            "response_mode": "query",
            "state": secrets.token_urlsafe(32),
            "nonce": secrets.token_urlsafe(32),
            "code_challenge": code_challenge,
            "code_challenge_method": "S256",
            "auth0Client": auth0_client,
        }
        auth_url = f"{AUTH_BASE}/api/accounts/authorize?{urlencode(params)}"
        dbg(f"  [OAuth] authorize URL: {auth_url[:120]}...")

        session.cookies.set("oai-did", oai_did or device_id, domain=".auth.openai.com")

        r = session.get(
            auth_url,
            headers={
                "User-Agent": UA,
                "Accept": "text/html,application/xhtml+xml",
                "auth0-client": auth0_client,
            },
            impersonate="chrome",
            allow_redirects=True,
            timeout=30,
        )
        dbg(f"  [OAuth] authorize -> status={r.status_code}, final_url={str(getattr(r, 'url', ''))[:150]}")

        # 从最终回调 URL 提取 code
        final_url = str(getattr(r, "url", "") or "")
        try:
            parsed = parse_qs(urlparse(final_url).query)
        except Exception:
            parsed = {}
        code = str((parsed.get("code") or [""])[0]).strip()

        if not code:
            dbg("  [OAuth] 未获取到 authorization code")
            try:
                body_snippet = r.text[:300] if r else "no response"
                dbg(f"  [OAuth] 最后响应体: {body_snippet}")
            except Exception:
                pass
            return None

        dbg(f"  [OAuth] 提取到 code: {code[:20]}...")

        r = session.post(
            f"{AUTH_BASE}/api/accounts/oauth/token",
            data={
                "client_id": client_id,
                "code_verifier": code_verifier,
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": redirect_uri,
            },
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "User-Agent": UA,
            },
            impersonate="chrome",
            timeout=30,
        )
        dbg(f"  [OAuth] token exchange -> status={r.status_code}")
        if r.status_code == 200:
            data = r.json()
            dbg(f"  [OAuth] 成功, keys={list(data.keys())}")
            return data
        else:
            dbg(f"  [OAuth] token exchange 失败: {r.text[:200]}")
        return None

    def _cp_upload(self, email, rt):
        d = json.dumps({"name": email, "refresh_token": rt}).encode()
        req = urllib.Request(
            self.codexproxy_url + "/api/admin/accounts",
            data=d,
            headers={"Content-Type": "application/json", "X-Admin-Key": self.codexproxy_key},
            method="POST",
        )
        try:
            with urllib.urlopen(req, timeout=15) as resp:
                return json.loads(resp.read().decode())
        except Exception as e:
            return {"error": str(e)}

    def _cp_status(self):
        req = urllib.Request(
            self.codexproxy_url + "/api/admin/accounts",
            headers={"X-Admin-Key": self.codexproxy_key},
        )
        try:
            with urllib.urlopen(req, timeout=10) as resp:
                accts = json.loads(resp.read().decode()).get("accounts", [])
            t = len(accts)
            a = sum(1 for x in accts if x.get("status") == "active")
            rl = sum(1 for x in accts if x.get("status") == "rate_limited")
            err = t - a - rl
            return (t, a, rl, err)
        except Exception:
            return (0, 0, 0, 0)

    def register(self, email=None, callback=None):
        cb = callback or self.log

        if email is None:
            email = self.imap_user
        email = _gen_alias_email(email)

        password = _gen_pw()
        fn, ln = _random_name()
        bd = _random_birthdate()
        oai_did = str(uuid.uuid4())

        cb("\n" + "==================================================")
        cb(f"  邮箱: {email}")
        cb(f"  密码: {password}")
        cb(f"  姓名: {fn} {ln}")
        cb(f"  生日: {bd}")

        # [1/9] Sentinel
        cb("[1/9] Sentinel...")
        t0 = time.time()
        try:
            clear_cache()
            sd = extract_sentinel(force_fresh=True, stop_event=self._stop_event)
        except Exception as e:
            cb(f"  [错误] Sentinel 异常: {e}")
            sd = None
        if not sd:
            cb("  [错误] Sentinel 提取失败")
            return (False, "Sentinel 提取失败")
        cb(f"  OK ({time.time() - t0:.1f}s)")

        # [2/9] CSRF
        session = curl_requests.Session()
        if PROXY_URL:
            session.proxies = {"http": PROXY_URL, "https": PROXY_URL}

        for pair in sd.get("cookie_str", "").split("; "):
            if "=" not in pair:
                continue
            k, v = pair.split("=", 1)
            if any(k.startswith(p) for p in ("oai-login-csrf", "oai-did", "oai-client-auth", "auth-session")):
                continue
            session.cookies.set(k, v, domain=".openai.com")

        bh = {"User-Agent": UA, "Accept": "application/json"}

        cb("[2/9] CSRF...")
        t0 = time.time()
        r = session.get(
            f"{CHAT_BASE}/auth/login",
            headers={**bh, "Accept": "text/html,application/xhtml+xml"},
            impersonate="chrome",
            timeout=30,
        )
        csrf = _parse_csrf(r.headers.get("Set-Cookie", ""))
        r = session.get(f"{CHAT_BASE}/api/auth/csrf", headers=bh, impersonate="chrome", timeout=15)
        if not csrf:
            csrf = _parse_csrf(r.headers.get("Set-Cookie", ""))
        if not csrf:
            csrf = "true"
        cb(f"  OK ({time.time() - t0:.1f}s)")

        # [3/9] Initiate registration
        cb("[3/9] 发起注册...")
        t0 = time.time()
        sp = {
            "prompt": "login",
            "ext-oai-did": oai_did,
            "auth_session_logging_id": str(uuid.uuid4()).replace("-", ""),
            "screen_hint": "login_or_signup",
            "login_hint": email,
        }
        r = session.post(
            f"{CHAT_BASE}/api/auth/signin/openai?" + urlencode(sp),
            data=urlencode({"csrfToken": csrf}),
            headers={**bh, "Content-Type": "application/x-www-form-urlencoded",
                     "Origin": CHAT_BASE, "Referer": f"{CHAT_BASE}/auth/login"},
            impersonate=_random_imp(),
            allow_redirects=False,
            timeout=30,
        )
        sb = {}
        try:
            sb = r.json()
        except Exception:
            pass

        ru = sb.get("url") or r.headers.get("Location", "")
        if not ru:
            cb("  [错误] 注册发起失败: " + r.text[:200])
            return (False, "注册发起失败: " + r.text[:200])
        cb(f"  OK ({time.time() - t0:.1f}s)")

        # [4/9] OAuth
        cb("[4/9] OAuth...")
        t0 = time.time()
        r = session.get(
            ru,
            headers={**bh, "Accept": "text/html,application/xhtml+xml",
                     "Origin": AUTH_BASE, "Referer": f"{CHAT_BASE}/"},
            impersonate=_random_imp(),
            timeout=30,
        )
        if "auth.openai.com" in r.url:
            cp = r.url.split("auth.openai.com")[-1]
        else:
            cp = r.url
        cb(f"  OK ({time.time() - t0:.1f}s)")

        if "/log-in" in cp:
            cb("  [提示] 账号已存在")
            return (False, "账号已存在")

        # [5/9] Email OTP
        cb("[5/9] 邮箱验证码...")
        t0 = time.time()
        code = self._poll_otp(email, timeout=120)
        if not code:
            cb("  [错误] 验证码获取失败（超时或未收到邮件）")
            return (False, "验证码获取失败")
        cb(f"  {code} ({time.time() - t0:.1f}s)")

        # [6/9] Validate email
        cb("[6/9] 验证邮箱...")
        t0 = time.time()
        r = session.post(
            f"{AUTH_BASE}/api/accounts/email-otp/validate",
            json={"code": code},
            headers={**bh, "Origin": AUTH_BASE,
                     "Referer": f"{AUTH_BASE}/email-verification",
                     "Content-Type": "application/json"},
            impersonate=_random_imp(),
            timeout=30,
        )
        if r.status_code != 200:
            cb(f"  [错误] 验证失败 [{r.status_code}]: {r.text[:200]}")
            return (False, f"验证失败 [{r.status_code}]: {r.text[:200]}")
        cb(f"  OK ({time.time() - t0:.1f}s)")
        continue_url = r.json().get("continue_url", "")

        # [7/9] Create account
        cb("[7/9] 创建账号...")
        t0 = time.time()
        ch = {**bh, "Origin": AUTH_BASE, "Referer": f"{AUTH_BASE}/about-you", "Content-Type": "application/json"}
        if sd.get("sentinel_token"):
            ch["openai-sentinel-token"] = sd["sentinel_token"]
        if sd.get("sentinel_so_token"):
            ch["openai-sentinel-so-token"] = sd["sentinel_so_token"]
        r = session.post(
            f"{AUTH_BASE}/api/accounts/create_account",
            json={"name": f"{fn} {ln}", "birthdate": bd},
            headers=ch,
            impersonate=_random_imp(),
            timeout=30,
        )
        if r.status_code != 200:
            cb(f"  [错误] 创建失败 [{r.status_code}]: {r.text[:200]}")
            return (False, f"创建失败 [{r.status_code}]: {r.text[:200]}")
        cb(f"  OK ({time.time() - t0:.1f}s)")
        continue_url = r.json().get("continue_url", "")

        # [8/9] Get session
        cb("[8/9] 建立会话...")
        t0 = time.time()
        cur = continue_url
        for _ in range(8):
            if not cur:
                break
            if cur.startswith("/"):
                cur = f"{CHAT_BASE}{cur}"
            r = session.get(
                cur,
                headers={**bh, "Accept": "text/html,application/xhtml+xml"},
                impersonate=_random_imp(),
                allow_redirects=False,
                timeout=30,
            )
            cur = r.headers.get("Location", "")

        si = None
        for _ in range(3):
            try:
                r = session.get(
                    f"{CHAT_BASE}/api/auth/session",
                    headers={**bh, "Accept": "application/json"},
                    impersonate=_random_imp(),
                    timeout=15,
                )
                if r.status_code == 200:
                    si = r.json()
                    break
            except Exception:
                time.sleep(1)

        cb(f"  OK ({time.time() - t0:.1f}s)")
        personal_at = si.get("accessToken", "") if si else ""
        personal_st = si.get("sessionToken", "") if si else ""

        # [8.5/9] OAuth PKCE 获取 refresh_token
        cb("[8.5/9] OAuth 获取 refresh_token...")
        t0 = time.time()
        rt = ""
        oauth_tokens = None
        try:
            oauth_tokens = self._oauth_get_tokens(session, email=email, oai_did=oai_did, log_cb=cb)
        except Exception as e:
            cb(f"  [警告] OAuth 异常: {e}")
        personal_id = ""
        if oauth_tokens:
            rt = oauth_tokens.get("refresh_token", "")
            if oauth_tokens.get("access_token"):
                personal_at = oauth_tokens["access_token"]
            personal_id = oauth_tokens.get("id_token", "")
            cb(f"  OK, refresh_token={'有' if rt else '无'} ({time.time() - t0:.1f}s)")
        else:
            cb(f"  [警告] 未获取到 refresh_token ({time.time() - t0:.1f}s)")

        # [9/9] Export JSON (no workspace)
        cb("[9/9] 导出 JSON...")
        if not personal_at:
            cb("  [警告] 未获取到 access_token，但仍会保存")

        ad = self.save_dir
        ad.mkdir(parents=True, exist_ok=True)
        cp_path = ad / f"{email}.json"
        cpa = {
            "type": "codex",
            "email": email,
            "password": password,
            "expired": "",
            "id_token": personal_id,
            "account_id": "",
            "disabled": False,
            "access_token": personal_at,
            "session_token": personal_st,
            "workspace_id": "",
            "last_refresh": datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M:%S +0800"),
            "refresh_token": rt,
        }
        json.dump(cpa, open(cp_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        cb(f"  保存: {cp_path}")

        # Upload to Codex2API
        if self.codexproxy_url and rt:
            ur = self._cp_upload(email, rt)
        else:
            ur = {}
        if ur.get("success"):
            cb("    上传成功")
        elif ur:
            cb("    上传: " + ur.get("message", ur.get("error", "?")))

        # Show pool status
        if self.codexproxy_url:
            t, a, rl, err = self._cp_status()
        else:
            t, a, rl, err = (0, 0, 0, 0)
        cb(f"\n  号池: 总计{t} 正常{a} 限流{rl} 异常{err}")
        cb(f"\n  === 注册完成 ===\n  {email}\n  {password}")

        return (True, {"email": email, "password": password, "access_token": personal_at, "session_token": personal_st, "refresh_token": rt})

    def batch(self, count=1, callback=None):
        cb = callback or self.log

        if self.mode in ACCOUNT_LIST_MODES and self.accounts:
            cb("\n" + "#" * 50)
            cb(f"  欧米freeGPT 批量注册 — {self.mode} 模式，{len(self.accounts)} 个邮箱")
            ok = 0
            for i, acct in enumerate(self.accounts):
                cb(f"\n--- 邮箱 {i + 1}/{len(self.accounts)}: {acct['email']} ---")
                if self.mode == "api":
                    self._current_api_url = acct.get("api_url", "")
                elif self.mode == "outlook_token":
                    self.imap_user = acct["email"]
                    self.imap_pass = acct["password"]
                    self.outlook_client_id = acct["client_id"]
                    self.outlook_refresh_token = acct["refresh_token"]
                    self._token_cache = {}
                else:
                    self.imap_user = acct["email"]
                    self.imap_pass = acct["password"]
                done = False
                for attempt in range(3):
                    success, data = self.register(acct["email"], callback=callback)
                    if success:
                        done = True
                        ok += 1
                        break
                    cb(f"  第 {attempt + 1} 次失败，重试...")
                    time.sleep(5)
                if not done:
                    cb("  该邮箱 3 次均失败，跳过")
            cb(f"\n  完成: {ok} 成功")
            return ok
        else:
            cb("\n" + "#" * 50)
            cb(f"  欧米freeGPT 批量注册 — {self.mode} 模式 x {count} 轮")
            ok = 0
            for i in range(count):
                cb(f"\n--- 第 {i + 1}/{count} 轮 ---")
                success, data = self.register(callback=callback)
                if success:
                    ok += 1
                else:
                    time.sleep(5)
            cb(f"\n  完成: {ok} 成功")
            return ok
