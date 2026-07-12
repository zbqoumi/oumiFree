# 欧米freeGPT注册机

> ChatGPT 账号批量自动注册工具，支持 Gmail / Outlook / API 接码等多种模式，多线程并发注册，注册后直接导出 JSON。

## 功能特性

- **多模式注册**：Gmail、Outlook、API 接码、Outlook 令牌、Token 生成
- **多线程并发**：可自定义并发线程数（1-20），大幅提升注册效率
- **自动收码**：Gmail IMAP / Outlook Graph API / Outlook IMAP XOAUTH2 / 第三方 API 接码
- **Sentinel Token**：内置 Playwright + Chrome 自动提取 Cloudflare Sentinel Token
- **TLS 指纹伪装**：使用 curl_cffi 模拟 Chrome 浏览器指纹，绕过反爬检测
- **自动上传**：注册成功后可选自动上传到 Codex2API 号池
- **OAuth PKCE**：自动获取 refresh_token，导出可直接使用的完整账号信息
- **暗色主题 GUI**：美观的 Tkinter 界面，实时日志输出

## 快速启动

```bash
cd free
chmod +x run.sh
./run.sh
```

或手动安装：

```bash
cd free
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python3 -m playwright install chromium
python3 start_blackcat.py
```

## 五种模式

| 模式 | 说明 | 输入格式 |
|------|------|----------|
| **Gmail** | 使用 Gmail 应用专用密码，自动 IMAP 收验证码 | 邮箱地址 + 应用专用密码 |
| **Outlook** | 通过 OAuth2 登录 Outlook，手动输入验证码 | 邮箱地址 + 应用专用密码 |
| **API 接码** | 批量邮箱 + 第三方 API 接码平台自动轮询 | 每行：`邮箱----API链接` |
| **Outlook 令牌** | 使用 Outlook refresh_token 自动 Graph/IMAP 接码 | 每行：`邮箱----密码----client_id----refresh_token` |
| **Token 生成** | 通过 Device Code Flow 批量获取 Outlook refresh_token | 每行：`邮箱----密码` |

## 注册控制

| 参数 | 说明 |
|------|------|
| **注册数量** | 批量注册的总次数（Gmail/Outlook 模式）或轮数（列表模式） |
| **并发线程** | 同时运行的注册线程数（1-20），线程数越大注册越快 |
| **开始注册** | 启动注册流程 |
| **停止** | 停止当前注册任务（优雅停止，等待当前任务完成） |

## 配置文件

### `blackcat_config.json`（自动生成）

保存界面配置，下次启动自动加载：

```json
{
    "mode": "gmail",
    "email": "your@gmail.com",
    "password": "your-app-password",
    "cp_url": "http://your-codex2api:port",
    "cp_key": "your-admin-key",
    "count": 10,
    "threads": 3,
    "proxy": "127.0.0.1:7897",
    "upload": true,
    "save_dir": "/path/to/codex",
    "api_list": ""
}
```

### `config.json`（全局配置）

```json
{
    "chatgpt": {
        "auth_base_url": "https://auth.openai.com",
        "chat_base_url": "https://chatgpt.com",
        "chat_web_client_id": "app_X8zY6vW2pQ9tR3dE7nK1jL5gH",
        "user_agent_chrome": "Mozilla/5.0 ..."
    },
    "http": {
        "proxy": {
            "default": "http://127.0.0.1:7897"
        }
    }
}
```

## 注册流程（9 步）

```
[1/9] Sentinel Token 提取（Playwright + Chrome 无头浏览器）
[2/9] 获取 CSRF Token
[3/9] 发起注册请求
[4/9] OAuth 授权跳转
[5/9] 邮箱验证码获取（IMAP / Graph API / 第三方 API）
[6/9] 验证邮箱
[7/9] 创建账号（随机姓名 + 生日）
[8/9] 建立会话 + OAuth PKCE 获取 refresh_token
[9/9] 导出 JSON（保存到本地 + 可选上传 Codex2API）
```

## 导出格式

注册成功后，每个账号保存为独立 JSON 文件：

```json
{
    "type": "codex",
    "email": "user+abc123@gmail.com",
    "password": "Xy7k3mNq!A1",
    "access_token": "eyJ...",
    "session_token": "...",
    "refresh_token": "...",
    "id_token": "eyJ...",
    "workspace_id": "",
    "last_refresh": "2025-07-12 18:00:00 +0800"
}
```

## 文件结构

```
free/
├── start_blackcat.py              # 主入口（Tkinter GUI）
├── sentinel.py                    # Sentinel Token 提取（Playwright + Chrome）
├── config.json                    # 全局配置（代理、URL 等）
├── blackcat_config.json           # 界面配置（自动生成）
├── requirements.txt               # Python 依赖
├── run.sh                         # 一键启动脚本
├── codex/                         # 注册结果输出目录
│   └── xxx@gmail.com.json
└── gmailreg/
    ├── __init__.py
    └── blackcat_engine.py         # 注册引擎核心
```

## 依赖

| 依赖 | 用途 |
|------|------|
| **Python 3.11+** | 运行环境 |
| **Google Chrome** | Sentinel Token 提取（Playwright 驱动） |
| **curl_cffi** | TLS 指纹伪装 HTTP 请求 |
| **playwright** | 浏览器自动化（Sentinel 提取） |
| **msal** | Microsoft Authentication Library（Outlook OAuth2） |

## 代理配置

工具支持 HTTP / SOCKS5 代理，在 `config.json` 中配置：

```json
{
    "http": {
        "proxy": {
            "default": "http://127.0.0.1:7897"
        }
    }
}
```

支持的代理格式：
- `http://host:port`
- `http://user:pass@host:port`
- `socks5://host:port`

## 注意事项

- **网络环境**：需要能访问 `auth.openai.com` 和 `chatgpt.com`，建议使用代理
- **Chrome 浏览器**：Sentinel Token 提取依赖本地 Chrome，请确保已安装
- **Gmail 应用专用密码**：不是 Gmail 登录密码，需在 Google 账户中生成
- **Outlook 令牌**：需要有效的 refresh_token，可通过 Token 生成模式获取
- **并发线程**：建议根据网络和代理情况设置，过高可能导致注册失败率上升
