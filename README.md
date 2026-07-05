# M365 Copilot

M365 Copilot 非官方 Python 客户端 + OpenAI 兼容 API 服务器。

基于逆向工程实现，通过 SignalR-over-WebSocket 协议与 Microsoft 365 Copilot 后端通信，支持 Claude Sonnet、GPT-5.5 等多种模型。

> ⚠️ 非官方项目，基于 [cramt/m365-copilot-proxy](https://github.com/cramt/m365-copilot-proxy) 的反向工程成果。

---

## 目录

- [项目结构](#项目结构)
- [前置要求](#前置要求)
- [安装](#安装)
- [获取 Token](#获取-token)
- [CLI 聊天](#cli-聊天)
- [API 服务器](#api-服务器)
- [可用模型](#可用模型)
- [多轮上下文](#多轮上下文)
- [Token 刷新](#token-刷新)
- [技术原理](#技术原理)
- [参考](#参考)

---

## 项目结构

```
m365-copilot-client/
├── m365_copilot/          # Python 核心库
│   ├── __init__.py
│   ├── auth.py            # 认证（MSAL PKCE / 浏览器提取）
│   ├── signalr.py         # SignalR WebSocket 协议实现
│   ├── session.py         # 多轮对话会话管理
│   └── cli.py             # 交互式 CLI
├── api.py                 # OpenAI 兼容 API 服务器
├── gettoken.py            # MSAL PKCE 手动 token 兑换工具
├── requirements.txt
├── pyproject.toml
└── README.md
```

## 前置要求

- **Python 3.11+**
- **Microsoft 365 Copilot 订阅**（付费版，含 Copilot 的 E3/E5 等）
- 已安装依赖：`pip install -r requirements.txt`

## 安装

```bash
cd ~/projects/m365-copilot-client
pip install -r requirements.txt
```

依赖：msal, websockets, httpx, click（CLI）/ fastapi, uvicorn（API）

## 获取 Token

连接 M365 Copilot 需要 Sydney JWT token。三种方式：

### 方式 A：浏览器提取（推荐）

最快的方式，从已登录的浏览器直接获取 token。

1. 打开 https://m365.cloud.microsoft 并登录
2. F12 → **Network** 标签 → 过滤栏输入 `WS`（只显示 WebSocket）
3. 在聊天框发一条消息
4. 点那条 `wss://substrate.office.com/m365Copilot/Chathub/...` 的连接
5. 在 **Headers** 面板的 Request URL 中找到 `access_token=***` 参数
6. 复制 `eyJ...` 开头的整个 JWT token（约 3000+ 字符）

保存到文件：

```bash
echo 'eyJ...' > ~/.config/m365-copilot/token.txt
```

或设为环境变量：

```bash
export M365_TOKEN='eyJ...'
```

### 方式 B：MSAL PKCE（备用）

通过 OAuth 授权码流程获取 token（需要浏览器登录）。

```bash
# 第一步：生成登录链接
python3 gettoken.py gen

# 第二步：打开链接登录，复制跳转 URL，执行：
python3 gettoken.py ex "https://login.microsoftonline.com/common/oauth2/nativeclient?code=..."

# 成功后自动保存到 ~/.config/m365-copilot/token.txt
```

token 保存后，`api.py` 和 CLI 会自动读取。

### 方式 C：CLI 认证命令

```bash
python3 -m m365_copilot.cli auth --force
```

查看已保存的 token 信息：

```bash
python3 -m m365_copilot.cli info
```

## CLI 聊天

```bash
# 单次对话
python3 -m m365_copilot.cli chat --tone claude-sonnet --text "你好"

# 交互模式（支持多轮对话）
python3 -m m365_copilot.cli chat --tone claude-sonnet
```

交互模式支持的命令：

| 命令 | 功能 |
|------|------|
| `!reset` | 重置对话（开始新会话） |
| `!info` | 查看当前对话信息（轮次、会话 ID） |
| `!models` | 查看可用模型列表 |
| `!quit` | 退出 |

### 模型选择

```bash
# Claude Sonnet（Anthropic 模型）
python3 -m m365_copilot.cli chat --tone claude-sonnet

# 最新 GPT-5.5
python3 -m m365_copilot.cli chat --tone gpt-5.5

# 深度思考模式
python3 -m m365_copilot.cli chat --tone think-deeper

# 默认自动路由
python3 -m m365_copilot.cli chat --tone auto
```

### Verbose 模式

```bash
python3 -m m365_copilot.cli chat -v
```

显示详细的调试日志（SignalR 帧类型、参数等）。

## API 服务器

启动 OpenAI 兼容的 API 服务器：

```bash
# 默认端口 23100
python3 api.py --port 23100

# 详细日志
python3 api.py --port 23100 -v
```

### API 端点

#### `GET /v1/models`

返回可用模型列表。

#### `POST /v1/chat/completions`

OpenAI 兼容的聊天补全接口。支持 `stream: true`（SSE 流式）。

```bash
curl http://127.0.0.1:23100/v1/chat/completions \
  -H "Content-Type: application/json" \
  -d '{
    "model": "claude-sonnet",
    "messages": [{"role": "user", "content": "你好"}],
    "stream": false
  }'
```

#### `GET /health`

健康检查。

### 配置到第三方客户端

以 LobeChat 为例：

1. 添加自定义 AI Provider → OpenAI
2. 接口地址：`http://127.0.0.1:23100/v1`
3. API Key：任意填写（暂未做认证）
4. 模型：`claude-sonnet` / `gpt-5.5` / `auto` 等

## 可用模型

| 模型 ID | Tone | 说明 |
|---------|------|------|
| `auto` | `magic` | 默认，自动路由 GPT-5 级别 |
| `quick` | `Gpt_Quick` | 快速模式 |
| `think-deeper` | `Gpt_Reasoning` | 深度推理 |
| `claude-sonnet` | `Claude_Sonnet` | **Anthropic Claude Sonnet 4.6**（独立模型） |
| `claude-sonnet-4.6` | `Claude_Sonnet` | 同上 |
| `claude-sonnet-think-deeper` | `Claude_Sonnet_Reasoning` | Claude + 推理 |
| `claude-opus` | `Claude_Opus` | Claude Opus |
| `gpt-5.5` | `Gpt_5_5_Chat` | 最新 GPT 系列 |
| `gpt-5.5-quick` | `Gpt_5_5_Chat` | |
| `gpt-5.5-think-deeper` | `Gpt_5_5_Reasoning` | |
| `gpt-5.4` | `Gpt_5_4_Reasoning` | |
| `gpt-5.4-quick` | `Gpt_5_4_Quick` | |
| `gpt-5.3` | `Gpt_5_3_Quick` | |
| `gpt-5.2` | `Gpt_5_2_Quick` | |

Tone 由服务端验证，未知值会返回错误。"Claude" 开头的 tone 会路由到真正的 Anthropic Claude 模型。

## 多轮上下文

**CLI 模式：** 交互模式下自动保持上下文，每轮对话共享同一 `conversation_id`。

**API 模式：** 通过消息历史前缀哈希自动匹配同一会话。使用客户端时连续发送消息即可保持上下文。也可通过 `user` 字段或 `X-Conversation-Id` 标头指定会话。

会话池最大 100 个，超出时自动淘汰最旧会话。

## Token 刷新

M365 Copilot 的 JWT token 有效期约 1 小时。过期后需重新获取：

```bash
# 浏览器提取，或
python3 gettoken.py gen
python3 gettoken.py ex "URL"
```

API 服务器启动时会检查 token 有效性，过期后需重启。

## 技术原理

### 通信协议

- **端点：** `wss://substrate.office.com/m365Copilot/Chathub/{oid}@{tid}`
- **传输：** SignalR JSON 协议，帧以 `0x1E`（Record Separator）分隔
- **认证：** access_token 放在 WebSocket URL 的 query string 中
- **必备标头：** `Origin: https://m365.cloud.microsoft` + 浏览器 User-Agent

### 数据帧类型

| Type | 方向 | 说明 |
|------|------|------|
| 1 | 双向 | Update / Invocation / Metrics |
| 2 | 服务端 | StreamItem（回合结束 + 完整状态） |
| 3 | 服务端 | Completion（完成 + 可选的错误） |
| 4 | 客户端 | Invocation（Chat 请求） |
| 6 | 双向 | Ping / Pong |
| 7 | 服务端 | Close |

### 握手流程

1. 客户端发送 `{"protocol":"json","version":1}\x1E`
2. 服务端回复 `{}\x1E`
3. 客户端发送 Chat invocation（type:4） + Metrics 帧（type:1） 在同一 WS message 中
4. 服务端回复 type:1 update 帧（增量文本 + 消息快照）
5. 服务端发送 type:2 StreamItem 或 type:3 Completion 结束回合

### 常见问题

- **Negotiate 403：** 需要将 token 放在 WS URL query string 中，并带浏览器 Origin/User-Agent
- **Disengaged：** 工具数过多或 prompt 像 jailbreak 时触发，返回 `messageType: "Disengaged"`
- **InternalError：** 第二轮对话失败，可能是 plugins 在非首轮不应出现（已在代码中修复）

## 参考

- [NodeLoc 逆向分析帖](https://www.nodeloc.com/t/topic/96870) — 原始 SignalR 协议分析
- [cramt/m365-copilot-proxy](https://github.com/cramt/m365-copilot-proxy) — TypeScript 反向代理实现（主要参考）
- [vsakkas/sydney.py](https://github.com/vsakkas/sydney.py) — Bing Chat / Sydney Python 客户端

## 许可

MIT
