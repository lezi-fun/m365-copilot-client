"""
M365 Copilot Token Exchange Tool

用法:
  python3 token.py gen           # 生成登录链接
  python3 token.py ex "URL"      # 用跳转 URL 换 token
"""
import base64, hashlib, json, os, secrets, subprocess, sys, urllib.parse
from pathlib import Path

CLIENT_ID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1"
REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"
SCOPES = ["https://substrate.office.com/sydney/M365Chat.Read",
          "https://substrate.office.com/sydney/sydney.readwrite"]

MODE = sys.argv[1] if len(sys.argv) > 1 else ""

if MODE == "gen":
    v = base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()
    c = base64.urlsafe_b64encode(hashlib.sha256(v.encode()).digest()).rstrip(b"=").decode()
    Path("/tmp/msal_v.txt").write_text(v)
    q = urllib.parse.urlencode({"client_id": CLIENT_ID, "response_type": "code",
        "redirect_uri": REDIRECT_URI, "scope": " ".join(SCOPES),
        "response_mode": "query", "code_challenge": c, "code_challenge_method": "S256"})
    print("=" * 60)
    print("打开链接登录，复制跳转 URL")
    print("=" * 60)
    print(f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?{q}")
    print()
    print('然后: python3 token.py ex "粘贴的URL"')

elif MODE == "ex":
    if not os.path.exists("/tmp/msal_v.txt"):
        print("先运行 python3 token.py gen"); sys.exit(1)
    if len(sys.argv) < 3:
        print('用法: python3 token.py ex "URL"'); sys.exit(1)

    v = Path("/tmp/msal_v.txt").read_text().strip()
    raw = sys.argv[2]

    parsed = urllib.parse.urlparse(raw)
    code = urllib.parse.parse_qs(parsed.query).get("code", [None])[0]
    if not code:
        code = urllib.parse.parse_qs(parsed.fragment).get("code", [None])[0]
    if not code:
        print("❌ URL 中没有 code"); sys.exit(1)

    print(f"✅ Auth code: {len(code)} chars")

    data = urllib.parse.urlencode({"client_id": CLIENT_ID, "code": code,
        "redirect_uri": REDIRECT_URI, "grant_type": "authorization_code",
        "code_verifier": v, "scope": " ".join(SCOPES)})

    # 尝试直连，失败则走代理（通过 HTTP_PROXY 环境变量配置）
    r = subprocess.run(["curl", "-s", "--max-time", "15",
        "https://login.microsoftonline.com/common/oauth2/v2.0/token", "-d", data],
        capture_output=True, text=True, timeout=20)
    if r.returncode != 0 or not r.stdout.strip():
        proxy = os.environ.get("HTTP_PROXY") or os.environ.get("HTTPS_PROXY")
        if proxy:
            r = subprocess.run(["curl", "-s", "--max-time", "15",
                "--proxy", proxy,
                "https://login.microsoftonline.com/common/oauth2/v2.0/token", "-d", data],
                capture_output=True, text=True, timeout=20)

    try:
        resp = json.loads(r.stdout)
    except json.JSONDecodeError:
        print(f"❌ curl 失败: {r.stdout[:300]}"); sys.exit(1)

    if "access_token" in resp:
        t = resp["access_token"]
        p = t.split(".")[1]
        pad = 4 - (len(p) % 4)
        if pad != 4: p += "=" * pad
        cl = json.loads(base64.urlsafe_b64decode(p))
        print(f"\n✅ {cl.get('name','?')} | {resp.get('expires_in','?')}s")
        dst = Path.home() / ".config" / "m365-copilot" / "token.txt"
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_text(t)
        print(f"   已保存 → {dst}")
    else:
        print(f"❌ {json.dumps(resp, indent=2)[:500]}")
else:
    print("用法:")
    print("  python3 token.py gen           # 生成链接")
    print('  python3 token.py ex "URL"      # 换 token')
