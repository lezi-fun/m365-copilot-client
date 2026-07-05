"""
M365 Copilot CLI — interactive chat client for M365 Copilot via SignalR.

Usage:
    m365-copilot auth           # Authenticate and save token
    m365-copilot info           # Show token info
    m365-copilot chat           # Interactive chat session
    m365-copilot chat --text "Hello"  # One-shot message
    m365-copilot list-models    # Show available model tones
"""

import asyncio
import json
import logging
import sys
from pathlib import Path
from typing import Optional

import click

from .auth import get_token, get_token_info, CONFIG_DIR as AUTH_CONFIG_DIR
from .session import CopilotSession, MODEL_TONES

# Configure logging
logging.basicConfig(
    level=logging.WARNING,
    format="%(levelname)s %(name)s: %(message)s",
)

TOKEN_FILE = AUTH_CONFIG_DIR / "token.txt"


def _load_token() -> Optional[str]:
    """Load token from cache or auth."""
    # Try file cache first
    if TOKEN_FILE.exists():
        try:
            with open(TOKEN_FILE) as f:
                token = f.read().strip()
                if token:
                    return token
        except Exception:
            pass

    return None


def _save_token(token: str):
    AUTH_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(TOKEN_FILE, "w") as f:
        f.write(token)


def _clear_token():
    if TOKEN_FILE.exists():
        TOKEN_FILE.unlink()


def _print_response(resp, show_raw: bool = False):
    """Pretty-print a ChatResponse."""
    if resp.disengaged:
        click.secho("\n[!] Disengaged — M365 refused to answer", fg="red", bold=True)
    elif resp.error:
        click.secho(f"\n[!] Error: {resp.error}", fg="red")
    elif resp.has_content:
        click.echo(f"\n{resp.text}")
    else:
        click.secho("\n[!] Empty response", fg="yellow")

    # Metadata
    if resp.throttle:
        click.echo(f"  [quota: {resp.throttle['current']}/{resp.throttle['max']}]")
    if resp.content_origin:
        click.echo(f"  [origin: {resp.content_origin}]")
    if resp.turn_count is not None:
        click.echo(f"  [turn: {resp.turn_count}]")
    if resp.turn_state:
        click.echo(f"  [state: {resp.turn_state}]")
    if resp.scores and any(v > 1e-10 for v in resp.scores.values()):
        click.echo(f"  [scores: {resp.scores}]")
    if resp.message_type:
        click.echo(f"  [msg_type: {resp.message_type}]")

    if show_raw:
        click.echo("\n--- Raw Frames ---")
        for i, frame in enumerate(resp.raw_frames):
            click.echo(f"  [{i}] type={frame.get('type')} target={frame.get('target', 'N/A')}")
            click.echo(f"      {json.dumps(frame, indent=2)[:500]}")


@click.group()
def cli():
    """M365 Copilot — Unofficial Python client."""


@cli.command()
@click.option("--force", is_flag=True, help="Force re-authentication")
@click.option("--method", type=click.Choice(["interactive", "device_code", "manual"]), default="interactive",
              help="Auth method: interactive (browser) or device_code (URL+code)")
def auth(force: bool, method: str):
    """Authenticate with Microsoft and save the Sydney token."""
    click.echo("Authenticating with Microsoft 365 Copilot...")
    token = get_token(force_refresh=force, method=method)
    if token:
        info = get_token_info(token)
        _save_token(token)
        click.secho("\n✅ Authentication successful!", fg="green", bold=True)
        click.echo(f"   User: {info['name']} ({info['upn']})")
        click.echo(f"   Expires in: {info['expires_in']} seconds")
        click.echo(f"   Audience: {info['aud']}")
    else:
        click.secho("\n❌ Authentication failed", fg="red", bold=True)
        sys.exit(1)


@cli.command()
def info():
    """Show current token information."""
    token = _load_token()
    if not token:
        click.secho("Not authenticated. Run: m365-copilot auth", fg="yellow")
        return

    try:
        info = get_token_info(token)
        click.echo("Token Info:")
        click.echo(f"  User:       {info['name']} ({info['upn']})")
        click.echo(f"  Object ID:  {info['oid']}")
        click.echo(f"  Tenant ID:  {info['tid']}")
        click.echo(f"  Audience:   {info['aud']}")
        click.echo(f"  Expires in: {info['expires_in']} seconds")
        if info["expires_in"] < 0:
            click.secho("  ⚠️  Token expired! Re-authenticate with: m365-copilot auth --force", fg="red")
    except Exception as e:
        click.secho(f"Error decoding token: {e}", fg="red")
        sys.exit(1)


@cli.command()
def list_models():
    """Show available model tones."""
    click.echo("Available model tones:\n")
    click.echo(f"  {'Model Name':28s} {'Tone'}")
    click.echo(f"  {'-'*28} {'-'*25}")
    for name, tone in sorted(MODEL_TONES.items(), key=lambda x: x[1]):
        click.echo(f"  {name:28s} {tone}")
    click.echo()
    click.echo("Usage: --tone <model_name> in chat command")


@cli.command()
@click.option("--text", "-t", help="One-shot message (omit for interactive mode)")
@click.option("--tone", default="auto", help="Model tone (see list-models)")
@click.option("--agent", help="Copilot Studio agent ID for tool calling")
@click.option("--code-interpreter", is_flag=True, help="Enable server-side Python execution")
@click.option("--raw", is_flag=True, help="Show raw SignalR frames")
@click.option("--verbose", "-v", is_flag=True, help="Verbose logging")
@click.option("--token", envvar="M365_TOKEN", help="Sydney token (or use cached auth)")
def chat(text: Optional[str], tone: str, agent: Optional[str],
         code_interpreter: bool, raw: bool, verbose: bool, token: Optional[str]):
    """Chat with M365 Copilot."""
    if verbose:
        logging.getLogger("m365").setLevel(logging.DEBUG)

    # Resolve token
    token = token or _load_token()
    if not token:
        click.secho("No token. Run: m365-copilot auth", fg="red")
        sys.exit(1)

    # Check token expiry
    info = None
    try:
        info = get_token_info(token)
        if info["expires_in"] < 60:
            click.secho("Token expired or expiring soon. Re-run: m365-copilot auth --force", fg="red")
            sys.exit(1)
    except Exception as e:
        click.secho(f"Invalid token: {e}", fg="red")
        sys.exit(1)

    # Resolve tone
    actual_tone = MODEL_TONES.get(tone, tone)

    click.echo(f"Connected as {info['name']} | model: {tone} → tone: {actual_tone}")
    if agent:
        click.echo(f"Agent: {agent}")

    async def run():
        session = CopilotSession(
            token=token,
            agent_id=agent or None,
            tone=actual_tone,
            enable_code_interpreter=code_interpreter,
        )

        def on_chunk(chunk: str):
            click.echo(chunk, nl=False)
            sys.stdout.flush()

        # One-shot mode
        if text:
            click.echo("\n" + "=" * 50)
            resp = await session.send(text, on_delta=None)
            click.echo(resp.text)
            _print_response(resp, show_raw=raw)
            return

        # Interactive mode
        click.echo("\nInteractive mode. Type your messages below.")
        click.echo("Commands: !reset, !models, !info, !quit\n")

        try:
            while True:
                line = click.prompt("You", prompt_suffix="> ")

                if line.startswith("!"):
                    cmd = line[1:].strip().lower()
                    if cmd in ("quit", "exit", "q"):
                        break
                    elif cmd == "reset":
                        session.reset()
                        click.secho("Conversation reset.", fg="blue")
                        continue
                    elif cmd == "models":
                        list_models()
                        continue
                    elif cmd == "info":
                        click.echo(f"  Turn: {session.turn_count}")
                        click.echo(f"  Conv: {session.conversation_id}")
                        click.echo(f"  Tone: {session.tone}")
                        continue
                    else:
                        click.echo(f"Unknown command: !{cmd}")
                        continue

                click.echo()
                resp = await session.send(line, on_delta=on_chunk)
                click.echo()
                _print_response(resp, show_raw=raw)

        except (EOFError, KeyboardInterrupt):
            click.echo("\nBye!")

    asyncio.run(run())


@cli.command()
def reset_auth():
    """Clear cached authentication and token."""
    _clear_token()
    # Also remove MSAL cache
    cache_file = Path(AUTH_CONFIG_DIR / "msal-cache.json")
    if cache_file.exists():
        cache_file.unlink()
    click.secho("Authentication cache cleared.", fg="green")


if __name__ == "__main__":
    cli()
