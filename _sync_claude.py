import json
import os
import sqlite3
from pathlib import Path

home = Path(os.environ["USERPROFILE"])
db = home / ".cc-switch" / "cc-switch.db"
out = home / ".claude" / "settings.json"

c = sqlite3.connect(db)
row = c.execute(
    "SELECT settings_config FROM providers WHERE app_type='claude' AND is_current=1"
).fetchone()
if not row:
    raise SystemExit("No current Claude provider")

config = json.loads(row[0])
env = config.setdefault("env", {})
# Some Claude Code paths only read ANTHROPIC_API_KEY
if env.get("ANTHROPIC_AUTH_TOKEN") and not env.get("ANTHROPIC_API_KEY"):
    env["ANTHROPIC_API_KEY"] = env["ANTHROPIC_AUTH_TOKEN"]

out.parent.mkdir(parents=True, exist_ok=True)
existing = {}
if out.exists():
    try:
        existing = json.loads(out.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        pass
for k, v in config.items():
    if k == "env":
        existing.setdefault("env", {}).update(v)
    else:
        existing[k] = v

out.write_text(json.dumps(existing, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

env.setdefault("CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC", "1")
env.setdefault("DISABLE_AUTOUPDATER", "1")

for key in (
    "ANTHROPIC_BASE_URL",
    "ANTHROPIC_API_KEY",
    "ANTHROPIC_AUTH_TOKEN",
    "CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC",
    "DISABLE_AUTOUPDATER",
):
    val = env.get(key)
    if val:
        os.environ[key] = val
        # Persist for all future terminals
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Environment",
            0,
            winreg.KEY_SET_VALUE,
        ) as hk:
            winreg.SetValueEx(hk, key, 0, winreg.REG_EXPAND_SZ, val)

print("Synced", out)
print("BASE_URL", env.get("ANTHROPIC_BASE_URL"))
