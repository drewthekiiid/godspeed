# AGENTS.md

## Cursor Cloud Agents — AWS SSM & SSH Fleet Access

This repository is configured with `ops/cloud-agent/` and `.cursor/environment.json` so Cursor Cloud Agents (`cursor.com/agents`) have full outbound AWS SSM (`SendCommand` + `StartSession`) and SSH/SCP/rsync-over-SSM access to the AWS EC2 fleet.

- **Default box for this repo:** `famp-research` (`i-07b4eecaa41cb53a8` in `us-east-1`, remote dir `/opt/madoffs-fantasy`)

```bash
bash ops/cloud-agent/box.sh status                          # Verify AWS STS identity + default box (famp-research)
bash ops/cloud-agent/box.sh fleet                           # Check EC2 & SSM status of all 9 fleet boxes
bash ops/cloud-agent/box.sh exec -- '<cmd>'                 # Run command via SSM on default box (famp-research)
bash ops/cloud-agent/box.sh --box <alias> exec -- '<cmd>'   # Run command via SSM on any fleet box
ssh famp-research                                                   # SSH over SSM tunnel (works with scp/rsync too)
bash ops/cloud-agent/box.sh --box <alias> start             # Boot a stopped on-demand box and wait for SSM Online
```

| Alias | Instance ID | Region | Role / Path |
|-------|-------------|--------|-------------|
| `polygamy` | `i-01c97deadad199159` | `us-east-1` | Poly-gamy live/shadow node (`/opt/polygamy`) |
| `famp-sniper` | `i-002255402e5ed0044` | `eu-west-1` | Poly-gamy / FAMP Dublin sniper (`54.220.211.64`) |
| `bpumper` | `i-001af733ab7600594` | `eu-central-1` | BPumper Frankfurt node (`/opt/bpumper`, `/var/lib/bpumper`) |
| `famp-research` | `i-07b4eecaa41cb53a8` | `us-east-1` | FAMP & Cryptologist research (`r7i.8xlarge`, `/opt/madoffs-fantasy`) |
| `famp-hl-live` | `i-0de47d4dca7713d99` | `ap-northeast-1` | Hyperliquid live Tokyo node (`/opt/madoffs-fantasy`) |
| `famp-hl-pilot` | `i-041be9b38ad6e05e7` | `ap-northeast-1` | Hyperliquid pilot & mainnet shadows (`famp-testmp`) |
| `famp-drift` | `i-0cefa97eec12f66bf` | `eu-central-1` | Drift Frankfurt harness (on-demand) |
| `famp-alpaca` | `i-007a1960d1581eef0` | `us-east-1` | Alpaca Virginia harness (on-demand) |
| `paperroute-box` | `i-08074a84c7478855e` | `us-east-1` | PaperRoute devbox (`/home/ubuntu/paperroute`, on-demand) |

## AWS Cloud Vault (MCP Servers, 162 API Keys, Repo `.env`s, 7 Agents & 98 Skills)

`install-aws-tools.sh` automatically runs `bash ops/cloud-agent/vault.sh pull` on Cloud VM boot to restore:
- `~/.copilot/mcp-config.json` & `~/.cursor/mcp.json` (all 19 MCP server configs)
- `~/.config/cloud-vault/global.env` (162 global API keys auto-sourced in shell) + `~/.cursor/ghl.env` + `~/.cursor/namecheap.env`
- Repo-local `.env` / `.env.local` / `ops/*.env` files for the current workspace
- `~/.copilot/agents/` (7 custom agents) + `~/.copilot/skills/` (98 skills) + `~/.cursor/tools/org-mcp`

```bash
bash ops/cloud-agent/vault.sh pull                  # Restore MCPs, global keys, agents/skills & current repo .env
bash ops/cloud-agent/vault.sh list                  # List all 19 MCPs, repo .env bundles, and 162 API key names
bash ops/cloud-agent/vault.sh get <KEY_NAME>        # Fetch a specific API key on demand (e.g., OPENAI_API_KEY, GHL_PIT)
bash ops/cloud-agent/vault.sh env <repo-slug>       # Materialize another repo's .env bundle on demand
```
