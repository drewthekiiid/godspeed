#!/usr/bin/env bash
# Unified AWS Cloud Vault CLI for VS Code Cloud / GitHub Copilot Coding Agent / Codespaces.
# Pulls (and pushes) MCP server configs, 160+ global API keys, per-repo .env files,
# custom Copilot agents (.agent.md), and 98 skills from AWS Secrets Manager & private S3.
#
# Usage:
#   bash ops/cloud-agent/vault.sh pull [--force] [repo-slug]   Restore MCPs, global keys, agents/skills, and repo .env files
#   bash ops/cloud-agent/vault.sh list                         List all MCP servers, repo .env bundles, and global API key names
#   bash ops/cloud-agent/vault.sh get <KEY_NAME>               Print the value of a specific API key from the vault
#   bash ops/cloud-agent/vault.sh env <repo-slug> [--force]    Materialize a specific repo's .env files into current directory
#   bash ops/cloud-agent/vault.sh status                       Verify vault reachability in AWS Secrets Manager & S3
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
# shellcheck source=ops/cloud-agent/config.sh
source "$SCRIPT_DIR/config.sh"

export VAULT_BUCKET="${VAULT_BUCKET:-drewthekiiid-cloud-vault-703890345115}"
export VAULT_SECRET_ID="${VAULT_SECRET_ID:-vscode-cloud/master-vault}"
export VAULT_REGION="${VAULT_REGION:-us-east-1}"

CMD="${1:-pull}"
shift || true

FORCE="0"
TARGET_SLUG=""
while [ $# -gt 0 ]; do
  case "$1" in
    -f|--force)
      FORCE="1"
      shift
      ;;
    *)
      TARGET_SLUG="$1"
      shift
      ;;
  esac
done

if [ -z "$TARGET_SLUG" ]; then
  TARGET_SLUG="$(basename "$REPO_ROOT")"
fi

case "$CMD" in
  status)
    aws secretsmanager describe-secret --secret-id "$VAULT_SECRET_ID" --region "$VAULT_REGION" --no-cli-pager --query '[Name,ARN,LastChangedDate]' --output text
    aws s3 ls "s3://${VAULT_BUCKET}/vault/" --region "$VAULT_REGION"
    ;;

  list)
    TMP_JSON="$(mktemp)"
    trap 'rm -f "$TMP_JSON"' EXIT
    aws s3 cp "s3://${VAULT_BUCKET}/vault/env-bundle.json" "$TMP_JSON" --region "$VAULT_REGION" --only-show-errors
    python3 - "$TMP_JSON" <<'PY'
import sys, json
with open(sys.argv[1], "r", encoding="utf-8") as f:
    d = json.load(f)
servers = list((d.get("mcp_config", {}).get("mcpServers") or {}).keys())
gkeys = sorted((d.get("global_env") or {}).keys())
bundles = d.get("repo_env_bundles") or {}
print(f"=== MCP Servers ({len(servers)}) ===")
print(", ".join(servers))
print(f"\n=== Repo .env Bundles ({len(bundles)}) ===")
for k in sorted(bundles.keys()):
    if bundles[k]:
        print(f"  {k:<24}: {list(bundles[k].keys())}")
print(f"\n=== Global API Keys ({len(gkeys)}) ===")
print(", ".join(gkeys))
PY
    ;;

  get)
    if [ -z "$TARGET_SLUG" ]; then
      echo "Usage: vault.sh get <KEY_NAME>" >&2
      exit 1
    fi
    SECRET_STR="$(aws secretsmanager get-secret-value --secret-id "$VAULT_SECRET_ID" --region "$VAULT_REGION" --no-cli-pager --query SecretString --output text)"
    python3 - "$TARGET_SLUG" "$SECRET_STR" <<'PY'
import sys, json
key, raw = sys.argv[1], sys.argv[2]
d = json.loads(raw)
genv = d.get("global_env") or {}
if key in genv:
    print(genv[key])
else:
    sys.stderr.write(f"Key {key!r} not found in global_env\n")
    sys.exit(1)
PY
    ;;

  pull|env)
    TMP_JSON="$(mktemp)"
    trap 'rm -f "$TMP_JSON"' EXIT
    if ! aws s3 cp "s3://${VAULT_BUCKET}/vault/env-bundle.json" "$TMP_JSON" --region "$VAULT_REGION" --only-show-errors 2>/dev/null; then
      echo "WARN: Could not fetch s3://${VAULT_BUCKET}/vault/env-bundle.json — falling back to Secrets Manager $VAULT_SECRET_ID" >&2
      aws secretsmanager get-secret-value --secret-id "$VAULT_SECRET_ID" --region "$VAULT_REGION" --no-cli-pager --query SecretString --output text > "$TMP_JSON"
    fi

    python3 - "$TMP_JSON" "$REPO_ROOT" "$TARGET_SLUG" "$FORCE" "$CMD" <<'PY'
import sys, os, json, shlex

json_path, repo_root, slug, force_str, mode = sys.argv[1:6]
force = (force_str == "1")
home = os.path.expanduser("~")

with open(json_path, "r", encoding="utf-8") as f:
    data = json.load(f)

if mode == "pull":
    # 1. Restore MCP configs (~/.copilot/mcp-config.json and ~/.cursor/mcp.json)
    mcp_cfg = data.get("mcp_config")
    if mcp_cfg:
        raw_mcp = json.dumps(mcp_cfg, indent=2)
        raw_mcp = raw_mcp.replace("/Users/thekiiid", home)
        for rel_mcp in [".copilot/mcp-config.json", ".cursor/mcp.json"]:
            mcp_dest = os.path.join(home, rel_mcp)
            os.makedirs(os.path.dirname(mcp_dest), exist_ok=True)
            if force or not os.path.exists(mcp_dest):
                with open(mcp_dest, "w", encoding="utf-8") as f:
                    f.write(raw_mcp + "\n")
                os.chmod(mcp_dest, 0o600)

    # 2. Restore ~/.cursor/ghl.env and ~/.cursor/namecheap.env
    for rel_df, content in (data.get("dot_files") or {}).items():
        df_dest = os.path.join(home, rel_df)
        os.makedirs(os.path.dirname(df_dest), exist_ok=True)
        if force or not os.path.exists(df_dest):
            with open(df_dest, "w", encoding="utf-8") as f:
                f.write(content)
            os.chmod(df_dest, 0o600)

    # 3. Write ~/.config/cloud-vault/global.env (160+ API keys)
    global_env = data.get("global_env") or {}
    if global_env:
        gdir = os.path.join(home, ".config", "cloud-vault")
        os.makedirs(gdir, exist_ok=True)
        gfile = os.path.join(gdir, "global.env")
        with open(gfile, "w", encoding="utf-8") as f:
            f.write("# Auto-generated by ops/cloud-agent/vault.sh — DO NOT COMMIT\n")
            for k in sorted(global_env.keys()):
                v = str(global_env[k]).replace("\n", "")
                f.write(f"export {k}={shlex.quote(v)}\n")
        os.chmod(gfile, 0o600)

# 4. Restore repo-specific .env files
bundles = data.get("repo_env_bundles") or {}
# Match exact slug or case-insensitive slug
matched_bundle = bundles.get(slug)
if matched_bundle is None:
    for k, v in bundles.items():
        if k.lower() == slug.lower() or k.lower().replace("'", "") == slug.lower().replace("'", ""):
            matched_bundle = v
            break

restored = []
if matched_bundle:
    for rel_path, content in matched_bundle.items():
        dest = os.path.join(repo_root, rel_path)
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        if force or not os.path.exists(dest):
            with open(dest, "w", encoding="utf-8") as f:
                f.write(content)
            os.chmod(dest, 0o600)
            restored.append(rel_path)

print(f"vault: MCP config + {len(data.get('global_env') or {})} global API keys synced; repo={slug} restored_files={restored}")
PY

    # 5. Restore ~/.copilot/agents, ~/.copilot/skills, and ~/.cursor/tools/org-mcp if missing (or --force)
    if [ "$CMD" = "pull" ]; then
      if [ "$FORCE" = "1" ] || [ ! -d "$HOME/.copilot/skills/ce-work" ] || [ ! -f "$HOME/.cursor/tools/org-mcp/server.mjs" ]; then
        TMP_TAR="$(mktemp).tar.gz"
        if aws s3 cp "s3://${VAULT_BUCKET}/vault/copilot-agents-skills.tar.gz" "$TMP_TAR" --region "$VAULT_REGION" --only-show-errors 2>/dev/null; then
          tar -xzf "$TMP_TAR" -C "$HOME"
          rm -f "$TMP_TAR"
          if [ -f "$HOME/.cursor/tools/org-mcp/package.json" ] && [ ! -d "$HOME/.cursor/tools/org-mcp/node_modules" ] && command -v npm >/dev/null 2>&1; then
            (cd "$HOME/.cursor/tools/org-mcp" && npm install --silent >/dev/null 2>&1 || true)
          fi
          echo "vault: restored ~/.copilot/agents (7), ~/.copilot/skills (98), and ~/.cursor/tools/org-mcp"
        else
          rm -f "$TMP_TAR"
        fi
      fi

      # Ensure shell profiles on Linux Cloud VMs auto-source ~/.config/cloud-vault/global.env
      if [ "$(uname -s)" = "Linux" ] && [ -f "$HOME/.config/cloud-vault/global.env" ]; then
        for rc in "$HOME/.bashrc" "$HOME/.profile" "$HOME/.zshrc"; do
          touch "$rc"
          grep -qF '.config/cloud-vault/global.env' "$rc" 2>/dev/null || \
            echo '[ -f "$HOME/.config/cloud-vault/global.env" ] && . "$HOME/.config/cloud-vault/global.env"' >> "$rc"
        done
      fi
    fi
    ;;

  *)
    echo "Usage: vault.sh [pull|list|get|env|status] [--force] [repo-or-key]" >&2
    exit 1
    ;;
esac
