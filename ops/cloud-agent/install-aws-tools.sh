#!/usr/bin/env bash
# Install AWS CLI v2 + Session Manager plugin + SSH-over-SSM fleet config on Cursor Cloud Agent VMs.
# Idempotent — safe to run on every `install` / `start` / update.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export PATH="$HOME/.local/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"
mkdir -p "$HOME/.local/bin"

if [ "$(uname -s)" = "Linux" ]; then
  MISSING_PKGS=()
  for bin_pkg in "ssh:openssh-client" "rsync:rsync" "jq:jq" "curl:curl" "unzip:unzip"; do
    bin="${bin_pkg%%:*}"
    pkg="${bin_pkg##*:}"
    if ! command -v "$bin" >/dev/null 2>&1; then
      MISSING_PKGS+=("$pkg")
    fi
  done
  if [ "${#MISSING_PKGS[@]}" -gt 0 ] && command -v apt-get >/dev/null 2>&1; then
    sudo apt-get update -qq && sudo apt-get install -y -qq "${MISSING_PKGS[@]}"
  fi
fi

if command -v aws >/dev/null 2>&1; then
  echo "aws: already installed ($(aws --version 2>&1 | head -1))"
else
  ARCH="$(uname -m)"
  case "$ARCH" in
    aarch64|arm64) PKG="awscli-exe-linux-aarch64.zip" ;;
    x86_64|amd64)  PKG="awscli-exe-linux-x86_64.zip" ;;
    *) echo "unsupported arch: $ARCH" >&2; exit 1 ;;
  esac
  echo "Installing AWS CLI v2 ($PKG)…"
  TMP="$(mktemp -d)"
  curl -fsSL "https://awscli.amazonaws.com/${PKG}" -o "$TMP/awscliv2.zip"
  unzip -q "$TMP/awscliv2.zip" -d "$TMP"
  "$TMP/aws/install" --update -i "$HOME/.local/aws-cli" -b "$HOME/.local/bin"
  if command -v sudo >/dev/null 2>&1; then
    sudo ln -sf "$HOME/.local/bin/aws" /usr/local/bin/aws 2>/dev/null || true
    sudo ln -sf "$HOME/.local/bin/aws_completer" /usr/local/bin/aws_completer 2>/dev/null || true
  fi
  rm -rf "$TMP"
  echo "aws: installed"
fi

if command -v session-manager-plugin >/dev/null 2>&1; then
  echo "session-manager-plugin: already installed"
else
  ARCH="$(uname -m)"
  case "$ARCH" in
    aarch64|arm64) DEB="session-manager-plugin.deb" URL="https://s3.amazonaws.com/session-manager-downloads/plugin/latest/ubuntu_arm64/session-manager-plugin.deb" ;;
    x86_64|amd64)  DEB="session-manager-plugin.deb" URL="https://s3.amazonaws.com/session-manager-downloads/plugin/latest/ubuntu_64bit/session-manager-plugin.deb" ;;
    *) echo "unsupported arch: $ARCH" >&2; exit 1 ;;
  esac
  echo "Installing session-manager-plugin…"
  TMP="$(mktemp -d)"
  curl -fsSL "$URL" -o "$TMP/$DEB"
  sudo dpkg -i "$TMP/$DEB" || sudo apt-get install -f -y
  rm -rf "$TMP"
  echo "session-manager-plugin: installed"
fi

# Configure ~/.ssh/config + ProxyCommand for all 9 AWS fleet instances
bash "$SCRIPT_DIR/setup-aws-ssh.sh"

# In VS Code Cloud / GitHub Copilot Coding Agent / Codespaces, persist injected secret env vars
# into ~/.aws/credentials and ~/.aws/config so all subsequent agent/terminal shells authenticate automatically.
if [ -n "${AWS_ACCESS_KEY_ID:-}" ] && [ -n "${AWS_SECRET_ACCESS_KEY:-}" ] && [ ! -f "$HOME/.aws/credentials" ]; then
  mkdir -p "$HOME/.aws"
  chmod 700 "$HOME/.aws"
  DEF_REG="${AWS_DEFAULT_REGION:-${AWS_REGION:-us-east-1}}"
  umask 077
  cat > "$HOME/.aws/credentials" <<EOF
[default]
aws_access_key_id = ${AWS_ACCESS_KEY_ID}
aws_secret_access_key = ${AWS_SECRET_ACCESS_KEY}

[famp-aws]
aws_access_key_id = ${AWS_ACCESS_KEY_ID}
aws_secret_access_key = ${AWS_SECRET_ACCESS_KEY}

[paperroute]
aws_access_key_id = ${AWS_ACCESS_KEY_ID}
aws_secret_access_key = ${AWS_SECRET_ACCESS_KEY}

[cursor-cloud-agent]
aws_access_key_id = ${AWS_ACCESS_KEY_ID}
aws_secret_access_key = ${AWS_SECRET_ACCESS_KEY}
EOF
  cat > "$HOME/.aws/config" <<EOF
[default]
region = ${DEF_REG}
output = json

[profile famp-aws]
region = ${DEF_REG}
output = json

[profile paperroute]
region = ${DEF_REG}
output = json

[profile cursor-cloud-agent]
region = ${DEF_REG}
output = json
EOF
  chmod 600 "$HOME/.aws/credentials" "$HOME/.aws/config"
  echo "aws: wrote ~/.aws/credentials and ~/.aws/config from cloud environment secrets"
fi

if [ -n "${GITHUB_PATH:-}" ]; then
  echo "$HOME/.local/bin" >> "$GITHUB_PATH"
fi
if [ -n "${GITHUB_ENV:-}" ]; then
  {
    echo "AWS_PAGER="
    echo "AWS_DEFAULT_REGION=${AWS_DEFAULT_REGION:-${AWS_REGION:-us-east-1}}"
    echo "AWS_REGION=${AWS_REGION:-${AWS_DEFAULT_REGION:-us-east-1}}"
    echo "COPILOT_AGENT_FIREWALL_ALLOW_LIST_ADDITIONS=amazonaws.com,21st.dev,services.leadconnectorhq.com,api.github.com,mcp.notion.com,mcp.clickup.com,pacelabintervals.com,app.base44.com"
    [ -n "${CLOUD_AGENT_BOX:-}" ] && echo "CLOUD_AGENT_BOX=${CLOUD_AGENT_BOX}"
  } >> "$GITHUB_ENV"
fi
if [ "$(uname -s)" = "Linux" ]; then
  for rc in "$HOME/.bashrc" "$HOME/.profile"; do
    touch "$rc"
    grep -qF 'export AWS_PAGER=""' "$rc" 2>/dev/null || echo 'export AWS_PAGER=""' >> "$rc"
    grep -qF '$HOME/.local/bin' "$rc" 2>/dev/null || echo 'export PATH="$HOME/.local/bin:$PATH"' >> "$rc"
  done
fi

# Check AWS identity if credentials are present, then pull MCPs, API keys, .env files, agents, and skills from Vault
# shellcheck source=ops/cloud-agent/config.sh
source "$SCRIPT_DIR/config.sh"
if aws sts get-caller-identity --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 --output text >/dev/null 2>&1; then
  echo "AWS identity: OK ($(aws sts get-caller-identity --no-cli-pager --query 'Arn' --output text 2>/dev/null))"
  if [ -f "$SCRIPT_DIR/vault.sh" ]; then
    bash "$SCRIPT_DIR/vault.sh" pull || echo "WARN: vault.sh pull encountered a non-fatal warning" >&2
  fi
else
  echo "AWS identity: not yet configured (ensure AWS_ACCESS_KEY_ID + AWS_SECRET_ACCESS_KEY are set in GitHub Copilot / Codespaces secrets)"
fi
