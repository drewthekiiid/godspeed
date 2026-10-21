#!/usr/bin/env bash
# OpenSSH ProxyCommand helper that tunnels SSH/SCP/rsync over AWS SSM (AWS-StartSSHSession).
# On fresh Cursor Cloud VMs, automatically generates ~/.ssh/id_ed25519 and registers its
# public key on the target instance via SSM SendCommand before opening the SSH tunnel.
set -euo pipefail

INSTANCE_ID="${1:?usage: ssm-ssh-proxy.sh <instance-id> <region> [port]}"
REGION="${2:?usage: ssm-ssh-proxy.sh <instance-id> <region> [port]}"
PORT="${3:-22}"

export AWS_PAGER=""
export PATH="$HOME/.local/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"

if [ -z "${AWS_ACCESS_KEY_ID:-}" ] || [ -z "${AWS_SECRET_ACCESS_KEY:-}" ]; then
  if [ -z "${AWS_PROFILE:-}" ]; then
    if grep -qs "cursor-cloud-agent" "$HOME/.aws/config" "$HOME/.aws/credentials" 2>/dev/null; then
      export AWS_PROFILE="cursor-cloud-agent"
    elif grep -qs "famp-aws" "$HOME/.aws/config" "$HOME/.aws/credentials" 2>/dev/null; then
      export AWS_PROFILE="famp-aws"
    elif grep -qs "paperroute" "$HOME/.aws/config" "$HOME/.aws/credentials" 2>/dev/null; then
      export AWS_PROFILE="paperroute"
    fi
  fi
else
  unset AWS_PROFILE
fi

mkdir -p "$HOME/.ssh"
chmod 700 "$HOME/.ssh"
if [ ! -f "$HOME/.ssh/id_ed25519" ]; then
  ssh-keygen -t ed25519 -N "" -C "cursor-cloud-agent@$(hostname)" -f "$HOME/.ssh/id_ed25519" >/dev/null 2>&1
fi

PUBKEY="$(cat "$HOME/.ssh/id_ed25519.pub")"
MARKER="$HOME/.ssh/.ssm_ssh_auth_${INSTANCE_ID}"

# Laptop key AAAAC3NzaC1lZDI1NTE5AAAAIK9YKycw/O5yLsx2IkaMt2O2UmH3ZyUaOTfUnVfAmWd1 is pre-authorized across the fleet.
if [ ! -f "$MARKER" ] && [[ "$PUBKEY" != *"AAAAC3NzaC1lZDI1NTE5AAAAIK9YKycw/O5yLsx2IkaMt2O2UmH3ZyUaOTfUnVfAmWd1"* ]]; then
  B64_KEY="$(printf '%s' "$PUBKEY" | base64 | tr -d '\n')"
  REMOTE_SCRIPT="K=\$(echo $B64_KEY | base64 -d); for D in /home/ubuntu/.ssh /root/.ssh; do if [ -d \"\$(dirname \"\$D\")\" ]; then mkdir -p \"\$D\"; chmod 700 \"\$D\"; touch \"\$D/authorized_keys\"; grep -qF \"\$K\" \"\$D/authorized_keys\" || printf '%s\n' \"\$K\" >> \"\$D/authorized_keys\"; chmod 600 \"\$D/authorized_keys\"; fi; done; chown -R ubuntu:ubuntu /home/ubuntu/.ssh 2>/dev/null || true"
  CMD_ID="$(aws ssm send-command \
    --instance-ids "$INSTANCE_ID" \
    --document-name AWS-RunShellScript \
    --comment "cloud-agent-ssh-key-bootstrap" \
    --parameters "commands=[\"$REMOTE_SCRIPT\"]" \
    --region "$REGION" \
    --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
    --query 'Command.CommandId' \
    --output text 2>/dev/null || true)"
  if [ -n "$CMD_ID" ]; then
    for _ in $(seq 1 15); do
      ST="$(aws ssm get-command-invocation \
        --command-id "$CMD_ID" \
        --instance-id "$INSTANCE_ID" \
        --region "$REGION" \
        --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
        --query 'Status' \
        --output text 2>/dev/null || echo Pending)"
      case "$ST" in
        Success)
          touch "$MARKER"
          break
          ;;
        Failed|Cancelled|TimedOut)
          break
          ;;
      esac
      sleep 1
    done
  fi
else
  touch "$MARKER" 2>/dev/null || true
fi

exec aws ssm start-session \
  --target "$INSTANCE_ID" \
  --document-name AWS-StartSSHSession \
  --parameters "portNumber=${PORT}" \
  --region "$REGION"
