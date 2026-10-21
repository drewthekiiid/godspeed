#!/usr/bin/env bash
# Unified AWS SSM + SSH-over-SSM control script for Cursor Cloud Agents and local sessions.
#
# Examples:
#   bash ops/cloud-agent/box.sh status
#   bash ops/cloud-agent/box.sh fleet
#   bash ops/cloud-agent/box.sh exec -- 'uptime && df -h'
#   bash ops/cloud-agent/box.sh --box bpumper exec -- 'sudo systemctl status bpumper-shadow'
#   bash ops/cloud-agent/box.sh --box polygamy exec -- 'sudo systemctl status polygamy-live'
#   bash ops/cloud-agent/box.sh --box famp-research exec -- 'ls -la /opt/madoffs-fantasy'
#   bash ops/cloud-agent/box.sh ssh 'hostname -f'
#   bash ops/cloud-agent/box.sh --box paperroute-box start
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=ops/cloud-agent/config.sh
source "$SCRIPT_DIR/config.sh"

BOX_OVERRIDE=""
while [ $# -gt 0 ]; do
  case "$1" in
    -b|--box)
      BOX_OVERRIDE="${2:?--box requires an alias or instance-id}"
      shift 2
      ;;
    *)
      break
      ;;
  esac
done

if [ -n "$BOX_OVERRIDE" ]; then
  cloud_agent_resolve_box "$BOX_OVERRIDE"
fi

INSTANCE="${TARGET_INSTANCE_ID}"
REGION="${TARGET_REGION}"
REPO="${TARGET_REPO_DIR}"
BOX_NAME="${TARGET_BOX_NAME}"
SSH_HOST="${TARGET_SSH_HOST}"

usage() {
  cat <<EOF
Usage:
  box.sh [--box <alias>] status                  Check AWS STS identity + box EC2/SSM reachability
  box.sh fleet                                   Show status of all 9 AWS fleet instances across regions
  box.sh [--box <alias>] exec -- '<command>'     Run a shell command on the box (ubuntu user via SSM)
  box.sh [--box <alias>] ssh [command...]        Connect or run command via SSH-over-SSM tunnel
  box.sh [--box <alias>] start                   Start a stopped EC2 box and wait for SSM Online
  box.sh [--box <alias>] stop                    Stop an on-demand EC2 box
  box.sh [--box <alias>] sync-repo               git pull --ff-only in $REPO on the box
  box.sh [--box <alias>] fetch-secrets           Pull app secrets on the box + restore cloud vault .env
  box.sh [--box <alias>] transcripts             List mirrored chat transcripts on the box
  box.sh vault [pull|list|get|env|status]        Pull MCP configs, 162 API keys, .env files, agents & skills from AWS Vault

Default box for this repo: $BOX_NAME ($INSTANCE in $REGION, ssh alias: $SSH_HOST)
Fleet aliases:
  polygamy (us-east-1) | famp-sniper (eu-west-1) | bpumper (eu-central-1)
  famp-research (us-east-1) | famp-hl-live (ap-northeast-1) | famp-hl-pilot (ap-northeast-1)
  famp-drift (eu-central-1) | famp-alpaca (us-east-1) | paperroute-box (us-east-1)
EOF
}

require_aws() {
  if ! command -v aws >/dev/null 2>&1; then
    echo "aws CLI missing — installing via ops/cloud-agent/install-aws-tools.sh..." >&2
    bash "$SCRIPT_DIR/install-aws-tools.sh"
  fi
}

run_ssm() {
  local remote_cmd="$1"
  local comment="${2:-cloud-agent}"
  local b64
  b64="$(printf '%s' "$remote_cmd" | base64 | tr -d '\n')"
  local cmd_id
  cmd_id="$(aws ssm send-command \
    --instance-ids "$INSTANCE" \
    --document-name AWS-RunShellScript \
    --comment "$comment" \
    --parameters "commands=[\"echo $b64 | base64 -d | sudo -iu ubuntu bash -s\"]" \
    --region "$REGION" \
    --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
    --query 'Command.CommandId' \
    --output text)"
  local status="Pending"
  for _ in $(seq 1 90); do
    status="$(aws ssm get-command-invocation \
      --command-id "$cmd_id" \
      --instance-id "$INSTANCE" \
      --region "$REGION" \
      --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
      --query 'Status' \
      --output text 2>/dev/null || echo Pending)"
    case "$status" in
      Success|Failed|Cancelled|TimedOut) break ;;
      InProgress|Pending) ;;
      *) ;;
    esac
    sleep 2
  done
  echo "box=$BOX_NAME ($INSTANCE) status=$status command_id=$cmd_id"
  aws ssm get-command-invocation \
    --command-id "$cmd_id" \
    --instance-id "$INSTANCE" \
    --region "$REGION" \
    --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
    --query 'StandardOutputContent' \
    --output text 2>/dev/null || true
  local stderr
  stderr="$(aws ssm get-command-invocation \
    --command-id "$cmd_id" \
    --instance-id "$INSTANCE" \
    --region "$REGION" \
    --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
    --query 'StandardErrorContent' \
    --output text 2>/dev/null || true)"
  if [ -n "$stderr" ]; then
    echo "--- stderr ---" >&2
    printf '%s\n' "$stderr" >&2
  fi
  case "$status" in
    Success) return 0 ;;
    *) return 1 ;;
  esac
}

CMD="${1:-help}"
shift || true

case "$CMD" in
  help|-h|--help)
    usage
    ;;
  exec)
    require_aws
    if [ "${1:-}" = "--" ]; then shift; fi
    if [ $# -lt 1 ]; then
      echo "box.sh exec requires a command string" >&2
      exit 1
    fi
    run_ssm "$*"
    ;;
  ssh)
    require_aws
    if [ ! -x "$HOME/.local/bin/cursor-ssm-ssh-proxy" ]; then
      bash "$SCRIPT_DIR/setup-aws-ssh.sh" >/dev/null
    fi
    if [ "${1:-}" = "--" ]; then shift; fi
    exec ssh "$SSH_HOST" "$@"
    ;;
  start)
    require_aws
    echo "Starting $BOX_NAME ($INSTANCE in $REGION)..."
    aws ec2 start-instances --instance-ids "$INSTANCE" --region "$REGION" --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 >/dev/null
    aws ec2 wait instance-running --instance-ids "$INSTANCE" --region "$REGION" --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30
    echo "EC2 running — waiting for SSM Online..."
    for _ in $(seq 1 30); do
      ping_st="$(aws ssm describe-instance-information \
        --filters "Key=InstanceIds,Values=$INSTANCE" \
        --region "$REGION" \
        --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
        --query 'InstanceInformationList[0].PingStatus' \
        --output text 2>/dev/null || echo None)"
      if [ "$ping_st" = "Online" ]; then
        echo "ssm: Online ($BOX_NAME ready)"
        exit 0
      fi
      sleep 2
    done
    echo "WARN: EC2 is running but SSM not yet Online; retry in a few seconds" >&2
    ;;
  stop)
    require_aws
    echo "Stopping $BOX_NAME ($INSTANCE in $REGION)..."
    aws ec2 stop-instances --instance-ids "$INSTANCE" --region "$REGION" --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 >/dev/null
    echo "Stop initiated for $BOX_NAME ($INSTANCE)"
    ;;
  sync-repo)
    require_aws
    run_ssm "cd '$REPO' && GIT_SSH_COMMAND='ssh -o BatchMode=yes -o StrictHostKeyChecking=accept-new' git pull --ff-only"
    ;;
  fetch-secrets)
    require_aws
    bash "$SCRIPT_DIR/vault.sh" pull --force
    ;;
  vault)
    require_aws
    exec bash "$SCRIPT_DIR/vault.sh" "$@"
    ;;
  transcripts)
    require_aws
    run_ssm "ls -la '${PAPERROUTE_BOX_TRANSCRIPTS_DIR}' 2>/dev/null || echo '(no transcripts mirrored yet)'"
    ;;
  fleet)
    require_aws
    printf "%-20s %-20s %-15s %-10s %-10s %-16s\n" "ALIAS" "INSTANCE_ID" "REGION" "EC2_STATE" "SSM" "PUBLIC_IP"
    printf "%-20s %-20s %-15s %-10s %-10s %-16s\n" "-----" "-----------" "------" "---------" "---" "---------"
    for alias in polygamy famp-sniper bpumper famp-research famp-hl-live famp-hl-pilot famp-drift famp-alpaca paperroute-box; do
      cloud_agent_resolve_box "$alias"
      ec2_info="$(aws ec2 describe-instances \
        --instance-ids "$TARGET_INSTANCE_ID" \
        --region "$TARGET_REGION" \
        --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
        --query 'Reservations[0].Instances[0].[State.Name,PublicIpAddress]' \
        --output text 2>/dev/null || echo "unknown -")"
      ec2_state="$(printf '%s' "$ec2_info" | awk '{print $1}')"
      pub_ip="$(printf '%s' "$ec2_info" | awk '{print $2}')"
      ssm_state="-"
      if [ "$ec2_state" = "running" ]; then
        ssm_state="$(aws ssm describe-instance-information \
          --filters "Key=InstanceIds,Values=$TARGET_INSTANCE_ID" \
          --region "$TARGET_REGION" \
          --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
          --query 'InstanceInformationList[0].PingStatus' \
          --output text 2>/dev/null || echo "-")"
        [ "$ssm_state" = "None" ] && ssm_state="-"
      fi
      printf "%-20s %-20s %-15s %-10s %-10s %-16s\n" "$alias" "$TARGET_INSTANCE_ID" "$TARGET_REGION" "$ec2_state" "$ssm_state" "${pub_ip:--}"
    done
    ;;
  status)
    require_aws
    echo "box=$BOX_NAME instance=$INSTANCE region=$REGION ssh_host=$SSH_HOST profile=${AWS_PROFILE:-<env-keys>}"
    aws sts get-caller-identity --region "$REGION" --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 2>/dev/null || {
      echo "sts: FAILED (configure AWS_ACCESS_KEY_ID + AWS_SECRET_ACCESS_KEY or CURSOR_AWS_ASSUME_IAM_ROLE_ARN in Cursor Cloud Agents secrets)"
      exit 1
    }
    ec2_state="$(aws ec2 describe-instances \
      --instance-ids "$INSTANCE" \
      --region "$REGION" \
      --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
      --query 'Reservations[0].Instances[0].State.Name' \
      --output text 2>/dev/null || echo unknown)"
    echo "ec2: $ec2_state"
    if [ "$ec2_state" != "running" ]; then
      echo "ssm: Offline (instance is $ec2_state — run 'bash ops/cloud-agent/box.sh --box $SSH_HOST start' to boot it)"
      exit 0
    fi
    probe_id="$(aws ssm send-command \
      --instance-ids "$INSTANCE" \
      --document-name AWS-RunShellScript \
      --comment "box-status-probe" \
      --parameters 'commands=["echo ok"]' \
      --region "$REGION" \
      --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
      --query 'Command.CommandId' \
      --output text 2>/dev/null || true)"
    if [ -n "$probe_id" ]; then
      probe_status="Pending"
      for _ in $(seq 1 15); do
        probe_status="$(aws ssm get-command-invocation \
          --command-id "$probe_id" \
          --instance-id "$INSTANCE" \
          --region "$REGION" \
          --no-cli-pager --cli-connect-timeout 10 --cli-read-timeout 30 \
          --query 'Status' \
          --output text 2>/dev/null || echo Pending)"
        case "$probe_status" in
          Success) echo "ssm: Online"; break ;;
          Failed|Cancelled|TimedOut) echo "ssm: $probe_status"; break ;;
        esac
        sleep 1
      done
    else
      echo "ssm: unreachable (SendCommand failed — check IAM)"
    fi
    ;;
  *)
    echo "unknown command: $CMD" >&2
    usage >&2
    exit 1
    ;;
esac
