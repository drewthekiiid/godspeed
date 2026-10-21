#!/usr/bin/env bash
# Non-secret AWS fleet & default box configuration for Cursor Cloud Agents.
# Works with both IAM Role Assume (CURSOR_AWS_ASSUME_IAM_ROLE_ARN -> AWS_PROFILE=cursor-cloud-agent)
# and Runtime Secret access keys (AWS_ACCESS_KEY_ID + AWS_SECRET_ACCESS_KEY), or local laptop profiles.

export AWS_PAGER=""
export PATH="$HOME/.local/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"

cloud_agent_aws_env() {
  export AWS_PAGER=""
  if [ -n "${AWS_ACCESS_KEY_ID:-}" ] && [ -n "${AWS_SECRET_ACCESS_KEY:-}" ]; then
    unset AWS_PROFILE
    return 0
  fi
  if [ -n "${AWS_PROFILE:-}" ]; then
    return 0
  fi
  if grep -qs "cursor-cloud-agent" "$HOME/.aws/config" "$HOME/.aws/credentials" 2>/dev/null; then
    export AWS_PROFILE="cursor-cloud-agent"
  elif grep -qs "famp-aws" "$HOME/.aws/config" "$HOME/.aws/credentials" 2>/dev/null; then
    export AWS_PROFILE="famp-aws"
  elif grep -qs "paperroute" "$HOME/.aws/config" "$HOME/.aws/credentials" 2>/dev/null; then
    export AWS_PROFILE="paperroute"
  fi
}

cloud_agent_aws_env

# Repo-specific primary defaults (overridden per repo or via BOX / CLOUD_AGENT_BOX env)
CLOUD_AGENT_DEFAULT_BOX="${CLOUD_AGENT_DEFAULT_BOX:-famp-research}"
CLOUD_AGENT_DEFAULT_INSTANCE_ID="${CLOUD_AGENT_DEFAULT_INSTANCE_ID:-i-07b4eecaa41cb53a8}"
CLOUD_AGENT_DEFAULT_REGION="${CLOUD_AGENT_DEFAULT_REGION:-us-east-1}"
CLOUD_AGENT_DEFAULT_REPO_DIR="${CLOUD_AGENT_DEFAULT_REPO_DIR:-/opt/madoffs-fantasy}"

# Fleet-wide instance IDs & regions
export POLYGAMY_BOX_INSTANCE_ID="${POLYGAMY_BOX_INSTANCE_ID:-i-01c97deadad199159}"
export POLYGAMY_AWS_REGION="${POLYGAMY_AWS_REGION:-us-east-1}"

export FAMP_SNIPER_AWS_INSTANCE_ID="${FAMP_SNIPER_AWS_INSTANCE_ID:-i-002255402e5ed0044}"
export FAMP_SNIPER_AWS_REGION="${FAMP_SNIPER_AWS_REGION:-eu-west-1}"

export BPUMPER_BOX_INSTANCE_ID="${BPUMPER_BOX_INSTANCE_ID:-i-001af733ab7600594}"
export BPUMPER_AWS_REGION="${BPUMPER_AWS_REGION:-eu-central-1}"

export FAMP_RESEARCH_AWS_INSTANCE_ID="${FAMP_RESEARCH_AWS_INSTANCE_ID:-i-07b4eecaa41cb53a8}"
export FAMP_RESEARCH_AWS_REGION="${FAMP_RESEARCH_AWS_REGION:-us-east-1}"
export FAMP_RESEARCH_AWS="${FAMP_RESEARCH_AWS:-ubuntu@famp-research}"

export FAMP_HL_AWS_INSTANCE_ID="${FAMP_HL_AWS_INSTANCE_ID:-i-0de47d4dca7713d99}"
export FAMP_HL_AWS_REGION="${FAMP_HL_AWS_REGION:-ap-northeast-1}"
export FAMP_HL_AWS="${FAMP_HL_AWS:-ubuntu@famp-hl-live}"

export FAMP_TESTMP_AWS_INSTANCE_ID="${FAMP_TESTMP_AWS_INSTANCE_ID:-i-041be9b38ad6e05e7}"
export FAMP_TESTMP_AWS_REGION="${FAMP_TESTMP_AWS_REGION:-ap-northeast-1}"
export FAMP_TESTMP_AWS="${FAMP_TESTMP_AWS:-ubuntu@famp-hl-pilot}"
export FAMP_DYDX_AWS="${FAMP_DYDX_AWS:-ubuntu@famp-hl-pilot}"
export FAMP_DYDX_AWS_REGION="${FAMP_DYDX_AWS_REGION:-ap-northeast-1}"

export FAMP_DRIFT_AWS_INSTANCE_ID="${FAMP_DRIFT_AWS_INSTANCE_ID:-i-0cefa97eec12f66bf}"
export FAMP_DRIFT_AWS_REGION="${FAMP_DRIFT_AWS_REGION:-eu-central-1}"
export FAMP_DRIFT_AWS="${FAMP_DRIFT_AWS:-ubuntu@famp-drift}"

export FAMP_ALPACA_AWS_INSTANCE_ID="${FAMP_ALPACA_AWS_INSTANCE_ID:-i-007a1960d1581eef0}"
export FAMP_ALPACA_AWS_REGION="${FAMP_ALPACA_AWS_REGION:-us-east-1}"
export FAMP_ALPACA_AWS="${FAMP_ALPACA_AWS:-ubuntu@famp-alpaca}"

export PAPERROUTE_BOX_INSTANCE_ID="${PAPERROUTE_BOX_INSTANCE_ID:-i-08074a84c7478855e}"
export PAPERROUTE_AWS_REGION="${PAPERROUTE_AWS_REGION:-us-east-1}"
export PAPERROUTE_APP_SECRET_ID="${PAPERROUTE_APP_SECRET_ID:-paperroute/app}"
export PAPERROUTE_BOX_REPO_DIR="${PAPERROUTE_BOX_REPO_DIR:-/home/ubuntu/paperroute}"
export PAPERROUTE_BOX_TRANSCRIPTS_DIR="${PAPERROUTE_BOX_TRANSCRIPTS_DIR:-${PAPERROUTE_BOX_REPO_DIR}/.cursor/agent-transcripts}"

cloud_agent_resolve_box() {
  local key="${1:-${BOX:-${CLOUD_AGENT_BOX:-$CLOUD_AGENT_DEFAULT_BOX}}}"
  case "$key" in
    polygamy|polygamy-box|polygamy-us-east-1|i-01c97deadad199159|54.211.249.196|3.94.152.251)
      TARGET_BOX_NAME="polygamy-us-east-1"
      TARGET_INSTANCE_ID="$POLYGAMY_BOX_INSTANCE_ID"
      TARGET_REGION="$POLYGAMY_AWS_REGION"
      TARGET_REPO_DIR="/opt/polygamy"
      TARGET_SSH_HOST="polygamy"
      ;;
    poly-sniper|famp-sniper|famp-poly-sniper|sniper|dublin|i-002255402e5ed0044|54.220.211.64)
      TARGET_BOX_NAME="famp-poly-sniper"
      TARGET_INSTANCE_ID="$FAMP_SNIPER_AWS_INSTANCE_ID"
      TARGET_REGION="$FAMP_SNIPER_AWS_REGION"
      TARGET_REPO_DIR="/opt/polygamy"
      TARGET_SSH_HOST="famp-sniper"
      ;;
    bpumper|bpumper-box|bpumper-eu-central-1|frankfurt|i-001af733ab7600594|3.70.171.85)
      TARGET_BOX_NAME="bpumper-eu-central-1"
      TARGET_INSTANCE_ID="$BPUMPER_BOX_INSTANCE_ID"
      TARGET_REGION="$BPUMPER_AWS_REGION"
      TARGET_REPO_DIR="/opt/bpumper"
      TARGET_SSH_HOST="bpumper"
      ;;
    famp-research|research|research-aws-use1|famp-research-us-east-1|i-07b4eecaa41cb53a8|13.220.83.95|54.91.6.66)
      TARGET_BOX_NAME="famp-research-us-east-1"
      TARGET_INSTANCE_ID="$FAMP_RESEARCH_AWS_INSTANCE_ID"
      TARGET_REGION="$FAMP_RESEARCH_AWS_REGION"
      TARGET_REPO_DIR="/opt/madoffs-fantasy"
      TARGET_SSH_HOST="famp-research"
      ;;
    famp-hl-live|hl-live|hl-aws-tokyo|famp-hl-live-ap-northeast-1|i-0de47d4dca7713d99|52.194.211.205)
      TARGET_BOX_NAME="famp-hl-live-ap-northeast-1"
      TARGET_INSTANCE_ID="$FAMP_HL_AWS_INSTANCE_ID"
      TARGET_REGION="$FAMP_HL_AWS_REGION"
      TARGET_REPO_DIR="/opt/madoffs-fantasy"
      TARGET_SSH_HOST="famp-hl-live"
      ;;
    famp-hl-pilot|hl-pilot|famp-testmp|testmp|testmp-aws-tokyo|dydx-aws-tokyo|famp-hl-pilot-ap-northeast-1|i-041be9b38ad6e05e7|18.183.225.54)
      TARGET_BOX_NAME="famp-hl-pilot-ap-northeast-1"
      TARGET_INSTANCE_ID="$FAMP_TESTMP_AWS_INSTANCE_ID"
      TARGET_REGION="$FAMP_TESTMP_AWS_REGION"
      TARGET_REPO_DIR="/opt/madoffs-fantasy"
      TARGET_SSH_HOST="famp-hl-pilot"
      ;;
    famp-drift|drift|drift-aws-fra|famp-drift-eu-central-1|i-0cefa97eec12f66bf|51.102.171.215)
      TARGET_BOX_NAME="famp-drift-eu-central-1"
      TARGET_INSTANCE_ID="$FAMP_DRIFT_AWS_INSTANCE_ID"
      TARGET_REGION="$FAMP_DRIFT_AWS_REGION"
      TARGET_REPO_DIR="/opt/changenstein-research"
      TARGET_SSH_HOST="famp-drift"
      ;;
    famp-alpaca|alpaca|alpaca-aws-use1|famp-alpaca-us-east-1|i-007a1960d1581eef0|100.27.209.44)
      TARGET_BOX_NAME="famp-alpaca-us-east-1"
      TARGET_INSTANCE_ID="$FAMP_ALPACA_AWS_INSTANCE_ID"
      TARGET_REGION="$FAMP_ALPACA_AWS_REGION"
      TARGET_REPO_DIR="/opt/changenstein-research"
      TARGET_SSH_HOST="famp-alpaca"
      ;;
    paperroute|paperroute-box|paperroute-devbox|i-08074a84c7478855e)
      TARGET_BOX_NAME="paperroute-devbox"
      TARGET_INSTANCE_ID="$PAPERROUTE_BOX_INSTANCE_ID"
      TARGET_REGION="$PAPERROUTE_AWS_REGION"
      TARGET_REPO_DIR="$PAPERROUTE_BOX_REPO_DIR"
      TARGET_SSH_HOST="paperroute-box"
      ;;
    *)
      TARGET_BOX_NAME="$key"
      TARGET_INSTANCE_ID="$CLOUD_AGENT_DEFAULT_INSTANCE_ID"
      TARGET_REGION="$CLOUD_AGENT_DEFAULT_REGION"
      TARGET_REPO_DIR="$CLOUD_AGENT_DEFAULT_REPO_DIR"
      TARGET_SSH_HOST="$CLOUD_AGENT_DEFAULT_BOX"
      ;;
  esac
  export TARGET_BOX_NAME TARGET_INSTANCE_ID TARGET_REGION TARGET_REPO_DIR TARGET_SSH_HOST
}

cloud_agent_resolve_box "${BOX:-${CLOUD_AGENT_BOX:-$CLOUD_AGENT_DEFAULT_BOX}}"
