#!/usr/bin/env bash
# Quick boot preflight for Cursor Cloud Agent VMs — refreshes SSH-over-SSM config and verifies AWS + box reachability.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
bash "$SCRIPT_DIR/setup-aws-ssh.sh" >/dev/null 2>&1 || true
bash "$SCRIPT_DIR/box.sh" status || {
  echo "WARN: AWS box pipeline not ready — set AWS_ACCESS_KEY_ID + AWS_SECRET_ACCESS_KEY or CURSOR_AWS_ASSUME_IAM_ROLE_ARN in Cursor Cloud Agents secrets" >&2
  exit 0
}
