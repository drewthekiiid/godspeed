#!/usr/bin/env bash
# Configure ~/.ssh/config and ~/.local/bin/cursor-ssm-ssh-proxy so that ssh, scp, and rsync
# transparently tunnel over AWS SSM for every EC2 box in the fleet (by alias, instance ID, or IP).
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

mkdir -p "$HOME/.local/bin" "$HOME/.ssh"
chmod 700 "$HOME/.ssh"

if [ ! -f "$HOME/.ssh/id_ed25519" ]; then
  ssh-keygen -t ed25519 -N "" -C "cursor-cloud-agent@$(hostname)" -f "$HOME/.ssh/id_ed25519" >/dev/null 2>&1
fi

cp "$SCRIPT_DIR/ssm-ssh-proxy.sh" "$HOME/.local/bin/aws-ssm-ssh-proxy"
cp "$SCRIPT_DIR/ssm-ssh-proxy.sh" "$HOME/.local/bin/cursor-ssm-ssh-proxy"
chmod 755 "$HOME/.local/bin/aws-ssm-ssh-proxy" "$HOME/.local/bin/cursor-ssm-ssh-proxy" "$SCRIPT_DIR/ssm-ssh-proxy.sh"

SSH_CONFIG="$HOME/.ssh/config"
touch "$SSH_CONFIG"
chmod 600 "$SSH_CONFIG"

BEGIN_MARK="# BEGIN CURSOR-CLOUD-AGENT-AWS-FLEET"
END_MARK="# END CURSOR-CLOUD-AGENT-AWS-FLEET"

FLEET_BLOCK="$(cat <<'EOF'
# BEGIN CURSOR-CLOUD-AGENT-AWS-FLEET
Host polygamy polygamy-box polygamy-us-east-1 i-01c97deadad199159 54.211.249.196 3.94.152.251
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-01c97deadad199159 us-east-1 %p'

Host famp-sniper poly-sniper famp-poly-sniper i-002255402e5ed0044 54.220.211.64
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-002255402e5ed0044 eu-west-1 %p'

Host bpumper bpumper-box bpumper-eu-central-1 i-001af733ab7600594 3.70.171.85
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-001af733ab7600594 eu-central-1 %p'

Host famp-research research-aws-use1 famp-research-us-east-1 i-07b4eecaa41cb53a8 13.220.83.95 54.91.6.66
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-07b4eecaa41cb53a8 us-east-1 %p'

Host famp-hl-live hl-aws-tokyo famp-hl-live-ap-northeast-1 i-0de47d4dca7713d99 52.194.211.205
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-0de47d4dca7713d99 ap-northeast-1 %p'

Host famp-hl-pilot famp-testmp testmp-aws-tokyo dydx-aws-tokyo famp-hl-pilot-ap-northeast-1 i-041be9b38ad6e05e7 18.183.225.54
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-041be9b38ad6e05e7 ap-northeast-1 %p'

Host famp-drift drift-aws-fra famp-drift-eu-central-1 i-0cefa97eec12f66bf 51.102.171.215
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-0cefa97eec12f66bf eu-central-1 %p'

Host famp-alpaca alpaca-aws-use1 famp-alpaca-us-east-1 i-007a1960d1581eef0 100.27.209.44
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-007a1960d1581eef0 us-east-1 %p'

Host paperroute paperroute-box paperroute-devbox i-08074a84c7478855e
  User ubuntu
  IdentityFile ~/.ssh/id_ed25519
  IdentitiesOnly yes
  StrictHostKeyChecking accept-new
  ConnectTimeout 15
  ServerAliveInterval 15
  ServerAliveCountMax 4
  ProxyCommand /bin/bash -c '"$HOME/.local/bin/aws-ssm-ssh-proxy" i-08074a84c7478855e us-east-1 %p'
# END CURSOR-CLOUD-AGENT-AWS-FLEET
EOF
)"

python3 - "$SSH_CONFIG" "$BEGIN_MARK" "$END_MARK" "$FLEET_BLOCK" <<'PY'
import sys, re

cfg_path, begin_mark, end_mark, block = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
with open(cfg_path, "r", encoding="utf-8") as f:
    content = f.read()

pattern = re.compile(re.escape(begin_mark) + r".*?" + re.escape(end_mark) + r"\n?", re.DOTALL)
if pattern.search(content):
    new_content = pattern.sub(block + "\n", content)
else:
    sep = "" if (not content or content.endswith("\n")) else "\n"
    new_content = content + sep + "\n" + block + "\n"

with open(cfg_path, "w", encoding="utf-8") as f:
    f.write(new_content.lstrip("\n"))
PY

chmod 600 "$SSH_CONFIG"
echo "ssh-over-ssm: configured 9 AWS fleet hosts in $SSH_CONFIG"
