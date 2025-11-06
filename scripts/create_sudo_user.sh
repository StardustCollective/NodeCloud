#!/usr/bin/env bash

set -euo pipefail

msg() { echo -e "[*] $*"; }
ok()  { echo -e "[+] $*"; }
err() { echo -e "[!] $*" >&2; }

require_root() {
  if [[ "${EUID:-$(id -u)}" -ne 0 ]]; then
    err "This script must be run as root. Abort."
    exit 1
  fi
}

prompt_yes_no() {
  local prompt="${1:-Continue? [y/N]} "
  local ans
  read -r -p "$prompt" ans || true
  [[ "${ans,,}" == "y" || "${ans,,}" == "yes" ]]
}

username_is_valid() {
  [[ "$1" =~ ^[a-z][-a-z0-9_]*$ ]] && [[ "${#1}" -le 32 ]]
}

ensure_authorized_keys() {
  local src_keys=""
  if [[ -f /root/.ssh/authorized_keys ]]; then
    src_keys="/root/.ssh/authorized_keys"
  elif [[ -n "${SUDO_USER-}" && -f "/home/$SUDO_USER/.ssh/authorized_keys" ]]; then
    src_keys="/home/$SUDO_USER/.ssh/authorized_keys"
  fi
  echo -n "$src_keys"
}

harden_permissions() {
  local user="$1"
  local home_dir
  home_dir="$(eval echo "~$user")"
  chmod 700 "$home_dir/.ssh"
  chmod 600 "$home_dir/.ssh/authorized_keys"
  chown -R "$user:$user" "$home_dir/.ssh"
}

disable_root_ssh_login() {
  local sshd_cfg="/etc/ssh/sshd_config"
  local backup="/etc/ssh/sshd_config.$(date +%Y%m%d-%H%M%S).bak"
  cp -a "$sshd_cfg" "$backup"
  ok "Backed up $sshd_cfg to $backup"

  if grep -qiE '^\s*PermitRootLogin' "$sshd_cfg"; then
    sed -i 's/^\s*PermitRootLogin\s\+.*/# &/I' "$sshd_cfg"
  fi
  echo "PermitRootLogin no" >> "$sshd_cfg"

  if ls /etc/ssh/sshd_config.d/*.conf >/dev/null 2>&1; then
    sed -i 's/^\s*PermitRootLogin\s\+yes/# &/I' /etc/ssh/sshd_config.d/*.conf || true
  fi

  if ! sshd -t 2>/tmp/sshd_test.err; then
    err "sshd config test failed:"
    cat /tmp/sshd_test.err >&2
    err "Restoring previous config."
    mv -f "$backup" "$sshd_cfg"
    exit 1
  fi

  systemctl restart ssh || systemctl restart sshd
  ok "Root SSH login disabled and SSH service restarted."
}

require_root

msg "This helper can create a non-root (sudo) user and migrate your current SSH keys."
msg "It can also disable root SSH login AFTER the keys are copied."
echo

if ! prompt_yes_no "Proceed to create a non-root (sudo) user and disable root SSH login? [y/N] "; then
  msg "No changes made. Exiting."
  exit 0
fi

# Ask for username
new_user=""
while true; do
  read -r -p "Enter the new username (lowercase letters, digits, -, _ ; must start with a letter): " new_user || true
  new_user="${new_user,,}"
  if [[ -z "$new_user" ]]; then
    err "Username cannot be empty."
    continue
  fi
  if ! username_is_valid "$new_user"; then
    err "Invalid username. Example valid names: alice, nodeadmin, dev_ops"
    continue
  fi
  if id "$new_user" >/dev/null 2>&1; then
    err "User '$new_user' already exists. Choose another."
    continue
  fi
  break
done

echo
msg "You'll now set a password for '$new_user'."

cat <<'EOF'
Password tips:
+ Allowed (and recommended): Letters (A-Z, a-z), Numbers (0-9), and symbols:
    !  @  #  %  ^  +  =  _  -  .  ,
x  Avoid when possible (these can possibly cause issues in some shells/tools):
    !!   !$   `   '   "   \   |   &   ;   <   >   *   ?   [   ]   ~   $
EOF
echo

adduser --disabled-password --gecos "" "$new_user"
ok "User '$new_user' created."

passwd "$new_user"

usermod -aG sudo "$new_user"
ok "User '$new_user' added to 'sudo' group."

src_keys="$(ensure_authorized_keys)"
if [[ -z "$src_keys" ]]; then
  err "No existing authorized_keys found for root (or SUDO_USER)."
  err "Skipping SSH key migration. You can create $new_user's keys later."
else
  dest_home="$(eval echo "~$new_user")"
  install -d -m 700 "$dest_home/.ssh"
  install -m 600 "$src_keys" "$dest_home/.ssh/authorized_keys"
  harden_permissions "$new_user"
  ok "SSH authorized_keys copied from '$src_keys' to $dest_home/.ssh/authorized_keys"
fi

echo
msg "Quick access check (optional): try this in another terminal BEFORE disabling root login:"
echo "    ssh ${new_user}@<server_ip>"
echo

if prompt_yes_no "Disable root SSH login now and restart SSH? (Recommended) [y/N] "; then
  disable_root_ssh_login
  ok "All done. Next time, log in as: ssh ${new_user}@<server_ip>"
else
  msg "Left root SSH login ENABLED (you can disable it later)."
fi

ok "Script completed."
