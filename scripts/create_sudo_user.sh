#!/usr/bin/env bash

set -euo pipefail

BOLD=$'\e[1m'
DIM=$'\e[2m'
ITALIC=$'\e[3m'      # Not supported in all terminals; harmless if ignored
UNDER=$'\e[4m'
RED=$'\e[31m'
GREEN=$'\e[32m'
YELLOW=$'\e[33m'
BLUE=$'\e[34m'
MAGENTA=$'\e[35m'
CYAN=$'\e[36m'
RESET=$'\e[0m'

msg() { echo -e "${CYAN}[*]${RESET} $*"; }
ok()  { echo -e "${GREEN}[+]${RESET} $*"; }
warn(){ echo -e "${YELLOW}[~]${RESET} $*"; }
err() { echo -e "${RED}[!]${RESET} $*" >&2; }

banner() {
  echo -e "${BOLD}${BLUE}────────────────────────────────────────────────────────────${RESET}"
  echo -e "${BOLD}${BLUE}  Sudo User Creator • SSH Key Migration • Root Hardening     ${RESET}"
  echo -e "${BOLD}${BLUE}────────────────────────────────────────────────────────────${RESET}"
}

require_root() {
  if [[ "${EUID:-$(id -u)}" -ne 0 ]]; then
    err "This script must be run as root. Abort."
    exit 1
  fi
}

prompt_yes_no() {
  local prompt="$1"
  local default="${2:-N}" # default N if not provided
  local hint="[y/N]"
  [[ "${default^^}" == "Y" ]] && hint="[Y/n]"

  local ans
  read -r -p "$(echo -e "${BOLD}${prompt}${RESET} ${DIM}${hint}${RESET} ")" ans || true
  ans="${ans,,}"

  if [[ -z "$ans" ]]; then
    [[ "${default^^}" == "Y" ]] && return 0 || return 1
  fi
  [[ "$ans" == "y" || "$ans" == "yes" ]]
}

read_with_default() {
  local prompt="$1"
  local def="$2"
  local ans
  read -r -p "$(echo -e "${BOLD}${prompt}${RESET} ${DIM}[${def}]${RESET}: ")" ans || true
  if [[ -z "$ans" ]]; then
    echo -n "$def"
  else
    echo -n "$ans"
  fi
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
  ok "Backed up ${BOLD}$sshd_cfg${RESET} to ${BOLD}$backup${RESET}"

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
  ok "Root SSH login ${BOLD}disabled${RESET} and SSH service restarted."
}

main() {
  clear
  banner
  require_root

  msg "This helper can ${BOLD}create a non-root (sudo) user${RESET} and ${BOLD}migrate your current SSH keys${RESET}."
  msg "It can also ${BOLD}disable root SSH login${RESET} after the new user is set up."
  echo

  # Default decisions
  local default_proceed="Y"        # default Proceed?  -> Yes
  local default_user="nodeadmin"   # default username -> nodeadmin
  local default_disable_root="Y"   # default Disable root SSH now? -> Yes

  if ! prompt_yes_no "Proceed to create a non-root (sudo) user and migrate SSH keys?" "$default_proceed"; then
    warn "No changes made. Exiting."
    exit 0
  fi

  local new_user=""
  while true; do
    new_user="$(read_with_default "Enter the new ${ITALIC}username${RESET} (lowercase; start with a letter; may contain digits, - or _)" "$default_user")"
    new_user="${new_user,,}"
    if ! username_is_valid "$new_user"; then
      err "Invalid username. Use lowercase, start with a letter, then letters/digits/-/_ (max 32). Examples: alice, nodeadmin, dev_ops"
      continue
    fi
    if id "$new_user" >/dev/null 2>&1; then
      warn "User '${BOLD}$new_user${RESET}' already exists."
      for i in {1..9}; do
        if ! id "${new_user}${i}" >/dev/null 2>&1; then
          local suggestion="${new_user}${i}"
          new_user="$(read_with_default "Pick another username" "$suggestion")"
          new_user="${new_user,,}"
          break
        fi
      done
      continue
    fi
    break
  done

  echo
  msg "You'll now set a password for '${BOLD}$new_user${RESET}'."
  cat <<'EOF'

Password tips:
  ✓ Allowed (and recommended): Letters (A-Z, a-z), Numbers (0-9), and symbols:
      !  @  #  %  ^  +  =  _  -  .  ,
  ⚠ Avoid when possible (can cause issues in some shells/tools):
      !!   !$   `   '   "   \   |   &   ;   <   >   *   ?   [   ]   ~   $

EOF

  adduser --disabled-password --gecos "" "$new_user"
  ok "User '${BOLD}$new_user${RESET}' created."

  passwd "$new_user"

  usermod -aG sudo "$new_user"
  ok "User '${BOLD}$new_user${RESET}' added to '${BOLD}sudo${RESET}' group."

  local src_keys dest_home
  src_keys="$(ensure_authorized_keys)"
  if [[ -z "$src_keys" ]]; then
    warn "No existing ${BOLD}authorized_keys${RESET} found for root (or SUDO_USER)."
    warn "Skipping SSH key migration. You can add keys later for ${BOLD}$new_user${RESET}."
  else
    dest_home="$(eval echo "~$new_user")"
    install -d -m 700 "$dest_home/.ssh"
    install -m 600 "$src_keys" "$dest_home/.ssh/authorized_keys"
    harden_permissions "$new_user"
    ok "SSH ${BOLD}authorized_keys${RESET} copied from '${BOLD}$src_keys${RESET}' to '${BOLD}$dest_home/.ssh/authorized_keys${RESET}'."
  fi

  echo
  msg "Quick access check (recommended): open another terminal and test:"
  echo -e "  ${BOLD}ssh ${new_user}@<server_ip>${RESET}"
  echo

  if prompt_yes_no "Disable root SSH login now and restart SSH?" "$default_disable_root"; then
    disable_root_ssh_login
    ok "All done. Next time, log in as: ${BOLD}ssh ${new_user}@<server_ip>${RESET}"
  else
    warn "Left root SSH login ${BOLD}ENABLED${RESET}. You can disable it later after testing."
  fi

  echo
  ok "Script completed."
}

main "$@"
