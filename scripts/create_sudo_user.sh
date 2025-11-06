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
  if [[ -t 1 ]]; then
    clear || true
    printf '\e[3J\e[H\e[2J' || true
  fi
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
  banner
  require_root

  msg "This script can ${BOLD}create a non-root (sudo) user${RESET} and ${BOLD}migrate your current SSH keys${RESET}."
  msg "It can also ${BOLD}disable root SSH login${RESET} after the new user is set up."
  echo

  local default_proceed="Y"        # Proceed? default Yes
  local default_user="nodeadmin"   # Suggested username
  local default_disable_root="Y"   # Disable root SSH after setup? default Yes

  if ! prompt_yes_no "Proceed to create a non-root (sudo) user and migrate SSH keys?" "$default_proceed"; then
    warn "No changes made. Exiting."
    return 0
  fi

  local new_user=""
  local prompt="Enter the new ${ITALIC}username${RESET} (lowercase; start with a letter; may contain digits, - or _)"
  local def="$default_user"

  while true; do
    new_user="$(read_with_default "$prompt" "$def")"
    new_user="${new_user,,}"

    if ! username_is_valid "$new_user"; then
      err "Invalid username. Use lowercase, start with a letter, then letters/digits/-/_ (max 32). Examples: alice, nodeadmin, dev_ops"
      continue
    fi

    if id "$new_user" >/dev/null 2>&1; then
      warn "User '${BOLD}$new_user${RESET}' already exists."
      local suggestion=""
      for i in {1..9}; do
        if ! id "${new_user}${i}" >/dev/null 2>&1; then
          suggestion="${new_user}${i}"
          break
        fi
      done
      def="${suggestion:-${new_user}1}"
      prompt="Enter the new ${ITALIC}username${RESET} (taken; suggested: ${BOLD}${def}${RESET})"
      continue
    fi

    break
  done

  echo
  msg "You'll now set a password for '${BOLD}$new_user${RESET}'."
  cat <<'EOF'

Password tips:
  ✓ Allowed (and recommended): Letters (A–Z, a–z), Numbers (0–9), and symbols:
      !  @  #  %  ^  +  =  _  -  .  ,
  ⚠ Avoid when possible (can be error-prone when pasted/typed in some shells/tools):
      !!   !$   `   '   "   \   |   &   ;   <   >   *   ?   [   ]   ~   $
(These are NOT forbidden — just more likely to cause copy/paste or shell interpretation issues.)

EOF

  local pw1="" pw2=""
  while true; do
    read -rs -p "Enter password for ${new_user}: " pw1; echo
    read -rs -p "Confirm password: " pw2; echo
    if [[ -z "$pw1" ]]; then
      err "Password cannot be empty."
      continue
    fi
    if [[ "$pw1" != "$pw2" ]]; then
      err "Passwords do not match. Try again."
      continue
    fi
    break
  done

  echo
  local disable_root_after="N"
  if prompt_yes_no "Disable root SSH login after setup?" "$default_disable_root"; then
    disable_root_after="Y"
  fi

  local src_keys=""
  src_keys="$(ensure_authorized_keys)"
  if [[ -z "$src_keys" ]]; then
    warn "No existing ${BOLD}authorized_keys${RESET} found for root (or SUDO_USER). Key migration will be skipped."
    if ! prompt_yes_no "Continue without migrating SSH keys?" "Y"; then
      warn "No changes made. Exiting."
      unset pw1 pw2
      return 0
    fi
  else
    ok "Will copy keys from: ${BOLD}$src_keys${RESET}"
  fi

  echo
  echo -e "${BOLD}${BLUE}Summary of actions (no changes yet):${RESET}"
  echo -e "  New username     : ${BOLD}$new_user${RESET}"
  echo -e "  SSH key source   : ${BOLD}${src_keys:-<none>}${RESET}"
  echo -e "  Disable root SSH : ${BOLD}$([[ "$disable_root_after" == "Y" ]] && echo Yes || echo No)${RESET}"
  echo -e "  Password         : ${BOLD}********${RESET}"
  echo
  if ! prompt_yes_no "Apply these changes now?" "Y"; then
    warn "Cancelled. No changes made."
    unset pw1 pw2
    return 0
  fi

  msg "Creating user '${BOLD}$new_user${RESET}'..."
  if ! adduser --disabled-password --gecos "" "$new_user" >/dev/null; then
    err "User creation failed. Aborting."
    unset pw1 pw2
    return 1
  fi
  ok "User '${BOLD}$new_user${RESET}' created."

  if ! printf '%s:%s\n' "$new_user" "$pw1" | chpasswd; then
    err "Failed to set password for '$new_user'. Rolling back."
    deluser --remove-home "$new_user" >/dev/null 2>&1 || true
    unset pw1 pw2
    return 1
  fi
  unset pw1 pw2
  ok "Password set for '${BOLD}$new_user${RESET}'."

  if usermod -aG sudo "$new_user" >/dev/null 2>&1; then
    ok "User '${BOLD}$new_user${RESET}' added to '${BOLD}sudo${RESET}' group."
  else
    err "Failed to add '${new_user}' to sudo group."
  fi

  local dest_home=""
  if [[ -n "$src_keys" ]]; then
    dest_home="$(eval echo "~$new_user")"
    install -d -m 700 "$dest_home/.ssh"
    install -m 600 "$src_keys" "$dest_home/.ssh/authorized_keys"
    harden_permissions "$new_user"
    ok "SSH ${BOLD}authorized_keys${RESET} copied to '${BOLD}$dest_home/.ssh/authorized_keys${RESET}'."
  fi

  echo
  msg "Quick access check (recommended): open another terminal and test:"
  echo -e "  ${BOLD}ssh ${new_user}@<server_ip>${RESET}"
  echo

  if [[ "$disable_root_after" == "Y" ]]; then
    disable_root_ssh_login
    ok "All done. Next time, log in as: ${BOLD}ssh ${new_user}@<server_ip>${RESET}"
  else
    warn "Left root SSH login ${BOLD}ENABLED${RESET}. You can disable it later after testing."
  fi

  echo
  ok "Script completed."
}

main "$@"
