#!/usr/bin/env bash
set -euo pipefail

BOLD=$'\e[1m'
DIM=$'\e[2m'
UNDER=$'\e[4m'
RED=$'\e[31m'
GREEN=$'\e[32m'
YELLOW=$'\e[33m'
BLUE=$'\e[34m'
MAGENTA=$'\e[35m'
CYAN=$'\e[36m'
RESET=$'\e[0m'

msg()  { echo -e "${CYAN}[*]${RESET} $*"; }
ok()   { echo -e "${GREEN}[+]${RESET} $*"; }
warn() { echo -e "${YELLOW}[~]${RESET} $*"; }
err()  { echo -e "${RED}[!]${RESET} $*" >&2; }

banner() {
  if [[ -t 1 ]]; then
    clear || true
    printf '\e[3J\e[H\e[2J' || true
  fi
  echo -e "${BOLD}${BLUE}────────────────────────────────────────────────────────────${RESET}"
  echo -e "${BOLD}${BLUE}     STARDUST COLLECTIVE • SERVER SETUP WIZARD          ${RESET}"
  echo -e "${BOLD}${BLUE}────────────────────────────────────────────────────────────${RESET}"
  echo
}

intro() {
  banner
  echo -e "${BOLD}Welcome to the Server Setup Wizard${RESET}"
  echo
  echo -e "This guided tool will help you:"
  echo -e "  ${GREEN}•${RESET} Create a secure ${BOLD}non-root sudo user${RESET} (recommended)"
  echo -e "  ${GREEN}•${RESET} ${BOLD}Back up your P12${RESET} file safely from this server"
  echo -e "  ${GREEN}•${RESET} View instructions to ${BOLD}upload your P12${RESET} from Windows/macOS"
  echo
  echo -e "You can run this wizard:"
  echo -e "  ${GREEN}•${RESET} On a ${BOLD}new server${RESET} (to set up a fresh environment)"
  echo -e "  ${GREEN}•${RESET} On an ${BOLD}old server${RESET} (to locate and back up your P12)"
  echo
  echo -e "${DIM}Nothing is changed until you pick an option and confirm.${RESET}"
  echo
  echo -e "${BOLD}Press any key to continue...${RESET}"
  IFS= read -rsn1 _
}

prompt_yes_no() {
  local prompt="$1"
  local def="${2:-Y}"
  local hint="[y/N]"
  [[ "${def^^}" == "Y" ]] && hint="[Y/n]"

  local ans
  read -r -p "$(echo -e "${BOLD}${prompt}${RESET} ${DIM}${hint}${RESET} ")" ans || true
  ans="${ans,,}"

  if [[ -z "$ans" ]]; then
    [[ "${def^^}" == "Y" ]] && return 0 || return 1
  fi
  [[ "$ans" == "y" || "$ans" == "yes" ]]
}

pause_any_key() {
  echo
  echo -e "${DIM}Press any key to return to the menu...${RESET}"
  IFS= read -rsn1 _ || true
}

MENU_ITEMS=(
  "Full Server Setup"
  "Create Non-Root Sudo User"
  "Backup P12 File on This Server"
  "Show P12 Upload Instructions (Windows / macOS)"
  "Exit"
)

draw_menu() {
  banner
  echo -e "${BOLD}Use the ${UNDER}↑ / ↓${RESET}${BOLD} arrow keys and press Enter to select:${RESET}"
  echo

  local idx=0
  for item in "${MENU_ITEMS[@]}"; do
    if [[ $idx -eq $1 ]]; then
      echo -e "${GREEN}> ${BOLD}${item}${RESET}"
    else
      echo -e "  ${item}"
    fi
    ((idx++))
  done

  echo
}

menu_loop() {
  local selected=0
  local num_items=${#MENU_ITEMS[@]}

  while true; do
    draw_menu "$selected"

    IFS= read -rsn1 key || continue

    if [[ $key == $'\x1b' ]]; then
      IFS= read -rsn2 -t 0.001 key2 || key2=""
      case "$key2" in
        "[A")
          ((selected--))
          ((selected < 0)) && selected=$((num_items-1))
          ;;
        "[B")
          ((selected++))
          ((selected >= num_items)) && selected=0
          ;;
      esac
    elif [[ $key == "" ]]; then
      case "$selected" in
        0) full_server_setup ;;
        1) create_non_root_user_only ;;
        2) backup_p12_on_this_server ;;
        3) show_p12_upload_help ;;
        4) exit 0 ;;
      esac
    fi
  done
}

run_create_sudo_script() {
  local script="create_sudo_user.sh"
  msg "Fetching latest ${BOLD}${script}${RESET} from GitHub..."
  rm -f "$script"
  if curl -fsSL -o "$script" "https://github.com/StardustCollective/NodeCloud/raw/main/scripts/create_sudo_user.sh"; then
    chmod +x "$script"
    echo
    ok "Launching ${BOLD}$script${RESET}..."
    sudo bash "$script"
  else
    err "Failed to download ${script}. Please check your network or URL."
    return 1
  fi
}

create_non_root_user_only() {
  banner
  echo -e "${BOLD}Create Non-Root Sudo User${RESET}"
  echo

  if [[ $EUID -ne 0 ]]; then
    warn "You are logged in as ${BOLD}$(id -un)${RESET}, not root."
    echo
    echo "This step is usually run when you first log in as ${BOLD}root${RESET}."
    echo "If you already created a sudo user, you can skip this."
    echo
    pause_any_key
    return 0
  fi

  if ! prompt_yes_no "Run the non-root user setup script now?" "Y"; then
    warn "Skipped user creation."
    pause_any_key
    return 0
  fi

  run_create_sudo_script || true
  echo
  echo -e "${BOLD}Next Steps:${RESET}"
  echo -e "  1) Close this SSH session."
  echo -e "  2) Log back in as your new sudo user (for example ${BOLD}nodeadmin${RESET})."
  echo -e "  3) Run this wizard again as the new user."
  echo
  pause_any_key
}

search_p12_files() {
  local search_roots=(
    /root
    /home
    /var/tessellation
    /opt
  )

  mapfile -t P12_FOUND < <(
    for base in "${search_roots[@]}"; do
      [[ -d "$base" ]] || continue
      find "$base" -maxdepth 5 \
        \( -name hash -o -name ordinal \) -prune -o \
        -type f -iname '*.p12' -print 2>/dev/null
    done | sort -u
  )
}

select_from_list() {
  local -n arr=$1
  local prompt="$2"

  if ((${#arr[@]} == 0)); then
    return 1
  fi

  local sel=0
  local count=${#arr[@]}

  while true; do
    banner
    echo -e "${BOLD}${prompt}${RESET}"
    echo
    echo -e "Use ${UNDER}↑ / ↓${RESET} and Enter to choose. ${DIM}(Esc to cancel)${RESET}"
    echo

    local i=0
    for path in "${arr[@]}"; do
      if [[ $i -eq $sel ]]; then
        echo -e "${GREEN}>${RESET} ${BOLD}${path}${RESET}"
      else
        echo -e "  ${path}"
      fi
      ((i++))
    done

    IFS= read -rsn1 key || continue
    if [[ $key == $'\x1b' ]]; then
      IFS= read -rsn2 -t 0.001 key2 || key2=""
      case "$key2" in
        "[A") ((sel--)); ((sel<0)) && sel=$((count-1));;
        "[B") ((sel++)); ((sel>=count)) && sel=0;;
        "")
             return 1 ;;
      esac
    elif [[ $key == "" ]]; then
      SELECTED_ITEM="${arr[$sel]}"
      return 0
    fi
  done
}

backup_p12_on_this_server() {
  banner
  echo -e "${BOLD}Backup P12 File on This Server${RESET}"
  echo
  echo "This will scan common locations on this server for *.p12 files:"
  echo "  ${DIM}/root, /home, /var/tessellation, /opt (depth ≤ 5, skipping hash/ordinal)${RESET}"
  echo "You can then pick one to copy into a safe backup folder."
  echo
  if ! prompt_yes_no "Start the scan now?" "Y"; then
    warn "Scan cancelled."
    pause_any_key
    return 0
  fi

  msg "Scanning for P12 files..."
  search_p12_files

  if ((${#P12_FOUND[@]} == 0)); then
    err "No P12 files found in the common search locations."
    echo
    echo "If your P12 is stored somewhere unusual, you may need to copy it manually."
    pause_any_key
    return 0
  fi

  if ! select_from_list P12_FOUND "Select the P12 file you want to back up:"; then
    warn "Selection cancelled."
    pause_any_key
    return 0
  fi

  local src="$SELECTED_ITEM"
  local backup_root="$HOME/p12-backups"
  mkdir -p "$backup_root"

  local base_name
  base_name="$(basename "$src")"
  local dest="$backup_root/$base_name"

  if [[ -e "$dest" ]]; then
    echo
    warn "A file named ${BOLD}$base_name${RESET} already exists in:"
    echo "  ${BOLD}$backup_root${RESET}"
    echo
    if ! prompt_yes_no "Overwrite this existing backup file?" "N"; then
      warn "Backup aborted; existing file left untouched."
      pause_any_key
      return 0
    fi
  fi

  msg "Copying:"
  echo "  From: ${src}"
  echo "  To  : ${dest}"
  cp -f "$src" "$dest"

  ok "P12 file backed up to: ${BOLD}$dest${RESET}"
  echo
  echo "From your local machine you can now download this backup with a command like:"
  echo
  echo -e "  ${GREEN}scp your_user@your_old_server_ip:${backup_root}/${base_name} ./ ${RESET}"
  echo
  echo "Or move it to your new server, for example:"
  echo
  echo -e "  ${GREEN}scp your_user@your_old_server_ip:${backup_root}/${base_name} your_user@your_new_server_ip:~/ ${RESET}"
  echo
  pause_any_key
}

show_p12_upload_help() {
  banner
  echo -e "${BOLD}Upload P12 from Windows or macOS${RESET}"
  echo
  echo "To keep things simple for non-technical users, the actual upload is done"
  echo "from your local computer using dedicated helper scripts that show a"
  echo "file browser (no typing of paths)."
  echo
  echo -e "${BOLD}On Windows:${RESET}"
  echo "  1) Open PowerShell on your local computer."
  echo "  2) Follow the instructions in the Windows P12 uploader:"
  echo "     https://github.com/StardustCollective/NodeCloud/blob/main/scripts/uploadP12/windows/ReadMe.md"
  echo
  echo -e "${BOLD}On macOS:${RESET}"
  echo "  1) Open Terminal on your Mac."
  echo "  2) Follow the instructions in the macOS P12 uploader:"
  echo "     https://github.com/StardustCollective/NodeCloud/blob/main/scripts/uploadP12/macos/ReadMe.md"
  echo
  echo "Those tools will:"
  echo "  • Let you pick the .p12 file with a graphical file chooser"
  echo "  • Verify the P12 password locally"
  echo "  • Show the P12 alias (friendlyName)"
  echo "  • Upload the P12 securely to your server home directory"
  echo
  echo -e "${BOLD}Tip:${RESET} Once the P12 is uploaded, re-run this wizard on the server"
  echo "to continue any remaining setup steps."
  echo
  pause_any_key
}

full_server_setup() {
  banner
  echo -e "${BOLD}Full Server Setup${RESET}"
  echo

  local current_user
  current_user="$(id -un)"

  if [[ $EUID -eq 0 ]]; then
    echo -e "You are currently logged in as ${BOLD}root${RESET}."
    echo
    echo "Recommended flow:"
    echo "  1) Create a non-root sudo user (for example: nodeadmin)"
    echo "  2) Disable root SSH access for better security"
    echo "  3) Reconnect as the new user"
    echo
    if prompt_yes_no "Run the non-root user setup script now?" "Y"; then
      run_create_sudo_script || true
    else
      warn "Skipped user creation. You can run that option later from the menu."
    fi
    echo
    echo -e "${BOLD}When you are done:${RESET}"
    echo "  • Disconnect from this session"
    echo "  • Log back in as your new sudo user"
    echo "  • Run this wizard again and pick 'Full Server Setup' or other options"
    echo
    pause_any_key
    return 0
  fi

  echo -e "You are logged in as non-root user: ${BOLD}${current_user}${RESET}"
  echo
  echo "Full setup from here usually means:"
  echo "  1) Ensuring your P12 is safely backed up from the old server"
  echo "  2) Uploading your P12 to this new server"
  echo "  3) Running your validator / node-specific tools afterwards"
  echo
  if prompt_yes_no "Do you still need to back up the P12 from an old server?" "Y"; then
    echo
    echo "To back up from the OLD server:"
    echo "  1) SSH into the OLD server."
    echo "  2) Download and run this same wizard there:"
    echo
    echo -e "     ${GREEN}curl -fsSL -o server_setup_wizard.sh \\"
    echo -e "       https://github.com/StardustCollective/NodeCloud/raw/main/scripts/server_setup_wizard.sh && \\"
    echo -e "     bash server_setup_wizard.sh${RESET}"
    echo
    echo "  3) Choose '${BOLD}Backup P12 File on This Server${RESET}' in the menu."
    echo "  4) Copy the backed-up P12 from ~/p12-backups to your local machine or directly to this new server."
    echo
  else
    warn "OK, skipping the backup step. Make sure your P12 is already stored safely."
    echo
  fi

  echo -e "${BOLD}Next: Upload P12 to this NEW server${RESET}"
  echo
  echo "From your local computer (Windows or macOS), use the P12 upload helper"
  echo "scripts which provide a graphical file chooser."
  echo
  echo -e "You can view the exact commands anytime by selecting:"
  echo -e "  ${BOLD}'Show P12 Upload Instructions (Windows / macOS)'${RESET} in the menu."
  echo
  echo -e "Once the P12 is uploaded to this user's home directory, you're ready to"
  echo "continue with Server / node setup."
  echo
  pause_any_key
}

main() {
  intro
  menu_loop
}

main "$@"
