#!/bin/bash

# ============================================================
#  STARDUST COLLECTIVE
#  SECURE P12 UPLOAD & SSH SETUP TOOL (macOS)
# ============================================================

CYAN="\033[36m"
GREEN="\033[32m"
YELLOW="\033[33m"
RED="\033[31m"
NC="\033[0m"

P12_ALIAS=""

clear
echo -e "${CYAN}==============================================================${NC}"
echo -e "${CYAN}                    STARDUST COLLECTIVE${NC}"
echo -e "${CYAN}           SECURE P12 UPLOAD & SSH SETUP TOOL${NC}"
echo -e "${CYAN}==============================================================${NC}"
echo ""
echo -e "${CYAN}This guided tool will help you:${NC}"
echo "  1) Select your .p12 file"
echo "  2) Verify its password locally"
echo "  3) Enter server connection info"
echo "  4) Choose SSH authentication (key or password)"
echo "  5) Upload your .p12 securely"
echo "  6) (Optional) Set up SSH key-based login"
echo ""
read -p "Press Enter to begin..."

choose_file() {
    local prompt="$1"
    local result

    result=$(osascript <<EOF
set f to choose file with prompt "$prompt"
POSIX path of f
EOF
)
    echo "$result"
}

verify_p12_password() {
    local file="$1"
    local pass="$2"

    echo "" | openssl pkcs12 -in "$file" -nokeys -passin pass:"$pass" -passout pass:"dummy" >/dev/null 2>&1
    if [[ $? -eq 0 ]]; then
        return 0
    fi

    echo "" | openssl pkcs12 -in "$file" -legacy -nokeys -passin pass:"$pass" -passout pass:"dummy" >/dev/null 2>&1
    return $?
}

extract_p12_alias() {
    local file="$1"
    local pass="$2"
    local alias

    alias=$(openssl pkcs12 -in "$file" -nokeys -info -passin pass:"$pass" 2>&1 | awk -F': ' '/friendlyName/ {print $2; exit}')
    if [[ -z "$alias" ]]; then
        alias=$(openssl pkcs12 -in "$file" -legacy -nokeys -info -passin pass:"$pass" 2>&1 | awk -F': ' '/friendlyName/ {print $2; exit}')
    fi

    if [[ -n "$alias" ]]; then
        P12_ALIAS="$alias"
        return 0
    else
        return 1
    fi
}

KNOWN_HOSTS="$HOME/.ssh/known_hosts"
mkdir -p "$HOME/.ssh"
touch "$KNOWN_HOSTS"

clean_known_host() {
    local server="$1"
    echo -e "${YELLOW}! Cleaning SSH host entry for $server${NC}"
    ssh-keygen -R "$server" >/dev/null 2>&1
}

generate_ssh_key() {
    local username="$1"
    local server="$2"

    mkdir -p "$HOME/.ssh"

    while true; do
        echo ""
        read -p "Enter a name for your new SSH key (no extension): " keyname

        if [[ -z "$keyname" ]]; then
            echo -e "${YELLOW}! Name cannot be empty.${NC}"
            continue
        fi

        local keypath="$HOME/.ssh/$keyname"

        if [[ -e "$keypath" ]]; then
            echo -e "${RED}- A key with that name already exists. Choose another.${NC}"
            continue
        fi

        echo -e "${CYAN}Generating SSH key pair...${NC}"
        ssh-keygen -t ed25519 -f "$keypath" -N "" -C "nodecloud-$username@$server"

        if [[ $? -eq 0 ]]; then
            echo -e "${GREEN}+ SSH key generated: $keypath${NC}"
            echo "$keypath"
            return
        else
            echo -e "${RED}- Failed to generate SSH key. Try again.${NC}"
        fi
    done
}

install_pubkey() {
    local username="$1"
    local server="$2"
    local keypath="$3"

    echo -e "${CYAN}Installing SSH public key on server...${NC}"

    local pubkey
    pubkey=$(ssh-keygen -y -f "$keypath" 2>/dev/null)
    if [[ $? -ne 0 ]]; then
        echo -e "${RED}- Failed to extract public key.${NC}"
        return 1
    fi

    ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile="$KNOWN_HOSTS" \
        "$username@$server" \
        "mkdir -p ~/.ssh && chmod 700 ~/.ssh && echo '$pubkey' >> ~/.ssh/authorized_keys && chmod 600 ~/.ssh/authorized_keys" \
        >/dev/null 2>&1

    local code=$?

    if [[ $code -ne 0 ]]; then
        ssh_output=$(ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile="$KNOWN_HOSTS" "$username@$server" true 2>&1)
        if [[ "$ssh_output" == *"REMOTE HOST IDENTIFICATION HAS CHANGED"* ]]; then
            clean_known_host "$server"
            ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile="$KNOWN_HOSTS" \
                "$username@$server" \
                "mkdir -p ~/.ssh && chmod 700 ~/.ssh && echo '$pubkey' >> ~/.ssh/authorized_keys && chmod 600 ~/.ssh/authorized_keys" \
                >/dev/null 2>&1
            code=$?
        fi
    fi

    if [[ $code -eq 0 ]]; then
        echo -e "${GREEN}+ SSH key installed.${NC}"
        return 0
    else
        echo -e "${RED}- Failed to install SSH key.${NC}"
        return 1
    fi
}

test_ssh_key() {
    local username="$1"
    local server="$2"
    local keypath="$3"

    echo -e "${CYAN}Testing SSH key authentication...${NC}"
    echo -e "${YELLOW}SSH may prompt you for the key passphrase.${NC}"

    ssh -i "$keypath" -o StrictHostKeyChecking=no -o UserKnownHostsFile="$KNOWN_HOSTS" \
        "$username@$server" "echo ok" >/dev/null 2>&1

    if [[ $? -eq 0 ]]; then
        echo -e "${GREEN}+ SSH key authentication succeeded.${NC}"
        return 0
    else
        echo -e "${RED}- SSH key authentication failed.${NC}"
        return 1
    fi
}

echo ""
echo -e "${CYAN}STEP 1: Select your .p12 file${NC}"

P12_FILE=$(choose_file "Select your .p12 file")
if [[ -z "$P12_FILE" ]]; then
    echo -e "${RED}- No file selected. Exiting.${NC}"
    exit 1
fi

echo -e "${GREEN}+ Selected: $P12_FILE${NC}"

for attempt in {1..12}; do
    read -s -p "Enter .p12 password (attempt $attempt of 12): " P12_PASS
    echo ""

    verify_p12_password "$P12_FILE" "$P12_PASS"
    if [[ $? -eq 0 ]]; then
        echo -e "${GREEN}+ Password verified.${NC}"

        if extract_p12_alias "$P12_FILE" "$P12_PASS"; then
            echo -e "${GREEN}+ P12 alias (friendlyName): ${P12_ALIAS}${NC}"
            echo -e "${YELLOW}! Make sure to write this alias down and keep it documented for future use.${NC}"
        else
            echo -e "${YELLOW}! No friendlyName/alias was found in this P12 file.${NC}"
        fi

        break
    else
        echo -e "${RED}- Incorrect password.${NC}"
    fi

    if [[ $attempt -eq 12 ]]; then
        echo -e "${RED}- Too many incorrect attempts. Exiting.${NC}"
        exit 1
    fi
done

echo ""
echo -e "${CYAN}STEP 3: Enter server IP or hostname${NC}"
read -p "Server IP: " SERVER
if [[ -z "$SERVER" ]]; then
    echo -e "${RED}- Server IP required. Exiting.${NC}"
    exit 1
fi

echo ""
echo -e "${CYAN}STEP 4: Enter server username${NC}"
read -p "Username: " USERNAME
if [[ -z "$USERNAME" ]]; then
    echo -e "${RED}- Username required. Exiting.${NC}"
    exit 1
fi

echo ""
echo -e "${CYAN}STEP 5: Choose SSH authentication method${NC}"
read -p "Use SSH private key? [Y/n]: " USE_KEY
[[ -z "$USE_KEY" ]] && USE_KEY="Y"

SSH_KEY=""

if [[ "$USE_KEY" =~ ^[Yy]$ ]]; then
    echo ""
    echo -e "${CYAN}Select your SSH private key...${NC}"

    SSH_KEY=$(choose_file "Select your SSH private key")

    if [[ -z "$SSH_KEY" ]]; then
        echo -e "${YELLOW}! No key selected. Falling back to password auth.${NC}"
        SSH_KEY=""
    else
        echo -e "${GREEN}+ Using SSH key: $SSH_KEY${NC}"
    fi
fi

echo ""
echo -e "${CYAN}STEP 6: Uploading your .p12 file${NC}"

UPLOAD_CMD=(scp)
[[ -n "$SSH_KEY" ]] && UPLOAD_CMD+=(-i "$SSH_KEY")
UPLOAD_CMD+=(
    -o StrictHostKeyChecking=no
    -o UserKnownHostsFile="$KNOWN_HOSTS"
    "$P12_FILE"
    "$USERNAME@$SERVER:~/"
)

"${UPLOAD_CMD[@]}" >/dev/null 2>&1
UPLOAD_CODE=$?

if [[ $UPLOAD_CODE -ne 0 ]]; then
    if ssh "$USERNAME@$SERVER" true 2>&1 | grep -q "REMOTE HOST IDENTIFICATION HAS CHANGED"; then
        clean_known_host "$SERVER"
        "${UPLOAD_CMD[@]}" >/dev/null 2>&1
        UPLOAD_CODE=$?
    fi
fi

if [[ $UPLOAD_CODE -eq 0 ]]; then
    echo -e "${GREEN}+ Upload successful.${NC}"
else
    echo -e "${RED}- Upload failed.${NC}"
    read -p "Press Enter to exit..."
    echo -e "${NC}"
    exit 1
fi

if [[ -z "$SSH_KEY" ]]; then
    echo ""
    echo -e "${CYAN}STEP 7: Optional SSH key setup${NC}"
    read -p "Set up SSH key-based login now? [Y/n]: " DO_SETUP
    [[ -z "$DO_SETUP" ]] && DO_SETUP="Y"

    if [[ "$DO_SETUP" =~ ^[Yy]$ ]]; then
        echo ""
        echo -e "${CYAN}Choose an option:${NC}"
        echo "  1) Generate a NEW SSH key pair"
        echo "  2) Import an EXISTING SSH key"
        echo "  3) Skip SSH key setup"
        read -p "Enter 1, 2, or 3: " OPTION

        case "$OPTION" in
            1)
                NEWKEY=$(generate_ssh_key "$USERNAME" "$SERVER")
                if install_pubkey "$USERNAME" "$SERVER" "$NEWKEY"; then
                    test_ssh_key "$USERNAME" "$SERVER" "$NEWKEY"
                fi
                ;;
            2)
                echo ""
                IMPORTED=$(choose_file "Select your SSH private key")
                if [[ -n "$IMPORTED" ]]; then
                    if install_pubkey "$USERNAME" "$SERVER" "$IMPORTED"; then
                        test_ssh_key "$USERNAME" "$SERVER" "$IMPORTED"
                    fi
                fi
                ;;
            *)
                echo -e "${YELLOW}! SSH key setup skipped.${NC}"
                ;;
        esac
    fi
fi

if [[ -n "$P12_ALIAS" ]]; then
    echo ""
    echo -e "${CYAN}REMINDER: The alias (friendlyName) for this P12 is: ${P12_ALIAS}${NC}"
    echo -e "${YELLOW}Please keep this alias documented somewhere safe for future use.${NC}"
fi

echo ""
echo -e "${CYAN}Login reminder:${NC}"
echo -e "${CYAN}You can log into your server using this command:${NC}"

if [[ -n "$SSH_KEY" ]]; then
    echo -e "${GREEN}ssh -i \"$SSH_KEY\" ${USERNAME}@${SERVER}${NC}"
else
    echo -e "${GREEN}ssh ${USERNAME}@${SERVER}${NC}"
fi

echo ""
read -p "Press Enter to exit..."
echo -e "${NC}"
