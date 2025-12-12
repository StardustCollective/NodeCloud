# Windows P12 Upload Tool

This tool guides you through securely uploading your `.p12` file to your Ubuntu server and optionally setting up SSH key-based login.

It uses simple prompts, Windows file pickers, and clear instructions-no advanced Linux knowledge required.

---

## 1️⃣ Quick Start (CMD Only)

IMPORTANT:

• Run this in **Windows Command Prompt (CMD)**
• Do NOT run inside PowerShell
• Do NOT run inside WSL
• Do NOT run this on your Ubuntu server

Open CMD and paste this command:

```cmd
powershell -NoProfile -Command "$s=\"$env:TEMP\upload-p12.ps1\"; iwr 'https://raw.githubusercontent.com/StardustCollective/NodeCloud/main/scripts/uploadP12/windows/upload-p12.ps1' -OutFile $s; powershell -NoProfile -ExecutionPolicy RemoteSigned -File $s; Remove-Item -LiteralPath $s -Force"
```

This will:

1. Download the latest script to your TEMP folder
2. Run it
3. Delete it when finished

Nothing is permanently installed on your system.

---

## 2️⃣ What the Tool Does

### STEP 1 - Select your `.p12` file

A Windows file picker opens.
Choose the `.p12` file you want to upload.

If the picker cannot open, you may type the file path manually.

---

### STEP 2 - Verify your `.p12` password

The script:

• Prompts for the `.p12` password (hidden input)
• Verifies it **locally**
• Never sends the password to the server
• Allows up to 12 attempts
• Shows the certificate **friendlyName / alias** if present (recommended to write down)

If the password cannot be verified, the script exits without uploading.

---

### STEP 3 - Enter server information

You will enter:

• Server IP or hostname
• SSH username (`root`, `nodeadmin`, etc.)

The `.p12` will be uploaded into that user’s **home directory (`~/`)**.

If your SSH known_hosts file contains an old entry for that server, the script removes it automatically to avoid “REMOTE HOST IDENTIFICATION HAS CHANGED” errors.

---

### STEP 4 - Choose your SSH authentication method

You may authenticate using:

---

### ✔ Option A - SSH Private Key (recommended)

You can browse and select a standard **OpenSSH** private key, such as:

• id_ed25519
• id_rsa
• Any key beginning with:
`-----BEGIN … PRIVATE KEY-----`

If you select a **PuTTY Private Key (`.ppk`)**, Windows cannot convert it automatically.
Instead, the script displays a guided instruction set:

---

### How to export a normal SSH key from a `.ppk` file (with PuTTYgen)

1. Open **PuTTYgen**
   (Install PuTTY from [https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html](https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html) if you do not have it)

2. In PuTTYgen:
   **Conversions → Import Key**
   Select your `.ppk` file

3. Enter your `.ppk` passphrase if prompted

4. Then export an OpenSSH key:
   **Conversions → Export OpenSSH Key (force new file format)**

5. Save it (recommended folder):
   C:\Users<username>.ssh\myserver-key

6. Return to the script and select the **newly exported key**

---

### ✔ Option B - Password SSH login

If you do not select a key, the script falls back to password authentication.

When `scp` or `ssh` run, you will be prompted for your server password.

---

### STEP 5 - Upload your `.p12` securely

The script:

• Uploads the file via `scp`
• Automatically accepts new host fingerprints
• Removes stale known_hosts entries
• Retries if necessary
• Prints clear upload success or failure messages

The file is uploaded into `~/` on the server.

---

### STEP 6 - Optional: SSH Key Setup (if you used password auth)

If you logged in with SSH password authentication, the script can help you set up key-based login.

You may:

#### Generate a new SSH key pair

• You choose the name
• A new keypair is created via `ssh-keygen`
• The public key is installed into `~/.ssh/authorized_keys` on the server
• Login is tested immediately

#### Import an existing SSH key

• Select an existing OpenSSH private key
• `.ppk` files again trigger the PuTTYgen instructions
• Public key is extracted and installed on the server
• Login is tested

#### Skip SSH key setup

• No changes made

---

## 3️⃣ Requirements

• Windows 10 or Windows 11
• PowerShell (preinstalled)
• OpenSSH Client
• PuTTYgen (only needed for `.ppk` users)

To install **OpenSSH Client**:

Settings → Apps → Optional Features → Add OpenSSH Client

---

## 4️⃣ Cleanup

This tool cleans up after itself:

• The downloaded script in `%TEMP%` is automatically deleted
• No services or components are installed
• No registry changes are made

Only permanent items created are:

• SSH keys you generate or import
• The `.p12` file you intentionally uploaded to the server

---

## 5️⃣ Support

Stardust Collective - @Proph151Music
[https://github.com/StardustCollective/NodeCloud/tree/main/scripts/uploadP12/windows/](https://github.com/StardustCollective/NodeCloud/tree/main/scripts/uploadP12/windows/)
