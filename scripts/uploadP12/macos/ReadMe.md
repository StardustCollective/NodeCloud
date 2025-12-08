# macOS P12 Upload Tool

This tool guides you through securely uploading your `.p12` file to an Ubuntu server and optionally setting up SSH key-based login. It uses simple prompts and native macOS file picker dialogs—no technical experience required.

---

## 1️⃣ Quick Start (Copy & Paste)

Run this in **Terminal**:

DO NOT run this command inside of your Ubuntu server! It will fail to work. It must be run in a fresh Terminal on your MacOS.

```bash
curl -fsSL https://raw.githubusercontent.com/StardustCollective/NodeCloud/main/scripts/uploadP12/macos/upload-p12-mac.sh -o upload-p12-mac.sh \
  && chmod +x upload-p12-mac.sh \
  && ./upload-p12-mac.sh
```

This will:

1. Download the latest macOS script
2. Make it executable
3. Run it
4. Leave no permanent installation on your system

---

## 2️⃣ What Happens When You Run It

The script walks you through several simple steps.

---

### **STEP 1 — Select your `.p12` file**

A macOS file picker opens.
Choose your validator `.p12` file.

---

### **STEP 2 — Enter your `.p12` password**

The script checks that the password is correct **before uploading anything**.

* Password is verified locally using OpenSSL
* Password is never sent to the server
* You get up to 12 attempts
* After the password is verified, the script reads the `.p12` **alias (friendlyName)** and displays it so you can write it down and keep it documented

---

### **STEP 3 — Enter your server information**

You’ll be asked for:

* **Server IP or hostname**
* **SSH username** (e.g., `root`, `nodeadmin`)

The `.p12` will be uploaded into that user’s **home directory (`~/`)**.

---

### **STEP 4 — Choose your SSH authentication method**

You may authenticate using:

#### ✔ An existing SSH private key

You can browse and select any OpenSSH private key in your system (`id_ed25519`, `id_rsa`, etc.).

PuTTY `.ppk` keys are **not supported on macOS**, since macOS does not include PuTTY or PuTTYgen.

#### ✔ Or your server password

If no key is selected, the tool falls back to password-based SSH automatically.

---

### **STEP 5 — Securely upload the `.p12` file**

The script:

* Uploads your file over `scp`
* Automatically accepts unknown SSH host fingerprints
* Automatically repairs known_hosts issues
* Retries the upload if needed

If the upload succeeds, you’ll see a confirmation message.

---

## 3️⃣ Optional: SSH Key Setup (If You Used Password Authentication)

If you logged in using a password, the tool can help you switch to SSH key-based login.

You may choose to:

### 🔹 Generate a new SSH key (recommended)

* You name the key
* Script ensures you never overwrite existing keys
* Public key is installed automatically on the server
* Script can test the key immediately
* Uses modern `ed25519` keys (secure and recommended)

### 🔹 Import an existing SSH key

* Browse to select an existing OpenSSH private key
* Public key is installed on the server
* Script can test authentication

SSH key login allows passwordless access to your server.

---

## 4️⃣ Requirements

* **macOS Monterey, Ventura, Sonoma, or later**
* Built-in macOS tools:

  * Terminal
  * OpenSSL
  * OpenSSH Client (`ssh`, `scp`, `ssh-keygen`)

All required tools come preinstalled on macOS.

---

## 5️⃣ Cleanup

* The optional one-liner downloads the script to your current folder
* You can delete `upload-p12-mac.sh` at any time
* No system modifications are made
* Only keys that **you choose to generate** remain permanently
