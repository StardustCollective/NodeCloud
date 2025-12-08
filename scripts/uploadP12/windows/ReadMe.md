# Windows P12 Upload Tool

This tool guides you through securely uploading your `.p12` file to an Ubuntu server and optionally setting up SSH key-based login. It uses simple prompts and file pickers—no technical experience required.

---

## 1️⃣ Quick Start (Copy & Paste)

Run this in **Command Prompt (CMD)**:

```cmd
powershell -NoProfile -Command "$s=\"$env:TEMP\upload-p12.ps1\"; iwr 'https://raw.githubusercontent.com/StardustCollective/NodeCloud/main/scripts/uploadP12/windows/upload-p12.ps1' -OutFile $s; powershell -NoProfile -ExecutionPolicy RemoteSigned -File $s; Remove-Item -LiteralPath $s -Force"
```

This will:

1. Download the latest script into your TEMP folder
2. Run it
3. Automatically delete it afterwards

Nothing is permanently installed on your system.

---

## 2️⃣ What Happens When You Run It

The script walks you through several simple steps.

---

### **STEP 1 — Select your `.p12` file**

A file browser opens.
Choose the `.p12` file you plan to upload to your server.

---

### **STEP 2 — Enter your `.p12` password**

The script checks that you know the correct password **before uploading anything**.

* Password is verified locally  
* Password is never sent to the server  
* You get multiple attempts if needed  
* After the password is verified, the script reads the `.p12` **alias (friendlyName)** and displays it so you can write it down and keep it documented

---

### **STEP 3 — Enter your server information**

You’ll be asked for:

* **Server IP or hostname**
* **SSH username** (e.g., `root`, `nodeadmin`)

The `.p12` will be uploaded into that user’s **home directory (`~/`)**.

---

### **STEP 4 — Choose your SSH authentication method**

You may:

#### ✔ Use an SSH private key

You can browse and select one.
The tool supports:

* OpenSSH private keys
* PuTTY `.ppk` keys (auto-converted to OpenSSH format)

If you choose a `.ppk` file:

* Script detects it
* Auto-converts it using PuTTYgen
* If PuTTYgen is not installed:

  * It will auto-install via Chocolatey (if available), **or**
  * Download PuTTYgen automatically

Your original `.ppk` is never modified.

#### ✔ Or use server password authentication

If no key is selected, the tool falls back to password-based SSH.

---

### **STEP 5 — Securely upload the `.p12` file**

The script:

* Uploads your file over SSH
* Automatically accepts unknown host fingerprints
* Automatically cleans up old/invalid host fingerprints
* Retries the upload if needed

If the upload succeeds, you’ll see a confirmation message.

---

## 3️⃣ Optional: SSH Key Setup (If You Used Password Authentication)

If you logged in with a password, the tool can help set up SSH key-based login.

You may choose to:

### 🔹 Generate a new SSH key

* You enter the key name
* The tool prevents overwriting existing keys
* Public key is installed automatically on the server
* Key authentication can be tested immediately

### 🔹 Import an existing SSH key

* Choose an existing private key
* `.ppk` keys are converted automatically
* Public key is installed on the server
* Authentication can be tested

SSH key login allows you to access your server without typing a password.

---

## 4️⃣ Requirements

* **Windows 10 or 11**
* **PowerShell** (preinstalled)
* **OpenSSH Client**

To install OpenSSH Client (if missing):

**Settings → Apps → Optional Features → Add OpenSSH Client**

---

## 5️⃣ Troubleshooting (Short List)

* Wrong `.p12` password → script asks again
* Wrong SSH password → SSH prompts again
* Incorrect server fingerprint → script fixes it automatically
* SSH key won’t authenticate → script shows the reason and offers alternatives
* PuTTY key detected → script auto-converts it
* Forgot your `.p12` alias → the script prints the alias (friendlyName) right after password verification and reminds you again at the end, so you can write it down

Most issues resolve themselves automatically.

---

## 6️⃣ Cleanup

* The one-liner **downloads the script temporarily**
* The script **deletes itself** before exiting
* Your system remains clean

Only things permanently created are the SSH keys you choose to generate or import.
