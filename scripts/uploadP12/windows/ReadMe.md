# Windows P12 Upload Tool

This tool helps you upload your `.p12` file to your Ubuntu server and (optionally) set up SSH key login. It runs in a guided step-by-step mode with simple prompts and file pickers.

---

## 1️⃣ Quick Start

Copy and paste this command into **Command Prompt (CMD)** and press **Enter**:

```cmd
powershell -NoProfile -Command "$s=\"$env:TEMP\upload-p12.ps1\"; iwr 'https://raw.githubusercontent.com/StardustCollective/NodeCloud/main/scripts/uploadP12/windows/upload-p12.ps1' -OutFile $s; powershell -NoProfile -ExecutionPolicy RemoteSigned -File $s; Remove-Item -LiteralPath $s -Force"
```

---

## 2️⃣ What you’ll be asked

The script guides you through four simple steps:

### **Step 1 – Choose your `.p12` file**

A file browser will pop up. Choose your validator `.p12`.

### **Step 2 – Enter your `.p12` password**

The script verifies the password before uploading.

### **Step 3 – Choose how to connect to your server**

You can:

* Use an SSH key (browse to select one), **or**
* Use your server password

### **Step 4 – Enter your server login info**

You’ll enter:

* Your server IP
* Your username (example: `root`, `nodeadmin`)

The script uploads the `.p12` to the server’s **home directory (`~`)**.

---

## 3️⃣ Optional: Set up SSH key login

If you logged in using a password, the script will ask if you want to set up SSH key-based login.

You can:

* **Generate a new SSH key**, or
* **Import an existing SSH key**

The tool installs the public key automatically and can test the login for you.

---

## 4️⃣ Requirements

* Windows 10 or 11
* PowerShell (built in)
* OpenSSH Client

  * If missing, install it under **Settings → Apps → Optional Features → Add OpenSSH Client**

---

## 5️⃣ Troubleshooting (quick)

* If you enter the wrong `.p12` password, you’ll be asked to try again.
* If your server’s SSH fingerprint changed, the tool will fix it automatically.
* If your SSH key doesn’t work, the tool will tell you and let you try another.

---

## 6️⃣ Cleanup

The one-liner deletes the script automatically.
Nothing is left on your system.
