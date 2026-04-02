# Office Click-to-Run + KMS Lab Deployment

This repository contains a lab-oriented script and configuration used to test deploying a modern Click-to-Run version of Microsoft Office while using a Volume Licensing (KMS) activation model.

The goal of this project is to reproduce and analyze behavior differences between MSI-based Office 2016 and Click-to-Run Office builds, particularly on Windows 11 systems.

---

## ⚠️ Disclaimer

This project is intended for **lab and educational purposes only**.

Always ensure that your deployment and licensing approach complies with your organization's Microsoft licensing agreement. The author does not recommend using this approach in production without proper validation and licensing review.

---

## 📌 Overview

In legacy environments, MSI-based Office 2016 can show stability issues on modern operating systems, especially when working with files stored on network shares.

This lab setup explores an alternative approach:

- Deploy a Click-to-Run Office build
- Apply Volume Licensing components
- Activate via KMS
- Compare behavior against MSI-based Office

---

## 📁 Contents

- `install-office-c2r-kms-lab.bat` – main deployment script  
- `S365x64OVP.xml` – Click-to-Run configuration file  

---

## ⚙️ How It Works

The script performs the following steps:

- Ensures administrative privileges  
- Deploys a Click-to-Run version of Office  
- Applies Volume Licensing components  
- Configures KMS activation  

The configuration file defines:

- Office edition and products  
- Language settings  
- Update behavior  
- Optional app exclusions  

---

## 🛠️ Required Changes Before Use

Before running the script, review and adjust the following:

### 1. KMS Server

Inside the script:


Replace with your KMS server:

/sethst:kms.example.local

---

### 2. Product Keys (Optional)

In `S365x64OVP.xml`:


PIDKEY="XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"


Replace with your actual Volume License keys if required.

In many KMS environments, this step is not necessary.

---

### 3. Source Path


SourcePath="Office64"


Make sure the `Office64` folder exists and contains Office installation files.

Alternatively, enable CDN fallback in the XML:


AllowCdnFallback="TRUE"


---

### 4. Language

```<Language ID="uk-ua" /> ```

Adjust to your preferred language:

en-us
ru-ru
etc.

▶️ Usage  
Place the script and configuration file in the same directory
Ensure Office installation files are available (if using local source)
Run the script as Administrator  
  
📊 Notes  
This setup is designed for testing and experimentation
Some steps are intentionally verbose to observe licensing behavior
The script is not optimized for production deployment

📖 Full Article  

A detailed breakdown of this experiment is available here:  

👉 https://www.hiddenobelisk.com/from-office-2016-to-modern-click-to-run-a-practical-migration-approach-using-volume-licensing-in-enterprise-environments/

💬 Feedback

If you have seen similar behavior or have improvements, feel free to open an issue or share your experience.
