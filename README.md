# 📊 Excel-PQ-AutoRefresher

**Excel-PQ-AutoRefresher** is a lightweight **PowerShell automation tool** that refreshes all Excel workbooks containing **Power Query connections** — fully unattended and on schedule.  
Designed for analysts, data teams, and professionals who rely on up-to-date Excel dashboards, it ensures that every connected dataset stays current without manual intervention.

With built-in logging, smart error handling, and automatic scheduling, the script sequentially opens each workbook, triggers refresh operations safely (waiting for completion), and records every step in a detailed log file. It intelligently skips files without Power Query connections, handles nested folders, and supports long-running refreshes without timeouts — all without requiring administrator permissions or Git setup.

---

## ✨ Key Features

- 🔁 **Sequential refresh** of all Excel files in main and subfolders  
- ⚙️ **Auto-skip** for files without Power Query connections  
- 📊 **Real-time progress bar** for visual feedback  
- 🕒 **Built-in scheduler** (runs automatically every few hours)  
- 🪶 **Lightweight setup** – no admin rights or Git needed  
- 🧠 **Smart wait system** – each workbook refresh takes as long as needed  
- 🧾 **Detailed logging** for tracking, auditing, and troubleshooting  
- 💬 **Error handling** and skip logic for stable unattended execution  

---

## 💼 Use Case Example

A data analyst managing weekly Excel reports can simply run **Schedule.ps1** once, and the system will automatically refresh every dataset, generate logs, and ensure all dashboards remain current — even when the system is offline.

---

## 🧩 File Structure

```

📁 Scripts
┣ 📜 ExcelRefresh.ps1      # Main refresh logic for Excel workbooks
┣ 📜 Schedule.ps1          # Creates a scheduled task to run the refresher automatically
┗ 📄 ExcelRefreshLog.txt   # Log file recording all refresh events

````

---

## ⚙️ Setup Instructions

### 1️⃣ Create a Working Folder
Create a folder (e.g., `F:\Scripts`) and place all scripts inside it.

### 2️⃣ Configure Paths in `ExcelRefresh.ps1`
Open and edit:
```powershell
$ExcelFolder = "F:\Scripts\ExcelFiles"              # Folder containing Excel files
$LogFile = "F:\Scripts\ExcelRefreshLog.txt"         # Log file path
````

### 3️⃣ Configure the Scheduler Script

In `Schedule.ps1`, set the correct script path:

```powershell
$ScriptPath = "F:\Scripts\ExcelRefresh.ps1"
```

### 4️⃣ Allow PowerShell Scripts to Run (no admin required)

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
```

### 5️⃣ Run the Scheduler

Navigate to your script folder:

```powershell
cd F:\Scripts
```

Then execute:

```powershell
.\Schedule.ps1
```

This sets up an automatic refresh task running every 3 hours for one year.

---

## 🧾 Log Example

Example entries from the generated log file:

```
[2025-11-02 12:01:47] === Excel Refresh Task Started ===
[2025-11-02 12:01:47] Refreshing F:\Scripts\ExcelFiles\All Groups (1).xlsx
[2025-11-02 12:02:05] Successfully refreshed: All Groups (1).xlsx
[2025-11-02 12:02:44] === All files refreshed successfully ===
```

---

## 🧠 Notes

* Files **without Power Query** are automatically skipped.
* Each workbook refreshes **in sequence** to avoid “pending refresh” popups.
* Fully compatible with **Excel Desktop**, **PowerShell 5+**, and **PowerShell 7+**.
* Can run **offline** — logs are written locally and tasks continue on schedule.

---

## 🔧 System Requirements

* Windows 10 or later
* Microsoft Excel (with Power Query support)
* PowerShell 5.1 or higher

---

## 🧑‍💻 Author

**Youssef Sherif Wahib**
PowerShell Automation | Excel Data Systems | Process Optimization
📧 *[Contact Me on LinkedIn](https://www.linkedin.com/in/youssef-sherif-wahib-7277191bb)*
