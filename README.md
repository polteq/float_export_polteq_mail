# Timesheet Automation (`float_export_polteq_mail`) 🚀

Automated processing of Float/People timesheet exports to Excel and PDF formats.

---

## 🚀 Snel aan de slag (NL)
Zie [START_GUIDE_NL.md](START_GUIDE_NL.md) voor een korte, lichte handleiding in het Nederlands.

---

## 🚀 Getting Started

### 1. Getting the Files
First, get the project files on your computer:
- **Option A (Recommended):** Clone the repository:
  `git clone https://github.com/polteq/float_export_polteq_mail.git`
- **Option B:** Download the [ZIP file](https://github.com/polteq/float_export_polteq_mail/archive/refs/heads/main.zip) and extract it to a folder of your choice.

### 2. Initial Setup
You only need to do this once. Ensure you have **Python 3** installed.

**Windows:**
Double-click `setup_windows.bat`.

**macOS:**
Double-click `setup_mac.command`.  
*Note: If you get an "appropriate access privileges" error, run `chmod +x setup_mac.command` in your terminal.*

This will:
- Install all necessary dependencies.
- Ask for your Name, Client Name, and Target Folder (e.g., your OneDrive folder).
- Create a **"Process Timesheet"** icon on your Desktop.

### 2. Daily Usage (Drag & Drop)
1. Download your `*-LoggedTime-*.csv` file from Float.
2. **Drag and drop** the CSV file onto the **"Process Timesheet"** icon on your Desktop.

**What happens next?**
1. **Excel & PDF:** Beautifully formatted files are saved to your configured Target Folder.
2. **Archiving:** Your original CSV is moved to the `converted/` folder to keep your desktop clean.
3. **Email:** A draft email opens in Outlook with the files attached, ready to send!

---

## 💡 Troubleshooting
- **macOS:** If the app doesn't open, Right-Click -> Open to allow execution the first time.
- **Windows:** Ensure Microsoft Excel is installed for the best PDF quality.
- **Config:** Want to change your name or client? Run the setup again or edit `config.json`.

---
**Maintainer:** Jorre van Munster
