# Timesheet Automation (`float_export_polteq_mail`)

Automated processing of Float timesheet exports to Excel and PDF formats.

## Overview

This tool automates the conversion of CSV timesheet exports from Float/People into formatted Excel and PDF files, automatically organizing them into a designated folder (like OneDrive), and drafting an email to send them.

It is designed to be completely **Cross-Platform (Windows & macOS)** and extremely easy to use via Drag-and-Drop.

---

## 🚀 Easy Setup & Usage

### 1. Initial Setup
You only need to do this once. Ensure you have **Python 3** installed.

**Windows:**
Double-click `setup_windows.bat`.

**macOS:**
Double-click `setup_mac.command`. *(If macOS blocks it, Right-Click -> Open)*.

This will:
- Install all necessary dependencies.
- Ask for your Name, Client Name, and the folder where you want to save the generated files.
- Create a shortcut/app on your Desktop called **"Process Timesheet"**.

### 2. Daily Usage (Drag & Drop)
When you download a new `*-LoggedTime-*.csv` file from Float, simply **drag and drop** the CSV file onto the **"Process Timesheet"** shortcut on your Desktop.

The tool will:
1. Generate a beautifully formatted Excel file.
2. Generate a high-quality PDF.
3. Save both files to your configured folder.
4. Open a draft email in your local Outlook with the files attached and ready to send!

---

## AI Assistants Guide
If you are an AI assistant trying to run this tool for the user, please refer to [AI_GUIDE.md](AI_GUIDE.md).

## Configuration
Your personal details (Name, Client, Target Folder) are saved in `config.json`. You can edit this file directly at any time if your details change, or run `python shareable_processor.py --setup` to reconfigure.

## Legacy System
The older macOS-only processor (`process_timesheet.py` and `run_timesheet_processor.sh`) is still included for backwards compatibility but is considered legacy.

---
**Maintainer:** Jorre van Munster
