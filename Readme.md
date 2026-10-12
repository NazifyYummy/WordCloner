# 📄 WordCloner ✉️

**Letter Issuance System** 🚀 — A desktop application for generating personalized Word documents from templates in bulk or one-by-one. It was specifically designed to **dramatically reduce the time required to issue repetitive letters and circulars in administrative offices and organizations** 🏢, where the same template must be filled with different data many times per day. ⏱️

WordCloner lets you define a `.docx` template with placeholder variables (e.g. `##name##`, `##date##`), then issue letters either through a clean 🎨 GUI or directly from the command line 💻. It also supports batch issuance from Excel workbooks 📊 — perfect for producing hundreds of personalized letters from a single spreadsheet. 🎯

---

## 📚 Table of Contents

- [✨ Features](#-features)
- [🖼️ Screenshots](#️-screenshots)
- [💾 Installation](#-installation)
- [🚀 Quick Start (GUI)](#-quick-start-gui)
- [📝 Template Syntax](#-template-syntax)
- [⭐ Quick Access](#-quick-access)
- [📊 Batch Issuance from Excel](#-batch-issuance-from-excel)
- [💻 Command Line Interface (CLI)](#-command-line-interface-cli)
- [⚙️ Settings](#️-settings)
- [🎬 Video Guide](#-video-guide)
- [🧪 Practice Files](#-practice-files)
- [🗂️ Project Structure](#️-project-structure)
- [🛠️ Troubleshooting](#️-troubleshooting)
- [🤖 AI Assistant Notes](#-ai-assistant-notes)
- [📜 License](#-license)

---

## ✨ Features

- 🏢 **Built for administrative efficiency** — Drastically cuts the time spent on issuing repetitive letters and circulars in offices and organizations.
- 📝 **Template-based document generation** — Replace `##variable##` placeholders inside any `.docx` file.
- ⭐ **Quick Access** — Pin your favorite templates and launch them from the home screen or even the system tray. 🖥️
- 🖥️ **System tray integration** — Runs quietly in the tray; quick-launch templates from the tray menu. 🔔
- 🔍 **Auto font sizing** — The app can infer the appropriate font size from the surrounding text of the placeholder.
- 🔧 **Configurable placeholder wrapper** — Default is `##` but you can use `{{ }}`, `%%`, or anything you like.
- **Native Aero Snap support** on Windows 11/10.

---

## 🖼️ Screenshots

> 📸 _Replace the placeholders below with real screenshots of your application._

| 🏠 Home | ✉️ Issue Letter |
|------|--------------|
| ![Home](docs/screenshots/home.png) | ![Issue](docs/screenshots/issue.png) |

| ➕ Add Letter | 📊 Batch Issuance |
|------------|----------------|
| ![Add](docs/screenshots/add.png) | ![Batch](docs/screenshots/batch.png) |

| ⚙️ Settings | ⭐ Quick Access |
|----------|--------------|
| ![Settings](docs/screenshots/settings.png) | ![QuickAccess](docs/screenshots/quick_access.png) |

---

## 💾 Installation

### 🪟 Windows (recommended)

1. ⬇️ Download **`WordCloner-setup-64bit.exe`** (Inno Setup installer).
2. 🖱️ Run the installer and follow the wizard.

---

## 🚀 Quick Start (GUI)

1. 🚀 **Launch WordCloner.** On the first run you will be asked to choose your language (**Persian** or **English** ).
2. ➕ Click **"Add Letter"** from the dropdown on the home screen.
3. 📝 Fill in:
   - 🏷️ **Letter name** — a friendly name shown in the dropdown.
   - 📄 **Word template file** — drag & drop or click to select a `.docx` file that contains `##variable##` placeholders.
4. ✅ Click **"Confirm and Save"**. WordCloner will scan the file 🔍, extract every variable, and add the template to your library 📚.
5. 🏠 Back on the home screen, click the dropdown and pick your template — or add it to ⭐ **Quick Access** for one-click use.
6. 🎨 Fill in the variable values, optionally adjust font size / bold / italic / color per variable, then click **"Confirm and Issue"** ✅.

📤 The output `.docx` is written to the output folder configured in ⚙️ **Settings** (default: 🖥️ Desktop).

---

## 📝 Template Syntax

✏️ Inside your `.docx` file, mark variables using the configured wrapper. By default, that wrapper is `##`:

```
Dear ##customer_name##,

Your invoice number ##invoice_no## for the amount of ##amount## is due on ##due_date##.

Sincerely,
##sender_name##
```

📏 Rules:

- 🔧 The **wrapper** is configurable in **Settings → Placeholder Wrapper**.
- ⚖️ The number of wrapper tokens must be **even** — every opening `##` must have a matching closing `##`.
- 🔤 Variable names are **case-sensitive** and **trimmed** of surrounding whitespace.
- 🔁 A variable that appears more than once in the document will be replaced everywhere.
- 🎨 Font size, bold, italic, and color for a variable are controlled **per variable** in the issue page, not inside the document.

---

## ⭐ Quick Access

⭐ Quick Access is a shortlist of templates you use most often.

- ➕ Click the **⊕** button next to the **"Quick Access"** header to add or remove templates.
- 🔀 Reorder items by dragging the grip handle on the left.
- ⚡ Launch a template instantly by clicking its card on the home page.
- 🖥️ The same list also appears in the **system tray menu** for fast issuance without opening the main window.

---

## 📊 Batch Issuance from Excel

📊 Generate many letters at once from an Excel workbook.

1. ✉️ Open a template and click **"Batch Excel Issue"** in the header.
2. 📋 Prepare your workbook:
   - 🔢 **Column 1** — output file name (`.docx` is added automatically if missing).
   - 🔤 **Columns 2, 3, 4, …** — values for the variables **in the order they first appear** in the template.
3. ✅ Optionally tick **"Header row"** if row 1 contains column titles.
4. 📥 Drag & drop the `.xlsx` / `.xlsm` file onto the drop zone.
5. 🎨 Optionally adjust font size, bold, italic, and color for each variable.
6. ✅ Click **"Confirm"** — WordCloner writes every document to a 📁 `Bulk Issuance` folder inside your configured output directory.

> ⚠️ **Note:** The number of data columns must match `1 + number of variables`. Extra non-empty columns trigger a validation error. 🚫

---

## 💻 Command Line Interface (CLI)

💻 When installed via the Inno Setup package, the `wordcloner` command is available system-wide. 🌐

### 🔢 Count variables in a template

```bat
wordcloner -i "letter_template.docx" --count
```

### 📄 Issue a single document

```bat
wordcloner -i "letter_template.docx" -o "output.docx" ^
  -v "Ali Rezaei" "INV-2024-001" "Operations"
```

- 🔢 Values are given **in the order the variables first appear** in the template.
- ␣ Pass a **quoted single space** `" "` for an intentionally blank value.
- ♻️ Use `--force` to overwrite an existing output file.

### 📊 Batch issuance from Excel

```bat
wordcloner -i "letter_template.docx" --excel "data.xlsx" --header-row
```

- 📊 `--excel` reads column 1 as the output filename and the remaining columns as variable values.
- ⏭️ `--header-row` skips row 1.
- 📁 `-o` can be used to override the output folder (must be a folder, not a file).

### 🧾 Options summary

| 🚩 Flag | 📖 Description |
|------|-------------|
| `-i, --input` | 📄 Path to the `.docx` template. |
| `-o, --output` | 📁 Output `.docx` file (single mode) or folder (batch mode). |
| `-v, --values` | 🔤 Values in first-appearance order. |
| `-w, --wrapper` | 🔧 Override the placeholder wrapper for this run. |
| `-x, --excel` | 📊 Excel workbook for batch issuance. |
| `--header-row` | ⏭️ Skip row 1 of the worksheet. |
| `-c, --count` | 🔢 Print the number of unique variables and exit. |
| `--force` | ♻️ Overwrite the single-document output file if it exists. |
| `--gui` | 🖥️ Force the GUI to launch (for shortcuts). |

### ❓ Getting help

```bat
wordcloner --help
```

---

## ⚙️ Settings

⚙️ Open **Settings** from the home screen.

| 🛠️ Setting | 📖 Description |
|---------|-------------|
| 🌍 **App Language** | Persian or English. Restarts the UI on change. 🔄 |
| 📁 **Output Path** | Where issued documents are written. |
| 🔧 **Placeholder Wrapper** | The token used to mark variables. Default `##`. |
| ⭐ **Quick Access** | Managed from the home screen, not Settings. |

---

## 🎬 Video Guide

🎥 A built-in video tutorial is bundled with the application.

- 🏠 On the home screen, look for the **"راهنما / Guide"** button **directly below the ⚙️ Settings button** in the top-right (or top-left in Persian RTL layout) corner. 👀
- 🖱️ Clicking it opens the language-appropriate video:
  -  **`Persian_guide.mp4`** for Persian UI
  -  **`English_guide.mp4`** for English UI
- 📁 The videos live in the `Guide/` folder next to the executable.

---

## 🧪 Practice Files

🧪 Sample material is provided inside the 📁 `Practice/` folder so you can try everything without preparing your own files first:

| 📄 File | 🎯 Purpose |
|------|---------|
| `WordCloner_Test_Template_EN.docx` |  English sample template with placeholders. |
| `WordCloner_Test_Template_FA.docx` |  Persian sample template with placeholders. |
| `WordCloner_Test_Data_EN.xlsx`   | 📊 English batch data ready for the template above. |
| `WordCloner_Test_Data_FA.xlsx`   | 📊 Persian batch data ready for the template above. |

🎯 **Try it now:**

1. 👀 Open the template in Word to inspect the `##variable##` tokens.
2. ➕ Add it to WordCloner using **Add Letter**.
3. ✍️ For a single document, fill the fields manually.
4. 📊 For batch mode, use the matching `.xlsx` file with **Batch Excel Issue**.

---

## 🤖 AI Assistant Notes

🤖 If you are an AI assistant helping a user with WordCloner, or a developer looking to modify it, please first read the file:

```
AI-Readme.md
```

📖 That file contains:
- 🗺️ A concise architectural map of the codebase.
- 🗃️ The data model for `templates.json` and `settings.json`.
- 🔍 The placeholder extraction pipeline.
- 🔌 Common extension points (adding languages, adding settings, adding output formats).
- ⚠️ Known limitations and gotchas.

💡 Paste `AI-Readme.md` (and, if needed, `WordCloner.py`) into your AI tool of choice to get accurate, context-aware answers. 🎯

---

## 📜 License

📜 See the [`LICENSE`](LICENSE) file for details.

---

**📄 WordCloner ✉️** — built to reduce the time spent issuing repetitive letters and circulars in administrative offices. 🏢⏱️🚀