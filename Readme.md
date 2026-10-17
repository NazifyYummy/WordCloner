
# 📄 WordCloner ✉️

**Letter Issuance System** 🚀 — A desktop application for generating personalized Word documents from templates, one at a time or in bulk. It was designed to **dramatically reduce the time required to issue repetitive letters and circulars in administrative offices and organizations** 🏢, where the same template must be filled with different data many times per day. ⏱️

WordCloner lets you define a `.docx` template with placeholder variables such as `##name##` and `##date##`, then issue letters through a clean 🎨 GUI, the command line 💻, or an Excel workbook 📊. Excel batch issuance is ideal for producing hundreds of personalized letters from one spreadsheet. 🎯

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
- [⚙️ Settings and Data Location](#️-settings-and-data-location)
- [🎬 Video Guide](#-video-guide)
- [🧪 Practice Files](#-practice-files)
- [🗂️ Project Structure](#️-project-structure)
- [🛠️ Troubleshooting](#️-troubleshooting)
- [🤖 AI Assistant Notes](#-ai-assistant-notes)
- [📜 License](#-license)

---

## ✨ Features

- 🏢 **Built for administrative efficiency** — Reduces the time spent issuing repetitive letters and circulars in offices and organizations.
- ⭐ **Quick Access** — Pin frequently used templates and open them from the home screen or the system-tray menu.
- 🔔 **System-tray integration** — Keep WordCloner available in the tray and launch pinned templates from its menu.
- 🚀 **Optional Windows startup** — Off by default. When enabled, WordCloner starts with its main window hidden; clicking the tray icon opens the window.
- 🔍 **Auto font sizing** — Infer an appropriate font size from the text surrounding a placeholder.
- 🎨 **Per-variable formatting** — Set font size, bold, italic, and text color for each variable. Font-size input accepts numbers; if left blank and focus moves away, it returns to Auto. The color dialog labels its value `HEX:`.
- **Native Aero Snap support** on Windows 10/11.

---

## 🖼️ Screenshots

> 📸 **Note about the screenshots / نکته درباره اسکرین‌شات‌ها**
>
> Since most readers are likely Persian speakers, the screenshots below are shown in **Persian (فارسی)**. You can change the interface language to **English** at any time via **Settings → App Language**. And don't ask why the whole README isn't in Persian :) .

<table>
  <tr>
    <td align="center" width="50%">
      <img src="https://github.com/user-attachments/assets/151d2842-a255-42d1-977f-ce141f560318"
           alt="Home Page — template selection and Quick Access"
           width="400" />
      <br/>
      <sub><b>🏠 Home Page</b><br/>
      Select a template, pin favorites to Quick Access, or open the settings</sub>
    </td>
    <td align="center" width="50%">
      <img src="https://github.com/user-attachments/assets/3bbef273-4d9f-45f7-ae4a-02f19aaf0900"
           alt="Quick Access from the System Tray"
           width="400" />
      <br/>
      <sub><b>📌 Quick Access from the System Tray</b><br/>
      Launch a template without opening the main window</sub>
    </td>
  </tr>
  <tr>
    <td align="center" colspan="2">
      <img src="https://github.com/user-attachments/assets/bcb15c1f-8bdd-4fb5-a3aa-8826c9339b1a"
           alt="Issuance Page — variable values and formatting"
           width="400" />
      <br/>
      <sub><b>📝 Issuance Page</b><br/>
      Fill in each variable, adjust font and color per field, then generate the final letter</sub>
    </td>
  </tr>
</table>


## 💾 Installation

Grab the latest **`WordCloner_Setup.exe`** from the [Releases](https://github.com/NazifyYummy/WordCloner/releases) page and run it. The installation wizard will handle the rest.

---

## 🚀 Quick Start (GUI)

1. 🚀 **Launch WordCloner.** On the first run, choose a language (**Persian** or **English**).
2. ➕ Select **Add Letter** from the dropdown on the home screen.
3. 📝 Fill in:
   - 🏷️ **Letter name** — a friendly name shown in the dropdown.
   - 📄 **Word template file** — drag and drop a `.docx` file containing `##variable##` placeholders, or click to browse for it.
4. ✅ Select **Confirm and Save**. WordCloner scans the file, extracts its variables, and adds the template to your library.
5. 🏠 Choose the template from the home-screen dropdown, or add it to ⭐ **Quick Access** for one-click access.
6. 🎨 Enter the variable values, optionally adjust font size, bold, italic, and color for each variable, then select **Confirm and Issue**.

📤 The output `.docx` is written to the folder configured in ⚙️ **Settings** (default: the Desktop).

---

## 📝 Template Syntax

✏️ In your `.docx` file, mark variables using the configured wrapper. The default wrapper is `##`:

```text
Dear ##customer_name##,

Your invoice number ##invoice_no## for the amount of ##amount## is due on ##due_date##.

Sincerely,
##sender_name##
```

📏 Rules:

- 🔧 The wrapper can be changed in **Settings → Placeholder Wrapper**.
- ⚖️ The number of wrapper tokens must be **even**: every opening wrapper must have a matching closing wrapper.
- 🔤 Variable names are **case-sensitive** and trimmed of surrounding whitespace.
- 🔁 Repeated occurrences of a variable are replaced with the same value. Supply a value only once for that unique variable.
- 🎨 Font size, bold, italic, and color are controlled per variable in the issue page; they are not specified by the placeholder text itself.

---

## ⭐ Quick Access

⭐ Quick Access is a shortlist of templates you use most often.

- ➕ Select the **⊕** button next to the **Quick Access** heading to add or remove templates.
- 🔀 Reorder items by dragging the grip handle.
- ⚡ Select a card on the home screen to open that template.
- 🖥️ The same valid Quick Access templates also appear in the system-tray menu.

---

## 📊 Batch Issuance from Excel

📊 Generate many letters at once from an Excel workbook.

1. ✉️ Open a template and select **Batch Excel Issue** in the header.
2. 📋 Prepare the active worksheet:
   - 🔢 **Column 1** — output filename; `.docx` is added automatically if missing.
   - 🔤 **Columns 2, 3, 4, …** — values for the **unique variables**, in the order they first appear in the template.
3. ✅ Enable **Header row** if row 1 contains column titles.
4. 📥 Drag and drop the `.xlsx` or `.xlsm` file onto the drop zone, or click the zone to browse for it.
5. 🎨 Optionally adjust font size, bold, italic, and color for each variable.
6. ✅ Select **Confirm**. WordCloner writes the documents to a `Bulk Issuance` folder inside your configured output directory.

> ⚠️ The worksheet needs one filename column plus one column for each **unique** template variable. Extra non-empty columns trigger a validation error. A repeated variable in the template needs only one Excel value.

---

## 💻 Command Line Interface (CLI)

You can simply learn how to use cli jsut by reading `help` like this:

```bat
wordcloner --help
```

The short `wordcloner` command is available when the Python package is installed in editable mode, or if you separately add the executable directory to `PATH`.

### 🔢 Count variables in a template

```bat
wordcloner -i "letter_template.docx" --count
```

### 📄 Issue a single document

```bat
wordcloner -i "letter_template.docx" -o "output.docx" ^
  -v "Ali Rezaei" "INV-2024-001" "Operations"
```

- 🔢 Values are supplied in the order the **unique variables first appear** in the template.
- 🔁 If a variable occurs more than once, its one supplied value is reused at every occurrence.
- ␣ Use a quoted single space (`" "`) for an intentionally blank CLI value; do not use an empty string (`""`).
- ♻️ Use `--force` to overwrite an existing single-document output file.

### 📊 Batch issuance from Excel

```bat
wordcloner -i "letter_template.docx" --excel "data.xlsx" --header-row
```

- 📊 Column 1 supplies the output filename; the remaining columns map to unique template variables.
- ⏭️ `--header-row` skips row 1.
- 📁 Use `-o` to override the batch output folder. In Excel mode, `-o` must name a folder, not a `.docx` file.

### 🧾 Options summary

| 🚩 Flag | 📖 Description |
|------|-------------|
| `-i, --input` | 📄 Path to the `.docx` template (required). |
| `-o, --output` | 📁 Output `.docx` file in single mode, or output folder in Excel mode. |
| `-v, --values` | 🔤 Values in first-appearance order of unique variables. |
| `-w, --wrapper` | 🔧 Override the placeholder wrapper for this run. |
| `-x, --excel` | 📊 Excel workbook for batch issuance. |
| `--header-row` | ⏭️ Skip row 1 of the active worksheet. |
| `-c, --count` | 🔢 Print the number of unique variables and exit. |
| `--force` | ♻️ Overwrite an existing single-document output file. |
| `--gui` | 🖥️ Force the GUI to launch. |


## ⚙️ Settings and Data Location

Open **Settings** from the home screen.

| 🛠️ Setting | 📖 Description |
|---------|-------------|
| 🌍 **App Language** | Persian or English. Changing it restarts the GUI. |
| 📁 **Output Path** | Folder where issued documents are written (default: Desktop). |
| 🔧 **Placeholder Wrapper** | Token used to mark variables; default: `##`. |
| 🚀 **Start with Windows** | Windows-only; off by default. When enabled, WordCloner starts hidden in the tray. |
| ⭐ **Quick Access** | Managed from the home screen, not Settings. |
| **Window size/state** | The last window size and maximized state are saved and restored automatically. |

### 📂 Application data

For the current installed Windows build, application data is stored at the user-confirmed location below:

```text
C:\ProgramData\WordCloner\
├── settings.json
├── templates.json
└── Words/
```

This data directory is separate from the executable installation directory and from the output folder selected in Settings. Keep the current installed data path unchanged; uninstalling the application leaves application data in place. This path description is for the installed build; a Python/source run may follow the path logic in `wordcloner_paths.py` or an explicitly configured `WORDCLONER_HOME`.

---

## 🎬 Video Guide

🎥 The Guide button opens a local video using the operating system's default video player.

- 🏠 On the home screen, look for **راهنما / Guide** directly below the **Settings** button near the upper corner of the window.
- 🖱️ The video is selected according to the GUI language:
  - **`Persian_guide.mp4`** for the Persian UI.
  - **`English_guide.mp4`** for the English UI.
- 📁 Place the videos in the `Guide/` folder beside the application files before building or packaging the release.
- The current Guide behavior uses these local MP4 files; it does not depend on `guide.json`.

---

## 🧪 Practice Files

Sample material is provided in the 📁 `examples/` folder:

| 📄 File | 🎯 Purpose |
|------|---------|
| `WordCloner_Test_Template.docx` | English sample template with placeholders. |
| `WordCloner_Test_Template_FA.docx` | Persian sample template with placeholders. |
| `WordCloner_Test_Data.xlsx` | English Excel batch data. |
| `WordCloner_Test_Data_FA.xlsx` | Persian Excel batch data. |

🎯 **Try it now:**

1. 👀 Open a template in Word and inspect its `##variable##` tokens.
2. ➕ Register it in WordCloner using **Add Letter**.
3. ✍️ For a single document, fill in the fields manually.
4. 📊 For batch mode, use the matching `.xlsx` file with **Batch Excel Issue**.

---

## 🗂️ Project Structure

```text
WordCloner.py                  Main GUI and CLI application
WordCloner.spec                PyInstaller build configuration
WordCloner_Setup.iss           Inno Setup installer and uninstall registration
app.ico                        Shared executable, window, Taskbar, and tray icon
pyproject.toml                 Python metadata and dependencies
wordcloner_icons.py            Local SVG icon loader
wordcloner_paths.py            Data path logic and legacy migration helpers
wordcloner_translations.py     Persian and English GUI strings
resources/icons/               Local SVG UI icons
Guide/                         Persian and English local video guides
examples/                      Sample Word templates and Excel workbooks
CLI_guide_Readme.md             Full CLI guide and examples
AI_Assistant_Guide.md           AI support/developer guide and code snapshot
```

---

## 🛠️ Troubleshooting

- **No variables found:** Check that the template contains at least one pair of the configured wrapper, such as `##name##`.
- **Unmatched wrappers:** Ensure every opening wrapper has a closing wrapper; the total number of wrapper tokens must be even.
- **The tray icon is not visible:** Check Windows' **Show hidden icons** area and notification-area settings.
---

## 🤖 AI Assistant Notes

🤖 If you are an AI assistant helping a user or developer with WordCloner, read **`AI-Readme.md`** first. It contains:

- 🗺️ An architectural map of the codebase.
- 🗃️ The data models for `templates.json` and `settings.json`.
- 🔍 The placeholder extraction and replacement pipeline.
- 🔌 Common extension points and build notes.
- ⚠️ Known limitations and troubleshooting guidance.
- 📄 A text snapshot of the codebase for context.

---

## 📜 License

See [`LICENSE`](LICENSE) for the project’s license terms if that file is included in your release. The current workspace does not contain a project-level `LICENSE` file; add the chosen license before redistribution.

---

**📄 WordCloner ✉️** — built to reduce the time spent issuing repetitive letters and circulars in administrative offices. 🏢⏱️🚀
