# Outlook Template Bartender

**Effortlessly Manage and Generate Outlook Emails with Custom Templates and Dynamic Variables.**

---

## 🍷 Introduction

Still battling with an aging office PC? Does opening Outlook feel like wading through molasses because of the endless flood of emails?

For many workers stuck with company-issued, outdated laptops or desktops, **Outlook freezes** and sluggish responses are an everyday frustration. Even something as simple as replying with a saved template turns into an ordeal — opening the file explorer, searching for the right template, carefully editing, worrying about CC mistakes — it all piles up.

**Outlook Template Bartender** is built exactly for this:

> 🖥️ **Lightweight, offline, no cloud or fancy AI gimmicks.**
>
> 🍹 **Quickly manage, organize, and reuse email templates without the lag.**
>
> 📨 **Avoid opening extra windows, misclicking, or CC-ing the wrong person ever again.**

Perfect for:

- Office warriors fighting old hardware & bloated inboxes daily
- Those who just need an efficient, simple template system without the bells & whistles
- Anyone fed up with Outlook’s sluggishness when doing repetitive replies


---

## ✨ Features

- 📝 **Manage Templates Easily**: Organize email templates by event types (categories).
- 🖼️ **Quick Copy & Edit Existing Emails**:
  - Copy an email from Outlook, paste directly into our editor.
  - Automatically preserves most of the **original HTML formatting**.
  - Modify, insert variables, save instantly.
- 🔄 **Dynamic Variables**: Insert `{receiver}`, `{date}`, `{project_name}` placeholders and auto-fill them during email generation.
- 📧 **Direct Outlook Integration**: Generate and open pre-filled emails in Outlook.
- 👤 **Multiple Account Support**: Choose your sender account easily.
- 🌐 **Multilingual Interface**: English & Traditional Chinese.
- 🔍 **Search & Filter**: Quickly find templates by name or content.
- 🗂️ **Template Organization**: Event types + tags for better classification.
- 📁 **Local Storage, No Cloud**: Secure SQLite database; no privacy concerns.
- 💻 **Lightweight & Offline**: Works fully offline, no login required.

---

## 📸 Screenshots

*Coming Soon...*

---

## 🚀 Installation

### Option 1: From Source (Recommended)

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Ash0Lam/OutlookTemplateBartender.git
   cd OutlookTemplateBartender
   ```

2. **(Recommended) Create Virtual Environment:**
   ```bash
   python -m venv venv
   ```

3. **Activate Virtual Environment:**

   - **On Windows:**
     ```bash
     venv\Scripts\activate
     ```
   - **On macOS/Linux:**
     ```bash
     source venv/bin/activate
     ```

4. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

5. **Run the application:**
   ```bash
   python main_db.py
   ```

### Option 2: Standalone Executable *(Coming Soon)*

---

## 📚 Usage

### Managing Templates

- **Add Event Types**: Categorize templates.
- **Add Templates**: Create email templates with dynamic variables.
- **Edit Templates**: Double-click or use right-click menu.
- **Quick Copy Existing Email**: Paste any Outlook email into our editor; most HTML formatting is preserved.
- **Search Templates**: Quickly find templates by name/content.

### Generating Emails

1. Select event type + template.
2. Fill in variable values.
3. Click **Generate Email**.
4. Outlook opens the email ready-to-send, with images & formatting intact.

---

## 🏗️ Project Structure

```
OutlookTemplateBartender/
├── db_manager.py        # Database operations
├── template_manager.py  # Template management logic
├── language_manager.py  # Multilingual support
├── email_generator.py   # Outlook email creation logic
├── image_manager.py     # Image handling
├── gui/
│   ├── main_window.py   # Main app interface
│   └── edit_template.py # Template editing window
├── assets/              # Icons & static assets
├── README.md
└── requirements.txt
```

---

## 📦 Requirements

- Windows OS
- Microsoft Outlook (Installed & Configured)
- Python 3.6+
- Required Libraries:

```
bottle==0.13.2
certifi==2025.1.31
cffi==1.17.1
charset-normalizer==3.4.1
clr_loader==0.2.7.post0
idna==3.10
pillow==10.4.0
proxy_tools==0.1.0
psutil==7.0.0
pycparser==2.22
pythonnet==3.0.5
pywebview==5.4
pywin32==310
requests==2.32.3
tkhtmlview==0.3.1
typing_extensions==4.12.2
urllib3==2.3.0
```

---

## 🌟 Why Outlook Template Bartender?

Your daily office tasks might not involve heavy machinery, but constant misclicks, sluggish Outlook responses, and CC mishaps drain energy all the same. Many tools exist – but they are either too complex, require cloud subscriptions, or lack flexibility.

This tool is **simple, fast, offline, privacy-respecting** – and leaves no mess.

---

## 📜 License

MIT License.

---

## 🙌 Credits

- Developed by [Ash](https://github.com/Ash0Lam)
- CKEditor for rich-text editing.
- Python win32com for Outlook integration.

---

## 📬 Contact

- GitHub: [https://github.com/Ash0Lam](https://github.com/Ash0Lam)
- Website: [ash0lam.github.io](https://ash0lam.github.io)

