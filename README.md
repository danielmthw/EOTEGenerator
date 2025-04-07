
# End-of-Term Email Generator

**End-of-Term Email Generator** (EOTEGenerator) is a Windows application designed to streamline the process of generating professional offboarding emails. This Python script uses the `tkinter` library to create a graphical user interface (GUI) for generating standardized End-of-Term (EOT) emails for team members who are offboarding. The app allows users to easily create offboarding emails, including the relevant details about the employee's last day, supervisor's name, and necessary instructions.

The release includes a compiled version of the app, `EOTEGenerator.exe`, that allows team members to generate emails without needing to install Python or any dependencies.

---

## Features

- **EOTEGenerator.exe** — ready to run, no installation required.
- **Auto-generates standardized EOT email templates** for smoother offboarding processes.
- **User-friendly interface** with buttons to:
  - Copy the email content to the clipboard.
  - Send the generated email directly via Outlook.
  - Clear or reset input fields.
- **Safe and self-contained** — no data is stored or transmitted. All email generation happens locally.
- **Calendar widget integration** for selecting the leaver's last working day.
- **Input validation** for proper name formatting.
- **Formatted email preview** directly in the app.

---

## Screenshot

Here’s what the **End-of-Term Email Generator** looks like in action:

![End-of-Term Email Generator](https://lh3.googleusercontent.com/d/1vGYNy1VlFx-Sxo7W7GDp0vDdBR6yx8Wo)
---

## Installation

You have two options to run the End-of-Term Email Generator:

### Option 1: Run as a Python Script

1. Clone the repository:

   ```bash
   git clone https://github.com/danielmthw/EOTEGenerator.git
   cd EOTEGenerator
   ```

2. Install Python dependencies (if needed) and run the script:

   ```bash
   pip install -r requirements.txt
   python EOTEGenerator.py
   ```

### Option 2: Download/Run as .Exe

For Windows users who prefer not to install Python, you can download the standalone executable:

- **EOTEGenerator.exe** — ready to run, no installation required.
- Download the latest version here:  
  [EOTEGenerator.exe - v1.1.0](https://github.com/danielmthw/EOTEGenerator/releases/download/v1.1.0/EOTEGenerator.exe)

---

## Usage

1. **Run the script**: Launch the Python file to open the GUI.
2. **Enter the necessary details**: Input the leaver’s name, supervisor’s name, and last working day.
3. **Generate the email**: Click "Generate Email" to create the email content.
4. **Copy or send the email**: Use the "Copy to Clipboard" button or "Open in Outlook" to easily send the email to the leaver and supervisor.

---

## Changelog

### v1.1.0 (April 2025)

- **Added calendar widget** (`tkcalendar`) for selecting leaver's last day.
- **Implemented input validation** for names (disallowing invalid characters).
- **Improved email formatting** for clarity and consistency.
- **Enhanced error handling** for missing fields and invalid dates.
- **Added real-time date label updates** below the calendar.

## Notes

- Email client integration works with your system’s default mail handler (e.g., Outlook).
- No installation required — just run and go.
- This app is currently available only as a **Windows** executable.

---

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/danielmthw/EOTEGenerator/blob/main/LICENSE) file for details.  
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)


