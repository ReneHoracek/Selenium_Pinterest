# Pinterest Smart Automation V4

Advanced Pinterest automation tool built with Python and Selenium.

## Features

- Automatic Pinterest login
- Video pin uploading
- Scheduled publishing
- Smart board rotation
- Smart color balancing algorithm
- Excel-based content management
- Cross-platform logging
- Retry/recovery system for unstable Pinterest UI
- Secure credential storage using keyring
- Automatic tag handling
- Video upload monitoring

---

## Technologies

- Python 3.10+
- Selenium
- Pandas
- OpenPyXL
- Keyring

---

## Installation

Clone repository:

```bash
git clone https://github.com/YOUR_USERNAME/pinterest-smart-automation.git
cd pinterest-smart-automation
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Required Packages

```bash
pip install selenium pandas openpyxl keyring
```

Google Chrome is required.

Also install matching ChromeDriver:
https://googlechromelabs.github.io/chrome-for-testing/

---

## Usage

Place video files into the project folder.

Supported formats:

- .mp4
- .mov
- .avi
- .mkv
- .webm

Run script:

```bash
python pinterest_smart_v7.py
```

---

## Workflow

### First Run

The script:

1. Scans videos
2. Asks for clothing colors
3. Creates formatted Excel planner
4. Generates schedule automatically

### Second Run

After editing Excel:

- fill titles
- fill tags
- set status to:

```text
PRIPRAVENO K NAHRANI
```

Run script again to upload pins automatically.

---

## Security

Credentials are stored securely using:

- Windows Credential Manager
- GNOME Keyring
- Linux Secret Service

Passwords are NOT stored directly inside source code unless manually added to fallback variables.

---

## Important

Before publishing to GitHub:

- remove personal credentials
- keep fallback credentials empty
- do not upload logs/videos/excel exports

---

## Disclaimer

This project is for educational and productivity purposes.

Users are responsible for complying with Pinterest Terms of Service.

---

## License

MIT License
