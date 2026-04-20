# msgtoeml
A Python tool for converting Microsoft Outlook `.msg` files to `.eml` files on Linux without requiring Outlook.

Designed for environments without Outlook (e.g. Ubuntu) and built using open‑source libraries only. Ideal for reviewing emails with tools like Thuderbird.


## Features

- Converts `.msg` to `.eml`
- Preserves core headers (From, To, Cc, Subject, Date)
- Supports plain text and HTML bodies
- Handles `bytes` vs `str` issues from `extract-msg`
- No Outlook, Wine, or Windows required


## Requirements

- Python 3.8+
- `extract-msg`

Install dependency:

```bash
pip install extract-msg

## Usage
```bash
python3 msgtoeml.py input.msg output.eml

