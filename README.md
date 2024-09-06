# SAP Automation Script

This project provides a class for automating interactions with SAP GUI using the SAP GUI Scripting API and the `win32` library. The `SapGui` class handles login, password change, logout, and other SAP GUI operations.

## Features

- Automates SAP login using provided credentials.
- Handles SAP password change prompts.
- Allows custom operations, such as waiting for specific elements or dialogs.
- Supports closing sessions and logging out of SAP.
- Retrieves and handles SAP element text.

## Installation

To use this script, make sure the following dependencies are installed:

```bash
pip install pywin32 pygetwindow
```

## Usage
Initializing the SAP Automation

To create an SAP session:

```python

from sap_gui import SapGui

sap_args = {
    "platform": "your_sap_system",
    "username": "your_username",
    "password": "your_password",
    "sap_client": "100",
    "sap_language": "EN",
    "sap_path": "C:\\Program Files\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
}

sap_session = SapGui(sap_args)

# Log in to SAP
if sap_session.sapLogin():
    print("Login successful!")
else:
    print("Login failed.")
```

## Handling SAP Password Change

The SapGui class automatically detects password change prompts and generates a new password in the format Month#Year (e.g., Maio#2024).
Logout

To safely log out of SAP:

```python
sap_session.sapLogout()
```

## Close SAP Session

To close the session:

```python
sap_session.close_connection()
```

## Requirements

- Python 3.6+
- SAP GUI Client installed
- SAP Scripting enabled
