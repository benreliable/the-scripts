# Mailbox Report

A desktop (Tkinter) viewer for Exchange Online mailboxes. It connects to your
tenant, lists user and shared mailboxes with their size, quotas and delegated
permissions (Full Access, Send As, Send on Behalf), lets you search/filter/sort,
and exports the result to CSV.

## What gets installed automatically

On first run the app checks for the `ExchangeOnlineManagement` PowerShell module
and installs it (current user scope) from the PowerShell Gallery if it is
missing — so technicians do not need to set anything up beforehand. The script
forces TLS 1.2 first, which is required for the Gallery to work on Windows
PowerShell 5.1.

You still need either **Windows PowerShell 5.1** (built into Windows) or
**PowerShell 7+** on the machine; the app uses whichever it finds.

## Running from source

1.  **Install Python** from [python.org](https://python.org). Use the standard
    installer and keep the *"tcl/tk and IDLE"* feature selected — that provides
    Tkinter, the GUI library this app needs. (If Tkinter is missing the app
    prints instructions instead of crashing.)
2.  Run it:
    ```bash
    python MailboxReport.py
    ```
3.  Click **Connect and Load Mailboxes** and sign in to Microsoft 365 when the
    browser window opens.

## Building a standalone .exe

For machines without Python, build a single-file executable. Run this on a
machine that *does* have Python + Tkinter:

```bash
pip install pyinstaller
python MailboxReport.py --build
```

The executable is written to `dist/MailboxReport.exe`. It still requires
PowerShell on the target machine (present on all supported Windows versions),
and will install the Exchange module on first run as described above.
