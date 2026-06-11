import argparse
import base64
import csv
import json
import os
import queue
import subprocess
import sys
import threading
from datetime import datetime

# Tkinter ships with the standard python.org installer, but some minimal,
# embedded or Microsoft Store Python builds omit it. Import defensively so we
# can show a helpful message instead of an opaque traceback (see main()).
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
except ImportError:
    tk = None
    ttk = messagebox = filedialog = None


POWERSHELL_SCRIPT = r'''
$ErrorActionPreference = "Stop"

# Windows PowerShell 5.1 negotiates TLS 1.0 by default, which the PowerShell
# Gallery rejects. Without this, Install-Module fails for any user who does not
# already have ExchangeOnlineManagement installed. Harmless on PowerShell 7+.
try {
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {}

function Send-Status {
    param([string]$Message)
    [Console]::Out.WriteLine("__STATUS__$Message")
    [Console]::Out.Flush()
}

Send-Status "Checking Exchange Online PowerShell module..."

$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement

if (-not $module) {
    Send-Status "ExchangeOnlineManagement not found. Installing module..."

    try {
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Send-Status "Installing NuGet package provider..."
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
        }

        try {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        } catch {}

        Send-Status "Installing ExchangeOnlineManagement module from PSGallery..."
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber

        Send-Status "ExchangeOnlineManagement installed successfully."
    }
    catch {
        throw "Failed to install ExchangeOnlineManagement module: $($_.Exception.Message)"
    }
}

Send-Status "Importing ExchangeOnlineManagement module..."
Import-Module ExchangeOnlineManagement

Send-Status "Opening Microsoft 365 sign-in window..."
Connect-ExchangeOnline -ShowBanner:$false

try {
    Send-Status "Querying Exchange Online for user and shared mailbox list. This can take a few minutes on large tenants..."

    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited `
        -RecipientTypeDetails UserMailbox,SharedMailbox `
        -Properties DisplayName,PrimarySmtpAddress,RecipientTypeDetails,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,GrantSendOnBehalfTo)

    $mailboxCount = $mailboxes.Count

    Send-Status ("Mailbox list loaded. Found {0} mailbox(es)." -f $mailboxCount)

    if ($mailboxCount -eq 0) {
        Send-Status "No mailboxes were returned from Exchange Online."
        [Console]::Out.WriteLine("__JSON__[]")
        [Console]::Out.Flush()
        return
    }

    $results = New-Object System.Collections.Generic.List[object]
    $current = 0

    foreach ($mbx in $mailboxes) {
        $current++

        $mailboxLabel = if ([string]::IsNullOrWhiteSpace([string]$mbx.PrimarySmtpAddress)) {
            [string]$mbx.DisplayName
        } else {
            [string]$mbx.PrimarySmtpAddress
        }

        Send-Status ("Processing mailbox {0} of {1}: {2}" -f $current, $mailboxCount, $mailboxLabel)

        $stats = $null
        $fullAccess = @()
        $sendAs = @()
        $sendOnBehalf = @()

        try {
            Send-Status ("Getting mailbox statistics for {0}" -f $mailboxLabel)
            $stats = Get-EXOMailboxStatistics -Identity $mbx.PrimarySmtpAddress
        }
        catch {
            Send-Status ("Could not get mailbox statistics for {0}" -f $mailboxLabel)
        }

        try {
            Send-Status ("Getting Full Access permissions for {0}" -f $mailboxLabel)
            $fullAccess = Get-MailboxPermission -Identity $mbx.PrimarySmtpAddress |
                Where-Object {
                    $_.User -ne "NT AUTHORITY\SELF" -and
                    $_.IsInherited -eq $false -and
                    $_.Deny -eq $false -and
                    $_.AccessRights -contains "FullAccess"
                } |
                ForEach-Object {
                    [PSCustomObject]@{
                        User = [string]$_.User
                        Permission = "FullAccess"
                    }
                }
        }
        catch {
            Send-Status ("Could not get Full Access permissions for {0}" -f $mailboxLabel)
        }

        try {
            Send-Status ("Getting Send As permissions for {0}" -f $mailboxLabel)
            $sendAs = Get-EXORecipientPermission -Identity $mbx.PrimarySmtpAddress |
                Where-Object {
                    $_.Trustee -ne $null -and
                    $_.AccessRights -contains "SendAs"
                } |
                ForEach-Object {
                    [PSCustomObject]@{
                        User = [string]$_.Trustee
                        Permission = "SendAs"
                    }
                }
        }
        catch {
            Send-Status ("Could not get Send As permissions for {0}" -f $mailboxLabel)
        }

        if ($mbx.GrantSendOnBehalfTo) {
            Send-Status ("Getting Send on Behalf permissions for {0}" -f $mailboxLabel)
            $sendOnBehalf = $mbx.GrantSendOnBehalfTo | ForEach-Object {
                [PSCustomObject]@{
                    User = [string]$_
                    Permission = "SendOnBehalf"
                }
            }
        }

        $allPermissions = @()
        $allPermissions += $fullAccess
        $allPermissions += $sendAs
        $allPermissions += $sendOnBehalf

        $results.Add([PSCustomObject]@{
            DisplayName = [string]$mbx.DisplayName
            PrimarySmtpAddress = [string]$mbx.PrimarySmtpAddress
            MailboxType = [string]$mbx.RecipientTypeDetails
            TotalItemSize = if ($stats) { [string]$stats.TotalItemSize } else { "" }
            ItemCount = if ($stats) { [string]$stats.ItemCount } else { "" }
            IssueWarningQuota = [string]$mbx.IssueWarningQuota
            ProhibitSendQuota = [string]$mbx.ProhibitSendQuota
            ProhibitSendReceiveQuota = [string]$mbx.ProhibitSendReceiveQuota
            Permissions = $allPermissions
        }) | Out-Null
    }

    Send-Status ("Finished processing {0} mailbox(es)." -f $mailboxCount)
    Send-Status "Preparing JSON output..."

    $json = $results.ToArray() | ConvertTo-Json -Depth 8 -Compress
    [Console]::Out.WriteLine("__JSON__$json")
    [Console]::Out.Flush()

    Send-Status "Disconnecting from Exchange Online..."
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
'''


class MailboxReportApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Exchange Online Mailbox Viewer")
        self.root.geometry("1450x850")

        self.mailboxes = []
        self.filtered_mailboxes = []
        self.result_queue = queue.Queue()
        self.is_loading = False

        self.sort_column = None
        self.sort_reverse = False

        self.build_ui()
        self.root.after(150, self.process_queue)

    def build_ui(self):
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill="x")

        self.load_button = ttk.Button(
            top_frame,
            text="Connect and Load Mailboxes",
            command=self.load_mailboxes
        )
        self.load_button.pack(side="left")

        self.export_button = ttk.Button(
            top_frame,
            text="Export CSV",
            command=self.export_csv
        )
        self.export_button.pack(side="left", padx=(10, 0))

        ttk.Label(top_frame, text="Search:").pack(side="left", padx=(20, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.apply_filter())
        ttk.Entry(top_frame, textvariable=self.search_var, width=35).pack(side="left")

        ttk.Label(top_frame, text="Mailbox Type:").pack(side="left", padx=(20, 5))
        self.type_var = tk.StringVar(value="All")
        type_combo = ttk.Combobox(
            top_frame,
            textvariable=self.type_var,
            values=["All", "UserMailbox", "SharedMailbox"],
            state="readonly",
            width=15
        )
        type_combo.pack(side="left")
        type_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

        ttk.Label(top_frame, text="Permission Filter:").pack(side="left", padx=(20, 5))
        self.permission_var = tk.StringVar(value="All")
        permission_combo = ttk.Combobox(
            top_frame,
            textvariable=self.permission_var,
            values=["All", "Any", "FullAccess", "SendAs", "SendOnBehalf", "None"],
            state="readonly",
            width=15
        )
        permission_combo.pack(side="left")
        permission_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

        ttk.Button(top_frame, text="Clear Filters", command=self.clear_filters).pack(side="left", padx=(15, 0))

        status_frame = ttk.Frame(self.root, padding=(10, 0, 10, 8))
        status_frame.pack(fill="x")

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var, anchor="w").pack(fill="x", pady=(0, 6))

        self.progress = ttk.Progressbar(status_frame, mode="indeterminate")
        self.progress.pack(fill="x")

        main_pane = ttk.PanedWindow(self.root, orient="horizontal")
        main_pane.pack(fill="both", expand=True, padx=10, pady=10)

        left_frame = ttk.Frame(main_pane)
        right_frame = ttk.Frame(main_pane)

        main_pane.add(left_frame, weight=4)
        main_pane.add(right_frame, weight=2)

        columns = (
            "DisplayName",
            "PrimarySmtpAddress",
            "MailboxType",
            "TotalItemSize",
            "ItemCount",
        )

        self.tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.pack(side="left", fill="both", expand=True)

        column_settings = [
            ("DisplayName", 240),
            ("PrimarySmtpAddress", 300),
            ("MailboxType", 130),
            ("TotalItemSize", 180),
            ("ItemCount", 110),
        ]

        for col, width in column_settings:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=width, anchor="w")

        y_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        y_scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=y_scroll.set)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_mailbox)

        detail_header = ttk.Label(right_frame, text="Mailbox Details", font=("Segoe UI", 12, "bold"))
        detail_header.pack(anchor="w", pady=(0, 10))

        self.details_text = tk.Text(right_frame, wrap="word", state="disabled")
        self.details_text.pack(fill="both", expand=True)

    def clear_filters(self):
        self.search_var.set("")
        self.type_var.set("All")
        self.permission_var.set("All")
        self.apply_filter()

    def set_loading(self, loading: bool, message: str = ""):
        self.is_loading = loading
        self.load_button.config(state="disabled" if loading else "normal")
        self.export_button.config(state="disabled" if loading else "normal")

        if message:
            self.status_var.set(message)

        if loading:
            self.progress.start(10)
        else:
            self.progress.stop()

    def load_mailboxes(self):
        if self.is_loading:
            return

        self.set_loading(True, "Starting Exchange Online connection...")
        threading.Thread(target=self._load_mailboxes_worker, daemon=True).start()

    def _load_mailboxes_worker(self):
        try:
            data = self.run_powershell_streaming(POWERSHELL_SCRIPT)
            data = self.normalize_mailbox_data(data)
            self.result_queue.put(("success", data))
        except Exception as exc:
            self.result_queue.put(("error", str(exc)))

    def normalize_mailbox_data(self, data):
        if data is None:
            return []

        if isinstance(data, dict):
            data = [data]
        elif not isinstance(data, list):
            data = [data]

        normalized = []
        for item in data:
            if isinstance(item, dict):
                perms = item.get("Permissions", [])
                if isinstance(perms, dict):
                    perms = [perms]
                elif not isinstance(perms, list):
                    perms = []
                item["Permissions"] = [p for p in perms if isinstance(p, dict)]
                normalized.append(item)

        return normalized

    def run_powershell_streaming(self, script: str):
        # Pass the script as a base64 UTF-16LE -EncodedCommand rather than a raw
        # -Command argument. This sidesteps argv quoting/length limits with the
        # large embedded script and bypasses script-file execution policy.
        encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        candidates = [
            ["pwsh", "-NoProfile", "-EncodedCommand", encoded],
            ["powershell", "-NoProfile", "-EncodedCommand", encoded],
        ]

        last_error = None

        for cmd in candidates:
            try:
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    stdin=subprocess.DEVNULL,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=1
                )

                json_payload = None
                output_lines = []

                assert process.stdout is not None
                for raw_line in process.stdout:
                    line = raw_line.strip()
                    if not line:
                        continue

                    output_lines.append(line)

                    if line.startswith("__STATUS__"):
                        self.result_queue.put(("status", line[len("__STATUS__"):]))
                    elif line.startswith("__JSON__"):
                        json_payload = line[len("__JSON__"):]

                process.wait()

                if process.returncode != 0:
                    raise RuntimeError("\n".join(output_lines[-20:]))

                if not json_payload:
                    raise RuntimeError("PowerShell finished but returned no mailbox data.")

                return json.loads(json_payload)

            except FileNotFoundError as exc:
                last_error = exc
            except json.JSONDecodeError as exc:
                raise RuntimeError("Failed to parse mailbox data returned from PowerShell.") from exc

        raise RuntimeError("Could not find PowerShell. Install PowerShell 7 or use Windows PowerShell.") from last_error

    def process_queue(self):
        try:
            while True:
                kind, payload = self.result_queue.get_nowait()

                if kind == "status":
                    self.status_var.set(payload)

                elif kind == "success":
                    self.mailboxes = payload
                    self.apply_filter()
                    self.set_loading(False, f"Loaded {len(self.mailboxes)} mailboxes.")

                elif kind == "error":
                    self.set_loading(False, "Failed to load mailboxes.")
                    messagebox.showerror("Error", payload)

        except queue.Empty:
            pass

        self.root.after(150, self.process_queue)

    def mailbox_matches_permission_filter(self, mailbox, permission_filter):
        permissions = mailbox.get("Permissions", [])
        permission_types = {str(p.get("Permission", "")) for p in permissions if isinstance(p, dict)}

        if permission_filter == "All":
            return True
        if permission_filter == "Any":
            return len(permission_types) > 0
        if permission_filter == "None":
            return len(permission_types) == 0
        return permission_filter in permission_types

    def apply_filter(self):
        search = self.search_var.get().strip().lower()
        mailbox_type = self.type_var.get()
        permission_filter = self.permission_var.get()

        self.filtered_mailboxes = []

        for mbx in self.mailboxes:
            if not isinstance(mbx, dict):
                continue

            if mailbox_type != "All" and mbx.get("MailboxType") != mailbox_type:
                continue

            if not self.mailbox_matches_permission_filter(mbx, permission_filter):
                continue

            haystack_parts = [
                str(mbx.get("DisplayName", "")),
                str(mbx.get("PrimarySmtpAddress", "")),
                str(mbx.get("MailboxType", "")),
                str(mbx.get("TotalItemSize", "")),
                str(mbx.get("ItemCount", "")),
            ]

            permissions = mbx.get("Permissions", [])
            for perm in permissions:
                haystack_parts.append(str(perm.get("User", "")))
                haystack_parts.append(str(perm.get("Permission", "")))

            haystack = " ".join(haystack_parts).lower()

            if search and search not in haystack:
                continue

            self.filtered_mailboxes.append(mbx)

        if self.sort_column:
            self.filtered_mailboxes.sort(
                key=lambda x: self.get_sort_value(x, self.sort_column),
                reverse=self.sort_reverse
            )

        self.refresh_tree()

    def get_sort_value(self, mailbox, column):
        value = mailbox.get(column, "")

        if column == "ItemCount":
            try:
                return int(str(value).replace(",", "").strip())
            except Exception:
                return -1

        return str(value).lower()

    def sort_by_column(self, column):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False

        self.apply_filter()

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for index, mbx in enumerate(self.filtered_mailboxes):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    mbx.get("DisplayName", ""),
                    mbx.get("PrimarySmtpAddress", ""),
                    mbx.get("MailboxType", ""),
                    mbx.get("TotalItemSize", ""),
                    mbx.get("ItemCount", ""),
                )
            )

        self.show_details(None)

    def on_select_mailbox(self, _event):
        selected = self.tree.selection()
        if not selected:
            self.show_details(None)
            return

        idx = int(selected[0])
        if idx < 0 or idx >= len(self.filtered_mailboxes):
            self.show_details(None)
            return

        mailbox = self.filtered_mailboxes[idx]
        self.show_details(mailbox)

    def show_details(self, mailbox):
        self.details_text.config(state="normal")
        self.details_text.delete("1.0", "end")

        if not mailbox:
            self.details_text.insert("end", "Select a mailbox to view details.")
            self.details_text.config(state="disabled")
            return

        lines = [
            f"Display Name: {mailbox.get('DisplayName', '')}",
            f"Email Address: {mailbox.get('PrimarySmtpAddress', '')}",
            f"Mailbox Type: {mailbox.get('MailboxType', '')}",
            f"Total Size: {mailbox.get('TotalItemSize', '')}",
            f"Item Count: {mailbox.get('ItemCount', '')}",
            "",
            "Quota Information",
            "-----------------",
            f"Issue Warning Quota: {mailbox.get('IssueWarningQuota', '')}",
            f"Prohibit Send Quota: {mailbox.get('ProhibitSendQuota', '')}",
            f"Prohibit Send/Receive Quota: {mailbox.get('ProhibitSendReceiveQuota', '')}",
            "",
            "Permissions",
            "-----------",
        ]

        permissions = mailbox.get("Permissions", [])
        if permissions:
            for perm in permissions:
                lines.append(f"{perm.get('User', '')} - {perm.get('Permission', '')}")
        else:
            lines.append("No explicit permissions found.")

        self.details_text.insert("end", "\n".join(lines))
        self.details_text.config(state="disabled")

    def export_csv(self):
        if not self.mailboxes:
            messagebox.showwarning("No Data", "Load mailbox data first.")
            return

        default_name = f"mailbox_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        path = filedialog.asksaveasfilename(
            title="Save CSV Report",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV files", "*.csv")]
        )

        if not path:
            return

        rows = self.flatten_for_csv(self.filtered_mailboxes if self.filtered_mailboxes else self.mailboxes)

        fieldnames = [
            "DisplayName",
            "PrimarySmtpAddress",
            "MailboxType",
            "TotalItemSize",
            "ItemCount",
            "IssueWarningQuota",
            "ProhibitSendQuota",
            "ProhibitSendReceiveQuota",
            "PermissionUser",
            "PermissionType",
        ]

        with open(path, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

        messagebox.showinfo("Export Complete", f"CSV exported to:\n{path}")

    @staticmethod
    def flatten_for_csv(mailboxes):
        rows = []

        for mbx in mailboxes:
            permissions = mbx.get("Permissions", [])

            if not permissions:
                rows.append({
                    "DisplayName": mbx.get("DisplayName", ""),
                    "PrimarySmtpAddress": mbx.get("PrimarySmtpAddress", ""),
                    "MailboxType": mbx.get("MailboxType", ""),
                    "TotalItemSize": mbx.get("TotalItemSize", ""),
                    "ItemCount": mbx.get("ItemCount", ""),
                    "IssueWarningQuota": mbx.get("IssueWarningQuota", ""),
                    "ProhibitSendQuota": mbx.get("ProhibitSendQuota", ""),
                    "ProhibitSendReceiveQuota": mbx.get("ProhibitSendReceiveQuota", ""),
                    "PermissionUser": "",
                    "PermissionType": "",
                })
                continue

            for perm in permissions:
                rows.append({
                    "DisplayName": mbx.get("DisplayName", ""),
                    "PrimarySmtpAddress": mbx.get("PrimarySmtpAddress", ""),
                    "MailboxType": mbx.get("MailboxType", ""),
                    "TotalItemSize": mbx.get("TotalItemSize", ""),
                    "ItemCount": mbx.get("ItemCount", ""),
                    "IssueWarningQuota": mbx.get("IssueWarningQuota", ""),
                    "ProhibitSendQuota": mbx.get("ProhibitSendQuota", ""),
                    "ProhibitSendReceiveQuota": mbx.get("ProhibitSendReceiveQuota", ""),
                    "PermissionUser": perm.get("User", ""),
                    "PermissionType": perm.get("Permission", ""),
                })

        return rows


def build_executable():
    """Compile this script into a standalone Windows .exe using PyInstaller."""
    try:
        import PyInstaller.__main__
    except ImportError:
        print("PyInstaller is required to build the executable.")
        print("Install it first:  pip install pyinstaller")
        return 1

    script_dir = os.path.dirname(os.path.abspath(__file__))
    name = "MailboxReport"

    pyinstaller_args = [
        "--name=%s" % name,
        "--onefile",
        "--windowed",
        "--distpath=%s" % os.path.join(script_dir, "dist"),
        "--workpath=%s" % os.path.join(script_dir, "build"),
        "--specpath=%s" % script_dir,
        os.path.abspath(__file__),
    ]

    PyInstaller.__main__.run(pyinstaller_args)
    print("\nBuild complete. Executable: %s" % os.path.join(script_dir, "dist", name + ".exe"))
    return 0


def main():
    parser = argparse.ArgumentParser(description="Exchange Online Mailbox Viewer")
    parser.add_argument(
        "--build",
        action="store_true",
        help="Build a standalone Windows executable with PyInstaller, then exit.",
    )
    args = parser.parse_args()

    if args.build:
        return build_executable()

    if tk is None:
        sys.stderr.write(
            "This application needs Python's Tkinter GUI library, which is not installed.\n"
            "Tkinter is included with the standard installer from https://python.org\n"
            "(make sure the 'tcl/tk and IDLE' feature is selected). If you cannot install\n"
            "it here, build a standalone .exe on a machine that has it and run that:\n"
            "    python MailboxReport.py --build\n"
        )
        return 1

    root = tk.Tk()
    style = ttk.Style(root)

    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    MailboxReportApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())