"""
M365 PST Archiver
Export Exchange Online mailboxes to PST format via Outlook COM

Usage:
    python m365_pst_archiver.py              # Launch GUI
    python m365_pst_archiver.py --build-exe  # Build standalone EXE
"""

import sys
import subprocess
import logging
import argparse
import shutil
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QProgressBar,
    QFileDialog, QMessageBox, QGroupBox, QStatusBar, QFrame,
    QListWidget, QListWidgetItem, QAbstractItemView,
)
from PySide6.QtCore import QThread, Signal, Qt, QTimer
from PySide6.QtGui import QFont

APP_VERSION = "1.3.0"

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    handlers=[logging.FileHandler('pst_archiver.log'), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
#  EXE builder
# ──────────────────────────────────────────────
def build_executable():
    print("\n" + "=" * 60)
    print("M365 PST Archiver - EXE Build")
    print("=" * 60 + "\n")
    try:
        import PyInstaller
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)

    script_path = Path(__file__).resolve()
    dist_dir, build_dir = Path("dist"), Path("build")
    for d in [dist_dir, build_dir, Path("M365_PST_Archiver.spec"), Path("__pycache__")]:
        if d.exists():
            try:
                shutil.rmtree(d) if d.is_dir() else d.unlink()
            except Exception as e:
                print(f"Could not remove {d}: {e}")
    try:
        subprocess.run([
            sys.executable, "-m", "PyInstaller",
            "--onefile", "--windowed",
            "--name", "M365_PST_Archiver",
            "--collect-all=PySide6",
            "--distpath", str(dist_dir),
            "--buildpath", str(build_dir),
            str(script_path)
        ], check=True)
        exe = dist_dir / "M365_PST_Archiver.exe"
        if exe.exists():
            print(f"Build successful: {exe} ({exe.stat().st_size/1024/1024:.1f} MB)")
            return 0
        return 1
    except Exception as e:
        print(f"Build failed: {e}")
        return 1


# ──────────────────────────────────────────────
#  Stylesheet
# ──────────────────────────────────────────────
APP_STYLE = """
    QMainWindow, QWidget, QDialog { font-size: 11px; }
    QGroupBox {
        font-weight: 600;
        border: 1px solid #cfcfcf;
        border-radius: 8px;
        margin-top: 10px;
        padding-top: 10px;
        background: #fafafa;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 4px 0 4px;
    }
    QLineEdit, QTextEdit {
        border: 1px solid #c8c8c8;
        border-radius: 6px;
        padding: 6px;
        background: white;
    }
    QListWidget {
        border: 1px solid #c8c8c8;
        border-radius: 6px;
        background: white;
    }
    QPushButton {
        padding: 4px 12px;
        border-radius: 6px;
        border: 1px solid #c0c0c0;
        background: #f4f4f4;
        min-height: 26px;
    }
    QPushButton:hover    { background: #ececec; }
    QPushButton:disabled { color: #888; background: #f0f0f0; }
    QPushButton#PrimaryButton {
        background: #1f6feb; color: white;
        border: 1px solid #1f6feb; font-weight: 600;
    }
    QPushButton#PrimaryButton:hover    { background: #1857b8; }
    QPushButton#PrimaryButton:disabled { background: #a0bbdd; border-color: #a0bbdd; }
    QPushButton#DangerButton {
        background: #d73a49; color: white;
        border: 1px solid #d73a49; font-weight: 600;
    }
    QPushButton#DangerButton:hover    { background: #b92f3d; }
    QPushButton#DangerButton:disabled { background: #e8a0a6; border-color: #e8a0a6; }
    QPushButton#SecondaryButton {
        background: #2da44e; color: white;
        border: 1px solid #2da44e; font-weight: 600;
    }
    QPushButton#SecondaryButton:hover { background: #238636; }
    QProgressBar {
        border: 1px solid #c8c8c8; border-radius: 6px;
        text-align: center; min-height: 20px; background: white;
    }
    QProgressBar::chunk { background: #1f6feb; border-radius: 5px; }
    QFrame#Banner   { border: 1px solid #d8d8d8; border-radius: 8px; background: #f7faff; }
    QFrame#InfoCard { border: 1px solid #d8d8d8; border-radius: 8px; background: white; }
    QLabel#CardTitle { color: #666; font-size: 10px; font-weight: 600; }
    QLabel#CardValue { font-size: 14px; font-weight: 600; }
    QLabel#SetupHeader    { font-size: 18px; font-weight: 700; }
    QLabel#SetupSubheader { color: #555; }
    QLabel#SetupStatus    { font-size: 14px; font-weight: 700; }
"""


# ──────────────────────────────────────────────
#  InfoCard
# ──────────────────────────────────────────────
class InfoCard(QFrame):
    def __init__(self, title: str, value: str = "-"):
        super().__init__()
        self.setObjectName("InfoCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("CardTitle")
        self.value_label = QLabel(value)
        self.value_label.setObjectName("CardValue")
        self.value_label.setWordWrap(True)
        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value: str):
        self.value_label.setText(value)


# ──────────────────────────────────────────────
#  Mailbox loader thread
# ──────────────────────────────────────────────
class MailboxLoaderWorker(QThread):
    loaded = Signal(list)
    error  = Signal(str)

    def __init__(self, admin_email: str):
        super().__init__()
        self.admin_email = admin_email

    def run(self):
        ps = (
            "try {"
            "  $mbx = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop"
            "    | Select-Object -ExpandProperty PrimarySmtpAddress;"
            "  $mbx | ForEach-Object { Write-Output $_ }"
            "} catch {"
            "  Write-Output ('ERROR:' + $_.Exception.Message)"
            "}"
        )
        try:
            result = subprocess.run(
                ["powershell.exe", "-ExecutionPolicy", "Bypass", "-Command", ps],
                capture_output=True, text=True, timeout=180
            )
            lines  = [l.strip() for l in result.stdout.splitlines() if l.strip()]
            errors = [l for l in lines if l.upper().startswith("ERROR:")]
            if errors:
                self.error.emit(errors[0].replace("ERROR:", "").replace("error:", ""))
                return
            mailboxes = sorted([l for l in lines if "@" in l])
            if not mailboxes:
                self.error.emit("No mailboxes returned. Ensure you are connected to Exchange Online.")
                return
            self.loaded.emit(mailboxes)
        except subprocess.TimeoutExpired:
            self.error.emit("Timed out fetching mailboxes (3 min limit).")
        except Exception as e:
            self.error.emit(str(e))


# ──────────────────────────────────────────────
#  Connection / auth dialog
# ──────────────────────────────────────────────
class ConnectionSetupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("M365 PST Archiver — Connect")
        self.resize(540, 320)
        self.admin_email: str = ""
        self._build_ui()
        self.setStyleSheet(APP_STYLE)

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        header = QLabel(f"M365 PST Archiver  v{APP_VERSION}")
        header.setObjectName("SetupHeader")
        sub = QLabel(
            "Connect to Exchange Online before exporting mailboxes.\n"
            "Enter your Admin UPN and authenticate interactively."
        )
        sub.setWordWrap(True)
        sub.setObjectName("SetupSubheader")
        layout.addWidget(header)
        layout.addWidget(sub)

        form_group = QGroupBox("Exchange Online Connection")
        form = QFormLayout(form_group)
        self.upn_input = QLineEdit()
        self.upn_input.setPlaceholderText("admin@company.com")
        form.addRow("Admin UPN", self.upn_input)
        layout.addWidget(form_group)

        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout(status_group)
        self.status_label = QLabel("Not connected")
        self.status_label.setObjectName("SetupStatus")
        self.detail_label = QLabel("Authenticate to continue")
        self.detail_label.setWordWrap(True)
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.detail_label)
        layout.addWidget(status_group)

        buttons = QHBoxLayout()
        self.auth_btn = QPushButton("Connect to Exchange Online")
        self.auth_btn.setObjectName("SecondaryButton")
        self.auth_btn.clicked.connect(self._authenticate)

        self.continue_btn = QPushButton("Open PST Archiver")
        self.continue_btn.setObjectName("PrimaryButton")
        self.continue_btn.setEnabled(False)
        self.continue_btn.clicked.connect(self.accept)

        self.exit_btn = QPushButton("Exit")
        self.exit_btn.clicked.connect(self.reject)

        buttons.addWidget(self.auth_btn)
        buttons.addStretch(1)
        buttons.addWidget(self.exit_btn)
        buttons.addWidget(self.continue_btn)
        layout.addLayout(buttons)

    def _authenticate(self):
        email = self.upn_input.text().strip()
        if not email:
            QMessageBox.warning(self, "Missing UPN", "Enter your Admin UPN before connecting.")
            return

        self.auth_btn.setEnabled(False)
        self.status_label.setText("Connecting...")
        self.detail_label.setText(
            "Running Exchange Online connection — a browser or PowerShell prompt may appear."
        )
        QApplication.processEvents()

        ps = (
            "try {"
            f" Import-Module ExchangeOnlineManagement -ErrorAction Stop;"
            f" Connect-ExchangeOnline -UserPrincipalName '{email}'"
            "  -ShowBanner:$false -ErrorAction Stop;"
            " Write-Output 'CONNECTED'"
            "} catch {"
            " Write-Output ('ERROR:' + $_.Exception.Message)"
            "}"
        )
        try:
            result = subprocess.run(
                ["powershell.exe", "-ExecutionPolicy", "Bypass", "-Command", ps],
                capture_output=True, text=True, timeout=120
            )
            out = result.stdout.strip()
            if "CONNECTED" in out:
                self.admin_email = email
                self.status_label.setText(f"Connected as {email}")
                self.detail_label.setText(
                    "Authentication successful. You may now open the archiver."
                )
                self.continue_btn.setEnabled(True)
                self.auth_btn.setEnabled(False)
            else:
                err = out.replace("ERROR:", "") or result.stderr.strip() or "Unknown error"
                self.status_label.setText("Connection failed")
                self.detail_label.setText(err)
                self.auth_btn.setEnabled(True)
                QMessageBox.critical(
                    self, "Connection Failed",
                    f"Could not connect to Exchange Online:\n\n{err}\n\n"
                    "Ensure the ExchangeOnlineManagement module is installed."
                )
        except subprocess.TimeoutExpired:
            self.status_label.setText("Timed out")
            self.detail_label.setText("Connection attempt exceeded 2 minutes.")
            self.auth_btn.setEnabled(True)
        except Exception as e:
            self.status_label.setText("Error")
            self.detail_label.setText(str(e))
            self.auth_btn.setEnabled(True)


# ──────────────────────────────────────────────
#  Archive worker
# ──────────────────────────────────────────────
class ArchiveWorker(QThread):
    progress  = Signal(str, int, int)
    log       = Signal(str)
    completed = Signal(str)
    error     = Signal(str)
    cancelled = Signal()

    def __init__(self, admin_email: str, mailboxes: list, destination: str):
        super().__init__()
        self.admin_email = admin_email
        self.mailboxes   = mailboxes
        self.destination = destination
        self._cancel     = False
        self._process    = None

    def cancel(self):
        self._cancel = True
        if self._process:
            self._process.terminate()

    def run(self):
        try:
            ps_script = self._generate_ps()
            ps_file   = Path(self.destination) / "_pst_archive_script.ps1"
            ps_file.write_text(ps_script, encoding="utf-8")
            self.log.emit(f"PS     | SCRIPT  | {ps_file}")
            self.progress.emit("Running", 0, len(self.mailboxes))

            self._process = subprocess.Popen(
                ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", str(ps_file)],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, encoding="utf-8", errors="replace"
            )

            for line in self._process.stdout:
                line = line.rstrip()
                if not line:
                    continue
                if line.startswith("PROGRESS:"):
                    parts = line.split(":")
                    if len(parts) >= 3:
                        self.progress.emit("Exporting", int(parts[1]), int(parts[2]))
                elif line.startswith("MAILBOX:"):
                    self.progress.emit(f"Processing: {line[8:]}", 0, 0)
                else:
                    self.log.emit(line)
                if self._cancel:
                    self._process.terminate()
                    break

            self._process.wait()
            for err in self._process.stderr:
                err = err.rstrip()
                if err:
                    self.log.emit(f"[WARN] {err}")

            if self._cancel:
                self.cancelled.emit()
            elif self._process.returncode == 0:
                self.completed.emit(self.destination)
            else:
                self.error.emit(f"Process exited with code {self._process.returncode}")

        except Exception as e:
            self.error.emit(str(e))

    def _generate_ps(self):
        """
        Build the PowerShell script as a list of plain string lines.

        KEY: The filename-sanitise regex must never contain a literal double-quote
        because that char terminates PowerShell single-quoted strings on some
        parser versions. Strategy: two -replace passes —
          Pass 1: strip  < > : / backslash | ? *   (no quote needed)
          Pass 2: strip the double-quote via [char]34  (no literal " in script)
        """
        mailbox_list = "', '".join(self.mailboxes)
        total = len(self.mailboxes)
        ts    = datetime.now().isoformat()

        L = []
        L.append(f"# M365 PST Archiver Script -- {ts}")
        L.append(f"$AdminEmail      = '{self.admin_email}'")
        L.append(f"$TargetMailboxes = @('{mailbox_list}')")
        L.append(f"$DestinationPath = '{self.destination}'")
        L.append(f"$Total           = {total}")
        L.append( "$Done            = 0")
        L.append("")
        L.append("if (-not (Test-Path $DestinationPath)) {")
        L.append("    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null")
        L.append("}")
        L.append("")
        L.append("function Copy-OlkFolderRecurse {")
        L.append("    param ($SourceFolder, $DestinationParent)")
        L.append("    try {")
        L.append("        foreach ($SubFolder in $SourceFolder.Folders) {")
        L.append("            $NewFolder = $SubFolder.CopyTo($DestinationParent)")
        L.append("            Copy-OlkFolderRecurse -SourceFolder $SubFolder -DestinationParent $NewFolder")
        L.append("        }")
        L.append("    } catch {")
        L.append("        Write-Host ('  Skipped folder: ' + $SourceFolder.Name)")
        L.append("    }")
        L.append("}")
        L.append("")
        L.append("Write-Host 'Initializing Outlook COM...'")
        L.append("try {")
        L.append("    $Outlook   = New-Object -ComObject Outlook.Application")
        L.append("    $Namespace = $Outlook.GetNamespace('MAPI')")
        L.append("} catch {")
        L.append("    Write-Host 'ERROR: Failed to initialize Outlook -- ensure Outlook is installed.'")
        L.append("    exit 1")
        L.append("}")
        L.append("")
        L.append("foreach ($Mailbox in $TargetMailboxes) {")
        L.append("    Write-Host ('MAILBOX:' + $Mailbox)")
        L.append("    Write-Host ('Processing: ' + $Mailbox)")
        L.append("    try {")
        L.append("        Write-Host '  Granting FullAccess...'")
        L.append("        Add-MailboxPermission -Identity $Mailbox -User $AdminEmail `")
        L.append("            -AccessRights FullAccess -AutoMapping $true `")
        L.append("            -Confirm:$false -ErrorAction SilentlyContinue")
        L.append("")
        L.append("        Write-Host '  Waiting for replication (45s)...'")
        L.append("        Start-Sleep -Seconds 45")
        L.append("")
        L.append("        $SourceStore = $Namespace.Stores | Where-Object { $_.DisplayName -eq $Mailbox }")
        L.append("        if ($SourceStore) {")
        # Two-pass sanitise — NO literal double-quote anywhere in this script
        L.append("            $SafeName = $Mailbox -replace '[<>:/\\\\|?*]', '_'")
        L.append("            $SafeName = $SafeName -replace [char]34, '_'")
        L.append("            $PSTPath  = Join-Path $DestinationPath ($SafeName + '.pst')")
        L.append("            Write-Host ('  Exporting to: ' + $PSTPath)")
        L.append("            $Namespace.AddStore($PSTPath)")
        L.append("            $NewPST   = $Namespace.Stores | Where-Object { $_.FilePath -eq $PSTPath }")
        L.append("            $DestRoot = $NewPST.GetRootFolder()")
        L.append("            foreach ($Folder in $SourceStore.GetRootFolder().Folders) {")
        L.append("                Copy-OlkFolderRecurse -SourceFolder $Folder -DestinationParent $DestRoot")
        L.append("            }")
        L.append("            $Namespace.RemoveStore($DestRoot)")
        L.append("            Write-Host ('  Done: ' + $PSTPath)")
        L.append("        } else {")
        L.append("            Write-Host '  ERROR: Mailbox not found in Outlook profile.'")
        L.append("        }")
        L.append("    } catch {")
        L.append("        Write-Host ('  ERROR: ' + $_)")
        L.append("    } finally {")
        L.append("        Remove-MailboxPermission -Identity $Mailbox -User $AdminEmail `")
        L.append("            -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue")
        L.append("    }")
        L.append("    $Done++")
        L.append("    Write-Host ('PROGRESS:' + $Done + ':' + $Total)")
        L.append("}")
        L.append("")
        L.append("Disconnect-ExchangeOnline -Confirm:$false")
        L.append("Write-Host 'All done.'")

        return "\n".join(L)


# ──────────────────────────────────────────────
#  Main window
# ──────────────────────────────────────────────
class MainWindow(QMainWindow):
    log_signal = Signal(str)

    def __init__(self, admin_email: str):
        super().__init__()
        self.setWindowTitle(f"M365 PST Archiver  v{APP_VERSION}")
        self.resize(860, 920)

        self.admin_email   = admin_email
        self.worker        = None
        self._loader       = None
        self._start_time   = None
        self._all_mailboxes: list = []

        self.log_signal.connect(self._append_log)
        self._build_ui()
        self.setStyleSheet(APP_STYLE)
        self._log("INFO", "APP", "START", "Application started")

    # ── UI ────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # Banner
        banner = QFrame()
        banner.setObjectName("Banner")
        bg = QGridLayout(banner)
        bg.setContentsMargins(12, 10, 12, 10)
        self.banner_user   = QLabel(f"User: {self.admin_email}")
        domain = self.admin_email.split("@")[-1] if "@" in self.admin_email else "-"
        self.banner_tenant = QLabel(f"Tenant: {domain}")
        self.banner_sel    = QLabel("Selected: 0")
        self.banner_status = QLabel("Status: Ready")
        bg.addWidget(self.banner_user,   0, 0)
        bg.addWidget(self.banner_tenant, 0, 1, Qt.AlignRight)
        bg.addWidget(self.banner_sel,    1, 0)
        bg.addWidget(self.banner_status, 1, 1, Qt.AlignRight)
        root.addWidget(banner)

        # Mailbox picker
        picker = QGroupBox("Mailboxes")
        pl = QVBoxLayout(picker)

        top_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter mailboxes...")
        self.search_input.textChanged.connect(self._filter_mailboxes)
        self.load_btn = QPushButton("Load Mailboxes")
        self.load_btn.setObjectName("SecondaryButton")
        self.load_btn.setFixedWidth(140)
        self.load_btn.clicked.connect(self._load_mailboxes)
        top_row.addWidget(self.search_input, 1)
        top_row.addWidget(self.load_btn)
        pl.addLayout(top_row)

        self.mailbox_list = QListWidget()
        self.mailbox_list.setFixedHeight(210)
        self.mailbox_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.mailbox_list.itemChanged.connect(self._on_item_changed)
        placeholder = QListWidgetItem("Click 'Load Mailboxes' to fetch from Exchange Online")
        placeholder.setFlags(Qt.NoItemFlags)
        self.mailbox_list.addItem(placeholder)
        pl.addWidget(self.mailbox_list)

        sel_row = QHBoxLayout()
        self.sel_all_btn  = QPushButton("Select All")
        self.sel_none_btn = QPushButton("Select None")
        self.sel_all_btn.setEnabled(False)
        self.sel_none_btn.setEnabled(False)
        self.sel_all_btn.clicked.connect(self._select_all)
        self.sel_none_btn.clicked.connect(self._select_none)
        self.sel_count_label = QLabel("0 selected")
        sel_row.addWidget(self.sel_all_btn)
        sel_row.addWidget(self.sel_none_btn)
        sel_row.addStretch(1)
        sel_row.addWidget(self.sel_count_label)
        pl.addLayout(sel_row)
        root.addWidget(picker)

        # Output folder
        out_group = QGroupBox("Output")
        out_form  = QFormLayout(out_group)
        out_row   = QHBoxLayout()
        self.output_input = QLineEdit(str(Path.home() / "PSTExports"))
        self.browse_btn   = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self._browse_output)
        out_row.addWidget(self.output_input, 1)
        out_row.addWidget(self.browse_btn)
        out_form.addRow("Output Folder", out_row)
        root.addWidget(out_group)

        # Progress
        prog_group  = QGroupBox("Progress")
        prog_layout = QVBoxLayout(prog_group)
        cards_row = QHBoxLayout()
        self.card_phase    = InfoCard("Phase",     "Idle")
        self.card_progress = InfoCard("Progress",  "0 / 0")
        self.card_current  = InfoCard("Current",   "-")
        self.card_duration = InfoCard("Duration",  "00:00")
        self.card_total    = InfoCard("Mailboxes", "0")
        for c in (self.card_phase, self.card_progress, self.card_current,
                  self.card_duration, self.card_total):
            cards_row.addWidget(c)
        prog_layout.addLayout(cards_row)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        prog_layout.addWidget(self.progress_bar)
        self.progress_detail = QLabel("Ready")
        prog_layout.addWidget(self.progress_detail)
        root.addWidget(prog_group)

        # Live log
        log_group  = QGroupBox("Live Log")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        root.addWidget(log_group, 1)

        # Buttons
        actions = QHBoxLayout()
        self.start_btn = QPushButton("Start Archive")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.setFixedWidth(120)
        self.start_btn.clicked.connect(self._start_archive)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setObjectName("DangerButton")
        self.cancel_btn.setFixedWidth(90)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_archive)

        self.open_btn = QPushButton("Open Output Folder")
        self.open_btn.setFixedWidth(150)
        self.open_btn.clicked.connect(self._open_output)

        actions.addWidget(self.start_btn)
        actions.addWidget(self.cancel_btn)
        actions.addStretch(1)
        actions.addWidget(self.open_btn)
        root.addLayout(actions)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

    # ── Mailbox loading ───────────────────────
    def _load_mailboxes(self):
        self.load_btn.setEnabled(False)
        self.sel_all_btn.setEnabled(False)
        self.sel_none_btn.setEnabled(False)
        self.mailbox_list.blockSignals(True)
        self.mailbox_list.clear()
        item = QListWidgetItem("Loading mailboxes from Exchange Online...")
        item.setFlags(Qt.NoItemFlags)
        self.mailbox_list.addItem(item)
        self.mailbox_list.blockSignals(False)
        self._log("INFO", "EXO", "FETCH", "Fetching mailbox list...")
        self.statusBar().showMessage("Loading mailboxes...")

        self._loader = MailboxLoaderWorker(self.admin_email)
        self._loader.loaded.connect(self._on_mailboxes_loaded)
        self._loader.error.connect(self._on_mailboxes_error)
        self._loader.start()

    def _on_mailboxes_loaded(self, mailboxes: list):
        self._all_mailboxes = mailboxes
        self._populate_list(mailboxes)
        self.load_btn.setEnabled(True)
        self.sel_all_btn.setEnabled(True)
        self.sel_none_btn.setEnabled(True)
        self._log("INFO", "EXO", "FETCH", f"Loaded {len(mailboxes)} mailbox(es)")
        self.statusBar().showMessage(f"Loaded {len(mailboxes)} mailboxes")

    def _on_mailboxes_error(self, msg: str):
        self.mailbox_list.blockSignals(True)
        self.mailbox_list.clear()
        err_item = QListWidgetItem(f"Error: {msg}")
        err_item.setFlags(Qt.NoItemFlags)
        self.mailbox_list.addItem(err_item)
        self.mailbox_list.blockSignals(False)
        self.load_btn.setEnabled(True)
        self._log("ERROR", "EXO", "FETCH", msg)
        self.statusBar().showMessage("Failed to load mailboxes")
        QMessageBox.critical(self, "Load Failed", f"Could not retrieve mailboxes:\n\n{msg}")

    def _populate_list(self, mailboxes: list):
        self.mailbox_list.blockSignals(True)
        self.mailbox_list.clear()
        for mb in mailboxes:
            item = QListWidgetItem(mb)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.mailbox_list.addItem(item)
        self.mailbox_list.blockSignals(False)
        self._update_sel_count()

    def _filter_mailboxes(self, text: str):
        if not self._all_mailboxes:
            return
        filtered = [m for m in self._all_mailboxes if text.lower() in m.lower()]
        self._populate_list(filtered)

    def _select_all(self):
        self.mailbox_list.blockSignals(True)
        for i in range(self.mailbox_list.count()):
            self.mailbox_list.item(i).setCheckState(Qt.Checked)
        self.mailbox_list.blockSignals(False)
        self._update_sel_count()

    def _select_none(self):
        self.mailbox_list.blockSignals(True)
        for i in range(self.mailbox_list.count()):
            self.mailbox_list.item(i).setCheckState(Qt.Unchecked)
        self.mailbox_list.blockSignals(False)
        self._update_sel_count()

    def _on_item_changed(self, _item):
        self._update_sel_count()

    def _update_sel_count(self):
        n = sum(
            1 for i in range(self.mailbox_list.count())
            if self.mailbox_list.item(i).checkState() == Qt.Checked
        )
        self.sel_count_label.setText(f"{n} selected")
        self.banner_sel.setText(f"Selected: {n}")

    def _get_selected_mailboxes(self):
        return [
            self.mailbox_list.item(i).text()
            for i in range(self.mailbox_list.count())
            if self.mailbox_list.item(i).checkState() == Qt.Checked
        ]

    # ── Helpers ───────────────────────────────
    def _log(self, level: str, module: str, event: str, message: str):
        ts   = datetime.now().strftime("%H:%M:%S")
        line = f"{ts} | {level:<5} | {module:<6} | {event:<8} | {message}"
        self.log_signal.emit(line)

    def _append_log(self, msg: str):
        color = "#222222"
        u = msg.upper()
        if "ERROR" in u or "FAILED" in u:
            color = "#cf222e"
        elif "WARN" in u:
            color = "#bf8700"
        elif "DONE" in u or "COMPLETE" in u or "SUCCESS" in u:
            color = "#1a7f37"
        self.log_text.append(f'<span style="color:{color}">{msg}</span>')
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _format_duration(self, seconds: float) -> str:
        s = int(seconds)
        m, s = divmod(s, 60)
        h, m = divmod(m, 60)
        return f"{h:02}:{m:02}:{s:02}" if h else f"{m:02}:{s:02}"

    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select PST Export Destination", self.output_input.text()
        )
        if folder:
            self.output_input.setText(folder)
            self._log("INFO", "CONFIG", "DEST", f"Output: {folder}")

    def _open_output(self):
        folder = self.output_input.text().strip()
        if folder and Path(folder).exists():
            import os
            os.startfile(folder)
        else:
            QMessageBox.warning(self, "Folder Not Found",
                                "The output folder does not exist yet.")

    # ── Archive ───────────────────────────────
    def _start_archive(self):
        mailboxes = self._get_selected_mailboxes()
        if not mailboxes:
            QMessageBox.warning(self, "No Mailboxes Selected",
                                "Tick at least one mailbox.\n"
                                "Use 'Load Mailboxes' if the list is empty.")
            return

        dest = self.output_input.text().strip()
        if not dest:
            QMessageBox.warning(self, "Missing Output Folder", "Select an output folder.")
            return

        Path(dest).mkdir(parents=True, exist_ok=True)
        self._start_time = datetime.now()

        self.card_phase.set_value("Running")
        self.card_progress.set_value(f"0 / {len(mailboxes)}")
        self.card_total.set_value(str(len(mailboxes)))
        self.card_current.set_value("-")
        self.card_duration.set_value("00:00")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_detail.setText("Starting...")
        self.banner_status.setText("Status: Archiving")

        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.load_btn.setEnabled(False)
        self.output_input.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.mailbox_list.setEnabled(False)
        self.sel_all_btn.setEnabled(False)
        self.sel_none_btn.setEnabled(False)

        self._log("INFO", "APP", "START", f"Starting archive — {len(mailboxes)} mailbox(es)")
        self.statusBar().showMessage("Archive in progress...")
        self._tick()

        self.worker = ArchiveWorker(self.admin_email, mailboxes, dest)
        self.worker.progress.connect(self._on_progress)
        self.worker.log.connect(self._append_log)
        self.worker.completed.connect(self._on_completed)
        self.worker.error.connect(self._on_error)
        self.worker.cancelled.connect(self._on_cancelled)
        self.worker.start()

    def _cancel_archive(self):
        if self.worker:
            self._log("WARN", "APP", "CANCEL", "Cancellation requested")
            self.worker.cancel()
            self.cancel_btn.setEnabled(False)

    def _tick(self):
        if self._start_time and not self.start_btn.isEnabled():
            elapsed = (datetime.now() - self._start_time).total_seconds()
            self.card_duration.set_value(self._format_duration(elapsed))
            QTimer.singleShot(1000, self._tick)

    def _on_progress(self, phase: str, done: int, total: int):
        self.card_phase.set_value(phase)
        if total > 0:
            pct = int(done / total * 100)
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(pct)
            self.card_progress.set_value(f"{done} / {total}")
            self.progress_detail.setText(f"{phase}  |  {done} of {total}  |  {pct}%")
            self.statusBar().showMessage(f"{phase} | {done}/{total} | {pct}%")
        else:
            self.progress_bar.setRange(0, 0)
            self.progress_detail.setText(phase)
            self.statusBar().showMessage(phase)
        if phase.startswith("Processing:"):
            self.card_current.set_value(phase.replace("Processing:", "").strip())

    def _on_completed(self, destination: str):
        self._reset()
        self.card_phase.set_value("Complete")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.progress_detail.setText("Complete")
        self.banner_status.setText("Status: Complete")
        self.statusBar().showMessage("Archive completed")
        self._log("INFO", "APP", "DONE", "Archive completed successfully")
        QMessageBox.information(self, "Archive Complete",
                                f"PST export finished.\n\nFiles saved to:\n{destination}")

    def _on_error(self, message: str):
        self._reset()
        self.card_phase.set_value("Error")
        self.progress_bar.setRange(0, 100)
        self.banner_status.setText("Status: Error")
        self.statusBar().showMessage("Archive failed")
        self._log("ERROR", "APP", "FAIL", message)
        QMessageBox.critical(self, "Archive Failed", message)

    def _on_cancelled(self):
        self._reset()
        self.card_phase.set_value("Cancelled")
        self.progress_bar.setRange(0, 100)
        self.banner_status.setText("Status: Cancelled")
        self.statusBar().showMessage("Cancelled")
        self._log("WARN", "APP", "CANCEL", "Archive cancelled by user")

    def _reset(self):
        self._start_time = None
        self.worker      = None
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.load_btn.setEnabled(True)
        self.output_input.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.mailbox_list.setEnabled(True)
        has_mbx = self.mailbox_list.count() > 0
        self.sel_all_btn.setEnabled(has_mbx)
        self.sel_none_btn.setEnabled(has_mbx)

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self, "Confirm Exit", "Close M365 PST Archiver?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        event.accept() if reply == QMessageBox.Yes else event.ignore()


# ──────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="M365 PST Archiver")
    parser.add_argument("--build-exe", action="store_true")
    parser.add_argument("--version",   action="store_true")
    args = parser.parse_args()

    if args.version:
        print(f"M365 PST Archiver v{APP_VERSION}")
        return 0
    if args.build_exe:
        return build_executable()

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    setup = ConnectionSetupDialog()
    if setup.exec() != QDialog.Accepted:
        sys.exit(0)

    window = MainWindow(admin_email=setup.admin_email)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()