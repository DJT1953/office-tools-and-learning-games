# Document Renamer.ps1
# The window for Document Renamer. Pick a folder, and its .rtf, .doc, .docx
# and .txt files (or all its files) get new names: made from the first
# characters of their text, from the names they have now, or from typed
# text, changed by replacing words, removing characters or changing capital
# letters, with text, dates, author, title or numbers added. A list shows
# every old and new name as the choices change. Then either renamed copies
# are saved in a "renamed" folder inside it (the originals stay as they are),
# or the files themselves are renamed, and the old names can be put back
# afterwards. Start it with the "Document Renamer" shortcut.
param([string]$Folder)

$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Word and the window both need a single-threaded (STA) PowerShell
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $argList = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $MyInvocation.MyCommand.Path))
    if ($Folder) { $argList += ('"{0}"' -f $Folder) }
    Start-Process $ps -ArgumentList $argList
    return
}

. (Join-Path $AppDir 'RenameEngine.ps1')
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# The window for choosing a folder (Windows' Open window, which shows the files
# in each folder, with a "Use this folder" button added), and a helper that
# makes Word retry when it is busy instead of failing ("call was rejected by callee")
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace DocRenamer
{
    [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IFileDialogEvents pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
    }

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    public class FileOpenDialogClass { }

    [ComImport, Guid("973510DB-7D7F-452B-8975-74A85828D354"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileDialogEvents
    {
        [PreserveSig] int OnFileOk(IFileDialog pfd);
        [PreserveSig] int OnFolderChanging(IFileDialog pfd, IShellItem psiFolder);
        [PreserveSig] int OnFolderChange(IFileDialog pfd);
        [PreserveSig] int OnSelectionChange(IFileDialog pfd);
        [PreserveSig] int OnShareViolation(IFileDialog pfd, IShellItem psi, out int pResponse);
        [PreserveSig] int OnTypeChange(IFileDialog pfd);
        [PreserveSig] int OnOverwrite(IFileDialog pfd, IShellItem psi, out int pResponse);
    }

    [ComImport, Guid("36116642-D713-4B97-9B83-7484A9D00433"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileDialogControlEvents
    {
        [PreserveSig] int OnItemSelected(IFileDialogCustomize pfdc, uint dwIDCtl, uint dwIDItem);
        [PreserveSig] int OnButtonClicked(IFileDialogCustomize pfdc, uint dwIDCtl);
        [PreserveSig] int OnCheckButtonToggled(IFileDialogCustomize pfdc, uint dwIDCtl, bool bChecked);
        [PreserveSig] int OnControlActivating(IFileDialogCustomize pfdc, uint dwIDCtl);
    }

    [ComImport, Guid("e6fdd21a-163f-4975-9c8c-a69f1ba37034"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileDialogCustomize
    {
        void EnableOpenDropDown(uint dwIDCtl);
        void AddMenu(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void AddPushButton(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void AddComboBox(uint dwIDCtl);
        void AddRadioButtonList(uint dwIDCtl);
        void AddCheckButton(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel, bool bChecked);
        void AddEditBox(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void AddSeparator(uint dwIDCtl);
        void AddText(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetControlLabel(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetControlState(uint dwIDCtl, out uint pdwState);
        void SetControlState(uint dwIDCtl, uint dwState);
        void GetEditBoxText(uint dwIDCtl, out IntPtr ppszText);
        void SetEditBoxText(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void GetCheckButtonState(uint dwIDCtl, out bool pbChecked);
        void SetCheckButtonState(uint dwIDCtl, bool bChecked);
        void AddControlItem(uint dwIDCtl, int dwIDItem, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void RemoveControlItem(uint dwIDCtl, int dwIDItem);
        void RemoveAllControlItems(uint dwIDCtl);
        void GetControlItemState(uint dwIDCtl, int dwIDItem, out uint pdwState);
        void SetControlItemState(uint dwIDCtl, int dwIDItem, uint dwState);
        void GetSelectedControlItem(uint dwIDCtl, out int pdwIDItem);
        void SetSelectedControlItem(uint dwIDCtl, int dwIDItem);
        void StartVisualGroup(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void EndVisualGroup();
        void MakeProminent(uint dwIDCtl);
        void SetControlItemText(uint dwIDCtl, uint dwIDItem, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    }

    [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleWindow
    {
        void GetWindow(out IntPtr phwnd);
        void ContextSensitiveHelp(bool fEnterMode);
    }

    // Listens for the "Use this folder" button, and for Open
    public class FolderButtonEvents : IFileDialogEvents, IFileDialogControlEvents
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

        public IFileDialog Dialog;
        public string Chosen;

        // The full path of a shell item when it is a folder on the computer, or null
        private static string FolderPath(IShellItem item)
        {
            try
            {
                string path;
                item.GetDisplayName(0x80058000, out path);
                if (!string.IsNullOrEmpty(path) && System.IO.Directory.Exists(path)) return path;
            }
            catch { }
            return null;
        }

        // A highlighted folder is the one chosen, as in Windows' own folder chooser;
        // otherwise the folder that is open
        public int OnButtonClicked(IFileDialogCustomize pfdc, uint dwIDCtl)
        {
            string path = null;
            try { IShellItem selected; Dialog.GetCurrentSelection(out selected); path = FolderPath(selected); } catch { }
            if (path == null) { try { IShellItem folder; Dialog.GetFolder(out folder); path = FolderPath(folder); } catch { } }
            if (path != null)
            {
                Chosen = path;
                Dialog.Close(0);
                return 0;
            }
            // Such as "This PC" or a library: not one folder on the computer
            IntPtr window = IntPtr.Zero;
            try { ((IOleWindow)Dialog).GetWindow(out window); } catch { }
            MessageBoxW(window, "This place is not a single folder on the computer. Open the folder that holds your files, then click \"Use this folder\".", "Document Renamer", 0x40);
            return 0;
        }

        // A file picked in the folder that is open chooses that folder - even a
        // shortcut to a file somewhere else. (A typed path is followed as usual.)
        public int OnFileOk(IFileDialog pfd)
        {
            try
            {
                string typed;
                pfd.GetFileName(out typed);
                if (!string.IsNullOrEmpty(typed) && typed.IndexOfAny(new char[] { '\\', '/', ':' }) < 0)
                {
                    IShellItem folder;
                    pfd.GetFolder(out folder);
                    string path = FolderPath(folder);
                    if (path != null) Chosen = path;
                }
            }
            catch { }
            return 0;
        }
        public int OnFolderChanging(IFileDialog pfd, IShellItem psiFolder) { return 0; }
        public int OnFolderChange(IFileDialog pfd) { return 0; }
        public int OnSelectionChange(IFileDialog pfd) { return 0; }
        public int OnShareViolation(IFileDialog pfd, IShellItem psi, out int pResponse) { pResponse = 0; return 0; }
        public int OnTypeChange(IFileDialog pfd) { return 0; }
        public int OnOverwrite(IFileDialog pfd, IShellItem psi, out int pResponse) { pResponse = 0; return 0; }
        public int OnItemSelected(IFileDialogCustomize pfdc, uint dwIDCtl, uint dwIDItem) { return 0; }
        public int OnCheckButtonToggled(IFileDialogCustomize pfdc, uint dwIDCtl, bool bChecked) { return 0; }
        public int OnControlActivating(IFileDialogCustomize pfdc, uint dwIDCtl) { return 0; }
    }

    public static class FolderPicker
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(string path, IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem item);

        // Shows the files in each folder while choosing. Returns the folder chosen
        // with "Use this folder", or the folder of a file chosen with Open, or null
        // if the window was cancelled.
        public static string Pick(IntPtr owner, string title, string startFolder)
        {
            IFileDialog dialog = (IFileDialog)new FileOpenDialogClass();
            FolderButtonEvents events = new FolderButtonEvents();
            uint cookie = 0;
            try
            {
                uint options;
                dialog.GetOptions(out options);
                dialog.SetOptions(options | 0x40 | 0x2000000);   // real folders only; don't add the file picked to Recent files
                dialog.SetTitle(title);
                IFileDialogCustomize custom = (IFileDialogCustomize)dialog;
                custom.AddPushButton(1, "Use this folder");
                custom.MakeProminent(1);
                events.Dialog = dialog;
                dialog.Advise(events, out cookie);
                if (!string.IsNullOrEmpty(startFolder) && System.IO.Directory.Exists(startFolder))
                {
                    try
                    {
                        IShellItem folder;
                        SHCreateItemFromParsingName(startFolder, IntPtr.Zero, typeof(IShellItem).GUID, out folder);
                        dialog.SetFolder(folder);
                    }
                    catch { }
                }
                int shown = dialog.Show(owner);
                if (events.Chosen != null) return events.Chosen;
                if (shown != 0) return null;
                IShellItem result;
                dialog.GetResult(out result);
                string path;
                result.GetDisplayName(0x80058000, out path);   // the chosen file's full path
                return System.IO.Path.GetDirectoryName(path);
            }
            finally
            {
                if (cookie != 0) { try { dialog.Unadvise(cookie); } catch { } }
                Marshal.ReleaseComObject(dialog);
            }
        }
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);
        [PreserveSig] int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
    }

    public class MessageFilter : IOleMessageFilter
    {
        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new MessageFilter(), out old); }
        public static void Revoke() { IOleMessageFilter old; CoRegisterMessageFilter(null, out old); }

        int IOleMessageFilter.HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo) { return 0; }
        int IOleMessageFilter.RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType)
        {
            if (rejectType == 2 && tickCount < 120000) return 250;   // busy: try again in a quarter second, for up to 2 minutes
            return -1;
        }
        int IOleMessageFilter.MessagePending(IntPtr taskCallee, int tickCount, int pendingType) { return 2; }
    }
}
"@


# ---------------------------------------------------------------- remembered settings

$SettingsFile = Join-Path $env:APPDATA 'Document Renamer\settings.json'
# The naming choices (see $script:DefaultOptions in RenameEngine.ps1), the last
# folder, and FileSet: 'Docs' for Word and text documents, 'All' for all files
$Defaults = @{} + $script:DefaultOptions
$Defaults.LastFolder = ''
$Defaults.FileSet = 'Docs'

function Get-SavedSettings {
    $s = @{} + $Defaults
    try {
        if (Test-Path -LiteralPath $SettingsFile) {
            $o = Get-Content -LiteralPath $SettingsFile -Raw | ConvertFrom-Json
            foreach ($p in $o.PSObject.Properties) { if ($s.ContainsKey($p.Name)) { $s[$p.Name] = $p.Value } }
        }
    } catch {}
    return $s
}

function Save-Settings([hashtable]$S) {
    try {
        $dir = Split-Path -Parent $SettingsFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        $S | ConvertTo-Json | Set-Content -LiteralPath $SettingsFile -Encoding UTF8
    } catch {}
}


# ---------------------------------------------------------------- small helpers

function Show-Message([string]$Text, [string]$Buttons = 'OK', [string]$Icon = 'Information') {
    return [System.Windows.Forms.MessageBox]::Show($Text, 'Document Renamer', $Buttons, $Icon)
}

function Select-Folder([string]$Start) {
    $title = 'Open the folder with the files to rename, then click "Use this folder"'
    try { return [DocRenamer.FolderPicker]::Pick([IntPtr]::Zero, $title, $Start) }
    catch {
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = $title
        if ($Start -and (Test-Path -LiteralPath $Start)) { $d.SelectedPath = $Start }
        if ($d.ShowDialog() -eq 'OK') { return $d.SelectedPath }
        return $null
    }
}

# "1 file", "3 files"
function Get-Count([int]$N, [string]$One, [string]$Many) {
    if ($N -eq 1) { return "1 $One" }
    return "$N $Many"
}

function Format-When([string]$When) {
    try { return ([datetime]::Parse($When, [Globalization.CultureInfo]::InvariantCulture)).ToString('MMMM d, yyyy, h:mm tt') }
    catch { return $When }
}

function New-Window([string]$Title) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding(14)
    if ($script:AppIcon) { $form.Icon = $script:AppIcon }
    return $form
}

function New-Label([string]$Text, [switch]$Bold, [single]$Size = 10, [int]$Width = 0) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true
    if ($Bold -or $Size -ne 10) {
        $style = [System.Drawing.FontStyle]::Regular
        if ($Bold) { $style = [System.Drawing.FontStyle]::Bold }
        $l.Font = New-Object System.Drawing.Font('Segoe UI', $Size, $style)
    }
    if ($Width -gt 0) { $l.MaximumSize = New-Object System.Drawing.Size($Width, 0) }
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 7, 3, 3)
    return $l
}

function New-Heading([string]$Text) {
    $l = New-Label $Text -Bold -Size 11
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 14, 3, 3)
    return $l
}

# A grid that grows to fit what is put in it
function New-Grid([int]$Columns) {
    $g = New-Object System.Windows.Forms.TableLayoutPanel
    $g.ColumnCount = $Columns
    $g.AutoSize = $true
    $g.AutoSizeMode = 'GrowAndShrink'
    $g.Margin = New-Object System.Windows.Forms.Padding(0)
    for ($i = 0; $i -lt $Columns; $i++) { [void]$g.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize'))) }
    return $g
}

# Puts controls on the next row of a grid: one in each column, or one across the whole row
function Add-Row($Grid, $Left, $Right, [switch]$Wide) {
    $row = $Grid.RowCount
    $Grid.RowCount = $row + 1
    if ($Wide) {
        $Grid.Controls.Add($Left, 0, $row)
        if ($Grid.ColumnCount -gt 1) { $Grid.SetColumnSpan($Left, $Grid.ColumnCount) }
    } else {
        if ($Left) { $Grid.Controls.Add($Left, 0, $row) }
        if ($Right) { $Grid.Controls.Add($Right, 1, $row) }
    }
}

# A row of controls side by side
function New-Strip {
    $p = New-Object System.Windows.Forms.FlowLayoutPanel
    $p.AutoSize = $true
    $p.WrapContents = $false
    $p.Margin = New-Object System.Windows.Forms.Padding(0)
    foreach ($c in $args) { $p.Controls.Add($c) }
    return $p
}

function New-Button([string]$Text) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.AutoSize = $true
    $b.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    return $b
}

# A box with a bold title around a two-column grid of choices. The grid is in the box's Tag.
function New-Group([string]$Title) {
    $box = New-Object System.Windows.Forms.GroupBox
    $box.Text = $Title
    $box.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $box.AutoSize = $true
    $box.AutoSizeMode = 'GrowAndShrink'
    $box.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 3, 10, 3)
    $grid = New-Grid 2
    $grid.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $grid.Dock = 'Fill'
    $box.Controls.Add($grid)
    $box.Tag = $grid
    return $box
}

function New-Radio([string]$Text, [bool]$Checked) {
    $r = New-Object System.Windows.Forms.RadioButton
    $r.Text = $Text
    $r.AutoSize = $true
    $r.Checked = $Checked
    $r.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 2)
    return $r
}

function New-Tick([string]$Text, [bool]$Checked) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.AutoSize = $true
    $c.Checked = $Checked
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 2)
    return $c
}

function New-TextBox([string]$Text, [int]$Width) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Text = $Text
    $t.Width = $Width
    $t.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    return $t
}

function New-Number([long]$Min, [long]$Max, $Value, [int]$Width = 55) {
    $n = New-Object System.Windows.Forms.NumericUpDown
    $n.Minimum = $Min
    $n.Maximum = $Max
    $n.Width = $Width
    $n.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    $v = 0L
    if (-not [long]::TryParse([string]$Value, [ref]$v)) { $v = $Min }
    $n.Value = [Math]::Min($Max, [Math]::Max($Min, $v))
    return $n
}

# A drop-down list. $Choices holds pairs of (setting value, text shown).
function New-Choice([object[]]$Choices, [string]$Selected, [int]$Width) {
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'
    $c.Width = $Width
    $c.MaxDropDownItems = 15
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    $keys = @()
    foreach ($pair in $Choices) { [void]$c.Items.Add($pair[1]); $keys += $pair[0] }
    $c.Tag = $keys
    $i = [array]::IndexOf($keys, [string]$Selected)
    if ($i -lt 0) { $i = 0 }
    $c.SelectedIndex = $i
    return $c
}

function Get-ChoiceKey($Combo) { return [string]$Combo.Tag[$Combo.SelectedIndex] }


# ---------------------------------------------------------------- reading the documents

function Invoke-ReadingWithProgress($Items) {
    $total = @($Items | Where-Object { Test-NeedsReading $_ }).Count
    if ($total -eq 0) { return }
    $form = New-Window 'Document Renamer - reading'
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.AutoSize = $true
    $panel.FlowDirection = 'TopDown'
    $panel.WrapContents = $false
    $label = New-Label 'Getting ready...'
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(520, 48)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Size = New-Object System.Drawing.Size(520, 22)
    $bar.Maximum = $total
    $stopBtn = New-Button 'Stop reading'
    $stopBtn.Anchor = 'Right'
    $panel.Controls.Add($label)
    $panel.Controls.Add($bar)
    $panel.Controls.Add($stopBtn)
    $form.Controls.Add($panel)

    # The buttons are handled while the reading runs, so they talk through $script:Progress
    $script:Progress = @{ Label = $label; Bar = $bar; StopButton = $stopBtn; Stop = $false; Done = $false }
    $stopBtn.Add_Click({
        $script:Progress.Stop = $true
        $script:Progress.StopButton.Enabled = $false
        $script:Progress.Label.Text = 'Stopping after this file...'
    })
    $form.Add_FormClosing({
        param($sender, $closeArgs)
        if (-not $script:Progress.Done) {
            $closeArgs.Cancel = $true
            $script:Progress.Stop = $true
            $script:Progress.StopButton.Enabled = $false
            $script:Progress.Label.Text = 'Stopping after this file...'
        }
    })
    $form.Show()
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    try {
        Read-DocumentStarts -Items $Items `
            -OnFile {
                param($number, $count, $item)
                $p = $script:Progress
                $p.Bar.Value = $number - 1
                if (-not $p.Stop) { $p.Label.Text = "Reading file $number of ${count}:`r`n$($item.Name)" }
                [System.Windows.Forms.Application]::DoEvents()
            } `
            -OnTick { [System.Windows.Forms.Application]::DoEvents() } `
            -ShouldStop { $script:Progress.Stop }
    } finally {
        $script:Progress.Done = $true
        $form.Close()
        $form.Dispose()
    }
}


# ---------------------------------------------------------------- the choices and the list of old and new names

# $Known: the files of this folder already read, by full path (from an earlier
# visit to this folder), or $null
function Show-PreviewForm([string]$Folder, $Items, [hashtable]$Last, $History, [hashtable]$Known) {
    $form = New-Window 'Document Renamer'

    # The window, from top to bottom: the folder and which files, the three
    # numbered steps that make the new names, the list of old and new names,
    # what to do, then the buttons
    $page = New-Grid 1
    $form.Controls.Add($page)
    $top = New-Grid 2
    $bottom = New-Grid 1
    Add-Row $page $top

    Add-Row $top (New-Label 'Rename the files in this folder' -Bold -Size 12) -Wide
    $changeBtn = New-Button 'Change folder...'
    Add-Row $top (New-Label 'Folder:') (New-Strip (New-Label $Folder -Width 760) $changeBtn)
    $docsRadio = New-Radio 'Word and text documents (.rtf, .doc, .docx, .txt)' ($Last.FileSet -ne 'All')
    $allRadio = New-Radio 'All files' ($Last.FileSet -eq 'All')
    Add-Row $top (New-Label 'Files:') (New-Strip $docsRadio $allRadio)

    $steps = New-Strip
    $steps.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    Add-Row $page $steps

    # --- 1. what the new name starts from
    $g1 = New-Group '1. Start with'
    $srcName = New-Radio 'The name it has now' ($Last.Source -eq 'Name')
    $srcText = New-Radio 'The first' ($Last.Source -ne 'Name' -and $Last.Source -ne 'Fixed')
    $length = [Math]::Min(30, [Math]::Max(1, [int]$Last.Length))
    $lengthBox = New-Choice @(foreach ($i in 1..30) { , @("$i", "$i") }) "$length" 55
    $srcFixed = New-Radio 'This text:' ($Last.Source -eq 'Fixed')
    $fixedBox = New-TextBox $Last.FixedText 170
    Add-Row $g1.Tag $srcName $null
    Add-Row $g1.Tag $srcText (New-Strip $lengthBox (New-Label 'characters of its text'))
    Add-Row $g1.Tag $srcFixed $fixedBox
    $steps.Controls.Add($g1)

    # --- 2. changes to it
    $g2 = New-Group '2. Change it'
    $findBox = New-TextBox $Last.Find 150
    $withBox = New-TextBox $Last.ReplaceWith 150
    Add-Row $g2.Tag (New-Label 'Replace:') $findBox
    Add-Row $g2.Tag (New-Label 'with:') $withBox
    $matchTick = New-Tick 'Only where the capital letters match' ([bool]$Last.MatchCase)
    Add-Row $g2.Tag $null $matchTick
    $firstBox = New-Number 0 99 $Last.RemoveFirst
    $lastBox = New-Number 0 99 $Last.RemoveLast
    Add-Row $g2.Tag (New-Label 'Remove:') (New-Strip (New-Label 'the first') $firstBox (New-Label 'characters'))
    Add-Row $g2.Tag $null (New-Strip (New-Label 'the last') $lastBox (New-Label 'characters'))
    $caseBox = New-Choice @(@('Leave', 'Leave them as they are'), @('Lower', 'all small letters'), @('Upper', 'ALL CAPITALS'),
        @('Title', 'Capital At The Start Of Each Word'), @('Sentence', 'Capital first letter only')) $Last.Case 230
    Add-Row $g2.Tag (New-Label 'Capital letters:') $caseBox
    $steps.Controls.Add($g2)

    # --- 3. what is added to it
    $g3 = New-Group '3. Add to it'
    $prefixBox = New-TextBox $Last.Prefix 170
    $suffixBox = New-TextBox $Last.Suffix 170
    Add-Row $g3.Tag (New-Label 'Before the name:') $prefixBox
    Add-Row $g3.Tag (New-Label 'After the name:') $suffixBox
    $infoBox = New-Choice @(@('None', 'nothing'), @('Created', 'date created'), @('Modified', 'date modified'), @('Today', "today's date"),
        @('Taken', 'date taken (photos, videos)'), @('Author', 'author (Word documents)'), @('Title', 'title (Word documents)'),
        @('Folder', 'the folder name')) $Last.Info 190
    $infoAtBox = New-Choice @(@('Start', 'at the start'), @('End', 'at the end')) $Last.InfoAt 105
    Add-Row $g3.Tag (New-Label 'Information:') (New-Strip $infoBox $infoAtBox)
    $now = Get-Date
    $dateBox = New-Choice @(foreach ($f in 'yyyy-MM-dd', 'yyyy-MM-dd HH.mm', 'MM-dd-yyyy', 'yyyyMMdd', 'MMMM d, yyyy') { , @($f, $now.ToString($f)) }) $Last.DateFormat 190
    Add-Row $g3.Tag (New-Label 'Dates look like:') $dateBox
    $numberTick = New-Tick 'Number them:' ([bool]$Last.Number)
    $numberAtBox = New-Choice @(@('Start', 'at the start'), @('End', 'at the end')) $Last.NumberAt 105
    $fromBox = New-Number 0 999999 $Last.NumberFrom 70
    Add-Row $g3.Tag $numberTick (New-Strip $numberAtBox (New-Label 'starting at') $fromBox)
    $digitsBox = New-Choice @(@('1', '1, 2, 3'), @('2', '01, 02, 03'), @('3', '001, 002, 003'), @('4', '0001, 0002'),
        @('5', '00001, 00002'), @('6', '000001, 000002')) ([string]$Last.NumberDigits) 120
    $stepBox = New-Number 1 1000 $Last.NumberStep
    Add-Row $g3.Tag $null (New-Strip $digitsBox (New-Label 'going up by') $stepBox)
    $orderBox = New-Choice @(@('Name', 'name'), @('Created', 'date created'), @('Modified', 'date modified')) $Last.NumberOrder 130
    Add-Row $g3.Tag $null (New-Strip (New-Label 'in order of') $orderBox)
    $joinerBox = New-Choice @(@(' - ', 'space, dash, space   ( - )'), @(' ', 'a space'), @('_', 'underscore   ( _ )'),
        @('-', 'dash, no spaces   ( - )'), @('', 'nothing')) $Last.Joiner 220
    Add-Row $g3.Tag (New-Label 'Between the parts:') $joinerBox
    $steps.Controls.Add($g3)

    # --- the list of old and new names
    $wide = 1010
    try { $wide = [Math]::Max(1010, $steps.GetPreferredSize([System.Drawing.Size]::Empty).Width - 6) } catch {}
    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.HeaderStyle = 'Nonclickable'
    $list.ShowItemToolTips = $true
    $list.Size = New-Object System.Drawing.Size($wide, 250)
    $list.Margin = New-Object System.Windows.Forms.Padding(3, 10, 3, 3)
    $noteWidth = [Math]::Max(250, $wide - 690)
    [void]$list.Columns.Add('Name now', 310)
    [void]$list.Columns.Add('New name', 350)
    [void]$list.Columns.Add('Note', $noteWidth)
    Add-Row $page $list
    $countLabel = New-Label '' -Width $wide
    Add-Row $page $countLabel
    Add-Row $page $bottom

    # --- what to do
    Add-Row $bottom (New-Heading 'What to do')
    $copyRadio = New-Radio 'Save renamed copies in a folder called "renamed" inside this folder (your files stay as they are)' $true
    $renameRadio = New-Radio 'Rename the files themselves (you can put the old names back afterwards)' $false
    $radioPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $radioPanel.AutoSize = $true
    $radioPanel.FlowDirection = 'TopDown'
    $radioPanel.WrapContents = $false
    $radioPanel.Controls.Add($copyRadio)
    $radioPanel.Controls.Add($renameRadio)
    Add-Row $bottom $radioPanel
    $modeNote = New-Label '' -Width $wide
    $modeNote.ForeColor = [System.Drawing.Color]::DimGray
    Add-Row $bottom $modeNote

    # --- the last renaming in this folder, if there was one
    $undoBtn = $null
    if ($History) {
        $when = Format-When $History.When
        $undoCount = @($History.Entries).Count
        $undoBtn = New-Button 'Put the old names back...'
        $undoLabel = New-Label ('Last time, on ' + $when + ', ' + (Get-Count $undoCount 'file' 'files') + ' in this folder got new names.')
        Add-Row $bottom (New-Strip $undoLabel $undoBtn)
    }

    # --- the buttons
    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.AutoSize = $true
    $buttons.FlowDirection = 'RightToLeft'
    $buttons.Dock = 'Fill'
    $buttons.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)
    $goBtn = New-Button 'Save renamed copies'
    $goBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $cancelBtn = New-Button 'Cancel'
    $cancelBtn.DialogResult = 'Cancel'
    $buttons.Controls.Add($goBtn)
    $buttons.Controls.Add($cancelBtn)
    Add-Row $bottom $buttons
    # No AcceptButton: pressing Enter after typing must not start the renaming
    $form.CancelButton = $cancelBtn

    $existing = @()
    try { $existing = @(Get-ChildItem -LiteralPath $Folder -Force | ForEach-Object { $_.Name }) } catch {}
    # Files already read keep their text when the list switches between documents and all files
    if (-not $Known) { $Known = @{} }
    $state = @{ Items = $Items; Known = $Known; Plan = $null; Mode = 'Copy'; Options = $null; Busy = $false; Again = $false }
    foreach ($it in $Items) { $state.Known[$it.Path] = $it }

    # A number box's number as typed so far, so the list follows along while typing
    $readNumber = {
        param($box)
        $v = 0
        if ([int]::TryParse($box.Text, [ref]$v) -and $v -ge $box.Minimum -and $v -le $box.Maximum) { return $v }
        return [int]$box.Value
    }

    $readOptions = {
        $source = 'Name'
        if ($srcText.Checked) { $source = 'Text' } elseif ($srcFixed.Checked) { $source = 'Fixed' }
        return @{
            Source = $source; Length = [int](Get-ChoiceKey $lengthBox); FixedText = $fixedBox.Text
            Find = $findBox.Text; ReplaceWith = $withBox.Text; MatchCase = $matchTick.Checked
            RemoveFirst = (& $readNumber $firstBox); RemoveLast = (& $readNumber $lastBox); Case = (Get-ChoiceKey $caseBox)
            Prefix = $prefixBox.Text; Suffix = $suffixBox.Text; Info = (Get-ChoiceKey $infoBox); InfoAt = (Get-ChoiceKey $infoAtBox)
            DateFormat = (Get-ChoiceKey $dateBox); Number = $numberTick.Checked; NumberAt = (Get-ChoiceKey $numberAtBox)
            NumberFrom = (& $readNumber $fromBox); NumberStep = (& $readNumber $stepBox); NumberDigits = [int](Get-ChoiceKey $digitsBox)
            NumberOrder = (Get-ChoiceKey $orderBox); Joiner = (Get-ChoiceKey $joinerBox) }
    }

    $loadItems = {
        try {
            $state.Items = Merge-KnownItems (Find-Files $Folder -AllFiles:$allRadio.Checked) $state.Known
        } catch { Show-Message ("The files in this folder could not be listed:`n`n" + $_.Exception.Message) 'OK' 'Warning' | Out-Null }
    }

    # Reads what the choices need and hasn't been read yet: the documents' text,
    # or the author, title and date taken
    $ensureData = {
        param($o)
        if ($o.Source -eq 'Text' -and @($state.Items | Where-Object { Test-NeedsReading $_ }).Count) {
            $form.Enabled = $false
            try { Invoke-ReadingWithProgress $state.Items } finally { $form.Enabled = $true; $form.Activate() }
        }
        if (@('Taken', 'Author', 'Title') -contains $o.Info) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            try { Read-FileProperties $state.Items $Folder } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
        }
    }

    $fillList = {
        $o = & $readOptions
        & $ensureData $o
        $mode = 'Copy'
        if ($renameRadio.Checked) { $mode = 'Rename' }
        $plan = Get-RenamePlan -Items $state.Items -Options $o -Mode $mode -ExistingNames $existing -Folder $Folder
        $state.Plan = $plan; $state.Mode = $mode; $state.Options = $o
        $list.BeginUpdate()
        $list.Items.Clear()
        foreach ($row in $plan) {
            $lvi = New-Object System.Windows.Forms.ListViewItem($row.Item.Name)
            $note = $row.Note
            if ($row.Action -eq 'Keep') {
                if ($mode -eq 'Copy') { [void]$lvi.SubItems.Add('(not copied)') } else { [void]$lvi.SubItems.Add('(keeps its name)') }
                if ($row.ProblemKind -eq 'Error') { $note = 'could not be read: ' + $note; $lvi.ForeColor = [System.Drawing.Color]::DarkRed }
                else { $lvi.ForeColor = [System.Drawing.Color]::DimGray }
            } else {
                [void]$lvi.SubItems.Add($row.NewName)
                if ($row.Action -eq 'Same') { $lvi.ForeColor = [System.Drawing.Color]::DimGray }
            }
            if ($row.Item.Note -and $o.Source -eq 'Text') {
                if ($note) { $note += '; ' }
                $note += $row.Item.Note
            }
            [void]$lvi.SubItems.Add($note)
            $lvi.ToolTipText = $note
            [void]$list.Items.Add($lvi)
        }
        $list.EndUpdate()

        $go = @($plan | Where-Object { $_.Action -eq 'Rename' }).Count
        $same = @($plan | Where-Object { $_.Action -eq 'Same' }).Count
        $keep = @($plan | Where-Object { $_.Action -eq 'Keep' }).Count
        if (@($plan).Count -eq 0) {
            if ($allRadio.Checked) { $text = 'There are no files in this folder.' }
            else { $text = 'There are no Word or text documents in this folder. Choose "All files" at the top to rename other kinds of files.' }
        } elseif ($mode -eq 'Copy') {
            $text = (Get-Count $go 'renamed copy' 'renamed copies') + ' will be saved.'
            if ($keep) { $text += ' ' + (Get-Count $keep 'file' 'files') + ' will not be copied (the notes say why).' }
        } else {
            $text = (Get-Count $go 'file' 'files') + ' will get a new name.'
            if ($same) { $text += ' ' + (Get-Count $same 'file already has its' 'files already have their') + ' new name.' }
            if ($keep) { $text += ' ' + (Get-Count $keep 'file keeps its' 'files keep their') + ' name (the notes say why).' }
        }
        if ($mode -eq 'Copy') {
            $goBtn.Text = 'Save renamed copies'
            $modeNote.Text = 'The copies go into a folder called "renamed" inside this folder. Your original files are not changed. ' +
                'If "renamed" already has a copy with the same name from an earlier run, it is replaced.'
        } else {
            $goBtn.Text = 'Rename the files'
            $modeNote.Text = 'The files in this folder get their new names. To undo it, click "Put the old names back" when it is done, ' +
                'or open this folder in Document Renamer again later.'
        }
        $countLabel.Text = $text
    }

    # Every change to a choice redoes the list. A change made while the list is
    # being redone (such as while documents are read) redoes it once more after.
    $refresh = {
        $typingPause.Stop()
        if ($state.Busy) { $state.Again = $true; return }
        $state.Busy = $true
        try {
            do { $state.Again = $false; & $fillList } while ($state.Again)
        } finally { $state.Busy = $false }
    }
    # While typing, the list is redone when the typing pauses, so a folder with
    # many files doesn't slow the typing down
    $typingPause = New-Object System.Windows.Forms.Timer
    $typingPause.Interval = 300
    $typingPause.Add_Tick({ & $refresh })
    $refreshSoon = { $typingPause.Stop(); $typingPause.Start() }

    # Typing in a box or picking from a list chooses the option it belongs to
    $lengthBox.Add_SelectedIndexChanged({ $srcText.Checked = $true; & $refresh })
    $fixedBox.Add_TextChanged({ $srcFixed.Checked = $true; & $refreshSoon })
    foreach ($r in $srcName, $srcText, $srcFixed, $copyRadio, $renameRadio) { $r.Add_CheckedChanged({ if ($this.Checked) { & $refresh } }) }
    $docsRadio.Add_CheckedChanged({ & $loadItems; & $refresh })
    foreach ($b in $findBox, $withBox, $prefixBox, $suffixBox) { $b.Add_TextChanged({ & $refreshSoon }) }
    foreach ($t in $matchTick, $numberTick) { $t.Add_CheckedChanged({ & $refresh }) }
    foreach ($n in $firstBox, $lastBox) { $n.Add_ValueChanged({ & $refresh }); $n.Add_KeyUp({ & $refreshSoon }) }
    foreach ($c in $caseBox, $infoBox, $infoAtBox, $dateBox, $joinerBox) { $c.Add_SelectedIndexChanged({ & $refresh }) }
    foreach ($c in $numberAtBox, $digitsBox, $orderBox) { $c.Add_SelectedIndexChanged({ $numberTick.Checked = $true; & $refresh }) }
    foreach ($n in $fromBox, $stepBox) {
        $n.Add_ValueChanged({ $numberTick.Checked = $true; & $refresh })
        $n.Add_KeyUp({ $numberTick.Checked = $true; & $refreshSoon })
    }
    & $refresh

    # On a small screen the window scrolls instead of running off the bottom
    $form.Add_Load({
        try {
            $area = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
            if ($form.Height -gt $area.Height) {
                $w = $form.Width + [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
                $form.AutoSize = $false
                $form.AutoScroll = $true
                $form.Size = New-Object System.Drawing.Size($w, $area.Height)
                $form.Location = New-Object System.Drawing.Point([Math]::Max($area.Left, $form.Left - 10), $area.Top)
            }
        } catch {}
    })
    $form.Add_Shown({ $goBtn.Focus() })

    $readFileSet = { if ($allRadio.Checked) { 'All' } else { 'Docs' } }
    # The choices go along to the next folder
    $changeBtn.Add_Click({
        $form.Tag = @{ Action = 'ChangeFolder'; Options = (& $readOptions); FileSet = (& $readFileSet); Known = $state.Known }
        $form.DialogResult = 'Retry'
    })
    $goBtn.Add_Click({
        & $refresh   # the list as it is now, in case a box was still being typed in
        $plan = $state.Plan
        $go = @($plan | Where-Object { $_.Action -eq 'Rename' }).Count
        if ($go -eq 0) {
            Show-Message 'With these choices, no file gets a new name. The Note column in the list says why for each file.' 'OK' 'Information' | Out-Null
            return
        }
        if ($state.Mode -eq 'Rename') {
            $answer = Show-Message ('Give ' + (Get-Count $go 'file' 'files') + " in this folder a new name?`n`n$Folder`n`n" +
                'You can put the old names back afterwards.') 'YesNo' 'Question'
            if ($answer -ne 'Yes') { return }
        }
        $form.Tag = @{ Action = 'Rename'; Mode = $state.Mode; Options = $state.Options; FileSet = (& $readFileSet); Plan = $plan }
        $form.DialogResult = 'OK'
    })
    if ($undoBtn) {
        $undoBtn.Add_Click({
            $answer = Show-Message ('Put back the old names of the ' + (Get-Count $undoCount 'file' 'files') + " that got new names on ${when}?") 'YesNo' 'Question'
            if ($answer -ne 'Yes') { return }
            $form.Tag = @{ Action = 'Undo' }
            $form.DialogResult = 'OK'
        })
    }

    $result = $form.ShowDialog()
    $choice = $form.Tag
    $typingPause.Dispose()
    $form.Dispose()
    if ($result -eq 'OK' -or $result -eq 'Retry') { return $choice }
    return $null
}


# ---------------------------------------------------------------- the summary

# $Mode: 'Copy', 'Rename' or 'Undo' (the old names were put back)
function Show-Summary($Results, [string]$Folder, [string]$Mode) {
    $done = @($Results | Where-Object { $_.Status -eq 'Done' })
    $same = @($Results | Where-Object { $_.Status -eq 'Same' })
    $keep = @($Results | Where-Object { $_.Status -eq 'Keep' })
    $fail = @($Results | Where-Object { $_.Status -eq 'Failed' })
    $outFolder = $Folder
    if ($Mode -eq 'Copy') { $outFolder = Join-Path $Folder $script:CopyFolderName }
    $words = @{
        Copy   = @{ Done = 'COPIED WITH THEIR NEW NAMES'; Keep = 'NOT COPIED'; Failed = 'COULD NOT BE COPIED'
                    Head = 'renamed copies saved'; HeadKeep = 'not copied'; HeadFail = 'could not be copied' }
        Rename = @{ Done = 'RENAMED'; Same = 'ALREADY HAD THIS NAME'; Keep = 'KEPT THEIR NAME'; Failed = 'COULD NOT BE RENAMED'
                    Head = 'renamed'; HeadSame = 'already had the name'; HeadKeep = 'kept their name'; HeadFail = 'could not be renamed' }
        Undo   = @{ Done = 'OLD NAMES PUT BACK'; Same = 'ALREADY HAD THEIR OLD NAME'; Failed = 'COULD NOT BE PUT BACK'
                    Head = 'old names put back'; HeadSame = 'already had their old name'; HeadFail = 'could not be put back' }
    }[$Mode]

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Folder:  $Folder")
    if ($Mode -eq 'Copy') { [void]$sb.AppendLine("Copies saved in:  $outFolder") }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("$($words.Done) ($($done.Count))")
    if ($done.Count -eq 0) { [void]$sb.AppendLine('   (none)') }
    foreach ($r in $done) {
        [void]$sb.AppendLine("   $($r.Name)   ->   $($r.NewName)")
        if ($r.Detail) { [void]$sb.AppendLine("      $($r.Detail)") }
    }
    if ($same.Count) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("$($words.Same) ($($same.Count))")
        foreach ($r in $same) { [void]$sb.AppendLine("   $($r.NewName)") }
    }
    foreach ($group in @(@($words.Keep, $keep), @($words.Failed, $fail))) {
        if ($group[1].Count -eq 0) { continue }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("$($group[0]) ($($group[1].Count))")
        foreach ($r in $group[1]) {
            if ($r.NewName) { [void]$sb.AppendLine("   $($r.Name)   ->   $($r.NewName)") } else { [void]$sb.AppendLine("   $($r.Name)") }
            [void]$sb.AppendLine("      $($r.Detail)")
        }
    }

    $form = New-Window 'Document Renamer - finished'
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.AutoSize = $true
    $panel.FlowDirection = 'TopDown'
    $panel.WrapContents = $false

    $headline = "Finished: $($done.Count) $($words.Head)"
    if ($same.Count) { $headline += ", $($same.Count) $($words.HeadSame)" }
    if ($keep.Count) { $headline += ", $($keep.Count) $($words.HeadKeep)" }
    if ($fail.Count) { $headline += ", $($fail.Count) $($words.HeadFail)" }
    $panel.Controls.Add((New-Label $headline -Bold -Size 12 -Width 860))
    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.ScrollBars = 'Both'
    $box.WordWrap = $false
    $box.BackColor = [System.Drawing.SystemColors]::Window
    $box.Size = New-Object System.Drawing.Size(860, 360)
    $box.Text = $sb.ToString()
    $panel.Controls.Add($box)

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.AutoSize = $true
    $buttons.FlowDirection = 'RightToLeft'
    $buttons.Width = 860
    $closeBtn = New-Button 'Close'
    $closeBtn.DialogResult = 'Cancel'
    $againBtn = New-Button 'Rename another folder'
    $againBtn.DialogResult = 'Retry'
    $openText = 'Open the folder'
    if ($Mode -eq 'Copy') { $openText = 'Open the "renamed" folder' }
    $openBtn = New-Button $openText
    $openBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $openBtn.Add_Click({ Start-Process -FilePath $outFolder })
    $openBtn.Enabled = (Test-Path -LiteralPath $outFolder -PathType Container)
    $buttons.Controls.Add($closeBtn)
    $buttons.Controls.Add($againBtn)
    if ($Mode -eq 'Rename' -and $done.Count -gt 0) {
        $undoBtn = New-Button 'Put the old names back'
        $undoBtn.Add_Click({
            $answer = Show-Message ('Put back the old names of the ' + (Get-Count $done.Count 'file' 'files') + ' that were just renamed?') 'YesNo' 'Question'
            if ($answer -eq 'Yes') { $form.DialogResult = 'Abort' }
        })
        $buttons.Controls.Add($undoBtn)
    }
    $buttons.Controls.Add($openBtn)
    $panel.Controls.Add($buttons)
    $form.Controls.Add($panel)
    $form.CancelButton = $closeBtn
    $form.Add_Shown({ $box.SelectionStart = 0; $box.SelectionLength = 0; $openBtn.Focus() })

    $result = $form.ShowDialog()
    $form.Dispose()
    if ($result -eq 'Retry') { return 'Again' }
    if ($result -eq 'Abort') { return 'Undo' }
    return 'Close'
}


# ---------------------------------------------------------------- start here

$script:AppIcon = $null
try { $script:AppIcon = New-Object System.Drawing.Icon (Join-Path $AppDir 'Document Renamer.ico') } catch {}

try {
    [DocRenamer.MessageFilter]::Register()
    $last = Get-SavedSettings
    $current = $Folder
    # After "Change folder...", Cancel - or "No" when the new folder has no files -
    # goes back to the folder before, with the files already read there
    $returnTo = $null; $returnKnown = $null
    $known = $null
    while ($true) {
        if (-not $current) {
            $current = Select-Folder $last.LastFolder
            if (-not $current -and $returnTo) { $current = $returnTo; $known = $returnKnown; $returnTo = $null }
            if (-not $current) { break }
        }
        $items = Find-Files $current -AllFiles:($last.FileSet -eq 'All')
        if ($known) { $items = Merge-KnownItems $items $known }
        $history = Get-LastRenaming $current
        $anyFile = $null -ne (Get-ChildItem -LiteralPath $current -File -ErrorAction SilentlyContinue | Select-Object -First 1)
        if (-not $anyFile -and -not $history) {
            $answer = Show-Message ("No files were found in:`n$current`n`n" +
                "Only files directly in this folder are looked at, not folders inside it.`n`nChoose a different folder?") 'YesNo' 'Question'
            $last.LastFolder = $current
            $current = $null
            $known = $null
            if ($answer -eq 'Yes') { continue }
            if ($returnTo) { $current = $returnTo; $known = $returnKnown; $returnTo = $null; continue }
            break
        }
        # Documents are read now when their text is wanted; otherwise only if the choice changes to it
        if ($last.Source -ne 'Name' -and $last.Source -ne 'Fixed') { Invoke-ReadingWithProgress $items }
        $returnTo = $null; $returnKnown = $null
        $choice = Show-PreviewForm $current $items $last $history $known
        $known = $null
        if ($null -eq $choice) { break }
        if ($choice.Action -eq 'ChangeFolder') {
            foreach ($k in @($choice.Options.Keys)) { $last[$k] = $choice.Options[$k] }
            $last.FileSet = $choice.FileSet
            $last.LastFolder = $current
            $returnTo = $current
            $returnKnown = $choice.Known
            $current = $null
            continue
        }
        $last.LastFolder = $current
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        if ($choice.Action -eq 'Undo') {
            $mode = 'Undo'
            $results = Undo-LastRenaming $current
        } else {
            $mode = $choice.Mode
            foreach ($k in @($choice.Options.Keys)) { $last[$k] = $choice.Options[$k] }
            $last.FileSet = $choice.FileSet
            $results = Invoke-Renaming -Plan $choice.Plan -Mode $mode -Folder $current
        }
        Save-Settings $last
        $next = Show-Summary $results $current $mode
        while ($next -eq 'Undo') {
            $results = Undo-LastRenaming $current
            $next = Show-Summary $results $current 'Undo'
        }
        if ($next -ne 'Again') { break }
        $current = $null
    }
} catch {
    $msg = $_.Exception.Message
    try { Add-Content -LiteralPath (Join-Path ([IO.Path]::GetTempPath()) 'DocumentRenamer-errors.txt') -Value ("{0}  {1}`r`n{2}`r`n" -f (Get-Date), $msg, $_.ScriptStackTrace) } catch {}
    Show-Message ("Sorry, something went wrong:`n`n$msg") 'OK' 'Error' | Out-Null
} finally {
    try { [DocRenamer.MessageFilter]::Revoke() } catch {}
}