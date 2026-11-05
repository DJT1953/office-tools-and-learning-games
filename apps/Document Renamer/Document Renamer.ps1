# Document Renamer.ps1
# The window for Document Renamer. Pick a folder, and each .rtf, .doc, .docx
# and .txt file in it gets a new name made from the first 30 characters of
# its text. A list shows every old and new name first. Then either renamed
# copies are saved in a "renamed" folder inside it (the originals stay as
# they are), or the files themselves are renamed, and the old names can be
# put back afterwards. Start it with the "Document Renamer" shortcut.
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

# The modern "Select folder" window, and a helper that makes Word retry when
# it is busy instead of failing ("call was rejected by callee")
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
        void Advise(IntPtr pfde, out uint pdwCookie);
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

    public static class FolderPicker
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(string path, IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem item);

        // Returns the chosen folder, or null if the window was cancelled
        public static string Pick(IntPtr owner, string title, string startFolder)
        {
            IFileDialog dialog = (IFileDialog)new FileOpenDialogClass();
            try
            {
                uint options;
                dialog.GetOptions(out options);
                dialog.SetOptions(options | 0x20 | 0x40);   // pick folders, real folders only
                dialog.SetTitle(title);
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
                if (dialog.Show(owner) != 0) return null;
                IShellItem result;
                dialog.GetResult(out result);
                string path;
                result.GetDisplayName(0x80058000, out path);   // the folder's full path
                return path;
            }
            finally { Marshal.ReleaseComObject(dialog); }
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
$Defaults = @{ Length = 30; LastFolder = '' }

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
    $title = 'Choose the folder with the documents to rename'
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


# ---------------------------------------------------------------- reading the documents

function Invoke-ReadingWithProgress($Items) {
    $total = @($Items | Where-Object { -not $_.Problem }).Count
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


# ---------------------------------------------------------------- the list of old and new names

function Show-PreviewForm([string]$Folder, $Items, [hashtable]$Last, $History) {
    $form = New-Window 'Document Renamer'

    # The window, from top to bottom: the folder and the name length, the list
    # of old and new names, what to do, then the buttons
    $page = New-Grid 1
    $form.Controls.Add($page)
    $top = New-Grid 2
    $bottom = New-Grid 1
    Add-Row $page $top

    Add-Row $top (New-Label 'Rename the documents in this folder' -Bold -Size 12) -Wide
    $changeBtn = New-Button 'Change folder...'
    Add-Row $top (New-Label 'Folder:') (New-Strip (New-Label $Folder -Width 760) $changeBtn)

    $lengthBox = New-Object System.Windows.Forms.NumericUpDown
    $lengthBox.Minimum = 5
    $lengthBox.Maximum = 150
    $lengthBox.Width = 64
    $lengthBox.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    $lengthBox.Value = [Math]::Min(150, [Math]::Max(5, [int]$Last.Length))
    Add-Row $top (New-Label 'New names:') (New-Strip (New-Label 'the first') $lengthBox (New-Label 'characters of the text in each document'))

    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.HeaderStyle = 'Nonclickable'
    $list.ShowItemToolTips = $true
    $list.Size = New-Object System.Drawing.Size(1010, 330)
    $list.Margin = New-Object System.Windows.Forms.Padding(3, 10, 3, 3)
    [void]$list.Columns.Add('Name now', 300)
    [void]$list.Columns.Add('New name', 330)
    [void]$list.Columns.Add('Note', 355)
    Add-Row $page $list
    $countLabel = New-Label '' -Width 1000
    Add-Row $page $countLabel
    Add-Row $page $bottom

    # --- what to do
    Add-Row $bottom (New-Heading 'What to do')
    $copyRadio = New-Object System.Windows.Forms.RadioButton
    $copyRadio.Text = 'Save renamed copies in a folder called "renamed" inside this folder (your files stay as they are)'
    $copyRadio.AutoSize = $true
    $copyRadio.Checked = $true
    $renameRadio = New-Object System.Windows.Forms.RadioButton
    $renameRadio.Text = 'Rename the files themselves (you can put the old names back afterwards)'
    $renameRadio.AutoSize = $true
    $radioPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $radioPanel.AutoSize = $true
    $radioPanel.FlowDirection = 'TopDown'
    $radioPanel.WrapContents = $false
    $radioPanel.Controls.Add($copyRadio)
    $radioPanel.Controls.Add($renameRadio)
    Add-Row $bottom $radioPanel
    $modeNote = New-Label '' -Width 1000
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
    # No AcceptButton: pressing Enter after typing a number must not start the renaming
    $form.CancelButton = $cancelBtn

    $existing = @()
    try { $existing = @(Get-ChildItem -LiteralPath $Folder -Force | ForEach-Object { $_.Name }) } catch {}
    $state = @{ Plan = $null; Mode = 'Copy'; Length = 30 }

    # The number as typed so far, so the list follows along while typing
    $readLength = {
        $v = 0
        if ([int]::TryParse($lengthBox.Text, [ref]$v) -and $v -ge $lengthBox.Minimum -and $v -le $lengthBox.Maximum) { return $v }
        return [int]$lengthBox.Value
    }

    $refresh = {
        $len = & $readLength
        $mode = 'Copy'
        if ($renameRadio.Checked) { $mode = 'Rename' }
        $plan = Get-RenamePlan -Items $Items -Length $len -Mode $mode -ExistingNames $existing
        $state.Plan = $plan; $state.Mode = $mode; $state.Length = $len
        $list.BeginUpdate()
        $list.Items.Clear()
        foreach ($row in $plan) {
            $lvi = New-Object System.Windows.Forms.ListViewItem($row.Item.Name)
            $note = $row.Note
            if ($row.Action -eq 'Keep') {
                if ($mode -eq 'Copy') { [void]$lvi.SubItems.Add('(not copied)') } else { [void]$lvi.SubItems.Add('(keeps its name)') }
                if ($row.Item.ProblemKind -eq 'Error') { $note = 'could not be read: ' + $note; $lvi.ForeColor = [System.Drawing.Color]::DarkRed }
                else { $lvi.ForeColor = [System.Drawing.Color]::DimGray }
            } else {
                [void]$lvi.SubItems.Add($row.NewName)
                if ($row.Action -eq 'Same') { $lvi.ForeColor = [System.Drawing.Color]::DimGray }
            }
            if ($row.Item.Note) {
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
        if ($mode -eq 'Copy') {
            $text = (Get-Count $go 'renamed copy' 'renamed copies') + ' will be saved.'
            if ($keep) { $text += ' ' + (Get-Count $keep 'document' 'documents') + ' will not be copied (the notes say why).' }
            $goBtn.Text = 'Save renamed copies'
            $modeNote.Text = 'The copies go into a folder called "renamed" inside this folder. Your original files are not changed. ' +
                'If "renamed" already has a copy with the same name from an earlier run, it is replaced.'
        } else {
            $text = (Get-Count $go 'file' 'files') + ' will get a new name.'
            if ($same) { $text += ' ' + (Get-Count $same 'file already has its' 'files already have their') + ' new name.' }
            if ($keep) { $text += ' ' + (Get-Count $keep 'file keeps its' 'files keep their') + ' name (the notes say why).' }
            $goBtn.Text = 'Rename the files'
            $modeNote.Text = 'The files in this folder get their new names. To undo it, click "Put the old names back" when it is done, ' +
                'or open this folder in Document Renamer again later.'
        }
        $countLabel.Text = $text
        $goBtn.Enabled = ($go -gt 0)
    }

    $lengthBox.Add_ValueChanged($refresh)
    $lengthBox.Add_KeyUp($refresh)
    $copyRadio.Add_CheckedChanged({ if ($this.Checked) { & $refresh } })
    $renameRadio.Add_CheckedChanged({ if ($this.Checked) { & $refresh } })
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

    $changeBtn.Add_Click({ $form.DialogResult = 'Retry' })
    $goBtn.Add_Click({
        & $refresh   # the list as it is now, in case the number was still being typed
        $plan = $state.Plan
        $go = @($plan | Where-Object { $_.Action -eq 'Rename' }).Count
        if ($go -eq 0) { return }
        if ($state.Mode -eq 'Rename') {
            $answer = Show-Message ('Give ' + (Get-Count $go 'file' 'files') + " in this folder a new name?`n`n$Folder`n`n" +
                'You can put the old names back afterwards.') 'YesNo' 'Question'
            if ($answer -ne 'Yes') { return }
        }
        $form.Tag = @{ Action = 'Rename'; Mode = $state.Mode; Length = $state.Length; Plan = $plan }
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
    $form.Dispose()
    if ($result -eq 'OK') { return $choice }
    if ($result -eq 'Retry') { return 'ChangeFolder' }
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
    while ($true) {
        if (-not $current) {
            $current = Select-Folder $last.LastFolder
            if (-not $current) { break }
        }
        $items = Find-Documents $current
        $history = Get-LastRenaming $current
        if (@($items | Where-Object { -not $_.Problem }).Count -eq 0 -and -not $history) {
            $answer = Show-Message ("No .rtf, .doc, .docx or .txt documents were found in:`n$current`n`n" +
                "Only files directly in this folder are looked at, not folders inside it.`n`nChoose a different folder?") 'YesNo' 'Question'
            $last.LastFolder = $current
            $current = $null
            if ($answer -eq 'Yes') { continue } else { break }
        }
        Invoke-ReadingWithProgress $items
        $choice = Show-PreviewForm $current $items $last $history
        if ($choice -eq 'ChangeFolder') { $last.LastFolder = $current; $current = $null; continue }
        if ($null -eq $choice) { break }
        $last.LastFolder = $current
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        if ($choice.Action -eq 'Undo') {
            $mode = 'Undo'
            $results = Undo-LastRenaming $current
        } else {
            $mode = $choice.Mode
            $last.Length = $choice.Length
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