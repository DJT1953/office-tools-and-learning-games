# Find and Replace.ps1
# The window for Find and Replace. Add the documents (.rtf, .doc, .docx and
# .txt), then choose what to replace in them: words or phrases everywhere they
# appear (like Find and Replace in Word), the beginning of each document up to
# some text, and the ending of each document from some text. A preview shows
# every change first. Then changed copies are saved in a "replaced" folder, or
# the documents themselves are changed (after a copy of each one is kept in
# "before replace"). The boxes also understand Word's special codes, such as ^l
# for a line break and ^p for a paragraph mark, and a "Format" button finds and
# changes how text looks, such as bold or a style. Start it with the "Find and
# Replace" shortcut. Documents dropped on the shortcut are put in the list.
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Paths)

$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Word and the window both need a single-threaded (STA) PowerShell
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $argList = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $MyInvocation.MyCommand.Path))
    foreach ($p in $Paths) { $argList += ('"{0}"' -f $p) }
    Start-Process $ps -ArgumentList $argList
    return
}

. (Join-Path $AppDir 'ReplaceEngine.ps1')
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# The modern "Select folder" window, and a helper that makes Word retry when
# it is busy instead of failing ("call was rejected by callee")
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace FindReplaceApp
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

$SettingsFile = Join-Path $env:APPDATA 'Find and Replace\settings.json'
$Defaults = @{
    LastFolder = ''
    WordsOn = $true; WordsFind = ''; WordsReplace = ''; WordsMatchCase = $false; WordsWhole = $false
    WordsFindFormat = @{}; WordsNewFormat = @{}
    BeginOn = $false; BeginFind = ''; BeginInclude = $false; BeginMatchCase = $false; BeginWhich = 'First'
    BeginSource = 'Text'; BeginReplace = ''; BeginFile = ''
    EndOn = $false; EndFind = ''; EndInclude = $false; EndMatchCase = $false; EndWhich = 'Last'
    EndSource = 'Text'; EndReplace = ''; EndFile = ''
}

function Get-SavedSettings {
    $s = @{} + $Defaults
    try {
        if (Test-Path -LiteralPath $SettingsFile) {
            $o = Get-Content -LiteralPath $SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $o.PSObject.Properties) { if ($s.ContainsKey($p.Name) -and $null -ne $p.Value) { $s[$p.Name] = $p.Value } }
            # The formats come back from the file in another form
            foreach ($k in 'WordsFindFormat', 'WordsNewFormat') {
                if ($s[$k] -is [hashtable]) { continue }
                $h = @{}
                foreach ($p in $s[$k].PSObject.Properties) { $h[$p.Name] = $p.Value }
                $s[$k] = $h
            }
        }
    } catch {}
    return $s
}

function Save-Settings([hashtable]$S) {
    try {
        $dir = Split-Path -Parent $SettingsFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        [IO.File]::WriteAllText($SettingsFile, ($S | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
    } catch {}
}


# ---------------------------------------------------------------- small helpers

function Show-Message([string]$Text, [string]$Buttons = 'OK', [string]$Icon = 'Information') {
    return [System.Windows.Forms.MessageBox]::Show($Text, 'Find and Replace', $Buttons, $Icon)
}

function Select-Folder([string]$Start) {
    $title = 'Choose a folder of documents to add'
    try { return [FindReplaceApp.FolderPicker]::Pick([IntPtr]::Zero, $title, $Start) }
    catch {
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = $title
        if ($Start -and (Test-Path -LiteralPath $Start)) { $d.SelectedPath = $Start }
        if ($d.ShowDialog() -eq 'OK') { return $d.SelectedPath }
        return $null
    }
}

# "1 document", "3 documents"
function Get-Count([int]$N, [string]$One, [string]$Many) {
    if ($N -eq 1) { return "1 $One" }
    return "$N $Many"
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
    # On a small screen the window scrolls instead of running off the bottom
    $form.Add_Load({
        try {
            $area = [System.Windows.Forms.Screen]::FromControl($this).WorkingArea
            if ($this.Height -gt $area.Height) {
                $w = $this.Width + [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
                $this.AutoSize = $false
                $this.AutoScroll = $true
                $this.Size = New-Object System.Drawing.Size($w, $area.Height)
                $this.Location = New-Object System.Drawing.Point([Math]::Max($area.Left, $this.Left - 10), $area.Top)
            }
        } catch {}
    })
    return $form
}

function New-Label([string]$Text, [switch]$Bold, [single]$Size = 10, [int]$Width = 0, [switch]$Grey) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true
    if ($Bold -or $Size -ne 10) {
        $style = [System.Drawing.FontStyle]::Regular
        if ($Bold) { $style = [System.Drawing.FontStyle]::Bold }
        $l.Font = New-Object System.Drawing.Font('Segoe UI', $Size, $style)
    }
    if ($Grey) { $l.ForeColor = [System.Drawing.Color]::DimGray }
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

# Gives the labels on the left of a grid the same width everywhere, so the boxes line up
function Set-LabelColumn($Grid) {
    $Grid.ColumnStyles[0].SizeType = 'Absolute'
    $Grid.ColumnStyles[0].Width = 132
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

function New-Button([string]$Text, [switch]$Bold) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.AutoSize = $true
    $b.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    if ($Bold) { $b.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold) }
    return $b
}

function New-Check([string]$Text, [switch]$Bold) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.AutoSize = $true
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 6, 12, 3)
    if ($Bold) { $c.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold) }
    return $c
}

function New-Radio([string]$Text, [int]$Width = 0) {
    $r = New-Object System.Windows.Forms.RadioButton
    $r.Text = $Text
    $r.AutoSize = $true
    if ($Width -gt 0) { $r.MaximumSize = New-Object System.Drawing.Size($Width, 0) }
    $r.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    return $r
}

function New-TextBox([int]$Width, [int]$Lines = 1) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Width = $Width
    $t.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    if ($Lines -gt 1) {
        $t.Multiline = $true
        $t.AcceptsReturn = $true
        $t.ScrollBars = 'Vertical'
        $t.Height = 20 * $Lines + 8
    } else { $t.MaxLength = 255 }   # Word looks for at most 255 characters
    return $t
}

# What the "Special" buttons offer, as "name|code"
$script:SpecialFind = @('Paragraph mark|^p', 'Manual line break|^l', 'Tab|^t', 'Manual page break|^m', 'Section break|^b', 'Column break|^n', '-',
    'Non-breaking space|^s', 'Non-breaking hyphen|^~', 'Optional hyphen|^-', 'Em dash|^+', 'En dash|^=', 'Caret (the ^ character)|^^', '-',
    'White space|^w', 'Any character|^?', 'Any digit|^#', 'Any letter|^$', '-', 'Picture|^g', 'Footnote mark|^f', 'Endnote mark|^e')
$script:SpecialNew = @('Paragraph mark|^p', 'Manual line break|^l', 'Tab|^t', 'Manual page break|^m', 'Column break|^n', '-',
    'Non-breaking space|^s', 'Non-breaking hyphen|^~', 'Optional hyphen|^-', 'Em dash|^+', 'En dash|^=', 'Caret (the ^ character)|^^')
$script:SpecialReplace = $script:SpecialNew + @('-', 'The text that was found|^&')

# A button like "Special" in Word's Find and Replace: it puts a code in the box
function New-SpecialButton($Box, [string[]]$Items) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = 'Special'
    $b.Width = 72
    $b.Height = 26
    $b.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 2)
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    foreach ($entry in $Items) {
        if ($entry -eq '-') { [void]$menu.Items.Add('-'); continue }
        $parts = $entry.Split('|')
        # A single & in a menu would underline the next letter instead of showing
        $item = $menu.Items.Add(($parts[0] + '     ' + $parts[1]).Replace('&', '&&'))
        $item.Tag = @{ Box = $Box; Code = $parts[1] }
        $item.Add_Click({
            $this.Tag.Box.SelectedText = $this.Tag.Code
            $this.Tag.Box.Focus()
        })
    }
    $b.Tag = $menu
    $b.Add_Click({ $this.Tag.Show($this, 0, $this.Height) })
    return $b
}

# The line under a box that says which format was chosen. It stays hidden until there is one.
function New-FormatLabel {
    $l = New-Label '' -Width 450
    $l.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 140)
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 3)
    $l.Tag = @{}
    $l.Visible = $false
    return $l
}

function Set-FormatLabel($Label, $Format) {
    $Label.Tag = $Format
    if (Test-HasFormat $Format) {
        $Label.Text = 'Format: ' + (Get-FormatText $Format)
        $Label.Visible = $true
    } else {
        $Label.Text = ''
        $Label.Visible = $false
    }
}

# A button like "Format" in Word's Find and Replace. $Kind: 'Find' or 'Replace'.
function New-FormatButton([string]$Kind, $Label) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = 'Format...'
    $b.Width = 72
    $b.Height = 26
    $b.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 2)
    $b.Tag = @{ Kind = $Kind; Label = $Label }
    $b.Add_Click({
        $t = $this.Tag
        $chosen = Show-FormatWindow $t.Kind $t.Label.Tag
        if ($null -ne $chosen) { Set-FormatLabel $t.Label $chosen }
    })
    return $b
}

function New-List([int]$Width, [int]$Height) {
    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.HeaderStyle = 'Nonclickable'
    $list.ShowItemToolTips = $true
    $list.HideSelection = $false
    $list.Size = New-Object System.Drawing.Size($Width, $Height)
    $list.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    return $list
}

# A row of buttons on the right-hand side
function New-ButtonRow([int]$Width) {
    $p = New-Object System.Windows.Forms.FlowLayoutPanel
    $p.AutoSize = $true
    $p.FlowDirection = 'RightToLeft'
    $p.WrapContents = $false
    $p.Width = $Width
    $p.MinimumSize = New-Object System.Drawing.Size($Width, 0)
    $p.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)
    foreach ($c in $args) { $p.Controls.Add($c) }
    return $p
}

# ---------------------------------------------------------------- the small window shown while working

# Runs Invoke-Documents (the preview, or with -Apply the changes) with a
# window that shows each document and has a Stop button
function Invoke-WithProgress($Items, [hashtable]$Settings, [switch]$Apply, [string]$SaveMode = 'Copy', [string]$Stamp = '') {
    if ($Apply) { $total = @($Items | Where-Object { -not $_.Problem -and $_.Preview -and $_.Preview.Changes }).Count; $verb = 'Changing' }
    else { $total = @($Items | Where-Object { -not $_.Problem }).Count; $verb = 'Looking at' }
    $form = New-Window ('Find and Replace - ' + $verb.ToLower())
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.AutoSize = $true
    $panel.FlowDirection = 'TopDown'
    $panel.WrapContents = $false
    $label = New-Label 'Getting ready...'
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(520, 48)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Size = New-Object System.Drawing.Size(520, 22)
    $bar.Maximum = [Math]::Max(1, $total)
    $stopBtn = New-Button 'Stop'
    $stopBtn.Anchor = 'Right'
    $panel.Controls.Add($label)
    $panel.Controls.Add($bar)
    $panel.Controls.Add($stopBtn)
    $form.Controls.Add($panel)

    # The button is handled while the work runs, so it talks through $script:Progress
    $script:Progress = @{ Label = $label; Bar = $bar; StopButton = $stopBtn; Stop = $false; Done = $false; Verb = $verb }
    $stopBtn.Add_Click({
        $script:Progress.Stop = $true
        $script:Progress.StopButton.Enabled = $false
        $script:Progress.Label.Text = 'Stopping after this document...'
    })
    $form.Add_FormClosing({
        param($sender, $closeArgs)
        if (-not $script:Progress.Done) {
            $closeArgs.Cancel = $true
            $script:Progress.Stop = $true
            $script:Progress.StopButton.Enabled = $false
            $script:Progress.Label.Text = 'Stopping after this document...'
        }
    })
    if ($total -gt 0) {
        $form.Show()
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Invoke-Documents -Items $Items -Settings $Settings -Apply:$Apply -SaveMode $SaveMode -Stamp $Stamp `
            -OnFile {
                param($number, $count, $item)
                $p = $script:Progress
                $p.Bar.Value = [Math]::Min($p.Bar.Maximum, $number - 1)
                if (-not $p.Stop) { $p.Label.Text = "$($p.Verb) document $number of ${count}:`r`n$($item.Name)" }
                $p.Label.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
            } `
            -OnTick { [System.Windows.Forms.Application]::DoEvents() } `
            -ShouldStop { $script:Progress.Stop }
    } finally {
        $script:Progress.Done = $true
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        $form.Close()
        $form.Dispose()
    }
}


# ---------------------------------------------------------------- the list of documents

# Adds documents, and the documents in folders, to the list. Returns what
# couldn't be added, so the window can say so.
function Add-DocPaths($List, [string[]]$Paths) {
    $known = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($lvi in $List.Items) { [void]$known.Add($lvi.Tag) }
    $out = @{ Added = 0; NotDocs = 0; EmptyFolders = @(); LastFolder = '' }
    $List.BeginUpdate()
    try {
        foreach ($p in $Paths) {
            if (-not $p) { continue }
            if (Test-Path -LiteralPath $p -PathType Container) {
                $docs = @(Get-FolderDocuments $p | Where-Object { -not [IO.Path]::GetFileName($_).StartsWith('~$') })
                if (-not $docs.Count) { $out.EmptyFolders += $p }
                foreach ($d in $docs) { if (Add-OneDoc $List $known $d) { $out.Added++ } }
                $out.LastFolder = $p
            } elseif ((Test-Path -LiteralPath $p -PathType Leaf) -and (Test-DocFile $p) -and -not [IO.Path]::GetFileName($p).StartsWith('~$')) {
                if (Add-OneDoc $List $known $p) { $out.Added++ }
                $out.LastFolder = Split-Path -Parent $p
            } else { $out.NotDocs++ }
        }
    } finally { $List.EndUpdate() }
    return $out
}

function Add-OneDoc($List, $Known, [string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $Known.Add($full)) { return $false }
    $lvi = New-Object System.Windows.Forms.ListViewItem([IO.Path]::GetFileName($full))
    [void]$lvi.SubItems.Add((Split-Path -Leaf (Split-Path -Parent $full)))
    $lvi.ToolTipText = $full
    $lvi.Tag = $full
    [void]$List.Items.Add($lvi)
    return $true
}

# Tells the user about files that weren't added
function Show-AddProblems($Out) {
    $lines = @()
    if ($Out.NotDocs) { $lines += (Get-Count $Out.NotDocs 'file was' 'files were') + ' not added: only .rtf, .doc, .docx and .txt documents can be changed.' }
    foreach ($f in $Out.EmptyFolders) { $lines += "No .rtf, .doc, .docx or .txt documents were found in:`n$f`n(Only documents directly in a folder are added, not ones in folders inside it.)" }
    if ($lines.Count) { [void](Show-Message ($lines -join "`n`n") 'OK' 'Information') }
}

# A document in the list as an item for the engine. One that has gone missing
# since it was added is skipped, with a note saying why.
function Get-ListItems($List) {
    $items = @()
    foreach ($lvi in $List.Items) {
        try { $items += New-DocItem $lvi.Tag }
        catch {
            $items += [pscustomobject]@{
                Path = $lvi.Tag; Name = $lvi.Text; Folder = (Split-Path -Parent $lvi.Tag); Type = ''
                Problem = 'skipped: the file is not there any more (it may have been moved, renamed or deleted)'; Preview = $null; Result = $null }
        }
    }
    return $items
}


# ---------------------------------------------------------------- the Format window

# Word's standard colors, with Word's number for each
$script:ColorChoices = @(@('Automatic color', -16777216), @('Black', 0), @('White', 16777215), @('Dark red', 0x0000C0), @('Red', 0x0000FF),
    @('Orange', 0x00C0FF), @('Yellow', 0x00FFFF), @('Light green', 0x50D092), @('Green', 0x50B000), @('Light blue', 0xF0B000),
    @('Blue', 0xC07000), @('Dark blue', 0x602000), @('Purple', 0xA03070))
$script:HighlightChoices = @(@('No highlight', 0), @('Yellow', 7), @('Bright green', 4), @('Turquoise', 3), @('Pink', 5), @('Blue', 2), @('Red', 6),
    @('Dark blue', 9), @('Teal', 10), @('Green', 11), @('Violet', 12), @('Dark red', 13), @('Dark yellow', 14), @('Gray 50%', 15), @('Gray 25%', 16), @('Black', 1))
$script:StyleChoices = @('Normal', 'Title', 'Subtitle', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Quote', 'Intense Quote', 'List Paragraph', 'No Spacing', 'Strong', 'Emphasis')

# A drop-down list. -Editable also lets you type in it.
function New-Choice([string[]]$Items, [int]$Width, [switch]$Editable) {
    $c = New-Object System.Windows.Forms.ComboBox
    if ($Editable) { $c.DropDownStyle = 'DropDown' } else { $c.DropDownStyle = 'DropDownList' }
    $c.Width = $Width
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    [void]$c.Items.AddRange($Items)
    $c.SelectedIndex = 0
    return $c
}

# Chooses what a format looks like. Returns the format (with nothing in it for
# "No formatting"), or $null when the window was cancelled.
function Show-FormatWindow([string]$Kind, $Current) {
    $isFind = ($Kind -eq 'Find')
    $blank = '(no change)'
    $title = 'Format to replace with'
    $intro = 'The found text gets these. Anything left at "(no change)" stays as it is. If "Replace with" is empty, the words stay and only their look changes.'
    if ($isFind) {
        $blank = '(any)'
        $title = 'Format to find'
        $intro = 'Only text with all of these is found. If "Find" is empty, all text with this format is found.'
    }
    $form = New-Window $title
    $form.StartPosition = 'CenterParent'
    $g = New-Grid 2
    $form.Controls.Add($g)
    Add-Row $g (New-Label $intro -Width 390) -Wide
    Set-LabelColumn $g

    $fonts = New-Object System.Drawing.Text.InstalledFontCollection
    $font = New-Choice (@($blank) + @($fonts.Families | ForEach-Object { $_.Name })) 250 -Editable
    $fonts.Dispose()
    $size = New-Choice @($blank, '8', '9', '10', '10.5', '11', '12', '14', '16', '18', '20', '22', '24', '26', '28', '36', '48', '72') 100 -Editable
    $bold = New-Choice @($blank, 'Bold', 'Not bold') 170
    $italic = New-Choice @($blank, 'Italic', 'Not italic') 170
    $underline = New-Choice @($blank, 'Underlined', 'Not underlined') 170
    $strike = New-Choice @($blank, 'Strikethrough', 'Not strikethrough') 170
    $colorValues = @{}
    $colorItems = @($blank)
    foreach ($c in $script:ColorChoices) { $colorItems += $c[0]; $colorValues[$c[0]] = $c[1] }
    $color = New-Choice ($colorItems + @('Other color...')) 170
    $highlightValues = @{}
    if ($isFind) { $highlight = New-Choice @($blank, 'Highlighted', 'Not highlighted') 170 }
    else {
        $highlightItems = @($blank)
        foreach ($c in $script:HighlightChoices) { $highlightItems += $c[0]; $highlightValues[$c[0]] = $c[1] }
        $highlight = New-Choice $highlightItems 170
    }
    $style = New-Choice (@($blank) + $script:StyleChoices) 250 -Editable
    Add-Row $g (New-Label 'Font:') $font
    Add-Row $g (New-Label 'Size:') $size
    Add-Row $g (New-Label 'Bold:') $bold
    Add-Row $g (New-Label 'Italic:') $italic
    Add-Row $g (New-Label 'Underline:') $underline
    Add-Row $g (New-Label 'Strikethrough:') $strike
    Add-Row $g (New-Label 'Color:') $color
    Add-Row $g (New-Label 'Highlight:') $highlight
    Add-Row $g (New-Label 'Style:') $style
    if (-not $isFind) { Add-Row $g (New-Label 'A paragraph style such as a heading changes the whole paragraph the text is in.' -Grey -Width 390) -Wide }

    $clearBtn = New-Button 'No formatting'
    $okBtn = New-Button 'OK' -Bold
    $cancelBtn = New-Button 'Cancel'
    $cancelBtn.DialogResult = 'Cancel'
    Add-Row $g (New-ButtonRow 390 $cancelBtn $okBtn $clearBtn) -Wide
    $form.AcceptButton = $okBtn
    $form.CancelButton = $cancelBtn

    # What was chosen before
    if ($Current.Font) { $font.Text = [string]$Current.Font }
    if ($Current.Size) { $size.Text = [string]$Current.Size }
    foreach ($row in @(@($bold, $Current.Bold), @($italic, $Current.Italic), @($underline, $Current.Underline), @($strike, $Current.Strike))) {
        if ($row[1] -eq 'Yes') { $row[0].SelectedIndex = 1 } elseif ($row[1] -eq 'No') { $row[0].SelectedIndex = 2 }
    }
    if ([string]$Current.Color -ne '') {
        $known = $colorValues.Keys | Where-Object { [int]$colorValues[$_] -eq [int]$Current.Color } | Select-Object -First 1
        if (-not $known) {
            $known = [string]$Current.ColorName
            $colorValues[$known] = [int]$Current.Color
            $color.Items.Insert($color.Items.Count - 1, $known)
        }
        $color.SelectedItem = $known
    }
    if ($isFind) {
        if ($Current.Highlight -eq 'Yes') { $highlight.SelectedIndex = 1 } elseif ($Current.Highlight -eq 'No') { $highlight.SelectedIndex = 2 }
    } elseif ([string]$Current.Highlight -ne '') {
        $known = $highlightValues.Keys | Where-Object { [int]$highlightValues[$_] -eq [int]$Current.Highlight } | Select-Object -First 1
        if ($known) { $highlight.SelectedItem = $known }
    }
    if ($Current.Style) { $style.Text = [string]$Current.Style }

    # "Other color..." asks for a color, and goes back if none is chosen
    $state = @{ LastColor = $color.SelectedIndex }
    $color.Add_SelectedIndexChanged({
        if ($color.Text -ne 'Other color...') { $state.LastColor = $color.SelectedIndex; return }
        $d = New-Object System.Windows.Forms.ColorDialog
        $d.FullOpen = $true
        if ($d.ShowDialog($form) -eq 'OK') {
            $name = "Color $($d.Color.R),$($d.Color.G),$($d.Color.B)"
            $colorValues[$name] = [int]$d.Color.R + 256 * [int]$d.Color.G + 65536 * [int]$d.Color.B
            if (-not $color.Items.Contains($name)) { $color.Items.Insert($color.Items.Count - 1, $name) }
            $color.SelectedItem = $name
        } else { $color.SelectedIndex = $state.LastColor }
    })
    $clearBtn.Add_Click({
        foreach ($c in @($font, $size, $bold, $italic, $underline, $strike, $color, $highlight, $style)) { $c.SelectedIndex = 0 }
    })
    $okBtn.Add_Click({
        $new = @{}
        $text = $font.Text.Trim()
        if ($text -and $text -ne $blank) { $new.Font = $text }
        $text = $size.Text.Trim()
        if ($text -and $text -ne $blank) {
            $number = 0.0
            $isNumber = [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::CurrentCulture, [ref]$number)
            if (-not $isNumber -or $number -lt 1 -or $number -gt 1638) {
                [void](Show-Message 'The size must be a number of points between 1 and 1638, such as 12 or 10.5.' 'OK' 'Warning')
                return
            }
            $new.Size = ([single]$number).ToString([Globalization.CultureInfo]::InvariantCulture)
        }
        foreach ($row in @(@('Bold', $bold), @('Italic', $italic), @('Underline', $underline), @('Strike', $strike))) {
            if ($row[1].SelectedIndex -eq 1) { $new[$row[0]] = 'Yes' } elseif ($row[1].SelectedIndex -eq 2) { $new[$row[0]] = 'No' }
        }
        if ($color.SelectedIndex -gt 0) {
            $new.Color = [int]$colorValues[$color.Text]
            $new.ColorName = $color.Text.ToLowerInvariant()
        }
        if ($highlight.SelectedIndex -gt 0) {
            if ($isFind) {
                $new.Highlight = 'Yes'
                if ($highlight.SelectedIndex -eq 2) { $new.Highlight = 'No' }
                $new.HighlightName = $highlight.Text.ToLowerInvariant()
            } else {
                $new.Highlight = [int]$highlightValues[$highlight.Text]
                $new.HighlightName = $highlight.Text.ToLowerInvariant()
                if ($new.Highlight -ne 0) { $new.HighlightName += ' highlight' }
            }
        }
        $text = $style.Text.Trim()
        if ($text -and $text -ne $blank) { $new.Style = $text }
        $form.Tag = $new
        $form.DialogResult = 'OK'
    })

    $result = $form.ShowDialog()
    $chosen = $form.Tag
    $form.Dispose()
    if ($result -eq 'OK') { return $chosen }
    return $null
}


# ---------------------------------------------------------------- the main window

# The controls for "The beginning" or "The ending"
function New-CutSection($Column, [string]$Title, [string]$FindLabel, [string]$Part) {
    $c = @{ Part = $Part }
    $c.On = New-Check $Title -Bold
    $c.On.Margin = New-Object System.Windows.Forms.Padding(3, 12, 3, 0)
    Add-Row $Column $c.On
    $g = New-Grid 2
    $g.Margin = New-Object System.Windows.Forms.Padding(22, 0, 0, 0)
    Set-LabelColumn $g
    $c.Find = New-TextBox 372
    Add-Row $g (New-Label $FindLabel) (New-Strip $c.Find (New-SpecialButton $c.Find $script:SpecialFind))
    $c.Include = New-Check 'and this text too'
    $c.Case = New-Check 'Match case'
    Add-Row $g $null (New-Strip $c.Include $c.Case)
    $c.Which = New-Object System.Windows.Forms.ComboBox
    $c.Which.DropDownStyle = 'DropDownList'
    $c.Which.Width = 62
    $c.Which.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
    [void]$c.Which.Items.AddRange(@('first', 'last'))
    Add-Row $g $null (New-Strip (New-Label 'If the text is there more than once, use the') $c.Which (New-Label 'one'))
    $c.UseText = New-Radio 'this text (leave it empty to just delete):'
    Add-Row $g (New-Label 'Replace it with:') $c.UseText
    $c.Replace = New-TextBox 372 3
    $c.ReplaceSpecial = New-SpecialButton $c.Replace $script:SpecialNew
    Add-Row $g $null (New-Strip $c.Replace $c.ReplaceSpecial)
    $c.UseFile = New-Radio 'the contents of a document:'
    $c.Choose = New-Button 'Choose...'
    $c.FileLabel = New-Label '(none chosen)' -Grey -Width 200
    $c.FileLabel.Tag = ''
    $row = New-Strip $c.UseFile $c.Choose $c.FileLabel
    Add-Row $g $null $row
    Add-Row $Column $g
    $c.Grid = $g
    return $c
}

function Set-CutSection($C, [hashtable]$Last, [string]$P) {
    $C.On.Checked = [bool]$Last["${P}On"]
    $C.Find.Text = [string]$Last["${P}Find"]
    $C.Include.Checked = [bool]$Last["${P}Include"]
    $C.Case.Checked = [bool]$Last["${P}MatchCase"]
    $C.Which.SelectedIndex = [int]($Last["${P}Which"] -eq 'Last')
    $C.Replace.Text = [string]$Last["${P}Replace"]
    Set-ChosenFile $C ([string]$Last["${P}File"])
    if ($Last["${P}Source"] -eq 'File' -and $C.FileLabel.Tag) { $C.UseFile.Checked = $true } else { $C.UseText.Checked = $true }
}

function Set-ChosenFile($C, [string]$Path) {
    if ($Path -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $C.FileLabel.Text = [IO.Path]::GetFileName($Path)
        $C.FileLabel.Tag = $Path
        $C.FileLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:Tips.SetToolTip($C.FileLabel, $Path)
    } else {
        $C.FileLabel.Text = '(none chosen)'
        $C.FileLabel.Tag = ''
        $C.FileLabel.ForeColor = [System.Drawing.Color]::DimGray
    }
}

function Get-CutSettings($C) {
    $which = 'First'; if ($C.Which.SelectedIndex -eq 1) { $which = 'Last' }
    $source = 'Text'; if ($C.UseFile.Checked) { $source = 'File' }
    return @{ On = $C.On.Checked; Find = $C.Find.Text; Include = $C.Include.Checked; MatchCase = $C.Case.Checked
              Which = $which; Source = $source; Replace = $C.Replace.Text; File = [string]$C.FileLabel.Tag }
}

# What the user asked for, in the form the engine wants
function Get-UiSettings($Ui) {
    return @{
        Words = @{ On = $Ui.WordsOn.Checked; Find = $Ui.WordsFind.Text; Replace = $Ui.WordsReplace.Text
                   MatchCase = $Ui.WordsCase.Checked; WholeWords = $Ui.WordsWhole.Checked
                   FindFormat = $Ui.WordsFindFormat.Tag; NewFormat = $Ui.WordsNewFormat.Tag }
        Begin = Get-CutSettings $Ui.Begin
        End   = Get-CutSettings $Ui.End
    }
}

# Copies the settings into the remembered ones
function Update-Last([hashtable]$Last, [hashtable]$S) {
    $Last.WordsOn = $S.Words.On; $Last.WordsFind = $S.Words.Find; $Last.WordsReplace = $S.Words.Replace
    $Last.WordsMatchCase = $S.Words.MatchCase; $Last.WordsWhole = $S.Words.WholeWords
    $Last.WordsFindFormat = $S.Words.FindFormat; $Last.WordsNewFormat = $S.Words.NewFormat
    foreach ($p in 'Begin', 'End') {
        $c = $S[$p]
        $Last["${p}On"] = $c.On; $Last["${p}Find"] = $c.Find; $Last["${p}Include"] = $c.Include; $Last["${p}MatchCase"] = $c.MatchCase
        $Last["${p}Which"] = $c.Which; $Last["${p}Source"] = $c.Source; $Last["${p}Replace"] = $c.Replace; $Last["${p}File"] = $c.File
    }
}

# Says what is missing before the preview can be made, or $null when nothing is
function Get-SettingsProblem([hashtable]$S, [int]$DocCount) {
    if ($DocCount -eq 0) { return 'First add the documents to change: click "Add documents..." or "Add a folder...", or drag them into the list.' }
    if (-not ($S.Words.On -or $S.Begin.On -or $S.End.On)) { return 'Tick at least one of the three: "Words, phrases or marks", "The beginning" or "The ending".' }
    if ($S.Words.On -and -not $S.Words.Find -and -not (Test-HasFormat $S.Words.FindFormat)) { return 'Type what to find, or click "Format..." to find text by how it looks (or untick "Words, phrases or marks").' }
    if ($S.Begin.On -and -not $S.Begin.Find) { return 'Type the text that the beginning goes up to (or untick "The beginning").' }
    if ($S.End.On -and -not $S.End.Find) { return 'Type the text that the ending starts from (or untick "The ending").' }
    foreach ($p in @(@('Begin', 'beginning'), @('End', 'ending'))) {
        $c = $S[$p[0]]
        if ($c.On -and $c.Source -eq 'File' -and -not ($c.File -and (Test-Path -LiteralPath $c.File -PathType Leaf))) {
            return "Choose the document whose contents go in place of the $($p[1]) (click ""Choose..."")."
        }
    }
    # Word quietly finds nothing for a code it doesn't know, so check them here
    $boxes = New-Object System.Collections.ArrayList
    if ($S.Words.On) { [void]$boxes.Add(@($S.Words.Find, 'Find', 'Find')); [void]$boxes.Add(@($S.Words.Replace, 'Replace', 'Replace with')) }
    if ($S.Begin.On) {
        [void]$boxes.Add(@($S.Begin.Find, 'Find', 'Everything before'))
        if ($S.Begin.Source -eq 'Text') { [void]$boxes.Add(@($S.Begin.Replace, 'NewText', 'the new beginning')) }
    }
    if ($S.End.On) {
        [void]$boxes.Add(@($S.End.Find, 'Find', 'Everything after'))
        if ($S.End.Source -eq 'Text') { [void]$boxes.Add(@($S.End.Replace, 'NewText', 'the new ending')) }
    }
    foreach ($b in $boxes) {
        $bad = Get-CodeProblem $b[0] $b[1]
        if ($bad) { return """$bad"" in ""$($b[2])"" is not a special code that works there.`n`nThe ""Special"" button lists the codes that do. To use the ^ character itself, type ^^." }
    }
    return $null
}
function Show-MainForm([hashtable]$Last, [string[]]$StartPaths) {
    $form = New-Window 'Find and Replace'
    $page = New-Grid 2
    $form.Controls.Add($page)
    Add-Row $page (New-Label 'Find and replace in your documents' -Bold -Size 13) -Wide
    $left = New-Grid 1
    $left.Margin = New-Object System.Windows.Forms.Padding(0, 0, 18, 0)
    $right = New-Grid 1
    Add-Row $page $left $right

    # Left: the documents
    Add-Row $left (New-Heading '1. The documents')
    $list = New-List 400 420
    $list.AllowDrop = $true
    [void]$list.Columns.Add('Document', 240)
    [void]$list.Columns.Add('Folder', 135)
    Add-Row $left $list
    $countLabel = New-Label '' -Width 400
    Add-Row $left $countLabel
    $addBtn = New-Button 'Add documents...'
    $folderBtn = New-Button 'Add a folder...'
    Add-Row $left (New-Strip $addBtn $folderBtn)
    $removeBtn = New-Button 'Remove from the list'
    $clearBtn = New-Button 'Clear the list'
    Add-Row $left (New-Strip $removeBtn $clearBtn)
    Add-Row $left (New-Label 'You can also drag documents or folders into the list.' -Grey -Width 400)

    # Right: what to replace
    Add-Row $right (New-Heading '2. What to replace (tick one or more)')
    $ui = @{}
    $ui.WordsOn = New-Check 'Words, phrases or marks, everywhere they appear (like Find and Replace in Word)' -Bold
    $ui.WordsOn.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 0)
    Add-Row $right $ui.WordsOn
    $wg = New-Grid 2
    $wg.Margin = New-Object System.Windows.Forms.Padding(22, 0, 0, 0)
    Set-LabelColumn $wg
    $ui.WordsFind = New-TextBox 294
    $ui.WordsFindFormat = New-FormatLabel
    Add-Row $wg (New-Label 'Find:') (New-Strip $ui.WordsFind (New-SpecialButton $ui.WordsFind $script:SpecialFind) (New-FormatButton 'Find' $ui.WordsFindFormat))
    Add-Row $wg $null $ui.WordsFindFormat
    $ui.WordsReplace = New-TextBox 294
    $ui.WordsNewFormat = New-FormatLabel
    Add-Row $wg (New-Label 'Replace with:') (New-Strip $ui.WordsReplace (New-SpecialButton $ui.WordsReplace $script:SpecialReplace) (New-FormatButton 'Replace' $ui.WordsNewFormat))
    Add-Row $wg $null $ui.WordsNewFormat
    $ui.WordsCase = New-Check 'Match case'
    $ui.WordsWhole = New-Check 'Whole words only'
    Add-Row $wg $null (New-Strip $ui.WordsCase $ui.WordsWhole)
    Add-Row $right $wg
    $ui.WordsGrid = $wg
    $ui.Begin = New-CutSection $right 'The beginning of each document' 'Everything before:' 'Begin'
    $ui.End = New-CutSection $right 'The ending of each document' 'Everything after:' 'End'

    $previewBtn = New-Button 'Preview the changes...' -Bold
    $closeBtn = New-Button 'Close'
    $closeBtn.DialogResult = 'Cancel'
    Add-Row $page (New-ButtonRow 1030 $closeBtn $previewBtn) -Wide

    # What was used last time
    $ui.WordsOn.Checked = [bool]$Last.WordsOn
    $ui.WordsFind.Text = [string]$Last.WordsFind
    $ui.WordsReplace.Text = [string]$Last.WordsReplace
    $ui.WordsCase.Checked = [bool]$Last.WordsMatchCase
    $ui.WordsWhole.Checked = [bool]$Last.WordsWhole
    Set-FormatLabel $ui.WordsFindFormat $Last.WordsFindFormat
    Set-FormatLabel $ui.WordsNewFormat $Last.WordsNewFormat
    Set-CutSection $ui.Begin $Last 'Begin'
    Set-CutSection $ui.End $Last 'End'

    # Greys out what isn't ticked, and keeps the count under the list up to date
    $sync = {
        $ui.WordsGrid.Enabled = $ui.WordsOn.Checked
        foreach ($c in $ui.Begin, $ui.End) {
            $c.Grid.Enabled = $c.On.Checked
            $c.Replace.Enabled = $c.UseText.Checked
            $c.ReplaceSpecial.Enabled = $c.UseText.Checked
            $c.Choose.Enabled = $c.UseFile.Checked
        }
        $n = $list.Items.Count
        if ($n) { $countLabel.Text = (Get-Count $n 'document' 'documents') + ' in the list' }
        else { $countLabel.Text = 'No documents in the list yet.' }
        $removeBtn.Enabled = ($list.SelectedItems.Count -gt 0)
        $clearBtn.Enabled = ($n -gt 0)
    }
    foreach ($box in @($ui.WordsOn, $ui.Begin.On, $ui.End.On, $ui.Begin.UseText, $ui.Begin.UseFile, $ui.End.UseText, $ui.End.UseFile)) {
        $box.Add_CheckedChanged($sync)
    }
    $list.Add_SelectedIndexChanged($sync)

    $afterAdd = {
        param($out)
        if ($out.LastFolder) { $Last.LastFolder = $out.LastFolder }
        & $sync
        Show-AddProblems $out
    }
    $addBtn.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Title = 'Choose the documents to change'
        $d.Multiselect = $true
        $d.Filter = 'Documents (*.docx, *.doc, *.rtf, *.txt)|*.docx;*.doc;*.rtf;*.txt|All files (*.*)|*.*'
        if ($Last.LastFolder -and (Test-Path -LiteralPath $Last.LastFolder)) { $d.InitialDirectory = $Last.LastFolder }
        if ($d.ShowDialog($form) -ne 'OK') { return }
        & $afterAdd (Add-DocPaths $list $d.FileNames)
    })
    $folderBtn.Add_Click({
        $f = Select-Folder $Last.LastFolder
        if ($f) { & $afterAdd (Add-DocPaths $list @($f)) }
    })
    $list.Add_DragEnter({
        param($sender, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' }
    })
    $list.Add_DragDrop({
        param($sender, $e)
        $dropped = [string[]]$e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        & $afterAdd (Add-DocPaths $list $dropped)
    })
    $removeBtn.Add_Click({
        foreach ($lvi in @($list.SelectedItems)) { $list.Items.Remove($lvi) }
        & $sync
    })
    $list.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq 'Delete') { foreach ($lvi in @($list.SelectedItems)) { $list.Items.Remove($lvi) }; & $sync }
    })
    $clearBtn.Add_Click({ $list.Items.Clear(); & $sync })

    foreach ($c in $ui.Begin, $ui.End) {
        $c.Choose.Tag = $c
        $c.Choose.Add_Click({
            $sec = $this.Tag
            $d = New-Object System.Windows.Forms.OpenFileDialog
            if ($sec.Part -eq 'Begin') { $d.Title = 'Choose the document that becomes the new beginning' }
            else { $d.Title = 'Choose the document that becomes the new ending' }
            $d.Filter = 'Documents (*.docx, *.doc, *.rtf, *.txt)|*.docx;*.doc;*.rtf;*.txt'
            $start = $Last.LastFolder
            if ($sec.FileLabel.Tag) { $start = Split-Path -Parent $sec.FileLabel.Tag }
            if ($start -and (Test-Path -LiteralPath $start)) { $d.InitialDirectory = $start }
            if ($d.ShowDialog($form) -eq 'OK') { Set-ChosenFile $sec $d.FileName }
        })
    }

    $previewBtn.Add_Click({
        $S = Get-UiSettings $ui
        $problem = Get-SettingsProblem $S $list.Items.Count
        if ($problem) { [void](Show-Message $problem 'OK' 'Warning'); return }
        Update-Last $Last $S
        Save-Settings $Last
        $items = Get-ListItems $list
        $form.Enabled = $false
        try { Invoke-WithProgress $items $S } finally { $form.Enabled = $true }
        Show-PreviewForm $items $S
        $form.Activate()
    })
    $form.Add_FormClosing({
        Update-Last $Last (Get-UiSettings $ui)
        Save-Settings $Last
    })

    if ($StartPaths) { $out = Add-DocPaths $list $StartPaths; if ($out.LastFolder) { $Last.LastFolder = $out.LastFolder } }
    & $sync
    $form.Add_Shown({
        # So that "Special" adds to the end of remembered text, not in front of it
        foreach ($box in @($ui.WordsFind, $ui.WordsReplace, $ui.Begin.Find, $ui.Begin.Replace, $ui.End.Find, $ui.End.Replace)) { $box.SelectionStart = $box.TextLength }
        if ($list.Items.Count -eq 0) { $addBtn.Focus() } else { $previewBtn.Focus() }
    })
    [void]$form.ShowDialog()
    $form.Dispose()
}


# ---------------------------------------------------------------- the preview

# How each kind of piece of the preview looks
$script:Rich = @{
    Name  = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)); Fore = [System.Drawing.Color]::Black; Back = [System.Drawing.Color]::White }
    Title = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)); Fore = [System.Drawing.Color]::FromArgb(31, 56, 100); Back = [System.Drawing.Color]::White }
    Note  = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Italic)); Fore = [System.Drawing.Color]::FromArgb(70, 70, 70); Back = [System.Drawing.Color]::White }
    Text  = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 10)); Fore = [System.Drawing.Color]::FromArgb(110, 110, 110); Back = [System.Drawing.Color]::White }
    Old   = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Strikeout)); Fore = [System.Drawing.Color]::FromArgb(165, 0, 0); Back = [System.Drawing.Color]::FromArgb(255, 226, 226) }
    New   = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 10)); Fore = [System.Drawing.Color]::FromArgb(0, 100, 0); Back = [System.Drawing.Color]::FromArgb(212, 242, 212) }
    Format = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 10)); Fore = [System.Drawing.Color]::FromArgb(0, 70, 140); Back = [System.Drawing.Color]::FromArgb(218, 234, 252) }
    Gap   = @{ Font = (New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Italic)); Fore = [System.Drawing.Color]::FromArgb(110, 110, 110); Back = [System.Drawing.Color]::FromArgb(236, 236, 236) }
}

function Add-RichText($Box, [string]$Text, [string]$Kind) {
    $st = $script:Rich[$Kind]
    if (-not $st) { $st = $script:Rich.Text }
    $Box.SelectionStart = $Box.TextLength
    $Box.SelectionLength = 0
    $Box.SelectionFont = $st.Font
    $Box.SelectionColor = $st.Fore
    $Box.SelectionBackColor = $st.Back
    $Box.SelectedText = $Text
}

# Shows one document's changes: red and crossed out is taken out, green is put in
function Show-Segments($Box, $Item) {
    $Box.Clear()
    $p = $Item.Preview
    Add-RichText $Box ($Item.Name + "`n") 'Name'
    if (-not $p) { return }
    if ($p.Status -eq 'Error') { Add-RichText $Box ("`nThis document could not be read: " + $p.Summary + '.') 'Note' }
    elseif ($p.Status -eq 'Skipped') { Add-RichText $Box ("`nThis document is " + $p.Summary + '.') 'Note' }
    else {
        if (-not $p.Changes) { Add-RichText $Box "Nothing in this document will change.`n" 'Note' }
        foreach ($seg in $p.Segments) {
            $text = $seg.Text
            if ($seg.Kind -eq 'Title') { $text = "`n" + $text + "`n" }
            Add-RichText $Box $text $seg.Kind
        }
    }
    $Box.SelectionStart = 0
    $Box.SelectionLength = 0
    $Box.ScrollToCaret()
}
function Show-PreviewForm($Items, [hashtable]$S) {
    $form = New-Window 'Find and Replace - check the changes'
    $page = New-Grid 1
    $form.Controls.Add($page)
    $changing = @($Items | Where-Object { $_.Preview -and $_.Preview.Changes })
    $bad = @($Items | Where-Object { $_.Preview -and $_.Preview.Status -eq 'Error' })
    $all = @($Items).Count

    Add-Row $page (New-Label 'Check the changes before they are made' -Bold -Size 13)
    if ($changing.Count) {
        $text = "$($changing.Count) of the $(Get-Count $all 'document' 'documents') will change. Nothing has been changed yet."
        if ($all -eq 1) { $text = 'The document will change. Nothing has been changed yet.' }
    } else {
        $text = 'Nothing was found to change. Click "Back" to change what to look for.'
    }
    if ($bad.Count) { $text += ' ' + (Get-Count $bad.Count 'document' 'documents') + ' could not be read (see the list).' }
    Add-Row $page (New-Label $text -Width 1040)

    $list = New-List 1040 210
    [void]$list.Columns.Add('Document', 270)
    [void]$list.Columns.Add('What will change', 745)
    foreach ($it in $Items) {
        $p = $it.Preview
        $lvi = New-Object System.Windows.Forms.ListViewItem($it.Name)
        $what = ''
        if (-not $p) { $what = 'not looked at'; $lvi.ForeColor = [System.Drawing.Color]::DimGray }
        elseif ($p.Status -eq 'Error') { $what = 'could not be read: ' + $p.Summary; $lvi.ForeColor = [System.Drawing.Color]::DarkRed }
        elseif ($p.Status -eq 'Skipped') { $what = $p.Summary; $lvi.ForeColor = [System.Drawing.Color]::DimGray }
        elseif (-not $p.Changes) { $what = 'nothing to change (' + $p.Summary + ')'; $lvi.ForeColor = [System.Drawing.Color]::DimGray }
        else { $what = $p.Summary }
        [void]$lvi.SubItems.Add($what)
        $lvi.ToolTipText = $it.Path + "`n" + $what
        $lvi.Tag = $it
        [void]$list.Items.Add($lvi)
    }
    Add-Row $page $list

    Add-Row $page (New-Label 'The selected document:   red and crossed out is taken out,   green is put in,   blue gets the format in [ ],   grey stays as it is.' -Grey -Width 1040)
    $box = New-Object System.Windows.Forms.RichTextBox
    $box.ReadOnly = $true
    $box.BackColor = [System.Drawing.Color]::White
    $box.DetectUrls = $false
    $box.ScrollBars = 'Vertical'
    $box.Size = New-Object System.Drawing.Size(1040, 300)
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    Add-Row $page $box

    Add-Row $page (New-Heading 'Where to save the changed documents')
    $copyRadio = New-Radio 'Save changed copies in a folder called "replaced" next to each document. Your documents stay exactly as they are.' 1040
    $inPlaceRadio = New-Radio 'Change the documents themselves. A copy of each one as it is now is kept first, in a folder called "before replace" next to it.' 1040
    $copyRadio.Checked = $true
    Add-Row $page $copyRadio
    Add-Row $page $inPlaceRadio

    $goBtn = New-Button 'Save the changed copies' -Bold
    $backBtn = New-Button 'Back'
    $backBtn.DialogResult = 'Cancel'
    Add-Row $page (New-ButtonRow 1046 $backBtn $goBtn)
    $form.CancelButton = $backBtn
    $goBtn.Enabled = ($changing.Count -gt 0)
    $copyRadio.Enabled = $goBtn.Enabled
    $inPlaceRadio.Enabled = $goBtn.Enabled
    $setGoText = {
        if ($inPlaceRadio.Checked) { $goBtn.Text = 'Change the documents' } else { $goBtn.Text = 'Save the changed copies' }
    }
    $copyRadio.Add_CheckedChanged($setGoText)
    $inPlaceRadio.Add_CheckedChanged($setGoText)

    $list.Add_SelectedIndexChanged({
        if ($list.SelectedItems.Count) { Show-Segments $box $list.SelectedItems[0].Tag }
    })
    $form.Add_Shown({
        # Start with the first document that will change
        $pick = 0
        for ($i = 0; $i -lt $list.Items.Count; $i++) { if ($list.Items[$i].Tag.Preview -and $list.Items[$i].Tag.Preview.Changes) { $pick = $i; break } }
        if ($list.Items.Count) { $list.Items[$pick].Selected = $true; $list.Items[$pick].Focused = $true; $list.EnsureVisible($pick) }
        if ($goBtn.Enabled) { $goBtn.Focus() } else { $backBtn.Focus() }
    })

    $goBtn.Add_Click({
        $mode = 'Copy'
        if ($inPlaceRadio.Checked) { $mode = 'InPlace' }
        if ($mode -eq 'InPlace') {
            $answer = Show-Message ('Change ' + (Get-Count $changing.Count 'document' 'documents') + " themselves?`n`n" +
                'A copy of each one as it is now is kept first, in a folder called "before replace" next to it, ' +
                'and you can put them back afterwards.') 'YesNo' 'Question'
            if ($answer -ne 'Yes') { return }
        }
        $stamp = (Get-Date).ToString('yyyy-MM-dd h.mm.ss tt', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
        $form.Enabled = $false
        try { Invoke-WithProgress $Items $S -Apply -SaveMode $mode -Stamp $stamp } finally { $form.Enabled = $true }
        Show-SummaryForm $Items $mode $stamp
        $form.DialogResult = 'OK'
    })

    [void]$form.ShowDialog()
    $form.Dispose()
}


# ---------------------------------------------------------------- the summary

function Show-SummaryForm($Items, [string]$Mode, [string]$Stamp) {
    $done = @($Items | Where-Object { $_.Result -and $_.Result.Status -eq 'Done' })
    $none = @($Items | Where-Object { $_.Result -and $_.Result.Status -eq 'None' })
    $skip = @($Items | Where-Object { $_.Result -and $_.Result.Status -eq 'Skipped' })
    $fail = @($Items | Where-Object { $_.Result -and $_.Result.Status -eq 'Failed' })

    $form = New-Window 'Find and Replace - finished'
    $page = New-Grid 1
    $form.Controls.Add($page)
    $headline = 'Finished: ' + (Get-Count $done.Count 'document' 'documents') + ' changed'
    if ($none.Count) { $headline += ', ' + $none.Count + ' with nothing to change' }
    if ($skip.Count) { $headline += ', ' + $skip.Count + ' skipped' }
    if ($fail.Count) { $headline += ', ' + $fail.Count + ' could not be changed' }
    Add-Row $page (New-Label $headline -Bold -Size 13 -Width 1040)
    if ($Mode -eq 'Copy') { $where = 'The changed copies are in a folder called "replaced" next to each document. Your documents themselves were not changed.' }
    else { $where = "The documents themselves were changed. A copy of each one as it was before is in ""before replace\$Stamp"" next to it." }
    if ($done.Count -eq 0) { $where = 'No document was changed.' }
    Add-Row $page (New-Label $where -Width 1040)

    $list = New-List 1040 300
    [void]$list.Columns.Add('Document', 240)
    [void]$list.Columns.Add('Result', 150)
    [void]$list.Columns.Add('Details', 625)
    $words = @{ Done = 'changed'; None = 'nothing to change'; Skipped = 'skipped'; Failed = 'could not be changed' }
    foreach ($it in $Items) {
        $r = $it.Result
        if (-not $r) { continue }
        $lvi = New-Object System.Windows.Forms.ListViewItem($it.Name)
        [void]$lvi.SubItems.Add($words[$r.Status])
        [void]$lvi.SubItems.Add($r.Detail)
        $tip = $it.Path + "`n" + $r.Detail
        if ($r.SavedTo -and $r.SavedTo -ne $it.Path) { $tip += "`nSaved as: " + $r.SavedTo }
        $lvi.ToolTipText = $tip
        if ($r.Status -eq 'Failed') { $lvi.ForeColor = [System.Drawing.Color]::DarkRed }
        elseif ($r.Status -ne 'Done') { $lvi.ForeColor = [System.Drawing.Color]::DimGray }
        [void]$list.Items.Add($lvi)
    }
    Add-Row $page $list

    # The folder to open: where the first changed document was saved
    $openFolder = ''
    if ($done.Count) { $openFolder = Split-Path -Parent $done[0].Result.SavedTo }
    elseif (@($Items).Count) { $openFolder = @($Items)[0].Folder }
    $openText = 'Open the folder'
    if ($Mode -eq 'Copy' -and $done.Count) { $openText = 'Open the "replaced" folder' }
    $openBtn = New-Button $openText -Bold
    $openBtn.Enabled = ($openFolder -and (Test-Path -LiteralPath $openFolder -PathType Container))
    $openBtn.Add_Click({ Start-Process -FilePath $openFolder })
    $closeBtn = New-Button 'Close'
    $closeBtn.DialogResult = 'Cancel'
    $buttons = @($closeBtn)
    if ($Mode -eq 'InPlace' -and $done.Count) {
        $undoBtn = New-Button 'Put the documents back as they were'
        $undoBtn.Add_Click({
            $answer = Show-Message ('Put back the ' + (Get-Count $done.Count 'document' 'documents') + ' that were just changed, as they were before?') 'YesNo' 'Question'
            if ($answer -ne 'Yes') { return }
            $lines = Restore-Originals $Items
            $ok = @($lines | Where-Object { $_.Ok })
            $bad = @($lines | Where-Object { -not $_.Ok })
            $msg = (Get-Count $ok.Count 'document was' 'documents were') + ' put back as they were.'
            if ($bad.Count) { $msg += "`n`nThese could not be put back:`n" + (($bad | ForEach-Object { "   $($_.Name): $($_.Detail)" }) -join "`n") }
            $msg += "`n`nThe copies are still in the ""before replace"" folder."
            [void](Show-Message $msg 'OK' 'Information')
            $undoBtn.Enabled = $false
        })
        $buttons += $undoBtn
    }
    $buttons += $openBtn
    Add-Row $page (New-ButtonRow 1046 @buttons)
    $form.CancelButton = $closeBtn
    $form.Add_Shown({ if ($openBtn.Enabled) { $openBtn.Focus() } else { $closeBtn.Focus() } })
    [void]$form.ShowDialog()
    $form.Dispose()
}


# ---------------------------------------------------------------- start here

$script:AppIcon = $null
try { $script:AppIcon = New-Object System.Drawing.Icon (Join-Path $AppDir 'Find and Replace.ico') } catch {}
$script:Tips = New-Object System.Windows.Forms.ToolTip

try {
    [FindReplaceApp.MessageFilter]::Register()
    $last = Get-SavedSettings
    Show-MainForm $last $Paths
} catch {
    $msg = $_.Exception.Message
    try { Add-Content -LiteralPath (Join-Path ([IO.Path]::GetTempPath()) 'FindAndReplace-errors.txt') -Value ("{0}  {1}`r`n{2}`r`n" -f (Get-Date), $msg, $_.ScriptStackTrace) } catch {}
    Show-Message ("Sorry, something went wrong:`n`n$msg") 'OK' 'Error' | Out-Null
} finally {
    try { [FindReplaceApp.MessageFilter]::Revoke() } catch {}
}