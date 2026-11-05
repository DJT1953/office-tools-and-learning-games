# Document Reformatter.ps1
# The window for Document Reformatter. Pick a folder, choose a font, size and
# Normal style, and every .rtf, .doc, .docx and .txt file in that folder is saved as
# a reformatted .rtf copy in a "fixed" folder inside it. The originals are
# never changed. Start it with the "Document Reformatter" shortcut.
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

. (Join-Path $AppDir 'ReformatEngine.ps1')
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# The modern "Select folder" window, and a helper that makes Word retry when
# it is busy instead of failing ("call was rejected by callee")
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace DocReformatter
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

$SettingsFile = Join-Path $env:APPDATA 'Document Reformatter\settings.json'
$Defaults = @{ FontName = 'Calibri'; FontSize = 12; Headings = 'Keep'; LineSpacing = 'Leave'; SpaceBetween = 'Leave'
               Indent = 'Leave'; Align = 'Leave'; TextColour = 'Leave'; RemoveHighlight = $false; ClearEmphasis = $false
               EvenSpacing = $false; RemoveEmptyLines = $false; WidowControl = $false; RemoveLinks = $true
               PaperSize = 'Leave'; Margins = 'Leave'; LastFolder = '' }

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
    return [System.Windows.Forms.MessageBox]::Show($Text, 'Document Reformatter', $Buttons, $Icon)
}

function Select-Folder([string]$Start) {
    $title = 'Choose the folder with the documents to reformat'
    try { return [DocReformatter.FolderPicker]::Pick([IntPtr]::Zero, $title, $Start) }
    catch {
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = $title
        if ($Start -and (Test-Path -LiteralPath $Start)) { $d.SelectedPath = $Start }
        if ($d.ShowDialog() -eq 'OK') { return $d.SelectedPath }
        return $null
    }
}

function Get-FontNames {
    $names = New-Object 'System.Collections.Generic.SortedSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $previewable = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($fam in (New-Object System.Drawing.Text.InstalledFontCollection).Families) {
        if ($fam.Name -and -not $fam.Name.StartsWith('@')) { [void]$names.Add($fam.Name); [void]$previewable.Add($fam.Name) }
    }
    # Fonts Office downloads by itself (such as Aptos) live here, not in Windows' own list
    $cloud = Join-Path $env:LOCALAPPDATA 'Microsoft\FontCache\4\CloudFonts'
    if (Test-Path -LiteralPath $cloud) { foreach ($d in Get-ChildItem -LiteralPath $cloud -Directory) { [void]$names.Add($d.Name) } }
    return @{ All = @($names); Previewable = $previewable }
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

# A drop-down list. $Choices holds pairs of (setting value, text shown).
function New-Choice([object[]]$Choices, [string]$Selected, [int]$Width = 330) {
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'
    $c.Width = $Width
    $keys = @()
    foreach ($pair in $Choices) { [void]$c.Items.Add($pair[1]); $keys += $pair[0] }
    $c.Tag = $keys
    $i = [array]::IndexOf($keys, [string]$Selected)
    if ($i -lt 0) { $i = 0 }
    $c.SelectedIndex = $i
    return $c
}

function Get-ChoiceKey($Combo) { return [string]$Combo.Tag[$Combo.SelectedIndex] }

function New-Heading([string]$Text) {
    $l = New-Label $Text -Bold -Size 11
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 14, 3, 3)
    return $l
}

# A tick box. Text with a line break in it gets the box at the top.
function New-Tick([string]$Text, [bool]$Checked) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.AutoSize = $true
    $c.Checked = $Checked
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 2)
    if ($Text.Contains("`n")) { $c.CheckAlign = 'TopLeft'; $c.TextAlign = 'TopLeft' }
    return $c
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

function New-Button([string]$Text) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.AutoSize = $true
    $b.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    return $b
}


# ---------------------------------------------------------------- the settings window

function Show-SettingsForm([string]$Folder, $Items, [hashtable]$Last, [hashtable]$Fonts) {
    $todo = @($Items | Where-Object { -not $_.SkipReason })
    $skipped = @($Items | Where-Object { $_.SkipReason })

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Document Reformatter'
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding(14)
    if ($script:AppIcon) { $form.Icon = $script:AppIcon }

    # The window, from top to bottom: the folder, the settings in two columns
    # side by side, then the buttons
    $page = New-Grid 1
    $form.Controls.Add($page)
    $top = New-Grid 2
    $sides = New-Grid 2
    $left = New-Grid 2
    $right = New-Grid 2
    $bottom = New-Grid 1
    $left.Margin = New-Object System.Windows.Forms.Padding(0, 0, 36, 0)
    Add-Row $sides $left $right
    foreach ($part in $top, $sides, $bottom) { Add-Row $page $part }

    # --- the folder and what was found in it
    Add-Row $top (New-Label 'Reformat the documents in this folder' -Bold -Size 12) -Wide
    $folderPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $folderPanel.AutoSize = $true
    $folderPanel.WrapContents = $false
    $folderPanel.Margin = New-Object System.Windows.Forms.Padding(0)
    $folderPanel.Controls.Add((New-Label $Folder -Width 700))
    $changeBtn = New-Button 'Change folder...'
    $folderPanel.Controls.Add($changeBtn)
    Add-Row $top (New-Label 'Folder:') $folderPanel

    $counts = @()
    foreach ($t in @(@('docx', 'Word (.docx)'), @('doc', 'old Word (.doc)'), @('rtf', 'Rich Text (.rtf)'), @('txt', 'plain text (.txt)'))) {
        $n = @($todo | Where-Object { $_.Type -eq $t[0] }).Count
        if ($n -gt 0) { $counts += "$n $($t[1])" }
    }
    $found = "$($todo.Count) document"
    if ($todo.Count -ne 1) { $found += 's' }
    $found += ': ' + ($counts -join ', ')
    if ($skipped.Count -eq 1) { $found += "`n1 other file will be skipped (the summary at the end says why)." }
    elseif ($skipped.Count -gt 1) { $found += "`n$($skipped.Count) other files will be skipped (the summary at the end says why)." }
    Add-Row $top (New-Label 'Found:') (New-Label $found -Width 800)

    $list = New-Object System.Windows.Forms.ListBox
    $list.Height = 100
    $list.SelectionMode = 'None'
    foreach ($it in $Items) {
        if ($it.SkipReason) { [void]$list.Items.Add("$($it.Name)    (will be skipped)") } else { [void]$list.Items.Add($it.Name) }
    }
    Add-Row $top $list -Wide

    # --- left column: font and size
    $first = New-Heading 'Text'
    $first.Margin = New-Object System.Windows.Forms.Padding(3, 10, 3, 3)
    Add-Row $left $first -Wide
    $fontBox = New-Object System.Windows.Forms.ComboBox
    $fontBox.DropDownStyle = 'DropDown'
    $fontBox.Width = 330
    $fontBox.MaxDropDownItems = 20
    $fontBox.Items.AddRange([object[]]$Fonts.All)
    $fontBox.AutoCompleteMode = 'SuggestAppend'
    $fontBox.AutoCompleteSource = 'ListItems'
    $fontBox.Text = [string]$Last.FontName
    Add-Row $left (New-Label 'Font:') $fontBox

    $preview = New-Label ''
    $preview.AutoSize = $false
    $preview.Size = New-Object System.Drawing.Size(470, 34)
    Add-Row $left $null $preview
    $updatePreview = {
        $name = $fontBox.Text.Trim()
        $shown = $false
        if ($Fonts.Previewable.Contains($name)) {
            try {
                $preview.Font = New-Object System.Drawing.Font($name, 13)
                $preview.Text = 'The quick brown fox jumps over the lazy dog. 0123456789'
                $preview.ForeColor = [System.Drawing.SystemColors]::ControlText
                $shown = $true
            } catch {}
        }
        if (-not $shown) {
            $preview.Font = $form.Font
            if ($Fonts.All -contains $name) {
                $preview.Text = '(no preview for this font here, but Word has it)'
                $preview.ForeColor = [System.Drawing.Color]::DimGray
            } else {
                $preview.Text = '(this font is not on this PC)'
                $preview.ForeColor = [System.Drawing.Color]::DarkRed
            }
        }
    }
    $fontBox.Add_TextChanged($updatePreview)
    $fontBox.Add_SelectedIndexChanged($updatePreview)
    & $updatePreview

    $sizePanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $sizePanel.AutoSize = $true
    $sizePanel.WrapContents = $false
    $sizePanel.Margin = New-Object System.Windows.Forms.Padding(0)
    $sizeBox = New-Object System.Windows.Forms.ComboBox
    $sizeBox.DropDownStyle = 'DropDown'
    $sizeBox.Width = 80
    foreach ($v in 8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 22, 24, 28) { [void]$sizeBox.Items.Add([string]$v) }
    $sizeBox.Text = [string]$Last.FontSize
    $sizePanel.Controls.Add($sizeBox)
    $sizePanel.Controls.Add((New-Label 'points  (12 is a common size for letters and reports)'))
    Add-Row $left (New-Label 'Size:') $sizePanel

    $headPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $headPanel.AutoSize = $true
    $headPanel.FlowDirection = 'TopDown'
    $headPanel.WrapContents = $false
    $headPanel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 0)
    $keepRadio = New-Object System.Windows.Forms.RadioButton
    $keepRadio.Text = 'Keep headings bigger (they get the new font but keep their size)'
    $keepRadio.AutoSize = $true
    $sameRadio = New-Object System.Windows.Forms.RadioButton
    $sameRadio.Text = 'Make headings the same size as the rest of the text'
    $sameRadio.AutoSize = $true
    if ($Last.Headings -eq 'Same') { $sameRadio.Checked = $true } else { $keepRadio.Checked = $true }
    $headPanel.Controls.Add($keepRadio)
    $headPanel.Controls.Add($sameRadio)
    Add-Row $left (New-Label 'Headings:') $headPanel

    # --- left column: the rest of the Normal style
    Add-Row $left (New-Heading 'Normal style for regular paragraphs') -Wide
    Add-Row $left (New-Label '"Leave as it is in each file" keeps whatever each document already has.' -Width 640) -Wide
    $leave = @('Leave', 'Leave as it is in each file')
    $lineBox = New-Choice @($leave, @('1', 'Single'), @('1.15', '1.15 lines'), @('1.5', '1.5 lines'), @('2', 'Double')) $Last.LineSpacing
    Add-Row $left (New-Label 'Line spacing:') $lineBox
    $spaceBox = New-Choice @($leave, @('0', 'No space'), @('6', 'Small space (6 pt)'), @('10', 'Medium space (10 pt)'), @('12', 'Large space (12 pt)')) $Last.SpaceBetween
    Add-Row $left (New-Label 'Space between paragraphs:') $spaceBox
    $indentBox = New-Choice @($leave, @('None', 'No indent'), @('Half', 'Indent the first line (half an inch)')) $Last.Indent
    Add-Row $left (New-Label 'First line indent:') $indentBox
    $alignBox = New-Choice @($leave, @('Left', 'Line up on the left'), @('Justify', 'Justified (straight edges on both sides)')) $Last.Align
    Add-Row $left (New-Label 'Alignment:') $alignBox
    $colourBox = New-Choice @($leave, @('Black', 'Black (headings keep their color)'), @('BlackAll', 'Black, headings too')) $Last.TextColour
    Add-Row $left (New-Label 'Text color:') $colourBox

    # --- right column: clean-up and page settings
    $first = New-Heading 'Clean up'
    $first.Margin = New-Object System.Windows.Forms.Padding(3, 10, 3, 3)
    Add-Row $right $first -Wide
    $highlightBox = New-Tick 'Remove highlighting and shading behind words' ([bool]$Last.RemoveHighlight)
    Add-Row $right $highlightBox -Wide
    $emphasisBox = New-Tick "Remove bold, italic and underline added by hand`n(headings and tables keep theirs)" ([bool]$Last.ClearEmphasis)
    Add-Row $right $emphasisBox -Wide
    $spacingBox = New-Tick "Even out letter spacing`n(undo letters squeezed together or stretched apart)" ([bool]$Last.EvenSpacing)
    Add-Row $right $spacingBox -Wide
    $emptyBox = New-Tick "Remove empty lines between paragraphs`n(tip: also choose a ""Space between paragraphs"")" ([bool]$Last.RemoveEmptyLines)
    Add-Row $right $emptyBox -Wide
    $widowBox = New-Tick "Don't leave one line of a paragraph on its own`nat the top or bottom of a page" ([bool]$Last.WidowControl)
    Add-Row $right $widowBox -Wide
    $linkBox = New-Tick 'Remove hyperlinks (keep the words, remove the link)' ([bool]$Last.RemoveLinks)
    Add-Row $right $linkBox -Wide

    Add-Row $right (New-Heading 'Page') -Wide
    $paperBox = New-Choice @($leave, @('Letter', 'Letter (8.5 x 11 inches)'), @('A4', 'A4 (21 x 29.7 cm)')) $Last.PaperSize 250
    Add-Row $right (New-Label 'Paper size:') $paperBox
    $marginBox = New-Choice @($leave, @('36', '0.5 inch all round'), @('54', '0.75 inch all round'), @('72', '1 inch all round'), @('90', '1.25 inches all round')) $Last.Margins 250
    Add-Row $right (New-Label 'Margins:') $marginBox

    # --- where the copies go, and the buttons
    $wide = 1060
    try { $wide = [Math]::Max(700, $sides.GetPreferredSize([System.Drawing.Size]::Empty).Width - 6) } catch {}
    $list.Width = $wide
    $note = New-Label ('The reformatted copies are saved as .rtf files in a folder called "fixed" inside this folder. ' +
        'Your original files are not changed. If "fixed" already has copies from an earlier run, they are replaced.') -Width $wide
    $note.ForeColor = [System.Drawing.Color]::DimGray
    $note.Margin = New-Object System.Windows.Forms.Padding(3, 16, 3, 6)
    Add-Row $bottom $note

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.AutoSize = $true
    $buttons.FlowDirection = 'RightToLeft'
    $buttons.Dock = 'Fill'
    $startBtn = New-Button 'Start reformatting'
    $startBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $cancelBtn = New-Button 'Cancel'
    $cancelBtn.DialogResult = 'Cancel'
    $buttons.Controls.Add($startBtn)
    $buttons.Controls.Add($cancelBtn)
    Add-Row $bottom $buttons
    $form.AcceptButton = $startBtn
    $form.CancelButton = $cancelBtn

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

    $changeBtn.Add_Click({ $form.DialogResult = 'Retry' })
    $startBtn.Add_Click({
        $name = $fontBox.Text.Trim()
        if (-not $name) { Show-Message 'Please choose a font.' 'OK' 'Warning' | Out-Null; return }
        $size = 0.0
        $typed = $sizeBox.Text.Trim().Replace(',', '.')
        if (-not [double]::TryParse($typed, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$size) -or $size -lt 1 -or $size -gt 400) {
            Show-Message 'Please type a font size between 1 and 400, for example 12.' 'OK' 'Warning' | Out-Null
            return
        }
        if ($Fonts.All -notcontains $name) {
            $answer = Show-Message ("The font ""$name"" is not installed on this PC, so Word would show a different font instead.`n`nUse it anyway?") 'YesNo' 'Warning'
            if ($answer -ne 'Yes') { return }
        }
        $headings = 'Keep'
        if ($sameRadio.Checked) { $headings = 'Same' }
        $form.Tag = @{ FontName = $name; FontSize = $size; Headings = $headings
                       LineSpacing = (Get-ChoiceKey $lineBox); SpaceBetween = (Get-ChoiceKey $spaceBox)
                       Indent = (Get-ChoiceKey $indentBox); Align = (Get-ChoiceKey $alignBox); TextColour = (Get-ChoiceKey $colourBox)
                       RemoveHighlight = $highlightBox.Checked; ClearEmphasis = $emphasisBox.Checked; EvenSpacing = $spacingBox.Checked
                       RemoveEmptyLines = $emptyBox.Checked; WidowControl = $widowBox.Checked; RemoveLinks = $linkBox.Checked
                       PaperSize = (Get-ChoiceKey $paperBox); Margins = (Get-ChoiceKey $marginBox) }
        $form.DialogResult = 'OK'
    })

    $result = $form.ShowDialog()
    $settings = $form.Tag
    $form.Dispose()
    if ($result -eq 'OK') { return $settings }
    if ($result -eq 'Retry') { return 'ChangeFolder' }
    return $null
}

# ---------------------------------------------------------------- working, then the summary

function Invoke-WithProgress($Items, [hashtable]$Settings, [string]$OutFolder) {
    $total = @($Items | Where-Object { -not $_.SkipReason }).Count
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Document Reformatter - working'
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding(14)
    if ($script:AppIcon) { $form.Icon = $script:AppIcon }
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.AutoSize = $true
    $panel.FlowDirection = 'TopDown'
    $panel.WrapContents = $false
    $label = New-Label 'Starting Microsoft Word...'
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(520, 48)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Size = New-Object System.Drawing.Size(520, 22)
    $bar.Maximum = [Math]::Max(1, $total)
    $stopBtn = New-Button 'Stop after this file'
    $stopBtn.Anchor = 'Right'
    $panel.Controls.Add($label)
    $panel.Controls.Add($bar)
    $panel.Controls.Add($stopBtn)
    $form.Controls.Add($panel)

    # The buttons are handled while the work runs, so they talk through $script:Progress
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
        $results = Invoke-Reformat -Items $Items -Settings $Settings -OutFolder $OutFolder `
            -OnFile {
                param($number, $count, $item)
                $p = $script:Progress
                $p.Bar.Value = $number - 1
                if (-not $p.Stop) { $p.Label.Text = "Working on file $number of ${count}:`r`n$($item.Name)" }
                [System.Windows.Forms.Application]::DoEvents()
            } `
            -OnTick { [System.Windows.Forms.Application]::DoEvents() } `
            -ShouldStop { $script:Progress.Stop }
    } finally {
        $script:Progress.Done = $true
        $form.Close()
        $form.Dispose()
    }
    return ,$results
}

function Show-Summary($Results, [string]$Folder, [string]$OutFolder) {
    $conv = @($Results | Where-Object { $_.Status -eq 'Converted' })
    $skip = @($Results | Where-Object { $_.Status -eq 'Skipped' })
    $fail = @($Results | Where-Object { $_.Status -eq 'Failed' })

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Folder:  $Folder")
    [void]$sb.AppendLine("Copies saved in:  $OutFolder")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("CONVERTED ($($conv.Count))")
    if ($conv.Count -eq 0) { [void]$sb.AppendLine('   (none)') }
    foreach ($r in $conv) {
        $line = "   $($r.Name)   ->   fixed\$($r.Output)"
        if ($r.Detail) { $line += "   ($($r.Detail))" }
        [void]$sb.AppendLine($line)
    }
    foreach ($group in @(@('SKIPPED', $skip), @('COULD NOT BE CONVERTED', $fail))) {
        if ($group[1].Count -eq 0) { continue }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("$($group[0]) ($($group[1].Count))")
        foreach ($r in $group[1]) { [void]$sb.AppendLine("   $($r.Name)"); [void]$sb.AppendLine("      $($r.Detail)") }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Document Reformatter - finished'
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding(14)
    if ($script:AppIcon) { $form.Icon = $script:AppIcon }
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.AutoSize = $true
    $panel.FlowDirection = 'TopDown'
    $panel.WrapContents = $false

    $headline = "Finished: $($conv.Count) converted"
    if ($skip.Count) { $headline += ", $($skip.Count) skipped" }
    if ($fail.Count) { $headline += ", $($fail.Count) could not be converted" }
    $panel.Controls.Add((New-Label $headline -Bold -Size 12))
    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.ScrollBars = 'Vertical'
    $box.WordWrap = $true
    $box.BackColor = [System.Drawing.SystemColors]::Window
    $box.Size = New-Object System.Drawing.Size(780, 340)
    $box.Text = $sb.ToString()
    $panel.Controls.Add($box)

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.AutoSize = $true
    $buttons.FlowDirection = 'RightToLeft'
    $buttons.Width = 780
    $closeBtn = New-Button 'Close'
    $closeBtn.DialogResult = 'Cancel'
    $againBtn = New-Button 'Reformat another folder'
    $againBtn.DialogResult = 'Retry'
    $openBtn = New-Button 'Open the "fixed" folder'
    $openBtn.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $openBtn.Add_Click({ Start-Process -FilePath $OutFolder })
    $openBtn.Enabled = ($conv.Count -gt 0)
    $buttons.Controls.Add($closeBtn)
    $buttons.Controls.Add($againBtn)
    $buttons.Controls.Add($openBtn)
    $panel.Controls.Add($buttons)
    $form.Controls.Add($panel)
    $form.CancelButton = $closeBtn
    $form.Add_Shown({ $box.SelectionStart = 0; $box.SelectionLength = 0; $openBtn.Focus() })

    $result = $form.ShowDialog()
    $form.Dispose()
    if ($result -eq 'Retry') { return 'Again' }
    return 'Close'
}


# ---------------------------------------------------------------- start here

$script:AppIcon = $null
try { $script:AppIcon = New-Object System.Drawing.Icon (Join-Path $AppDir 'Document Reformatter.ico') } catch {}

try {
    [DocReformatter.MessageFilter]::Register()
    $fonts = Get-FontNames
    $last = Get-SavedSettings
    $current = $Folder
    while ($true) {
        if (-not $current) {
            $current = Select-Folder $last.LastFolder
            if (-not $current) { break }
        }
        $items = Find-Documents $current
        if (@($items | Where-Object { -not $_.SkipReason }).Count -eq 0) {
            $answer = Show-Message ("No .rtf, .doc, .docx or .txt documents were found in:`n$current`n`n" +
                "Only files directly in this folder are looked at, not folders inside it.`n`nChoose a different folder?") 'YesNo' 'Question'
            $last.LastFolder = $current
            $current = $null
            if ($answer -eq 'Yes') { continue } else { break }
        }
        $choice = Show-SettingsForm $current $items $last $fonts
        if ($choice -eq 'ChangeFolder') { $last.LastFolder = $current; $current = $null; continue }
        if ($null -eq $choice) { break }
        foreach ($k in @($choice.Keys)) { $last[$k] = $choice[$k] }
        $last.LastFolder = $current
        Save-Settings $last
        $outFolder = Join-Path $current 'fixed'
        $results = Invoke-WithProgress $items $choice $outFolder
        if ((Show-Summary $results $current $outFolder) -ne 'Again') { break }
        $current = $null
    }
} catch {
    $msg = $_.Exception.Message
    try { Add-Content -LiteralPath (Join-Path ([IO.Path]::GetTempPath()) 'DocumentReformatter-errors.txt') -Value ("{0}  {1}`r`n{2}`r`n" -f (Get-Date), $msg, $_.ScriptStackTrace) } catch {}
    Show-Message ("Sorry, something went wrong:`n`n$msg") 'OK' 'Error' | Out-Null
} finally {
    try { [DocReformatter.MessageFilter]::Revoke() } catch {}
}
