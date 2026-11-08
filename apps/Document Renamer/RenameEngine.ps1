# RenameEngine.ps1
# The part of Document Renamer that does the real work: it finds the files in
# a folder (the .rtf, .doc, .docx and .txt documents, or all files), reads the
# start of each document when its text is wanted (with Microsoft Word, or
# LibreOffice when Word can't open a file), works out the new names from the
# choices made in the window, and then saves renamed copies or renames the
# files themselves. Every renaming of the files themselves is written down,
# so that the old names can be put back. "Document Renamer.ps1" (the window)
# loads this file.

$script:DocExtensions  = @('.rtf', '.doc', '.docx', '.txt')
$script:CopyFolderName = 'renamed'
# How much of each document is read, and how much of its cleaned-up start is
# kept: plenty for a long name, even when a document starts with empty lines
$script:ReadChars      = 4000
$script:KeepChars      = 300
# Short work folder in Temp: LibreOffice crashes on very long paths
$script:WorkRoot       = Join-Path ([IO.Path]::GetTempPath()) 'DocumentRenamer'
# The record of renamings, for "Put the old names back"
$script:HistoryFile    = Join-Path $env:APPDATA 'Document Renamer\history.json'
$script:HistoryMax     = 100
# Longest new name, without the file type, so the full path stays within what Windows allows
$script:MaxNameChars   = 200

# The naming choices, as the window passes them to Get-RenamePlan:
#   Source       'Name' (the name it has now), 'Text' (the first Length characters
#                of the document's text) or 'Fixed' (FixedText, the same for every file)
#   Find, ReplaceWith, MatchCase       replace words in that name
#   RemoveFirst, RemoveLast            characters taken off its start and end
#   Case         'Leave', 'Lower', 'Upper', 'Title' (Each Word) or 'Sentence' (first letter only)
#   Prefix, Suffix                     text added before and after
#   Info         'None', 'Created', 'Modified', 'Today', 'Taken', 'Author', 'Title' or 'Folder',
#                added at InfoAt 'Start' or 'End'; dates are written as DateFormat
#   Number       $true to number the files: at NumberAt 'Start' or 'End', from NumberFrom,
#                going up by NumberStep, with NumberDigits digits, in NumberOrder
#                'Name', 'Created' or 'Modified'
#   Joiner       what goes between the parts, such as ' - '
$script:DefaultOptions = @{
    Source = 'Text'; Length = 30; FixedText = ''
    Find = ''; ReplaceWith = ''; MatchCase = $false; RemoveFirst = 0; RemoveLast = 0; Case = 'Leave'
    Prefix = ''; Suffix = ''; Info = 'None'; InfoAt = 'End'; DateFormat = 'yyyy-MM-dd'
    Number = $false; NumberAt = 'Start'; NumberFrom = 1; NumberStep = 1; NumberDigits = 3; NumberOrder = 'Name'
    Joiner = ' - ' }


# ---------------------------------------------------------------- finding files

# Each file found gets these, filled in when they are needed:
#   Skip                          why it is never renamed (Word's temporary files), or $null
#   Text, TextRead                the start of a document's words, once reading was tried
#   TextProblem, TextProblemKind  why there is no text ('Skip' or 'Error' for a reading problem)
#   Props                         author, title and date taken, once Read-FileProperties ran
function Find-Files {
    param([string]$Folder, [switch]$AllFiles)
    $files = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction Stop |
        Where-Object { $AllFiles -or $script:DocExtensions -contains $_.Extension.ToLowerInvariant() })

    $found = foreach ($f in $files) {
        $skip = $null
        if ($f.Name.StartsWith('~$')) { $skip = 'skipped: a temporary file Word makes while a document is open' }
        [pscustomobject]@{
            Path = $f.FullName; Name = $f.Name; BaseName = $f.BaseName; Extension = $f.Extension
            NameKey = (Get-NaturalKey $f.Name); Type = $f.Extension.TrimStart('.').ToLowerInvariant()
            IsDocument = ($script:DocExtensions -contains $f.Extension.ToLowerInvariant())
            Size = $f.Length; Created = $f.CreationTime; Modified = $f.LastWriteTime; Skip = $skip
            Text = $null; TextRead = $false; TextProblem = $null; TextProblemKind = $null; Note = ''; Props = $null }
    }
    $items = New-Object System.Collections.ArrayList
    foreach ($it in @($found | Sort-Object NameKey)) { [void]$items.Add($it) }
    return ,$items
}

# Files already read keep what was read: each file found that is in $Known
# (by its full path) is swapped for the one there, and the others are added to it
function Merge-KnownItems([object[]]$Found, [hashtable]$Known) {
    $merged = New-Object System.Collections.ArrayList
    foreach ($it in $Found) {
        if ($Known.ContainsKey($it.Path)) { [void]$merged.Add($Known[$it.Path]) }
        else { $Known[$it.Path] = $it; [void]$merged.Add($it) }
    }
    return ,$merged
}

# Sorts names the way File Explorer does: "Photo 2" before "Photo 10"
function Get-NaturalKey([string]$Name) {
    return [regex]::Replace($Name, '\d+', { param($m) $m.Value.PadLeft(20, '0') })
}

function Test-NeedsReading($Item) {
    return ($Item.IsDocument -and -not $Item.Skip -and -not $Item.TextRead)
}

function Set-TextProblem($Item, [string]$Problem, [string]$Kind) {
    $Item.TextProblem = $Problem
    $Item.TextProblemKind = $Kind
    $Item.TextRead = $true
}


# ---------------------------------------------------------------- reading the start of each document

function Read-DocumentStarts {
    param(
        [object[]]$Items,          # from Find-Files; each document not read yet gets its Text, or a TextProblem
        [scriptblock]$OnFile,      # & $OnFile <number> <total> <item>, before each file
        [scriptblock]$OnTick,      # called now and then so the window stays responsive
        [scriptblock]$ShouldStop   # returns $true when the user pressed Stop
    )
    $todo = @($Items | Where-Object { Test-NeedsReading $_ })
    # Word is only started when there is a Word or Rich Text file to read
    $readers = @{ Word = $null; WordTried = $false; LibreOffice = (Find-LibreOffice) }
    try {
        for ($i = 0; $i -lt $todo.Count; $i++) {
            $it = $todo[$i]
            if ($ShouldStop -and (& $ShouldStop)) {
                for ($j = $i; $j -lt $todo.Count; $j++) { Set-TextProblem $todo[$j] 'not read, because you pressed Stop' 'Skip' }
                break
            }
            if ($OnFile) { & $OnFile ($i + 1) $todo.Count $it }
            if ($it.Size -eq 0) { Set-TextProblem $it 'the file is empty (0 bytes)' 'Skip'; continue }
            try {
                if ($it.Type -eq 'txt') { $raw = Read-TextFile $it.Path }
                else {
                    $got = Read-OfficeFile $readers $it.Path $OnTick
                    $raw = $got.Text
                    $it.Note = $got.Note
                }
                $it.Text = Get-CleanText $raw
                $it.TextRead = $true
            } catch {
                Set-TextProblem $it (Get-FriendlyError (Get-InnerMessage $_.Exception)) 'Error'
            }
        }
    } finally {
        Stop-Word $readers.Word
    }
}

# Reads a Word or Rich Text file with Word. If Word can't open it, LibreOffice tries.
function Read-OfficeFile([hashtable]$Readers, [string]$Path, [scriptblock]$OnTick) {
    if (-not $Readers.WordTried) {
        $Readers.WordTried = $true
        try { $Readers.Word = Start-Word } catch {}
    }
    $why = 'NOWORD'
    if ($Readers.Word) {
        try { return @{ Text = (Read-WithWord $Readers.Word $Path); Note = '' } }
        catch {
            $why = Get-InnerMessage $_.Exception
            # If Word itself crashed, start a fresh copy for the next file
            if (-not (Test-WordAlive $Readers.Word)) { $Readers.Word = $null; try { $Readers.Word = Start-Word } catch {} }
            if ($why -match 'password') { throw 'PASSWORD:' }
        }
    }
    if (-not $Readers.LibreOffice) { throw ('OPEN:' + $why) }
    try { $text = Read-WithLibreOffice $Readers.LibreOffice $Path $OnTick }
    catch { throw ('OPEN:' + $why + ' | LibreOffice: ' + (Get-InnerMessage $_.Exception)) }
    $note = 'read with LibreOffice, because Word could not open it'
    if ($why -eq 'NOWORD') { $note = 'read with LibreOffice (Microsoft Word is not available)' }
    return @{ Text = $text; Note = $note }
}

function Read-WithWord($Word, [string]$Path) {
    $doc = Open-InWord $Word $Path
    try {
        $end = $doc.Content.End
        $text = Get-RangeText $doc ([Math]::Min($end, $script:ReadChars))
        # A document that starts with a lot of empty space: read further
        if (-not (Get-CleanText $text) -and $end -gt $script:ReadChars) { $text = Get-RangeText $doc ([Math]::Min($end, 200000)) }
        # Word opens a file it doesn't understand as plain text. A damaged file comes out as junk.
        if (@(2, 3, 4, 5, 7) -contains $doc.SaveFormat -and ((Test-ScrambledText $text) -or (Test-MostlySymbols $text))) { throw 'Word found only scrambled text.' }
        return $text
    } finally {
        Close-WordDocument $doc
    }
}

# The words in the main text (not headers, footers or text boxes), without
# hidden text, with fields showing what they show on the page, and without
# words that tracked changes deleted. (Word only accepts changes in a
# document it shows on screen, so the deleted parts are skipped instead.)
function Get-RangeText($Doc, [int]$End) {
    $deleted = New-Object System.Collections.ArrayList
    try {
        foreach ($rev in $Doc.Range(0, $End).Revisions) {
            # 2 = deleted, 14 = moved away from here
            if ($rev.Type -eq 2 -or $rev.Type -eq 14) { [void]$deleted.Add([pscustomobject]@{ Start = $rev.Range.Start; End = $rev.Range.End }) }
        }
    } catch {}
    $sb = New-Object System.Text.StringBuilder
    $pos = 0
    foreach ($d in @($deleted | Sort-Object Start)) {
        if ($d.Start -gt $pos) { [void]$sb.Append((Get-PlainRangeText $Doc $pos ([Math]::Min($d.Start, $End)))) }
        $pos = [Math]::Max($pos, $d.End)
    }
    if ($End -gt $pos) { [void]$sb.Append((Get-PlainRangeText $Doc $pos $End)) }
    return $sb.ToString()
}

function Get-PlainRangeText($Doc, [int]$Start, [int]$End) {
    $r = $Doc.Range($Start, $End)
    try { $r.TextRetrievalMode.IncludeHiddenText = $false } catch {}
    try { $r.TextRetrievalMode.IncludeFieldCodes = $false } catch {}
    return [string]$r.Text
}

# Plain text is read straight from the file. Only the start is needed.
function Read-TextFile([string]$Path) {
    $max = 65536
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $cut = ($fs.Length -gt $max)
        $bytes = New-Object byte[] ([int][Math]::Min($fs.Length, $max))
        $got = 0
        while ($got -lt $bytes.Length) {
            $n = $fs.Read($bytes, $got, $bytes.Length - $got)
            if ($n -le 0) { break }
            $got += $n
        }
    } finally { $fs.Dispose() }
    if ($got -eq 0) { $bytes = New-Object byte[] 0 }   # the file was emptied since it was listed
    elseif ($got -lt $bytes.Length) { $bytes = [byte[]]($bytes[0..($got - 1)]) }
    $text = [Text.Encoding]::GetEncoding((Get-TextEncoding $bytes $cut)).GetString($bytes)
    $text = $text.TrimStart([char]0xFEFF)
    if (Test-ScrambledText $text) { throw 'NOTTEXT:' }
    return $text
}

# Plain text files don't say which alphabet they were saved in, so look at the
# bytes. Returns the code page number. $Cut: the bytes are only the start of the file.
function Get-TextEncoding([byte[]]$Bytes, [bool]$Cut) {
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { return 65001 }   # UTF-8
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return 1200 }   # Unicode (UTF-16)
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return 1201 }
    # Unicode without a marker at the start: for English text, every other byte is zero
    $check = [Math]::Min($Bytes.Length, 4000)
    $oddZeros = 0; $evenZeros = 0
    for ($i = 0; $i -lt $check; $i++) { if ($Bytes[$i] -eq 0) { if ($i % 2) { $oddZeros++ } else { $evenZeros++ } } }
    if ($check -ge 4 -and $oddZeros * 4 -gt $check) { return 1200 }
    if ($check -ge 4 -and $evenZeros * 4 -gt $check) { return 1201 }
    # UTF-8 follows strict rules for every letter outside plain English. When only the
    # start of the file was read, its last letter may be cut in half, so leave it out.
    $n = $Bytes.Length
    if ($Cut -and $n -gt 0) {
        $p = $n - 1
        while ($p -gt 0 -and ($n - $p) -lt 4 -and ($Bytes[$p] -band 0xC0) -eq 0x80) { $p-- }
        $n = $p
    }
    try { [void](New-Object Text.UTF8Encoding($false, $true)).GetString($Bytes, 0, $n); return 65001 } catch {}
    return [Text.Encoding]::Default.CodePage   # the PC's usual Windows alphabet, such as 1252
}

# A file that isn't really text (or a damaged one that LibreOffice reads as
# plain text) comes out as scrambled symbols. That counts as "could not be read".
function Test-ScrambledText([string]$Text) {
    if ($Text.Contains([string][char]0xE0 + [char]0xA1 + [char]0xB1)) { return $true }   # the start of an old Word file, read as text
    if ($Text.Length -lt 20) { return $false }
    if ($Text.Length -gt 4000) { $Text = $Text.Substring(0, 4000) }
    $odd = [regex]::Matches($Text, '[\x00\x03\x04\x06\x0E-\x12\x16-\x1F\uFFFD]').Count
    return ($odd * 20 -gt $Text.Length)   # more than 5% junk characters
}

# When a file isn't really a document, Word and LibreOffice read it as plain
# text, with its junk turned into odd symbols. Real writing is mostly letters
# and numbers (80% or more, not counting spaces).
function Test-MostlySymbols([string]$Text) {
    if ($Text.Length -gt 4000) { $Text = $Text.Substring(0, 4000) }
    $t = [regex]::Replace($Text, '\s', '')
    if ($t.Length -lt 20) { return $false }
    $plain = [regex]::Matches($t, '[\p{L}\p{N}]').Count
    return ($plain * 10 -lt $t.Length * 7)   # less than 70%
}


# ---------------------------------------------------------------- Word itself

function Start-Word {
    try { $w = New-Object -ComObject Word.Application -ErrorAction Stop }
    catch { throw 'Microsoft Word could not be started.' }
    $w.Visible = $false
    $w.DisplayAlerts = 0                                   # never show pop-up questions
    try { $w.AutomationSecurity = 3 } catch {}             # never run macros inside the documents
    try { $w.Options.UpdateLinksAtOpen = $false } catch {}
    try { $w.Options.ConfirmConversions = $false } catch {}
    return $w
}

function Stop-Word($Word) {
    if ($null -eq $Word) { return }
    try {
        # Only close Word if nothing else is open in it - never the user's own documents
        if ($Word.Documents.Count -eq 0) { $noSave = 0; $Word.Quit([ref]$noSave) }
    } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Word) } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Test-WordAlive($Word) {
    try { $null = $Word.Documents.Count; return $true } catch { return $false }
}

function Open-InWord($Word, [string]$Path) {
    $m = [Type]::Missing
    $pw = '#no-password#'   # a wrong password makes protected files fail at once instead of asking
    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    # Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible,
    # OpenAndRepair, DocumentDirection, NoEncodingDialog
    return $Word.Documents.Open($Path, $false, $true, $false, $pw, $pw, $false, $pw, $pw, 0, $m, $false, $false, $m, $true)
}

function Close-WordDocument($Doc) {
    if ($null -eq $Doc) { return }
    try { $Doc.Close(0) } catch { try { $noSave = 0; $Doc.Close([ref]$noSave) } catch {} }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Doc) } catch {}
}

function Get-InnerMessage($Ex) {
    while ($Ex.InnerException) { $Ex = $Ex.InnerException }
    return (($Ex.Message -replace '\s+', ' ').Trim())
}

function Get-FriendlyError([string]$Message) {
    $kind = ''; $detail = $Message
    if ($Message -match '^(OPEN|PASSWORD|NOTTEXT):(.*)$') { $kind = $Matches[1]; $detail = $Matches[2].Trim() }
    switch ($kind) {
        'PASSWORD' { return 'it is protected with a password' }
        'NOTTEXT'  { return 'this .txt file does not contain readable text' }
        'OPEN' {
            if ($detail -like 'NOWORD*') { return 'Microsoft Word or LibreOffice is needed to read this type of file' }
            if ($detail -match 'RPC|remote procedure|disconnected|0x800706BA|0x80010108') { return 'Word stopped responding while reading it' }
            if ($detail -match 'File Block') { return "Word's security settings (File Block) do not allow this old type of file to be opened" }
            $text = 'Word could not open it. It may be damaged, or it may not really be a Word document'
            if ($detail -match 'LibreOffice') { $text += '. LibreOffice could not read it either' }
            return $text
        }
    }
    return 'something went wrong: ' + $detail
}


# ---------------------------------------------------------------- LibreOffice backup

function Find-LibreOffice {
    $candidates = @()
    try {
        $dir = (Get-ItemProperty 'HKLM:\SOFTWARE\LibreOffice\UNO\InstallPath' -ErrorAction Stop).'(default)'
        if ($dir) { $candidates += (Join-Path $dir 'soffice.com') }
    } catch {}
    $candidates += (Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.com')
    if (${env:ProgramFiles(x86)}) { $candidates += (Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\soffice.com') }
    foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { return $c } }
    return $null
}

# Has LibreOffice turn a file Word couldn't open into plain text; returns the text
function Read-WithLibreOffice([string]$Soffice, [string]$Path, [scriptblock]$OnTick) {
    $work = Join-Path $script:WorkRoot 'work'
    $inDir = Join-Path $work 'in'
    $outDir = Join-Path $work 'out'
    if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $inDir, $outDir | Out-Null
    $copy = Join-Path $inDir ('source' + [IO.Path]::GetExtension($Path))
    Copy-Item -LiteralPath $Path -Destination $copy

    # LibreOffice gets its own settings folder, so it works even if you have LibreOffice open
    $profileUrl = 'file:///' + ((Join-Path $script:WorkRoot 'lo-profile') -replace '\\', '/')
    $argLine = '--headless --norestore --nolockcheck "-env:UserInstallation={0}" --convert-to "txt:Text (encoded):UTF8" --outdir "{1}" "{2}"' -f $profileUrl, $outDir, $copy
    $proc = Start-Process -FilePath $Soffice -ArgumentList $argLine -WindowStyle Hidden -PassThru
    $deadline = (Get-Date).AddMinutes(3)
    while (-not $proc.WaitForExit(250)) {
        if ($OnTick) { & $OnTick }
        if ((Get-Date) -gt $deadline) { try { $proc.Kill() } catch {}; throw 'LibreOffice took too long.' }
    }
    $result = Join-Path $outDir 'source.txt'
    if (-not (Test-Path -LiteralPath $result)) { throw 'LibreOffice could not read it either.' }
    $text = [IO.File]::ReadAllText($result, [Text.Encoding]::UTF8)
    if ((Test-ScrambledText $text) -or (Test-MostlySymbols $text)) { throw 'LibreOffice found only scrambled text.' }
    return $text
}


# ---------------------------------------------------------------- author, title and date taken

# Windows reads these from inside Word documents, photos and videos (the same
# details File Explorer shows). Rich Text files keep theirs in the file's own
# text, so those are read directly. Each file gets Props: Author, Title, Taken.
function Read-FileProperties([object[]]$Items, [string]$Folder) {
    $todo = @($Items | Where-Object { $null -eq $_.Props })
    if ($todo.Count -eq 0) { return }
    $shellFolder = $null
    try { $shellFolder = (New-Object -ComObject Shell.Application).Namespace([string]$Folder) } catch {}
    foreach ($it in $todo) {
        $p = @{ Author = ''; Title = ''; Taken = $null }
        if ($shellFolder) {
            try {
                $file = $shellFolder.ParseName([string]$it.Name)
                if ($file) {
                    $p.Author = Join-PropertyText $file.ExtendedProperty('System.Author')
                    $p.Title = Join-PropertyText $file.ExtendedProperty('System.Title')
                    $taken = $file.ExtendedProperty('System.Photo.DateTaken')
                    if ($taken -isnot [datetime]) { $taken = $file.ExtendedProperty('System.Media.DateEncoded') }
                    # Windows gives these dates in world time (UTC)
                    if ($taken -is [datetime]) { $p.Taken = [datetime]::SpecifyKind($taken, [DateTimeKind]::Utc).ToLocalTime() }
                }
            } catch {}
        }
        if ($it.Type -eq 'rtf' -and -not ($p.Author -and $p.Title)) {
            try {
                $info = Get-RtfInfo $it.Path
                if (-not $p.Author) { $p.Author = $info.Author }
                if (-not $p.Title) { $p.Title = $info.Title }
            } catch {}
        }
        $it.Props = $p
    }
}

# Several authors come as a list
function Join-PropertyText($Value) {
    if ($null -eq $Value) { return '' }
    return ((@($Value) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }) -join ', ')
}

# The author and title in a Rich Text file's {\info ...} part
function Get-RtfInfo([string]$Path) {
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $bytes = New-Object byte[] ([int][Math]::Min($fs.Length, 65536))
        $got = $fs.Read($bytes, 0, $bytes.Length)
    } finally { $fs.Dispose() }
    $raw = [Text.Encoding]::GetEncoding(1252).GetString($bytes, 0, $got)
    $info = @{ Author = ''; Title = '' }
    foreach ($key in 'Author', 'Title') {
        $m = [regex]::Match($raw, '\{\\' + $key.ToLowerInvariant() + ' ?((?:[^{}\\]|\\[^{}])*)\}')
        if ($m.Success) { $info[$key] = ConvertFrom-RtfText $m.Groups[1].Value }
    }
    return $info
}

# Rich Text writes letters outside plain English as codes such as \'e9. A
# Unicode code (backslash, u, number) is followed by a stand-in letter for
# older programs - often the same letter as a \'e9 code - which is skipped.
function ConvertFrom-RtfText([string]$Text) {
    $ansi = [Text.Encoding]::GetEncoding(1252)
    $t = [regex]::Replace($Text, "\\u(-?\d+) ?(\\'[0-9a-fA-F]{2}|[^\\{}])?", { param($m) [string][char]((([int]$m.Groups[1].Value) + 65536) % 65536) })
    $t = [regex]::Replace($t, "\\'([0-9a-fA-F]{2})", { param($m) $ansi.GetString([byte[]]@([Convert]::ToByte($m.Groups[1].Value, 16))) })
    $t = [regex]::Replace($t, '(?<!\\)\\[a-zA-Z]+-?\d* ?', '')
    $t = [regex]::Replace($t, '\\([\\{}])', '$1')
    return $t.Trim()
}


# ---------------------------------------------------------------- the new names

# Turns the start of a document into one line that can be used in a file name.
# Keeps at most $script:KeepChars characters; Get-NameFromText cuts it to length.
function Get-CleanText([string]$Raw) {
    if (-not $Raw) { return '' }
    $t = $Raw
    if ($t.Length -gt 200000) { $t = $t.Substring(0, 200000) }
    $t = $t.Replace([string][char]30, '-')   # Word's non-breaking hyphen
    # Line and page breaks, tabs, table cells and unusual spaces become spaces
    $t = [regex]::Replace($t, '[\x07\x09-\x0D\x85\u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000]', ' ')
    # Leave out what can't be seen: pictures, field and footnote marks, optional
    # hyphens, direction marks, and symbols from fonts such as Wingdings
    $t = [regex]::Replace($t, '[\x00-\x1F\x7F\u00AD\u200B-\u200F\u202A-\u202E\u2060-\u2064\uFEFF\uFFFD\uE000-\uF8FF]', '')
    $t = [regex]::Replace($t, ' {2,}', ' ').Trim()
    $t = Get-SafeText $t
    $t = [regex]::Replace($t, ' {2,}', ' ').Trim()
    $t = $t.TrimStart(' ', '.', '-')           # such as a "- " typed in front of a list
    if ($t.Length -gt $script:KeepChars) { $t = $t.Substring(0, $script:KeepChars) }
    return $t
}

# Changes or leaves out the characters Windows doesn't allow in a file name
function Get-SafeText([string]$Text) {
    if (-not $Text) { return '' }
    $t = [regex]::Replace($Text, '[\x00-\x1F\x7F]', '')
    $t = [regex]::Replace($t, '[/\\|]', '-')
    $t = $t.Replace('"', "'")
    $t = [regex]::Replace($t, '[<>*?]', '')
    $t = [regex]::Replace($t, ' ?: ', ' - ')   # "Re: Meeting" becomes "Re - Meeting"
    $t = [regex]::Replace($t, ':$', '')
    $t = $t.Replace(':', '-')                  # "10:30" becomes "10-30"
    return $t
}

# The first characters of a document's cleaned-up start
function Get-NameFromText([string]$Text, [int]$Length) {
    if (-not $Text) { return '' }
    $name = Get-Shortened $Text $Length
    return $name.TrimEnd(' ', '.', '-')
}

# Text cut to at most $Length characters, without cutting an emoji in half
function Get-Shortened([string]$Text, [int]$Length) {
    if ($Text.Length -le $Length) { return $Text }
    $cut = $Length
    if ([char]::IsHighSurrogate($Text[$cut - 1])) { $cut-- }
    return $Text.Substring(0, $cut)
}

# A finished name made acceptable to Windows: it can't end with a space or a
# period, and a few names (CON, PRN, NUL, COM1, LPT1...) are kept for Windows itself
function Get-AllowedName([string]$Name) {
    $Name = $Name.TrimStart(' ').TrimEnd(' ', '.')
    if ($Name -match '^(CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[0-9\u00B9\u00B2\u00B3]|LPT[0-9\u00B9\u00B2\u00B3]) *(\.|$)') { $Name = '_' + $Name }
    return $Name
}

# The choices with every missing one filled in from $script:DefaultOptions
function Complete-Options([hashtable]$Options) {
    $o = @{} + $script:DefaultOptions
    if ($Options) { foreach ($k in @($Options.Keys)) { if ($o.ContainsKey($k) -and $null -ne $Options[$k]) { $o[$k] = $Options[$k] } } }
    foreach ($k in 'Source', 'FixedText', 'Find', 'ReplaceWith', 'Case', 'Prefix', 'Suffix', 'Info', 'InfoAt', 'DateFormat', 'NumberAt', 'NumberOrder', 'Joiner') { $o[$k] = [string]$o[$k] }
    foreach ($k in 'MatchCase', 'Number') { $o[$k] = [bool]$o[$k] }
    $o.Length = [Math]::Max(1, [int]$o.Length)
    $o.RemoveFirst = [Math]::Max(0, [int]$o.RemoveFirst)
    $o.RemoveLast = [Math]::Max(0, [int]$o.RemoveLast)
    $o.NumberFrom = [Math]::Max(0, [long]$o.NumberFrom)
    $o.NumberStep = [Math]::Max(1, [long]$o.NumberStep)
    $o.NumberDigits = [Math]::Min(10, [Math]::Max(1, [int]$o.NumberDigits))
    if (-not $o.DateFormat) { $o.DateFormat = 'yyyy-MM-dd' }
    # Typed text made safe for a file name once here, rather than for every file
    foreach ($k in 'FixedText', 'ReplaceWith', 'Prefix', 'Suffix') { $o[$k] = Get-SafeText $o[$k] }
    $o.FixedText = $o.FixedText.Trim()
    return $o
}

function Get-ReplacedText([string]$Text, [string]$Find, [string]$With, [bool]$MatchCase) {
    if (-not $Find) { return $Text }
    $how = [Text.RegularExpressions.RegexOptions]::IgnoreCase
    if ($MatchCase) { $how = [Text.RegularExpressions.RegexOptions]::None }
    return [regex]::Replace($Text, [regex]::Escape($Find), $With.Replace('$', '$$'), $how)
}

function Remove-Ends([string]$Text, [int]$First, [int]$Last) {
    if ($First -ge $Text.Length) { return '' }
    $Text = $Text.Substring($First)
    if ($Last -ge $Text.Length) { return '' }
    return $Text.Substring(0, $Text.Length - $Last)
}

function Convert-LetterCase([string]$Text, [string]$Case) {
    $culture = [Globalization.CultureInfo]::CurrentCulture
    switch ($Case) {
        'Lower' { return $Text.ToLower($culture) }
        'Upper' { return $Text.ToUpper($culture) }
        'Title' { return $culture.TextInfo.ToTitleCase($Text.ToLower($culture)) }
        'Sentence' {
            $t = $Text.ToLower($culture)
            $first = [regex]::Match($t, '\p{L}')
            if ($first.Success) { $t = $t.Substring(0, $first.Index) + $t.Substring($first.Index, 1).ToUpper($culture) + $t.Substring($first.Index + 1) }
            return $t
        }
    }
    return $Text
}

# The part of the name the other choices are added to. Returns Main, or a
# Problem (why this file gets no new name) with its Kind ('Skip' or 'Error').
function Get-MainName($Item, [hashtable]$O) {
    switch ($O.Source) {
        'Text' {
            if (-not $Item.IsDocument) { return @{ Problem = 'not a Word or text document, so it has no text to name it by'; Kind = 'Skip' } }
            if ($Item.TextProblem) { return @{ Problem = $Item.TextProblem; Kind = $Item.TextProblemKind } }
            if (-not $Item.TextRead) { return @{ Problem = 'its text has not been read'; Kind = 'Skip' } }
            $main = Get-NameFromText $Item.Text $O.Length
            if (-not $main) { return @{ Problem = 'no words found at the start of this document'; Kind = 'Skip' } }
        }
        'Fixed' { $main = $O.FixedText }
        default { $main = $Item.BaseName }
    }
    $main = Get-ReplacedText $main $O.Find $O.ReplaceWith $O.MatchCase
    $main = Remove-Ends $main $O.RemoveFirst $O.RemoveLast
    $main = Convert-LetterCase $main $O.Case
    return @{ Main = $main }
}

# The information added to a name, as text, or '' when the file doesn't have it.
# Author, title and date taken need Read-FileProperties first.
function Get-InfoText($Item, [string]$Info, [string]$DateFormat, [string]$Folder) {
    $date = $null
    switch ($Info) {
        'Created'  { $date = $Item.Created }
        'Modified' { $date = $Item.Modified }
        'Today'    { $date = Get-Date }
        'Taken'    { if ($Item.Props) { $date = $Item.Props.Taken } }
        'Author'   { if ($Item.Props) { return (Get-SafeText $Item.Props.Author).Trim() } }
        'Title'    { if ($Item.Props) { return (Get-SafeText $Item.Props.Title).Trim() } }
        'Folder'   { return (Get-SafeText (Split-Path -Leaf ([string]$Folder).TrimEnd('\'))).Trim() }
    }
    if ($null -eq $date) { return '' }
    return (Get-SafeText $date.ToString($DateFormat, [Globalization.CultureInfo]::CurrentCulture))
}

# Puts the name together: before-text, number and information at the start,
# the main name, then information, number and after-text at the end, with
# the Joiner between them. Returns Base (the name without its file type) and Notes.
function Get-JoinedName([string]$Main, $Item, [hashtable]$O, [string]$NumberText, [string]$Folder) {
    $notes = New-Object System.Collections.ArrayList
    $info = ''
    if ($O.Info -and $O.Info -ne 'None') {
        $info = Get-InfoText $Item $O.Info $O.DateFormat $Folder
        if (-not $info) {
            $what = @{ Taken = 'date taken'; Author = 'author'; Title = 'title' }[$O.Info]
            if ($what) { [void]$notes.Add("no $what found in this file, so none was added") }
        }
    }
    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add($O.Prefix)
    if ($NumberText -and $O.NumberAt -ne 'End') { [void]$parts.Add($NumberText) }
    if ($O.InfoAt -ne 'End') { [void]$parts.Add($info) }
    [void]$parts.Add($Main)
    if ($O.InfoAt -eq 'End') { [void]$parts.Add($info) }
    if ($NumberText -and $O.NumberAt -eq 'End') { [void]$parts.Add($NumberText) }
    [void]$parts.Add($O.Suffix)
    # With a joiner, spaces typed around a part would double up next to it
    $used = @(foreach ($p in $parts) { $t = [string]$p; if ($O.Joiner) { $t = $t.Trim() }; if ($t) { $t } })
    $base = Get-AllowedName ($used -join $O.Joiner)
    if ($base.Length -gt $script:MaxNameChars) {
        $base = Get-AllowedName (Get-Shortened $base $script:MaxNameChars)
        [void]$notes.Add("shortened to $($script:MaxNameChars) characters")
    }
    return @{ Base = $base; Notes = $notes }
}

# The order files are numbered in
function Get-SortedItems([object[]]$Items, [string]$Order) {
    switch ($Order) {
        'Created'  { return @($Items | Sort-Object Created, NameKey) }
        'Modified' { return @($Items | Sort-Object Modified, NameKey) }
    }
    return @($Items | Sort-Object NameKey)
}

# Works out every file's new name from the choices in $Options (see
# $script:DefaultOptions). $Mode is 'Copy' (renamed copies go into the
# "renamed" folder) or 'Rename' (the files themselves get the new names).
# $ExistingNames: the names of everything in the folder now, for 'Rename'.
# The plan lists the files in numbering order. Each row has the Item, its
# NewName, a Note, a ProblemKind for 'Keep' rows, and an Action:
#   'Rename' - gets the new name (for 'Copy': a copy with the new name is saved)
#   'Same'   - already has this name, so nothing happens (only for 'Rename')
#   'Keep'   - no new name: no words were found, the file couldn't be read...
function Get-RenamePlan {
    param([object[]]$Items, [hashtable]$Options, [string]$Mode, [string[]]$ExistingNames, [string]$Folder)
    $O = Complete-Options $Options
    if (@('Taken', 'Author', 'Title') -contains $O.Info) { Read-FileProperties $Items $Folder }

    # First, each file's name before any " (2)" is added
    $plan = New-Object System.Collections.ArrayList
    $numbered = 0
    foreach ($it in (Get-SortedItems $Items $O.NumberOrder)) {
        $row = [pscustomobject]@{ Item = $it; NewName = ''; Action = 'Keep'; Note = ''; ProblemKind = ''; Base = ''; Notes = $null }
        [void]$plan.Add($row)
        if ($it.Skip) { $row.Note = $it.Skip; $row.ProblemKind = 'Skip'; continue }
        $m = Get-MainName $it $O
        if ($m.Problem) { $row.Note = $m.Problem; $row.ProblemKind = $m.Kind; continue }
        $numberText = ''
        if ($O.Number) {
            $numberText = ($O.NumberFrom + $numbered * $O.NumberStep).ToString().PadLeft($O.NumberDigits, '0')
            $numbered++
        }
        $joined = Get-JoinedName $m.Main $it $O $numberText $Folder
        if (-not $joined.Base) { $row.Note = 'the new name would be empty'; $row.ProblemKind = 'Skip'; continue }
        $row.Base = $joined.Base
        $row.Notes = New-Object System.Collections.ArrayList
        foreach ($n in $joined.Notes) { [void]$row.Notes.Add($n) }
    }
    $named = @($plan | Where-Object { $_.Base })

    # A new name is never used twice. When renaming in place it also can't be the
    # name of anything else in the folder, except the files getting new names now:
    # their names come free, because Invoke-Renaming moves them all together.
    $taken = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if ($Mode -eq 'Rename') {
        $moving = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $named) { [void]$moving.Add($row.Item.Name) }
        foreach ($n in $ExistingNames) { if (-not $moving.Contains($n)) { [void]$taken.Add($n) } }
        # A file whose new name is the name it has keeps it first, so running the
        # same choices again changes nothing
        foreach ($row in $named) {
            $name = $row.Base + $row.Item.Extension
            if ($name -ieq $row.Item.Name -and $taken.Add($name)) { $row.NewName = $name }
        }
    }
    foreach ($row in $named) {
        if (-not $row.NewName) {
            $name = $row.Base + $row.Item.Extension
            $copyNumber = 1
            while ($taken.Contains($name)) {
                $copyNumber++
                $name = '{0} ({1}){2}' -f $row.Base, $copyNumber, $row.Item.Extension
            }
            [void]$taken.Add($name)
            $row.NewName = $name
            if ($copyNumber -gt 1) { [void]$row.Notes.Insert(0, "($copyNumber) added, because another file already has this name") }
        }
        if ($Mode -eq 'Rename' -and $row.NewName -ceq $row.Item.Name) { $row.Action = 'Same'; [void]$row.Notes.Insert(0, 'already has this name') }
        else { $row.Action = 'Rename' }
        $row.Note = $row.Notes -join '; '
    }
    return ,$plan
}


# ---------------------------------------------------------------- renaming

# Carries out the plan. Returns one result per document:
# Name, NewName, Status ('Done', 'Same', 'Keep' or 'Failed') and Detail
function Invoke-Renaming {
    param([object[]]$Plan, [string]$Mode, [string]$Folder)
    $results = New-Object System.Collections.ArrayList
    $work = @($Plan | Where-Object { $_.Action -eq 'Rename' })
    $record = $null
    if ($Mode -eq 'Copy') {
        $outFolder = Join-Path $Folder $script:CopyFolderName
        if ($work.Count -and -not (Test-Path -LiteralPath $outFolder -PathType Container)) {
            try { New-Item -ItemType Directory -Path $outFolder -ErrorAction Stop | Out-Null }
            catch { throw ('The "renamed" folder could not be made here: ' + $outFolder + "`n`n(" + (Get-InnerMessage $_.Exception) + ')') }
        }
    }

    $moves = @()
    if ($Mode -ne 'Copy' -and $work.Count) {
        $moves = New-FileMoves $Folder @($work | ForEach-Object { , @($_.Item.Name, $_.NewName) })
        # Written down before anything is renamed, with the temporary names too,
        # so the old names are known even if something goes wrong half way
        $record = [pscustomobject]@{
            Id = [guid]::NewGuid().ToString(); Folder = (Get-FolderKey $Folder); When = (Get-Date).ToString('s')
            Entries = @($moves | ForEach-Object { [pscustomobject]@{ Old = $_.From; New = $_.To; Temp = $_.Temp } }) }
        try { Add-HistoryRecord $record }
        catch { throw ('Nothing was renamed, because the list of old names could not be saved in ' + $script:HistoryFile + "`n`n(" + (Get-InnerMessage $_.Exception) + ')') }
        Move-FileNames $Folder $moves
    }

    $byOldName = @{}
    foreach ($m in $moves) { $byOldName[$m.From] = $m }
    foreach ($row in $Plan) {
        if ($row.Action -eq 'Same') { [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Same' $row.Note)); continue }
        if ($row.Action -ne 'Rename') { [void]$results.Add((New-Result $row.Item.Name '' 'Keep' $row.Note)); continue }
        if ($Mode -eq 'Copy') {
            try {
                Copy-Renamed $row.Item.Path (Join-Path $outFolder $row.NewName)
                [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Done' $row.Note))
            } catch {
                [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Failed' (Get-FileError $_.Exception)))
            }
        } else {
            $m = $byOldName[$row.Item.Name]
            if ($m.Done) { [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Done' $row.Note)) }
            else { [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Failed' $m.Error)) }
        }
    }
    if ($record) {
        # The record keeps every file that has a new name now - including one
        # left with its temporary name - so "Put the old names back" can find it
        $record.Entries = @(foreach ($m in $moves) {
            if ($m.Done) { [pscustomobject]@{ Old = $m.From; New = $m.To } }
            elseif ($m.AtTemp) { [pscustomobject]@{ Old = $m.From; New = $m.Temp } }
        })
        try { Set-HistoryRecord $record } catch {}
    }
    return ,$results
}

# The renamings for Move-FileNames: one for each pair of (name now, new name),
# each with a temporary name that is not in use
function New-FileMoves([string]$Folder, [object[]]$Pairs) {
    $inUse = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($n in @(Get-ChildItem -LiteralPath $Folder -Force -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })) { [void]$inUse.Add($n) }
    foreach ($p in $Pairs) { [void]$inUse.Add($p[1]) }
    foreach ($p in $Pairs) {
        $temp = $p[0] + '.renaming'
        $i = 2
        while ($inUse.Contains($temp)) { $temp = '{0}.renaming{1}' -f $p[0], $i; $i++ }
        [void]$inUse.Add($temp)
        [pscustomobject]@{ From = $p[0]; To = $p[1]; Temp = $temp; AtTemp = $false; Done = $false; Error = '' }
    }
}

# Renames files in one folder, even when a file's new name is the name another
# file in the list has now (names may even swap round): every file first gets
# its temporary name, then its new one. A file that can't be renamed keeps its
# name, and so does any file whose new name it would block. Each move ends with
# Done, or with Error saying why not (AtTemp if it was left with its temporary name).
function Move-FileNames([string]$Folder, [object[]]$Moves) {
    $leaving = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $Moves) { [void]$leaving.Add($m.From) }
    # A new name that belongs to a file that isn't moving can't be given
    foreach ($m in $Moves) {
        if (-not $m.Error -and -not $leaving.Contains($m.To) -and (Test-Path -LiteralPath (Join-Path $Folder $m.To))) {
            $m.Error = 'another file with the new name is already there.'
        }
    }
    foreach ($m in $Moves) {
        if ($m.Error) { continue }
        try { [IO.File]::Move((Join-Path $Folder $m.From), (Join-Path $Folder $m.Temp)); $m.AtTemp = $true }
        catch { $m.Error = Get-FileError $_.Exception }
    }
    # A file that didn't move still has its name, so a file meant to get that name
    # goes back to its own old name - which may block another, and so on
    $held = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $Moves) { if (-not $m.AtTemp) { [void]$held.Add($m.From) } }
    do {
        $again = $false
        foreach ($m in $Moves) {
            if ($m.AtTemp -and -not $m.Error -and $held.Contains($m.To)) {
                $m.Error = 'its new name still belongs to a file that could not be renamed.'
                try {
                    [IO.File]::Move((Join-Path $Folder $m.Temp), (Join-Path $Folder $m.From))
                    $m.AtTemp = $false
                    [void]$held.Add($m.From)
                    $again = $true
                } catch { $m.Error += ' ' + (Get-StuckNote $m) }
            }
        }
    } while ($again)
    foreach ($m in $Moves) {
        if (-not $m.AtTemp -or $m.Error) { continue }
        try {
            [IO.File]::Move((Join-Path $Folder $m.Temp), (Join-Path $Folder $m.To))
            $m.AtTemp = $false
            $m.Done = $true
        } catch {
            $m.Error = Get-FileError $_.Exception
            try { [IO.File]::Move((Join-Path $Folder $m.Temp), (Join-Path $Folder $m.From)); $m.AtTemp = $false }
            catch { $m.Error += ' ' + (Get-StuckNote $m) }
        }
    }
}

function Get-StuckNote($Move) {
    return 'For now it is called "' + $Move.Temp + '"; "Put the old names back" gives it its old name.'
}

function New-Result([string]$Name, [string]$NewName, [string]$Status, [string]$Detail) {
    [pscustomobject]@{ Name = $Name; NewName = $NewName; Status = $Status; Detail = $Detail }
}

# Saves a copy with the new name. A copy left there by an earlier run is replaced.
function Copy-Renamed([string]$From, [string]$To) {
    if (Test-Path -LiteralPath $To -PathType Leaf) {
        $old = Get-Item -LiteralPath $To -Force
        if ($old.IsReadOnly) { $old.IsReadOnly = $false }
    }
    [IO.File]::Copy($From, $To, $true)
}

function Get-FileError($Ex) {
    $inner = $Ex
    while ($inner.InnerException) { $inner = $inner.InnerException }
    $msg = Get-InnerMessage $Ex
    if ($msg -match 'being used by another process') { return 'the file is open in another program (such as Word). Close it and try again.' }
    if ($inner -is [UnauthorizedAccessException]) { return 'Windows did not allow it. The file may be read-only, or you may not have permission to change it.' }
    if ($inner -is [IO.FileNotFoundException] -or $inner -is [IO.DirectoryNotFoundException]) { return 'the file is not there any more (it may have been moved, renamed or deleted).' }
    if ($msg -match 'already exists') { return 'another file with the new name is already there.' }
    return $msg
}


# ---------------------------------------------------------------- putting the old names back

function Get-FolderKey([string]$Folder) {
    return [IO.Path]::GetFullPath($Folder).TrimEnd('\')
}

function Read-History {
    $list = New-Object System.Collections.ArrayList
    if (Test-Path -LiteralPath $script:HistoryFile) {
        $raw = [IO.File]::ReadAllText($script:HistoryFile, [Text.Encoding]::UTF8)
        if ($raw.Trim()) { foreach ($r in (ConvertFrom-Json $raw)) { [void]$list.Add($r) } }
    }
    return ,$list
}

function Save-History($Records) {
    $dir = Split-Path -Parent $script:HistoryFile
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -ErrorAction Stop | Out-Null }
    $keep = @($Records)
    if ($keep.Count -gt $script:HistoryMax) { $keep = @($keep[($keep.Count - $script:HistoryMax)..($keep.Count - 1)]) }
    $json = ConvertTo-Json -InputObject $keep -Depth 6
    if ($keep.Count -eq 0) { $json = '[]' }
    # Written to a new file first, so a half-written list never replaces a good one
    $tmp = $script:HistoryFile + '.new'
    [IO.File]::WriteAllText($tmp, $json, (New-Object Text.UTF8Encoding($false)))
    if (Test-Path -LiteralPath $script:HistoryFile) { [IO.File]::Replace($tmp, $script:HistoryFile, [NullString]::Value) }
    else { [IO.File]::Move($tmp, $script:HistoryFile) }
}

function Add-HistoryRecord($Record) {
    $all = New-Object System.Collections.ArrayList
    try { $all = Read-History }
    catch {
        # A damaged list is kept aside, and a new one started
        try { Copy-Item -LiteralPath $script:HistoryFile -Destination ($script:HistoryFile + '.damaged') -Force } catch {}
    }
    [void]$all.Add($Record)
    Save-History $all
}

# Replaces the record with the same Id. One with nothing left in it is removed.
function Set-HistoryRecord($Record) {
    $all = New-Object System.Collections.ArrayList
    foreach ($r in (Read-History)) {
        if ($r.Id -ne $Record.Id) { [void]$all.Add($r) }
        elseif (@($Record.Entries).Count) { [void]$all.Add($Record) }
    }
    Save-History $all
}

# The latest renaming of the files in this folder, or $null
function Get-LastRenaming([string]$Folder) {
    $last = $null
    try {
        $key = Get-FolderKey $Folder
        foreach ($r in (Read-History)) { if ($r.Folder -ieq $key -and @($r.Entries).Count) { $last = $r } }
    } catch {}
    return $last
}

# Puts back the old names from the latest renaming in this folder. Results as
# for Invoke-Renaming, with Name = the name now and NewName = the old name.
function Undo-LastRenaming([string]$Folder) {
    $results = New-Object System.Collections.ArrayList
    $record = Get-LastRenaming $Folder
    if (-not $record) { return ,$results }
    # Where each file is now: under its new name, or (if a renaming stopped half
    # way) under its temporary name
    $pairs = New-Object System.Collections.ArrayList
    foreach ($e in @($record.Entries)) {
        $now = $null
        if (Test-Path -LiteralPath (Join-Path $Folder $e.New) -PathType Leaf) { $now = [string]$e.New }
        elseif ($e.Temp -and (Test-Path -LiteralPath (Join-Path $Folder $e.Temp) -PathType Leaf)) { $now = [string]$e.Temp }
        if (-not $now) {
            if (Test-Path -LiteralPath (Join-Path $Folder $e.Old) -PathType Leaf) { [void]$results.Add((New-Result $e.New $e.Old 'Same' 'already has its old name')) }
            else { [void]$results.Add((New-Result $e.New $e.Old 'Failed' 'the file is not there any more (it may have been moved, renamed or deleted).')) }
            continue
        }
        if ($now -ceq $e.Old) { [void]$results.Add((New-Result $e.New $e.Old 'Same' 'already has its old name')); continue }
        [void]$pairs.Add(@($now, [string]$e.Old))
    }
    $moves = @()
    if ($pairs.Count) {
        $moves = New-FileMoves $Folder $pairs
        foreach ($m in $moves) {
            # Moving out of the way of another file is fine; a different file in the way is not
            if (-not ($moves | Where-Object { $_.From -ieq $m.To }) -and (Test-Path -LiteralPath (Join-Path $Folder $m.To))) {
                $m.Error = 'another file is called "' + $m.To + '" now, so this one keeps its new name.'
            }
        }
        Move-FileNames $Folder $moves
    }
    $left = New-Object System.Collections.ArrayList
    foreach ($m in $moves) {
        if ($m.Done) { [void]$results.Add((New-Result $m.From $m.To 'Done' '')); continue }
        [void]$results.Add((New-Result $m.From $m.To 'Failed' $m.Error))
        $name = $m.From
        if ($m.AtTemp) { $name = $m.Temp }
        [void]$left.Add([pscustomobject]@{ Old = $m.To; New = $name })
    }
    $record.Entries = @($left)
    try { Set-HistoryRecord $record } catch {}
    return ,$results
}