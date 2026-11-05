# RenameEngine.ps1
# The part of Document Renamer that does the real work: it finds the .rtf,
# .doc, .docx and .txt files in a folder, reads the start of each one (with
# Microsoft Word, or LibreOffice when Word can't open a file), works out the
# new names, and then saves renamed copies or renames the files themselves.
# Every renaming of the files themselves is written down, so that the old
# names can be put back. "Document Renamer.ps1" (the window) loads this file.

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


# ---------------------------------------------------------------- finding files

# Each document gets a Text (the start of its words, once read), or a Problem
# saying why it can't be renamed, with ProblemKind 'Skip' or 'Error'
function Find-Documents {
    param([string]$Folder)
    $files = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction Stop |
        Where-Object { $script:DocExtensions -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object Name)

    $items = New-Object System.Collections.ArrayList
    foreach ($f in $files) {
        $skip = $null; $kind = $null
        if ($f.Name.StartsWith('~$')) { $skip = 'skipped: a temporary file Word makes while a document is open' }
        elseif ($f.Length -eq 0) { $skip = 'skipped: the file is empty (0 bytes)' }
        if ($skip) { $kind = 'Skip' }
        [void]$items.Add([pscustomobject]@{
            Path = $f.FullName; Name = $f.Name; Extension = $f.Extension
            Type = $f.Extension.TrimStart('.').ToLowerInvariant()
            Text = $null; Problem = $skip; ProblemKind = $kind; Note = '' })
    }
    return ,$items
}


# ---------------------------------------------------------------- reading the start of each document

function Read-DocumentStarts {
    param(
        [object[]]$Items,          # from Find-Documents; each one gets its Text, or a Problem
        [scriptblock]$OnFile,      # & $OnFile <number> <total> <item>, before each file
        [scriptblock]$OnTick,      # called now and then so the window stays responsive
        [scriptblock]$ShouldStop   # returns $true when the user pressed Stop
    )
    $todo = @($Items | Where-Object { -not $_.Problem })
    # Word is only started when there is a Word or Rich Text file to read
    $readers = @{ Word = $null; WordTried = $false; LibreOffice = (Find-LibreOffice) }
    try {
        for ($i = 0; $i -lt $todo.Count; $i++) {
            $it = $todo[$i]
            if ($ShouldStop -and (& $ShouldStop)) {
                for ($j = $i; $j -lt $todo.Count; $j++) { $todo[$j].Problem = 'not read, because you pressed Stop'; $todo[$j].ProblemKind = 'Skip' }
                break
            }
            if ($OnFile) { & $OnFile ($i + 1) $todo.Count $it }
            try {
                if ($it.Type -eq 'txt') { $raw = Read-TextFile $it.Path }
                else {
                    $got = Read-OfficeFile $readers $it.Path $OnTick
                    $raw = $got.Text
                    $it.Note = $got.Note
                }
                $it.Text = Get-CleanText $raw
            } catch {
                $it.Problem = Get-FriendlyError (Get-InnerMessage $_.Exception)
                $it.ProblemKind = 'Error'
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
    if ($got -lt $bytes.Length) { $bytes = [byte[]]($bytes[0..($got - 1)]) }
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
    # Characters Windows doesn't allow in a file name
    $t = [regex]::Replace($t, '[/\\|]', '-')
    $t = $t.Replace('"', "'")
    $t = [regex]::Replace($t, '[<>*?]', '')
    $t = [regex]::Replace($t, ' {2,}', ' ').Trim()
    $t = [regex]::Replace($t, ' ?: ', ' - ')   # "Re: Meeting" becomes "Re - Meeting"
    $t = [regex]::Replace($t, ':$', '')
    $t = $t.Replace(':', '-')                  # "10:30" becomes "10-30"
    $t = $t.TrimStart(' ', '.', '-')           # such as a "- " typed in front of a list
    if ($t.Length -gt $script:KeepChars) { $t = $t.Substring(0, $script:KeepChars) }
    return $t
}

# The new name, without the file type, from a document's cleaned-up start
function Get-NameFromText([string]$Text, [int]$Length) {
    if (-not $Text) { return '' }
    $name = $Text
    if ($name.Length -gt $Length) {
        $cut = $Length
        if ([char]::IsHighSurrogate($name[$cut - 1])) { $cut-- }   # don't cut an emoji in half
        $name = $name.Substring(0, $cut)
    }
    # Windows doesn't allow a name to end with a space or a period
    $name = $name.TrimEnd(' ', '.', '-')
    # Names Windows keeps for itself, such as CON, PRN, NUL, COM1 and LPT1
    if ($name -match '^(CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[0-9\u00B9\u00B2\u00B3]|LPT[0-9\u00B9\u00B2\u00B3]) *(\.|$)') { $name = '_' + $name }
    return $name
}

# Works out every document's new name. $Mode is 'Copy' (renamed copies go into
# the "renamed" folder) or 'Rename' (the files themselves get the new names).
# $ExistingNames: the names of everything in the folder now, for 'Rename'.
# Each row of the plan has the Item, its NewName, a Note and an Action:
#   'Rename' - gets the new name (for 'Copy': a copy with the new name is saved)
#   'Same'   - already has this name, so nothing happens (only for 'Rename')
#   'Keep'   - no new name: no words were found, or the file couldn't be read
function Get-RenamePlan {
    param([object[]]$Items, [int]$Length, [string]$Mode, [string[]]$ExistingNames)
    $Length = [Math]::Max(1, $Length)
    # A new name is never used twice. When renaming in place it also can't be the
    # name of anything else in the folder - not even of a file that is getting a
    # new name itself - so no file ever has to wait for another to move out of the way.
    $taken = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if ($Mode -eq 'Rename') { foreach ($n in $ExistingNames) { [void]$taken.Add($n) } }

    $plan = New-Object System.Collections.ArrayList
    foreach ($it in $Items) {
        $row = [pscustomobject]@{ Item = $it; NewName = ''; Action = 'Keep'; Note = '' }
        [void]$plan.Add($row)
        if ($it.Problem) { $row.Note = $it.Problem; continue }
        $base = Get-NameFromText $it.Text $Length
        if (-not $base) { $row.Note = 'no words found at the start of this document'; continue }
        $name = $base + $it.Extension
        $number = 1
        # When renaming in place, a file's own name is free for itself
        while ($taken.Contains($name) -and -not ($Mode -eq 'Rename' -and $name -ieq $it.Name)) {
            $number++
            $name = '{0} ({1}){2}' -f $base, $number, $it.Extension
        }
        [void]$taken.Add($name)
        $row.NewName = $name
        if ($Mode -eq 'Rename' -and $name -ceq $it.Name) { $row.Action = 'Same'; $row.Note = 'already has this name' }
        else {
            $row.Action = 'Rename'
            if ($number -gt 1) { $row.Note = "($number) added, because another file already has this name" }
        }
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
    } elseif ($work.Count) {
        # Written down before anything is renamed, so the old names are known
        # even if something goes wrong half way
        $record = [pscustomobject]@{
            Id = [guid]::NewGuid().ToString(); Folder = (Get-FolderKey $Folder); When = (Get-Date).ToString('s')
            Entries = @($work | ForEach-Object { [pscustomobject]@{ Old = $_.Item.Name; New = $_.NewName } }) }
        try { Add-HistoryRecord $record }
        catch { throw ('Nothing was renamed, because the list of old names could not be saved in ' + $script:HistoryFile + "`n`n(" + (Get-InnerMessage $_.Exception) + ')') }
    }

    $done = New-Object System.Collections.ArrayList
    foreach ($row in $Plan) {
        if ($row.Action -eq 'Same') { [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Same' $row.Note)); continue }
        if ($row.Action -ne 'Rename') { [void]$results.Add((New-Result $row.Item.Name '' 'Keep' $row.Note)); continue }
        try {
            if ($Mode -eq 'Copy') { Copy-Renamed $row.Item.Path (Join-Path $outFolder $row.NewName) }
            else { [IO.File]::Move($row.Item.Path, (Join-Path $Folder $row.NewName)) }
            [void]$done.Add($row)
            [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Done' $row.Note))
        } catch {
            [void]$results.Add((New-Result $row.Item.Name $row.NewName 'Failed' (Get-FileError $_.Exception)))
        }
    }
    if ($record) {
        $record.Entries = @($done | ForEach-Object { [pscustomobject]@{ Old = $_.Item.Name; New = $_.NewName } })
        try { Set-HistoryRecord $record } catch {}
    }
    return ,$results
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
    $left = New-Object System.Collections.ArrayList
    foreach ($e in @($record.Entries)) {
        $now = Join-Path $Folder $e.New
        $old = Join-Path $Folder $e.Old
        if (-not (Test-Path -LiteralPath $now -PathType Leaf)) {
            if (Test-Path -LiteralPath $old -PathType Leaf) { [void]$results.Add((New-Result $e.New $e.Old 'Same' 'already has its old name')) }
            else { [void]$results.Add((New-Result $e.New $e.Old 'Failed' 'the file is not there any more (it may have been moved, renamed or deleted).')) }
            continue
        }
        if ($e.New -ine $e.Old -and (Test-Path -LiteralPath $old)) {
            [void]$results.Add((New-Result $e.New $e.Old 'Failed' ('another file is called "' + $e.Old + '" now, so this one keeps its new name.')))
            [void]$left.Add($e)
            continue
        }
        try {
            [IO.File]::Move($now, $old)
            [void]$results.Add((New-Result $e.New $e.Old 'Done' ''))
        } catch {
            [void]$results.Add((New-Result $e.New $e.Old 'Failed' (Get-FileError $_.Exception)))
            [void]$left.Add($e)
        }
    }
    $record.Entries = @($left)
    try { Set-HistoryRecord $record } catch {}
    return ,$results
}