# ReplaceEngine.ps1
# The part of Find and Replace that does the real work. In each .rtf, .doc,
# .docx and .txt file it finds what is to be replaced:
#   - words or phrases, everywhere they appear (like Find and Replace in Word)
#   - the beginning: everything from the start of the document up to some text
#   - the ending: everything from some text to the end of the document
# First it only looks (the preview), then it makes the changes and saves the
# documents. Word's special ^ codes (^p, ^l, ^t and so on) work in all the boxes,
# and words, phrases and marks can also be found and replaced by their formatting.
# Word documents are changed by Microsoft Word in the background, .txt files
# directly. "Find and Replace.ps1" (the window) loads this file.

$script:DocExtensions    = @('.rtf', '.doc', '.docx', '.txt')
$script:CopyFolderName   = 'replaced'
$script:BackupFolderName = 'before replace'
# How much the preview shows
$script:ShowWords  = 40    # word-or-phrase replacements in each document
$script:ShowAround = 45    # characters before and after each of them
$script:ShowEdge   = 350   # characters from each end of a removed beginning or ending
$script:ShowKept   = 150   # characters of the text that stays, next to it
$script:MaxHits    = 100000


# ---------------------------------------------------------------- the documents

function Test-DocFile([string]$Path) {
    return ($script:DocExtensions -contains [IO.Path]::GetExtension($Path).ToLowerInvariant())
}

# The documents directly in a folder (not in the folders inside it)
function Get-FolderDocuments([string]$Folder) {
    return @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction Stop |
        Where-Object { Test-DocFile $_.Name } | Sort-Object Name | ForEach-Object { $_.FullName })
}

# One document in the list. Preview and Result are filled in later.
function New-DocItem([string]$Path) {
    $f = Get-Item -LiteralPath $Path -ErrorAction Stop
    $problem = $null
    if ($f.Name.StartsWith('~$')) { $problem = 'skipped: a temporary file Word makes while a document is open' }
    elseif ($f.Length -eq 0) { $problem = 'skipped: the file is empty (0 bytes)' }
    return [pscustomobject]@{
        Path = $f.FullName; Name = $f.Name; Folder = $f.DirectoryName
        Type = $f.Extension.TrimStart('.').ToLowerInvariant()
        Problem = $problem; Preview = $null; Result = $null }
}


# ---------------------------------------------------------------- going through the documents

# $Settings comes from the window:
#   Words = @{ On; Find; Replace; MatchCase; WholeWords; FindFormat; NewFormat }
#   Begin = @{ On; Find; Include; MatchCase; Which ('First' or 'Last'); Source ('Text' or 'File'); Replace; File }
#   End   = the same as Begin
#   Between = @{ On; From; To; IncludeFrom; IncludeTo; MatchCase }
# Without -Apply each document only gets its Preview. With -Apply the changes
# are made and each document gets its Result. $SaveMode 'Copy' saves the changed
# documents in a "replaced" folder next to them. 'InPlace' changes the documents
# themselves, after keeping a copy of each one in "before replace\<Stamp>".
function Invoke-Documents {
    param(
        [object[]]$Items,
        [hashtable]$Settings,
        [switch]$Apply,
        [string]$SaveMode = 'Copy',
        [string]$Stamp,
        [scriptblock]$OnFile,      # & $OnFile <number> <total> <item>, before each file
        [scriptblock]$OnTick,      # called after each file so the window stays responsive
        [scriptblock]$ShouldStop   # returns $true when the user pressed Stop
    )
    foreach ($it in $Items) {
        if (-not $it.Problem) { continue }
        if ($Apply) { $it.Result = New-Result 'Skipped' $it.Problem } else { $it.Preview = New-Preview 'Skipped' $it.Problem }
    }
    $todo = @($Items | Where-Object { -not $_.Problem })
    # Documents the preview found nothing to change in (or couldn't read) are left alone
    if ($Apply) {
        foreach ($it in $todo) {
            $p = $it.Preview
            if (-not $p) { $it.Result = New-Result 'Skipped' 'not looked at in the preview' }
            elseif ($p.Status -eq 'Error') { $it.Result = New-Result 'Failed' $p.Summary }
            elseif ($p.Status -eq 'Skipped') { $it.Result = New-Result 'Skipped' $p.Summary }
            elseif (-not $p.Changes) { $it.Result = New-Result 'None' $p.Summary }
        }
        $todo = @($todo | Where-Object { $_.Preview -and $_.Preview.Changes })
    }
    # Word is only started when there is a Word or Rich Text file to work on
    $state = @{ Word = $null; WordTried = $false; Sources = @{} }
    try {
        for ($i = 0; $i -lt $todo.Count; $i++) {
            $it = $todo[$i]
            if ($ShouldStop -and (& $ShouldStop)) {
                for ($j = $i; $j -lt $todo.Count; $j++) {
                    if ($Apply) { $todo[$j].Result = New-Result 'Skipped' 'not changed, because you pressed Stop' }
                    else { $todo[$j].Preview = New-Preview 'Skipped' 'not looked at, because you pressed Stop' }
                }
                break
            }
            if ($OnFile) { & $OnFile ($i + 1) $todo.Count $it }
            try {
                if ($it.Type -eq 'txt') { $out = Invoke-TextDocument $state $it $Settings $Apply.IsPresent $SaveMode $Stamp }
                else { $out = Invoke-WordDocument $state $it $Settings $Apply.IsPresent $SaveMode $Stamp }
                if ($Apply) { $it.Result = $out } else { $it.Preview = $out }
            } catch {
                $msg = Get-FriendlyError (Get-InnerMessage $_.Exception)
                if ($Apply) { $it.Result = New-Result 'Failed' $msg } else { $it.Preview = New-Preview 'Error' $msg }
            }
            if ($OnTick) { & $OnTick }
        }
    } finally {
        Stop-Word $state.Word
    }
}

function New-Preview([string]$Status, [string]$Summary) {
    return [pscustomobject]@{
        Status = $Status; Summary = $Summary; Changes = $false; Segments = (New-Object System.Collections.ArrayList)
        Begin = $null; End = $null; Between = $null; WordsCount = 0; WordsNote = ''; KeepStart = 0; KeepEnd = 0; Hits = $null }
}

# Status: 'Done', 'None' (nothing to change), 'Skipped' or 'Failed'
function New-Result([string]$Status, [string]$Detail, [string]$SavedTo = '', [string]$Backup = '') {
    return [pscustomobject]@{ Status = $Status; Detail = $Detail; SavedTo = $SavedTo; Backup = $Backup }
}

# Where a changed document is saved. For 'InPlace' a copy of the document is
# kept in "before replace" first, and that copy is what gets opened and changed.
function Get-SaveTarget($Item, [string]$SaveMode, [string]$Stamp) {
    if ($SaveMode -eq 'Copy') {
        $dir = Join-Path $Item.Folder $script:CopyFolderName
        New-FolderIfMissing $dir
        $path = [string](Join-Path $dir $Item.Name)
        Clear-ReadOnly $path
        return @{ Open = $Item.Path; Save = $path; Backup = '' }
    }
    Assert-CanChange $Item.Path
    $dir = Join-Path (Join-Path $Item.Folder $script:BackupFolderName) $Stamp
    New-FolderIfMissing $dir
    $backup = [string](Join-Path $dir $Item.Name)
    [IO.File]::Copy($Item.Path, $backup, $true)
    return @{ Open = $backup; Save = $Item.Path; Backup = $backup }
}

function New-FolderIfMissing([string]$Dir) {
    if (Test-Path -LiteralPath $Dir -PathType Container) { return }
    try { New-Item -ItemType Directory -Path $Dir -ErrorAction Stop | Out-Null }
    catch { throw ('FOLDER:' + $Dir) }
}

function Clear-ReadOnly([string]$Path) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        $f = Get-Item -LiteralPath $Path -Force
        if ($f.IsReadOnly) { $f.IsReadOnly = $false }
    }
}

# A document can only be changed if it isn't read-only and no other program has it open
function Assert-CanChange([string]$Path) {
    if ((Get-Item -LiteralPath $Path -Force).IsReadOnly) { throw 'READONLY:' }
    try { $fs = [IO.File]::Open($Path, 'Open', 'ReadWrite', 'None'); $fs.Dispose() }
    catch {
        if ((Get-InnerMessage $_.Exception) -match 'being used by another process') { throw 'INUSE:' }
        throw
    }
}

# The copy kept of a document that turned out to need no change isn't needed
function Remove-UnneededBackup([string]$Backup) {
    if (-not $Backup) { return }
    try {
        [IO.File]::Delete($Backup)
        $stampDir = [IO.Path]::GetDirectoryName($Backup)
        if (@([IO.Directory]::GetFileSystemEntries($stampDir)).Count -eq 0) {
            [IO.Directory]::Delete($stampDir)
            $top = [IO.Path]::GetDirectoryName($stampDir)
            if (@([IO.Directory]::GetFileSystemEntries($top)).Count -eq 0) { [IO.Directory]::Delete($top) }
        }
    } catch {}
}

# Puts back the documents that were changed in place, from the copies kept in
# "before replace". Returns one line per document.
function Restore-Originals([object[]]$Items) {
    $lines = New-Object System.Collections.ArrayList
    foreach ($it in $Items) {
        if (-not ($it.Result -and $it.Result.Status -eq 'Done' -and $it.Result.Backup)) { continue }
        try {
            Clear-ReadOnly $it.Path
            [IO.File]::Copy($it.Result.Backup, $it.Path, $true)
            [void]$lines.Add([pscustomobject]@{ Name = $it.Name; Ok = $true; Detail = 'put back as it was' })
        } catch {
            [void]$lines.Add([pscustomobject]@{ Name = $it.Name; Ok = $false; Detail = (Get-FileError $_.Exception) })
        }
    }
    return ,$lines
}

# ---------------------------------------------------------------- Word's special ^ codes

# The codes that can be typed, as in Word's Find and Replace: ^p is a paragraph
# mark, ^l a line break, ^t a tab, and so on. ^^ is a ^ itself. Numbers work too,
# as in Word: ^13, ^0150 (Windows alphabet) and ^u8212 (Unicode).
$script:FindCodes    = @('p', 't', 'l', 'n', 'm', 'b', 's', '~', '-', '+', '=', '^', 'w', '#', '$', '?', '%', 'v', 'a', 'f', 'e', 'g', 'd')
$script:ReplaceCodes = @('p', 't', 'l', 'n', 'm', 's', '~', '-', '+', '=', '^', '%', 'v', '&')

# Splits what was typed into pieces: Kind 'Text' (plain text) or 'Code' (what
# follows a ^, in small letters, or nothing for a ^ with nothing after it)
function Get-CodeTokens([string]$Text) {
    $tokens = New-Object System.Collections.ArrayList
    $pos = 0
    foreach ($m in [regex]::Matches($Text, '\^([uU]\d{1,5}|0\d{1,3}|[1-9]\d{0,2}|[\s\S])?')) {
        if ($m.Index -gt $pos) { [void]$tokens.Add(@{ Kind = 'Text'; Value = $Text.Substring($pos, $m.Index - $pos) }) }
        [void]$tokens.Add(@{ Kind = 'Code'; Value = $m.Groups[1].Value.ToLowerInvariant() })
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $Text.Length) { [void]$tokens.Add(@{ Kind = 'Text'; Value = $Text.Substring($pos) }) }
    return ,$tokens
}

# The first ^ code that can't be used, or $null when they are all fine. $Kind:
# 'Find', 'Replace', or 'NewText' (a new beginning or ending, where ^& means nothing).
function Get-CodeProblem([string]$Text, [string]$Kind) {
    $allowed = $script:ReplaceCodes
    if ($Kind -eq 'Find') { $allowed = $script:FindCodes }
    foreach ($t in (Get-CodeTokens $Text)) {
        if ($t.Kind -ne 'Code') { continue }
        $c = $t.Value
        if ($c -match '^u(\d+)$') { if ([int]$Matches[1] -lt 1 -or [int]$Matches[1] -gt 65535) { return '^' + $c }; continue }
        if ($c -match '^\d+$') { if ([int]$c -lt 1 -or [int]$c -gt 255) { return '^' + $c }; continue }
        if ($allowed -notcontains $c -or ($Kind -eq 'NewText' -and $c -eq '&')) { return '^' + $c }
    }
    return $null
}

# What a ^ code in a "replace with" box stands for. $Eol is the kind of line
# break of a .txt file (a .txt file has no other kinds of break), empty for Word.
function Get-CodeChar([string]$Code, [string]$Eol) {
    if ($Code -match '^u(\d+)$') { return [string][char][int]$Matches[1] }
    if ($Code -match '^\d+$') {
        $n = [int]$Code
        if ($n -eq 13) { $Code = 'p' } elseif ($n -eq 11) { $Code = 'l' } elseif ($n -eq 14) { $Code = 'n' }
        elseif ($n -lt 128) { return [string][char]$n }
        elseif ($Code[0] -eq '0') { return [Text.Encoding]::GetEncoding(1252).GetString([byte[]]@($n)) }
        else { return [Text.Encoding]::GetEncoding(437).GetString([byte[]]@($n)) }
    }
    switch ($Code) {
        'p' { if ($Eol) { return $Eol }; return "`r" }
        'l' { if ($Eol) { return $Eol }; return [string][char]11 }
        'n' { if ($Eol) { return $Eol }; return [string][char]14 }
        't' { return "`t" }
        'm' { return [string][char]12 }
        's' { return [string][char]0xA0 }
        '~' { if ($Eol) { return [string][char]0x2011 }; return [string][char]30 }
        '-' { if ($Eol) { return [string][char]0xAD }; return [string][char]31 }
        '+' { return [string][char]0x2014 }
        '=' { return [string][char]0x2013 }
        '^' { return '^' }
        '%' { return [string][char]0xA7 }
        'v' { return [string][char]0xB6 }
    }
    return ''
}

# Replaces the ^ codes in what was typed with what they stand for. ^& stands
# for $Found. Lines typed in a box become paragraphs (or a .txt file's lines).
function Expand-Codes([string]$Text, [string]$Found = '', [string]$Eol = '') {
    $para = "`r"
    if ($Eol) { $para = $Eol }
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in (Get-CodeTokens $Text)) {
        if ($t.Kind -eq 'Text') { [void]$sb.Append($t.Value.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", $para)) }
        elseif ($t.Value -eq '&') { [void]$sb.Append($Found) }
        else { [void]$sb.Append((Get-CodeChar $t.Value $Eol)) }
    }
    return $sb.ToString()
}

# What one find is replaced with. A replacement that includes the found text
# itself (^&) is used as it is; otherwise capitals follow the found text, like Word.
function Get-NewText([string]$Found, [string]$Typed, [bool]$MatchCase, [string]$Eol = '') {
    $new = Expand-Codes $Typed $Found $Eol
    foreach ($t in (Get-CodeTokens $Typed)) { if ($t.Kind -eq 'Code' -and $t.Value -eq '&') { return $new } }
    return Get-CaseAdapted $Found $new $MatchCase
}

# The search pattern for a ^ code in a .txt file. Things a .txt file can't
# have, such as pictures and footnotes, are simply never found.
function Get-CodeRegex([string]$Code) {
    if ($Code -match '^(u\d+|\d+)$') {
        $ch = Get-CodeChar $Code ''
        if ($ch -eq "`r") { return '(?>\r\n|\n|\r)' }
        return [regex]::Escape($ch)
    }
    switch ($Code) {
        'p' { return '(?>\r\n|\n|\r)' }
        't' { return '\t' }
        'l' { return '[\x0B' + [char]0x2028 + ']' }
        'n' { return '\x0E' }
        'm' { return '\f' }
        's' { return [string][char]0xA0 }
        '~' { return '[\x1E' + [char]0x2011 + ']' }
        '-' { return '[\x1F' + [char]0xAD + ']' }
        '+' { return [string][char]0x2014 }
        '=' { return [string][char]0x2013 }
        '^' { return '\^' }
        'w' { return '[ \t' + [char]0xA0 + ']+' }
        '#' { return '[0-9]' }
        '$' { return '\p{L}' }
        '?' { return '(?>\r\n|[\s\S])' }
        '%' { return [string][char]0xA7 }
        'v' { return [string][char]0xB6 }
    }
    return '(?!)'
}

# ---------------------------------------------------------------- formatting

# A format is a hashtable with any of: Font, Size, Bold, Italic, Underline and
# Strike ('Yes' or 'No'), Color and ColorName, Highlight and HighlightName, and
# Style. When finding, Highlight is 'Yes' or 'No'; when replacing, it is one of
# Word's highlight color numbers (0 takes the highlight away).
function Test-HasFormat($F) {
    if (-not $F) { return $false }
    foreach ($k in 'Font', 'Size', 'Bold', 'Italic', 'Underline', 'Strike', 'Color', 'Highlight', 'Style') {
        if ($null -ne $F[$k] -and [string]$F[$k] -ne '') { return $true }
    }
    return $false
}

# A format in words, such as "Arial, 12 pt, bold, red"
function Get-FormatText($F) {
    $parts = @()
    if ($F.Font) { $parts += [string]$F.Font }
    if ($F.Size) { $parts += "$($F.Size) pt" }
    foreach ($p in @(@('Bold', 'bold', 'not bold'), @('Italic', 'italic', 'not italic'), @('Underline', 'underlined', 'not underlined'), @('Strike', 'strikethrough', 'not strikethrough'))) {
        if ($F[$p[0]] -eq 'Yes') { $parts += $p[1] } elseif ($F[$p[0]] -eq 'No') { $parts += $p[2] }
    }
    if ([string]$F.Color -ne '') { $parts += [string]$F.ColorName }
    if ([string]$F.Highlight -ne '') { $parts += [string]$F.HighlightName }
    if ($F.Style) { $parts += "style $($F.Style)" }
    return ($parts -join ', ')
}

# A replacement with no text but with a format keeps the found text and only
# changes how it looks, like Word
function Test-KeepsText($W) {
    return (-not $W.Replace) -and (Test-HasFormat $W.NewFormat)
}

# What is looked for, in words: "cat" [bold]
function Get-WordsLabel($W, [int]$Max) {
    $parts = @()
    if ($W.Find) { $parts += '"' + (Get-Short $W.Find $Max) + '"' }
    if (Test-HasFormat $W.FindFormat) { $parts += '[' + (Get-FormatText $W.FindFormat) + ']' }
    return ($parts -join ' ')
}

function Set-FindFormat($Find, $F) {
    if ($F.Font) { $Find.Font.Name = [string]$F.Font }
    if ($F.Size) { $Find.Font.Size = [single]$F.Size }
    if ($F.Bold) { $Find.Font.Bold = [int]($F.Bold -eq 'Yes') }
    if ($F.Italic) { $Find.Font.Italic = [int]($F.Italic -eq 'Yes') }
    if ($F.Underline) { $Find.Font.Underline = [int]($F.Underline -eq 'Yes') }
    if ($F.Strike) { $Find.Font.StrikeThrough = [int]($F.Strike -eq 'Yes') }
    if ([string]$F.Color -ne '') { $Find.Font.Color = [int]$F.Color }
    if ($F.Highlight) { if ($F.Highlight -eq 'Yes') { $Find.Highlight = -1 } else { $Find.Highlight = 0 } }
    if ($F.Style) {
        try { $Find.Style = [string]$F.Style } catch { throw 'NOSTYLE:' }
    }
}

# Gives text a format. The style goes first, so that the rest is not undone by it.
function Set-RangeFormat($R, $F) {
    if (-not (Test-HasFormat $F)) { return }
    if ($F.Style) { $R.Style = [string]$F.Style }
    if ($F.Font) { $R.Font.Name = [string]$F.Font }
    if ($F.Size) { $R.Font.Size = [single]$F.Size }
    if ($F.Bold) { $R.Font.Bold = [int]($F.Bold -eq 'Yes') }
    if ($F.Italic) { $R.Font.Italic = [int]($F.Italic -eq 'Yes') }
    if ($F.Underline) { $R.Font.Underline = [int]($F.Underline -eq 'Yes') }
    if ($F.Strike) { $R.Font.StrikeThrough = [int]($F.Strike -eq 'Yes') }
    if ([string]$F.Color -ne '') { $R.Font.Color = [int]$F.Color }
    if ([string]$F.Highlight -ne '') { $R.HighlightColorIndex = [int]$F.Highlight }
}

# Says so when a style to give the text isn't in this document. Word's own
# styles are always there; one a person made may only be in some documents.
function Get-StyleProblem($Doc, $F) {
    if (-not ($F -and $F.Style)) { return '' }
    try { [void]$Doc.Styles.Item([string]$F.Style); return '' }
    catch { return "the style ""$($F.Style)"" is not in this document" }
}


# ---------------------------------------------------------------- sections between two phrases

# Pairs up the places the two phrases were found. Each first phrase goes with
# the next last phrase after it, and the search goes on after that one, so a
# first phrase inside a section just goes with it. Each pair is the first
# phrase's start and end, then the last phrase's start and end. Unclosed counts
# the first phrases with no last phrase after them.
function Get-SectionPairs($FromHits, $ToHits) {
    $pairs = New-Object System.Collections.ArrayList
    $next = 0
    $pos = -1
    $unclosed = 0
    foreach ($a in $FromHits) {
        if ($a[0] -lt $pos) { continue }
        while ($next -lt $ToHits.Count -and $ToHits[$next][0] -lt $a[1]) { $next++ }
        if ($next -ge $ToHits.Count) { $unclosed++; continue }
        [void]$pairs.Add(@($a[0], $a[1], $ToHits[$next][0], $ToHits[$next][1]))
        $pos = $ToHits[$next][1]
    }
    return @{ Pairs = $pairs; Unclosed = $unclosed }
}

function New-BetweenInfo($FromHits, $Found) {
    return [pscustomobject]@{
        FromFound = ($FromHits.Count -gt 0); ToFound = ($Found.Pairs.Count -gt 0); Unclosed = $Found.Unclosed
        Sections = (New-Object System.Collections.ArrayList); Count = 0; Words = 0; Blocked = 0 }
}

# How long the line or paragraph break is at a position, or just before it (0 for none)
function Get-BreakAt([scriptblock]$Char, [int]$Pos) {
    $c = & $Char $Pos
    if ($c -eq "`r" -and (& $Char ($Pos + 1)) -eq "`n") { return 2 }
    if ($c -match '^[\r\n\f\v]$') { return 1 }
    return 0
}

function Get-BreakBefore([scriptblock]$Char, [int]$Pos) {
    $c = & $Char ($Pos - 1)
    if ($c -eq "`n" -and (& $Char ($Pos - 2)) -eq "`r") { return 2 }
    if ($c -match '^[\r\n\f\v]$') { return 1 }
    return 0
}

# Moves the edges of a section about to be deleted, so that the text on each
# side joins up the way a person would leave it: no doubled spaces, no empty
# paragraph left behind, and no words run together. $Char gets the character at
# a position. $First and $TextEnd are the ends of the text that stays. $GrowLeft
# and $GrowRight say the section may take in more on that side, which it must
# not do into a phrase that stays. Without $CanJoin (inside a table cell) no
# mark can go. Returns the start and end.
function Get-TidySection([int]$Start, [int]$End, [scriptblock]$Char, [int]$First, [int]$TextEnd, [bool]$GrowLeft, [bool]$GrowRight, [bool]$CanJoin) {
    $blank = '^[ \t]$'
    $closing = '^[,.;:!?)\]]$'
    $atStart = ($Start -le $First) -or ((Get-BreakBefore $Char $Start) -gt 0)
    $atEnd = ($End -ge $TextEnd) -or ((Get-BreakAt $Char $End) -gt 0)
    if ($atStart -and $atEnd) {
        # A whole paragraph goes, so its mark goes too
        if ($GrowRight -and $CanJoin -and $End -lt $TextEnd) { $End += Get-BreakAt $Char $End }
        return @($Start, $End)
    }
    if ($atEnd) {
        # A section starting with a break keeps that break, and the next one goes
        # instead, so the paragraph before it keeps its own mark and look
        $firstBreak = Get-BreakAt $Char $Start
        if ($firstBreak -and $GrowRight -and $CanJoin) {
            $Start += $firstBreak
            if ($End -lt $TextEnd) { $End += Get-BreakAt $Char $End }
            return @($Start, $End)
        }
        while ($GrowLeft -and $Start -gt $First -and (& $Char ($Start - 1)) -match $blank) { $Start-- }
        return @($Start, $End)
    }
    if ($atStart) {
        while ($GrowRight -and $End -lt $TextEnd -and (& $Char $End) -match $blank) { $End++ }
        return @($Start, $End)
    }
    $left = & $Char ($Start - 1)
    $right = & $Char $End
    if ($left -match $blank) {
        if ($right -match $blank) {
            # One of the two spaces goes
            if ($GrowRight) { while ($End -lt $TextEnd -and (& $Char $End) -match $blank) { $End++ } }
            elseif ($GrowLeft) { while ($Start -gt $First -and (& $Char ($Start - 1)) -match $blank) { $Start-- } }
        } elseif ($right -match $closing -and $GrowLeft) {
            while ($Start -gt $First -and (& $Char ($Start - 1)) -match $blank) { $Start-- }
        }
        return @($Start, $End)
    }
    if (($right -match $blank) -or ($right -match $closing)) { return @($Start, $End) }
    # Text on both sides would touch: one break or space from the section's edges stays
    $edge = Get-BreakAt $Char $Start
    if ($edge) { return @(($Start + $edge), $End) }
    $edge = Get-BreakBefore $Char $End
    if ($edge) { return @($Start, ($End - $edge)) }
    if ((& $Char ($End - 1)) -match $blank) { return @($Start, ($End - 1)) }
    if ((& $Char $Start) -match $blank) { return @(($Start + 1), $End) }
    return @($Start, $End)
}

# Which words or phrases to replace, leaving out those in a section that goes
function Get-HitsOutsideSections($Hits, $Between) {
    if (-not $Between -or -not $Between.Sections.Count) { return ,$Hits }
    $kept = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($h in $Hits) {
        while ($i -lt $Between.Sections.Count -and $Between.Sections[$i].End -le $h[0]) { $i++ }
        if ($i -lt $Between.Sections.Count -and $Between.Sections[$i].Start -lt $h[1]) { continue }
        [void]$kept.Add($h)
    }
    return ,$kept
}

# True when a section can be deleted cleanly from a Word document: all of it
# outside tables (whole tables inside it just go), or all of it in one table cell
function Test-SectionCanGo($Doc, [int]$Start, [int]$End) {
    $startInTable = [bool]$Doc.Range($Start, $Start).Information(12)
    if ($startInTable -ne [bool]$Doc.Range($End, $End).Information(12)) { return $false }
    if (-not $startInTable) { return $true }
    # (A range ending at the very start of the next cell still counts as one cell, so compare positions)
    return ($End -lt $Doc.Range($Start, $Start).Cells.Item(1).Range.End)
}

# Deletes one section of a Word document. Range.Delete would keep a paragraph
# mark when a section runs from the middle of one paragraph into the middle of
# another, leaving a sentence split in two, so the text is cleared instead. That
# leaves tables behind, so tables wholly inside it go first (Range.Tables also
# gives the table around a section in a cell), and only Delete can take the last mark.
function Remove-WordSection($Doc, [int]$Start, [int]$End) {
    $r = $Doc.Range($Start, $End)
    foreach ($table in @($r.Tables)) {
        if ($table.Range.Start -ge $Start -and $table.Range.End -le $End) { $table.Delete() }
    }
    if ($r.End -ge $Doc.Content.End) { [void]$r.Delete() } else { $r.Text = '' }
}

# Finds the sections to delete from the main text of a Word document
function Get-WordSections($Doc, $Main, [int]$From, [int]$To, [hashtable]$B) {
    $docEnd = $Main.End
    $char = {
        param($Pos)
        if ($Pos -lt 0 -or $Pos -ge $docEnd) { return '' }
        $c = [string]$Doc.Range($Pos, $Pos + 1).Text
        # The end of a table cell comes back as two characters
        if ($c.Length -gt 1) { return $c.Substring(0, 1) }
        return $c
    }
    $fromHits = Find-WordHits $Main $From $To $B.From $B.MatchCase $false
    $found = Get-SectionPairs $fromHits (Find-WordHits $Main $From $To $B.To $B.MatchCase $false)
    $info = New-BetweenInfo $fromHits $found
    $lastEnd = $From
    foreach ($pair in $found.Pairs) {
        $s = $pair[1]; if ($B.IncludeFrom) { $s = $pair[0] }
        $e = $pair[2]; if ($B.IncludeTo) { $e = $pair[3] }
        if ($e -le $s) { continue }
        if (-not (Test-SectionCanGo $Doc $s $e)) { $info.Blocked++; continue }
        $span = Get-TidySection $s $e $char $From $To $B.IncludeFrom $B.IncludeTo (-not [bool]$Doc.Range($s, $s).Information(12))
        $s = [Math]::Max($span[0], $lastEnd)
        $e = $span[1]
        if ($e -le $s) { continue }
        # A whole last paragraph goes with the document's last mark, or an empty one stays
        if ($e -eq $docEnd - 1 -and $To -eq $docEnd - 1 -and ($s -eq $From -or (Get-BreakBefore $char $s) -gt 0)) { $e = $docEnd }
        [void]$info.Sections.Add([pscustomobject]@{ Start = $s; End = $e })
        $info.Words += $Doc.Range($s, $e).ComputeStatistics(0)
        $lastEnd = $e
    }
    $info.Count = $info.Sections.Count
    return $info
}

# Finds the sections to delete from a .txt file
function Get-TextSections([string]$Plain, [int]$From, [int]$To, [hashtable]$B) {
    $char = {
        param($Pos)
        if ($Pos -lt 0 -or $Pos -ge $Plain.Length) { return '' }
        return [string]$Plain[$Pos]
    }
    $fromHits = Find-TextHits $Plain $From $To $B.From $B.MatchCase $false
    $found = Get-SectionPairs $fromHits (Find-TextHits $Plain $From $To $B.To $B.MatchCase $false)
    $info = New-BetweenInfo $fromHits $found
    $lastEnd = $From
    foreach ($pair in $found.Pairs) {
        $s = $pair[1]; if ($B.IncludeFrom) { $s = $pair[0] }
        $e = $pair[2]; if ($B.IncludeTo) { $e = $pair[3] }
        if ($e -le $s) { continue }
        $span = Get-TidySection $s $e $char $From $To $B.IncludeFrom $B.IncludeTo $true
        $s = [Math]::Max($span[0], $lastEnd)
        $e = $span[1]
        if ($e -le $s) { continue }
        [void]$info.Sections.Add([pscustomobject]@{ Start = $s; End = $e })
        $info.Words += Get-WordCount $Plain.Substring($s, $e - $s)
        $lastEnd = $e
    }
    $info.Count = $info.Sections.Count
    return $info
}

# One line for the list of documents, saying which sections go
function Get-BetweenSummary($Info, [hashtable]$B) {
    $first = '"' + (Get-Short $B.From 24) + '"'
    $last = '"' + (Get-Short $B.To 24) + '"'
    if (-not $Info.FromFound) { return "sections: $first not found" }
    if (-not $Info.ToFound) { return "sections: $last not found after $first" }
    if ($Info.Count) { $line = (Get-Plural $Info.Count 'section' 'sections') + ' deleted (' + (Get-Plural $Info.Words 'word' 'words') + ')' }
    else { $line = 'sections: nothing to delete between the phrases' }
    if ($Info.Blocked) { $line += ' (' + (Get-Plural $Info.Blocked 'section crosses' 'sections cross') + ' the edge of a table, left as is)' }
    return $line
}

# The heading and explanation in the preview. Returns $true when there are sections to show.
function Add-BetweenNote($Plan, [hashtable]$B) {
    $first = '"' + (Get-Short $B.From 40) + '"'
    $last = '"' + (Get-Short $B.To 40) + '"'
    $info = $Plan.Between
    Add-Seg $Plan 'Title' "SECTIONS BETWEEN $first AND $last"
    if (-not $info.FromFound) { Add-Seg $Plan 'Note' "$first was not found (or only in a part that is cut off).`n"; return $false }
    if (-not $info.ToFound) { Add-Seg $Plan 'Note' "$last was not found after $first, so nothing is deleted.`n"; return $false }
    if ($info.Blocked) { Add-Seg $Plan 'Note' ((Get-Plural $info.Blocked 'section crosses' 'sections cross') + " the edge of a table, so it stays as it is.`n") }
    if ($info.Unclosed) { Add-Seg $Plan 'Note' ("$first is there " + (Get-Plural $info.Unclosed 'more time' 'more times') + " with no $last after it, and that stays.`n") }
    if (-not $info.Count) { Add-Seg $Plan 'Note' "There is nothing to delete between the phrases.`n"; return $false }
    Add-Seg $Plan 'Note' ((Get-Plural $info.Count 'section is' 'sections are') + " deleted:`n")
    return $true
}


# ---------------------------------------------------------------- Word documents

function Invoke-WordDocument($State, $Item, [hashtable]$S, [bool]$Apply, [string]$SaveMode, [string]$Stamp) {
    $word = Get-Word $State
    $target = $null
    if ($Apply) { $target = Get-SaveTarget $Item $SaveMode $Stamp }
    $openPath = $Item.Path
    if ($target) { $openPath = $target.Open }
    $doc = Open-InWord $word $openPath
    try {
        # Word opens a file it doesn't understand as plain text
        if (@(2, 3, 4, 5, 7) -contains $doc.SaveFormat) { throw 'NOTDOC:' }
        $plan = Get-WordPlan $State $doc $S (-not $Apply)
        if (-not $Apply) { return $plan }
        if (-not $plan.Changes) {
            Remove-UnneededBackup $target.Backup
            return (New-Result 'None' $plan.Summary)
        }
        $plan.WordsCount = Set-WordChanges $State $doc $plan $S
        $plan.Summary = Get-PlanSummary $plan $S
        # The Compatibility Checker would wait for an answer nobody can see
        try { $doc.CheckCompatibility = $false } catch {}
        # Word needs plain text here: a path straight from Join-Path makes SaveAs2 hang
        $m = [Type]::Missing   # no LockComments and no Password
        try { $doc.SaveAs2([string]$target.Save, [int]$doc.SaveFormat, $m, $m, $false) }
        catch { throw ('SAVE:' + (Get-InnerMessage $_.Exception)) }
        return (New-Result 'Done' $plan.Summary $target.Save $target.Backup)
    } finally {
        Close-WordDocument $doc
    }
}

# Works out what will change in a Word document. Positions are Word's
# character positions in the main text (Range.Start and Range.End).
function Get-WordPlan($State, $Doc, [hashtable]$S, [bool]$ForPreview) {
    $plan = New-Preview 'None' ''
    $main = $Doc.Content
    $keepStart = 0
    $keepEnd = $main.End - 1   # the document's last paragraph mark always stays
    if ($S.Begin.On) {
        $plan.Begin = Get-WordCut $Doc $main $keepStart $keepEnd $S.Begin 'Begin'
        if ($plan.Begin.Change) { $keepStart = $plan.Begin.Cut }
    }
    if ($S.End.On) {
        $plan.End = Get-WordCut $Doc $main $keepStart $keepEnd $S.End 'End'
        if ($plan.End.Change) { $keepEnd = $plan.End.Cut }
    }
    $plan.KeepStart = $keepStart
    $plan.KeepEnd = $keepEnd
    if ($plan.Begin -and $plan.Begin.Change -and $plan.Begin.Removes) { $plan.Begin.Words = $Doc.Range(0, $plan.Begin.Cut).ComputeStatistics(0) }
    if ($plan.End -and $plan.End.Change -and $plan.End.Removes) { $plan.End.Words = $Doc.Range($plan.End.Cut, $main.End).ComputeStatistics(0) }
    if ($S.Between.On) { $plan.Between = Get-WordSections $Doc $main $keepStart $keepEnd $S.Between }

    $hits = New-Object System.Collections.ArrayList
    if ($S.Words.On) { $plan.WordsNote = Get-StyleProblem $Doc $S.Words.NewFormat }
    if ($S.Words.On -and -not $plan.WordsNote) {
        $replacing = -not (Test-KeepsText $S.Words)
        foreach ($story in (Get-WordStories $Doc)) {
            # The last paragraph mark of each part always stays
            if ($story.Where) { $found = Find-WordHits $story.Range $story.Range.Start ($story.Range.End - 1) $S.Words.Find $S.Words.MatchCase $S.Words.WholeWords $replacing $S.Words.FindFormat }
            else { $found = Get-HitsOutsideSections (Find-WordHits $story.Range $keepStart $keepEnd $S.Words.Find $S.Words.MatchCase $S.Words.WholeWords $replacing $S.Words.FindFormat) $plan.Between }
            $plan.WordsCount += $found.Count
            foreach ($h in $found) {
                if ($hits.Count -ge $script:ShowWords) { break }
                [void]$hits.Add([pscustomobject]@{ Story = $story.Range; Where = $story.Where; Start = $h[0]; End = $h[1] })
            }
        }
    }
    $plan.Hits = $hits
    $plan.Changes = ($plan.Begin -and $plan.Begin.Change) -or ($plan.End -and $plan.End.Change) -or ($plan.Between -and $plan.Between.Count -gt 0) -or ($plan.WordsCount -gt 0)
    if ($plan.Changes) { $plan.Status = 'Change' }
    $plan.Summary = Get-PlanSummary $plan $S
    if ($ForPreview) { Add-WordSegments $State $Doc $plan $S }
    # Word's own objects are not kept once the document is closed
    $plan.Hits = $null
    return $plan
}

function New-CutInfo {
    return [pscustomobject]@{ Found = $false; Count = 0; Cut = 0; ParaMode = $false; NeedSpace = $false; Removes = $false; Change = $false; Problem = ''; Hit = $null; Words = 0 }
}

# Finds where the beginning ends (or the ending starts). Returns Found, Count
# (how many times the text is there), Cut (the position where the document is
# cut), ParaMode (the cut is at the start or end of a paragraph, so new text
# goes in paragraphs of its own), Removes, Change, Problem, Hit and Words.
function Get-WordCut($Doc, $Main, [int]$From, [int]$To, [hashtable]$C, [string]$Part) {
    $info = New-CutInfo
    $found = Find-WordHits $Main $From $To $C.Find $C.MatchCase $false
    $info.Count = $found.Count
    if (-not $found.Count) { return $info }
    $info.Found = $true
    if ($C.Which -eq 'Last') { $h = $found[$found.Count - 1] } else { $h = $found[0] }
    $info.Hit = $h
    if ($Part -eq 'Begin') {
        if ($C.Include) {
            # The text is replaced too: so are the spaces after it, or the end of its paragraph
            $cut = $h[1]
            while ($cut -lt $To -and (Get-DocChar $Doc $cut) -match '^[ \t]$') { $cut++ }
            if ($cut -lt $To -and (Get-DocChar $Doc $cut) -eq "`r") { $cut++ }
        } else { $cut = $h[0] }
        $info.ParaMode = ($cut -eq 0) -or ((Get-DocChar $Doc ($cut - 1)) -match '^[\r\f\v]$') -or (Test-MultiLineNew $C)
        $info.NeedSpace = (-not $info.ParaMode) -and ((Get-DocChar $Doc $cut) -match '^\S')
        $info.Removes = ($cut -gt 0)
    } else {
        if ($C.Include) {
            # The text is replaced too: so are the spaces in front of it
            $cut = $h[0]
            while ($cut -gt $From -and (Get-DocChar $Doc ($cut - 1)) -match '^[ \t]$') { $cut-- }
        } else { $cut = $h[1] }
        $before = "`r"
        if ($cut -gt 0) { $before = Get-DocChar $Doc ($cut - 1) }
        $info.ParaMode = ($before -match '^[\r\f\v]$') -or ((Get-DocChar $Doc $cut) -eq "`r") -or (Test-MultiLineNew $C)
        $info.NeedSpace = (-not $info.ParaMode) -and ($before -match '^\S')
        $info.Removes = ($cut -lt $To)
    }
    $info.Cut = $cut
    $info.Change = $info.Removes -or (Test-HasNewText $C)
    # Cutting a table in two would leave half a table behind
    if ($Doc.Range($cut, $cut).Information(12)) { $info.Problem = 'table'; $info.Change = $false }
    return $info
}

function Get-DocChar($Doc, [int]$Pos) {
    if ($Pos -lt 0) { return '' }
    return [string]$Doc.Range($Pos, $Pos + 1).Text
}

# New text of more than one line, or a whole document, goes in paragraphs of its own
function Test-MultiLineNew([hashtable]$C) {
    if ($C.Source -eq 'File') { return [bool]$C.File }
    return (Expand-Codes ([string]$C.Replace)) -match '[\r\n\v\f]'
}

function Test-HasNewText([hashtable]$C) {
    if ($C.Source -eq 'File') { return [bool]$C.File }
    return -not [string]::IsNullOrEmpty($C.Replace)
}

# Every place Word keeps text: the main text first (Where = ''), then
# footnotes, endnotes, comments, text boxes, headers and footers.
function Get-WordStories($Doc) {
    $list = New-Object System.Collections.ArrayList
    [void]$list.Add([pscustomobject]@{ Range = $Doc.Content; Where = '' })
    $names = @{ 2 = 'a footnote'; 3 = 'an endnote'; 4 = 'a comment'; 5 = 'a text box' }
    foreach ($sr in $Doc.StoryRanges) {
        $type = [int]$sr.StoryType
        if (-not $names.ContainsKey($type)) { continue }
        $r = $sr
        while ($r) {
            [void]$list.Add([pscustomobject]@{ Range = $r; Where = $names[$type] })
            $r = $r.NextStoryRange
        }
    }
    # Headers and footers. A section whose header is "Same as Previous" shares it.
    $first = $true
    foreach ($sec in $Doc.Sections) {
        foreach ($kind in 1, 2, 3) {
            foreach ($pair in @(@($sec.Headers.Item($kind), 'a header'), @($sec.Footers.Item($kind), 'a footer'))) {
                $hf = $pair[0]
                try {
                    if (-not $hf.Exists) { continue }
                    if (-not $first -and $hf.LinkToPrevious) { continue }
                    if ($hf.Range.Text.Length -le 1) { continue }
                    [void]$list.Add([pscustomobject]@{ Range = $hf.Range; Where = $pair[1] })
                } catch {}
            }
        }
        $first = $false
    }
    return ,$list
}

# Every place the text is found between $From and $To in one part of the
# document, as pairs of positions. With $WholeWords, only where it is not part
# of a longer word. With $Replacing, only what Word can really replace. $Format
# is what the text must look like: with no $Text, everything that looks like it.
function Find-WordHits($Story, [int]$From, [int]$To, [string]$Text, [bool]$MatchCase, [bool]$WholeWords, [bool]$Replacing = $false, $Format = $null) {
    $hits = New-Object System.Collections.ArrayList
    if ($To -le $From) { return ,$hits }
    $r = $Story.Duplicate
    $r.SetRange($From, $From)
    try { Set-FindOptions $r.Find $Text $MatchCase $Format }
    catch {
        # A document without the style has nothing in that style to find
        if ((Get-InnerMessage $_.Exception) -eq 'NOSTYLE:') { return ,$hits }
        throw
    }
    $last = -1
    while ($r.Find.Execute()) {
        $start = $r.Start
        $end = $r.End
        if ($start -le $last) { break }
        if ($end -gt $To) {
            # Text found by its format alone can run on to the last mark: the part before it counts
            if ($Text -or $start -ge $To) { break }
            $r.End = $To
            $end = $To
        }
        $last = $start
        $found = [string]$r.Text
        # Text found by its format alone can include the paragraph mark after
        # it. The mark stays, so that replacing the text doesn't join paragraphs.
        if (-not $Text -and $found.EndsWith("`r")) {
            $found = $found.Substring(0, $found.Length - 1)
            $end--
        }
        $ok = ($found.Length -gt 0) -and ((-not $WholeWords) -or (-not $Text) -or (Test-WordBoundary $Story $start $end $found))
        if ($ok -and $Replacing) { $ok = Test-CanReplace $Story $start $end $found }
        if ($ok) {
            [void]$hits.Add(@($start, $end))
            if ($hits.Count -ge $script:MaxHits) { break }
        }
        $r.Collapse(0)
    }
    return ,$hits
}

function Set-FindOptions($Find, [string]$Text, [bool]$MatchCase, $Format = $null) {
    $Find.ClearFormatting()
    $Find.Replacement.ClearFormatting()
    $Find.Text = ConvertTo-WordFindText $Text
    $Find.Replacement.Text = ''
    $Find.Forward = $true
    $Find.Wrap = 0              # stop at the end
    $Find.Format = $false
    $Find.MatchCase = $MatchCase
    $Find.MatchWholeWord = $false
    $Find.MatchWildcards = $false
    $Find.MatchSoundsLike = $false
    $Find.MatchAllWordForms = $false
    try { $Find.MatchPrefix = $false; $Find.MatchSuffix = $false; $Find.IgnorePunct = $false; $Find.IgnoreSpace = $false } catch {}
    if (Test-HasFormat $Format) {
        Set-FindFormat $Find $Format
        $Find.Format = $true
    }
}

# The text for Word to look for. ^ codes go to Word as they are (Word only knows
# them in small letters). A line break or tab in the text stands for the same mark.
function ConvertTo-WordFindText([string]$Text) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in (Get-CodeTokens $Text)) {
        if ($t.Kind -eq 'Code') { [void]$sb.Append('^' + $t.Value) }
        else { [void]$sb.Append($t.Value.Replace("`r`n", '^p').Replace("`n", '^p').Replace("`r", '^p').Replace("`t", '^t')) }
    }
    if ($sb.Length -gt 255) { throw 'TOOLONG:' }
    return $sb.ToString()
}

# Word needs some marks to stay: the end of a table cell, and a paragraph mark
# right before a table (replacing it would pull the text before it into the table)
function Test-CanReplace($Story, [int]$Start, [int]$End, [string]$Found) {
    if ($Found.IndexOf([char]7) -ge 0) { return $false }
    if (-not $Found.EndsWith("`r")) { return $true }
    $r = $Story.Duplicate
    # The mark that ends a footnote or endnote, right before the next one
    $type = [int]$Story.StoryType
    if ($type -eq 2 -or $type -eq 3) {
        $r.SetRange($End, $End + 1)
        if ([string]$r.Text -eq [string][char]2) { return $false }
    }
    $r.SetRange($End, $End)
    if (-not $r.Information(12)) { return $true }
    $r.SetRange($Start, $Start)
    return [bool]$r.Information(12)
}

function ConvertTo-WordText([string]$Text) {
    return $Text.Replace("`r`n", "`r").Replace("`n", "`r")
}

# True when the found text ($Text) is a whole word: not straight after or before a letter or number
function Test-WordBoundary($Story, [int]$Start, [int]$End, [string]$Text) {
    $r = $Story.Duplicate
    if ($Text -match '^[\p{L}\p{N}]' -and $Start -gt $Story.Start) {
        $r.SetRange($Start - 1, $Start)
        if ([string]$r.Text -match '[\p{L}\p{N}]$') { return $false }
    }
    if ($Text -match '[\p{L}\p{N}]$' -and $End -lt $Story.End) {
        $r.SetRange($End, $End + 1)
        if ([string]$r.Text -match '^[\p{L}\p{N}]') { return $false }
    }
    return $true
}

# Like Word: unless "Match case" is ticked, a replacement typed in small letters
# follows the found text - "Cat" becomes "Dog" and "CAT" becomes "DOG"
function Get-CaseAdapted([string]$Found, [string]$New, [bool]$MatchCase) {
    if ($MatchCase -or -not $New) { return $New }
    if ($New -cne $New.ToLowerInvariant()) { return $New }
    $letters = [regex]::Replace($Found, '[^\p{L}]', '')
    if ($letters.Length -ge 2 -and $letters -ceq $letters.ToUpperInvariant()) { return $New.ToUpperInvariant() }
    if ($letters.Length -ge 1 -and [char]::IsUpper($letters[0])) {
        $m = [regex]::Match($New, '\p{L}')
        if ($m.Success) { return $New.Substring(0, $m.Index) + $m.Value.ToUpperInvariant() + $New.Substring($m.Index + 1) }
    }
    return $New
}
# Makes the changes in the plan. Parts are cut from the back of the document to
# the front, so the positions of those still to go don't move: the ending, the
# sections between phrases, then the beginning. The words and phrases are
# replaced in what is left, and then the new beginning and ending are put in
# (so the word replacements don't change them). Returns how many words or
# phrases were replaced.
function Set-WordChanges($State, $Doc, $Plan, [hashtable]$S) {
    try { $Doc.TrackRevisions = $false } catch {}
    $endLook = $null; $beginLook = $null
    $endNew = $null; $beginNew = $null
    if ($Plan.End -and $Plan.End.Change) {
        $endNew = Get-NewPart $State $S.End
        if ($Plan.End.ParaMode -and $endNew.Kind -eq 'Text') { $endLook = Get-BodyLook $Doc ([Math]::Max(0, $Plan.End.Cut - 1)) $false $Plan.KeepStart }
        if ($Plan.End.Removes) { [void]$Doc.Range($Plan.End.Cut, $Doc.Content.End).Delete() }
    }
    if ($Plan.Between) {
        $gone = 0
        for ($i = $Plan.Between.Sections.Count - 1; $i -ge 0; $i--) {
            $sec = $Plan.Between.Sections[$i]
            try { Remove-WordSection $Doc $sec.Start $sec.End; $gone++ } catch {}
        }
        $Plan.Between.Count = $gone
    }
    if ($Plan.Begin -and $Plan.Begin.Change) {
        $beginNew = Get-NewPart $State $S.Begin
        if ($Plan.Begin.ParaMode -and $beginNew.Kind -eq 'Text') { $beginLook = Get-BodyLook $Doc $Plan.Begin.Cut $true $Doc.Content.End }
        if ($Plan.Begin.Removes) { [void]$Doc.Range(0, $Plan.Begin.Cut).Delete() }
    }
    $count = 0
    if ($S.Words.On -and -not $Plan.WordsNote) {
        foreach ($story in (Get-WordStories $Doc)) { $count += Invoke-WordReplaceAll $story.Range $S.Words }
    }
    if ($beginNew) { Add-WordBeginning $Doc $Plan.Begin $beginNew $beginLook }
    if ($endNew) { Add-WordEnding $Doc $Plan.End $endNew $endLook }
    return $count
}

# Goes from the last find to the first, so the ones still to do don't move and
# nothing is ever looked at twice. Word refuses a few marks (such as the end of
# a footnote): those just stay, and are not counted.
function Invoke-WordReplaceAll($Story, [hashtable]$W) {
    $keepText = Test-KeepsText $W
    $hits = Find-WordHits $Story $Story.Start ($Story.End - 1) $W.Find $W.MatchCase $W.WholeWords (-not $keepText) $W.FindFormat
    $r = $Story.Duplicate
    $count = 0
    for ($i = $hits.Count - 1; $i -ge 0; $i--) {
        $r.SetRange($hits[$i][0], $hits[$i][1])
        try {
            if (-not $keepText) { $r.Text = Get-NewText ([string]$r.Text) $W.Replace $W.MatchCase }
            Set-RangeFormat $r $W.NewFormat
            $count++
        } catch {}
    }
    return $count
}

# What goes in place of the beginning or ending: Kind 'Text' (typed, or the
# words of a .txt file), 'File' (a Word document, put in with its formatting)
# or 'None' (nothing: the part is just deleted)
function Get-NewPart($State, [hashtable]$C) {
    if ($C.Source -eq 'File' -and $C.File) {
        if ([IO.Path]::GetExtension($C.File) -ieq '.txt') { return @{ Kind = 'Text'; Text = (Get-SourceText $State $C.File) } }
        return @{ Kind = 'File'; Path = $C.File }
    }
    if ($C.Replace) { return @{ Kind = 'Text'; Text = (Expand-Codes $C.Replace) } }
    return @{ Kind = 'None' }
}

# The look (paragraph style, paragraph settings and font) that typed text gets
# when it goes in paragraphs of its own: that of the nearest paragraph of
# ordinary text (not a heading) that stays - after the cut for the beginning,
# before it for the ending
function Get-BodyLook($Doc, [int]$Pos, [bool]$Forward, [int]$Limit) {
    try {
        $p = $Doc.Range($Pos, $Pos).Paragraphs.Item(1)
        $first = $p
        $pick = $null
        for ($i = 0; $i -lt 30 -and $p; $i++) {
            if ($Forward -and $p.Range.Start -ge $Limit) { break }
            if (-not $Forward -and $p.Range.End -le $Limit) { break }
            if ($p.OutlineLevel -eq 10 -and $p.Range.Text -match '[\p{L}\p{N}]') { $pick = $p; break }
            if ($Forward) { $p = $p.Next(1) } else { $p = $p.Previous(1) }
        }
        if (-not $pick) { $pick = $first }
        $chars = $pick.Range.Characters
        $c = $chars.Item(1)
        foreach ($ch in $chars) { if ($ch.Text -match '[\p{L}\p{N}]') { $c = $ch; break } }
        return @{ Style = $pick.Style; Para = $pick.Range.ParagraphFormat.Duplicate; Font = $c.Font.Duplicate }
    } catch { return $null }
}

function Set-Look($Range, $Look) {
    if (-not $Look) { return }
    try { $Range.Style = $Look.Style } catch {}
    try { $Range.ParagraphFormat = $Look.Para } catch {}
    try { $Range.Font = $Look.Font } catch {}
}

function Add-WordBeginning($Doc, $Cut, $New, $Look) {
    if ($New.Kind -eq 'File') {
        $r = $Doc.Range(0, 0)
        try { $r.InsertFile([string]$New.Path, '', $false, $false, $false) } catch { throw ('INSERT:' + (Get-InnerMessage $_.Exception)) }
        return
    }
    if ($New.Kind -ne 'Text' -or -not $New.Text) { return }
    $t = ConvertTo-WordText $New.Text
    if ($Cut.ParaMode -and -not $t.EndsWith("`r")) { $t += "`r" }
    elseif ($Cut.NeedSpace -and $t -notmatch '\s$') { $t += ' ' }
    $r = $Doc.Range(0, 0)
    $r.InsertBefore($t)
    if ($Cut.ParaMode) { Set-Look $r $Look }
}

function Add-WordEnding($Doc, $Cut, $New, $Look) {
    if ($New.Kind -eq 'File') {
        if ($Doc.Content.End -gt 1) { $Doc.Content.InsertParagraphAfter() }
        $e = $Doc.Content.End
        $r = $Doc.Range($e - 1, $e - 1)
        try { $r.InsertFile([string]$New.Path, '', $false, $false, $false) } catch { throw ('INSERT:' + (Get-InnerMessage $_.Exception)) }
        # The document put in brings its own last paragraph mark: drop the empty paragraph left over
        $lastPara = $Doc.Paragraphs.Last.Range
        if ($Doc.Paragraphs.Count -gt 1 -and $lastPara.Text -eq "`r") { [void]$lastPara.Delete() }
        return
    }
    if ($New.Kind -ne 'Text' -or -not $New.Text) { return }
    $t = (ConvertTo-WordText $New.Text).TrimEnd("`r")
    if (-not $t) { return }
    $e = $Doc.Content.End
    $hasText = ($e -gt 1)
    if ($Cut.ParaMode -and $hasText) { $t = "`r" + $t }
    elseif ($Cut.NeedSpace -and $t -notmatch '^\s') { $t = ' ' + $t }
    $r = $Doc.Range($e - 1, $e - 1)
    $r.InsertAfter($t)
    if ($Cut.ParaMode) {
        $start = $r.Start
        if ($hasText) { $start++ }
        Set-Look ($Doc.Range($start, $Doc.Content.End)) $Look
    }
}


# ---------------------------------------------------------------- what the preview shows

# The preview is a list of pieces, each with a Kind:
#   Title - a heading,  Text - text that stays,  Old - text that goes,
#   New - text that comes in,  Gap - "... more ...",  Note - an explanation
function Add-Seg($Plan, [string]$Kind, [string]$Text) {
    [void]$Plan.Segments.Add([pscustomobject]@{ Kind = $Kind; Text = $Text })
}

# Makes Word's text readable in the preview
function Format-ShowText([string]$Text) {
    $t = $Text.Replace("`r" + [char]7, "`t").Replace([string][char]7, "`t")
    $t = $t.Replace("`r`n", "`n").Replace("`r", "`n").Replace([string][char]11, "`n").Replace([string][char]12, "`n")
    $t = $t.Replace([string][char]30, '-').Replace([string][char]31, '')
    return [regex]::Replace($t, '[\x00-\x08\x0E-\x1F]', '')
}

# Shows marks that can't be seen, in what is found and what replaces it
function Format-MarkText([string]$Text) {
    $t = $Text.Replace("`r`n", '[paragraph]').Replace("`r", '[paragraph]').Replace("`n", '[paragraph]').Replace("`t", '[tab]')
    $t = $t.Replace([string][char]11, '[line break]').Replace([string][char]12, '[page break]').Replace([string][char]14, '[column break]')
    $t = $t.Replace([string][char]0xA0, '[non-breaking space]').Replace([string][char]30, '[non-breaking hyphen]').Replace([string][char]0x2011, '[non-breaking hyphen]')
    $t = $t.Replace([string][char]31, '[optional hyphen]').Replace([string][char]0xAD, '[optional hyphen]')
    $t = $t.Replace([string][char]1, '[picture]').Replace([string][char]2, '[footnote mark]')
    return [regex]::Replace($t, '[\x00-\x08\x0E-\x1F]', '')
}

function Format-Count([int]$N) { return $N.ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US')) }

function Get-Plural([int]$N, [string]$One, [string]$Many) {
    if ($N -eq 1) { return "1 $One" }
    return "$(Format-Count $N) $Many"
}

function Get-Short([string]$Text, [int]$Max = 30) {
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max - 3).TrimEnd() + '...' }
    return $t
}

# One line for the list of documents, saying what will change
function Get-PlanSummary($Plan, [hashtable]$S) {
    $parts = @()
    if ($S.Begin.On) { $parts += Get-CutSummary $Plan.Begin $S.Begin 'beginning' }
    if ($S.Between.On) { $parts += Get-BetweenSummary $Plan.Between $S.Between }
    if ($S.Words.On) {
        $q = Get-WordsLabel $S.Words 24
        $verb = 'replaced'
        if (Test-KeepsText $S.Words) { $verb = 'reformatted' }
        if ($Plan.WordsCount) {
            $line = "$q $verb $(Get-Plural $Plan.WordsCount 'time' 'times')"
            if ($Plan.WordsNote) { $line += " ($($Plan.WordsNote))" }
            $parts += $line
        }
        elseif ($Plan.WordsNote) { $parts += "${q}: $($Plan.WordsNote)" }
        else { $parts += "$q not found" }
    }
    if ($S.End.On) { $parts += Get-CutSummary $Plan.End $S.End 'ending' }
    return ($parts -join ';  ')
}

function Get-CutSummary($Cut, [hashtable]$C, [string]$Part) {
    if (-not $Cut -or -not $Cut.Found) { return "${Part}: text not found" }
    if ($Cut.Problem -eq 'table') { return "${Part}: the text is inside a table (left as it is)" }
    if (-not $Cut.Change) {
        if ($Part -eq 'beginning') { return 'beginning: nothing before the text' }
        return 'ending: nothing after the text'
    }
    if (-not $Cut.Removes) { return "new $Part added" }
    $size = Get-Plural $Cut.Words 'word' 'words'
    if (Test-HasNewText $C) { return "$Part replaced ($size)" }
    return "$Part deleted ($size)"
}

function Add-WordSegments($State, $Doc, $Plan, [hashtable]$S) {
    $main = $Doc.Content
    if ($S.Begin.On) {
        Add-Seg $Plan 'Title' 'THE BEGINNING'
        $b = $Plan.Begin
        if (Add-CutNote $Plan $b $S.Begin 'Begin') {
            if ($b.Removes) { Add-RangeSegs $Plan $main 0 $b.Cut 'Old' }
            Add-NewSeg $State $Plan $S.Begin
            Add-Seg $Plan 'Text' ((Format-ShowText (Get-StoryText $main $b.Cut ([Math]::Min($b.Cut + $script:ShowKept, $Plan.KeepEnd)))) + ' ...')
        }
    }
    if ($S.Between.On -and (Add-BetweenNote $Plan $S.Between)) {
        $shown = 0
        foreach ($sec in $Plan.Between.Sections) {
            if ($shown -ge $script:ShowWords) { break }
            $shown++
            $stop = [Math]::Min($sec.End, $main.End - 1)
            $s0 = [Math]::Max($Plan.KeepStart, $sec.Start - $script:ShowAround)
            $e1 = [Math]::Min($Plan.KeepEnd, $stop + $script:ShowAround)
            Add-Seg $Plan 'Text' ('... ' + ((Format-ShowText (Get-StoryText $main $s0 $sec.Start)) -replace '\s+', ' '))
            Add-RangeSegs $Plan $main $sec.Start $stop 'Old'
            Add-Seg $Plan 'Text' (((Format-ShowText (Get-StoryText $main $stop $e1)) -replace '\s+', ' ') + " ...`n")
        }
        if ($Plan.Between.Count -gt $shown) { Add-Seg $Plan 'Gap' ('... and ' + (Format-Count ($Plan.Between.Count - $shown)) + " more`n") }
    }
    if ($S.Words.On) {
        Add-Seg $Plan 'Title' ('WORDS, PHRASES AND MARKS: ' + (Get-WordsLabel $S.Words 40))
        if ($Plan.WordsNote) { Add-Seg $Plan 'Note' "Nothing is changed here, because $($Plan.WordsNote)." }
        elseif (-not $Plan.WordsCount) { Add-Seg $Plan 'Note' 'Not found in this document (or only in a part that is cut off).' }
        else {
            Add-Seg $Plan 'Note' ('Found ' + (Get-Plural $Plan.WordsCount 'time' 'times') + ":`n")
            foreach ($h in $Plan.Hits) {
                if ($h.Where) { Add-Seg $Plan 'Note' ("In $($h.Where): ") }
                # Around a find in the main text, only show text that stays
                $lo = $h.Story.Start; $hi = $h.Story.End
                if (-not $h.Where) { $lo = $Plan.KeepStart; $hi = $Plan.KeepEnd }
                $s0 = [Math]::Max($lo, $h.Start - $script:ShowAround)
                $e1 = [Math]::Min($hi, $h.End + $script:ShowAround)
                Add-HitSegs $Plan $S.Words (Format-ShowText (Get-StoryText $h.Story $s0 $h.Start)) (Get-StoryText $h.Story $h.Start $h.End) (Format-ShowText (Get-StoryText $h.Story $h.End $e1)) '' $true
            }
            if ($Plan.WordsCount -gt $Plan.Hits.Count) { Add-Seg $Plan 'Gap' ('... and ' + (Format-Count ($Plan.WordsCount - $Plan.Hits.Count)) + " more`n") }
        }
    }
    if ($S.End.On) {
        Add-Seg $Plan 'Title' 'THE ENDING'
        $e = $Plan.End
        if (Add-CutNote $Plan $e $S.End 'End') {
            Add-Seg $Plan 'Text' ('... ' + (Format-ShowText (Get-StoryText $main ([Math]::Max($Plan.KeepStart, $e.Cut - $script:ShowKept)) $e.Cut)))
            if ($e.Removes) { Add-RangeSegs $Plan $main $e.Cut ($main.End - 1) 'Old' }
            Add-NewSeg $State $Plan $S.End
        }
    }
}

# One find in the preview, with a little of the text around it. Text that is
# only given a new format shows once, with the new format after it.
function Add-HitSegs($Plan, $W, [string]$Before, [string]$Found, [string]$After, [string]$Eol, [bool]$ShowFormat) {
    Add-Seg $Plan 'Text' ('... ' + ($Before -replace '\s+', ' '))
    $shown = Format-MarkText $Found
    if ($shown.Length -gt 240) { $shown = $shown.Substring(0, 120) + ' ... ' + $shown.Substring($shown.Length - 100) }
    $format = ''
    if ($ShowFormat -and (Test-HasFormat $W.NewFormat)) { $format = ' [' + (Get-FormatText $W.NewFormat) + ']' }
    if (Test-KeepsText $W) { Add-Seg $Plan 'Format' ($shown + $format) }
    else {
        Add-Seg $Plan 'Old' $shown
        Add-Seg $Plan 'New' (Format-MarkText (Get-NewText $Found $W.Replace $W.MatchCase $Eol))
        if ($format) { Add-Seg $Plan 'Format' $format }
    }
    Add-Seg $Plan 'Text' (($After -replace '\s+', ' ') + " ...`n")
}

# The explanation above a beginning or ending. Returns $true when the
# preview should go on to show the change.
function Add-CutNote($Plan, $Cut, [hashtable]$C, [string]$Part) {
    $q = '"' + (Get-Short $C.Find 60) + '"'
    if (-not $Cut.Found) {
        if ($Part -eq 'Begin') { Add-Seg $Plan 'Note' "$q was not found, so the beginning stays as it is.`n" }
        else { Add-Seg $Plan 'Note' "$q was not found after the beginning, so the ending stays as it is.`n" }
        return $false
    }
    if ($Cut.Problem -eq 'table') {
        Add-Seg $Plan 'Note' "$q is inside a table. The document can't be cut there, so it stays as it is.`n"
        return $false
    }
    $times = ''
    if ($Cut.Count -gt 1) {
        $which = 'first'; if ($C.Which -eq 'Last') { $which = 'last' }
        $times = "$q is there $(Get-Plural $Cut.Count 'time' 'times'); the $which one is used. "
    }
    if (-not $Cut.Change) {
        if ($Part -eq 'Begin') { Add-Seg $Plan 'Note' ($times + "$q is at the very start, so there is nothing before it to replace.`n") }
        else { Add-Seg $Plan 'Note' ($times + "$q is at the very end, so there is nothing after it to replace.`n") }
        return $false
    }
    if ($Part -eq 'Begin') { $what = 'Everything before it'; if ($C.Include) { $what = 'It and everything before it' } }
    else { $what = 'Everything after it'; if ($C.Include) { $what = 'It and everything after it' } }
    $how = 'is deleted'
    if (Test-HasNewText $C) { $how = 'is replaced' }
    if (-not $Cut.Removes) { Add-Seg $Plan 'Note' ($times + "There is nothing to remove, so the new text is just added:`n") }
    else { Add-Seg $Plan 'Note' ($times + "$what $how" + ":`n") }
    return $true
}

# A removed part: its start and its end, with the middle left out when it is long
function Add-RangeSegs($Plan, $Story, [int]$Start, [int]$End, [string]$Kind) {
    $n = $End - $Start
    if ($n -le 2 * $script:ShowEdge + 100) { Add-Seg $Plan $Kind (Format-ShowText (Get-StoryText $Story $Start $End)); return }
    Add-Seg $Plan $Kind (Format-ShowText (Get-StoryText $Story $Start ($Start + $script:ShowEdge)))
    Add-Seg $Plan 'Gap' ("`n ... " + (Format-Count ($n - 2 * $script:ShowEdge)) + " more characters ... `n")
    Add-Seg $Plan $Kind (Format-ShowText (Get-StoryText $Story ($End - $script:ShowEdge) $End))
}

function Add-NewSeg($State, $Plan, [hashtable]$C) {
    if (-not (Test-HasNewText $C)) { Add-Seg $Plan 'Note' "`n(nothing is put in its place)`n"; return }
    if ($C.Source -eq 'File') {
        if ([IO.Path]::GetExtension($C.File) -ine '.txt') {
            Add-Seg $Plan 'New' ("`n[the contents of " + [IO.Path]::GetFileName($C.File) + ', with its own formatting]')
            Add-Seg $Plan 'Text' "`n"
            return
        }
        $text = Get-SourceText $State $C.File
    } else { $text = Expand-Codes $C.Replace }
    # Paragraphs show as new lines, and other marks by name
    $lines = $text.Replace("`r`n", "`n").Replace("`r", "`n").Split("`n")
    Add-Seg $Plan 'New' (($lines | ForEach-Object { Format-MarkText $_ }) -join "`n")
    Add-Seg $Plan 'Text' "`n"
}

function Get-StoryText($Story, [int]$Start, [int]$End) {
    if ($End -le $Start) { return '' }
    $r = $Story.Duplicate
    $r.SetRange($Start, $End)
    return [string]$r.Text
}

# The words of a .txt file or Word document chosen as "the contents of a
# document", for putting into .txt files. Read once, then remembered.
function Get-SourceText($State, [string]$Path) {
    if ($State.Sources.ContainsKey($Path)) { return $State.Sources[$Path] }
    if ([IO.Path]::GetExtension($Path) -ieq '.txt') { $text = (Read-TextDocument $Path).Text }
    else {
        $doc = Open-InWord (Get-Word $State) $Path
        try { $text = Format-ShowText ([string]$doc.Content.Text) } finally { Close-WordDocument $doc }
        $text = $text.TrimEnd("`n").Replace("`n", "`r`n")
    }
    $State.Sources[$Path] = $text
    return $text
}

# ---------------------------------------------------------------- plain text (.txt) files

function Invoke-TextDocument($State, $Item, [hashtable]$S, [bool]$Apply, [string]$SaveMode, [string]$Stamp) {
    $target = $null
    if ($Apply) { $target = Get-SaveTarget $Item $SaveMode $Stamp }
    $openPath = $Item.Path
    if ($target) { $openPath = $target.Open }
    $doc = Read-TextDocument $openPath
    $plan = Get-TextPlan $State $doc $S (-not $Apply)
    if (-not $Apply) { return $plan }
    if (-not $plan.Changes) {
        Remove-UnneededBackup $target.Backup
        return (New-Result 'None' $plan.Summary)
    }
    $newText = Get-TextResult $State $doc $plan $S
    $note = Save-TextDocument $doc $newText $target.Save
    $detail = $plan.Summary
    if ($note) { $detail += ' (' + $note + ')' }
    return (New-Result 'Done' $detail $target.Save $target.Backup)
}

# Reads a whole .txt file. Remembers how it was saved (its alphabet, whether it
# starts with a Unicode marker, and its kind of line breaks), so that it can be
# saved the same way.
function Read-TextDocument([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $cp = Get-TextEncoding $bytes
    $enc = [Text.Encoding]::GetEncoding($cp)
    $pre = $enc.GetPreamble()
    $bom = $false
    if ($pre.Length -and $bytes.Length -ge $pre.Length) {
        $bom = $true
        for ($i = 0; $i -lt $pre.Length; $i++) { if ($bytes[$i] -ne $pre[$i]) { $bom = $false; break } }
    }
    $skip = 0; if ($bom) { $skip = $pre.Length }
    $text = $enc.GetString($bytes, $skip, $bytes.Length - $skip)
    if (Test-ScrambledText $text) { throw 'NOTTEXT:' }
    $eol = "`r`n"
    if (-not $text.Contains("`r`n")) {
        if ($text.Contains("`n")) { $eol = "`n" } elseif ($text.Contains("`r")) { $eol = "`r" }
    }
    return [pscustomobject]@{ Text = $text; CodePage = $cp; Bom = $bom; Eol = $eol }
}

# Plain text files don't say which alphabet they were saved in, so look at the bytes
function Get-TextEncoding([byte[]]$Bytes) {
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { return 65001 }   # UTF-8
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return 1200 }   # Unicode (UTF-16)
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return 1201 }
    # Unicode without a marker at the start: for English text, every other byte is zero
    $check = [Math]::Min($Bytes.Length, 4000)
    $oddZeros = 0; $evenZeros = 0
    for ($i = 0; $i -lt $check; $i++) { if ($Bytes[$i] -eq 0) { if ($i % 2) { $oddZeros++ } else { $evenZeros++ } } }
    if ($check -ge 4 -and $oddZeros * 4 -gt $check) { return 1200 }
    if ($check -ge 4 -and $evenZeros * 4 -gt $check) { return 1201 }
    # UTF-8 follows strict rules for every letter outside plain English
    try { [void](New-Object Text.UTF8Encoding($false, $true)).GetString($Bytes); return 65001 } catch {}
    return [Text.Encoding]::Default.CodePage   # the PC's usual Windows alphabet, such as 1252
}

# A file that isn't really text comes out as scrambled symbols
function Test-ScrambledText([string]$Text) {
    if ($Text.Length -lt 20) { return $false }
    if ($Text.Length -gt 4000) { $Text = $Text.Substring(0, 4000) }
    $odd = [regex]::Matches($Text, '[\x00\x03\x04\x06\x0E-\x12\x16-\x1F]').Count
    return ($odd * 20 -gt $Text.Length)   # more than 5% junk characters
}

# Saves the text the way the file was saved before. If the new text has letters
# the old alphabet can't hold, it is saved as UTF-8 instead (and returns a note saying so).
function Save-TextDocument($Doc, [string]$Text, [string]$Path) {
    $note = ''
    $enc = [Text.Encoding]::GetEncoding($Doc.CodePage, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ReplacementFallback)
    $bom = $Doc.Bom
    try { $body = $enc.GetBytes($Text) }
    catch {
        $enc = New-Object Text.UTF8Encoding($true)
        $body = $enc.GetBytes($Text)
        $bom = $true
        $note = 'saved as Unicode (UTF-8), because the new text has letters the old format could not hold'
    }
    $pre = [byte[]]@()
    if ($bom) { $pre = $enc.GetPreamble() }
    $all = New-Object byte[] ($pre.Length + $body.Length)
    [Array]::Copy($pre, 0, $all, 0, $pre.Length)
    [Array]::Copy($body, 0, $all, $pre.Length, $body.Length)
    Clear-ReadOnly $Path
    [IO.File]::WriteAllBytes($Path, $all)
    return $note
}

# A search pattern for the text as typed. Straight and curly quotes count as
# the same, a typed line break matches any kind of line break, and ^ codes work.
function New-FindRegex([string]$Text, [bool]$MatchCase, [bool]$WholeWords) {
    $single = "['" + [char]0x2018 + [char]0x2019 + ']'
    $double = '["' + [char]0x201C + [char]0x201D + ']'
    $sb = New-Object System.Text.StringBuilder
    $tokens = Get-CodeTokens $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    foreach ($t in $tokens) {
        if ($t.Kind -eq 'Code') { [void]$sb.Append((Get-CodeRegex $t.Value)); continue }
        foreach ($ch in $t.Value.ToCharArray()) {
            $code = [int]$ch
            if ($code -eq 0x27 -or $code -eq 0x2018 -or $code -eq 0x2019) { [void]$sb.Append($single) }
            elseif ($code -eq 0x22 -or $code -eq 0x201C -or $code -eq 0x201D) { [void]$sb.Append($double) }
            elseif ($code -eq 10) { [void]$sb.Append('(?>\r\n|\n|\r)') }
            else { [void]$sb.Append([regex]::Escape([string]$ch)) }
        }
    }
    $pattern = $sb.ToString()
    if ($WholeWords -and $tokens.Count) {
        $first = $tokens[0]; $last = $tokens[$tokens.Count - 1]
        if ($first.Kind -eq 'Text' -and $first.Value -match '^[\p{L}\p{N}]') { $pattern = '(?<![\p{L}\p{N}])' + $pattern }
        if ($last.Kind -eq 'Text' -and $last.Value -match '[\p{L}\p{N}]$') { $pattern = $pattern + '(?![\p{L}\p{N}])' }
    }
    $opts = [Text.RegularExpressions.RegexOptions]::CultureInvariant
    if (-not $MatchCase) { $opts = $opts -bor [Text.RegularExpressions.RegexOptions]::IgnoreCase }
    return New-Object Text.RegularExpressions.Regex($pattern, $opts)
}

function Find-TextHits([string]$Text, [int]$From, [int]$To, [string]$Find, [bool]$MatchCase, [bool]$WholeWords) {
    $hits = New-Object System.Collections.ArrayList
    if ($To -le $From) { return ,$hits }
    $rx = New-FindRegex $Find $MatchCase $WholeWords
    $m = $rx.Match($Text, $From)
    while ($m.Success -and $m.Index + $m.Length -le $To -and $hits.Count -lt $script:MaxHits) {
        if ($m.Length -eq 0) { break }
        [void]$hits.Add(@($m.Index, ($m.Index + $m.Length)))
        $m = $m.NextMatch()
    }
    return ,$hits
}

function Test-LineBreak([string]$Text, [int]$Pos) {
    if ($Pos -lt 0 -or $Pos -ge $Text.Length) { return $false }
    return ($Text[$Pos] -eq "`n" -or $Text[$Pos] -eq "`r" -or $Text[$Pos] -eq [char]12)
}

# The same as Get-WordCut, for the text of a .txt file
function Get-TextCut([string]$Text, [int]$From, [int]$To, [hashtable]$C, [string]$Part) {
    $info = New-CutInfo
    $found = Find-TextHits $Text $From $To $C.Find $C.MatchCase $false
    $info.Count = $found.Count
    if (-not $found.Count) { return $info }
    $info.Found = $true
    if ($C.Which -eq 'Last') { $h = $found[$found.Count - 1] } else { $h = $found[0] }
    $info.Hit = $h
    if ($Part -eq 'Begin') {
        if ($C.Include) {
            $cut = $h[1]
            while ($cut -lt $To -and ($Text[$cut] -eq ' ' -or $Text[$cut] -eq "`t")) { $cut++ }
            if ($cut -lt $To -and $Text[$cut] -eq "`r") { $cut++; if ($cut -lt $To -and $Text[$cut] -eq "`n") { $cut++ } }
            elseif ($cut -lt $To -and $Text[$cut] -eq "`n") { $cut++ }
        } else { $cut = $h[0] }
        $info.ParaMode = ($cut -eq 0) -or (Test-LineBreak $Text ($cut - 1)) -or (Test-MultiLineNew $C)
        $info.NeedSpace = (-not $info.ParaMode) -and $cut -lt $Text.Length -and -not [char]::IsWhiteSpace($Text[$cut])
        $info.Removes = ($cut -gt 0)
    } else {
        if ($C.Include) {
            $cut = $h[0]
            while ($cut -gt $From -and ($Text[$cut - 1] -eq ' ' -or $Text[$cut - 1] -eq "`t")) { $cut-- }
        } else { $cut = $h[1] }
        $info.ParaMode = ($cut -eq 0) -or (Test-LineBreak $Text ($cut - 1)) -or (Test-LineBreak $Text $cut) -or (Test-MultiLineNew $C)
        $info.NeedSpace = (-not $info.ParaMode) -and $cut -gt 0 -and -not [char]::IsWhiteSpace($Text[$cut - 1])
        $info.Removes = ($cut -lt $To)
    }
    $info.Cut = $cut
    $info.Change = $info.Removes -or (Test-HasNewText $C)
    return $info
}

function Get-TextPlan($State, $Doc, [hashtable]$S, [bool]$ForPreview) {
    $plan = New-Preview 'None' ''
    $text = $Doc.Text
    $keepStart = 0
    $keepEnd = $text.Length
    if ($S.Begin.On) {
        $plan.Begin = Get-TextCut $text $keepStart $keepEnd $S.Begin 'Begin'
        if ($plan.Begin.Change) { $keepStart = $plan.Begin.Cut }
    }
    if ($S.End.On) {
        $plan.End = Get-TextCut $text $keepStart $keepEnd $S.End 'End'
        if ($plan.End.Change) { $keepEnd = $plan.End.Cut }
    }
    $plan.KeepStart = $keepStart
    $plan.KeepEnd = $keepEnd
    if ($plan.Begin -and $plan.Begin.Removes -and $plan.Begin.Change) { $plan.Begin.Words = Get-WordCount $text.Substring(0, $plan.Begin.Cut) }
    if ($plan.End -and $plan.End.Removes -and $plan.End.Change) { $plan.End.Words = Get-WordCount $text.Substring($plan.End.Cut) }
    if ($S.Between.On) { $plan.Between = Get-TextSections $text $keepStart $keepEnd $S.Between }
    $plan.Hits = New-Object System.Collections.ArrayList
    # A .txt file has no formatting: nothing in it can be found by its format or be given one
    if ($S.Words.On -and ((Test-HasFormat $S.Words.FindFormat) -or (Test-KeepsText $S.Words))) { $plan.WordsNote = 'a .txt file has no formatting' }
    elseif ($S.Words.On) {
        $plan.Hits = Get-HitsOutsideSections (Find-TextHits $text $keepStart $keepEnd $S.Words.Find $S.Words.MatchCase $S.Words.WholeWords) $plan.Between
        if ($plan.Hits.Count -and (Test-HasFormat $S.Words.NewFormat)) { $plan.WordsNote = 'a .txt file has no formatting, so the new format is left out' }
        # Like the last paragraph mark in Word, the file's last line break always stays
        $n = $plan.Hits.Count
        if ($n -and $plan.Hits[$n - 1][1] -eq $text.Length -and (Test-LineBreak $text ($text.Length - 1))) { $plan.Hits.RemoveAt($n - 1) }
        $plan.WordsCount = $plan.Hits.Count
    }
    $plan.Changes = ($plan.Begin -and $plan.Begin.Change) -or ($plan.End -and $plan.End.Change) -or ($plan.Between -and $plan.Between.Count -gt 0) -or ($plan.WordsCount -gt 0)
    if ($plan.Changes) { $plan.Status = 'Change' }
    $plan.Summary = Get-PlanSummary $plan $S
    if ($ForPreview) { Add-TextSegments $State $text $plan $S }
    return $plan
}

function Get-WordCount([string]$Text) {
    $joins = "'" + [char]0x2019 + '-'
    return [regex]::Matches($Text, '[\p{L}\p{N}]+(?:[' + $joins + '][\p{L}\p{N}]+)*').Count
}

# The whole new text of a .txt file
function Get-TextResult($State, $Doc, $Plan, [hashtable]$S) {
    $text = $Doc.Text
    $eol = $Doc.Eol
    $sb = New-Object System.Text.StringBuilder
    if ($Plan.Begin -and $Plan.Begin.Change) {
        $new = Get-NewPlainText $State $S.Begin $eol
        if ($new) {
            if ($Plan.Begin.ParaMode -and -not $new.EndsWith($eol)) { $new += $eol }
            elseif ($Plan.Begin.NeedSpace -and $new -notmatch '\s$') { $new += ' ' }
            [void]$sb.Append($new)
        }
    }
    # The words to replace and the sections to delete, in the order they come
    $edits = New-Object System.Collections.ArrayList
    if ($S.Words.On) {
        foreach ($h in $Plan.Hits) { [void]$edits.Add([pscustomobject]@{ Start = $h[0]; End = $h[1]; New = (Get-NewText $text.Substring($h[0], $h[1] - $h[0]) $S.Words.Replace $S.Words.MatchCase $eol) }) }
    }
    if ($Plan.Between) {
        foreach ($sec in $Plan.Between.Sections) { [void]$edits.Add([pscustomobject]@{ Start = $sec.Start; End = $sec.End; New = '' }) }
    }
    $pos = $Plan.KeepStart
    foreach ($ed in ($edits | Sort-Object Start)) {
        [void]$sb.Append($text.Substring($pos, $ed.Start - $pos))
        [void]$sb.Append($ed.New)
        $pos = $ed.End
    }
    [void]$sb.Append($text.Substring($pos, $Plan.KeepEnd - $pos))
    if ($Plan.End -and $Plan.End.Change) {
        $new = Get-NewPlainText $State $S.End $eol
        if ($new) {
            $new = $new.TrimEnd("`r", "`n")
            $keptEndsLine = ($Plan.KeepEnd -gt 0 -and (Test-LineBreak $text ($Plan.KeepEnd - 1)))
            if ($Plan.End.ParaMode -and $sb.Length -gt 0 -and -not $keptEndsLine) { $new = $eol + $new }
            elseif ($Plan.End.NeedSpace -and $new -notmatch '^\s') { $new = ' ' + $new }
            [void]$sb.Append($new)
        }
        # A file that ended with a line break still does
        if ((Test-LineBreak $text ($text.Length - 1)) -and $sb.Length -gt 0 -and -not (Test-LineBreak $sb.ToString() ($sb.Length - 1))) { [void]$sb.Append($eol) }
    }
    return $sb.ToString()
}

function Get-NewPlainText($State, [hashtable]$C, [string]$Eol) {
    if ($C.Source -eq 'File' -and $C.File) { return (Get-SourceText $State $C.File).Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", $Eol) }
    if ($C.Replace) { return Expand-Codes $C.Replace '' $Eol }
    return ''
}

function Add-TextSegments($State, [string]$Text, $Plan, [hashtable]$S) {
    if ($S.Begin.On) {
        Add-Seg $Plan 'Title' 'THE BEGINNING'
        $b = $Plan.Begin
        if (Add-CutNote $Plan $b $S.Begin 'Begin') {
            if ($b.Removes) { Add-TextRangeSegs $Plan $Text 0 $b.Cut 'Old' }
            Add-NewSeg $State $Plan $S.Begin
            Add-Seg $Plan 'Text' ((Format-ShowText $Text.Substring($b.Cut, [Math]::Min($script:ShowKept, $Plan.KeepEnd - $b.Cut))) + ' ...')
        }
    }
    if ($S.Between.On -and (Add-BetweenNote $Plan $S.Between)) {
        $shown = 0
        foreach ($sec in $Plan.Between.Sections) {
            if ($shown -ge $script:ShowWords) { break }
            $shown++
            $s0 = [Math]::Max($Plan.KeepStart, $sec.Start - $script:ShowAround)
            $e1 = [Math]::Min($Plan.KeepEnd, $sec.End + $script:ShowAround)
            Add-Seg $Plan 'Text' ('... ' + ((Format-ShowText $Text.Substring($s0, $sec.Start - $s0)) -replace '\s+', ' '))
            Add-TextRangeSegs $Plan $Text $sec.Start $sec.End 'Old'
            Add-Seg $Plan 'Text' (((Format-ShowText $Text.Substring($sec.End, $e1 - $sec.End)) -replace '\s+', ' ') + " ...`n")
        }
        if ($Plan.Between.Count -gt $shown) { Add-Seg $Plan 'Gap' ('... and ' + (Format-Count ($Plan.Between.Count - $shown)) + " more`n") }
    }
    if ($S.Words.On) {
        Add-Seg $Plan 'Title' ('WORDS, PHRASES AND MARKS: ' + (Get-WordsLabel $S.Words 40))
        if ($Plan.WordsNote) { Add-Seg $Plan 'Note' "Note: $($Plan.WordsNote).`n" }
        if (-not $Plan.WordsCount) { if (-not $Plan.WordsNote) { Add-Seg $Plan 'Note' 'Not found in this document (or only in a part that is cut off).' } }
        else {
            Add-Seg $Plan 'Note' ('Found ' + (Get-Plural $Plan.WordsCount 'time' 'times') + ":`n")
            $shown = 0
            foreach ($h in $Plan.Hits) {
                if ($shown -ge $script:ShowWords) { break }
                $shown++
                $s0 = [Math]::Max($Plan.KeepStart, $h[0] - $script:ShowAround)
                $e1 = [Math]::Min($Plan.KeepEnd, $h[1] + $script:ShowAround)
                Add-HitSegs $Plan $S.Words (Format-ShowText $Text.Substring($s0, $h[0] - $s0)) $Text.Substring($h[0], $h[1] - $h[0]) (Format-ShowText $Text.Substring($h[1], $e1 - $h[1])) "`n" $false
            }
            if ($Plan.WordsCount -gt $shown) { Add-Seg $Plan 'Gap' ('... and ' + (Format-Count ($Plan.WordsCount - $shown)) + " more`n") }
        }
    }
    if ($S.End.On) {
        Add-Seg $Plan 'Title' 'THE ENDING'
        $e = $Plan.End
        if (Add-CutNote $Plan $e $S.End 'End') {
            $s0 = [Math]::Max($Plan.KeepStart, $e.Cut - $script:ShowKept)
            Add-Seg $Plan 'Text' ('... ' + (Format-ShowText $Text.Substring($s0, $e.Cut - $s0)))
            if ($e.Removes) { Add-TextRangeSegs $Plan $Text $e.Cut $Text.Length 'Old' }
            Add-NewSeg $State $Plan $S.End
        }
    }
}

function Add-TextRangeSegs($Plan, [string]$Text, [int]$Start, [int]$End, [string]$Kind) {
    $n = $End - $Start
    if ($n -le 2 * $script:ShowEdge + 100) { Add-Seg $Plan $Kind (Format-ShowText $Text.Substring($Start, $n)); return }
    Add-Seg $Plan $Kind (Format-ShowText $Text.Substring($Start, $script:ShowEdge))
    Add-Seg $Plan 'Gap' ("`n ... " + (Format-Count ($n - 2 * $script:ShowEdge)) + " more characters ... `n")
    Add-Seg $Plan $Kind (Format-ShowText $Text.Substring($End - $script:ShowEdge, $script:ShowEdge))
}

# ---------------------------------------------------------------- Word itself

function Get-Word($State) {
    if (-not $State.WordTried) {
        $State.WordTried = $true
        try { $State.Word = Start-Word } catch {}
    }
    if (-not $State.Word) { throw 'NOWORD:' }
    # If Word crashed on the last document, start a fresh copy
    if (-not (Test-WordAlive $State.Word)) {
        $State.Word = $null
        try { $State.Word = Start-Word } catch { throw 'NOWORD:' }
    }
    return $State.Word
}

function Start-Word {
    try { $w = New-Object -ComObject Word.Application -ErrorAction Stop }
    catch { throw 'Microsoft Word could not be started.' }
    $w.Visible = $false
    $w.DisplayAlerts = 0                                   # never show pop-up questions
    try { $w.AutomationSecurity = 3 } catch {}             # never run macros inside the documents
    try { $w.Options.UpdateLinksAtOpen = $false } catch {}
    try { $w.Options.ConfirmConversions = $false } catch {}
    try { $w.Options.SaveNormalPrompt = $false } catch {}
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

# Opens a document read-only: Word never saves over the file it opened
function Open-InWord($Word, [string]$Path) {
    $m = [Type]::Missing
    $pw = '#no-password#'   # a wrong password makes protected files fail at once instead of asking
    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    # Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible,
    # OpenAndRepair, DocumentDirection, NoEncodingDialog
    try { return $Word.Documents.Open([string]$Path, $false, $true, $false, $pw, $pw, $false, $pw, $pw, 0, $m, $false, $false, $m, $true) }
    catch {
        $why = Get-InnerMessage $_.Exception
        if ($why -match 'password') { throw 'PASSWORD:' }
        throw ('OPEN:' + $why)
    }
}

function Close-WordDocument($Doc) {
    if ($null -eq $Doc) { return }
    try { $Doc.Close(0) } catch { try { $noSave = 0; $Doc.Close([ref]$noSave) } catch {} }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Doc) } catch {}
}


# ---------------------------------------------------------------- messages

function Get-InnerMessage($Ex) {
    while ($Ex.InnerException) { $Ex = $Ex.InnerException }
    return (($Ex.Message -replace '\s+', ' ').Trim())
}

function Get-FriendlyError([string]$Message) {
    $kind = ''; $detail = $Message
    if ($Message -match '^(NOWORD|PASSWORD|OPEN|NOTDOC|NOTTEXT|TOOLONG|INUSE|READONLY|SAVE|INSERT|FOLDER):(.*)$') { $kind = $Matches[1]; $detail = $Matches[2].Trim() }
    switch ($kind) {
        'NOWORD'   { return 'Microsoft Word is needed to change this type of file, and it could not be started' }
        'PASSWORD' { return 'it is protected with a password' }
        'NOTDOC'   { return 'Word could not read it as a document. It may be damaged, or it may not really be a Word document' }
        'NOTTEXT'  { return 'this .txt file does not contain readable text' }
        'TOOLONG'  { return 'the text to find is too long for Word documents (Word looks for at most 255 characters)' }
        'INUSE'    { return 'the document is open in another program (such as Word). Close it and try again' }
        'READONLY' { return 'the document is read-only, so it can''t be changed' }
        'FOLDER'   { return 'a folder for the changed documents could not be made: ' + $detail }
        'INSERT'   { return 'the document chosen for the new text could not be put in: ' + $detail }
        'SAVE' {
            if ($detail -match 'already open|in use|being used|locked') { return 'it could not be saved, because the file it is saved to is open in another program. Close it and try again' }
            return 'Word could not save it: ' + $detail
        }
        'OPEN' {
            if ($detail -match 'RPC|remote procedure|disconnected|0x800706BA|0x80010108') { return 'Word stopped responding while working on it' }
            if ($detail -match 'File Block') { return "Word's security settings (File Block) do not allow this old type of file to be opened" }
            return 'Word could not open it. It may be damaged, or it may not really be a Word document'
        }
    }
    if ($Message -match 'being used by another process') { return 'the file is open in another program. Close it and try again' }
    if ($Message -match 'valid special character|search item was not found') { return 'Word does not accept a ^ code in the text to find' }
    return 'something went wrong: ' + $detail
}

function Get-FileError($Ex) {
    $inner = $Ex
    while ($inner.InnerException) { $inner = $inner.InnerException }
    $msg = Get-InnerMessage $Ex
    if ($msg -match 'being used by another process') { return 'the file is open in another program (such as Word). Close it and try again.' }
    if ($inner -is [UnauthorizedAccessException]) { return 'Windows did not allow it. The file may be read-only, or you may not have permission to change it.' }
    if ($inner -is [IO.FileNotFoundException] -or $inner -is [IO.DirectoryNotFoundException]) { return 'the file is not there any more (it may have been moved, renamed or deleted).' }
    return $msg
}