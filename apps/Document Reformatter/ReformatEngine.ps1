# ReformatEngine.ps1
# The part of Document Reformatter that does the real work: it finds the .rtf,
# .doc, .docx and .txt files in a folder and uses Microsoft Word to reformat each one and save
# it as .rtf in the same folder, moving the original into an "Originals" folder.
# If Word can't open a file, LibreOffice (when installed) reads it
# first. "Document Reformatter.ps1" (the window) loads this file.

# Word's numbers for the things we use (Word calls them wd... constants)
$script:WdFormatRTF       = 6
$script:WdStyleNormal     = -1
$script:WdOutlineBodyText = 10
$script:WdColorAutomatic  = -16777216
$script:WdUndefined       = 9999999
$script:WdFieldHyperlink  = 88
$script:WdWithInTable     = 12
$script:WdNoProtection    = -1
$script:WdOpenFormatEncodedText = 5
# Paper sizes in points (72 points = 1 inch), standing upright
$script:PaperSizes = @{ Letter = @(612, 792); A4 = @(595.3, 841.9) }
# Heading 1-9, Title and Subtitle
$script:WdHeadingStyleIds = @(-2, -3, -4, -5, -6, -7, -8, -9, -10, -63, -75)
# Word splits a document into "stories". Main text and text boxes get the new
# size; headers, footers and footnotes only get the new font.
$script:WdSizedStories    = @(1, 5)

$script:DocExtensions = @('.rtf', '.doc', '.docx', '.txt')
$script:OriginalsName = 'Originals'   # the folder the originals are moved into
# Short work folder in Temp: LibreOffice crashes on very long paths
$script:WorkRoot      = Join-Path ([IO.Path]::GetTempPath()) 'DocumentReformatter'


# ---------------------------------------------------------------- finding files

function Find-Documents {
    param([string]$Folder)
    $files = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction Stop |
        Where-Object { $script:DocExtensions -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object Name)

    $items = New-Object System.Collections.ArrayList
    foreach ($f in $files) {
        $skip = $null
        if ($f.Name.StartsWith('~$')) { $skip = 'Temporary file that Word makes while a document is open - not a real document.' }
        elseif ($f.Length -eq 0) { $skip = 'The file is empty (0 bytes).' }
        [void]$items.Add([pscustomobject]@{
            Path = $f.FullName; Name = $f.Name; BaseName = $f.BaseName
            Type = $f.Extension.TrimStart('.').ToLowerInvariant()
            OutputName = $f.BaseName + '.rtf'; SkipReason = $skip })
    }

    # Report.doc and Report.docx would both become Report.rtf, so files that
    # share a name get their old type added: "Report (doc).rtf", "Report (docx).rtf".
    # Skipped files count too, because they stay in the folder under their name.
    $counts = @{}
    foreach ($it in $items) { $counts[$it.BaseName] = 1 + [int]$counts[$it.BaseName] }
    foreach ($it in $items) {
        if (-not $it.SkipReason -and $counts[$it.BaseName] -gt 1) { $it.OutputName = '{0} ({1}).rtf' -f $it.BaseName, $it.Type }
    }
    return ,$items
}


# ---------------------------------------------------------------- the batch

# Each reformatted document is saved in the folder itself, and its original is
# moved into the "Originals" folder inside it.
function Invoke-Reformat {
    param(
        [object[]]$Items,          # from Find-Documents
        [hashtable]$Settings,
        [string]$Folder,           # the folder the documents are in
        [scriptblock]$OnFile,      # & $OnFile <number> <total> <item>, before each file
        [scriptblock]$OnTick,      # called now and then so the window stays responsive
        [scriptblock]$ShouldStop   # returns $true when the user pressed Stop
    )
    $results = New-Object System.Collections.ArrayList
    $todo = @($Items | Where-Object { -not $_.SkipReason })
    foreach ($it in @($Items | Where-Object { $_.SkipReason })) { [void]$results.Add((New-Result $it 'Skipped' $it.SkipReason)) }
    if ($todo.Count -eq 0) { return ,$results }

    $saveDir = Join-Path $script:WorkRoot 'saved'
    New-Item -ItemType Directory -Force -Path $saveDir | Out-Null
    $libre = Find-LibreOffice
    $word = Start-Word
    try {
        # Only once Word has started, so a PC without Word isn't left with an empty folder
        $originals = Join-Path $Folder $script:OriginalsName
        if (-not (Test-Path -LiteralPath $originals)) {
            try { New-Item -ItemType Directory -Path $originals -ErrorAction Stop | Out-Null }
            catch { throw ('The "' + $script:OriginalsName + '" folder could not be made here: ' + $originals + "`n`n(" + (Get-InnerMessage $_.Exception) + ')') }
        }
        for ($i = 0; $i -lt $todo.Count; $i++) {
            $it = $todo[$i]
            if ($ShouldStop -and (& $ShouldStop)) {
                for ($j = $i; $j -lt $todo.Count; $j++) { [void]$results.Add((New-Result $todo[$j] 'Skipped' 'Not done, because you pressed Stop.')) }
                break
            }
            if ($OnFile) { & $OnFile ($i + 1) $todo.Count $it }
            # Saved in the work folder first: the original may have the very name the new file gets
            $out = Join-Path $saveDir 'reformatted.rtf'
            if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
            try {
                $note = Convert-OneDocument $word $it.Path $out $Settings $libre $OnTick
                $placed = Move-IntoPlace $it $out $Folder $originals
                [void]$results.Add((New-Result $it 'Converted' $note $placed.Output $placed.Original))
            } catch {
                [void]$results.Add((New-Result $it 'Failed' (Get-FriendlyError (Get-InnerMessage $_.Exception))))
                # If Word itself crashed, start a fresh copy for the next file
                if (-not (Test-WordAlive $word)) { $word = Start-Word }
            }
        }
    } finally {
        Stop-Word $word
    }
    return ,$results
}

function New-Result($Item, [string]$Status, [string]$Detail, [string]$Output = '', [string]$Original = '') {
    [pscustomobject]@{ Name = $Item.Name; Path = $Item.Path; Status = $Status; Detail = $Detail; Output = $Output; Original = $Original }
}

# Moves the original into the Originals folder, then the reformatted file into
# the folder. Nothing is ever overwritten: a name already taken gets " (2)",
# " (3)"... added. If the new file can't be put in place, the original goes back.
# Returns the names used: @{ Output = ...; Original = ... }
function Move-IntoPlace($Item, [string]$Reformatted, [string]$Folder, [string]$Originals) {
    $kept = Get-FreeName $Originals $Item.Name
    try { Move-Item -LiteralPath $Item.Path -Destination $kept -ErrorAction Stop }
    catch {
        Remove-Item -LiteralPath $Reformatted -Force -ErrorAction SilentlyContinue
        throw ('MOVE:' + (Get-InnerMessage $_.Exception))
    }
    $target = Get-FreeName $Folder $Item.OutputName
    try { Move-Item -LiteralPath $Reformatted -Destination $target -ErrorAction Stop }
    catch {
        $why = Get-InnerMessage $_.Exception
        Remove-Item -LiteralPath $Reformatted -Force -ErrorAction SilentlyContinue
        try { Move-Item -LiteralPath $kept -Destination $Item.Path -ErrorAction Stop }
        catch { throw ('KEPT:' + $script:OriginalsName + '\' + (Split-Path -Leaf $kept)) }
        throw ('PLACE:' + $why)
    }
    return @{ Output = (Split-Path -Leaf $target); Original = (Split-Path -Leaf $kept) }
}

# Folder\Name, or Folder\Name (2).ext, (3)... if that is taken
function Get-FreeName([string]$Folder, [string]$Name) {
    $path = Join-Path $Folder $Name
    $base = [IO.Path]::GetFileNameWithoutExtension($Name)
    $ext = [IO.Path]::GetExtension($Name)
    for ($n = 2; Test-Path -LiteralPath $path; $n++) { $path = Join-Path $Folder ('{0} ({1}){2}' -f $base, $n, $ext) }
    return $path
}


# ---------------------------------------------------------------- Word itself

function Start-Word {
    try { $w = New-Object -ComObject Word.Application -ErrorAction Stop }
    catch { throw 'Microsoft Word could not be started. This app needs Word to be installed on this PC.' }
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
    $format = 0; $encoding = $m
    if ([IO.Path]::GetExtension($Path) -ieq '.txt') { $format = $script:WdOpenFormatEncodedText; $encoding = Get-TextEncoding $Path }
    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    # Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible,
    # OpenAndRepair, DocumentDirection, NoEncodingDialog
    return $Word.Documents.Open($Path, $false, $true, $false, $pw, $pw, $false, $pw, $pw, $format, $encoding, $false, $false, $m, $true)
}

# Plain text files don't say which alphabet they were saved in, so look at the bytes.
# Returns the code page number Word understands.
function Get-TextEncoding([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return 65001 }   # UTF-8
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return 1200 }   # Unicode (UTF-16)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return 1201 }
    # Unicode without a marker at the start: for English text, every other byte is zero
    $check = [Math]::Min($bytes.Length, 4000)
    $oddZeros = 0; $evenZeros = 0
    for ($i = 0; $i -lt $check; $i++) { if ($bytes[$i] -eq 0) { if ($i % 2) { $oddZeros++ } else { $evenZeros++ } } }
    if ($check -ge 4 -and $oddZeros * 4 -gt $check) { return 1200 }
    if ($check -ge 4 -and $evenZeros * 4 -gt $check) { return 1201 }
    try { [void](New-Object Text.UTF8Encoding($false, $true)).GetString($bytes); return 65001 } catch {}
    return [Text.Encoding]::Default.CodePage   # the PC's usual Windows alphabet, such as 1252
}

# Word runs lines together when they end with only a carriage return (old Mac
# files), so such a file is copied with Windows line endings first.
# Returns the path for Word to open.
function Get-WordReadyText([string]$Path) {
    $text = [Text.Encoding]::GetEncoding((Get-TextEncoding $Path)).GetString([IO.File]::ReadAllBytes($Path))
    if ($text -notmatch "`r(?!`n)") { return $Path }
    $dir = Join-Path $script:WorkRoot 'text'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $copy = Join-Path $dir 'source.txt'
    $text = [regex]::Replace($text.TrimStart([char]0xFEFF), "`r`n|`r|`n", "`r`n")
    [IO.File]::WriteAllText($copy, $text, (New-Object Text.UTF8Encoding($true)))
    return $copy
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
    if ($Message -match '^(OPEN|EDIT|SAVE|MOVE|PLACE|KEPT|LOCKED|NOTTEXT):(.*)$') { $kind = $Matches[1]; $detail = $Matches[2].Trim() }
    if ($detail -match 'RPC|remote procedure|disconnected|0x800706BA|0x80010108') { return 'Word stopped responding while working on this file.' }
    switch ($kind) {
        'OPEN' {
            if ($detail -match 'password') { return 'This file is protected with a password, so it was left alone.' }
            if ($detail -match 'File Block') { return "Word's security settings (File Block) do not allow this old type of file to be opened." }
            $text = 'Word could not open this file. It may be damaged, or it may not really be a Word document.'
            if ($detail -match 'LibreOffice') { $text += ' LibreOffice could not read it either.' }
            return $text
        }
        'LOCKED' { return 'This file is locked against changes with a password (Restrict Editing), so it was left alone.' }
        'NOTTEXT' { return 'This .txt file does not contain readable text, so it was left alone.' }
        'SAVE'   { return 'The reformatted file could not be saved, so the original was left as it was.' }
        'MOVE'   { return 'The original could not be moved into the "' + $script:OriginalsName + '" folder (it may be open in another program), so it was left as it was. Close it and try again.' }
        'PLACE'  { return 'The reformatted file could not be put in the folder, so the original was put back as it was.' }
        'KEPT'   { return 'The reformatted file could not be put in the folder. Your original is safe, unchanged, in ' + $detail + '.' }
        'EDIT'   { return 'Something went wrong while changing this file. (Word said: ' + $detail + ')' }
    }
    return 'Something went wrong: ' + $detail
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

# Has LibreOffice turn a file Word couldn't open into a .docx; returns its path
function Convert-WithLibreOffice([string]$Soffice, [string]$Path, [scriptblock]$OnTick) {
    $work = Join-Path $script:WorkRoot 'work'
    $inDir = Join-Path $work 'in'
    $outDir = Join-Path $work 'out'
    if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $inDir, $outDir | Out-Null
    $copy = Join-Path $inDir ('source' + [IO.Path]::GetExtension($Path))
    Copy-Item -LiteralPath $Path -Destination $copy

    # LibreOffice gets its own settings folder, so it works even if you have LibreOffice open
    $profileUrl = 'file:///' + ((Join-Path $script:WorkRoot 'lo-profile') -replace '\\', '/')
    $argLine = '--headless --norestore --nolockcheck "-env:UserInstallation={0}" --convert-to docx --outdir "{1}" "{2}"' -f $profileUrl, $outDir, $copy
    $proc = Start-Process -FilePath $Soffice -ArgumentList $argLine -WindowStyle Hidden -PassThru
    $deadline = (Get-Date).AddMinutes(3)
    while (-not $proc.WaitForExit(250)) {
        if ($OnTick) { & $OnTick }
        if ((Get-Date) -gt $deadline) { try { $proc.Kill() } catch {}; throw 'LibreOffice took too long.' }
    }
    $result = Join-Path $outDir 'source.docx'
    if (-not (Test-Path -LiteralPath $result)) { throw 'LibreOffice could not read it either.' }
    return $result
}


# ---------------------------------------------------------------- one document

function Convert-OneDocument {
    param($Word, [string]$Path, [string]$OutPath, [hashtable]$S, [string]$LibreOffice, [scriptblock]$OnTick)
    $notes = New-Object System.Collections.ArrayList
    $isText = ([IO.Path]::GetExtension($Path) -ieq '.txt')
    $stats = @{ EmptyLines = 0; Breaks = 0 }
    $doc = $null
    $openPath = $Path
    if ($isText) { $openPath = Get-WordReadyText $Path }
    try { $doc = Open-InWord $Word $openPath }
    catch {
        $why = Get-InnerMessage $_.Exception
        if ($why -match 'password' -or -not $LibreOffice) { throw ('OPEN:' + $why) }
        try { $copy = Convert-WithLibreOffice $LibreOffice $Path $OnTick }
        catch { throw ('OPEN:' + $why + ' | LibreOffice: ' + (Get-InnerMessage $_.Exception)) }
        try { $doc = Open-InWord $Word $copy } catch { throw ('OPEN:' + (Get-InnerMessage $_.Exception)) }
        if (Test-ScrambledText $doc) { Close-WordDocument $doc; throw ('OPEN:' + $why + ' | LibreOffice: only scrambled text') }
        [void]$notes.Add('read with LibreOffice, because Word could not open it')
    }
    if ($isText -and (Test-ScrambledText $doc)) { Close-WordDocument $doc; throw 'NOTTEXT:' }
    try {
        try { Unlock-Document $doc } catch { throw 'LOCKED:' }
        try {
            $links = 0
            if ($S.RemoveLinks) { $links = Remove-Hyperlinks $doc }
            if ($S.SplitParagraphs) { $stats.Breaks = Split-RunTogether $doc.Content $OnTick }
            Set-DocumentFormatting $doc $S $stats $OnTick
        } catch { throw ('EDIT:' + (Get-InnerMessage $_.Exception)) }
        try { $doc.SaveAs2($OutPath, $script:WdFormatRTF) } catch { throw ('SAVE:' + (Get-InnerMessage $_.Exception)) }
    } finally {
        Close-WordDocument $doc
    }
    if ($S.RemoveLinks -and -not $isText) {
        if ($links -eq 0) { [void]$notes.Add('no links found') }
        elseif ($links -eq 1) { [void]$notes.Add('1 link removed') }
        else { [void]$notes.Add("$links links removed") }
    }
    if ($stats.Breaks -eq 1) { [void]$notes.Add('1 paragraph break added where the text ran together') }
    elseif ($stats.Breaks -gt 1) { [void]$notes.Add("$($stats.Breaks) paragraph breaks added where the text ran together") }
    if ($stats.EmptyLines -eq 1) { [void]$notes.Add('1 empty line removed') }
    elseif ($stats.EmptyLines -gt 1) { [void]$notes.Add("$($stats.EmptyLines) empty lines removed") }
    return ($notes -join '; ')
}

# A damaged file that LibreOffice reads as plain text comes out as scrambled
# symbols. That counts as "could not be read" rather than saving junk.
function Test-ScrambledText($Doc) {
    $text = [string]$Doc.Content.Text
    if ($text.Contains([string][char]0xE0 + [char]0xA1 + [char]0xB1)) { return $true }   # the start of an old Word file, read as text
    if ($text.Length -lt 20) { return $false }
    $odd = [regex]::Matches($text, '[\x00\x03\x04\x06\x0E-\x12\x16-\x1F\uFFFD]').Count
    return ($odd * 20 -gt $text.Length)   # more than 5% junk characters
}

function Unlock-Document($Doc) {
    try { if ($Doc.Final) { $Doc.Final = $false } } catch {}
    # Works for files locked without a password; fails (as it should) for password-locked ones
    if ($Doc.ProtectionType -ne $script:WdNoProtection) { $Doc.Unprotect('#no-password#') }
    $Doc.TrackRevisions = $false
}

# Every part of the document: main text, headers, footers, footnotes, text boxes...
function Get-Stories($Doc) {
    # Touching the first header makes Word list every header and footer
    try { $null = $Doc.Sections.Item(1).Headers.Item(1).Range.StoryType } catch {}
    $list = New-Object System.Collections.ArrayList
    foreach ($story in $Doc.StoryRanges) {
        $r = $story
        while ($null -ne $r) { [void]$list.Add($r); $r = $r.NextStoryRange }
    }
    return ,$list
}

# Removes every link but keeps its words, and makes them look like normal text
function Remove-Hyperlinks($Doc) {
    $count = 0
    foreach ($story in (Get-Stories $Doc)) {
        $links = $story.Hyperlinks
        for ($i = $links.Count; $i -ge 1; $i--) {
            $h = $links.Item($i)
            $r = $null
            try { $r = $h.Range } catch {}
            $h.Delete()
            Clear-LinkLook $r
            $count++
        }
        # Some files hold links as plain HYPERLINK fields that Word doesn't list above
        $fields = $story.Fields
        for ($i = $fields.Count; $i -ge 1; $i--) {
            $f = $fields.Item($i)
            if ($f.Type -eq $script:WdFieldHyperlink) {
                $r = $f.Result
                $f.Unlink()
                Clear-LinkLook $r
                $count++
            }
        }
    }
    return $count
}

function Clear-LinkLook($Range) {
    if ($null -eq $Range) { return }
    try { $Range.Font.Underline = 0; $Range.Font.Color = $script:WdColorAutomatic } catch {}
}


# ---------------------------------------------------------------- run-together paragraphs

# Text copied from a web page sometimes loses its line breaks, so a whole story
# turns into one giant paragraph. Often each lost break left two spaces behind
# after the end of a sentence, so long paragraphs are split there. Whatever is
# still very long after that is split at sentence ends, every few sentences,
# at the start of a quote where there is one.
$script:LongParagraph = 1500   # characters; paragraphs longer than this are looked at
$script:HugeParagraph = 3000   # still longer than this: split at sentence ends
$script:PieceLength   = 700    # about how long those pieces are
$script:ShortestPiece = 150    # never split off less than this
# The end of a sentence, maybe followed by closing quotes or brackets, then the
# spaces after it (the "gap") and the first thing of the next sentence
$script:SentenceGap = '[.!?' + [char]0x2026 + '][''"' + [char]0x201D + [char]0x2019 + ')\]]*(?<gap>[ \xA0]+)(?=\S)'
$script:OpeningQuotes = '"' + [char]0x201C + [char]0x2018   # a new line of dialogue
$script:OpeningMarks  = $script:OpeningQuotes + '(['        # anything that can start a sentence

# Splits the long paragraphs of the main text; returns how many breaks were added
function Split-RunTogether($Story, [scriptblock]$OnTick) {
    $longest = 0
    foreach ($piece in ([string]$Story.Text).Split([char]13)) { if ($piece.Length -gt $longest) { $longest = $piece.Length } }
    if ($longest -le $script:LongParagraph) { return 0 }   # nothing to do: most documents stop here

    $long = New-Object System.Collections.ArrayList
    $n = 0
    foreach ($p in $Story.Paragraphs) {
        $n++
        if ($OnTick -and ($n % 25) -eq 0) { & $OnTick }
        $r = $p.Range
        if ($r.End - $r.Start -le $script:LongParagraph) { continue }
        if ($p.OutlineLevel -lt $script:WdOutlineBodyText) { continue }   # headings stay as they are
        if ($r.Information($script:WdWithInTable)) { continue }            # so do tables
        [void]$long.Add(@($r.Start, $r.End))
    }
    $added = 0
    # From the end backwards, so the places noted for earlier paragraphs stay right
    for ($i = $long.Count - 1; $i -ge 0; $i--) {
        $para = $Story.Duplicate
        $para.SetRange($long[$i][0], $long[$i][1])
        $text = [string]$para.Text
        $breaks = Get-ParagraphBreaks $text
        for ($j = $breaks.Count - 1; $j -ge 0; $j--) {
            if ($OnTick -and ($j % 25) -eq 0) { & $OnTick }
            if (Split-ParagraphAt $para $text $breaks[$j][0] $breaks[$j][1]) { $added++ }
        }
    }
    return $added
}

# Where one long paragraph should be split: a list of (gap start, gap end)
# places in its text. The spaces in each gap become the paragraph break.
function Get-ParagraphBreaks([string]$Text) {
    $gaps = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Text, $script:SentenceGap)) {
        $g = $m.Groups['gap']
        $next = $Text[$g.Index + $g.Length]
        [void]$gaps.Add(@{ From = $g.Index; To = $g.Index + $g.Length
            Double = ($g.Length -ge 2 -and -not [char]::IsLower($next))
            Quote = ($script:OpeningQuotes.IndexOf($next) -ge 0)
            Starts = ([char]::IsUpper($next) -or [char]::IsDigit($next) -or $script:OpeningMarks.IndexOf($next) -ge 0) })
    }
    $textEnd = $Text.TrimEnd([char]13, [char]7).Length
    $chosen = New-Object System.Collections.ArrayList

    # 1. Two spaces after a sentence - unless the writer puts two spaces after
    #    (nearly) every sentence, when they don't mean anything
    $doubles = @($gaps | Where-Object { $_.Double })
    if ($doubles.Count -gt 0 -and $doubles.Count -lt 0.8 * $gaps.Count) {
        foreach ($g in $doubles) { [void]$chosen.Add($g) }
    }

    # 2. Pieces that are still very long: every few sentences
    $all = New-Object System.Collections.ArrayList
    $pieceStart = 0
    $k = 0
    foreach ($stop in @($chosen) + @(@{ From = $textEnd; To = $textEnd; Last = $true })) {
        if ($stop.From - $pieceStart -gt $script:HugeParagraph) {
            $since = $pieceStart
            while ($k -lt $gaps.Count -and $gaps[$k].From -lt $stop.From) {
                $g = $gaps[$k]; $k++
                if ($g.From -le $since -or -not $g.Starts) { continue }
                if ($stop.From - $g.To -lt $script:ShortestPiece) { break }
                $length = $g.From - $since
                if ($length -ge $script:PieceLength -or ($g.Quote -and $length -ge $script:ShortestPiece)) {
                    [void]$all.Add(@($g.From, $g.To))
                    $since = $g.To
                }
            }
        }
        while ($k -lt $gaps.Count -and $gaps[$k].From -le $stop.From) { $k++ }
        if (-not $stop.Last) { [void]$all.Add(@($stop.From, $stop.To)) }
        $pieceStart = $stop.To
    }
    return ,@($all | Sort-Object { $_[0] })
}

# Turns the spaces between From and To (places in Text) into a paragraph break.
# Word counts some things, such as fields, differently from the text, so the
# spot is checked first; a spot that doesn't match is left alone.
function Split-ParagraphAt($Para, [string]$Text, [int]$From, [int]$To) {
    $a = [Math]::Max(0, $From - 4)
    $z = [Math]::Min($Text.Length, $To + 4)
    $check = $Para.Duplicate
    $check.SetRange($Para.Start + $a, $Para.Start + $z)
    if ([string]$check.Text -ne $Text.Substring($a, $z - $a)) { return $false }
    $gap = $Para.Duplicate
    $gap.SetRange($Para.Start + $From, $Para.Start + $To)
    $gap.InsertParagraph()
    return $true
}


# ---------------------------------------------------------------- fonts and styles

function Test-ParagraphChange([hashtable]$S) {
    foreach ($k in 'LineSpacing', 'SpaceBetween', 'Indent', 'Align') {
        $v = [string]$S[$k]
        if ($v -and $v -ne 'Leave') { return $true }
    }
    return $false
}

# Most settings need each paragraph looked at, to tell headings and tables
# apart. When none of them is chosen, the whole story gets the new size at once.
function Test-NeedsParagraphPass([hashtable]$S, $Story) {
    if ($S.Headings -ne 'Same' -or (Test-ParagraphChange $S)) { return $true }
    if ($S.TextColour -eq 'Black' -or $S.EvenSpacing -or $S.ClearEmphasis) { return $true }
    return ($S.RemoveEmptyLines -and [int]$Story.StoryType -eq 1)
}

function Set-DocumentFormatting($Doc, [hashtable]$S, [hashtable]$Stats, [scriptblock]$OnTick) {
    $font = [string]$S.FontName
    $size = [single]$S.FontSize
    $sameHeadings = ($S.Headings -eq 'Same')
    $paraChange = Test-ParagraphChange $S
    $colourAll = ($S.TextColour -eq 'BlackAll')
    $colourBody = ($S.TextColour -eq 'Black')

    $headingNames = @{}
    $titleNames = @{}   # Title and Subtitle have no outline level, so paragraphs are checked by style name
    foreach ($id in $script:WdHeadingStyleIds) {
        try {
            $st = $Doc.Styles.Item($id)
            $headingNames[[string]$st.NameLocal] = $true
            if (($id -eq -63 -or $id -eq -75) -and $st.InUse) { $titleNames[[string]$st.NameLocal] = $true }
        } catch {}
    }
    $normal = $Doc.Styles.Item($script:WdStyleNormal)
    $normalName = [string]$normal.NameLocal

    Set-PageLook $Doc $S

    # 1. Styles. Note how the other styles look, change Normal, then put the
    #    others back, so headings, footers etc. built on Normal don't shift.
    $snapshot = Get-StyleSnapshot $Doc $paraChange
    $normal.Font.Name = $font
    $normal.Font.Size = $size
    if ($paraChange) { Set-ParagraphLook $normal.ParagraphFormat $S $true }
    if ($colourAll -or $colourBody) { $normal.Font.Color = $script:WdColorAutomatic }
    if ($S.EvenSpacing) { Set-EvenSpacing $normal }
    if ($S.WidowControl) { $normal.ParagraphFormat.WidowControl = $true }
    foreach ($e in $snapshot) {
        if ($e.Name -eq $normalName) { continue }
        $isHeading = $headingNames.ContainsKey($e.Name)
        Restore-StyleLook $e $font $size ($sameHeadings -and $isHeading)
        if ($colourAll -and $isHeading) { try { $e.Style.Font.Color = $script:WdColorAutomatic } catch {} }
    }

    # 2. Text formatted by hand rather than by a style
    foreach ($story in (Get-Stories $Doc)) {
        try { $story.Font.Name = $font } catch {}
        if ($S.RemoveHighlight) { Clear-Highlight $story $OnTick }
        if ($colourAll) { try { $story.Font.Color = $script:WdColorAutomatic } catch {} }
        if ($script:WdSizedStories -notcontains [int]$story.StoryType) {
            # Headers, footers and footnotes have no headings, so all their text is ordinary text
            if ($colourBody) { try { $story.Font.Color = $script:WdColorAutomatic } catch {} }
            if ($S.EvenSpacing) { Set-EvenSpacing $story }
            continue
        }
        if ($S.WidowControl) { try { $story.ParagraphFormat.WidowControl = $true } catch {} }
        if (Test-NeedsParagraphPass $S $story) { Format-BodyParagraphs $story $S $titleNames $Stats $OnTick }
        else { $story.Font.Size = $size }
    }
}

# Paper size and margins, for every section of the document
function Set-PageLook($Doc, [hashtable]$S) {
    $paper = $script:PaperSizes[[string]$S.PaperSize]
    $margin = [string]$S.Margins
    $setMargin = ($margin -and $margin -ne 'Leave')
    if (-not $paper -and -not $setMargin) { return }
    foreach ($sec in $Doc.Sections) {
        $ps = $sec.PageSetup
        if ($paper) {
            $w = $paper[0]; $h = $paper[1]
            if ($ps.Orientation -eq 1) { $w = $paper[1]; $h = $paper[0] }   # pages turned sideways stay sideways
            $ps.PageWidth = $w
            $ps.PageHeight = $h
        }
        if ($setMargin) {
            $m = [single]$margin
            $ps.TopMargin = $m; $ps.BottomMargin = $m; $ps.LeftMargin = $m; $ps.RightMargin = $m
        }
    }
}

# Highlighter colours, and shading behind words and paragraphs. Shading in
# tables (on the cells or on the paragraphs inside them) is part of the
# table's design, so it stays. White or very light words that sat on the
# removed colour get the normal (automatic) colour first, so they don't
# vanish on the white page.
# Word only says whether a stretch of text has highlighting for a few thousand
# characters at a time (longer stretches come back as "mixed"), so the story
# is checked in blocks and only blocks with something in them are gone
# through paragraph by paragraph.
function Clear-Highlight($Range, [scriptblock]$OnTick) {
    $blockSize = 2000
    $doneTo = -1
    for ($at = $Range.Start; $at -lt $Range.End; $at += $blockSize) {
        if ($OnTick) { & $OnTick }
        $block = $Range.Duplicate
        $block.SetRange($at, [Math]::Min($at + $blockSize, $Range.End))
        if (-not (Test-HasShading $block)) { continue }
        foreach ($p in $block.Paragraphs) {
            $r = $p.Range
            if ($r.Start -lt $doneTo) { continue }   # already done as part of the block before
            $doneTo = $r.End
            if ($r.Information($script:WdWithInTable)) { if (-not (Test-HasMarking $r)) { continue } }
            elseif (Test-Shaded $p.Shading) {
                Reset-LightText $r   # the whole paragraph sat on colour
                Clear-Shading $p.Shading
                continue
            }
            elseif (-not (Test-HasMarking $r)) { continue }
            foreach ($w in $r.Words) {
                if ([int]$w.HighlightColorIndex -ne 0 -or (Test-Shaded $w.Font.Shading)) { Reset-LightText $w }
            }
        }
    }
    try { $Range.HighlightColorIndex = 0 } catch {}
    try { Clear-Shading $Range.Font.Shading } catch {}
}

function Test-HasShading($Range) {
    try { return ((Test-HasMarking $Range) -or (Test-Shaded $Range.ParagraphFormat.Shading)) } catch { return $true }
}

# Highlighting or shading on the words themselves
function Test-HasMarking($Range) {
    try { return ([int]$Range.HighlightColorIndex -ne 0 -or (Test-Shaded $Range.Font.Shading)) } catch { return $true }
}

# Shading that is really there: a colour or a pattern, not "none" (automatic,
# or -1 in tables) and not plain white
function Test-Shaded($Shading) {
    $texture = [int]$Shading.Texture
    if ($texture -ne 0 -and $texture -ne -1) { return $true }
    return (@($script:WdColorAutomatic, -1, 16777215) -notcontains [int]$Shading.BackgroundPatternColor)
}

function Clear-Shading($Shading) {
    try { $Shading.Texture = 0; $Shading.BackgroundPatternColor = $script:WdColorAutomatic; $Shading.ForegroundPatternColor = $script:WdColorAutomatic } catch {}
}
function Reset-LightText($Range) {
    $c = [int]$Range.Font.Color
    if ($c -ne $script:WdUndefined) {
        if (Test-LightColour $c) { $Range.Font.Color = $script:WdColorAutomatic }
        return
    }
    foreach ($ch in $Range.Characters) { if (Test-LightColour ([int]$ch.Font.Color)) { $ch.Font.Color = $script:WdColorAutomatic } }
}

# White or nearly white. Word stores a theme colour as 0xD in the top 4 bits,
# the theme slot in the next 4 (1, 3, 12 and 14 are the light ones) and how
# much it is lightened in the lowest byte (smaller = lighter).
function Test-LightColour([int]$Colour) {
    if ($Colour -eq $script:WdColorAutomatic -or $Colour -eq $script:WdUndefined) { return $false }
    $v = [int64]$Colour -band 4294967295
    if (($v -shr 28) -eq 13) { return ((@(1, 3, 12, 14) -contains [int](($v -shr 24) -band 15)) -or ($v -band 255) -lt 64) }
    if ($v -gt 16777215) { return $false }
    $r = $v -band 255; $g = ($v -shr 8) -band 255; $b = ($v -shr 16) -band 255
    return (0.299 * $r + 0.587 * $g + 0.114 * $b) -gt 200
}

# Letters squeezed together, spread apart or stretched go back to normal.
# Works on a range of text or on a style.
function Set-EvenSpacing($Target) {
    try { $Target.Font.Spacing = 0; $Target.Font.Scaling = 100 } catch {}
}

function Get-StyleSnapshot($Doc, [bool]$WithParagraph) {
    $snap = New-Object System.Collections.ArrayList
    foreach ($st in $Doc.Styles) {
        try {
            if (-not $st.InUse) { continue }
            $type = [int]$st.Type
            if (@(1, 2, 5, 6) -notcontains $type) { continue }   # paragraph, character, linked styles
            $e = @{ Style = $st; Name = [string]$st.NameLocal; IsCharacter = ($type -eq 2); Size = $null; Para = $null }
            if (-not $e.IsCharacter) {
                $e.Size = $st.Font.Size
                if ($WithParagraph) {
                    $pf = $st.ParagraphFormat
                    $e.Para = @{ Rule = $pf.LineSpacingRule; Spacing = $pf.LineSpacing; Before = $pf.SpaceBefore
                                 After = $pf.SpaceAfter; First = $pf.FirstLineIndent; Align = $pf.Alignment }
                }
            }
            [void]$snap.Add($e)
        } catch {}
    }
    return ,$snap
}

function Restore-StyleLook($E, [string]$Font, [single]$Size, [bool]$UseNewSize) {
    $st = $E.Style
    try { $st.Font.Name = $Font } catch {}
    if ($E.IsCharacter) { return }
    try {
        if ($UseNewSize) { $st.Font.Size = $Size }
        elseif ($null -ne $E.Size -and $E.Size -ne $script:WdUndefined) { $st.Font.Size = $E.Size }
    } catch {}
    if ($E.Para) {
        try {
            $pf = $st.ParagraphFormat
            $pf.LineSpacingRule = $E.Para.Rule
            if (@(3, 4, 5) -contains [int]$E.Para.Rule) { $pf.LineSpacing = $E.Para.Spacing }   # at least / exactly / multiple
            $pf.SpaceBefore = $E.Para.Before
            $pf.SpaceAfter = $E.Para.After
            $pf.FirstLineIndent = $E.Para.First
            $pf.Alignment = $E.Para.Align
        } catch {}
    }
}

# Works on a style's ParagraphFormat or on a single paragraph
function Set-ParagraphLook($Target, [hashtable]$S, [bool]$AllowIndent) {
    switch ([string]$S.LineSpacing) {
        '1'    { $Target.LineSpacingRule = 0 }
        '1.15' { $Target.LineSpacingRule = 5; $Target.LineSpacing = 13.8 }
        '1.5'  { $Target.LineSpacingRule = 1 }
        '2'    { $Target.LineSpacingRule = 2 }
    }
    if ([string]$S.SpaceBetween -and [string]$S.SpaceBetween -ne 'Leave') {
        try { $Target.SpaceBeforeAuto = $false; $Target.SpaceAfterAuto = $false } catch {}
        $Target.SpaceBefore = 0
        $Target.SpaceAfter = [single]$S.SpaceBetween
    }
    if ($AllowIndent -and [string]$S.Indent -and [string]$S.Indent -ne 'Leave') {
        if ($S.Indent -eq 'Half') { $Target.FirstLineIndent = 36 } else { $Target.FirstLineIndent = 0 }
    }
    switch ([string]$S.Align) {
        'Left'    { $Target.Alignment = 0 }
        'Justify' { $Target.Alignment = 3 }
    }
}

# Regular paragraphs get the new size, paragraph settings, colour, letter
# spacing and emphasis. Headings keep their size (unless "same size" was
# picked) and always keep their spacing and look. Table cells keep spacing and
# bold set on them directly (they still follow the Normal style), and bulleted
# lists keep their indents. Blank lines next to a table stay, so two tables
# never join into one.
# Neighbouring paragraphs are changed together, which is much faster.
function Format-BodyParagraphs($Story, [hashtable]$S, [hashtable]$TitleNames, [hashtable]$Stats, [scriptblock]$OnTick) {
    $sameHeadings = ($S.Headings -eq 'Same')
    $paraChange = Test-ParagraphChange $S
    $indentChange = ([string]$S.Indent -and [string]$S.Indent -ne 'Leave')
    $colourBody = ($S.TextColour -eq 'Black')
    $evenSpacing = [bool]$S.EvenSpacing
    $clearEmphasis = [bool]$S.ClearEmphasis
    $removeEmpty = ($S.RemoveEmptyLines -and [int]$Story.StoryType -eq 1)   # main text only
    $needTable = ($paraChange -or $clearEmphasis -or $removeEmpty)
    $needStyle = ($TitleNames.Count -gt 0 -or $clearEmphasis)
    $emphasisOf = @{}   # style name -> how bold, italic and underlined the style itself is
    $sizeRun = @{ Start = -1; End = -1 }
    $lookRun = @{ Start = -1; End = -1; Indent = $true }
    $charRun = @{ Start = -1; End = -1; Key = '' }
    $blanks = New-Object System.Collections.ArrayList
    $maybeBlank = $null
    # Where floating pictures and shapes are attached, so their lines are kept
    $anchors = New-Object System.Collections.Generic.List[int]
    if ($removeEmpty) {
        try { foreach ($shp in $Story.Document.Shapes) { $anchors.Add([int]$shp.Anchor.Start) } }
        catch { $removeEmpty = $false }   # can't tell where a shape sits, so leave every blank line alone
    }
    $prevInTable = $false
    $n = 0
    foreach ($p in $Story.Paragraphs) {
        $n++
        if ($OnTick -and ($n % 25) -eq 0) { & $OnTick }
        $r = $p.Range
        $start = $r.Start; $end = $r.End
        $heading = ($p.OutlineLevel -lt $script:WdOutlineBodyText)
        $styleName = ''
        if ($needStyle) {
            try {
                $pst = $p.Style
                $styleName = [string]$pst.NameLocal
                if ($clearEmphasis -and -not $emphasisOf.ContainsKey($styleName)) { $emphasisOf[$styleName] = Get-StyleEmphasis $pst }
            } catch {}
            if (-not $heading) { $heading = $TitleNames.ContainsKey($styleName) }
        }
        $inTable = $false
        if ($needTable -and (-not $heading -or $removeEmpty)) { $inTable = [bool]$r.Information($script:WdWithInTable) }

        # size
        if ($heading -and -not $sameHeadings) { Close-Run $Story $sizeRun $S 'Size' }
        else { Add-ToRun $sizeRun $start $end }

        # line spacing, space between, indent and alignment
        if ($paraChange) {
            if ($heading -or $inTable) { Close-Run $Story $lookRun $S 'Look' }
            else {
                $indentOk = $true
                if ($indentChange) { $indentOk = ($r.ListFormat.ListType -eq 0) }
                if ($lookRun.Start -ge 0 -and $lookRun.Indent -ne $indentOk) { Close-Run $Story $lookRun $S 'Look' }
                if ($lookRun.Start -lt 0) { $lookRun.Indent = $indentOk }
                Add-ToRun $lookRun $start $end
            }
        }

        # colour, letter spacing, and bold/italic/underline
        $key = ''
        if (-not $heading) {
            if ($colourBody) { $key += 'C' }
            if ($evenSpacing) { $key += 'S' }
            if ($clearEmphasis -and -not $inTable) {
                $look = $emphasisOf[$styleName]
                if (-not $look) { $look = '0,0,0' }
                $key += 'E' + $look
            }
        }
        if ($charRun.Start -ge 0 -and $charRun.Key -ne $key) { Close-Run $Story $charRun $S 'Char' }
        if ($key) {
            if ($charRun.Start -lt 0) { $charRun.Key = $key }
            Add-ToRun $charRun $start $end
        }

        # blank lines; one is only removed once we know the next paragraph isn't in a table
        if ($removeEmpty) {
            if ($maybeBlank -and -not $inTable) { [void]$blanks.Add($maybeBlank) }
            $maybeBlank = $null
            if (-not $inTable -and -not $prevInTable -and (Test-EmptyParagraph $r $start $end $anchors)) { $maybeBlank = @($start, $end) }
            $prevInTable = $inTable
        }
    }
    Close-Run $Story $sizeRun $S 'Size'
    Close-Run $Story $lookRun $S 'Look'
    Close-Run $Story $charRun $S 'Char'

    # From the end backwards, so the places noted for earlier blank lines stay right
    for ($i = $blanks.Count - 1; $i -ge 0; $i--) {
        if ($OnTick -and ($i % 25) -eq 0) { & $OnTick }
        $rr = $Story.Duplicate
        $rr.SetRange($blanks[$i][0], $blanks[$i][1])
        [void]$rr.Delete()
    }
    $Stats.EmptyLines += $blanks.Count
}

# "bold,italic,underline" as the paragraph's own style has them, for example "0,0,0"
function Get-StyleEmphasis($Style) {
    $f = $Style.Font
    $values = foreach ($v in $f.Bold, $f.Italic, $f.Underline) { if ([int]$v -eq $script:WdUndefined) { 0 } else { [int]$v } }
    return ($values -join ',')
}

# Only spaces or nothing, and no picture, shape, field or form box in it either.
# Floating shapes are found by where they are attached ($Anchors), because
# Word doesn't list them by paragraph for a document opened without a window.
function Test-EmptyParagraph($Range, [int]$Start, [int]$End, $Anchors) {
    if ($End - $Start -gt 200) { return $false }
    if ([string]$Range.Text -notmatch '^[ \t\r\n\u000B\u00A0]*$') { return $false }
    foreach ($a in $Anchors) { if ($a -ge $Start -and $a -lt $End) { return $false } }
    try {
        if ($Range.InlineShapes.Count -or $Range.Fields.Count -or $Range.ContentControls.Count) { return $false }
    } catch { return $false }
    return $true
}

function Add-ToRun([hashtable]$Run, [int]$Start, [int]$End) {
    if ($Run.Start -lt 0) { $Run.Start = $Start }
    $Run.End = $End
}

# Applies the new size, the paragraph settings, or the text look to one stretch of paragraphs
function Close-Run($Story, [hashtable]$Run, [hashtable]$S, [string]$What) {
    if ($Run.Start -lt 0) { return }
    $rr = $Story.Duplicate
    $rr.SetRange($Run.Start, $Run.End)
    if ($What -eq 'Size') { $rr.Font.Size = [single]$S.FontSize }
    elseif ($What -eq 'Look') { Set-ParagraphLook $rr.ParagraphFormat $S $Run.Indent }
    else {
        if ($Run.Key.Contains('C')) { $rr.Font.Color = $script:WdColorAutomatic }
        if ($Run.Key.Contains('S')) { Set-EvenSpacing $rr }
        if ($Run.Key -match 'E(-?\d+),(-?\d+),(-?\d+)') {
            $rr.Font.Bold = [int]$Matches[1]; $rr.Font.Italic = [int]$Matches[2]; $rr.Font.Underline = [int]$Matches[3]
        }
    }
    $Run.Start = -1
}
