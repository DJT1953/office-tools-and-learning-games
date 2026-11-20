# Tests Document Reformatter's engine on made-up documents, through Word:
# - giant paragraphs are split where the line breaks were lost
# - normal paragraphs, headings and tables are left alone
# - the reformatted files stay in the folder and the originals move into "Originals"
# Run:  powershell -NoProfile -STA -ExecutionPolicy Bypass -File test_reformat.ps1
# Needs Microsoft Word. Exit code 0 = all checks passed.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReformatEngine.ps1')
$failures = 0
function Check([bool]$ok, [string]$what) {
    if ($ok) { Write-Host "  ok    $what" } else { Write-Host "  FAIL  $what"; $script:failures++ }
}

$q = [char]0x201C; $qq = [char]0x201D   # curly opening and closing quotes
# Made-up sentences; paragraph p, sentence s
function Sentence([int]$p, [int]$s) {
    switch ($s % 4) {
        0 { "Paragraph $p begins here with sentence $s of the sample." }
        1 { "${q}This is spoken line $s in paragraph $p,$qq said the speaker." }
        2 { "The weather in part $p stayed calm and grey for most of sentence $s." }
        default { "Number $s ends paragraph $p's thought, and then it moves on." }
    }
}
# Paragraphs of 5 sentences, joined with $Join (as text copied from a web page)
function New-Story([int]$Paragraphs, [string]$Join, [string]$InsideJoin = ' ') {
    $paras = foreach ($p in 1..$Paragraphs) { (0..4 | ForEach-Object { Sentence $p $_ }) -join $InsideJoin }
    return ($paras -join $Join)
}

function New-WordFile($Word, [string]$Path, [scriptblock]$Fill, [int]$Format) {
    $d = $Word.Documents.Add()
    & $Fill $d
    $d.SaveAs2($Path, $Format)
    $d.Close(0)
}

function Get-Paragraphs($Word, [string]$Path) {
    $d = $Word.Documents.Open($Path, $false, $true, $false)
    $out = foreach ($p in $d.Paragraphs) { [pscustomobject]@{ Text = ([string]$p.Range.Text).TrimEnd([char]13, [char]7); Heading = ($p.OutlineLevel -lt 10); Table = [bool]$p.Range.Information(12) } }
    $d.Close(0)
    return ,@($out)
}

# A paragraph cut into lines of up to $Width characters, as e-mails and old text files are
function Wrap([string]$Text, [int]$Width = 70) {
    $lines = @(); $line = ''
    foreach ($w in $Text -split ' ') {
        if ($line -and ($line.Length + 1 + $w.Length) -gt $Width) { $lines += $line; $line = $w }
        elseif ($line) { $line += ' ' + $w } else { $line = $w }
    }
    return @($lines + $line)
}
# Joins the lines of $Text as the engine would, without Word (the last paragraph end stays)
function Join-Text([string]$Text) {
    $width = Get-WrappedLineLength $Text
    if ($width -eq 0) { return $Text }
    $edits = Get-LineJoins $Text $width ($Text -notmatch '[\r\v][ \t\xA0]*[\r\v]')
    for ($j = $edits.Count - 1; $j -ge 0; $j--) {
        $e = $edits[$j]; $Text = $Text.Substring(0, $e.From) + $e.With + $Text.Substring($e.To)
    }
    return $Text
}
# Five made-up paragraphs, each 4-6 wrapped lines
$plainParas = foreach ($p in 1..5) { (0..4 | ForEach-Object { Sentence $p $_ }) -join ' ' }
$lineEnds = @("`r", " `r", "`v", " `v")   # paragraph break, space then paragraph break, line break, space then line break

Write-Host 'Joining lines (no Word needed)'
$i = 0
$broken = (@(foreach ($para in $plainParas) { (@(Wrap $para) | ForEach-Object { $_ }) -join $lineEnds[$i++ % 4] }) -join "`r`r") + "`r"
Check ((Get-WrappedLineLength $broken) -gt 0) 'text broken at every line end is recognised'
Check ((Join-Text $broken) -ceq (($plainParas -join "`r") + "`r")) 'each kind of single break becomes a space; two in a row end the paragraph'
$gaps = @("`r`r", "`r `r", "`v`v", "`v `r", "`r`r`r", " `v`r`r ")
$i = 0
$mixed = $plainParas[0]; foreach ($para in $plainParas[1..4]) { $mixed += $gaps[$i++ % $gaps.Count] + $para }
$wrappedMixed = ((($mixed -split '(\r[\r\v ]*|\v[\r\v ]*)') | ForEach-Object { if ($_ -match '^[\r\v ]+$') { $_ } else { (Wrap $_) -join "`r" } }) -join '') + "`r"
Check ((Join-Text $wrappedMixed) -ceq (($plainParas -join "`r") + "`r")) 'two, three or more breaks in a row (with or without spaces) make one paragraph end'
$indented = ((Wrap $plainParas[0]) -join "`r") + "`r    " + ((Wrap $plainParas[1]) -join "`r") + "`r`t" + ((Wrap $plainParas[2]) -join "`r") + "`r"
$want = $plainParas[0] + "`r    " + $plainParas[1] + "`r`t" + $plainParas[2] + "`r"
Check ((Join-Text $indented) -ceq $want) 'a paragraph break followed by an indent stays as it is, indent and all'
$lineIndent = ((Wrap $plainParas[0]) -join "`r") + "`v`t" + ((Wrap $plainParas[1]) -join "`r") + "`r`r  " + ((Wrap $plainParas[2]) -join "`r") + "`r"
Check ((Join-Text $lineIndent) -ceq ($plainParas[0] + "`r`t" + $plainParas[1] + "`r  " + $plainParas[2] + "`r")) 'a line break or a double break before an indent ends the paragraph and keeps the indent'
$stray = (@(foreach ($para in $plainParas[0..2]) { (Wrap $para) -join "`r " }) -join "`r`r") + "`r"
Check ((Join-Text $stray) -ceq (($plainParas[0..2] -join "`r") + "`r")) 'a single space at the start of a line is not taken for an indent'
$crlf = (@(foreach ($para in $plainParas) { (Wrap $para) -join "`r`n" }) -join "`r`n`r`n") + "`r`n"
Check ((Join-Text $crlf) -ceq (($plainParas -join "`r") + "`r`n")) 'plain-text line ends (a paragraph end with a line feed) count as one break'
$few = ((Wrap $plainParas[0]) -join "`r") + "`r"
Check ((Get-WrappedLineLength $few) -eq 0) 'a single short paragraph (under 5 cut lines) is left alone'
$noBlank = (@(foreach ($para in $plainParas) { (Wrap $para) -join "`r" }) -join "`r") + "`r"
Check ((Join-Text $noBlank) -ceq (($plainParas -join "`r") + "`r")) 'with no blank lines at all, a short line ending a sentence ends the paragraph'
# Capital letters as a clue (no blank lines): lines of exact lengths. Wrapped lines are
# 70 long; a "fairly short" line is 60 (between 75% and 90% of 70), a "short" one 45.
$fill = 'went along the road past the old farm and over the hill to the town where it rained '
function Line([string]$Start, [int]$Length, [string]$End = 'x') {
    $body = ($Start + ' ' + ($fill * 3)).Substring(0, $Length - $End.Length)
    if ($body.EndsWith(' ')) { $body = $body.Substring(0, $body.Length - 1) + 'y' }
    return $body + $End
}
$cap = @(
    (Line 'it' 70), (Line 'and' 70), (Line 'so' 70), (Line 'then' 60 '.'),   # fairly short, then "The" (an ordinary word)
    (Line 'The' 70), (Line 'and' 70), (Line 'so' 70), (Line 'then' 60 '.'),  # fairly short, then "Mary" (a name)
    (Line 'Mary' 70), (Line 'and' 70), (Line 'so' 45 'etc.'),                # short, ends "etc.", then a small letter
    (Line 'and' 70), (Line 'so' 45 '.'),                                       # short, then "Mary"
    (Line 'Mary' 70), (Line 'and' 70), (Line 'so' 70), (Line 'then' 45 '.')
) -join "`r"
$joined = Join-Text ($cap + "`r")
$c = $cap -split "`r"
Check ($joined.Contains($c[3] + "`r" + $c[4])) 'a fairly short line ending a sentence, then a capital on an ordinary word: new paragraph'
Check ($joined.Contains($c[7] + ' ' + $c[8])) 'the same, but the next line starts with a name: the lines join'
Check ($joined.Contains($c[10] + ' ' + $c[11])) 'a short line ending "etc.", then a small letter: the lines join'
Check ($joined.Contains($c[12] + "`r" + $c[13])) 'a short line ending a sentence, then a name: still a new paragraph'
Check ((Get-SentenceStart ([string][char]0x201C + 'Hello') (Get-OrdinaryWords 'x')) -eq 'yes' -and (Get-SentenceStart 'I went' (Get-OrdinaryWords 'i went')) -eq 'maybe') 'an opening quote starts a sentence; "I" says nothing either way'
$normal = (($plainParas + $plainParas) -join "`r") + "`r"
Check ((Get-WrappedLineLength $normal) -eq 0) 'a normal document (one break after each long paragraph) is not taken for broken lines'
$letter = "John Smith`r12 Long Road Over The Hill`rSpringfield, Some State 12345`r`rDear Sir`r" + (($plainParas[0..2]) -join "`r") + "`rYours truly`rJane Doe`r"
Check ((Get-WrappedLineLength $letter) -eq 0) 'a letter with a short address block is not taken for broken lines'

Write-Host 'Where breaks go (no Word needed)'
$s = 'This is a plain sentence that goes on for a while. '
$b = Get-ParagraphBreaks (($s * 20) + 'End here.  Next part starts. ' + ($s * 20))
Check ($b.Count -eq 1 -and $b[0][1] - $b[0][0] -eq 2) "one double space: exactly one break, over the two spaces (got $($b.Count))"
$b = Get-ParagraphBreaks ($s * 10)
Check ($b.Count -eq 0) 'an ordinary long paragraph with no marks is left alone'
$b = Get-ParagraphBreaks (($s * 20) + 'end here.  and on it goes. ' + ($s * 20))
Check ($b.Count -eq 0) 'two spaces before a small letter are not a break'
$b = Get-ParagraphBreaks ((($s * 10) + '(A bracketed aside.) ' + ($s * 2)) * 6)
Check (@($b | Where-Object { $_[0] -gt 0 -and (((($s * 10) + '(A bracketed aside.) ' + ($s * 2)) * 6).Substring($_[1], 1)) -eq '(' }).Count -eq 0) 'a bracket is not taken for the start of dialogue'
$long = ($s * 200)
$b = Get-ParagraphBreaks $long
$pieces = @(); $at = 0; foreach ($x in $b) { $pieces += $x[0] - $at; $at = $x[1] }; $pieces += $long.TrimEnd().Length - $at
Check (($pieces | Measure-Object -Maximum).Maximum -le 1400 -and ($pieces | Measure-Object -Minimum).Minimum -ge 150) "no marks at all: pieces between 150 and 1400 characters ($(($pieces | Measure-Object -Minimum).Minimum) to $(($pieces | Measure-Object -Maximum).Maximum))"

Write-Host 'Names and messages (no Word needed)'
$tmp = Join-Path ([IO.Path]::GetTempPath()) "reformatter-names-$PID"
New-Item -ItemType Directory -Force $tmp | Out-Null
Check ((Split-Path -Leaf (Get-FreeName $tmp 'A.rtf')) -eq 'A.rtf') 'a free name is used as it is'
Set-Content (Join-Path $tmp 'A.rtf') 'x'; Set-Content (Join-Path $tmp 'A (2).rtf') 'x'
Check ((Split-Path -Leaf (Get-FreeName $tmp 'A.rtf')) -eq 'A (3).rtf') 'a taken name gets the next free number'
Check ((Get-FriendlyError 'KEPT:Originals\A (2).rtf') -match 'safe, unchanged, in Originals\\A \(2\)\.rtf') 'if the original cannot be put back, the message says where it is'
Check ((Get-FriendlyError 'MOVE:in use') -match 'could not be moved into the "Originals" folder') 'a file that cannot be moved gets a plain message'
# An original that another program holds open stays where it is, and nothing is left behind
$orig = Join-Path $tmp 'Busy.txt'; Set-Content $orig 'busy'
$made = Join-Path $tmp 'made.rtf'; Set-Content $made 'new'
New-Item -ItemType Directory -Force (Join-Path $tmp 'Originals') | Out-Null
$lock = [IO.File]::Open($orig, 'Open', 'Read', 'None')
try {
    $err = ''
    try { [void](Move-IntoPlace ([pscustomobject]@{ Name = 'Busy.txt'; Path = $orig; OutputName = 'Busy.rtf' }) $made $tmp (Join-Path $tmp 'Originals')) }
    catch { $err = $_.Exception.Message }
} finally { $lock.Close() }
Check ($err -like 'MOVE:*') "an original open in another program is reported ($err)"
Check ((Test-Path $orig) -and -not (Test-Path (Join-Path $tmp 'Originals\Busy.txt'))) 'that original stays where it was'
Check (-not (Test-Path $made) -and -not (Test-Path (Join-Path $tmp 'Busy.rtf'))) 'and no half-done new file is left behind'
Remove-Item -LiteralPath $tmp -Recurse -Force

$root = Join-Path ([IO.Path]::GetTempPath()) "reformatter-test-$PID"
$docs = Join-Path $root 'docs'
New-Item -ItemType Directory -Force $docs | Out-Null
$word = Start-Word
try {
    $rtf = 6; $docx = 16
    # 1. 40 paragraphs run together; each lost break left two spaces
    $lost = New-Story 40 '  '
    New-WordFile $word (Join-Path $docs 'Lost breaks.rtf') { param($d) $d.Content.Text = "A Heading Line`r" + $lost; $d.Paragraphs.Item(1).Style = -2 } $rtf
    # 2. 30 paragraphs run together with only one space: no marks at all
    $none = New-Story 30 ' '
    New-WordFile $word (Join-Path $docs 'No marks.docx') { param($d) $d.Content.Text = $none } $docx
    # 3. A writer who puts two spaces after every sentence, all run together:
    #    the two spaces mean nothing here, so it is only split every few sentences
    $habit = New-Story 20 '  ' '  '
    New-WordFile $word (Join-Path $docs 'Two spaces habit.docx') { param($d) $d.Content.Text = $habit } $docx
    # 4. Normal short paragraphs and a long table cell, which must stay as they are
    New-WordFile $word (Join-Path $docs 'Normal.docx') {
        param($d)
        $d.Content.Text = (New-Story 8 "`r") + "`r"
        $t = $d.Tables.Add($d.Paragraphs.Item($d.Paragraphs.Count).Range, 1, 2)
        $t.Cell(1, 1).Range.Text = New-Story 10 '  '
        $t.Cell(1, 2).Range.Text = 'Short cell'
    } $docx
    # 5. Plain text with the lost breaks, and an old copy of its name already in Originals
    [IO.File]::WriteAllText((Join-Path $docs 'Notes.txt'), (New-Story 12 '  '), (New-Object Text.UTF8Encoding($true)))
    New-Item -ItemType Directory -Force (Join-Path $docs 'Originals') | Out-Null
    Set-Content -LiteralPath (Join-Path $docs 'Originals\Notes.txt') -Value 'an older original'
    # 6. An empty file (skipped) that shares its name with a real document
    New-Item -ItemType File (Join-Path $docs 'Report.rtf') | Out-Null
    New-WordFile $word (Join-Path $docs 'Report.docx') { param($d) $d.Content.Text = 'A small report.' } $docx
} finally { Stop-Word $word }

$settings = @{ FontName = 'Calibri'; FontSize = 12; Headings = 'Keep'; LineSpacing = 'Leave'; SpaceBetween = '6'; Indent = 'Leave'
               Align = 'Leave'; TextColour = 'Leave'; RemoveHighlight = $false; ClearEmphasis = $false; EvenSpacing = $false
               RemoveEmptyLines = $true; WidowControl = $false; RemoveLinks = $false; SplitParagraphs = $true; JoinLines = $false; PaperSize = 'Leave'; Margins = 'Leave' }
$items = Find-Documents $docs
$results = Invoke-Reformat -Items $items -Settings $settings -Folder $docs
$byName = @{}; foreach ($r in $results) { $byName[$r.Name] = $r }

Write-Host 'Files'
foreach ($n in 'Lost breaks.rtf', 'No marks.docx', 'Two spaces habit.docx', 'Normal.docx', 'Notes.txt', 'Report.docx') {
    Check ($byName[$n].Status -eq 'Converted') "$n was converted ($($byName[$n].Status) $($byName[$n].Detail))"
}
Check ($byName['Report.rtf'].Status -eq 'Skipped') 'the empty Report.rtf was skipped'
Check (Test-Path -LiteralPath (Join-Path $docs 'Report.rtf')) 'the empty Report.rtf is still in the folder'
Check ($byName['Report.docx'].Output -eq 'Report (docx).rtf') "Report.docx became 'Report (docx).rtf', not over the empty Report.rtf"
foreach ($n in 'Lost breaks.rtf', 'No marks.docx', 'Two spaces habit.docx', 'Normal.docx', 'Report.docx') {
    Check (Test-Path -LiteralPath (Join-Path $docs "Originals\$n")) "original $n moved into Originals"
}
Check ((Get-Content -LiteralPath (Join-Path $docs 'Originals\Notes.txt')) -eq 'an older original') 'the older Originals\Notes.txt was not overwritten'
Check (Test-Path -LiteralPath (Join-Path $docs 'Originals\Notes (2).txt')) 'the new original was kept as Originals\Notes (2).txt'
Check ($byName['Notes.txt'].Original -eq 'Notes (2).txt') 'the summary knows the original was kept as Notes (2).txt'
foreach ($n in 'Lost breaks.rtf', 'No marks.rtf', 'Two spaces habit.rtf', 'Normal.rtf', 'Notes.rtf') {
    Check (Test-Path -LiteralPath (Join-Path $docs $n)) "reformatted $n is in the folder"
}
Check (-not (Test-Path -LiteralPath (Join-Path $docs 'fixed'))) 'no "fixed" folder is made any more'

$word = Start-Word
try {
    Write-Host 'Paragraphs'
    $p = Get-Paragraphs $word (Join-Path $docs 'Lost breaks.rtf')
    Check ($p[0].Heading -and $p[0].Text -eq 'A Heading Line') 'the heading is kept'
    $body = @($p | Where-Object { -not $_.Heading -and $_.Text })
    Check ($body.Count -eq 40) "lost breaks: 40 paragraphs, as written (got $($body.Count))"
    Check (@($body | Where-Object { $_.Text -notmatch '^Paragraph \d+ begins' }).Count -eq 0) 'lost breaks: every paragraph starts where the writer started it'
    Check (@($body | Where-Object { $_.Text -match '  ' }).Count -eq 0) 'lost breaks: no double spaces left behind'

    $p = Get-Paragraphs $word (Join-Path $docs 'No marks.rtf'); $p = @($p | Where-Object { $_.Text })
    $longest = ($p | ForEach-Object { $_.Text.Length } | Measure-Object -Maximum).Maximum
    Check ($p.Count -ge 8) "no marks: split into readable pieces ($($p.Count) paragraphs)"
    Check ($longest -le 1400) "no marks: no paragraph longer than 1400 characters (longest $longest)"
    Check (@($p | Where-Object { $_.Text -notmatch "^[A-Z$q]" -or $_.Text -notmatch "[.$qq]$" }).Count -eq 0) 'no marks: every piece starts a sentence and ends one'
    Check ((($p | ForEach-Object { $_.Text }) -join ' ') -eq $none) 'no marks: not a word lost or changed'

    $p = Get-Paragraphs $word (Join-Path $docs 'Two spaces habit.rtf'); $p = @($p | Where-Object { $_.Text })
    Check ($p.Count -ge 3 -and $p.Count -le 20) "two-space habit: split every few sentences, not at every sentence ($($p.Count) paragraphs from 100 sentences)"

    $p = Get-Paragraphs $word (Join-Path $docs 'Normal.rtf')
    $plain = @($p | Where-Object { -not $_.Table -and $_.Text })
    Check ($plain.Count -eq 8) "normal: the 8 paragraphs are unchanged (got $($plain.Count))"
    $cells = @($p | Where-Object { $_.Table })
    Check (@($cells | Where-Object { $_.Text.Length -gt 1500 }).Count -eq 1) 'normal: the long table cell is not split'

    $p = Get-Paragraphs $word (Join-Path $docs 'Notes.rtf'); $p = @($p | Where-Object { $_.Text })
    Check ($p.Count -eq 12) "plain text: 12 paragraphs (got $($p.Count))"
} finally { Stop-Word $word }

Write-Host 'Joining lines, in Word'
$jdir = Join-Path $root 'join'
New-Item -ItemType Directory -Force $jdir | Out-Null
$word = Start-Word
try {
    # Mixed line ends inside paragraphs, mixed gaps between them, a heading and a table
    $i = 0; $body = ''
    foreach ($k in 0..4) {
        $body += (@(Wrap $plainParas[$k]) -join $lineEnds[$i++ % 4])
        if ($k -lt 4) { $body += @("`r`r", "`r `r", "`v`v", "`v`r`r")[$k] }
    }
    New-WordFile $word (Join-Path $jdir 'Broken lines.docx') {
        param($d)
        $d.Content.Text = "Chapter One`r" + $body + "`r"
        $d.Paragraphs.Item(1).Style = -2
        $t = $d.Tables.Add($d.Paragraphs.Item($d.Paragraphs.Count).Range, 1, 1)
        $t.Cell(1, 1).Range.Text = "Cell line one`vCell line two"
    } $docx
    # One Word paragraph per line; the first line of each real paragraph is indented by Word
    New-WordFile $word (Join-Path $jdir 'Indented.docx') {
        param($d)
        $d.Content.Text = (@(foreach ($para in $plainParas) { (Wrap $para) -join "`r" }) -join "`r")
        $n = 1
        foreach ($para in $plainParas) { $d.Paragraphs.Item($n).FirstLineIndent = 36; $n += @(Wrap $para).Count }
    } $docx
    # A plain-text file with blank lines between paragraphs
    [IO.File]::WriteAllText((Join-Path $jdir 'Broken.txt'), ((@(foreach ($para in $plainParas) { (Wrap $para) -join "`r`n" }) -join "`r`n`r`n") + "`r`n"))
    # A normal document, which must not change
    New-WordFile $word (Join-Path $jdir 'Normal paragraphs.docx') { param($d) $d.Content.Text = ($plainParas -join "`r") } $docx
} finally { Stop-Word $word }

$settings.JoinLines = $true
$settings.SplitParagraphs = $true
$results = Invoke-Reformat -Items (Find-Documents $jdir) -Settings $settings -Folder $jdir
$byName = @{}; foreach ($r in $results) { $byName[$r.Name] = $r }
foreach ($n in 'Broken lines.docx', 'Indented.docx', 'Broken.txt', 'Normal paragraphs.docx') {
    Check ($byName[$n].Status -eq 'Converted') "$n was converted ($($byName[$n].Status) $($byName[$n].Detail))"
}
$word = Start-Word
try {
    Check ($byName['Broken lines.docx'].Detail -match 'lines joined: \d+ line breaks made spaces') "the summary says the lines were joined ($($byName['Broken lines.docx'].Detail))"
    $p = Get-Paragraphs $word (Join-Path $jdir 'Broken lines.rtf')
    $plain = @($p | Where-Object { -not $_.Heading -and -not $_.Table -and $_.Text })
    Check ($p[0].Heading -and $p[0].Text -eq 'Chapter One') 'the heading keeps its own line'
    Check (($plain | ForEach-Object { $_.Text }) -join '|' -ceq ($plainParas -join '|')) "every kind of line end joined, every gap one paragraph end ($($plain.Count) paragraphs)"
    $cell = @($p | Where-Object { $_.Table -and $_.Text }) | Select-Object -First 1
    Check ($cell -and $cell.Text -ceq "Cell line one`vCell line two") 'the line break inside the table stays'

    $d = $word.Documents.Open((Join-Path $jdir 'Indented.rtf'), $false, $true, $false)
    $ind = @(foreach ($q2 in $d.Paragraphs) { if (([string]$q2.Range.Text).Trim()) { [pscustomobject]@{ Text = ([string]$q2.Range.Text).TrimEnd([char]13); Indent = [double]$q2.FirstLineIndent } } })
    $d.Close(0)
    Check (($ind | ForEach-Object { $_.Text }) -join '|' -ceq ($plainParas -join '|')) "Word-indented paragraphs: a new paragraph starts at each indent ($($ind.Count) paragraphs)"
    Check (@($ind | Where-Object { $_.Indent -ne 36 }).Count -eq 0) "and each keeps its indent ($(($ind | ForEach-Object { $_.Indent }) -join ', '))"

    $p = Get-Paragraphs $word (Join-Path $jdir 'Broken.rtf'); $p = @($p | Where-Object { $_.Text })
    Check (($p | ForEach-Object { $_.Text }) -join '|' -ceq ($plainParas -join '|')) "plain text: lines joined, blank lines make paragraphs ($($p.Count) paragraphs)"

    $p = Get-Paragraphs $word (Join-Path $jdir 'Normal paragraphs.rtf'); $p = @($p | Where-Object { $_.Text })
    Check (($p | ForEach-Object { $_.Text }) -join '|' -ceq ($plainParas -join '|')) 'a normal document is left as it was'
    Check ($byName['Normal paragraphs.docx'].Detail -notmatch 'joined') 'and the summary does not say it was joined'
} finally { Stop-Word $word }
$settings.JoinLines = $false
$settings.SplitParagraphs = $true

Write-Host 'Turned off'
$off = Join-Path $root 'off'
New-Item -ItemType Directory -Force $off | Out-Null
Copy-Item -LiteralPath (Join-Path $docs 'Originals\Lost breaks.rtf') -Destination $off
Copy-Item -LiteralPath (Join-Path $jdir 'Originals\Broken.txt') -Destination $off
$settings.SplitParagraphs = $false
$settings.JoinLines = $false
$results = Invoke-Reformat -Items (Find-Documents $off) -Settings $settings -Folder $off
$word = Start-Word
try {
    $p = Get-Paragraphs $word (Join-Path $off 'Lost breaks.rtf'); $p = @($p | Where-Object { $_.Text })
    Check ($p.Count -eq 2) "with the tick box off, the giant paragraph stays as it was ($($p.Count) paragraphs)"
    $p = Get-Paragraphs $word (Join-Path $off 'Broken.rtf'); $p = @($p | Where-Object { $_.Text })
    $lineCount = @(foreach ($para in $plainParas) { Wrap $para }).Count
    Check ($p.Count -eq $lineCount) "with the join tick box off, every line stays on its own ($($p.Count) of $lineCount)"
} finally { Stop-Word $word }

Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
if ($failures) { Write-Host "$failures check(s) FAILED"; exit 1 }
Write-Host 'All checks passed'
exit 0
