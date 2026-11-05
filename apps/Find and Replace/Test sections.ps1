# Test sections.ps1
# Checks "delete the sections between two phrases" in ReplaceEngine.ps1: first on
# plain text, then in real .rtf, .docx and .doc files using Microsoft Word, in a
# temporary folder. Run it with Windows PowerShell, from any folder:
#   powershell -NoProfile -STA -ExecutionPolicy Bypass -File "Test sections.ps1"
# It prints each check, then how many failed, and exits with 1 if any did.
$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $AppDir 'ReplaceEngine.ps1')

$script:Failed = 0
function Show-Plain([string]$Text) { return $Text.Replace("`r", '<CR>').Replace("`n", '<LF>').Replace([string][char]7, '<CELL>') }
function Test-Same([string]$Name, $Got, $Want) {
    if ([string]$Got -ceq [string]$Want) { "  ok    $Name"; return }
    $script:Failed++
    "  FAIL  $Name"
    "        got  [$(Show-Plain ([string]$Got))]"
    "        want [$(Show-Plain ([string]$Want))]"
}
function Get-Off { return @{ On = $false; Find = ''; Replace = ''; Include = $false; MatchCase = $false; WholeWords = $false; Which = 'First'; Source = 'Text'; File = '' } }
function New-Settings([string]$From, [string]$To, [string]$Keep = 'both', [bool]$MatchCase = $false) {
    return @{
        Words = Get-Off; Begin = Get-Off; End = Get-Off
        Between = @{ On = $true; From = $From; To = $To; MatchCase = $MatchCase
                     IncludeFrom = ($Keep -eq 'last' -or $Keep -eq 'none'); IncludeTo = ($Keep -eq 'first' -or $Keep -eq 'none') }
    }
}
# What a text becomes. $Keep says which phrases stay: both, first, last or none.
function Get-TextAfter([string]$Text, $S) {
    $doc = [pscustomobject]@{ Text = $Text; Eol = "`n" }
    if ($Text.Contains("`r`n")) { $doc.Eol = "`r`n" }
    $state = @{ Word = $null; WordTried = $false; Sources = @{} }
    $plan = Get-TextPlan $state $doc $S $true
    if (-not $plan.Changes) { return $Text }
    return (Get-TextResult $state $doc $plan $S)
}

"--- pairing the phrases"
$r = Get-SectionPairs @(@(0, 1), @(4, 5), @(9, 10)) (,@(6, 7))
Test-Same 'a first phrase inside a section goes with it' (($r.Pairs | ForEach-Object { $_ -join ',' }) -join ' ') '0,1,6,7'
Test-Same 'a first phrase with no last phrase after it is counted' $r.Unclosed 1
$r = Get-SectionPairs @(@(0, 3), @(6, 9), @(12, 15), @(18, 21)) @(@(0, 3), @(6, 9), @(12, 15), @(18, 21))
Test-Same 'the same phrase at both ends pairs up in twos' (($r.Pairs | ForEach-Object { $_ -join ',' }) -join ' ') '0,3,6,9 12,15,18,21'

"--- plain text: which phrases stay"
Test-Same 'both stay, one space between them' (Get-TextAfter 'aaa FROM bbb TO ccc' (New-Settings 'FROM' 'TO')) 'aaa FROM TO ccc'
Test-Same 'neither stays, no doubled space' (Get-TextAfter 'aaa FROM bbb TO ccc' (New-Settings 'FROM' 'TO' 'none')) 'aaa ccc'
Test-Same 'first stays' (Get-TextAfter 'aaa FROM bbb TO ccc' (New-Settings 'FROM' 'TO' 'first')) 'aaa FROM ccc'
Test-Same 'last stays' (Get-TextAfter 'aaa FROM bbb TO ccc' (New-Settings 'FROM' 'TO' 'last')) 'aaa TO ccc'
Test-Same 'brackets with no spaces stay touching' (Get-TextAfter 'a [x] b [yy] c' (New-Settings '[' ']')) 'a [] b [] c'
Test-Same 'several sections, same phrase at both ends' (Get-TextAfter '*** a *** b *** c ***' (New-Settings '***' '***')) '*** *** b *** ***'

"--- plain text: joining up the text on each side"
Test-Same 'whole line goes with its line break' (Get-TextAfter "one`nFROM x TO`ntwo`n" (New-Settings 'FROM' 'TO' 'none')) "one`ntwo`n"
Test-Same 'whole line at the very start' (Get-TextAfter "FROM x TO`ntwo`n" (New-Settings 'FROM' 'TO' 'none')) "two`n"
Test-Same 'whole last line, file end stays' (Get-TextAfter "one`nFROM x TO`n" (New-Settings 'FROM' 'TO' 'none')) "one`n"
Test-Same 'several lines, phrases stay on their own lines' (Get-TextAfter "FROM`nb1`nb2`nTO`nrest`n" (New-Settings 'FROM' 'TO')) "FROM`nTO`nrest`n"
Test-Same 'the same with Windows line endings' (Get-TextAfter "FROM`r`nb1`r`nb2`r`nTO`r`nrest`r`n" (New-Settings 'FROM' 'TO')) "FROM`r`nTO`r`nrest`r`n"
Test-Same 'first stays, section ends a line' (Get-TextAfter "FROM`nb`nTO`nrest`n" (New-Settings 'FROM' 'TO' 'first')) "FROM`nrest`n"
Test-Same 'first stays, Windows line endings' (Get-TextAfter "FROM`r`nb`r`nTO`r`nrest`r`n" (New-Settings 'FROM' 'TO' 'first')) "FROM`r`nrest`r`n"
Test-Same 'last stays, section starts a line' (Get-TextAfter "a`nFROM b`nTO c`n" (New-Settings 'FROM' 'TO' 'last')) "a`nTO c`n"
Test-Same 'text before a comma closes up' (Get-TextAfter 'see FROM x TO, then' (New-Settings 'FROM' 'TO' 'none')) 'see, then'
Test-Same 'first stays before a comma' (Get-TextAfter 'see FROM x TO, then' (New-Settings 'FROM' 'TO' 'first')) 'see FROM, then'
Test-Same 'section ends the file, no trailing space' (Get-TextAfter 'aaa FROM x TO' (New-Settings 'FROM' 'TO' 'none')) 'aaa'
Test-Same 'two sections in a row do not overlap' (Get-TextAfter 'A x B A y B' (New-Settings 'A' 'B' 'none' $true)) ''
Test-Same 'tabs found by a code' (Get-TextAfter "x`ty`tz" (New-Settings '^t' '^t')) "x`t`tz"

"--- plain text: what is left alone"
Test-Same 'no last phrase: nothing changes' (Get-TextAfter 'FROM and more' (New-Settings 'FROM' 'TO')) 'FROM and more'
Test-Same 'last phrase only before the first: nothing changes' (Get-TextAfter 'TO then FROM' (New-Settings 'FROM' 'TO')) 'TO then FROM'
Test-Same 'phrases already next to each other: nothing changes' (Get-TextAfter 'FROM TO' (New-Settings 'FROM' 'TO')) 'FROM TO'
Test-Same 'first phrase left over stays' (Get-TextAfter 'A x B y A z' (New-Settings 'A' 'B' 'none' $true)) 'y A z'
Test-Same 'match case off finds any capitals' (Get-TextAfter 'a from x to b' (New-Settings 'FROM' 'TO' 'none')) 'a b'
Test-Same 'match case on does not' (Get-TextAfter 'a from x to b' (New-Settings 'FROM' 'TO' 'none' $true)) 'a from x to b'

"--- plain text: together with the other kinds of replacing"
$S = New-Settings 'FROM' 'TO'
$S.Words = @{ On = $true; Find = 'x'; Replace = 'y'; MatchCase = $false; WholeWords = $false }
$doc = [pscustomobject]@{ Text = 'x FROM x TO x'; Eol = "`n" }
$state = @{ Word = $null; WordTried = $false; Sources = @{} }
$plan = Get-TextPlan $state $doc $S $true
Test-Same 'word replacing skips what is deleted' (Get-TextResult $state $doc $plan $S) 'y FROM TO y'
Test-Same 'and does not count it' $plan.WordsCount 2
$S = New-Settings 'PP' 'QQ' 'none' $true
$S.Begin = @{ On = $true; Find = 'START'; Include = $false; MatchCase = $true; Which = 'First'; Source = 'Text'; Replace = ''; File = '' }
$S.End = @{ On = $true; Find = 'STOP'; Include = $false; MatchCase = $true; Which = 'Last'; Source = 'Text'; Replace = ''; File = '' }
Test-Same 'only sections in the text that stays go' (Get-TextAfter 'PP j QQ START PP x QQ STOP PP y QQ' $S) 'START STOP'

"--- preview"
$doc = [pscustomobject]@{ Text = 'aaa FROM bbb TO ccc'; Eol = "`n" }
$plan = Get-TextPlan $state $doc (New-Settings 'FROM' 'TO') $true
Test-Same 'summary' $plan.Summary '1 section deleted (1 word)'
Test-Same 'taken-out text is shown' (($plan.Segments | Where-Object { $_.Kind -eq 'Old' } | ForEach-Object { $_.Text }) -join '|') ' bbb'
$doc = [pscustomobject]@{ Text = 'FROM only'; Eol = "`n" }
Test-Same 'summary when the last phrase is missing' (Get-TextPlan $state $doc (New-Settings 'FROM' 'TO') $false).Summary 'sections: "TO" not found after "FROM"'
$doc = [pscustomobject]@{ Text = 'nothing here'; Eol = "`n" }
Test-Same 'summary when the first phrase is missing' (Get-TextPlan $state $doc (New-Settings 'FROM' 'TO') $false).Summary 'sections: "FROM" not found'

# ------------------------------------------------------------ the checks in the window

"--- what the window asks for before the preview"
$errs = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $AppDir 'Find and Replace.ps1'), [ref]$null, [ref]$errs)
$fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SettingsProblem' }, $true)
. ([scriptblock]::Create($fn.Extent.Text))
Test-Same 'first phrase missing' (Get-SettingsProblem (New-Settings '' 'TO') 1) 'Type both phrases that the sections are between (or untick "Sections between two phrases").'
Test-Same 'last phrase missing' (Get-SettingsProblem (New-Settings 'FROM' '') 1) 'Type both phrases that the sections are between (or untick "Sections between two phrases").'
Test-Same 'a code that does not exist' ((Get-SettingsProblem (New-Settings 'FROM' '^z') 1) -split "`n")[0] '"^z" in "Last phrase" is not a special code that works there.'
Test-Same 'both phrases typed' (Get-SettingsProblem (New-Settings 'FROM' '^p') 1) ''
$S = New-Settings 'FROM' 'TO'; $S.Between.On = $false
Test-Same 'nothing ticked' (Get-SettingsProblem $S 1) 'Tick at least one of the four: "Words, phrases or marks", "The beginning", "Sections between two phrases" or "The ending".'

# ------------------------------------------------------------ real documents, using Word

$dir = Join-Path ([IO.Path]::GetTempPath()) ('FRsections' + [guid]::NewGuid().ToString('N').Substring(0, 8))
[void][IO.Directory]::CreateDirectory($dir)
$rtf = @'
{\rtf1\ansi\deff0{\fonttbl{\f0 Calibri;}}
{\stylesheet{\s0 Normal;}{\s1\fs32\b heading 1;}}
{\header\pard\plain\f0\fs18 Header FROM stays UNTIL here\par}
\pard\plain\s1\f0\fs32\b Heading FROM one\par
\pard\plain\f0\fs22 Goes one\par
\pard\plain\f0\fs22 Goes UNTIL two\par
\pard\plain\f0\fs22 Mid aaa FROM bbb UNTIL ccc end.\par
\pard\plain\f0\fs22 FROM whole paragraph UNTIL\par
\pard\plain\f0\fs22 Before table\par
\trowd\cellx3000\cellx6000\pard\intbl\plain\f0 CELL1 FROM in cell UNTIL stays\cell\pard\intbl\plain\f0 FROM across\cell\row
\trowd\cellx3000\cellx6000\pard\intbl\plain\f0 UNTIL out\cell\pard\intbl\plain\f0 B2\cell\row
\pard\plain\f0\fs22 FROM into the UNTIL table\par
\pard\plain\f0\fs22 FROM last whole paragraph UNTIL\par
}
'@
$word = $null
function Get-WordAfter([string]$Name, $S) {
    $items = @(New-DocItem (Join-Path $dir $Name))
    Invoke-Documents -Items $items -Settings $S
    Invoke-Documents -Items $items -Settings $S -Apply -SaveMode 'Copy'
    $saved = [string]$items[0].Result.SavedTo
    if (-not $saved) { return [pscustomobject]@{ Summary = $items[0].Result.Detail; Paragraphs = @(); Tables = -1; Header = ''; Format = -1 } }
    $doc = Open-InWord $word $saved
    try {
        $paras = @(foreach ($p in $doc.Paragraphs) { $p.Style.NameLocal + ':' + ([string]$p.Range.Text).TrimEnd("`r", [char]7) })
        return [pscustomobject]@{
            Summary = $items[0].Result.Detail; Paragraphs = $paras; Tables = $doc.Tables.Count
            Header = ([string]$doc.Sections.Item(1).Headers.Item(1).Range.Text).Trim(); Format = $doc.SaveFormat }
    } finally { Close-WordDocument $doc }
}
try {
    $word = Start-Word
    [IO.File]::WriteAllText((Join-Path $dir 'a.rtf'), $rtf, [Text.Encoding]::ASCII)
    $d = Open-InWord $word (Join-Path $dir 'a.rtf')
    $m = [Type]::Missing
    $d.SaveAs2([string](Join-Path $dir 'a.docx'), 16, $m, $m, $false)
    $d.SaveAs2([string](Join-Path $dir 'a.doc'), 0, $m, $m, $false)
    Close-WordDocument $d

    foreach ($name in 'a.rtf', 'a.docx', 'a.doc') {
        "--- Word, $name, both phrases stay"
        $r = Get-WordAfter $name (New-Settings 'FROM' 'UNTIL')
        Test-Same 'header is not touched' $r.Header 'Header FROM stays UNTIL here'
        Test-Same 'file type is kept' $r.Format @{ 'a.rtf' = 6; 'a.docx' = 12; 'a.doc' = 0 }[$name]
        Test-Same 'what is left' ($r.Paragraphs -join ' | ') ('Normal:Heading FROM UNTIL two | Normal:Mid aaa FROM UNTIL ccc end. | Normal:FROM UNTIL | Normal:Before table | ' +
            'Normal:CELL1 FROM UNTIL stays | Normal:FROM across | Normal: | Normal:UNTIL out | Normal:B2 | Normal: | Normal:FROM UNTIL table | Normal:FROM UNTIL')
    }

    "--- Word, a.docx, neither phrase stays"
    $r = Get-WordAfter 'a.docx' (New-Settings 'FROM' 'UNTIL' 'none')
    Test-Same 'what is left' ($r.Paragraphs -join ' | ') ('Normal:Heading two | Normal:Mid aaa ccc end. | Normal:Before table | Normal:CELL1 stays | ' +
        'Normal:FROM across | Normal: | Normal:UNTIL out | Normal:B2 | Normal: | Normal:table')
    Test-Same 'the table stays' $r.Tables 1

    "--- Word, a.docx, one phrase stays"
    $r = Get-WordAfter 'a.docx' (New-Settings 'FROM whole' 'paragraph UNTIL' 'first')
    Test-Same 'first stays: nothing joins' $r.Paragraphs[4] 'Normal:FROM whole'
    $r = Get-WordAfter 'a.docx' (New-Settings 'Goes one' 'Mid aaa' 'last')
    Test-Same 'last stays: the paragraph before is not touched' ($r.Paragraphs[0..1] -join ' | ') 'Heading 1:Heading FROM one | Normal:Mid aaa FROM bbb UNTIL ccc end.'

    "--- Word, a section crossing the edge of a table stays"
    $r = Get-WordAfter 'a.docx' (New-Settings 'across' 'out')
    Test-Same 'two cells' $r.Summary 'sections: nothing to delete between the phrases (1 section crosses the edge of a table, left as is)'
    $r = Get-WordAfter 'a.docx' (New-Settings 'Before table' 'stays')
    Test-Same 'outside to inside' $r.Summary 'sections: nothing to delete between the phrases (1 section crosses the edge of a table, left as is)'
    $r = Get-WordAfter 'a.docx' (New-Settings 'Before table' 'into the' 'none')
    Test-Same 'a whole table inside a section goes' $r.Tables 0
    Test-Same 'and the text after it stays' $r.Paragraphs[5] 'Normal:UNTIL table'
    $r = Get-WordAfter 'a.docx' (New-Settings 'FROM whole' 'Before table' 'none')
    Test-Same 'a whole paragraph right before a table goes' ($r.Paragraphs[3..4] -join ' | ') 'Normal:Mid aaa FROM bbb UNTIL ccc end. | Normal:CELL1 FROM in cell UNTIL stays'
    Test-Same 'and the table stays' $r.Tables 1

    "--- Word, the document's last paragraph"
    $r = Get-WordAfter 'a.docx' (New-Settings 'FROM last' 'paragraph UNTIL' 'none')
    Test-Same 'no empty paragraph is left' $r.Paragraphs[$r.Paragraphs.Count - 1] 'Normal:FROM into the UNTIL table'

    "--- changing the files themselves, and putting them back"
    $hash = { param($f) [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([IO.File]::ReadAllBytes($f))) }
    [IO.File]::WriteAllText((Join-Path $dir 'n.txt'), ("Caf" + [char]0xE9 + " FROM x UNTIL end`r`nnext`r`n"), (New-Object Text.UTF8Encoding($false)))
    $paths = @((Join-Path $dir 'a.docx'), (Join-Path $dir 'n.txt'))
    $before = @($paths | ForEach-Object { & $hash $_ })
    $items = @($paths | ForEach-Object { New-DocItem $_ })
    $S = New-Settings 'FROM' 'UNTIL' 'none'
    Invoke-Documents -Items $items -Settings $S
    Invoke-Documents -Items $items -Settings $S -Apply -SaveMode 'InPlace' -Stamp 'test'
    Test-Same 'the Word file is changed' $items[0].Result.Status 'Done'
    Test-Same 'the .txt file keeps its letters and line breaks' ([IO.File]::ReadAllText($paths[1])) ("Caf" + [char]0xE9 + " end`r`nnext`r`n")
    [void](Restore-Originals $items)
    Test-Same 'the Word file is put back exactly' (& $hash $paths[0]) $before[0]
    Test-Same 'the .txt file is put back exactly' (& $hash $paths[1]) $before[1]

    "--- Word, with the other kinds of replacing"
    $S = New-Settings 'FROM' 'UNTIL'
    $S.Words = @{ On = $true; Find = 'bbb'; Replace = 'NO'; MatchCase = $false; WholeWords = $false }
    $r = Get-WordAfter 'a.docx' $S
    Test-Same 'word replacing skips what is deleted' ($r.Summary -replace '\(\d+ words\)', '(N words)') '6 sections deleted (N words) (1 section crosses the edge of a table, left as is);  "bbb" not found'
    $S = New-Settings 'aaa' 'ccc'
    $S.Words = @{ On = $true; Find = 'end'; Replace = 'END'; MatchCase = $true; WholeWords = $false }
    $S.Begin = @{ On = $true; Find = 'Mid'; Include = $false; MatchCase = $true; Which = 'First'; Source = 'Text'; Replace = 'New start'; File = '' }
    $S.End = @{ On = $true; Find = 'Before table'; Include = $true; MatchCase = $true; Which = 'First'; Source = 'Text'; Replace = ''; File = '' }
    $r = Get-WordAfter 'a.docx' $S
    Test-Same 'beginning, section, word and ending all at once' ($r.Paragraphs -join ' | ') 'Normal:New start | Normal:Mid aaa ccc END. | Normal:FROM whole paragraph UNTIL'
} finally {
    Stop-Word $word
    try { [IO.Directory]::Delete($dir, $true) } catch {}
}

""
if ($script:Failed) { "$script:Failed check(s) FAILED"; exit 1 }
"All checks passed"
