# Test Renamer.ps1
# Checks the naming choices in RenameEngine.ps1 on test files it makes in a
# temporary folder. Run it in Windows PowerShell with -STA:
#   powershell.exe -STA -ExecutionPolicy Bypass -File "Test Renamer.ps1"
# It prints "ok" or "FAIL" for each check, and leaves your own files alone.
$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $AppDir 'RenameEngine.ps1')
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Root = Join-Path ([IO.Path]::GetTempPath()) 'DocumentRenamerTest'
if ([IO.Directory]::Exists($Root)) { [IO.Directory]::Delete($Root, $true) }
[void][IO.Directory]::CreateDirectory($Root)
$script:HistoryFile = Join-Path $Root 'history.json'
$script:Failures = 0

function Check([string]$Label, $Got, $Want) {
    if ([string]$Got -ceq [string]$Want) { "ok    $Label" }
    else { $script:Failures++; "FAIL  $Label`n        got:  [$Got]`n        want: [$Want]" }
}

# The new names a plan gives, in its order, as "old>new" (or "old>(keep)")
function Get-PlanText($Plan) {
    return (@($Plan | ForEach-Object { if ($_.Action -eq 'Keep') { $_.Item.Name + '>(keep)' } else { $_.Item.Name + '>' + $_.NewName } }) -join ' | ')
}

function New-Folder([string]$Name) {
    $f = Join-Path $Root $Name
    [void][IO.Directory]::CreateDirectory($f)
    return $f
}

function New-TextFile([string]$Folder, [string]$Name, [string]$Text) {
    $p = Join-Path $Folder $Name
    [IO.File]::WriteAllText($p, $Text, (New-Object Text.UTF8Encoding($false)))
    return $p
}

# A small Word document with an author and title, built without Word
function New-Docx([string]$Path, [string]$Text, [string]$Author, [string]$Title) {
    $parts = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'
        'word/document.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' + $Text + '</w:t></w:r></w:p></w:body></w:document>'
        'docProps/core.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + $Title + '</dc:title><dc:creator>' + $Author + '</dc:creator></cp:coreProperties>'
    }
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($k in $parts.Keys) {
            $w = New-Object IO.StreamWriter(($zip.CreateEntry($k)).Open(), (New-Object Text.UTF8Encoding($false)))
            $w.Write($parts[$k]); $w.Dispose()
        }
    } finally { $zip.Dispose() }
}

# A tiny photo whose camera details say when it was taken
function New-Photo([string]$Path, [string]$Taken) {
    $bmp = New-Object System.Drawing.Bitmap 8, 8
    try {
        $item = [Runtime.Serialization.FormatterServices]::GetUninitializedObject([System.Drawing.Imaging.PropertyItem])
        $value = [Text.Encoding]::ASCII.GetBytes($Taken + [char]0)
        $item.Id = 0x9003; $item.Type = 2; $item.Value = $value; $item.Len = $value.Length
        $bmp.SetPropertyItem($item)
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    } finally { $bmp.Dispose() }
}

function Get-Plan($Folder, [hashtable]$Options, [string]$Mode = 'Copy', [switch]$AllFiles, [switch]$Read) {
    $items = Find-Files $Folder -AllFiles:$AllFiles
    if ($Read) { Read-DocumentStarts -Items $items }
    Read-FileProperties $items $Folder
    $existing = @(Get-ChildItem -LiteralPath $Folder -Force | ForEach-Object { $_.Name })
    return Get-RenamePlan -Items $items -Options $Options -Mode $Mode -ExistingNames $existing -Folder $Folder
}

try {
    '--- names made from text (the original way)'
    $U = [char]0xD83D; $L = [char]0xDE00   # an emoji (two halves)
    Check 'trailing period/space' (Get-NameFromText (Get-CleanText 'Hello world. Bye') 12) 'Hello world'
    Check 'leading bullet' (Get-NameFromText (Get-CleanText "  - Buy milk`r`n- eggs") 30) 'Buy milk - eggs'
    Check 'quotes and slashes' (Get-NameFromText (Get-CleanText 'He said "a/b\c|d" <ok>*?') 40) "He said 'a-b-c-d' ok"
    Check 'colons' (Get-NameFromText (Get-CleanText 'Time: 10:30 am. Note:') 40) 'Time - 10-30 am. Note'
    Check 'nbsp and tabs' (Get-NameFromText (Get-CleanText ('a' + [char]0xA0 + [char]0xA0 + "`tb`v`fc")) 30) 'a b c'
    Check 'emoji not cut' (Get-NameFromText (Get-CleanText ('abcd' + $U + $L + 'efg')) 5) 'abcd'
    Check 'emoji kept' (Get-NameFromText (Get-CleanText ('abcd' + $U + $L + 'efg')) 6) ('abcd' + $U + $L)
    Check 'reserved name' (Get-AllowedName (Get-NameFromText (Get-CleanText 'nul') 30)) '_nul'
    Check 'reserved with dot' (Get-AllowedName 'COM1.backup notes') '_COM1.backup notes'
    Check 'not reserved' (Get-AllowedName 'Console log') 'Console log'
    Check 'only symbols' (Get-NameFromText (Get-CleanText ' ... ??? *** ') 30) ''
    Check 'exactly 30' (Get-NameFromText (Get-CleanText ('x' * 50)) 30) ('x' * 30)
    Check 'cut then trim' (Get-NameFromText (Get-CleanText 'Chapter 1: The Beginning') 11) 'Chapter 1'

    '--- finding files'
    $mixed = New-Folder 'mixed'
    New-TextFile $mixed 'notes.txt' "Shopping list`r`nmilk" | Out-Null
    New-TextFile $mixed 'Photo 10.jpg' 'x' | Out-Null
    New-TextFile $mixed 'Photo 2.jpg' 'x' | Out-Null
    New-TextFile $mixed 'Photo 1.jpg' 'x' | Out-Null
    New-TextFile $mixed '~$notes.docx' 'x' | Out-Null
    New-TextFile $mixed 'empty.txt' '' | Out-Null
    $found = Find-Files $mixed
    Check 'documents only' (($found | ForEach-Object Name) -join ', ') '~$notes.docx, empty.txt, notes.txt'
    $found = Find-Files $mixed -AllFiles
    Check 'all files, Explorer order' (($found | ForEach-Object Name) -join ', ') '~$notes.docx, empty.txt, notes.txt, Photo 1.jpg, Photo 2.jpg, Photo 10.jpg'

    $known = @{}
    $first = Merge-KnownItems (Find-Files $mixed) $known
    $first[2].Text = 'already read'
    $second = Merge-KnownItems (Find-Files $mixed -AllFiles) $known
    Check 'files read before keep their text' (($second | Where-Object { $_.Name -eq 'notes.txt' }).Text) 'already read'
    Check 'files found later are added' ($known.Count) 6

    '--- 1. start with'
    $p = Get-Plan $mixed @{ Source = 'Text' } -AllFiles -Read
    Check 'text: documents named, others kept' (Get-PlanText $p) '~$notes.docx>(keep) | empty.txt>(keep) | notes.txt>Shopping list milk.txt | Photo 1.jpg>(keep) | Photo 2.jpg>(keep) | Photo 10.jpg>(keep)'
    Check 'text: why a photo keeps its name' (($p | Where-Object { $_.Item.Name -eq 'Photo 1.jpg' }).Note) 'not a Word or text document, so it has no text to name it by'
    Check 'text: empty file' (($p | Where-Object { $_.Item.Name -eq 'empty.txt' }).Note) 'the file is empty (0 bytes)'
    $p = Get-Plan $mixed @{ Source = 'Text'; Length = 8 } -Read
    Check 'text: first 8 characters' (($p | Where-Object { $_.Item.Name -eq 'notes.txt' }).NewName) 'Shopping.txt'
    $p = Get-Plan $mixed @{ Source = 'Name' } -AllFiles
    Check 'name now: copies keep the name' (Get-PlanText $p) '~$notes.docx>(keep) | empty.txt>empty.txt | notes.txt>notes.txt | Photo 1.jpg>Photo 1.jpg | Photo 2.jpg>Photo 2.jpg | Photo 10.jpg>Photo 10.jpg'
    $p = Get-Plan $mixed @{ Source = 'Name' } -Mode 'Rename' -AllFiles
    Check 'name now: in place it is the same' ((@($p | Where-Object Action -eq 'Same')).Count) 5
    $p = Get-Plan $mixed @{ Source = 'Fixed'; FixedText = 'Trip' } -AllFiles
    Check 'typed text: repeats get (2)' (($p | Where-Object { $_.Item.Name -like 'Photo*' } | ForEach-Object NewName) -join ', ') 'Trip.jpg, Trip (2).jpg, Trip (3).jpg'
    $p = Get-Plan $mixed @{ Source = 'Fixed'; FixedText = '' } -AllFiles
    Check 'typed text: empty name is kept' (($p | Where-Object { $_.Item.Name -eq 'notes.txt' }).Note) 'the new name would be empty'
    $p = Get-Plan $mixed @{ Source = 'Fixed'; FixedText = 'con' }
    Check 'typed text: Windows-only names' (($p | Where-Object { $_.Item.Name -eq 'empty.txt' }).NewName) '_con.txt'

    '--- 2. change it'
    $photos = New-Folder 'photos'
    foreach ($n in 'IMG_0001.jpg', 'IMG_0002.jpg', "grandma's APPLE pie.txt") { New-TextFile $photos $n 'x' | Out-Null }
    $p = Get-Plan $photos @{ Source = 'Name'; Find = 'img_'; ReplaceWith = 'Photo ' } -AllFiles
    Check 'replace, any capitals' (Get-PlanText $p) "grandma's APPLE pie.txt>grandma's APPLE pie.txt | IMG_0001.jpg>Photo 0001.jpg | IMG_0002.jpg>Photo 0002.jpg"
    $p = Get-Plan $photos @{ Source = 'Name'; Find = 'img_'; ReplaceWith = 'Photo '; MatchCase = $true } -AllFiles
    Check 'replace, matching capitals' (($p | ForEach-Object NewName) -join ', ') "grandma's APPLE pie.txt, IMG_0001.jpg, IMG_0002.jpg"
    $p = Get-Plan $photos @{ Source = 'Name'; Find = '_'; ReplaceWith = 'a/b$1' } -AllFiles
    Check 'replace with odd characters' (($p | Where-Object { $_.Item.Name -eq 'IMG_0001.jpg' }).NewName) 'IMGa-b$10001.jpg'
    $p = Get-Plan $photos @{ Source = 'Name'; RemoveFirst = 4; RemoveLast = 2 } -AllFiles
    Check 'remove first 4 and last 2' (($p | Where-Object { $_.Item.Name -eq 'IMG_0001.jpg' }).NewName) '00.jpg'
    $p = Get-Plan $photos @{ Source = 'Name'; RemoveFirst = 40 } -AllFiles
    Check 'remove everything' (($p | Where-Object { $_.Item.Name -eq 'IMG_0001.jpg' }).Note) 'the new name would be empty'
    $pie = "grandma's APPLE pie"
    Check 'lowercase' (Convert-LetterCase $pie 'Lower') "grandma's apple pie"
    Check 'capitals' (Convert-LetterCase $pie 'Upper') "GRANDMA'S APPLE PIE"
    Check 'each word' (Convert-LetterCase $pie 'Title') "Grandma's Apple Pie"
    Check 'first letter only' (Convert-LetterCase ('- ' + $pie) 'Sentence') "- Grandma's apple pie"

    '--- 3. add to it'
    $p = Get-Plan $photos @{ Source = 'Name'; Prefix = 'Holiday'; Suffix = 'final' } -AllFiles
    Check 'before and after, joined by " - "' (($p | Where-Object { $_.Item.Name -eq 'IMG_0001.jpg' }).NewName) 'Holiday - IMG_0001 - final.jpg'
    $p = Get-Plan $photos @{ Source = 'Name'; Prefix = 'Inv-'; Joiner = '' } -AllFiles
    Check 'nothing between the parts' (($p | Where-Object { $_.Item.Name -eq 'IMG_0001.jpg' }).NewName) 'Inv-IMG_0001.jpg'
    $p = Get-Plan $photos @{ Source = 'Name'; Prefix = '  a/b:c  '; Joiner = '_' } -AllFiles
    Check 'typed text made safe, spaces trimmed' (($p | Where-Object { $_.Item.Name -eq 'IMG_0001.jpg' }).NewName) 'a-b-c_IMG_0001.jpg'
    $p = Get-Plan $photos @{ Source = 'Name'; Suffix = 'x' * 250 } -AllFiles
    $long = $p | Where-Object { $_.Item.Name -eq 'IMG_0001.jpg' }
    Check 'very long names shortened' ($long.NewName.Length) (200 + 4)
    Check 'very long names say so' ($long.Note) 'shortened to 200 characters'

    '--- numbers'
    $p = Get-Plan $mixed @{ Source = 'Name'; Number = $true } -AllFiles
    Check 'numbers at the start, Explorer order' (($p | ForEach-Object NewName) -join ', ') ', 001 - empty.txt, 002 - notes.txt, 003 - Photo 1.jpg, 004 - Photo 2.jpg, 005 - Photo 10.jpg'
    $p = Get-Plan $mixed @{ Source = 'Fixed'; FixedText = 'Trip'; Number = $true; NumberAt = 'End'; NumberFrom = 10; NumberStep = 5; NumberDigits = 2; Joiner = ' ' } -AllFiles
    Check 'numbers at the end, from 10 by 5' (($p | Where-Object { $_.Item.Name -like 'Photo*' } | ForEach-Object NewName) -join ', ') 'Trip 20.jpg, Trip 25.jpg, Trip 30.jpg'
    $byDate = New-Folder 'bydate'
    $d = 0
    foreach ($n in 'c.txt', 'a.txt', 'b.txt') {
        $f = New-TextFile $byDate $n 'x'
        [IO.File]::SetCreationTime($f, (Get-Date '2020-01-01').AddDays($d))
        [IO.File]::SetLastWriteTime($f, (Get-Date '2021-01-01').AddDays(-$d))
        $d++
    }
    $p = Get-Plan $byDate @{ Source = 'Name'; Number = $true; NumberDigits = 1; NumberOrder = 'Created' }
    Check 'numbered by date created' (Get-PlanText $p) 'c.txt>1 - c.txt | a.txt>2 - a.txt | b.txt>3 - b.txt'
    $p = Get-Plan $byDate @{ Source = 'Name'; Number = $true; NumberDigits = 1; NumberOrder = 'Modified' }
    Check 'numbered by date modified' (Get-PlanText $p) 'b.txt>1 - b.txt | a.txt>2 - a.txt | c.txt>3 - c.txt'

    '--- information'
    $p = Get-Plan $byDate @{ Source = 'Name'; Info = 'Created'; InfoAt = 'Start' }
    Check 'date created at the start' (($p | Where-Object { $_.Item.Name -eq 'a.txt' }).NewName) '2020-01-02 - a.txt'
    $p = Get-Plan $byDate @{ Source = 'Name'; Info = 'Modified'; DateFormat = 'MMMM d, yyyy' }
    Check 'date modified at the end, spelled out' (($p | Where-Object { $_.Item.Name -eq 'a.txt' }).NewName) 'a - December 31, 2020.txt'
    $p = Get-Plan $byDate @{ Source = 'Name'; Info = 'Today'; DateFormat = 'yyyyMMdd' }
    Check "today's date" (($p | Where-Object { $_.Item.Name -eq 'a.txt' }).NewName) ('a - ' + (Get-Date).ToString('yyyyMMdd') + '.txt')
    $p = Get-Plan $byDate @{ Source = 'Name'; Info = 'Folder'; Number = $true; Prefix = 'X' }
    Check 'order: before, number, info, name' (($p | Where-Object { $_.Item.Name -eq 'a.txt' }).NewName) 'X - 001 - a - bydate.txt'
    $p = Get-Plan $byDate @{ Source = 'Name'; Info = 'Folder'; InfoAt = 'Start'; Number = $true; NumberAt = 'End'; Suffix = 'Y' }
    Check 'order: info, name, number, after' (($p | Where-Object { $_.Item.Name -eq 'a.txt' }).NewName) 'bydate - a - 001 - Y.txt'

    $info = New-Folder 'info'
    New-Docx (Join-Path $info 'plan.docx') 'Quarterly plan for the shop' 'Ann Smith' 'Shop Plan: 2027'
    [IO.File]::WriteAllText((Join-Path $info 'letter.rtf'), ('{\rtf1\ansi{\info{\title Caf\''e9 menu}{\author Jos\''e9 Lee}}\f0 Dear Raj,\par}'), [Text.Encoding]::ASCII)
    New-TextFile $info 'plain.txt' 'Nothing about who wrote it' | Out-Null
    New-Photo (Join-Path $info 'beach.jpg') '2019:07:04 15:30:00'
    $p = Get-Plan $info @{ Source = 'Name'; Info = 'Author' } -AllFiles
    Check 'author from Word and Rich Text' (Get-PlanText $p) ('beach.jpg>beach.jpg | letter.rtf>letter - Jos' + [char]0xE9 + ' Lee.rtf | plain.txt>plain.txt | plan.docx>plan - Ann Smith.docx')
    Check 'missing author says so' (($p | Where-Object { $_.Item.Name -eq 'plain.txt' }).Note) 'no author found in this file, so none was added'
    $p = Get-Plan $info @{ Source = 'Name'; Info = 'Title'; InfoAt = 'Start' } -AllFiles
    Check 'title, made safe' (($p | Where-Object { $_.Item.Name -eq 'plan.docx' }).NewName) 'Shop Plan - 2027 - plan.docx'
    Check 'title from Rich Text' (($p | Where-Object { $_.Item.Name -eq 'letter.rtf' }).NewName) ('Caf' + [char]0xE9 + ' menu - letter.rtf')
    $p = Get-Plan $info @{ Source = 'Name'; Info = 'Taken'; DateFormat = 'yyyy-MM-dd HH.mm' } -AllFiles
    Check 'date a photo was taken' (($p | Where-Object { $_.Item.Name -eq 'beach.jpg' }).NewName) 'beach - 2019-07-04 15.30.jpg'
    Check 'no date taken says so' (($p | Where-Object { $_.Item.Name -eq 'plain.txt' }).Note) 'no date taken found in this file, so none was added'

    '--- renaming in place with the new choices, and putting the names back'
    $p = Get-Plan $photos @{ Source = 'Fixed'; FixedText = 'Beach'; Number = $true; NumberAt = 'End'; NumberDigits = 2 } -Mode 'Rename' -AllFiles
    $res = Invoke-Renaming -Plan $p -Mode 'Rename' -Folder $photos
    Check 'renamed' ((Get-ChildItem $photos -File | ForEach-Object Name | Sort-Object) -join ', ') 'Beach - 01.txt, Beach - 02.jpg, Beach - 03.jpg'
    $p2 = Get-Plan $photos @{ Source = 'Fixed'; FixedText = 'Beach'; Number = $true; NumberAt = 'End'; NumberDigits = 2 } -Mode 'Rename' -AllFiles
    Check 'running it again changes nothing' ((@($p2 | Where-Object Action -eq 'Same')).Count) 3
    $undo = Undo-LastRenaming $photos
    Check 'old names back' ((Get-ChildItem $photos -File | ForEach-Object Name | Sort-Object) -join ', ') "grandma's APPLE pie.txt, IMG_0001.jpg, IMG_0002.jpg"

    '--- names that move round: renumbering, swapping, running again'
    function Get-Contents($Folder) { (@(Get-ChildItem $Folder -File | Sort-Object Name | ForEach-Object { $_.Name + '=' + [IO.File]::ReadAllText($_.FullName) }) -join ', ') }
    $renum = New-Folder 'renumber'
    foreach ($n in 'Trip 1.jpg', 'Trip 2.jpg', 'Trip 3.jpg') { New-TextFile $renum $n $n | Out-Null }
    $shift = @{ Source = 'Fixed'; FixedText = 'Trip'; Number = $true; NumberDigits = 1; NumberFrom = 2; NumberAt = 'End'; Joiner = ' ' }
    $p = Get-Plan $renum $shift -Mode 'Rename' -AllFiles
    Check 'renumber: no (2) names' (Get-PlanText $p) 'Trip 1.jpg>Trip 2.jpg | Trip 2.jpg>Trip 3.jpg | Trip 3.jpg>Trip 4.jpg'
    $res = Invoke-Renaming -Plan $p -Mode 'Rename' -Folder $renum
    Check 'renumber: each file moved up one' (Get-Contents $renum) 'Trip 2.jpg=Trip 1.jpg, Trip 3.jpg=Trip 2.jpg, Trip 4.jpg=Trip 3.jpg'
    Check 'renumber: no temporary names left' (@(Get-ChildItem $renum -File -Filter '*.renaming*').Count) 0
    $undo = Undo-LastRenaming $renum
    Check 'renumber: put back' (Get-Contents $renum) 'Trip 1.jpg=Trip 1.jpg, Trip 2.jpg=Trip 2.jpg, Trip 3.jpg=Trip 3.jpg'

    $swap = New-Folder 'swap'
    New-TextFile $swap 'a.txt' 'b' | Out-Null
    New-TextFile $swap 'b.txt' 'a' | Out-Null
    $p = Get-Plan $swap @{ Source = 'Text' } -Mode 'Rename' -Read
    Check 'swap: plan' (Get-PlanText $p) 'a.txt>b.txt | b.txt>a.txt'
    $res = Invoke-Renaming -Plan $p -Mode 'Rename' -Folder $swap
    Check 'swap: names swapped' (Get-Contents $swap) 'a.txt=a, b.txt=b'
    $undo = Undo-LastRenaming $swap
    Check 'swap: put back' (Get-Contents $swap) 'a.txt=b, b.txt=a'

    $again = New-Folder 'again'
    New-TextFile $again 'Same start.txt' 'Same start' | Out-Null
    New-TextFile $again 'Same start (2).txt' 'Same start' | Out-Null
    $p = Get-Plan $again @{ Source = 'Text' } -Mode 'Rename' -Read
    Check 'running again keeps "(2)" names' ((@($p | ForEach-Object Action)) -join ', ') 'Same, Same'

    $chain = New-Folder 'chain'
    foreach ($n in '1.txt', '2.txt', '3.txt') { New-TextFile $chain $n $n | Out-Null }
    $up = @{ Source = 'Fixed'; FixedText = ''; Number = $true; NumberDigits = 1; NumberFrom = 2 }
    $p = Get-Plan $chain $up -Mode 'Rename'
    $lock = [IO.File]::Open((Join-Path $chain '3.txt'), 'Open', 'Read', 'Read')
    try { $res = Invoke-Renaming -Plan $p -Mode 'Rename' -Folder $chain } finally { $lock.Dispose() }
    Check 'locked file: nothing moved' (Get-Contents $chain) '1.txt=1.txt, 2.txt=2.txt, 3.txt=3.txt'
    Check 'locked file: why' (($res | Where-Object { $_.Name -eq '3.txt' }).Detail) 'the file is open in another program (such as Word). Close it and try again.'
    Check 'locked file: the ones it blocked' (($res | Where-Object { $_.Name -eq '1.txt' }).Detail) 'its new name still belongs to a file that could not be renamed.'
    Check 'locked file: nothing to put back' ($null -eq (Get-LastRenaming $chain)) $true
    $res = Invoke-Renaming -Plan (Get-Plan $chain $up -Mode 'Rename') -Mode 'Rename' -Folder $chain
    Check 'unlocked: all moved' (Get-Contents $chain) '2.txt=1.txt, 3.txt=2.txt, 4.txt=3.txt'

    $stuck = New-Folder 'stuck'
    New-TextFile $stuck 'x.txt.renaming' 'x' | Out-Null
    Add-HistoryRecord ([pscustomobject]@{ Id = 'stuck'; Folder = (Get-FolderKey $stuck); When = (Get-Date).ToString('s')
        Entries = @([pscustomobject]@{ Old = 'x.txt'; New = 'y.txt'; Temp = 'x.txt.renaming' }) })
    $undo = Undo-LastRenaming $stuck
    Check 'stopped half way: old name back from the temporary name' (Get-Contents $stuck) 'x.txt=x'

    $one = New-Folder 'one'
    New-TextFile $one 'shopping.txt' 'Shopping' | Out-Null
    $p = Get-Plan $one @{ Source = 'Text' } -Mode 'Rename' -Read
    $res = Invoke-Renaming -Plan $p -Mode 'Rename' -Folder $one
    Check 'one file, capital letter only' ((Get-ChildItem $one -File).Name) 'Shopping.txt'
    $undo = Undo-LastRenaming $one
    Check 'one file, put back' ((Get-ChildItem $one -File).Name) 'shopping.txt'

    '--- odds and ends'
    $U = '\' + 'u'
    Check 'Rich Text: code with a stand-in letter' (ConvertFrom-RtfText ('Caf' + $U + '233e menu')) ('Caf' + [char]0xE9 + ' menu')
    Check "Rich Text: code with a stand-in \'e9" (ConvertFrom-RtfText ('Caf' + $U + "233\'e9 menu")) ('Caf' + [char]0xE9 + ' menu')
    Check 'Rich Text: code with a stand-in ?' (ConvertFrom-RtfText ('Caf' + $U + '233? menu')) ('Caf' + [char]0xE9 + ' menu')
    Check "Rich Text: \'e9 codes" (ConvertFrom-RtfText "\'e9t\'e9") ([string][char]0xE9 + 't' + [char]0xE9)
    $items = Find-Files $info -AllFiles
    $p = Get-RenamePlan -Items $items -Options @{ Source = 'Name'; Info = 'Author' } -Mode 'Copy' -ExistingNames @() -Folder $info
    Check 'author read without being asked first' (($p | Where-Object { $_.Item.Name -eq 'plan.docx' }).NewName) 'plan - Ann Smith.docx'
    $many = New-Folder 'many'
    1..1000 | ForEach-Object { [IO.File]::WriteAllText((Join-Path $many "IMG_$_.jpg"), 'x') }
    $items = Find-Files $many -AllFiles
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $p = Get-RenamePlan -Items $items -Options @{ Source = 'Name'; Number = $true; Prefix = 'Trip' } -Mode 'Copy' -ExistingNames @() -Folder $many
    "info  1000 files: the list takes $($clock.ElapsedMilliseconds) ms to work out"
    Check '1000 files: last one' ($p[999].NewName) 'Trip - 1000 - IMG_1000.jpg'
} finally {
    try { [IO.Directory]::Delete($Root, $true) } catch {}
}
''
if ($script:Failures) { "$($script:Failures) check(s) FAILED"; exit 1 } else { 'All checks passed.' }
