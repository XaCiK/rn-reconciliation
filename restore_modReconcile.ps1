$ErrorActionPreference = 'Stop'
$transcript = 'C:\Users\o_antonov\.cursor\projects\c-Users-o-antonov-NEW-RN\agent-transcripts\2228858f-2fa8-414c-a236-a9626eab9148\subagents\a01a69d8-a862-4085-a93a-9c8cb0723cbf.jsonl'
$outFile = 'C:\Users\o_antonov\NEW\RN\modReconcile.bas'
$cp1251 = [System.Text.Encoding]::GetEncoding(1251)

$line = (Get-Content -Path $transcript -Encoding UTF8)[2]
if ($line -notmatch '"contents":"') { throw 'contents field not found' }

$start = $line.IndexOf('"contents":"') + 12
$sb = New-Object System.Text.StringBuilder
$i = $start
while ($i -lt $line.Length) {
    $ch = $line[$i]
    if ($ch -eq '\') {
        $i++
        if ($i -ge $line.Length) { break }
        $esc = $line[$i]
        switch ($esc) {
            'n' { [void]$sb.Append("`n") }
            'r' { [void]$sb.Append("`r") }
            't' { [void]$sb.Append("`t") }
            '"' { [void]$sb.Append('"') }
            '\' { [void]$sb.Append('\') }
            '/' { [void]$sb.Append('/') }
            'u' {
                $hex = $line.Substring($i + 1, 4)
                [void]$sb.Append([char][int]('0x' + $hex))
                $i += 4
            }
            default { [void]$sb.Append($esc) }
        }
    }
    elseif ($ch -eq '"') {
        break
    }
    else {
        [void]$sb.Append($ch)
    }
    $i++
}

$content = $sb.ToString()
if ($content.Length -lt 1000) { throw "Extracted content too short: $($content.Length)" }

$freezeOldBank = "    wsRpt.Columns(`"A:P`").AutoFit`n    wsRpt.Rows(`"1:1`").RowHeight = 28`n    wsRpt.Range(`"A1`").Select`n    ActiveWindow.FreezePanes = False`n    wsRpt.Range(`"A2`").Select`n    ActiveWindow.FreezePanes = True`n    wsRpt.AutoFilterMode = False"
$freezeNewBank = "    wsRpt.Columns(`"A:P`").AutoFit`n    wsRpt.Rows(`"1:1`").RowHeight = 28`n    ApplyFreezeTopRow wsRpt`n    wsRpt.AutoFilterMode = False"
$freezeOldRn = "    wsRpt.Columns(`"A:M`").AutoFit`n    wsRpt.Columns(`"E`").ColumnWidth = 35`n    wsRpt.Rows(`"1:1`").RowHeight = 28`n    wsRpt.Range(`"A2`").Select`n    ActiveWindow.FreezePanes = True`n    wsRpt.AutoFilterMode = False"
$freezeNewRn = "    wsRpt.Columns(`"A:M`").AutoFit`n    wsRpt.Columns(`"E`").ColumnWidth = 35`n    wsRpt.Rows(`"1:1`").RowHeight = 28`n    ApplyFreezeTopRow wsRpt`n    wsRpt.AutoFilterMode = False"

$content = $content.Replace($freezeOldBank, $freezeNewBank)
$content = $content.Replace($freezeOldRn, $freezeNewRn)

# Light gray headers, black text; no blue row fills
$content = $content.Replace('Private Const CLR_HEADER As Long = 9851952', 'Private Const CLR_HEADER As Long = 14277081')
$content = $content.Replace('Private Const CLR_HEADER_FG As Long = 16777215', 'Private Const CLR_HEADER_FG As Long = 0')
$content = $content.Replace('Private Const CLR_ERR As Long = 14483456', 'Private Const CLR_ERR As Long = 16764057')
$content = $content.Replace('Private Const CLR_GSM As Long = 13434828', 'Private Const CLR_GSM As Long = 13561798')
$content = $content.Replace('            If isGSM Then clr = CLR_ALT Else clr = -1', '            clr = -1')

# Fix dictionary: nefteprodukt -> GSM not card purchase; refresh if sheet corrupted
$dictRows = (Get-Content -Path (Join-Path $PSScriptRoot 'dict_rows_utf8.txt') -Raw -Encoding UTF8).TrimEnd()
$content = [regex]::Replace(
    $content,
    '(?ms)(    r = 2\r?\n)(.*?)(\r?\n    ws\.Columns\("A:G"\)\.AutoFit)',
    { param($m) $m.Groups[1].Value + $dictRows + "`n" + $m.Groups[3].Value }
)

$oldEnsure = "    ElseIf ws.Cells(2, 1).Value = `"`" Or ws.Cells(ws.Rows.Count, 1).End(xlUp).Row < 40 Then`n        PopulateDefaultDictionary ws`n    End If`n`n    Set EnsureDictionarySheet = ws`nEnd Function"
$newEnsure = @'
    ElseIf ws.Cells(2, 1).Value = "" Or ws.Cells(ws.Rows.Count, 1).End(xlUp).Row < 40 Or DictionaryNeedsRefresh(ws) Then
        PopulateDefaultDictionary ws
    End If

    Set EnsureDictionarySheet = ws
End Function

Private Function DictionaryNeedsRefresh(ws As Worksheet) As Boolean
    Dim hdr As String
    hdr = Trim$(CStr(ws.Cells(1, 1).Value))
    If Len(hdr) = 0 Then DictionaryNeedsRefresh = True: Exit Function
    If InStr(hdr, "?") > 0 Then DictionaryNeedsRefresh = True: Exit Function
    DictionaryNeedsRefresh = False
End Function
'@
$content = $content.Replace($oldEnsure, $newEnsure)

if ($content -notmatch 'Private Sub ApplyFreezeTopRow') {
    $insert = @"

Private Sub ApplyFreezeTopRow(ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 1
    ActiveWindow.SplitColumn = 0
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Sub ColorBankRow(ws As Worksheet, rowNum As Long, status As String, isGSM As Boolean)
"@
    $content = $content.Replace(
        'Private Sub ColorBankRow(ws As Worksheet, rowNum As Long, status As String, isGSM As Boolean)',
        $insert
    )
}

$content = $content -replace "`r?`n", "`r`n"
[System.IO.File]::WriteAllBytes($outFile, $cp1251.GetBytes($content))

$bytes = [System.IO.File]::ReadAllBytes($outFile)
Write-Host "Bytes: $($bytes.Length)"
Write-Host "First3: $($bytes[0]),$($bytes[1]),$($bytes[2])"

$patterns = @{
    Slovar = @(0xD1,0xEB,0xEE,0xE2,0xE0,0xF0,0xFC)
    Nefteprodukt = @(0xED,0xE5,0xF4,0xF2,0xE5,0xEF,0xF0,0xEE,0xE4,0xF3,0xEA,0xF2)
    Zachisleno = @(0xC7,0xE0,0xF7,0xE8,0xF1,0xEB,0xE5,0xED,0xEE)
    ApplyFreeze = @(0x41,0x70,0x70,0x6C,0x79,0x46,0x72,0x65,0x65,0x7A,0x65,0x54,0x6F,0x70,0x52,0x6F,0x77)
}

function Test-Bytes($haystack, $needle) {
    for ($i = 0; $i -le $haystack.Length - $needle.Length; $i++) {
        $ok = $true
        for ($j = 0; $j -lt $needle.Length; $j++) {
            if ($haystack[$i + $j] -ne $needle[$j]) { $ok = $false; break }
        }
        if ($ok) { return $true }
    }
    return $false
}

foreach ($key in $patterns.Keys) {
    if (Test-Bytes $bytes $patterns[$key]) { Write-Host "OK: $key" } else { Write-Host "FAIL: $key" }
}
