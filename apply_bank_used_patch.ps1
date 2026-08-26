$ErrorActionPreference = 'Stop'
$root = 'C:\Users\o_antonov\NEW\RN'
$basFile = Join-Path $root 'modReconcile.bas'
$patchDir = Join-Path $root 'patches'
$cp1251 = [System.Text.Encoding]::GetEncoding(1251)
$utf8 = New-Object System.Text.UTF8Encoding $false

function Read-Utf8($name) {
    return [System.IO.File]::ReadAllText((Join-Path $patchDir $name), $utf8).TrimEnd()
}

$content = $cp1251.GetString([System.IO.File]::ReadAllBytes($basFile))
if ($content -match 'FindUsedBankMatch') { Write-Host 'Already patched'; exit 0 }

$rnNew = Read-Utf8 'patch_rn_else.txt'
$content = [regex]::Replace(
    $content,
    '(?ms)        ElseIf status <> "[^"]+" Then\r?\n            note = "[^"]*"\r?\n        End If\r?\n\r?\n        wsRpt\.Cells\(outRow, 1\)\.Value',
    ($rnNew + "`r`n`r`n        wsRpt.Cells(outRow, 1).Value"),
    1
)

$bankNew = Read-Utf8 'patch_bank_else.txt'
$content = [regex]::Replace(
    $content,
    '(?ms)(            Set rnMatch = FindRNMatch\(rnData, sumVal, inn, kpp, contractId, CDate\(dt\)\)\r?\n            If Not rnMatch Is Nothing Then\r?\n                matchStatus = "[^"]+"\r?\n                rnTransId = rnMatch\("TransId"\)\r?\n                rnMatch\("Used"\) = True\r?\n            )Else\r?\n                matchStatus = "[^"]+"\r?\n                note = "[^"]*"\r?\n            End If',
    ('$1' + $bankNew),
    1
)

$findRnHelper = @'
Private Function ItemMatchesBankRN(item As Object, sumVal As Double, inn As String, kpp As String, contractId As String, bankDate As Date) As Boolean
    ItemMatchesBankRN = False
    If Abs(item("Sum") - sumVal) > 0.01 Then Exit Function
    If inn <> "" And item("INN") <> "" And inn <> item("INN") Then Exit Function
    If kpp <> "" And item("KPP") <> "" And kpp <> item("KPP") Then Exit Function
    If contractId <> "" And item("ContractId") <> "" And contractId <> item("ContractId") Then Exit Function
    If CDate(item("Date")) < bankDate Then Exit Function
    If CDate(item("Date")) > bankDate + 7 Then Exit Function
    ItemMatchesBankRN = True
End Function

Private Function FindRNMatch(rnData As Collection, sumVal As Double, inn As String, kpp As String, contractId As String, bankDate As Date) As Variant
    Dim i As Long, item As Object
    For i = 1 To rnData.Count
        Set item = rnData(i)
        If item("Used") Then GoTo NextI
        If ItemMatchesBankRN(item, sumVal, inn, kpp, contractId, bankDate) Then
            Set FindRNMatch = item
            Exit Function
        End If
NextI:
    Next i
    Set FindRNMatch = Nothing
End Function
'@

$findRnNew = Read-Utf8 'patch_find_used_rn.txt'
if ($content -notmatch 'ItemMatchesBankRN') {
    $content = [regex]::Replace(
        $content,
        '(?ms)Private Function FindRNMatch\(rnData As Collection, sumVal As Double, inn As String, kpp As String, contractId As String, bankDate As Date\) As Variant\r?\n    Dim i As Long, item As Object\r?\n    For i = 1 To rnData\.Count\r?\n        Set item = rnData\(i\)\r?\n        If item\("Used"\) Then GoTo NextI\r?\n        If Abs\(item\("Sum"\) - sumVal\) > 0\.01 Then GoTo NextI\r?\n        If inn <> "" And item\("INN"\) <> "" And inn <> item\("INN"\) Then GoTo NextI\r?\n        If kpp <> "" And item\("KPP"\) <> "" And kpp <> item\("KPP"\) Then GoTo NextI\r?\n        If contractId <> "" And item\("ContractId"\) <> "" And contractId <> item\("ContractId"\) Then GoTo NextI\r?\n        If CDate\(item\("Date"\)\) < bankDate Then GoTo NextI\r?\n        If CDate\(item\("Date"\)\) > bankDate \+ 7 Then GoTo NextI\r?\n        Set FindRNMatch = item\r?\n        Exit Function\r?\nNextI:\r?\n    Next i\r?\n    Set FindRNMatch = Nothing\r?\nEnd Function',
        ($findRnHelper + "`r`n" + $findRnNew),
        1
    )
}

$findBankHelper = @'
Private Function ItemMatchesRNBank(item As Object, sumVal As Double, inn As String, kpp As String, contractId As String, rnDate As Date) As Boolean
    ItemMatchesRNBank = False
    If Abs(item("Sum") - sumVal) > 0.01 Then Exit Function
    If inn <> "" And item("INN") <> "" And inn <> item("INN") Then Exit Function
    If kpp <> "" And item("KPP") <> "" And kpp <> item("KPP") Then Exit Function
    If contractId <> "" And item("ContractId") <> "" And contractId <> item("ContractId") Then Exit Function
    If CDate(item("Date")) > rnDate Then Exit Function
    If CDate(item("Date")) < rnDate - 7 Then Exit Function
    ItemMatchesRNBank = True
End Function

Private Function FindBankMatch(bankData As Collection, sumVal As Double, inn As String, kpp As String, contractId As String, rnDate As Date) As Variant
    Dim i As Long, item As Object
    For i = 1 To bankData.Count
        Set item = bankData(i)
        If item("Used") Then GoTo NextBI
        If ItemMatchesRNBank(item, sumVal, inn, kpp, contractId, rnDate) Then
            Set FindBankMatch = item
            Exit Function
        End If
NextBI:
    Next i
    Set FindBankMatch = Nothing
End Function
'@

$findBankNew = Read-Utf8 'patch_find_used_bank.txt'
if ($content -notmatch 'ItemMatchesRNBank') {
    $content = [regex]::Replace(
        $content,
        '(?ms)Private Function FindBankMatch\(bankData As Collection, sumVal As Double, inn As String, kpp As String, contractId As String, rnDate As Date\) As Variant\r?\n    Dim i As Long, item As Object\r?\n    For i = 1 To bankData\.Count\r?\n        Set item = bankData\(i\)\r?\n        If item\("Used"\) Then GoTo NextBI\r?\n        If Abs\(item\("Sum"\) - sumVal\) > 0\.01 Then GoTo NextBI\r?\n        If inn <> "" And item\("INN"\) <> "" And inn <> item\("INN"\) Then GoTo NextBI\r?\n        If kpp <> "" And item\("KPP"\) <> "" And kpp <> item\("KPP"\) Then GoTo NextBI\r?\n        If contractId <> "" And item\("ContractId"\) <> "" And contractId <> item\("ContractId"\) Then GoTo NextBI\r?\n        If CDate\(item\("Date"\)\) > rnDate Then GoTo NextBI\r?\n        If CDate\(item\("Date"\)\) < rnDate - 7 Then GoTo NextBI\r?\n        Set FindBankMatch = item\r?\n        Exit Function\r?\nNextBI:\r?\n    Next i\r?\n    Set FindBankMatch = Nothing\r?\nEnd Function',
        ($findBankHelper + "`r`n" + $findBankNew),
        1
    )
}

$colorRn = Read-Utf8 'patch_color_rn_line.txt'
$content = [regex]::Replace(
    $content,
    '(?ms)(Private Sub ColorRNRow\(ws As Worksheet, rowNum As Long, status As String\)\r?\n    Dim clr As Long\r?\n    Select Case status\r?\n        Case "[^"]+": clr = CLR_OK\r?\n        Case "[^"]+": clr = CLR_ERR\r?\n        Case "[^"]+": clr = CLR_WARN\r?\n)(        Case Else: clr = -1)',
    ('$1' + $colorRn + "`r`n" + '$2'),
    1
)

$colorBank = Read-Utf8 'patch_color_bank_line.txt'
$content = [regex]::Replace(
    $content,
    '(?ms)(Private Sub ColorBankRow\(ws As Worksheet, rowNum As Long, status As String, isGSM As Boolean\)\r?\n    Dim clr As Long\r?\n    Select Case status\r?\n        Case "[^"]+": clr = CLR_OK\r?\n        Case "[^"]+": clr = CLR_ERR\r?\n        Case "[^"]+": clr = CLR_WARN\r?\n)(        Case Else)',
    ('$1' + $colorBank + "`r`n" + '$2'),
    1
)

$sumRn = Read-Utf8 'patch_summary_rn_line.txt'
$content = [regex]::Replace(
    $content,
    '(?ms)(        ElseIf key = "[^"]+" Then\r?\n            wsSum\.Range\(wsSum\.Cells\(r, 1\), wsSum\.Cells\(r, 3\)\)\.Interior\.Color = CLR_WARN\r?\n)(        End If\r?\n        r = r \+ 1\r?\n    Next key\r?\n\r?\n    r = r \+ 3)',
    ('$1' + $sumRn + "`r`n" + '$2'),
    1
)

$sumBank = Read-Utf8 'patch_summary_bank_line.txt'
$content = [regex]::Replace(
    $content,
    '(?ms)(        ElseIf key = "[^"]+" Then\r?\n            wsSum\.Range\(wsSum\.Cells\(r, 1\), wsSum\.Cells\(r, 3\)\)\.Interior\.Color = CLR_WARN\r?\n)(        End If\r?\n        r = r \+ 1\r?\n    Next key\r?\n\r?\n    r = r \+ 2)',
    ('$1' + $sumBank + "`r`n" + '$2'),
    1
)

if ($content -notmatch 'FindUsedBankMatch') { throw 'Patch failed' }

$content = $content -replace "`r?`n", "`r`n"
[System.IO.File]::WriteAllBytes($basFile, $cp1251.GetBytes($content))
Write-Host 'Patch OK'
