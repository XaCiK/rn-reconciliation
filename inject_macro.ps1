# Inject VBA macro into rn_card.xlsx and save as rn_card.xlsm
$srcPath = "C:\Users\o_antonov\NEW\RN\rn_card.xlsx"
$dstPath = "C:\Users\o_antonov\NEW\RN\rn_card.xlsm"
$basPath = "C:\Users\o_antonov\NEW\RN\modReconcile.bas"

# Enable VBA project access for current user (Excel)
$secPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Security",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Security"
)
foreach ($sp in $secPaths) {
    if (Test-Path $sp) {
        Set-ItemProperty -Path $sp -Name "AccessVBOM" -Value 1 -Type DWord -Force
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1  # msoAutomationSecurityLow

try {
    $wb = $excel.Workbooks.Open($srcPath)

    # Remove existing module if present
    $vbProj = $wb.VBProject
    foreach ($comp in @($vbProj.VBComponents)) {
        if ($comp.Name -eq "modReconcile") {
            $vbProj.VBComponents.Remove($comp)
        }
    }

    $module = $vbProj.VBComponents.Import($basPath)
    if ($module.Name -ne "modReconcile") { $module.Name = "modReconcile" }

    # Add launch button on last sheet (link template)
    $btnSheet = $wb.Worksheets.Item($wb.Worksheets.Count)

    # Remove old buttons
    foreach ($sh in @($btnSheet.Shapes)) {
        if ($sh.Type -eq 8) { $sh.Delete() }
    }

    $btn = $btnSheet.Shapes.AddShape(1, 10, 30, 220, 40)
    $btn.Name = "btnReconcile"
    $btn.TextFrame.Characters().Text = [char]0x0421 + [char]0x0444 + [char]0x043e + [char]0x0440 + [char]0x043c + [char]0x0438 + [char]0x0440 + [char]0x043e + [char]0x0432 + [char]0x0430 + [char]0x0442 + [char]0x044c + " " + [char]0x043e + [char]0x0442 + [char]0x0447 + [char]0x0451 + [char]0x0442 + [char]0x044b
    $btn.TextFrame.Characters().Font.Size = 12
    $btn.TextFrame.Characters().Font.Bold = $true
    $btn.TextFrame.Characters().Font.Color = 16777215
    $btn.Fill.ForeColor.RGB = 9851952
    $btn.OnAction = "SformirovatOtchety"

    # Save as xlsm (macro-enabled)
    if (Test-Path $dstPath) { Remove-Item $dstPath -Force }
    $wb.SaveAs($dstPath, 52) # xlOpenXMLWorkbookMacroEnabled
    Write-Host "OK: Saved to $dstPath"
    $wb.Close($false)
}
catch {
    Write-Host "VBA injection failed: $($_.Exception.Message)"
    Write-Host "Manual import: Alt+F11 -> File -> Import -> modReconcile.bas"
    Write-Host "Then save as .xlsm"
    if ($wb) { $wb.Close($false) }
    exit 1
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
