Option Explicit
Dim excel, wb, vbProj, module, btnSheet, btn
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1

Set wb = excel.Workbooks.Open("C:\Users\o_antonov\NEW\RN\rn_card.xlsx")
Set vbProj = wb.VBProject

On Error Resume Next
Dim comp
For Each comp In vbProj.VBComponents
    If comp.Name = "modReconcile" Then vbProj.VBComponents.Remove comp
Next
On Error GoTo 0

Set module = vbProj.VBComponents.Import("C:\Users\o_antonov\NEW\RN\modReconcile.bas")

Set btnSheet = wb.Worksheets(wb.Worksheets.Count)
Set btn = btnSheet.Shapes.AddShape(1, 10, 30, 220, 40)
btn.Name = "btnReconcile"
btn.OnAction = "SformirovatOtchety"
btn.TextFrame.Characters.Text = "Run Reports"
btn.Fill.ForeColor.RGB = 9851952
btn.TextFrame.Characters.Font.Color = 16777215
btn.TextFrame.Characters.Font.Bold = True

On Error Resume Next
CreateObject("Scripting.FileSystemObject").DeleteFile "C:\Users\o_antonov\NEW\RN\rn_card.xlsm", True
On Error GoTo 0
wb.SaveAs "C:\Users\o_antonov\NEW\RN\rn_card.xlsm", 52
wb.Close False
excel.Quit

WScript.Echo "OK: rn_card.xlsm created"
