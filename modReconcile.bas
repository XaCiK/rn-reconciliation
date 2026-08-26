Attribute VB_Name = "modReconcile"
Option Explicit

' ============================================================================
'  RN-Card Reconciliation Macro
'  ����: ����� + ���� + ��� + ��� + ID ��������
'  �����������: https://lkp.rn-card.ru/Contract/?client=ID&cid=ID+1&act=open#info
' ============================================================================

Private Type ClientInfo
    ID As Long
    INN As String
    KPP As String
    ShortName As String
    FullName As String
    ContractCount As Long
End Type

Private Type TagInfo
    Tag As String
    Summarize As Boolean
    MatchedKeyword As String
End Type

Private Const CLR_HEADER As Long = 14277081
Private Const CLR_HEADER_FG As Long = 0
Private Const CLR_OK As Long = 13561798
Private Const CLR_WARN As Long = 10092543
Private Const CLR_ERR As Long = 16764057
Private Const CLR_ALT As Long = 15921906
Private Const CLR_GSM As Long = 13561798

Private Const SHEET_SLOVAR As String = "RN_Slovar"
Private Const SHEET_OTCHET_BANK As String = "RN_Otchet_Bank"
Private Const SHEET_OTCHET_RN As String = "RN_Otchet_RN"
Private Const SHEET_ITOG As String = "RN_Itog"
Private Const SHEET_NAME As String = "Name"

Private Const BK_ID As Long = 1
Private Const BK_DATE As Long = 2
Private Const BK_SUM As Long = 4
Private Const BK_INN As Long = 5
Private Const BK_CPNAME As Long = 6
Private Const BK_PURPOSE As Long = 7

Private Const RN_DATE As Long = 1
Private Const RN_OPTYPE As Long = 3
Private Const RN_SUM As Long = 4
Private Const RN_CLIENT As Long = 5
Private Const RN_CONTRACT As Long = 7
Private Const RN_CONTRACTID As Long = 8
Private Const RN_TRANSID As Long = 9

Private Const NM_ID As Long = 1
Private Const NM_SHORTNAME As Long = 3
Private Const NM_FULLNAME As Long = 4
Private Const NM_INN As Long = 6
Private Const NM_KPP As Long = 7

Private mClients() As ClientInfo
Private mClientCount As Long
Private mDictKeywords As Collection
Private mDictTags As Collection
Private mDictSummarize As Collection
Private mDictPriority As Collection

' ============================================================================
Public Sub ObnovitSlovar()
    Dim ws As Worksheet
    Set ws = EnsureDictionarySheet
    PopulateDefaultDictionary ws
    MsgBox "������� ������� �� ����� " & SHEET_SLOVAR & ".", vbInformation
End Sub

Public Sub SformirovatOtchety()
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean

    On Error GoTo ErrHandler

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    LoadClients
    LoadDictionary

    BuildBankReport
    BuildRNReport
    BuildSummary
    RefreshDictionaryStats

    Application.ScreenUpdating = True
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevScreen

    MsgBox "������ ������������:" & vbCrLf & "  - " & SHEET_OTCHET_BANK & vbCrLf & "  - " & SHEET_OTCHET_RN & vbCrLf & "  - " & SHEET_ITOG & vbCrLf & vbCrLf & "������� �����: " & SHEET_SLOVAR, vbInformation, "RN-Card ������"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevScreen
    MsgBox "������: " & Err.Description, vbCritical, "RN-Card ������"
End Sub

' ============================================================================
Private Sub LoadClients()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, n As Long

    Set ws = GetSheetByName(SHEET_NAME)
    If ws Is Nothing Then Err.Raise vbObjectError + 1, , "���� " & SHEET_NAME & " �� ������"

    lastRow = ws.Cells(ws.Rows.Count, NM_ID).End(xlUp).Row
    n = lastRow - 1
    If n < 1 Then Exit Sub

    ReDim mClients(1 To n)
    mClientCount = n

    For i = 2 To lastRow
        With mClients(i - 1)
            .ID = CLng(Val(ws.Cells(i, NM_ID).Value))
            .ShortName = Trim$(CStr(ws.Cells(i, NM_SHORTNAME).Value))
            .FullName = Trim$(CStr(ws.Cells(i, NM_FULLNAME).Value))
            .INN = NormalizeINN(CStr(ws.Cells(i, NM_INN).Value))
            .KPP = NormalizeKPP(CStr(ws.Cells(i, NM_KPP).Value))
            .ContractCount = CLng(Val(ws.Cells(i, 8).Value))
        End With
    Next i
End Sub

' ============================================================================
Private Sub LoadDictionary()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set mDictKeywords = New Collection
    Set mDictTags = New Collection
    Set mDictSummarize = New Collection
    Set mDictPriority = New Collection

    Set ws = EnsureDictionarySheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        Dim kw As String, tg As String, sumFlag As String
        kw = LCase$(Trim$(CStr(ws.Cells(i, 1).Value)))
        tg = Trim$(CStr(ws.Cells(i, 2).Value))
        sumFlag = UCase$(Trim$(CStr(ws.Cells(i, 4).Value)))
        If Len(kw) > 0 And Len(tg) > 0 Then
            mDictKeywords.Add kw, kw & CStr(i)
            mDictTags.Add tg, "t" & CStr(i)
            mDictSummarize.Add (sumFlag = "��" Or sumFlag = "YES" Or sumFlag = "1"), "s" & CStr(i)
            mDictPriority.Add CLng(Val(ws.Cells(i, 5).Value)), "p" & CStr(i)
        End If
    Next i
End Sub

' ============================================================================
Private Function EnsureDictionarySheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SLOVAR)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_SLOVAR
        PopulateDefaultDictionary ws
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

' ============================================================================
Private Sub PopulateDefaultDictionary(ws As Worksheet)
    Dim r As Long

    ws.Cells.Clear
    ws.Cells(1, 1).Value = "�������� �����"
    ws.Cells(1, 2).Value = "���"
    ws.Cells(1, 3).Value = "��������"
    ws.Cells(1, 4).Value = "�����������"
    ws.Cells(1, 5).Value = "���������"
    ws.Cells(1, 6).Value = "���-��"
    ws.Cells(1, 7).Value = "�����"

    r = 2
    r = DictAddRow(ws, r, "������������ �������", "�����_��������", "������� ��������� �����", "���", 115)
    r = DictAddRow(ws, r, "��������� ����", "�����_��������", "������� ��������� �����", "���", 114)
    r = DictAddRow(ws, r, "��������� ����", "�����_��������", "������� ��������� ����", "���", 113)
    r = DictAddRow(ws, r, "����.����", "�����_��������", "����. ����� (�������)", "���", 112)
    r = DictAddRow(ws, r, "�� ���� ����", "�����_��������", "�� ����. ����� (�������)", "���", 111)
    r = DictAddRow(ws, r, "����� �����", "�����_��������", "����������� �����", "���", 110)
    r = DictAddRow(ws, r, "���������� ��������", "����������", "���������� ���������", "���", 100)
    r = DictAddRow(ws, r, "���������� ���������", "����������", "���������� ���������", "���", 99)
    r = DictAddRow(ws, r, "���������� ��������", "����������", "���������� ���������", "���", 98)
    r = DictAddRow(ws, r, "�� ����������", "����������", "������ �� ����������", "���", 97)
    r = DictAddRow(ws, r, "������ �������", "������", "������ ���������", "���", 96)
    r = DictAddRow(ws, r, "������ �����", "������", "������ ���������", "���", 95)
    r = DictAddRow(ws, r, "������ ���", "������", "������ �����.", "���", 94)
    r = DictAddRow(ws, r, "������ �����", "������", "������ ��������", "���", 93)
    r = DictAddRow(ws, r, "������", "������", "������ (�����)", "���", 85)
    r = DictAddRow(ws, r, "������������", "���_������", "����� �� �������������", "��", 95)
    r = DictAddRow(ws, r, "�������������", "���_������", "�������������", "��", 94)
    r = DictAddRow(ws, r, "��������������", "���_������", "��������������", "��", 93)
    r = DictAddRow(ws, r, "���", "���_������", "������ �� ���", "��", 92)
    r = DictAddRow(ws, r, "�������", "���_������", "�������", "��", 91)
    r = DictAddRow(ws, r, "������� �� ����", "���_������", "������� �� �����", "��", 90)
    r = DictAddRow(ws, r, "���������� �� ����", "���_������", "���������� �����", "��", 89)
    r = DictAddRow(ws, r, "���������� ����", "���_������", "���������� ����", "��", 88)
    r = DictAddRow(ws, r, "�� ����", "���_������", "�� �����", "��", 87)
    r = DictAddRow(ws, r, "���������� ��", "���_������", "���������� ��", "��", 86)
    r = DictAddRow(ws, r, "��-����", "���_������", "��-����", "��", 85)
    r = DictAddRow(ws, r, "�� ����", "���_������", "�� �����", "��", 84)
    r = DictAddRow(ws, r, "��������", "���_������", "���������� �������", "��", 83)
    r = DictAddRow(ws, r, "����������", "���_������", "����������", "��", 82)
    r = DictAddRow(ws, r, "���������� �� �������", "���_������", "���������� �� �������", "��", 81)
    r = DictAddRow(ws, r, "������� �� �������", "���_������", "�������� �� �������", "��", 80)
    r = DictAddRow(ws, r, "���������� �� �������", "���_������", "���������� �� ��������", "��", 79)
    r = DictAddRow(ws, r, "�� ��������", "���_������", "������ �� ��������", "��", 78)
    r = DictAddRow(ws, r, "�������������", "���_������", "�������������", "��", 77)
    r = DictAddRow(ws, r, "���������", "���_������", "������������� �������", "��", 76)
    r = DictAddRow(ws, r, "���-��", "���_������", "���-��", "��", 75)
    r = DictAddRow(ws, r, "���", "���_������", "���", "��", 74)
    r = DictAddRow(ws, r, "������", "���_������", "������", "��", 73)
    r = DictAddRow(ws, r, "������", "���_������", "������", "��", 72)
    r = DictAddRow(ws, r, "����������", "���_������", "������������ �� ��-����", "��", 71)
    r = DictAddRow(ws, r, "��������", "���_������", "��������", "��", 70)
    r = DictAddRow(ws, r, "��������", "���_������", "���������� �����", "��", 69)
    r = DictAddRow(ws, r, "������", "���_������", "����������", "��", 68)
    r = DictAddRow(ws, r, "��������", "������", "�������� �����", "���", 50)
    r = DictAddRow(ws, r, "�������� ����", "������", "���. �����", "���", 49)
    r = DictAddRow(ws, r, "��������������", "������", "��������������", "���", 40)
    r = DictAddRow(ws, r, "�����", "������", "�����", "���", 39)
    r = DictAddRow(ws, r, "������ ������", "������_������", "������ � ������ ������", "���", 10)

    ws.Columns("A:G").AutoFit
    StyleHeaderRow ws, 1, 7
End Sub

Private Function DictAddRow(ws As Worksheet, r As Long, kw As String, tg As String, desc As String, sumFlag As String, pri As Long) As Long
    ws.Cells(r, 1).Value = kw
    ws.Cells(r, 2).Value = tg
    ws.Cells(r, 3).Value = desc
    ws.Cells(r, 4).Value = sumFlag
    ws.Cells(r, 5).Value = pri
    DictAddRow = r + 1
End Function

' ============================================================================
Private Function ClassifyPurpose(purpose As String) As TagInfo
    Dim result As TagInfo
    Dim p As String
    Dim bestPri As Long: bestPri = -1
    Dim i As Long, kw As String, pri As Long

    result.Tag = "���_���������"
    result.Summarize = False
    result.MatchedKeyword = ""

    p = LCase$(purpose)
    If Len(p) = 0 Then
        ClassifyPurpose = result
        Exit Function
    End If

    For i = 1 To mDictKeywords.Count
        kw = mDictKeywords(i)
        pri = mDictPriority(i)
        If InStr(1, p, kw, vbTextCompare) > 0 Then
            If pri > bestPri Then
                bestPri = pri
                result.Tag = mDictTags(i)
                result.Summarize = mDictSummarize(i)
                result.MatchedKeyword = kw
            End If
        End If
    Next i

    ClassifyPurpose = result
End Function

' ============================================================================
Private Sub BuildBankReport()
    Dim wsSrc As Worksheet, wsRpt As Worksheet
    Dim lastRow As Long, outRow As Long, i As Long
    Dim purpose As String, tag As TagInfo
    Dim inn As String, kpp As String, contractNum As String, contractId As String
    Dim sumVal As Double, dt As Variant
    Dim matchStatus As String
    Dim rnData As Collection

    Set wsSrc = GetBankSheet
    If wsSrc Is Nothing Then Err.Raise vbObjectError + 2, , "���� ����� �� ������"

    Set rnData = LoadRNTopups
    Set wsRpt = GetOrCreateSheet(SHEET_OTCHET_BANK)
    wsRpt.Cells.Clear

    wsRpt.Cells(1, 1).Value = "N"
    wsRpt.Cells(1, 2).Value = "ID �����"
    wsRpt.Cells(1, 3).Value = "���� �����"
    wsRpt.Cells(1, 4).Value = "�����"
    wsRpt.Cells(1, 5).Value = "���"
    wsRpt.Cells(1, 6).Value = "���"
    wsRpt.Cells(1, 7).Value = "����������"
    wsRpt.Cells(1, 8).Value = "�������"
    wsRpt.Cells(1, 9).Value = "ID ��������"
    wsRpt.Cells(1, 10).Value = "���"
    wsRpt.Cells(1, 11).Value = "����"
    wsRpt.Cells(1, 12).Value = "�����������"
    wsRpt.Cells(1, 13).Value = "���������� �������"
    wsRpt.Cells(1, 14).Value = "������ ������"
    wsRpt.Cells(1, 15).Value = "ID ���������� ��"
    wsRpt.Cells(1, 16).Value = "����������"
    StyleHeaderRow wsRpt, 1, 16

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, BK_ID).End(xlUp).Row
    outRow = 2

    For i = 2 To lastRow
        sumVal = CDbl(Val(wsSrc.Cells(i, BK_SUM).Value))
        dt = wsSrc.Cells(i, BK_DATE).Value
        inn = NormalizeINN(CStr(wsSrc.Cells(i, BK_INN).Value))
        purpose = CStr(wsSrc.Cells(i, BK_PURPOSE).Value)

        tag = ClassifyPurpose(purpose)
        contractNum = ExtractContract(purpose)
        contractId = FindContractId(contractNum)
        kpp = LookupKPP(inn, contractId)

        matchStatus = "-"
        Dim rnTransId As String: rnTransId = ""
        Dim note As String: note = ""

        If tag.Summarize Then
            Dim rnMatch As Variant
            Set rnMatch = FindRNMatch(rnData, sumVal, inn, kpp, contractId, CDate(dt))
            If Not rnMatch Is Nothing Then
                matchStatus = "���������"
                rnTransId = rnMatch("TransId")
                rnMatch("Used") = True
                        Else
                Dim usedRn As Variant
                Set usedRn = FindUsedRNMatch(rnData, sumVal, inn, kpp, contractId, CDate(dt))
                If Not usedRn Is Nothing Then
                    matchStatus = "�� ��� ������������"
                    note = "���������� �� ��� ��������� � ������� ������� ����� (�����. ID: " & usedRn("TransId") & ")"
                Else
                    matchStatus = "�� ���������"
                    note = "��� ���������� � ��-����"
                End If
            End If
        ElseIf tag.Tag = "�����_��������" Then
            matchStatus = "�������� (�� �������)"
            note = "������������� � �����"
        Else
            matchStatus = "�� �� ��-����"
            note = "�� �����������"
        End If

        wsRpt.Cells(outRow, 1).Value = outRow - 1
        wsRpt.Cells(outRow, 2).Value = wsSrc.Cells(i, BK_ID).Value
        wsRpt.Cells(outRow, 3).Value = dt
        wsRpt.Cells(outRow, 3).NumberFormat = "dd.mm.yyyy"
        wsRpt.Cells(outRow, 4).Value = sumVal
        wsRpt.Cells(outRow, 4).NumberFormat = "#,##0.00"
        wsRpt.Cells(outRow, 5).Value = inn
        wsRpt.Cells(outRow, 6).Value = kpp

        Dim bankClient As ClientInfo
        bankClient = LookupClientByINN(inn)
        If bankClient.ID > 0 Then
            AddContractHyperlink wsRpt.Cells(outRow, 7), bankClient.ID, CStr(wsSrc.Cells(i, BK_CPNAME).Value)
        Else
            wsRpt.Cells(outRow, 7).Value = CStr(wsSrc.Cells(i, BK_CPNAME).Value)
        End If
        wsRpt.Cells(outRow, 8).Value = contractNum
        wsRpt.Cells(outRow, 9).Value = contractId
        wsRpt.Cells(outRow, 10).Value = tag.Tag
        wsRpt.Cells(outRow, 11).Value = tag.MatchedKeyword
        wsRpt.Cells(outRow, 12).Value = IIf(tag.Summarize, "��", "���")
        wsRpt.Cells(outRow, 13).Value = purpose
        wsRpt.Cells(outRow, 14).Value = matchStatus
        wsRpt.Cells(outRow, 15).Value = rnTransId
        wsRpt.Cells(outRow, 16).Value = note

        ColorBankRow wsRpt, outRow, matchStatus, tag.Summarize
        outRow = outRow + 1
    Next i

    wsRpt.Columns("A:P").AutoFit
    wsRpt.Rows("1:1").RowHeight = 28
    ApplyFreezeTopRow wsRpt
    wsRpt.AutoFilterMode = False
    wsRpt.Range(wsRpt.Cells(1, 1), wsRpt.Cells(outRow - 1, 16)).AutoFilter
End Sub

' ============================================================================
Private Sub BuildRNReport()
    Dim wsSrc As Worksheet, wsRpt As Worksheet
    Dim lastRow As Long, outRow As Long, i As Long
    Dim opType As String, sumVal As Double, dt As Variant
    Dim clientName As String, contractId As String, transId As String
    Dim inn As String, kpp As String, contractNum As String
    Dim status As String, note As String
    Dim bankMatches As Collection
    Dim dupKeys As Collection

    Set wsSrc = GetRNSheet
    If wsSrc Is Nothing Then Err.Raise vbObjectError + 3, , "���� ��-���� �� ������"

    Set bankMatches = LoadBankGSMMatches
    Set dupKeys = New Collection
    Set wsRpt = GetOrCreateSheet(SHEET_OTCHET_RN)
    wsRpt.Cells.Clear

    wsRpt.Cells(1, 1).Value = "N"
    wsRpt.Cells(1, 2).Value = "����/�����"
    wsRpt.Cells(1, 3).Value = "�����"
    wsRpt.Cells(1, 4).Value = "���"
    wsRpt.Cells(1, 5).Value = "������"
    wsRpt.Cells(1, 6).Value = "�������"
    wsRpt.Cells(1, 7).Value = "ID ��������"
    wsRpt.Cells(1, 8).Value = "ID ����������"
    wsRpt.Cells(1, 9).Value = "���"
    wsRpt.Cells(1, 10).Value = "���"
    wsRpt.Cells(1, 11).Value = "������ ������"
    wsRpt.Cells(1, 12).Value = "���� ID"
    wsRpt.Cells(1, 13).Value = "����������"
    StyleHeaderRow wsRpt, 1, 13

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, RN_DATE).End(xlUp).Row
    outRow = 2

    For i = 2 To lastRow
        opType = CStr(wsSrc.Cells(i, RN_OPTYPE).Value)
        If Not IsPopoln(opType) Then GoTo NextRN

        sumVal = CDbl(Val(wsSrc.Cells(i, RN_SUM).Value))
        dt = wsSrc.Cells(i, RN_DATE).Value
        clientName = CStr(wsSrc.Cells(i, RN_CLIENT).Value)
        contractNum = CStr(wsSrc.Cells(i, RN_CONTRACT).Value)
        contractId = CStr(wsSrc.Cells(i, RN_CONTRACTID).Value)
        transId = CStr(wsSrc.Cells(i, RN_TRANSID).Value)

        Dim ci As ClientInfo
        ci = LookupClientByName(clientName)
        inn = ci.INN
        kpp = ci.KPP

        status = "�� ������� � �����"
        note = ""
        Dim bankId As String: bankId = ""

        Dim dupKey As String
        dupKey = BuildMatchKey(sumVal, dt, inn, kpp, contractId)
        On Error Resume Next
        dupKeys.Add dupKey, dupKey
        If Err.Number <> 0 Then
            status = "�����"
            note = "��������� ���������� � ��� �� ������"
            Err.Clear
        End If
        On Error GoTo 0

        Dim bm As Variant
        Set bm = FindBankMatch(bankMatches, sumVal, inn, kpp, contractId, CDate(dt))
        If Not bm Is Nothing Then
            If status = "�����" Then
                note = note & "; ���� �����."
            Else
                status = "������� � ������"
            End If
            bankId = bm("BankId")
            bm("Used") = True
        ElseIf status <> "�����" Then
            Dim usedBm As Variant
            Set usedBm = FindUsedBankMatch(bankMatches, sumVal, inn, kpp, contractId, CDate(dt))
            If Not usedBm Is Nothing Then
                status = "���� ��� �����������"
                note = "���������� ����� ��� ����������� � ������ ����������� �� (���� ID: " & usedBm("BankId") & ")"
            Else
                note = "� ����� ��� ����������� ������� ��� �������������"
            End If
        End If

        wsRpt.Cells(outRow, 1).Value = outRow - 1
        wsRpt.Cells(outRow, 2).Value = dt
        wsRpt.Cells(outRow, 2).NumberFormat = "dd.mm.yyyy hh:mm"
        wsRpt.Cells(outRow, 3).Value = sumVal
        wsRpt.Cells(outRow, 3).NumberFormat = "#,##0.00"
        wsRpt.Cells(outRow, 4).Value = opType
        wsRpt.Cells(outRow, 6).Value = contractNum
        wsRpt.Cells(outRow, 7).Value = contractId
        wsRpt.Cells(outRow, 8).Value = transId
        wsRpt.Cells(outRow, 9).Value = inn
        wsRpt.Cells(outRow, 10).Value = kpp
        wsRpt.Cells(outRow, 11).Value = status
        wsRpt.Cells(outRow, 12).Value = bankId
        wsRpt.Cells(outRow, 13).Value = note

        If ci.ID > 0 Then
            AddContractHyperlink wsRpt.Cells(outRow, 5), ci.ID, clientName
        Else
            wsRpt.Cells(outRow, 5).Value = clientName
        End If

        ColorRNRow wsRpt, outRow, status
        outRow = outRow + 1
NextRN:
    Next i

    wsRpt.Columns("A:M").AutoFit
    wsRpt.Columns("E").ColumnWidth = 35
    wsRpt.Rows("1:1").RowHeight = 28
    ApplyFreezeTopRow wsRpt
    wsRpt.AutoFilterMode = False
    If outRow > 2 Then
        wsRpt.Range(wsRpt.Cells(1, 1), wsRpt.Cells(outRow - 1, 13)).AutoFilter
    End If
End Sub

' ============================================================================
Private Sub BuildSummary()
    Dim wsBank As Worksheet, wsRN As Worksheet, wsSum As Worksheet
    Dim lastRow As Long, i As Long
    Dim periodFrom As Date, periodTo As Date
    Dim tagTotals As Object, statusTotals As Object
    Dim key As Variant

    Set wsBank = ThisWorkbook.Worksheets(SHEET_OTCHET_BANK)
    Set wsRN = ThisWorkbook.Worksheets(SHEET_OTCHET_RN)
    Set wsSum = GetOrCreateSheet(SHEET_ITOG)
    wsSum.Cells.Clear

    Set tagTotals = CreateObject("Scripting.Dictionary")
    Set statusTotals = CreateObject("Scripting.Dictionary")

    lastRow = wsBank.Cells(wsBank.Rows.Count, 1).End(xlUp).Row
    periodFrom = DateSerial(2099, 12, 31)
    periodTo = DateSerial(1900, 1, 1)

    For i = 2 To lastRow
        Dim dt As Variant: dt = wsBank.Cells(i, 3).Value
        If IsDate(dt) Then
            If CDate(dt) < periodFrom Then periodFrom = CDate(dt)
            If CDate(dt) > periodTo Then periodTo = CDate(dt)
        End If

        Dim tg As String: tg = CStr(wsBank.Cells(i, 10).Value)
        Dim sm As Double: sm = CDbl(Val(wsBank.Cells(i, 4).Value))
        Dim sumFlag As String: sumFlag = CStr(wsBank.Cells(i, 12).Value)
        Dim st As String: st = CStr(wsBank.Cells(i, 14).Value)

        If Not tagTotals.Exists(tg) Then
            tagTotals.Add tg, Array(0&, 0#, 0&, 0#)
        End If
        Dim ta As Variant: ta = tagTotals(tg)
        ta(0) = ta(0) + 1
        ta(1) = ta(1) + sm
        If sumFlag = "��" Then
            ta(2) = ta(2) + 1
            ta(3) = ta(3) + sm
        End If
        tagTotals(tg) = ta

        If Not statusTotals.Exists(st) Then
            statusTotals.Add st, Array(0&, 0#)
        End If
        Dim sa As Variant: sa = statusTotals(st)
        sa(0) = sa(0) + 1
        sa(1) = sa(1) + sm
        statusTotals(st) = sa
    Next i

    wsSum.Range("A1").Value = "������� ����� �� ������ ��-����"
    wsSum.Range("A1").Font.Size = 16
    wsSum.Range("A1").Font.Bold = True
    wsSum.Range("A1").Font.Color = CLR_HEADER

    wsSum.Range("A2").Value = "������: " & Format(periodFrom, "dd.mm.yyyy") & " - " & Format(periodTo, "dd.mm.yyyy")
    wsSum.Range("A2").Font.Size = 11
    wsSum.Range("A2").Font.Italic = True

    wsSum.Range("A3").Value = "������������: " & Format(Now, "dd.mm.yyyy hh:mm")
    wsSum.Range("A3").Font.Size = 10
    wsSum.Range("A3").Font.Color = 8421504

    Dim r As Long: r = 5
    wsSum.Cells(r, 1).Value = "������������� �� �����"
    StyleSectionTitle wsSum, r, 6
    r = r + 1

    wsSum.Cells(r, 1).Value = "���"
    wsSum.Cells(r, 2).Value = "���-��"
    wsSum.Cells(r, 3).Value = "����� �����"
    wsSum.Cells(r, 4).Value = "� ������ (���)"
    wsSum.Cells(r, 5).Value = "����� ���"
    StyleHeaderRow wsSum, r, 5
    r = r + 1

    Dim totalGSM As Double: totalGSM = 0
    Dim countGSM As Long: countGSM = 0
    Dim totalOther As Double: totalOther = 0

    For Each key In tagTotals.Keys
        ta = tagTotals(key)
        wsSum.Cells(r, 1).Value = key
        wsSum.Cells(r, 2).Value = ta(0)
        wsSum.Cells(r, 3).Value = ta(1)
        wsSum.Cells(r, 3).NumberFormat = "#,##0.00"
        wsSum.Cells(r, 4).Value = ta(2)
        wsSum.Cells(r, 5).Value = ta(3)
        wsSum.Cells(r, 5).NumberFormat = "#,##0.00"
        If ta(2) > 0 Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 5)).Interior.Color = CLR_GSM
            totalGSM = totalGSM + ta(3)
            countGSM = countGSM + ta(2)
        Else
            totalOther = totalOther + ta(1)
        End If
        r = r + 1
    Next key

    r = r + 1
    wsSum.Cells(r, 1).Value = "����� ��� (������)"
    wsSum.Cells(r, 1).Font.Bold = True
    wsSum.Cells(r, 4).Value = countGSM
    wsSum.Cells(r, 5).Value = totalGSM
    wsSum.Cells(r, 5).NumberFormat = "#,##0.00"
    wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 5)).Interior.Color = CLR_OK
    wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 5)).Font.Bold = True

    r = r + 3
    wsSum.Cells(r, 1).Value = "������� ������ � �����"
    StyleSectionTitle wsSum, r, 4
    r = r + 1

    wsSum.Cells(r, 1).Value = "������"
    wsSum.Cells(r, 2).Value = "���-��"
    wsSum.Cells(r, 3).Value = "�����"
    StyleHeaderRow wsSum, r, 3
    r = r + 1

    Dim matchedSum As Double: matchedSum = 0
    Dim unmatchedSum As Double: unmatchedSum = 0

    For Each key In statusTotals.Keys
        sa = statusTotals(key)
        wsSum.Cells(r, 1).Value = key
        wsSum.Cells(r, 2).Value = sa(0)
        wsSum.Cells(r, 3).Value = sa(1)
        wsSum.Cells(r, 3).NumberFormat = "#,##0.00"
        If key = "���������" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_OK
            matchedSum = sa(1)
        ElseIf key = "�� ���������" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_ERR
            unmatchedSum = sa(1)
        ElseIf key = "�������� (�� �������)" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_WARN
        ElseIf key = "�� ��� ������������" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_WARN
        End If
        r = r + 1
    Next key

    r = r + 2
    wsSum.Cells(r, 1).Value = "������� ������ � ��-����"
    StyleSectionTitle wsSum, r, 4
    r = r + 1

    Dim rnStatus As Object
    Set rnStatus = CreateObject("Scripting.Dictionary")
    lastRow = wsRN.Cells(wsRN.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        st = CStr(wsRN.Cells(i, 11).Value)
        sm = CDbl(Val(wsRN.Cells(i, 3).Value))
        If Not rnStatus.Exists(st) Then rnStatus.Add st, Array(0&, 0#)
        sa = rnStatus(st)
        sa(0) = sa(0) + 1
        sa(1) = sa(1) + sm
        rnStatus(st) = sa
    Next i

    wsSum.Cells(r, 1).Value = "������"
    wsSum.Cells(r, 2).Value = "���-��"
    wsSum.Cells(r, 3).Value = "�����"
    StyleHeaderRow wsSum, r, 3
    r = r + 1

    For Each key In rnStatus.Keys
        sa = rnStatus(key)
        wsSum.Cells(r, 1).Value = key
        wsSum.Cells(r, 2).Value = sa(0)
        wsSum.Cells(r, 3).Value = sa(1)
        wsSum.Cells(r, 3).NumberFormat = "#,##0.00"
        If key = "������� � ������" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_OK
        ElseIf key = "�� ������� � �����" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_ERR
        ElseIf key = "�����" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_WARN
        ElseIf key = "���� ��� �����������" Then
            wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 3)).Interior.Color = CLR_WARN
        End If
        r = r + 1
    Next key

    r = r + 3
    wsSum.Cells(r, 1).Value = "���� ������:"
    wsSum.Cells(r, 1).Font.Bold = True
    wsSum.Cells(r, 2).Value = "����� + ���� + ��� + ��� + ID ��������"
    wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 2)).Font.Size = 10

    r = r + 1
    wsSum.Cells(r, 1).Value = "�����������:"
    wsSum.Cells(r, 1).Font.Bold = True
    wsSum.Cells(r, 2).Value = "https://lkp.rn-card.ru/Contract/?client=ID&cid=ID+1&act=open#info"

    wsSum.Columns("A:F").AutoFit
    wsSum.Columns("A").ColumnWidth = 30
    wsSum.Columns("B").ColumnWidth = 18
End Sub

' ============================================================================
Private Sub RefreshDictionaryStats()
    Dim wsDict As Worksheet, wsBank As Worksheet
    Dim lastBank As Long, lastDict As Long, i As Long
    Dim tag As String
    Dim tagSums As Object

    Set wsDict = ThisWorkbook.Worksheets(SHEET_SLOVAR)
    Set wsBank = ThisWorkbook.Worksheets(SHEET_OTCHET_BANK)
    Set tagSums = CreateObject("Scripting.Dictionary")

    lastBank = wsBank.Cells(wsBank.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastBank
        tag = CStr(wsBank.Cells(i, 10).Value)
        If Not tagSums.Exists(tag) Then tagSums.Add tag, Array(0&, 0#)
        Dim a As Variant: a = tagSums(tag)
        a(0) = a(0) + 1
        a(1) = a(1) + CDbl(Val(wsBank.Cells(i, 4).Value))
        tagSums(tag) = a
    Next i

    lastDict = wsDict.Cells(wsDict.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastDict
        tag = CStr(wsDict.Cells(i, 2).Value)
        wsDict.Cells(i, 6).Value = ""
        wsDict.Cells(i, 7).Value = ""
        If tagSums.Exists(tag) Then
            a = tagSums(tag)
            wsDict.Cells(i, 6).Value = a(0)
            wsDict.Cells(i, 7).Value = a(1)
            wsDict.Cells(i, 7).NumberFormat = "#,##0.00"
        End If
    Next i
End Sub

' ============================================================================
Private Function LoadRNTopups() As Collection
    Dim ws As Worksheet, col As New Collection
    Dim lastRow As Long, i As Long

    Set ws = GetRNSheet
    If ws Is Nothing Then Set LoadRNTopups = col: Exit Function

    lastRow = ws.Cells(ws.Rows.Count, RN_DATE).End(xlUp).Row
    For i = 2 To lastRow
        If IsPopoln(CStr(ws.Cells(i, RN_OPTYPE).Value)) Then
            Dim ci As ClientInfo
            ci = LookupClientByName(CStr(ws.Cells(i, RN_CLIENT).Value))
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            item("Sum") = CDbl(Val(ws.Cells(i, RN_SUM).Value))
            item("Date") = CDate(ws.Cells(i, RN_DATE).Value)
            item("INN") = ci.INN
            item("KPP") = ci.KPP
            item("ContractId") = CStr(ws.Cells(i, RN_CONTRACTID).Value)
            item("TransId") = CStr(ws.Cells(i, RN_TRANSID).Value)
            item("Used") = False
            col.Add item
        End If
    Next i
    Set LoadRNTopups = col
End Function

Private Function LoadBankGSMMatches() As Collection
    Dim ws As Worksheet, col As New Collection
    Dim lastRow As Long, i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_OTCHET_BANK)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 12).Value) = "��" Then
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            item("Sum") = CDbl(Val(ws.Cells(i, 4).Value))
            item("Date") = CDate(ws.Cells(i, 3).Value)
            item("INN") = CStr(ws.Cells(i, 5).Value)
            item("KPP") = CStr(ws.Cells(i, 6).Value)
            item("ContractId") = CStr(ws.Cells(i, 9).Value)
            item("BankId") = CStr(ws.Cells(i, 2).Value)
            item("Used") = False
            col.Add item
        End If
    Next i
    Set LoadBankGSMMatches = col
End Function

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

Private Function FindUsedRNMatch(rnData As Collection, sumVal As Double, inn As String, kpp As String, contractId As String, bankDate As Date) As Variant
    Dim i As Long, item As Object
    For i = 1 To rnData.Count
        Set item = rnData(i)
        If Not item("Used") Then GoTo NextUI
        If ItemMatchesBankRN(item, sumVal, inn, kpp, contractId, bankDate) Then
            Set FindUsedRNMatch = item
            Exit Function
        End If
NextUI:
    Next i
    Set FindUsedRNMatch = Nothing
End Function

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

Private Function FindUsedBankMatch(bankData As Collection, sumVal As Double, inn As String, kpp As String, contractId As String, rnDate As Date) As Variant
    Dim i As Long, item As Object
    For i = 1 To bankData.Count
        Set item = bankData(i)
        If Not item("Used") Then GoTo NextUBI
        If ItemMatchesRNBank(item, sumVal, inn, kpp, contractId, rnDate) Then
            Set FindUsedBankMatch = item
            Exit Function
        End If
NextUBI:
    Next i
    Set FindUsedBankMatch = Nothing
End Function

' ============================================================================
Private Function LookupClientByINN(inn As String) As ClientInfo
    Dim i As Long
    Dim noClient As ClientInfo
    If inn = "" Then
        LookupClientByINN = noClient
        Exit Function
    End If
    For i = 1 To mClientCount
        If mClients(i).INN = inn Then
            LookupClientByINN = mClients(i)
            Exit Function
        End If
    Next i
    LookupClientByINN = noClient
End Function

Private Function LookupClientByName(name As String) As ClientInfo
    Dim i As Long, n As String
    Dim noClient As ClientInfo
    n = Trim$(name)
    For i = 1 To mClientCount
        If mClients(i).ShortName = n Or mClients(i).FullName = n Then
            LookupClientByName = mClients(i)
            Exit Function
        End If
    Next i
    LookupClientByName = noClient
End Function

Private Function LookupKPP(inn As String, contractId As String) As String
    Dim i As Long
    For i = 1 To mClientCount
        If mClients(i).INN = inn Then
            LookupKPP = mClients(i).KPP
            Exit Function
        End If
    Next i
    LookupKPP = ""
End Function

Private Function FindContractId(contractNum As String) As String
    Dim ws As Worksheet, lastRow As Long, i As Long
    If contractNum = "" Then
        FindContractId = ""
        Exit Function
    End If

    Set ws = GetRNSheet
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, RN_DATE).End(xlUp).Row
    Dim norm As String: norm = NormalizeContract(contractNum)

    For i = 2 To lastRow
        Dim c As String
        c = NormalizeContract(CStr(ws.Cells(i, RN_CONTRACT).Value))
        If c = norm Then
            FindContractId = CStr(ws.Cells(i, RN_CONTRACTID).Value)
            Exit Function
        End If
    Next i
    FindContractId = ""
End Function

' ============================================================================
Private Function IsPopoln(opType As String) As Boolean
    Dim s As String
    Dim p As Long
    s = LCase$(opType)
    p = InStr(1, s, ChrW(1087) & ChrW(1086) & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1085), vbTextCompare)
    IsPopoln = (p > 0)
End Function

Private Function ExtractContract(text As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    re.Pattern = "(\d{1,3}\.\d{2}[/\.]\d{2,4})"
    If re.Test(text) Then
        Set matches = re.Execute(text)
        ExtractContract = matches(0).SubMatches(0)
        Exit Function
    End If

    re.Pattern = "(?:���(?:����)?\.?\s*(?:�|N)?\s*)([\w\./-]+)"
    If re.Test(text) Then
        Set matches = re.Execute(text)
        ExtractContract = matches(0).SubMatches(0)
        Exit Function
    End If

    ExtractContract = ""
End Function

Private Function NormalizeContract(s As String) As String
    Dim r As String
    r = LCase$(Trim$(s))
    r = Replace(r, " ", "")
    r = Replace(r, "�", "")
    r = Replace(r, "���", "")
    r = Replace(r, "�������", "")
    NormalizeContract = r
End Function

Private Function NormalizeINN(s As String) As String
    Dim r As String
    r = Trim$(s)
    If InStr(r, "E") > 0 Then r = Format(CDbl(r), "0")
    NormalizeINN = r
End Function

Private Function NormalizeKPP(s As String) As String
    Dim r As String
    r = Trim$(s)
    If InStr(r, "E") > 0 Then r = Format(CDbl(r), "0")
    NormalizeKPP = r
End Function

Private Function BuildMatchKey(sumVal As Double, dt As Variant, inn As String, kpp As String, contractId As String) As String
    BuildMatchKey = Format(sumVal, "0.00") & "|" & Format(CDate(dt), "yyyy-mm-dd hh:nn:ss") & "|" & inn & "|" & kpp & "|" & contractId
End Function

' ============================================================================
Private Sub AddContractHyperlink(cell As Range, clientId As Long, displayText As String)
    Dim url As String
    url = "https://lkp.rn-card.ru/Contract/?client=" & clientId & "&cid=" & (clientId + 1) & "&act=open#info"
    cell.Hyperlinks.Delete
    cell.Hyperlinks.Add Anchor:=cell, Address:=url, TextToDisplay:=displayText
    cell.Font.Color = 12611584
    cell.Font.Underline = xlUnderlineStyleSingle
End Sub

' ============================================================================
Private Function GetSheetByName(name As String) As Worksheet
    On Error Resume Next
    Set GetSheetByName = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function GetOrCreateSheet(name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function GetBankSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Trim$(CStr(ws.Cells(1, 1).Value)) = "ID" Then
            If ws.UsedRange.Columns.Count = 7 Then
                Set GetBankSheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function GetRNSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.UsedRange.Columns.Count = 11 Then
            Set GetRNSheet = ws
            Exit Function
        End If
    Next ws
End Function

' ============================================================================
Private Sub StyleHeaderRow(ws As Worksheet, rowNum As Long, colCount As Long)
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, colCount))
        .Interior.Color = CLR_HEADER
        .Font.Color = CLR_HEADER_FG
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Sub StyleSectionTitle(ws As Worksheet, rowNum As Long, colCount As Long)
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, colCount))
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = CLR_HEADER
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub


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
    Dim clr As Long
    Select Case status
        Case "���������": clr = CLR_OK
        Case "�� ���������": clr = CLR_ERR
        Case "�������� (�� �������)": clr = CLR_WARN
        Case "�� ��� ������������": clr = CLR_WARN
        Case Else
            clr = -1
    End Select
    If clr >= 0 Then
        ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 16)).Interior.Color = clr
    End If
End Sub

Private Sub ColorRNRow(ws As Worksheet, rowNum As Long, status As String)
    Dim clr As Long
    Select Case status
        Case "������� � ������": clr = CLR_OK
        Case "�� ������� � �����": clr = CLR_ERR
        Case "�����": clr = CLR_WARN
        Case "���� ��� �����������": clr = CLR_WARN
        Case Else: clr = -1
    End Select
    If clr >= 0 Then
        ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 13)).Interior.Color = clr
    End If
End Sub
