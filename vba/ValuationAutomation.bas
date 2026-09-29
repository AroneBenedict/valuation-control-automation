Attribute VB_Name = "ValuationAutomation"
Option Explicit

'==============================================================================
' Position Price Verification - Reporting Automation (Excel VBA)
' SYNTHETIC DATA PROJECT. Requires Excel for Windows (uses Scripting.Dictionary).
'
' Expects sheets:  Raw_Data (A:I)   and   Thresholds (A2:B5 tolerances, E1 report date)
' Creates/refreshes:  Clean_Data, Exceptions   (deleted and rebuilt on every run)
' Appends to:         Run_Log (audit trail - one row per run, never deleted)
' Exports:            Clean_Positions.csv (next to the saved workbook)
'
' Macros to run:  RunAll  (does everything)  |  RunValuationControls  |  ExportCleanCSV
'==============================================================================

Private Const RAW_SHEET As String = "Raw_Data"
Private Const THR_SHEET As String = "Thresholds"
Private Const CLEAN_SHEET As String = "Clean_Data"
Private Const EXC_SHEET As String = "Exceptions"
Private Const LOG_SHEET As String = "Run_Log"
Private Const NCOLS As Long = 14

Public Sub RunAll()
    Dim msg As String
    msg = ExecuteControls()
    msg = msg & vbCrLf & vbCrLf & ExportCsv()
    MsgBox msg, vbInformation, "Valuation controls"
End Sub

Public Sub RunValuationControls()
    MsgBox ExecuteControls(), vbInformation, "Valuation controls"
End Sub

Public Sub ExportCleanCSV()
    MsgBox ExportCsv(), vbInformation, "Export"
End Sub

'------------------------------------------------------------------------------
' Core routine: clean -> test -> write Clean_Data, Exceptions, Run_Log
'------------------------------------------------------------------------------
Private Function ExecuteControls() As String
    Dim wsRaw As Worksheet, wsThr As Worksheet, wsClean As Worksheet, wsExc As Worksheet, wsLog As Worksheet
    Dim tol As Object, canon As Object, seen As Object
    Dim data As Variant, outClean() As Variant, outExc() As Variant
    Dim r As Long, i As Long, c As Long, lastRow As Long, nClean As Long, nExc As Long, nextLog As Long
    Dim reportDate As Date
    Dim key As String, id As String, ac As String, status As String, staleFlag As String
    Dim qty As Double, bookP As Double, indP As Double, variance As Double, mv As Double, tolerance As Double
    Dim isMapped As Boolean, hasInd As Boolean, validNums As Boolean, dateOk As Boolean
    Dim dupCount As Long, noPrice As Long, unmapped As Long, staleCount As Long, invalid As Long, okCount As Long, rawRows As Long

    Application.ScreenUpdating = False
    On Error GoTo Fail

    Set wsRaw = ThisWorkbook.Worksheets(RAW_SHEET)
    Set wsThr = ThisWorkbook.Worksheets(THR_SHEET)
    reportDate = CDate(wsThr.Range("E1").Value)

    ' --- 1. load tolerances (case-insensitive lookup) --------------------------
    Set tol = CreateObject("Scripting.Dictionary"): tol.CompareMode = vbTextCompare
    Set canon = CreateObject("Scripting.Dictionary"): canon.CompareMode = vbTextCompare
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    r = 2
    Do While Len(Trim$(CStr(wsThr.Cells(r, 1).Value))) > 0
        key = Trim$(CStr(wsThr.Cells(r, 1).Value))
        canon(key) = key
        tol(key) = CDbl(wsThr.Cells(r, 2).Value)
        r = r + 1
    Loop

    ' --- 2. read raw data into memory (fast) -----------------------------------
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1, , "Raw_Data has no rows."
    data = wsRaw.Range("A1:I" & lastRow).Value
    rawRows = lastRow - 1
    ReDim outClean(1 To lastRow - 1, 1 To NCOLS)
    ReDim outExc(1 To lastRow - 1, 1 To NCOLS + 1)

    ' --- 3. clean + test each row ----------------------------------------------
    For r = 2 To lastRow
        id = Trim$(CStr(data(r, 1)))
        If Len(id) = 0 Then GoTo NextRow

        If seen.Exists(id) Then                 ' duplicate PositionID: keep first only
            dupCount = dupCount + 1
            GoTo NextRow
        End If
        seen.Add id, True

        ac = Trim$(CStr(data(r, 4)))            ' standardise asset class (trim + case)
        If canon.Exists(ac) Then
            ac = canon(ac): isMapped = True
        Else
            isMapped = False
        End If

        validNums = IsNum(data(r, 7)) And IsNum(data(r, 8))
        hasInd = IsNum(data(r, 9))
        If hasInd Then hasInd = (CDbl(data(r, 9)) <> 0)
        qty = 0: bookP = 0: indP = 0: variance = 0: mv = 0: tolerance = 0
        If validNums Then qty = CDbl(data(r, 7)): bookP = CDbl(data(r, 8))
        If hasInd Then indP = CDbl(data(r, 9))
        If isMapped Then tolerance = tol(ac)

        If validNums And hasInd Then
            variance = (bookP - indP) / indP
            mv = qty * (bookP - indP)
        End If

        If Not validNums Then
            status = "INVALID DATA": invalid = invalid + 1
        ElseIf Not isMapped Then
            status = "UNMAPPED ASSET CLASS": unmapped = unmapped + 1
        ElseIf Not hasInd Then
            status = "NO INDEPENDENT PRICE": noPrice = noPrice + 1
        ElseIf Abs(variance) > tolerance Then
            status = "BREACH"
        Else
            status = "OK": okCount = okCount + 1
        End If

        dateOk = IsDate(data(r, 2))
        If Not dateOk Then
            staleFlag = "INVALID DATE"
        ElseIf CLng(CDate(data(r, 2))) <> CLng(reportDate) Then
            staleFlag = "STALE": staleCount = staleCount + 1
        Else
            staleFlag = "CURRENT"
        End If

        nClean = nClean + 1
        outClean(nClean, 1) = id
        If dateOk Then outClean(nClean, 2) = CDate(data(r, 2)) Else outClean(nClean, 2) = data(r, 2)
        outClean(nClean, 3) = Trim$(CStr(data(r, 3)))
        outClean(nClean, 4) = ac
        outClean(nClean, 5) = Trim$(CStr(data(r, 5)))
        outClean(nClean, 6) = Trim$(CStr(data(r, 6)))
        If validNums Then outClean(nClean, 7) = qty: outClean(nClean, 8) = bookP
        If hasInd Then outClean(nClean, 9) = indP
        If validNums And hasInd Then outClean(nClean, 10) = variance: outClean(nClean, 12) = mv
        If isMapped Then outClean(nClean, 11) = tolerance
        outClean(nClean, 13) = status
        outClean(nClean, 14) = staleFlag

        If status = "BREACH" Then
            nExc = nExc + 1
            For c = 1 To NCOLS: outExc(nExc, c) = outClean(nClean, c): Next c
            outExc(nExc, NCOLS + 1) = Abs(mv)
        End If
NextRow:
    Next r

    ' --- 4. write Clean_Data ------------------------------------------------------
    Set wsClean = ResetSheet(CLEAN_SHEET)
    wsClean.Range("A1").Resize(1, NCOLS).Value = Array("PositionID", "ValuationDate", "Desk", "AssetClass", "Instrument", "Currency", _
        "Quantity", "BookPrice", "IndependentPrice", "VariancePct", "TolerancePct", "MVDiff", "Status", "DateFlag")
    wsClean.Range("A2").Resize(lastRow - 1, NCOLS).Value = outClean
    StyleSheet wsClean, NCOLS, lastRow

    ' --- 5. write Exceptions (sorted by absolute market-value difference) --------
    Set wsExc = ResetSheet(EXC_SHEET)
    wsExc.Range("A1").Resize(1, NCOLS + 1).Value = Array("PositionID", "ValuationDate", "Desk", "AssetClass", "Instrument", "Currency", _
        "Quantity", "BookPrice", "IndependentPrice", "VariancePct", "TolerancePct", "MVDiff", "Status", "DateFlag", "AbsMVDiff")
    If nExc > 0 Then wsExc.Range("A2").Resize(lastRow - 1, NCOLS + 1).Value = outExc
    StyleSheet wsExc, NCOLS + 1, lastRow
    If nExc > 1 Then
        With wsExc.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsExc.Range("O2:O" & nExc + 1), SortOn:=xlSortOnValues, Order:=xlDescending
            .SetRange wsExc.Range("A1:O" & nExc + 1)
            .Header = xlYes
            .Apply
        End With
    End If

    ' --- 6. audit trail -----------------------------------------------------------
    Set wsLog = GetOrCreateSheet(LOG_SHEET)
    If Len(CStr(wsLog.Cells(1, 1).Value)) = 0 Then
        wsLog.Range("A1").Resize(1, 10).Value = Array("RunTime", "ReportDate", "RawRows", "UniquePositions", "DuplicatesRemoved", _
            "MissingIndepPrice", "UnmappedClass", "StaleDate", "Breaches", "BreachRate")
        wsLog.Range("A1:J1").Font.Bold = True
    End If
    nextLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(nextLog, 1).Value = Now
    wsLog.Cells(nextLog, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsLog.Cells(nextLog, 2).Value = reportDate
    wsLog.Cells(nextLog, 2).NumberFormat = "yyyy-mm-dd"
    wsLog.Cells(nextLog, 3).Value = rawRows
    wsLog.Cells(nextLog, 4).Value = nClean
    wsLog.Cells(nextLog, 5).Value = dupCount
    wsLog.Cells(nextLog, 6).Value = noPrice
    wsLog.Cells(nextLog, 7).Value = unmapped
    wsLog.Cells(nextLog, 8).Value = staleCount
    wsLog.Cells(nextLog, 9).Value = nExc
    If (okCount + nExc) > 0 Then wsLog.Cells(nextLog, 10).Value = nExc / (okCount + nExc)
    wsLog.Cells(nextLog, 10).NumberFormat = "0.0%"
    wsLog.Columns("A:J").AutoFit

    ExecuteControls = "Raw rows: " & rawRows & vbCrLf & _
        "Duplicates removed: " & dupCount & vbCrLf & _
        "Unique positions: " & nClean & vbCrLf & _
        "Missing independent price: " & noPrice & vbCrLf & _
        "Unmapped asset class: " & unmapped & vbCrLf & _
        "Stale valuation date: " & staleCount & vbCrLf & _
        "OK: " & okCount & "   Breaches: " & nExc & vbCrLf & _
        "Breach rate: " & Format$(IIf((okCount + nExc) > 0, nExc / (okCount + nExc), 0), "0.0%")
    Application.ScreenUpdating = True
    Exit Function

Fail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ExecuteControls = "Run failed: " & Err.Description
End Function

'------------------------------------------------------------------------------
' Export the first 9 (cleaned input) columns for Access / Power BI
'------------------------------------------------------------------------------
Private Function ExportCsv() As String
    Dim ws As Worksheet, f As Integer, p As String
    Dim r As Long, c As Long, lastRow As Long, line As String, v As Variant, s As String

    If Len(ThisWorkbook.Path) = 0 Then
        ExportCsv = "CSV not exported: save the workbook first.": Exit Function
    End If
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(CLEAN_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    p = ThisWorkbook.Path & Application.PathSeparator & "Clean_Positions.csv"
    f = FreeFile
    Open p For Output As #f
    Print #f, "PositionID,ValuationDate,Desk,AssetClass,Instrument,Currency,Quantity,BookPrice,IndependentPrice"
    For r = 2 To lastRow
        line = ""
        For c = 1 To 9
            v = ws.Cells(r, c).Value
            If IsEmpty(v) Then
                s = ""
            ElseIf c = 2 And IsDate(v) Then
                s = Format$(CDate(v), "yyyy-mm-dd")
            ElseIf c >= 7 And IsNumeric(v) Then
                s = Trim$(Str$(CDbl(v)))             ' Str$ always uses "." as decimal separator
            Else
                s = CStr(v)
                If InStr(s, ",") > 0 Then s = """" & s & """"
            End If
            If c > 1 Then line = line & ","
            line = line & s
        Next c
        Print #f, line
    Next r
    Close #f
    ExportCsv = "Exported " & (lastRow - 1) & " rows to:" & vbCrLf & p
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    ExportCsv = "Export failed: " & Err.Description
End Function

'------------------------------ helpers ---------------------------------------
Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Then IsNum = False: Exit Function
    IsNum = (Len(Trim$(CStr(v))) > 0) And IsNumeric(v)
End Function

Private Function ResetSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    Set ResetSheet = ws
End Function

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub StyleSheet(ByVal ws As Worksheet, ByVal nCol As Long, ByVal lastRow As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCol))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd"
    ws.Range("G2:G" & lastRow).NumberFormat = "#,##0"
    ws.Range("H2:I" & lastRow).NumberFormat = "#,##0.00000"
    ws.Range("J2:K" & lastRow).NumberFormat = "0.000%"
    ws.Range("L2:L" & lastRow).NumberFormat = "#,##0.00;(#,##0.00)"
    If nCol > 14 Then ws.Range("O2:O" & lastRow).NumberFormat = "#,##0.00"
    ws.Columns(1).Resize(, nCol).AutoFit
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
End Sub
