Attribute VB_Name = "Common"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
#End If

Public Const Undefined& = -999
Public Const clmSortingIndex = 5

Public Sub ProgressBar(stringToShow$)
    Dim timePeriod_s&, i&
    
    frmProgressBar.Caption = stringToShow & "... in progress..."
    frmProgressBar.Show
    
    timePeriod_s = RandomRange(5, 20)
    
    frmProgressBar.lblTimeRemains.Caption = CStr(timePeriod_s) & " s"
    frmProgressBar.lblProgress.Width = frmProgressBar.Width - _
        2 * frmProgressBar.lblProgress.Left
    DoEvents
    
    For i = 1 To timePeriod_s
        Call Sleep(1000)
        
        frmProgressBar.lblTimeRemains.Caption = CStr(timePeriod_s - i) & " s"
        frmProgressBar.lblProgress.Width = _
            (frmProgressBar.Width - 2 * frmProgressBar.lblProgress.Left) * _
            (timePeriod_s - i) / timePeriod_s
        
        DoEvents
    Next i
    Call Unload(frmProgressBar)
End Sub

Public Function RandomRange&(lowerBound&, upperBound&)
    RandomRange = CLng((upperBound - lowerBound + 1) * Rnd + lowerBound)
End Function

Public Sub ScreenUpdatingOff()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

Public Sub ScreenUpdatingOn()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Public Function Convert_date(inputDate$) As Date
    Dim MyDate$
    
    inputDate = Replace(inputDate, ", ", " ")
    MyDate = Split(inputDate, " ")
    
    'month is always mydate(0)
    MyDate(0) = ""
    MyDate(0) = InStr("...JAN/FEB/MAR/APR/MAY/JUN/JUL/AUG/SEP/OCT/NOV/DEC", UCase(MyDate(0))) / 4
    
    'mydate(1) is day
    'mydate(2) is year
    
    Convert_date = CDate(Format(Join(MyDate, "/"), "dd/mm/yyyy"))
End Function

Public Sub Sorting()
    Dim sortingRange As Range
    Dim nRows&, nColumns&, sortedColumn&
    Dim strRange As String
    Dim Offset%
    Dim wsShares As Worksheet
    
    Set wsShares = ThisWorkbook.Worksheets(wsShares_Name)
    
    nRows = 1
    While wsShares.Cells(nRows + 2, 1) <> ""
        nRows = nRows + 1
    Wend
    nRows = nRows - 1
    
    Offset = Application.WorksheetFunction.Match("Sorting", wsShares.Range("A:A"), 0)
    
    sortedColumn = wsShares.Cells(Offset + wsShares.Cells(1, clmSortingIndex), 2)
    
    nColumns = wsShares.Cells(2, wsShares.Columns.Count).End(xlToLeft).Column
    
    strRange = wsShares.Range(wsShares.Cells(3, 1), wsShares.Cells(nRows + 2, nColumns)).Address(ReferenceStyle:=xlA1)
    
    wsShares.Sort.SortFields.Clear
    
    Set sortingRange = _
        Range(wsShares.Range(wsShares.Cells(3, sortedColumn), wsShares.Cells(nRows + 2, sortedColumn)).Address(ReferenceStyle:=xlA1))
    
    If wsShares.Cells(Offset + wsShares.Cells(1, clmSortingIndex), 3) Then
        wsShares.Sort.SortFields.Add Key:=sortingRange, _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    Else
        wsShares.Sort.SortFields.Add Key:=sortingRange, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End If
    
    With wsShares.Sort
        .SetRange Range(strRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    GetIdealTransactions
End Sub
