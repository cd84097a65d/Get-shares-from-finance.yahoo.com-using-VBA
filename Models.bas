Attribute VB_Name = "Models"
Option Explicit

Const RepresentationOffset& = 2
Const minInvestment = 2000
Const bankComission = 20
Const Threshold = 0.1

Dim wsSimulations As Worksheet
Dim wsTimeSeries As Worksheet
Dim wsShares As Worksheet

' estimates maximal performance of selected share in % and shows the points of buys/sales
' this was rather complicated and unstable
' this part is necessary for training of neural network
'
' the curve is splitted to buy/sell pairs that gives maximal performance
' i need a new algo to do this
' possible candidate: iterative procedure which tests the influence of removing of one
' buy/selling pair on the final performance
Function EstimateIdealTransactions#(dates() As Date, high#(), low#(), average#(), _
        buyingIndex&(), sellingIndex&(), performancePerTransaction#)
    Dim i&, j&, k&, m&, n&, indexMax&, indexMin&
    Dim sellingIndex2Remove&(), buyingIndex2Remove&()
    Dim tmpBool As Boolean
    Dim transactionsChanged As Boolean

    transactionsChanged = False
    ReDim buyingIndex(UBound(dates))
    ReDim sellingIndex(UBound(dates))

    ' initial set of transactions
    ' first transaction should be buy:
    For j = 1 To UBound(average)
        If average(j) <> Undefined Then
            For k = j + 1 To UBound(average)
                If average(k) <> Undefined Then
                    If average(j) < average(k) Then
                        buyingIndex(1) = j
                        sellingIndex(1) = k
                        GoTo Step2
                    End If
                    Exit For
                End If
            Next k
        End If
    Next j

Step2:
    For i = k + 1 To UBound(high) - 1
        If average(i) <> Undefined And average(i + 1) <> Undefined Then
            ' always a sequence of buy/sell, buy/sell ...
            If average(i) < average(i + 1) Then
                buyingIndex(i) = i
                sellingIndex(i + 1) = i + 1
            End If
        End If
    Next i

    transactionsChanged = True
    While transactionsChanged = True
Step3:
        transactionsChanged = False
        For i = 1 To UBound(buyingIndex) - 1
            If buyingIndex(i) <> 0 And buyingIndex(i + 1) = 0 Then
                If average(buyingIndex(i)) > average(buyingIndex(i) + 1) Then
                    buyingIndex(i) = 0
                    buyingIndex(i + 1) = buyingIndex(i + 1) + 1
                    transactionsChanged = True
                End If
            End If
        Next i

        For i = 1 To UBound(sellingIndex) - 1
            If sellingIndex(i) <> 0 And sellingIndex(i + 1) = 0 Then
                If average(sellingIndex(i)) < average(sellingIndex(i) + 1) Then
                    sellingIndex(i + 1) = sellingIndex(i) + 1
                    sellingIndex(i) = 0
                    transactionsChanged = True
                End If
            End If
        Next i
    Wend

Step4:
    ReDim sellingIndex2Remove(0)
    ReDim buyingIndex2Remove(0)
    For i = 1 To UBound(buyingIndex)
        For j = 1 To UBound(sellingIndex)
            If buyingIndex(i) = sellingIndex(j) And buyingIndex(i) <> 0 Then
                indexMax = buyingIndex(i - 1)
                indexMin = sellingIndex(j)
                For k = 0 To UBound(buyingIndex)
                    If j + k < UBound(sellingIndex) Then
                        If buyingIndex(i + k) = 0 Then
                            Exit For
                        End If

                        If buyingIndex(i + k) = sellingIndex(j + k) Then
                            If average(indexMax) > average(buyingIndex(i + k)) Then
                                If k > 0 Then
                                    ReDim Preserve buyingIndex2Remove(UBound(buyingIndex2Remove) + 1)
                                    buyingIndex2Remove(UBound(buyingIndex2Remove)) = indexMax
                                End If

                                indexMax = buyingIndex(i + k)
                            Else
                                ReDim Preserve buyingIndex2Remove(UBound(buyingIndex2Remove) + 1)
                                buyingIndex2Remove(UBound(buyingIndex2Remove)) = i + k
                            End If

                            If average(sellingIndex(j + k + 1)) > average(indexMin) Then
                                ReDim Preserve sellingIndex2Remove(UBound(sellingIndex2Remove) + 1)
                                sellingIndex2Remove(UBound(sellingIndex2Remove)) = indexMin

                                indexMin = sellingIndex(j + k + 1)
                            Else
                                ReDim Preserve sellingIndex2Remove(UBound(sellingIndex2Remove) + 1)
                                sellingIndex2Remove(UBound(sellingIndex2Remove)) = j + k + 1
                            End If
                        End If
                    End If
                Next k

                For k = 1 To UBound(buyingIndex2Remove)
                    buyingIndex(buyingIndex2Remove(k)) = 0
                Next k
                For k = 1 To UBound(sellingIndex2Remove)
                    sellingIndex(sellingIndex2Remove(k)) = 0
                Next k

                GoTo Step4
            End If
        Next j
    Next i

    Call CompactIndex(buyingIndex)
    Call CompactIndex(sellingIndex)

    ' loop over transactions
    transactionsChanged = True
    While transactionsChanged = True
Step5:
        transactionsChanged = False
        For i = 1 To UBound(buyingIndex) - 1
            If high(buyingIndex(i)) > low(sellingIndex(i)) Then
                buyingIndex(i) = 0
                sellingIndex(i) = 0
                transactionsChanged = True

                Call CompactIndex(buyingIndex)
                Call CompactIndex(sellingIndex)

                GoTo Step5
            End If




'            ' rule 1: first points are above the next (market is falling)
'            If average(buyingIndex(i)) > average(buyingIndex(i + 1)) And _
'                average(sellingIndex(i)) > average(sellingIndex(i + 1)) Then
'                buyingIndex(i) = 0
'                sellingIndex(i) = 0
'                transactionsChanged = True
'
'                Call CompactIndex(buyingIndex)
'                Call CompactIndex(sellingIndex)
'
'                GoTo Step5
'            End If
'
'            ' rule 2: if the profit is below the bank comission, then remove buy/sell transaction
'            If (average(sellingIndex(i)) - average(buyingIndex(i))) / average(buyingIndex(i)) < _
'                Threshold Then
'                buyingIndex(i) = 0
'                sellingIndex(i) = 0
'
'                Call CompactIndex(buyingIndex)
'                Call CompactIndex(sellingIndex)
'
'                transactionsChanged = True
'
'                GoTo Step5
'            End If
        Next i
    Wend

    ' final analysis to estimate performance
    EstimateIdealTransactions = 1
    performancePerTransaction = 0
    For i = 1 To UBound(buyingIndex)
        EstimateIdealTransactions = EstimateIdealTransactions * _
            (1# + (low(sellingIndex(i)) - high(buyingIndex(i))) / high(buyingIndex(i)))
        performancePerTransaction = performancePerTransaction + _
            (low(sellingIndex(i)) - high(buyingIndex(i))) / high(buyingIndex(i))
    Next i
    EstimateIdealTransactions = EstimateIdealTransactions - 1
    performancePerTransaction = performancePerTransaction / UBound(buyingIndex) / 2
End Function

' TODO: make new algo of min/max search!
Sub GetIdealTransactions()
    Dim rowRepresentation&, clmStart&, clmEnd&, i&, performancePerTransaction#
    Dim dates() As Date, average#(), high#(), low#()
    Dim buyingIndex&(), sellingIndex&()
    
    ScreenUpdatingOn
    
    Set wsShares = ThisWorkbook.Worksheets(wsShares_Name)
    
    ScreenUpdatingOff
    
    rowRepresentation = Application.WorksheetFunction.Match("Representation", _
        wsShares.Range("A:A"), 0) + RepresentationOffset
    
    clmStart = 19
    For clmStart = 19 To 16000
        If Not IsError(wsShares.Cells(rowRepresentation - 1, clmStart)) And _
            Not IsEmpty(wsShares.Cells(rowRepresentation - 1, clmStart)) Then
            Exit For
        End If
    Next clmStart
    
    For clmEnd = clmStart To 16000
        If IsEmpty(wsShares.Cells(rowRepresentation - 1, clmEnd)) Then
            clmEnd = clmEnd - 1
            Exit For
        End If
    Next clmEnd
    
    ReDim dates(clmEnd - clmStart + 1)
    ReDim high(clmEnd - clmStart + 1)
    ReDim low(clmEnd - clmStart + 1)
    ReDim average(clmEnd - clmStart + 1)
    
    For i = clmStart To clmEnd
        dates(i - clmStart + 1) = wsShares.Cells(rowRepresentation - 1, i)
        If IsNumeric(wsShares.Cells(rowRepresentation, i)) Then
            average(i - clmStart + 1) = wsShares.Cells(rowRepresentation, i)
        Else
            average(i - clmStart + 1) = Undefined
        End If
        If IsNumeric(wsShares.Cells(rowRepresentation + 2, i)) Then
            high(i - clmStart + 1) = wsShares.Cells(rowRepresentation + 2, i)
        Else
            high(i - clmStart + 1) = Undefined
        End If
        If IsNumeric(wsShares.Cells(rowRepresentation + 3, i)) Then
            low(i - clmStart + 1) = wsShares.Cells(rowRepresentation + 3, i)
        Else
            low(i - clmStart + 1) = Undefined
        End If
    Next i
    
    wsShares.Range(wsShares.Cells(rowRepresentation + 5, clmStart), _
        wsShares.Cells(rowRepresentation + 6, clmEnd)).ClearContents
    
    wsShares.Cells(rowRepresentation - 1, 7) = _
        EstimateIdealTransactions(dates, high, low, average, buyingIndex, sellingIndex, _
        performancePerTransaction)
    wsShares.Cells(rowRepresentation - 1, 6) = performancePerTransaction
    
    ' todo: cleanup
    For i = 1 To UBound(buyingIndex)
        wsShares.Cells(rowRepresentation + 5, buyingIndex(i) + clmStart - 1) = _
            wsShares.Cells(rowRepresentation, buyingIndex(i) + clmStart - 1)
    Next i
    For i = 1 To UBound(buyingIndex)
        wsShares.Cells(rowRepresentation + 6, sellingIndex(i) + clmStart - 1) = _
            wsShares.Cells(rowRepresentation, sellingIndex(i) + clmStart - 1)
    Next i
    
    ScreenUpdatingOn
End Sub

Sub CompactIndex(sellingIndex&())
    Dim i&, counter&
    
    counter = 0
    For i = 1 To UBound(sellingIndex)
        If sellingIndex(i) <> 0 Then
            counter = counter + 1
            sellingIndex(counter) = sellingIndex(i)
        End If
    Next i
    
    ReDim Preserve sellingIndex(counter)
End Sub

Sub Model()
    Dim ihStart&, ihEnd&, invHorizon&, day_&
    Dim open_#(), high_#(), low_#(), close_#(), average_#()
    Dim totalFreeMoney#, nShares&, totalMOneyInShares#
    Dim ter#, minInvest#, maxInvest#
    Dim tickers$(), tickers_sorted$(), performanceArray#()
    Dim i&, j&, k&, nDays&, currentDay&, tickerCanBeBought As Boolean
    Dim days() As Date
    Dim portfolio As New Dictionary
    Dim tmpShare As New Share
    Dim adaptivePerformanceLimit#, nTransactions&
    Dim item
    Dim listToRemove$()
    
    Set wsSimulations = ThisWorkbook.Worksheets("Simulations")
    Set wsTimeSeries = ThisWorkbook.Worksheets("TimeSeries")
    
    ' overtake settings
    ter = wsSimulations.Cells(3, 2)
    ihStart = wsSimulations.Cells(4, 2)
    ihEnd = wsSimulations.Cells(5, 2)
    minInvest = wsSimulations.Cells(6, 2)
    maxInvest = wsSimulations.Cells(7, 2)
    
    ' clear previous results
    i = wsSimulations.Cells(16, wsSimulations.Columns.Count).End(xlToLeft).Column
    wsSimulations.Range(wsSimulations.Cells(15, 1), wsSimulations.Cells(15, i)).ClearContents
    For j = 1 To i
        k = wsSimulations.Cells(wsSimulations.Rows.Count, j).End(xlUp).row
        If k > 14 Then
            wsSimulations.Range(wsSimulations.Cells(16, j), wsSimulations.Cells(k, j)).ClearContents
        End If
    Next j
    
    ScreenUpdatingOff
    
    ' todo: if favorites = -1 then remove these shares from consideration
    
    ' overtake and clean up all data from "TimeSeries"
    Call OvertakeAndCleanup(tickers, days, open_, high_, low_, close_, average_)
    nShares = UBound(tickers)
    nDays = UBound(days)
    ReDim tickers_sorted(nShares)
    ReDim performanceArray(nShares)
    
    frmProgressBar.Show
    frmProgressBar.Caption = "In progress..."
    
    frmProgressBar.lblTimeRemains.Caption = "0%"
    frmProgressBar.lblProgress.Width = frmProgressBar.Width - _
        2 * frmProgressBar.lblProgress.Left
    DoEvents
    
    Dim nSteps&, currentStep&
    
    ' i'm too lazy to derive the formula :-)
    nSteps = 0
    For invHorizon = ihStart To ihEnd
        ' loop over days
        For day_ = 1 + invHorizon To nDays
            nSteps = nSteps + 1
        Next day_
    Next invHorizon
    
    ' loop over investment horizons
    ' todo: add ProgressBar
    currentStep = 0
    For invHorizon = ihStart To ihEnd
        Set portfolio = CreateObject("Scripting.Dictionary")
        totalFreeMoney = wsSimulations.Cells(2, 2)
        nTransactions = 0
        
        ' loop over days
        For day_ = 1 + invHorizon To nDays
            ' get performance for selected day and investment horizon
            Call GetPerformance(day_, invHorizon, average_, performanceArray)
            
            ' get sorted tickers according to the performance inside the last invest. horizon
            Call GetSortedTickersInsideHorizon(tickers, performanceArray, tickers_sorted)
            
            ' buy/sell shares according to sorted performance
            ' sell
            ReDim listToRemove(0)
            For i = 1 To portfolio.Count
                If GetIndex(portfolio.Items(i - 1).ticker, tickers_sorted) > 33 Then
                    If low_(GetIndex(portfolio.Items(i - 1).ticker, tickers), day_) > 0 Then
                        ReDim Preserve listToRemove(UBound(listToRemove) + 1)
                        
                        totalFreeMoney = totalFreeMoney + _
                            low_(GetIndex(portfolio.Items(i - 1).ticker, tickers), day_) * _
                            portfolio.Items(i - 1).Number
                        
                        wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 1) = _
                            "Sell"
                        wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 2) = _
                            portfolio.Items(i - 1).ticker
                        wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 3) = _
                            days(day_)
                        wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 4) = _
                            low_(GetIndex(portfolio.Items(i - 1).ticker, tickers), day_)
                        wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 5) = _
                            portfolio.Items(i - 1).Number
                        
                        nTransactions = nTransactions + 1
                        listToRemove(UBound(listToRemove)) = portfolio.Keys(i - 1)
                    End If
                End If
            Next i
            
            Call RemoveFromFromPortfolio(portfolio, listToRemove)
            
            ' buy only if the market performance is positive for at least of the half of shares
            If NumberOfSharesWithPositivePerformance&(performanceArray) > nShares / 2 Then
                ' buy
                For i = 1 To UBound(performanceArray)
                    tickerCanBeBought = True
                    For Each item In portfolio.Items
                        If item.ticker = tickers_sorted(i) Then
                            If item.Date_ + invHorizon < days(day_) Then
                                tickerCanBeBought = False
                            End If
                        End If
                    Next item
                    
                    If tickerCanBeBought Then
                        ' cannot buy share without money
                        If totalFreeMoney > minInvest And performanceArray(i) > 0.01 Then
                            Set tmpShare = New Share
                            tmpShare.ticker = tickers_sorted(i)
                            tmpShare.Price = high_(GetIndex(tmpShare.ticker, tickers), day_)
                            tmpShare.Number = Int(minInvest / tmpShare.Price)
                            tmpShare.Date_ = days(day_)
                            totalFreeMoney = totalFreeMoney - tmpShare.Number * tmpShare.Price
                            ' nTransactions is needed only as a unique descriptor
                            Call portfolio.Add(CStr(nTransactions), tmpShare)
                            
                            wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 1) = "Buy"
                            wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 2) = tmpShare.ticker
                            wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 3) = tmpShare.Date_
                            wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 4) = tmpShare.Price
                            wsSimulations.Cells(16 + nTransactions, (invHorizon - 1) * 5 + 5) = tmpShare.Number
                            
                            nTransactions = nTransactions + 1
                        End If
                    End If
                Next i
            End If
            
            currentStep = currentStep + 1
            Call ProgressBar_Percent(Round(100# * currentStep / nSteps, 0))
        Next day_
        
        For i = 1 To portfolio.Count
            portfolio.Items(i - 1).Price = _
                low_(GetIndex(tickers(i), tickers), UBound(days))
            totalFreeMoney = totalFreeMoney + portfolio.Items(i - 1).Number * _
                portfolio.Items(i - 1).Price
        Next i
        wsSimulations.Cells(15, (invHorizon - 1) * 5 + 1) = _
            (totalFreeMoney / wsSimulations.Cells(2, 2)) - 1#
        
        Set portfolio = Nothing
    Next invHorizon
    
    Call Unload(frmProgressBar)
    
    ScreenUpdatingOn
End Sub

Function NumberOfSharesWithPositivePerformance&(performanceArray#())
    Dim i&
    
    For i = 1 To UBound(performanceArray)
        If performanceArray(i) <= 0 Then
            NumberOfSharesWithPositivePerformance& = i - 1
            Exit Function
        End If
    Next i
End Function

' overtake and clean up all data from "TimeSeries"
Sub OvertakeAndCleanup(tickers$(), days() As Date, open_#(), high_#(), low_#(), close_#(), average_#())
    Dim currentDay&, i&, j&, nShares&, nDays&
    
    ReDim tickers(10000)
    nShares = 3
    While wsTimeSeries.Cells(nShares, 1) <> ""
        tickers(nShares - 2) = wsTimeSeries.Cells(nShares, 1)
        nShares = nShares + 1
    Wend
    nShares = nShares - 3
    ReDim Preserve tickers(nShares)
    
    ReDim days(clmData_Open_End)
    currentDay = 0
    For i = clmData_Open_End - 5 To 1 Step -1   ' 5 means that only 255 from 260 dates are analyzed (this is true, because the year contains max 253 working days)
        ' check that 2/3 of data for current date are valid
        ' othervise this date will be not taken in consideration:
        j = Application.WorksheetFunction.Count( _
            wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Open_End + 1 - i), _
            wsTimeSeries.Cells(2 + nShares, clmData_Open_End + 1 - i)))
        
        If j > 2 * nShares / 3 Then
            currentDay = currentDay + 1
            days(currentDay) = wsTimeSeries.Cells(2, clmData_Open_End + 1 - i)
        End If
    Next i
    
    nDays = currentDay
    ReDim Preserve days(nDays)
    ReDim open_#(nShares, nDays): ReDim high_#(nShares, nDays)
    ReDim low_#(nShares, nDays): ReDim close_#(nShares, nDays)
    ReDim average_#(nShares, nDays)
    
    currentDay = 0
    For i = clmData_Open_End - 5 To 1 Step -1   ' 5 means that only 255 from 260 dates are analyzed (this is true, because the year contains max 253 working days)
        ' check that 2/3 of data for current date are valid
        ' othervise this date will be not taken in consideration:
        j = Application.WorksheetFunction.Count( _
            wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Open_End + 1 - i), _
            wsTimeSeries.Cells(2 + nShares, clmData_Open_End + 1 - i)) _
        )
        
        If j > 2 * nShares / 3 Then
            currentDay = currentDay + 1
            For j = 1 To nShares
                If wsTimeSeries.Cells(j + 2, clmData_Open_End + 1 - i) = "" Then
                    open_(j, currentDay) = Undefined
                Else
                    open_(j, currentDay) = wsTimeSeries.Cells(j + 2, clmData_Open_End + 1 - i)
                End If
                If wsTimeSeries.Cells(j + 2, clmData_High_End + 1 - i) = "" Then
                    high_(j, currentDay) = Undefined
                Else
                    high_(j, currentDay) = wsTimeSeries.Cells(j + 2, clmData_High_End + 1 - i)
                End If
                If wsTimeSeries.Cells(j + 2, clmData_Low_End + 1 - i) = "" Then
                    low_(j, currentDay) = Undefined
                Else
                    low_(j, currentDay) = wsTimeSeries.Cells(j + 2, clmData_Low_End + 1 - i)
                End If
                If wsTimeSeries.Cells(j + 2, clmData_Close_End + 1 - i) = "" Then
                    close_(j, currentDay) = Undefined
                Else
                    close_(j, currentDay) = wsTimeSeries.Cells(j + 2, clmData_Close_End + 1 - i)
                End If
                If wsTimeSeries.Cells(j + 2, clmData_Average_End + 1 - i) = "" Then
                    average_(j, currentDay) = Undefined
                Else
                    average_(j, currentDay) = wsTimeSeries.Cells(j + 2, clmData_Average_End + 1 - i)
                End If
            Next j
        End If
    Next i
End Sub

Sub RemoveFromFromPortfolio(portfolio As Dictionary, listToRemove$())
    Dim i&
    
    For i = 1 To UBound(listToRemove)
        portfolio.Remove (listToRemove(i))
    Next i
End Sub

Function GetIndex(ticker, tickers$())
    For GetIndex = 1 To UBound(tickers)
        If ticker = tickers(GetIndex) Then
            Exit Function
        End If
    Next GetIndex
End Function

Sub GetPerformance(day_&, invHorizon&, average#(), performanceArray#())
    Dim i&
    
    For i = 1 To UBound(performanceArray)
        If average(i, day_) = Undefined Or average(i, day_ - invHorizon) = Undefined Then
            performanceArray(i) = 0#
        Else
            performanceArray(i) = Round((average(i, day_) - average(i, day_ - invHorizon)) / _
                average(i, day_ - invHorizon), 5)
        End If
    Next i
End Sub

' simple bubble sorting procedure:
Sub GetSortedTickersInsideHorizon(tickers$(), performanceArray#(), tickers_sorted$())
    Dim n&, i&, j&, newN&, tmpDbl#, tmpString$
    
    n = UBound(tickers)
    
    For i = 1 To n
        tickers_sorted(i) = tickers(i)
    Next i
    
    For j = 1 To n
        For i = 2 To n
            If performanceArray(i - 1) < performanceArray(i) Then
                tmpDbl = performanceArray(i)
                performanceArray(i) = performanceArray(i - 1)
                performanceArray(i - 1) = tmpDbl
                
                tmpString = tickers_sorted(i)
                tickers_sorted(i) = tickers_sorted(i - 1)
                tickers_sorted(i - 1) = tmpString
            End If
        Next i
    Next j
End Sub

Public Sub ProgressBar_Percent(prc#)
    frmProgressBar.lblTimeRemains.Caption = CStr(prc) & "%"
        frmProgressBar.lblProgress.Width = _
            (frmProgressBar.Width - 2 * frmProgressBar.lblProgress.Left) * prc / 100#
    DoEvents
End Sub

