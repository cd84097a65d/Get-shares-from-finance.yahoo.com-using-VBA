Attribute VB_Name = "Shares"
Option Explicit

' important links:
' https://stackoverflow.com/questions/44030983/yahoo-finance-url-not-working/44050039
' Samir Khan: http://investexcel.net/multiple-stock-quote-downloader-for-excel/

' TODOs

Const clmName& = 1
Const clmTicker& = 2
'Const clmFav& = 3
Const clmCountry& = 4
Const clmSector& = 5
Const clmIndustry& = 6
Const clmPrice& = 7
Const clmCurrency& = 8
Const clmCapitalization& = 9
Const clmPrice_Euro& = 10
Const clmTimeStamp& = 11
Const clmProfile& = 12
Public Const clmData_Open_Start& = 260
Public Const clmData_High_Start& = 520
Public Const clmData_Low_Start& = 780
Public Const clmData_Close_Start& = 1040
Public Const clmData_Average_Start& = 1300
Public Const Offset = 2     ' row offset of data in sheets "Shares" and "TimeSeries"

Const sort_1_Criterium& = 1 ' 23 days
Const sort_2_Criterium& = 2 ' 46 days
Const sort_3_Criterium& = 3 ' 69 days
Const sort_3m& = 4          ' 3 m
Const sort_6m& = 5          ' 6 m
Const sort_9m& = 6          ' 9 m
Const sort_timeStamp& = 7   ' Time stamp


Public Const wsShares_Name$ = "Shares"
Const wsTimeSeries_Name$ = "TimeSeries"

Dim wsShares As Worksheet
Dim wsTimeSeries As Worksheet

Sub GetProfile(ticker$, Name$, industry$, sector$, _
               profile$, currency_$, country$)
    Dim objRequest As Variant
    Dim tickerURL$
    Dim resultFromYahoo$
    
    Call getCookieCrumb
    
    tickerURL = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/" & ticker & _
        "?modules=summaryProfile%2CsummaryDetail%2Cprice"
    
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objRequest
        .Open "GET", tickerURL, False
        .setRequestHeader "Cookie", Samir_Khan.cookie
        .send
        .waitForResponse
        resultFromYahoo = .responseText
    End With
    
    ' Worksheets("Lapa2").Cells(6, 4) = resultFromYahoo
    ' name
    Name = ""
    If InStr(1, resultFromYahoo, "shortName") > 0 Then
        Name = Split(Split(resultFromYahoo, """shortName"":""")(1), """,""")(0)
    End If
    If InStr(1, resultFromYahoo, "longName") > 0 And Name = "" Then
        Name = Split(Split(resultFromYahoo, """longName"":""")(1), """,""")(0)
    End If
    ' sector
    sector = Split(Split(resultFromYahoo, """sector"":""")(1), """,""")(0)
    ' industry
    industry = Split(Split(resultFromYahoo, """industry"":""")(1), """,""")(0)
    ' price, euro
    ' currency
    currency_ = Split(Split(resultFromYahoo, """currency"":""")(1), """,""")(0)
    ' profile
    profile = Split(Split(resultFromYahoo, """longBusinessSummary"":""")(1), """,""")(0)
    ' country
    country = Split(Split(resultFromYahoo, """country"":""")(1), """,""")(0)
End Sub

Sub UpdateSelected()
    Dim period1$, period2$
    Dim outDates_reference() As Date    ' reference dates from AAPL (why? see comments below)
    Dim outDates() As Date
    Dim outTimeSeries#()
    Dim i&
    Dim ticker$, Name$, profile$, sector$
    Dim currency_$, Price#, country$, capitalization#, industry$
    
    Set wsShares = Worksheets(wsShares_Name)
    Set wsTimeSeries = Worksheets(wsTimeSeries_Name)
    
    getCookieCrumb
    
    GetCurrencies
    
    period2 = CStr((Int(DateTime.Now) - CDate("01.01.1970")) * 60 * 60 * 24)
    period1 = CStr((Int(DateTime.Now - 365) - CDate("01.01.1970")) * 60 * 60 * 24)
    
    ' problem: different markets/countries - different holidays
    ' lazy solution on getting the dates:
    ' I would take the apple (AAPL) and update the table with its dates
    ' this must be enough: the handle with apple is very frequent
    Call getYahooFinanceData("AAPL", period1, period2, "1d", outDates_reference, outTimeSeries)
    Call UpdateDates(outDates_reference)
    
    i = wsShares.Cells(1, 2)
    
    ticker = wsShares.Cells(i + Offset, clmTicker)
    Name = wsShares.Cells(i + Offset, clmName)
    currency_ = wsShares.Cells(i + Offset, clmCurrency)
    
    If wsShares.Cells(i + Offset, clmName) = "" Then
        Call GetProfile(ticker, Name, industry, sector, profile, currency_, country)
        wsShares.Cells(i + Offset, clmName) = Split(Name, " (")(0)
        wsShares.Cells(i + Offset, clmProfile) = profile
        wsShares.Cells(i + Offset, clmSector) = sector
        wsShares.Cells(i + Offset, clmIndustry) = industry
        wsShares.Cells(i + Offset, clmCurrency) = currency_
        wsShares.Cells(i + Offset, clmCountry) = country
    End If
    Application.StatusBar = CStr(i + Offset) & ", " & Name
    
    Call GetPrice(ticker, Price, capitalization)
    wsShares.Cells(i + Offset, clmCapitalization) = capitalization
    wsShares.Cells(i + Offset, clmPrice) = Price
    wsShares.Cells(i + Offset, clmPrice_Euro) = Price / GetCurrency(currency_)
    
    Call getYahooFinanceData(ticker, period1, period2, "1d", outDates, outTimeSeries)
    Call UpdateSheets(outDates_reference, outDates, outTimeSeries, ticker)
    
    wsShares.Cells(i + Offset, clmTimeStamp) = DateTime.Now
    Application.StatusBar = ""
    GetIdealTransactions
End Sub

Sub GetPrice(ticker$, Price#, capitalization#)
    Dim objRequest
    Dim tickerURL$
    Dim resultFromYahoo$, price_$, capitalization_$
    
    Call getCookieCrumb
    
    tickerURL = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/" & ticker & _
        "?modules=price"
    
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objRequest
        .Open "GET", tickerURL, False
        .setRequestHeader "Cookie", Samir_Khan.cookie
        .send
        .waitForResponse
        resultFromYahoo = .responseText
    End With
    
    ' price
    price_ = Split(Split(resultFromYahoo, """regularMarketPrice"":{""raw"":")(1), ",""")(0)
    Price = Val(price_)
    capitalization = Val(Split(Split(resultFromYahoo, """marketCap"":{""raw"":")(1), ",""")(0))
    capitalization_ = Split(resultFromYahoo, """marketCap"":{""raw"":")(1)
    capitalization_ = Split(capitalization_, """longFmt"":""")(1)
    capitalization_ = Split(capitalization_, """")(0)
    capitalization_ = Replace(capitalization_, ",", "")
    capitalization = Val(capitalization_)
End Sub

Sub Update_Shares()
    Dim period1$, period2$
    Dim outDates_reference() As Date    ' reference dates from AAPL (why? see comments below)
    Dim outDates() As Date
    Dim outTimeSeries#()
    
    Set wsShares = Worksheets(wsShares_Name)
    Set wsTimeSeries = Worksheets(wsTimeSeries_Name)
    
    getCookieCrumb
    
    GetCurrencies
    
    period2 = CStr((Int(DateTime.Now) - CDate("01.01.1970")) * 60 * 60 * 24)
    period1 = CStr((Int(DateTime.Now - 365) - CDate("01.01.1970")) * 60 * 60 * 24)
    
    ' problem: different markets/countries - different holidays
    ' lazy solution on getting the dates:
    ' I would take the apple (AAPL) and update the table with its dates
    ' this must be enough: the handle with apple is very frequent
    Call getYahooFinanceData("AAPL", period1, period2, "1d", outDates_reference, outTimeSeries)
    Call UpdateDates(outDates_reference)
    
    ' update shares sorted by first criterium
    Call GetSetsOfShares(sort_1_Criterium, period1, period2, outDates_reference)
    
    ' update shares sorted by second criterium
    Call GetSetsOfShares(sort_2_Criterium, period1, period2, outDates_reference)
    
    ' update shares sorted by third criterium
    Call GetSetsOfShares(sort_3_Criterium, period1, period2, outDates_reference)
    
    ' update shares sorted by 3 months
    Call GetSetsOfShares(sort_3m, period1, period2, outDates_reference)
    
    ' update shares sorted by 6 months
    Call GetSetsOfShares(sort_6m, period1, period2, outDates_reference)
    
    ' update shares sorted by 9 months
    Call GetSetsOfShares(sort_9m, period1, period2, outDates_reference)
    
    ' update shares sorted by time stamp
    Call GetSetsOfShares(sort_timeStamp, period1, period2, outDates_reference)
    
    ' at the end sort using the first criterium
    wsShares.Cells(1, clmSortingIndex) = sort_1_Criterium
    Sorting
    
    Application.StatusBar = ""
End Sub

' sort according to "sortingCriterium" and update the first 20 shares (20 is read from "A1")
' by updating of 20 first shares in 7 different sortings 65 shares were updated
Sub GetSetsOfShares(sortingCriterium&, period1$, period2$, outDates_reference() As Date)
    Dim i&, nToUpdate&
    Dim ticker$, Name$, profile$, sector$
    Dim currency_$, Price#, country$, capitalization#, industry$
    Dim outDates() As Date
    Dim outTimeSeries#()
    
    wsShares.Cells(1, clmSortingIndex) = sortingCriterium
    Sorting
    
    nToUpdate = wsShares.Cells(1, 1)
    
    For i = 1 To nToUpdate
        If Int(wsShares.Cells(i + Offset, clmTimeStamp)) < DateTime.Date Then
            ticker = wsShares.Cells(i + Offset, clmTicker)
            Name = wsShares.Cells(i + Offset, clmName)
            currency_ = wsShares.Cells(i + Offset, clmCurrency)
            
            If wsShares.Cells(i + Offset, clmName) = "" Then
                Call GetProfile(ticker, Name, industry, sector, profile, currency_, country)
                wsShares.Cells(i + Offset, clmName) = Split(Name, " (")(0)
                wsShares.Cells(i + Offset, clmProfile) = profile
                wsShares.Cells(i + Offset, clmSector) = sector
                wsShares.Cells(i + Offset, clmIndustry) = industry
                wsShares.Cells(i + Offset, clmCurrency) = currency_
                wsShares.Cells(i + Offset, clmCountry) = country
            End If
            Application.StatusBar = CStr(i + Offset) & ", " & Name
            
            Call GetPrice(ticker, Price, capitalization)
            wsShares.Cells(i + Offset, clmCapitalization) = capitalization
            wsShares.Cells(i + Offset, clmPrice) = Price
            wsShares.Cells(i + Offset, clmPrice_Euro) = Price / GetCurrency(currency_)
            
            Call getYahooFinanceData(ticker, period1, period2, "1d", outDates, outTimeSeries)
            Call UpdateSheets(outDates_reference, outDates, outTimeSeries, ticker)
            
            Call ProgressBar(Name)
            
            wsShares.Cells(i + Offset, clmTimeStamp) = DateTime.Now
        End If
    Next i
End Sub

Sub UpdateDates(outDates_reference() As Date)
    Dim i&, nColumns&, shift&, lastRow&
    
    Application.ScreenUpdating = False
    i = clmData_Open_Start
    While wsTimeSeries.Cells(2, i) <> outDates_reference(1)
        i = i - 1
    Wend
    
    shift = i - (clmData_Open_Start - UBound(outDates_reference) + 1)
    lastRow = wsTimeSeries.Cells(wsTimeSeries.Rows.Count, clmData_Open_Start).End(xlUp).Row
    
    If shift <> 0 Then
        wsTimeSeries.Range(wsTimeSeries.Cells(3, shift + (clmData_Open_Start + 1 - UBound(outDates_reference))), _
            wsTimeSeries.Cells(lastRow, clmData_Open_Start)).Copy
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Open_Start - UBound(outDates_reference) + 1), _
            wsTimeSeries.Cells(3, clmData_Open_Start - UBound(outDates_reference) + 1)).PasteSpecial xlPasteValues
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Open_Start - shift + 1), _
            wsTimeSeries.Cells(lastRow, clmData_Open_Start)).ClearContents
        
        wsTimeSeries.Range(wsTimeSeries.Cells(3, shift + (clmData_High_Start + 1 - UBound(outDates_reference))), _
            wsTimeSeries.Cells(lastRow, clmData_High_Start)).Copy
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_High_Start - UBound(outDates_reference) + 1), _
            wsTimeSeries.Cells(3, clmData_High_Start - UBound(outDates_reference) + 1)).PasteSpecial xlPasteValues
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_High_Start - shift + 1), _
            wsTimeSeries.Cells(lastRow, clmData_High_Start)).ClearContents
        
        wsTimeSeries.Range(wsTimeSeries.Cells(3, shift + (clmData_Low_Start + 1 - UBound(outDates_reference))), _
            wsTimeSeries.Cells(lastRow, clmData_Low_Start)).Copy
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Low_Start - UBound(outDates_reference) + 1), _
            wsTimeSeries.Cells(3, clmData_Low_Start - UBound(outDates_reference) + 1)).PasteSpecial xlPasteValues
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Low_Start - shift + 1), _
            wsTimeSeries.Cells(lastRow, clmData_Low_Start)).ClearContents
        
        wsTimeSeries.Range(wsTimeSeries.Cells(3, shift + (clmData_Close_Start + 1 - UBound(outDates_reference))), _
            wsTimeSeries.Cells(lastRow, clmData_Close_Start)).Copy
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Close_Start - UBound(outDates_reference) + 1), _
            wsTimeSeries.Cells(3, clmData_Close_Start - UBound(outDates_reference) + 1)).PasteSpecial xlPasteValues
        wsTimeSeries.Range(wsTimeSeries.Cells(3, clmData_Close_Start - shift + 1), _
            wsTimeSeries.Cells(lastRow, clmData_Close_Start)).ClearContents
        
        nColumns = UBound(outDates_reference)
        For i = 1 To UBound(outDates_reference)
            wsTimeSeries.Cells(2, clmData_Open_Start - nColumns + i) = outDates_reference(i)
            wsTimeSeries.Cells(2, clmData_High_Start - nColumns + i) = outDates_reference(i)
            wsTimeSeries.Cells(2, clmData_Low_Start - nColumns + i) = outDates_reference(i)
            wsTimeSeries.Cells(2, clmData_Close_Start - nColumns + i) = outDates_reference(i)
        Next i
    End If
    Application.ScreenUpdating = True
End Sub

Sub UpdateSheets(outDates_reference() As Date, outDates() As Date, outTimeSeries#(), ticker$)
    Dim i&, j&, nColumns&, timeSeriesLine&
    
    Application.ScreenUpdating = False
    nColumns = UBound(outDates_reference)
    
    ' find the line in wsTimeSeries that corresponds to the ticker
    timeSeriesLine = 3
    While wsTimeSeries.Cells(timeSeriesLine, 1) <> ticker
        timeSeriesLine = timeSeriesLine + 1
    Wend
    
    For i = 1 To UBound(outDates)
        ' search for
        For j = 1 To UBound(outDates_reference)
            ' look for the same dates:
            If outDates(i) = outDates_reference(j) Then
                If outTimeSeries(i, 1) = Undefined Then
                    wsTimeSeries.Cells(timeSeriesLine, clmData_Open_Start - nColumns + j) = ""
                Else
                    wsTimeSeries.Cells(timeSeriesLine, clmData_Open_Start - nColumns + j) = _
                        outTimeSeries(i, 1)
                End If
                If outTimeSeries(i, 2) = Undefined Then
                    wsTimeSeries.Cells(timeSeriesLine, clmData_High_Start - nColumns + j) = ""
                Else
                    wsTimeSeries.Cells(timeSeriesLine, clmData_High_Start - nColumns + j) = _
                        outTimeSeries(i, 2)
                End If
                If outTimeSeries(i, 3) = Undefined Then
                    wsTimeSeries.Cells(timeSeriesLine, clmData_Low_Start - nColumns + j) = ""
                Else
                    wsTimeSeries.Cells(timeSeriesLine, clmData_Low_Start - nColumns + j) = _
                        outTimeSeries(i, 3)
                End If
                If outTimeSeries(i, 4) = Undefined Then
                    wsTimeSeries.Cells(timeSeriesLine, clmData_Close_Start - nColumns + j) = ""
                Else
                    wsTimeSeries.Cells(timeSeriesLine, clmData_Close_Start - nColumns + j) = _
                        outTimeSeries(i, 4)
                End If
                Exit For
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub OpenProfileInfo()
    ProfileInfo.lblProfileInfo = Worksheets(wsShares_Name).Cells( _
                                 2 + Worksheets(wsShares_Name).Cells(1, 2), clmProfile)
    ProfileInfo.Show
End Sub
