Attribute VB_Name = "Currencies"
Option Explicit

Const wsCurrenciesName$ = "Currencies"

Private currencyName$(), currencyRate#()
Private currenciesAreUpdated As Boolean
Private wsCurrencies As Worksheet

Sub GetCurrencies()
    Dim objRequest As Variant
    Dim resultFromYahoo$, tickerURL$, i&
    
    Set wsCurrencies = ThisWorkbook.Worksheets(wsCurrenciesName)
    
    If Int(DateTime.Now) = ThisWorkbook.Worksheets("Currencies").Cells(1, 3) Then
        Exit Sub
    End If
    
    getCookieCrumb
    
    i = 2
    While wsCurrencies.Cells(i, 1) <> ""
        tickerURL = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/EUR" & _
            wsCurrencies.Cells(i, 1) & "%3DX?modules=price"
        
        Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        With objRequest
            .Open "GET", tickerURL, False
            .send
            .waitForResponse
            resultFromYahoo = .ResponseText
        End With
        
        resultFromYahoo = Split(resultFromYahoo, """regularMarketPrice"":{""raw"":")(1)
        resultFromYahoo = Split(resultFromYahoo, ",""fmt"":")(0)
        wsCurrencies.Cells(i, 2) = resultFromYahoo
        
        i = i + 1
    Wend
    
    currenciesAreUpdated = False
    wsCurrencies.Cells(1, 3) = Int(DateTime.Now)
End Sub

Public Function GetCurrency#(currency_$)
    Dim i&, j&, k&
    Dim tmpDbl#
    
    GetCurrency = 0#
    
    If Not currenciesAreUpdated Then
        k = wsCurrencies.Range("A" & wsCurrencies.Rows.Count).End(xlUp).row
        GetCurrencies
        j = 0
        For i = 1 To k
            If wsCurrencies.Cells(i, 1) <> "" Then
                j = j + 1
                ReDim Preserve currencyName(j)
                ReDim Preserve currencyRate(j)
                currencyName(j) = wsCurrencies.Cells(i, 1)
                currencyRate(j) = wsCurrencies.Cells(i, 2)
            End If
        Next i
        
        currenciesAreUpdated = True
    End If
    
    For i = 1 To UBound(currencyName)
        If currencyName(i) = currency_ Then
            tmpDbl = currencyRate(i)
            Exit For
        End If
    Next i
    GetCurrency = tmpDbl
End Function
