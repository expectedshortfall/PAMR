VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MarketStateFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pMemoryLimit As Integer

'=============================================================
'
'   METHODS
'
'=============================================================

Public Function LoadMarketStateProvider(inDate As Date) As MarketStateProvider
    
    Set LoadMarketStateProvider = Factory.CreateMarketStateProvider(inDate)
   
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
        Dim wbRates As Workbook, wbRef As Workbook
        Dim wsRates As Worksheet, wsRef As Worksheet
        Dim tmpMarketState As MarketState
    
        Set wbRates = Workbooks.Open(ThisWorkbook.Path + "\rates.xls", ReadOnly:=True)
        Set wsRates = wbRates.Worksheets(1)
        
        Set wbRef = Workbooks.Open(ThisWorkbook.Path + "\ref.xls", ReadOnly:=True)
        Set wsRef = wbRef.Worksheets(1)
            
        Dim i As Integer: i = FindIndexForDate(inDate, wsRates)
        Dim counter As Integer
        Dim currentDate As Date
        
        While (Not wsRates.Cells(i, 1).Value = Empty) And (counter <= pMemoryLimit)
            
            currentDate = wsRates.Cells(i, 1).Value
            
            Set tmpMarketState = Factory.CreateMarketSate(currentDate)
            With tmpMarketState
                Dim j As Integer: j = 4 'start reading from 4th column
                
                .AddCcyRate CCY.PLN, 1                         'ADD EURPLN
                .AddCcyRate CCY.EUR, wsRates.Cells(i, 2).Value 'ADD EURPLN
                .AddCcyRate CCY.USD, wsRates.Cells(i, 3).Value 'ADD USDPLN
                
                While Not wsRates.Cells(i, j).Value = Empty
                    .AddCurve GetCurve(wsRates, i, j)
                Wend
                               
                'Get reference rates only if there were rates in rates.xls for the same date
                If Not IsError(FindIndexForDate(currentDate, wsRef)) Then
                    
                    Dim K As Integer: K = FindIndexForDate(currentDate, wsRef)
                    
                    j = 2 'start reading from second columns
                    
                    While Not wsRef.Cells(K, j).Value = Empty
                        .AddCurve GetCurve(wsRef, K, j, False)
                    Wend
                
                End If
            End With
            
            LoadMarketStateProvider.AddMarketState tmpMarketState
            i = i + 1
            counter = counter + 1
        Wend

        wbRates.Close False: Set wbRates = Nothing
        wbRef.Close False: Set wbRef = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
            
End Function

'=============================================================
'
'   PRIVATE METHODS
'
'=============================================================
Private Function GetCurve(ws As Worksheet, row As Integer, ByRef column As Integer, Optional isRateFile As Boolean = True) As Curve
        
        Dim curveName As String: curveName = GetRatesName(ws.Cells(1, column).Value, isRateFile)
        Dim i As Integer
        
        Set GetCurve = Factory.CreateCurve(curveName, ws.Cells(row, 1))
        While curveName = GetRatesName(ws.Cells(1, column + i).Value, isRateFile)
            
            Dim newPeriod As Period
            Dim newRate As Double
            
            newPeriod = StringToPeriod(ws.Cells(1, column + i).Value)
            newRate = ws.Cells(row, column + i).Value
            
            GetCurve.AddRate newPeriod, newRate
            
            i = i + 1
        
        Wend
        
        column = column + i
        
End Function

Private Function GetRatesName(inString As String, Optional isRateFile As Boolean = True) As String
    If isRateFile Then
        GetRatesName = Left(inString, 3)
    Else
        If inString = "" Then
            GetRatesName = ""
        Else
            GetRatesName = Left(inString, Len(inString) - 2)
        End If
    End If
End Function

Private Function FindIndexForDate(inDate As Date, ws As Worksheet) As Integer
   Dim rng As Range: Set rng = ws.Range("A:A")
   
   If Not IsError(Application.Match(CLng(inDate), rng, 0)) Then
        FindIndexForDate = Application.Match(CLng(inDate), rng, 0)
    Else
        MsgBox "Date " + CStr(inDate) + " has not been found"
        FindIndexForDate = xlErrNA
    End If

End Function

'=============================================================
'
'   STRUCTRUAL
'
'=============================================================
Private Sub Class_Initialize()
    pMemoryLimit = 257
End Sub
