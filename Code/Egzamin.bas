Attribute VB_Name = "Egzamin"
Option Explicit

Public Sub Zadanie2()
    
    Dim valuationDate As Date: valuationDate = DateSerial(2013, 2, 5)

    Dim NPV As Double
    
    Dim im As InputManager: Set im = New InputManager
    
    Dim MP As MarketStateProvider: Set MP = im.LoadMarketStateProvider(valuationDate)
    Dim v As Portfolio: Set v = im.LoadPortfolio()
    
    NPV = v.GetNPV(valuationDate, MP)
    
    Debug.Print NPV
    ThisWorkbook.Worksheets("Temp").Range("B2").Value = NPV
        
    Set v = Nothing
    Set MP = Nothing
    

End Sub

