Attribute VB_Name = "Egzamin"
Option Explicit

Public Sub Zadanie2()
    
    Dim valuationDate As Date: valuationDate = DateSerial(2013, 2, 5)

    Dim npv As Double
    
    Dim im As InputManager: Set im = New InputManager
    
    Dim MP As MarketStateProvider: Set MP = im.LoadMarketStateProvider(valuationDate)
    Dim v As Portfolio: Set v = im.LoadPortfolio()
    
    npv = v.GetNPV(valuationDate, MP)
    
    Debug.Print npv
    ThisWorkbook.Worksheets("Temp").Range("B2").Value = npv
        
    Set v = Nothing
    Set MP = Nothing
    

End Sub

