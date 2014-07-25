Attribute VB_Name = "Tests"
Option Explicit

Sub TestNPV()
    'VALUATION DATE     2013-02-05
    Dim im As InputManager: Set im = New InputManager
    
    Dim MP As MarketStateProvider: Set MP = im.LoadMarketStateProvider(#2/5/2013#)
    Dim v As Portfolio: Set v = im.LoadPortfolio()
    
    Debug.Print v.GetNPVByCCY(#2/5/2013#, MP, "PLN")
        
    Set v = Nothing
    Set MP = Nothing
    
End Sub

Sub TestDiscountFactor()
    'VALUATION DATE     2013-02-05
    'FORWRAD DATE       2013-04-02
    
    Dim im As InputManager: Set im = New InputManager
    Dim M As MarketState: Set M = im.LoadMarketState(#2/5/2013#)
        
    Dim DF As DiscountFactor: Set DF = Factory.CreateRateManager("PLN", "PL3", cAct365).CreateDiscountFactor()
    
    Debug.Assert Math.Round(DF.Calculate(#2/5/2013#, #4/2/2013#, M), 15) = _
                 Math.Round(0.994373644710383, 15)
    
    Set M = Nothing
    Set im = Nothing
    Set DF = Nothing

End Sub


Sub TestForwardRate()
    'VALUATION DATE     2013-02-05
    'VALUE DATE         2013-04-02
    'MATURITY DATE      2013-07-02
    
    Dim im As InputManager: Set im = New InputManager
    Dim M As MarketState: Set M = im.LoadMarketState(#2/5/2013#)
        
    Dim rm As RateManager: Set rm = Factory.CreateRateManager("PLN", "PL3", cAct365)
    
    Debug.Assert Math.Round(rm.GetForwardRate(#2/5/2013#, #4/2/2013#, #7/2/2013#, M), 12) = _
                 Math.Round(323.125845894649, 12)
    
    Set M = Nothing
    Set im = Nothing
    Set rm = Nothing

End Sub

Sub TestLoadMarketState()
    'VALUATION DATE     2013-02-05
    
    Dim inDate As Date: inDate = #2/5/2013#
    
    Dim im As InputManager: Set im = New InputManager
    Dim M As MarketState: Set M = im.LoadMarketState(inDate)
    
    Debug.Print M.GetCurve("EURLIBOR").ToString

End Sub

Sub TestMarketStateProvider()
    'VALUATION DATE     2013-02-05
    
    Dim inDate As Date: inDate = #2/5/2013#
    Dim im As InputManager: Set im = New InputManager
    Dim MP As MarketStateProvider: Set MP = im.LoadMarketStateProvider(inDate)

    Debug.Print MP.GetCurrentCurve("PLN").ToString
    Debug.Print MP.GetCurrentCurve("USD").ToString
    Debug.Print MP.GetCurrentCurve("EURIBOR").ToString

End Sub


Sub TestZGitHelperExport()
    Dim helper As zGitHelper: Set helper = New zGitHelper
    Dim destPath As String
    
    helper.ExportModules
    
End Sub

Sub TestZGitHelperImport()
    Dim helper As zGitHelper: Set helper = New zGitHelper
    Dim destPath As String
    
    helper.ImportModules
    
End Sub


