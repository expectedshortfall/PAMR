Attribute VB_Name = "Tests"
Option Explicit

Sub TestNPV()
    'VALUATION DATE     2013-02-05
    Dim im As InputManager: Set im = New InputManager
    
    Dim MP As MarketStateProvider: Set MP = im.LoadMarketStateProvider(#2/5/2013#)
    Dim v As Portfolio: Set v = im.LoadPortfolio()
    
    Dim NPV As Double: NPV = v.GetNPVByCCY(#2/5/2013#, MP, CCY.PLN)
    Debug.Assert Math.Round(NPV, 10) = _
                 Math.Round(284.44105901793, 10)
                 
    NPV = v.GetNPVByOrigin(#2/5/2013#, MP, FRA)
    Debug.Assert Math.Round(NPV, 10) = _
                 Math.Round(284.44105901793, 10)

End Sub

Sub TestCurve()
    Dim c As Curve: Set c = Factory.CreateCurve("TEST", #1/1/2014#)
    
    c.AddRate pSN, 3
    c.AddRate p1M, 4
    c.ShiftParallel 1
                
    Debug.Assert c.GetRateForTenor(pSN) = 4
    Debug.Assert c.GetRateForTenor(p1M) = 5
        
    c.ShiftOnTenor p1M, 10
        
    Debug.Assert c.GetRateForTenor(pSN) = 4
    Debug.Assert c.GetRateForTenor(p1M) = 15
        
End Sub

Sub TestDiscountFactor()
    'VALUATION DATE     2013-02-05
    'FORWRAD DATE       2013-04-02
    
    Dim inDate As Date: inDate = #2/5/2013#
    Dim forwardDate As Date: forwardDate = #4/2/2013#
    
    Dim im As InputManager: Set im = New InputManager
    Dim MSP As MarketStateProvider: Set MSP = im.LoadMarketStateProvider(inDate)
    Dim MS As MarketState: Set MS = MSP.GetMarketStateFromDate(inDate)
        
    Dim DF As DiscountFactor: Set DF = Factory.CreateRateManager("PLN", "PL3", cAct365).CreateDiscountFactor()
    
    Debug.Assert Math.Round(DF.Calculate(inDate, forwardDate, MS), 15) = _
                 Math.Round(0.994373644710383, 15)
    
End Sub


Sub TestForwardRate()
    'VALUATION DATE     2013-02-05
    'VALUE DATE         2013-04-02
    'MATURITY DATE      2013-07-02
    
    Dim inDate As Date: inDate = #2/5/2013#
    Dim startDate As Date: startDate = #4/2/2013#
    Dim endDate As Date: endDate = #7/2/2013#
    
    Dim im As InputManager: Set im = New InputManager
    Dim MSP As MarketStateProvider: Set MSP = im.LoadMarketStateProvider(inDate)
    Dim MS As MarketState: Set MS = MSP.GetMarketStateFromDate(inDate)
        
    Dim rm As RateManager: Set rm = Factory.CreateRateManager("PLN", "PL3", cAct365)
    
    Debug.Assert Math.Round(rm.GetForwardRate(inDate, startDate, endDate, MS), 12) = _
                 Math.Round(323.125845894649, 12)
    
End Sub

Sub TestMarketState()
    'VALUATION DATE     2013-02-05
    
    Dim inDate As Date: inDate = #2/5/2013#
    
    Dim im As InputManager: Set im = New InputManager
    Dim MSP As MarketStateProvider: Set MSP = im.LoadMarketStateProvider(inDate)
    
    Dim srcMS As MarketState: Set srcMS = MSP.GetMarketStateFromDate(inDate)
    Dim dstMS As MarketState: Set dstMS = srcMS.Clone
    
    dstMS.ShiftMarketParallel 10
    dstMS.ShiftCurveParallel "EURLIBOR", -10
    dstMS.ShiftCurveOnTenor "EURLIBOR", p1M, 10
    dstMS.ShiftFXRate CCY.EUR, 0.15
    
    Debug.Print srcMS.ToString
    Debug.Print dstMS.ToString

End Sub
