Attribute VB_Name = "Tests"
Option Explicit

Sub TestNPV()
    'VALUATION DATE     2013-02-05
    Dim inDate As Date: inDate = #2/5/2013#
    Dim IM As InputManager: Set IM = New InputManager
        
    Debug.Assert IM.GetValuationDate = inDate
    
    Dim MP As MarketStateProvider: Set MP = IM.GetMarketStateProvider
    Dim v As Portfolio: Set v = IM.GetPortfolio
    
    Dim NPV As Double: NPV = v.GetNPVByCCY(inDate, MP, CCY.PLN)
    Debug.Assert Math.Round(NPV, 10) = _
                 Math.Round(284.44105901793, 10)
                 
    NPV = v.GetNPVByOrigin(inDate, MP, FRA)
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
    Dim IM As InputManager: Set IM = New InputManager
        
    Debug.Assert IM.GetValuationDate = inDate
        
    Dim MSP As MarketStateProvider: Set MSP = IM.GetMarketStateProvider
    Dim MS As MarketState: Set MS = MSP.GetCurrentMarketState
        
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
    
    Dim IM As InputManager: Set IM = New InputManager
    Debug.Assert IM.GetValuationDate = inDate
        
    Dim MSP As MarketStateProvider: Set MSP = IM.GetMarketStateProvider
    Dim MS As MarketState: Set MS = MSP.GetCurrentMarketState
        
    Dim RM As RateManager: Set RM = Factory.CreateRateManager("PLN", "PL3", cAct365)
    
    Debug.Assert Math.Round(RM.GetForwardRate(inDate, startDate, endDate, MS), 12) = _
                 Math.Round(323.125845894649, 12)
    
End Sub

Sub TestMarketState()
    'VALUATION DATE     2013-02-05
    Dim inDate As Date: inDate = #2/5/2013#
    Dim IM As InputManager: Set IM = New InputManager
    
    Debug.Assert IM.GetValuationDate = inDate
    
    Dim MSP As MarketStateProvider: Set MSP = IM.GetMarketStateProvider
    Dim srcMS As MarketState: Set srcMS = MSP.GetMarketStateFromHistory(#5/6/2013#)
    Dim dstMS As MarketState: Set dstMS = srcMS.Clone
    
    dstMS.ShiftMarketParallel 10
    dstMS.ShiftCurveParallel "EURLIBOR", -10
    dstMS.ShiftCurveOnTenor "EURLIBOR", p1M, 10
    dstMS.ShiftFXRate CCY.EUR, 0.15
    
    Debug.Print srcMS.ToString
    Debug.Print dstMS.ToString

End Sub

'=============================================================
'
'   RISK MEASURES
'
'=============================================================
Sub TestBPV()
    'VALUATION DATE     2013-02-05
    Dim inDate As Date: inDate = #2/5/2013#
    Dim IM As InputManager: Set IM = New InputManager
    Dim RM As RiskManager: Set RM = New RiskManager: RM.SetValues IM
    
    Dim bpvPLN As IRiskMeasure: Set bpvPLN = Factory.CreateBPV("PLN")
       
    Debug.Assert Math.Round(bpvPLN.Calculate(RM), 10) = _
                 Math.Round(12.29145263543, 10)
    
End Sub

Sub TestDictionary()
    Dim a As Scripting.Dictionary: Set a = New Scripting.Dictionary
    Dim b() As Double, i As Integer
    Dim arr() As Double: arr = FakeArr(2)
    Dim c As Curve
    Dim D As Object
        
    a.Add "bpv", New Curve
    a.Add "VAR", arr
        
    Debug.Print VarType(GetFromDict(a, "VAR")); " "; vbArray; " "; vbVariant
    Debug.Print VarType(GetFromDict(a, "ES")); " "; vbBoolean
    
    b = IIf(VarType(GetFromDict(a, "VAR")) = vbNull, FakeArr(3), GetFromDict(a, "VAR"))
    Set D = IIf(VarType(GetFromDict(a, "ES")) = vbNull, New Curve, GetFromDict(a, "VAR"))

    For i = LBound(b) To UBound(b)
        Debug.Print b(i)
    Next i

'    Set c = GetFromDict(a, "bpv")

End Sub

Function FakeArr(length As Integer) As Double()
    Dim arr() As Double: ReDim arr(0 To length)
    Dim i As Integer
    
    For i = LBound(arr) To UBound(arr)
        arr(i) = i
    Next i
    
    FakeArr = arr
    
End Function

Function GetFromDict(ByRef dict As Scripting.Dictionary, key As String) As Variant
    If dict.Exists(key) Then
        If VarType(dict.Item(key)) = vbObject Then
            Set GetFromDict = dict.Item(key)
        Else
            GetFromDict = dict.Item(key)
        End If
    Else
        GetFromDict = Null
    End If
End Function
