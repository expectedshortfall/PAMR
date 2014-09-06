VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FXOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInstrument

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pOrigin As Origin
Private pName As String

Private tradeDate As Date
Private maturityDate As Date        'tau
Private settlementDate As Date      'T
Private nominal As Double
Private theta As Integer
Private underlying As CCY
Private K As Double
Private premium As Double
Private Const settlementCCy As Currency = CCY.PLN

Private sigma As Double                         'implied volatility
Private Const MAX_ITERATIONS As Integer = 1000   'Max iteration of imp volatility callculations

Private Type sFXOption
    S0 As Double
    K As Double
    sigma As Double
    r_d_tau As Double
    r_d_T As Double
    r_f_tau As Double
    r_f_T As Double
    tau As Double
    T As Double
    theta As Double
End Type

Dim pRateManager1 As RateManager
Dim pRateManager2 As RateManager

'=============================================================
'
'   IMPLEMENTING IINSTRUMENT INTERFACE
'
'=============================================================
Public Function IInstrument_GetCashFlows(inValueDate As Date, MP As MarketStateProvider) As CashFlowCollection
    
    Dim res As CashFlowCollection: Set res = New CashFlowCollection
    
    If inValueDate <= tradeDate Or inValueDate >= settlementDate Then
        Set IInstrument_GetCashFlows = res: Exit Function
    End If
        
    CalculateImpliedVolatility MP 'First implied volatility needs to be calculated
    
    Dim o As sFXOption: o = PackDataIntoStructure(MP.GetCurrentMarketState())
    Dim cf1 As Double: cf1 = nominal * (theta * MP.GetCurrentCcyRate(underlying) * WorksheetFunction.NormSDist(theta * d1(o)))
    Dim cf2 As Double: cf2 = nominal * (-theta * K * WorksheetFunction.NormSDist(theta * d2(o)))
    
    res.AddCashFlow Factory.CreateCashFlow(Origin.FXSpot, underlying, cf1, settlementDate, pRateManager1.CreateDiscountFactor)
    res.AddCashFlow Factory.CreateCashFlow(Origin.FXSpot, settlementCCy, cf2, settlementDate, pRateManager2.CreateDiscountFactor)
    
    Set IInstrument_GetCashFlows = res

End Function

Public Property Get IInstrument_Name() As String
    IInstrument_Name = pName
End Property
'=============================================================
'
'   METHODS
'
'=============================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialization, should be used only in factory module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialize(inTradeDate As Date, inMaturityDate As Date, inSettlementDate As Date, inTheta As Integer, _
                      inNominal As Double, inUnderlying As CCY, inK As Double, inPremium As Double, _
                      inRateManager1 As RateManager, inRateManager2 As RateManager)
 
    tradeDate = inTradeDate
    maturityDate = inMaturityDate
    settlementDate = inSettlementDate
    nominal = inNominal
    theta = inTheta
    underlying = inUnderlying
    K = inK
    premium = inPremium
    
    Set pRateManager1 = inRateManager1
    Set pRateManager2 = inRateManager2
    
End Sub

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pOrigin = Origin.FXOption
    pName = "FXOption"
End Sub

Private Sub Class_Terminate()
    Set pRateManager1 = Nothing
    Set pRateManager2 = Nothing
End Sub

Private Sub CalculateImpliedVolatility(MSP As MarketStateProvider)
            
    Dim MS As MarketState: Set MS = MSP.GetMarketStateFromHistory(tradeDate)
    Dim o As sFXOption: o = PackDataIntoStructure(MS, 0.1)
    
    sigma = getImpliedVolatility(o, premium / nominal) 'this is not converging
    
End Sub

Private Function PackDataIntoStructure(MS As MarketState, Optional inSigma As Double = -1) As sFXOption
    Dim o As sFXOption
    
    o.S0 = MS.GetCcyRate(underlying)
    o.sigma = IIf(inSigma = -1, sigma, 0.01)
    o.K = K
    
    o.T = pRateManager1.GetYearFraction(tradeDate, settlementDate)
    o.tau = pRateManager1.GetYearFraction(tradeDate, maturityDate)
    o.theta = theta
    
    o.r_f_tau = MS.GetCurve(pRateManager1.forwardCurveName).GetRate(maturityDate) * 0.0001
    o.r_f_T = MS.GetCurve(pRateManager1.discountCurveName).GetRate(settlementDate) * 0.0001
    
    o.r_d_tau = MS.GetCurve(pRateManager2.forwardCurveName).GetRate(maturityDate) * 0.0001
    o.r_d_T = MS.GetCurve(pRateManager2.discountCurveName).GetRate(settlementDate) * 0.0001
    
    PackDataIntoStructure = o
    
End Function

'=============================================================
'
'   Price of FX option under BS assumptions
'
'=============================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       PARAMETR d1 W MODELU Garman'a-Kohlhagen'a
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function d1(o As sFXOption)
    d1 = (Math.Log(o.S0 / o.K) + (o.r_d_tau - o.r_f_tau + 0.5 * o.sigma ^ 2) * o.tau) / (o.sigma * Math.Sqr(o.tau))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       PARAMETR d1 W MODELU Garman'a-Kohlhagen'a
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function d2(o As sFXOption)
    d2 = d1(o) - o.sigma * Math.Sqr(o.tau)
End Function

Private Function Garman–Kohlhagen(o As sFXOption) As Double

    Garman–Kohlhagen = o.theta * (o.S0 * Math.Exp(-o.r_f_T * o.T) * WorksheetFunction.NormSDist(theta * d1(o)) - _
                                   o.K * Math.Exp(-o.r_d_T * o.T) * WorksheetFunction.NormSDist(theta * d2(o)))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       VEGA W MODELU Garman'a-Kohlhagen'a
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function vega(o As sFXOption) As Double
    vega = Math.Exp(-o.r_f_T * o.T) * o.S0 * Math.Sqr(o.T) * WorksheetFunction.NormDist(d1(o), 0, 1, False)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Calculate implied volatitlity from option private values
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function getImpliedVolatility(o As sFXOption, marketPrice As Double, Optional eps = 10 ^ (-3)) As Double

    Dim oPrice As Double
    Dim i As Integer
    Dim v As Double
    Dim res As Double
    Dim tmpO As sFXOption: tmpO = o

    Do
        ' Cena z modelu B-S dla aktualnej zmiennoœci
        oPrice = Garman–Kohlhagen(tmpO)
        res = Abs(oPrice - marketPrice)
    
        If (res < eps) Then
            getImpliedVolatility = tmpO.sigma
            Exit Do
        Else
            ' Bronimy siê przed dzieleniem przez zero
            v = vega(tmpO)
            If (v < 10 ^ (-260)) Then
                getImpliedVolatility = -1
                Exit Do
            End If
    
            tmpO.sigma = tmpO.sigma - (oPrice - marketPrice) / v
    
            ' zmiennosc musi byc dodatnia
            If (tmpO.sigma <= 0) Then
                getImpliedVolatility = -2
                Exit Do
            End If
    
            'wymuszamy przerwanie petli
            If (i >= MAX_ITERATIONS) Then
                getImpliedVolatility = -3
                Exit Do
            End If
            i = i + 1
        End If
    Loop
End Function

