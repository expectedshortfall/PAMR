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
Private underlying As String
Private K As Double
Private premium As Double
Private Const settlementCCy As Currency = CCY.PLN

Private Const MAX_ITERATIONS As Integer = 100
'=============================================================
'
'   IMPLEMENTING IINSTRUMENT INTERFACE
'
'=============================================================
Public Function IInstrument_GetCashFlows(inValueDate As Date, MP As MarketStateProvider) As CashFlowCollection
    
    Dim res As CashFlowCollection: Set res = New CashFlowCollection
    
    If inValueDate <= tradeDate Or inValueDate >= maturityDate Then
        Set IInstrument_GetCashFlows = res: Exit Function
    End If
    
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
                      inNominal As Double, inUnderlying As String, inK As Double, inPremium As Double)
 
    tradeDate = inTradeDate
    maturityDate = inMaturityDate
    settlementDate = inSettlementDate
    nominal = inNominal
    theta = inTheta
    underlying = inUnderlying
    K = inK
    premium = inPremium
    
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Price of FX option under BS assumptions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Garman–Kohlhagen(S0 As Double, sigma As Double, r_d_tau As Double, r_f_tau As Double, _
                                  r_d_T As Double, r_f_T As Double, tau As Double, T As Double) As Double
    Dim d1 As Double
    Dim d2 As Double
    Dim p As Double
    
    d1 = (Math.Log(S0 / K) + (r_d_tau - r_f_tau + 0.5 * sigma ^ 2) * tau) / (sigma * Math.Sqr(tau))
    d2 = d1 - sigma * Math.Sqr(tau)
    
    p = theta * (S0 * Math.Exp(-r_f_T * T) * WorksheetFunction.NormSDist(theta * d1) - _
                  K * Math.Exp(-r_d_T * T) * WorksheetFunction.NormSDist(theta * d2))
    
    Garman–Kohlhagen = p

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Calculate implied volatitlity from option private values
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TODO: dla trade date'u wyliczyæ zmiennoœæ implikowan¹
'Function getImpliedVolatility(Optional eps = 10 ^ (-3)) As Double
'
'    Dim bsPrice As Double
'    Dim i As Integer
'    Dim v As Double
'    Dim res As Double
'    Dim tmpO As New COption
'
'    tmpO.S0 = o.S0
'    tmpO.K = o.K
'    tmpO.r = o.r
'    tmpO.q = o.q
'    tmpO.T = o.T
'    tmpO.sigma = sigma_def
'    i = 0
'
'   Do
'        ' Cena z modelu B-S dla aktualnej zmiennoœci
'        bsPrice = BlackScholesPrice(tmpO, payoffType)
'        res = Abs(bsPrice - marketPrice)
'
'        If (res < eps) Then
'            getImpliedVolatility = tmpO.sigma
'            Exit Do
'        Else
'            ' Bronimy siê przed dzieleniem przez zero
'            v = vega(tmpO, payoffType)
'            If (v < 10 ^ (-260)) Then
'                getImpliedVolatility = -1
'                Exit Do
'            End If
'
'            tmpO.sigma = tmpO.sigma - (bsPrice - marketPrice) / v
'
'            ' zmiennosc musi byc dodatnia
'            If (tmpO.sigma <= 0) Then
'                getImpliedVolatility = -1
'                Exit Do
'            End If
'
'            'wymuszamy przerwanie petli
'            If (i >= MAX_ITERATIONS) Then
'                getImpliedVolatility = -2
'                Exit Do
'            End If
'            i = i + 1
'        End If
'    Loop
'End Function

'http://www.riskglossary.com/link/garman_kohlhagen_1983.htm
' ******************************************************************************************************
'       VEGA W MODELU B-S
' ******************************************************************************************************
'Function vega(o As COption, payoffType As String) As Double
'    Dim theta As Integer
'    theta = GetTheta(payoffType)
'    vega = Math.Exp(-o.q * o.T) * o.S0 * Math.Sqr(o.T) * WorksheetFunction.NormDist(d1(o), 0, 1, False)
'End Function

'Option parameters

' Public S0 As Double
' Public sigma As Double
' Public r As Double
' Public q As Double
' Public K As Double
' Public T As Double
' Public N As Integer
'
'
'Property Get h() As Double
'    h = T / N
'End Property
'
'Property Get m() As Double
'    m = Exp(r * h())
'End Property
'
'Property Get v() As Double
'    v = Exp(sigma ^ 2 * h())
'End Property
'
'Property Get alpha() As Double
'    alpha = r - 0.5 * sigma ^ 2
'End Property

