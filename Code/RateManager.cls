VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RateManager"
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
Private pDcc As DayCountConvention
Private pDiscountCurveName As String
Private pForwardCurveName As String


Public Sub Initialize(inDiscountCurveName As String, inForwardCurveName As String, inDcc As DayCountConvention)
    
    pDiscountCurveName = inDiscountCurveName
    pForwardCurveName = inForwardCurveName
    pDcc = inDcc

End Sub

Public Function CreateDiscountFactor() As DiscountFactor
    
    Set CreateDiscountFactor = New DiscountFactor
    CreateDiscountFactor.Initialize discountCurveName:=pDiscountCurveName, dcc:=pDcc

End Function

Public Function CreateCompoundFactor() As CompoundFactor
    
    Set CreateCompoundFactor = New CompoundFactor
    CreateCompoundFactor.Initialize curveName:=pForwardCurveName, dcc:=pDcc
    
End Function

Public Function GetRateForTenor(tenor As Period, ByRef M As MarketState) As Double
    GetRateForTenor = M.GetCurve(pForwardCurveName).GetRateForTenor(tenor)
End Function

Public Function GetForwardRate(valuationDate As Date, _
                               startDate As Date, _
                               endDate As Date, _
                               ByRef M As MarketState) As Double
                    
        GetForwardRate = M.GetCurve(pForwardCurveName).GetForwardRate(valuationDate, startDate, endDate, pDcc)

End Function

Public Function GetFDF(startDate As Date, endDate As Date, forwardRate As Double)
    GetFDF = Exp(-forwardRate * YearFrac(startDate, endDate, pDcc) / 10000)
End Function

Public Function GetYearFraction(startDate As Date, endDate As Date) As Double
    GetYearFraction = YearFrac(startDate, endDate, pDcc)
End Function
