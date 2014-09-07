VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mVAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1
Option Explicit
Implements IRiskMeasure

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Dim pValue As Double
Dim pName As String

Dim pAlpha As Double
Dim pIndex As Integer
Dim pSensitivities() As Double
Dim pTenors As Collection
Dim pCurveNames As Collection
Dim pCurrencies As Collection

Const pHistoryLength As Integer = 250

'=============================================================
'
'   IMPLEMENTING IRiskMeasure INTERFACE
'
'=============================================================
Public Function IRiskMeasure_Calculate(ds As IDataSource)
        
    Dim MSP As MarketStateProvider: Set MSP = ds.MarketStateProvider
    Dim history() As MarketState: history = MSP.GetHistory(pHistoryLength)
    Dim returns() As Double: ReDim returns(1 To (pHistoryLength - 1))
    
    CalculateSensitivities ds.ValueDate, ds.Portfolio, MSP
    
    Dim i As Long
    For i = LBound(history) To (UBound(history) - 1)
        returns(i) = WorksheetFunction.SumProduct(pSensitivities, CalculateRiskFactorChanges(history(i), history(i + 1)))
    Next i
    
    Call ArrayFunctions.QuickSort(returns, LBound(returns), UBound(returns))
    
    Call ds.AddDataToCache("VAR", returns)
    
    pValue = returns(pIndex)

End Function

Public Sub IRiskMeasure_Initialize(param As Variant)
    
    If Not VarType(param) = vbDouble Then
        Err.Raise vbObject + 513, "VaR::IRiskMeasure_Initialize", "Param is not a double"
    End If
    
    pAlpha = param
    pIndex = WorksheetFunction.Ceiling((1 - pAlpha) * pHistoryLength, 1)
    
End Sub

Public Property Get IRiskMeasure_Name() As String
    IRiskMeasure_Name = pName
End Property

Public Property Get IRiskMeasure_Value() As Double
    IRiskMeasure_Value = pValue
End Property

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pName = "Value-at-Risk"
    
    Set pTenors = New Collection
        pTenors.Add Period.pSN
        pTenors.Add Period.p1M
        pTenors.Add Period.p2M
        pTenors.Add Period.p3M
        pTenors.Add Period.p6M
        pTenors.Add Period.p9M
        pTenors.Add Period.p1Y
        pTenors.Add Period.p2Y
        pTenors.Add Period.p3Y
        pTenors.Add Period.p4Y
        pTenors.Add Period.p5Y
        pTenors.Add Period.p7Y
        pTenors.Add Period.p10Y
        pTenors.Add Period.p20Y
    
    Set pCurveNames = New Collection
        pCurveNames.Add "PLN"
        pCurveNames.Add "PL3"
        pCurveNames.Add "PL6"
        pCurveNames.Add "PLX"
        pCurveNames.Add "EUR"
        pCurveNames.Add "USD"
        
    Set pCurrencies = New Collection
        pCurrencies.Add EUR
        pCurrencies.Add USD

    ReDim pSensitivities(1 To (pTenors.Count * pCurveNames.Count + pCurrencies.Count))

End Sub

Private Sub Class_Terminate()
    
    Set pTenors = Nothing
    Set pCurveNames = Nothing
    Set pCurrencies = Nothing
End Sub

Private Sub CalculateSensitivities(inValueDate As Date, p As Portfolio, MSP As MarketStateProvider)
    
    Dim v As Double: v = p.GetNPV(inValueDate, MSP)
    
    Dim curveName As Variant, tenor As Variant, c As Variant
    Dim i As Integer: i = 1
    
    For Each c In pCurrencies
        pSensitivities(i) = p.GetNPV(inValueDate, MSP.ShiftFXRate(CLng(c), 0.01)) - v
        i = i + 1
    Next c
    
    For Each curveName In pCurveNames
        For Each tenor In pTenors
            pSensitivities(i) = p.GetNPV(inValueDate, MSP.ShiftCurveOnTenor(CStr(curveName), CLng(tenor), 1)) - v
            i = i + 1
        Next tenor
    Next curveName

End Sub

Private Function CalculateRiskFactorChanges(ms_next As MarketState, ms_prev As MarketState) As Double()
    
    Dim res() As Double:  ReDim res(1 To UBound(pSensitivities))
    Dim curveName As Variant, tenor As Variant, c As Variant
    Dim i As Integer: i = 1
    
    For Each c In pCurrencies
        res(i) = ms_next.GetCcyRate(CLng(c)) - ms_prev.GetCcyRate(CLng(c))
        i = i + 1
    Next c
    
    For Each curveName In pCurveNames
        For Each tenor In pTenors
            res(i) = ms_next.GetCurve(CStr(curveName)).GetRateForTenor(CLng(tenor)) - _
                     ms_prev.GetCurve(CStr(curveName)).GetRateForTenor(CLng(tenor))
            i = i + 1
        Next tenor
    Next curveName
    
    CalculateRiskFactorChanges = res

End Function

