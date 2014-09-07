VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mStressTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRiskMeasure

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Dim pValue As Double
Dim pName As String

Dim pSensitivities() As Double
Dim pTenors As Collection
Dim pCurveNames As Collection
Dim pCurrencies As Collection

'=============================================================
'
'   METHODS
'
'=============================================================

'=============================================================
'
'   IMPLEMENTING IRiskMeasure INTERFACE
'
'=============================================================
Public Function IRiskMeasure_Calculate(ds As IDataSource)

    CalculateSensitivities ds.ValueDate, ds.Portfolio, ds.MarketStateProvider
    OutputSensitivities
    
    pValue = ThisWorkbook.Worksheets("StressTest").Range("stressTestOutput").Value
    
End Function

Public Sub IRiskMeasure_Initialize(param As Variant)
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
    pName = "StressTest"
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
    
    Dim c As Collection: Set c = New Collection
        c.Add "PL3"
        c.Add "PL6"
        c.Add "PLX"
        
        pCurveNames.Add c
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
    Dim v_shifted As Double
    
    Dim names As Variant, tenor As Variant, c As Variant
    Dim i As Integer: i = 1
    
    For Each c In pCurrencies
        pSensitivities(i) = p.GetNPV(inValueDate, MSP.ShiftFXRate(CLng(c), 0.01)) - v
        i = i + 1
    Next c
    
    For Each names In pCurveNames
        For Each tenor In pTenors
            If VarType(names) = vbString Then
                v_shifted = p.GetNPV(inValueDate, MSP.ShiftCurveOnTenor(CStr(names), CLng(tenor), 1))
            Else
                v_shifted = p.GetNPV(inValueDate, MSP.ShiftCurvesOnTenor(Types.CastToCollection(names), CLng(tenor), 1))
            End If
            pSensitivities(i) = v_shifted - v
            i = i + 1
        Next tenor
    Next names

End Sub

Private Sub OutputSensitivities()
    Dim i As Integer
    For i = LBound(pSensitivities) To UBound(pSensitivities)
        ThisWorkbook.Worksheets("StressTest").Range("sensivityAreaStart").Offset(0, i - 1).Value = pSensitivities(i)
    Next i
End Sub
