VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements IRiskMeasure

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Dim pValue As Double
Dim pName As String
Dim pCCY As CCY

Dim pTenors As Collection
Dim pSensitivities() As Double

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
    
    Dim L As Double: L = 0
    Dim S As Double: S = 0
    Dim Rot As Double: Rot = 0

    Call CalculateSensitivities(ds.ValueDate, ds.Portfolio, ds.MarketStateProvider)
        
    Dim i As Integer
    For i = LBound(pSensitivities) To UBound(pSensitivities)
        L = L + WorksheetFunction.Max(pSensitivities(i), 0)
        S = S + WorksheetFunction.Min(pSensitivities(i), 0)
    Next i
       
    Dim yf As Double
    For i = LBound(pSensitivities) To UBound(pSensitivities)
       yf = YearFrac(ds.ValueDate, DateMove(ds.ValueDate, i - 1, Following), cAct365)
       Rot = Rot + yf * pSensitivities(i) * w(pSensitivities(i), L, S)
    Next i

End Function

Public Sub IRiskMeasure_Initialize(param As Variant)
    If Not VarType(param) = vbString Then
        Err.Raise vbObject + 513, "BPV::IRiskMeasure_Initialize", "Param is not a string"
    End If
    
    pCCY = Types.StringToCCY(CStr(param))
    pName = pName + param
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
    pName = "Rotation_"
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
      
    ReDim pSensitivities(1 To pTenors.Count)

End Sub

Private Sub Class_Terminate()
    Set pTenors = Nothing
End Sub


Private Sub CalculateSensitivities(inValueDate As Date, p As Portfolio, MSP As MarketStateProvider)
    
    Dim v As Double: v = p.GetNPVByCCY(inValueDate, MSP, pCCY)
    
    Dim tenor As Variant
    Dim i As Integer: i = 1
    
    For Each tenor In pTenors
        pSensitivities(i) = p.GetNPVByCCY(inValueDate, MSP.ShiftMarketOnTenor(CLng(tenor), 1), pCCY) - v
        i = i + 1
    Next tenor

End Sub

Private Function w(x As Double, L As Double, S As Double) As Double
    w = WorksheetFunction.Min(L, S) * IIf(x > 0, 1 / L, 1 / S)
End Function

 

