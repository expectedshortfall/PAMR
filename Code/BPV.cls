VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRiskMeasureCalculator

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Dim pName As String
Dim pCCY As CCY

'=============================================================
'
'   METHODS
'
'=============================================================
Public Sub Initialize(inCcy As CCY)
    pCCY = inCcy
End Sub

'=============================================================
'
'   IMPLEMENTING IRiskMeasureCalculator INTERFACE
'
'=============================================================
Public Function IRiskMeasureCalculator_Calculate(valueDate As Date, P As Portfolio, MSP As MarketStateProvider) As Double
    
    Dim currentValue As Double: currentValue = P.GetNPVByCCY(valueDate, MSP, pCCY)
    Dim shiftedValue As Double
    
    IRiskMeasureCalculator_Calculate = (currentValue - shiftedValue)

End Function

Public Property Get IRiskMeasureCalculator_Name() As String
    IRiskMeasureCalculator_Name = pName
End Property

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pName = "BPV"
End Sub
