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
Dim pCCY As String

'=============================================================
'
'   METHODS
'
'=============================================================
Public Sub Initialize(inCCY As String)
    pCCY = inCCY
End Sub

'=============================================================
'
'   IMPLEMENTING IRiskMeasureCalculator INTERFACE
'
'=============================================================
Public Function IRiskMeasureCalculator_Calculate(valueDate As Date, P As Portfolio, MP As MarketStateProvider) As Double
    Dim currentValue As Double
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
