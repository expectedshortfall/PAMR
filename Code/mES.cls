VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mES"
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

'=============================================================
'
'   METHODS
'
'=============================================================

'=============================================================
'
'   IMPLEMENTING IRiskMeasureCalculator INTERFACE
'
'=============================================================
Public Function IRiskMeasureCalculator_Calculate(valueDate As Date, P As Portfolio, MP As MarketStateProvider) As Double

End Function

Public Sub IRiskMeasureCalculator_Initialize(param As Variant)

End Sub

Public Property Get IRiskMeasureCalculator_Name() As String
    IRiskMeasureCalculator_Name = pName
End Property

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pName = "Expected Shortfall"
End Sub

 


