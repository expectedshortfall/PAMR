VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRiskMeasureCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'
'   PUBLIC IRISK_MEASURE INTERFACE
'
'=============================================================
Public Function Calculate(valueDate As Date, P As Portfolio, MP As MarketStateProvider) As Double
End Function

Public Sub Initialize(param As Variant)
End Sub

Public Property Get Name() As String
End Property
