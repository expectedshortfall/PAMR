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

'=============================================================
'
'   IMPLEMENTING IRiskMeasureCalculator INTERFACE
'
'=============================================================
Public Function IRiskMeasureCalculator_Calculate(valueDate As Date, P As Portfolio, MP As MarketStateProvider) As Double

End Function

Public Sub IRiskMeasureCalculator_Initialize(param As Variant)
    If Not VarType(param) = vbString Then
        Err.Raise vbObject + 513, "BPV::IRiskMeasureCalculator_Initialize", "Param is not a string"
    End If
    
    pCCY = Types.StringToCCY(CStr(param))
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
    pName = "Rotation_" + Types.CcyToString(pCCY)
End Sub

 

