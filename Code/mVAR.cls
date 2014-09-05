VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mVAR"
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

Dim pAlpha As Double
Dim pIndex As Integer
Const pHistoryLength As Integer = 250
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
Public Function IRiskMeasure_Calculate(dataSource As IDataSource)

End Function

Public Sub IRiskMeasure_Initialize(param As Variant)
    
    If Not VarType(param) = vbDouble Then
        Err.Raise vbObject + 513, "VaR::IRiskMeasure_Initialize", "Param is not a double"
    End If
    
    pAlpha = param
    pIndex = WorksheetFunction.Ceiling(pAlpha * pHistoryLength, 1)
    
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
End Sub

 


