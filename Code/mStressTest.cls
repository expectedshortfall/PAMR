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

'    - PLN_GOVT -> PLN
'    - PLN_SWAP -> PLNX, PLN3, PLN6
'    - EUR -> EUR
'    - USD -> USD


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
End Sub

 


