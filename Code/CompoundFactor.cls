VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompoundFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pDcc As DayCountConvention
Private pCurveName As String

'=============================================================
'
'   METHODS
'
'=============================================================
Public Function Calculate(startDate As Date, endDate As Date, ByRef M As MarketState) As Double
    
    Dim r As Double: r = M.GetCurve(pCurveName).GetRate(endDate)
    Calculate = Exp(r * YearFrac(startDate, endDate, pDcc) / 10000)

End Function

Public Sub Initialize(curveName As String, dcc As DayCountConvention)
    
    pCurveName = curveName
    pDcc = dcc

End Sub







