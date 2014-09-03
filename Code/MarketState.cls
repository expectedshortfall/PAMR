VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MarketState"
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
Private pCcyRates As Scripting.Dictionary
Private pCurves As Scripting.Dictionary
Private pDate As Date

'=============================================================
'
'   METHODS
'
'=============================================================
Public Property Get SnapshotDate() As Date
    SnapshotDate = pDate
End Property

Public Sub Initialize(inDate As Date)
    pDate = inDate
End Sub

Public Function GetCcyRate(ccyRate As CCY) As Double
    If pCcyRates.Exists(ccyRate) Then
         GetCcyRate = pCcyRates.Item(ccyRate)
    Else
        GetCcyRate = -1
    End If
End Function

Public Sub AddCcyRate(ccyRate As CCY, rateValue As Double)
    If Not pCcyRates.Exists(ccyRate) Then
        pCcyRates.Add ccyRate, rateValue
    End If
End Sub

Public Function GetCurve(curveName As String) As Curve
    If pCurves.Exists(curveName) Then
         Set GetCurve = pCurves.Item(curveName)
    Else
         MsgBox "There is no " + curveName + " curve in the Market"
    End If
End Function

Public Sub AddCurve(curveObj As Curve)
    If Not pCurves.Exists(curveObj.Name) Then
        pCurves.Add curveObj.Name, curveObj
    End If
End Sub

'=============================================================
'
'   STRUCTRUAL
'
'=============================================================
Private Sub Class_Initialize()
    Set pCcyRates = New Scripting.Dictionary
    Set pCurves = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set pCcyRates = Nothing
    Set pCurves = Nothing
End Sub
