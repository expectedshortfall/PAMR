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

Public Function GetCcyRate(inCcy As CCY) As Double
    If Not pCcyRates.Exists(inCcy) Then
        Err.Raise vbObjectError + 513, "MarketState::GetCcyRate", "Currency: " + Types.CcyToString(inCcy) + " is not present in market"
    Else
        GetCcyRate = pCcyRates.Item(inCcy)
    End If
End Function

Public Sub AddCcyRate(inCcy As CCY, rateValue As Double)
    If Not pCcyRates.Exists(inCcy) Then
        pCcyRates.Add inCcy, rateValue
    End If
End Sub

Public Function GetCurve(curveName As String) As Curve
    If Not pCurves.Exists(curveName) Then
        Err.Raise vbObjectError + 513, "MarketState::GetCurve", "Curve: " + curveName + " is not present in market"
    Else
        Set GetCurve = pCurves.Item(curveName)
    End If
End Function

Public Sub AddCurve(curveObj As Curve)
    If Not pCurves.Exists(curveObj.Name) Then
        pCurves.Add curveObj.Name, curveObj
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift Currency Rate by given multiplier
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftFXRate(inCcy As CCY, multiplier As Double)
    If Not pCcyRates.Exists(inCcy) Then
        Err.Raise vbObjectError + 513, "MarketState::ShiftFXRate", "Currency: " + Types.CcyToString(inCcy) + " is not present in market"
    Else
        pCcyRates.Item(inCcy) = pCcyRates.Item(inCcy) * (1 + multiplier)
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift all curves in parallel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftMarketParallel(shift As Double)
    Dim v As Variant
    
    For Each v In pCurves.keys
        Me.ShiftCurveParallel CStr(v), shift
    Next
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift single curve in parallel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftCurveParallel(curveName As String, shift As Double)
    
    If Not pCurves.Exists(curveName) Then
        Err.Raise vbObjectError + 513, "MarketState::ShiftCurveParallel", "Curve: " + curveName + " is not present in market"
    Else
        Types.CastToCurve(pCurves.Item(curveName)).ShiftParallel shift
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift single rate on a given tenor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftCurveOnTenor(curveName As String, inTenor As Period, shift As Double)
    
    If Not pCurves.Exists(curveName) Then
        Err.Raise vbObjectError + 513, "MarketState::ShiftCurveOnTenor", "Curve: " + curveName + " is not present in market"
    Else
        Types.CastToCurve(pCurves.Item(curveName)).ShiftOnTenor inTenor, shift
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Human readable string representation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
    Dim v As Variant
    Dim res As String: res = "Market state as of " + CStr(pDate) + vbNewLine
    
    For Each v In pCcyRates.keys
        res = res + Types.CcyToString(CLng(v)) + " : " + CStr(pCcyRates.Item(v)) + vbNewLine
    Next
    
    For Each v In pCurves.keys
        res = res + Types.CastToCurve(pCurves.Item(v)).ToString + vbNewLine
    Next
    
    ToString = res
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cloning
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As MarketState
    Dim v As Variant
    Dim myClone As MarketState: Set myClone = Factory.CreateMarketSate(pDate)
    
    For Each v In pCcyRates.keys
        myClone.AddCcyRate CLng(v), pCcyRates.Item(v)
    Next
    
    For Each v In pCurves.keys
        myClone.AddCurve Types.CastToCurve(pCurves.Item(v)).Clone
    Next
    
    Set Clone = myClone
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialization, should be used only in factory module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialize(inDate As Date)
    pDate = inDate
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
