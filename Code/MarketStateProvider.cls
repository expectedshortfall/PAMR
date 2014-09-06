VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MarketStateProvider"
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
Private pMarketStates As Scripting.Dictionary
Private pCurrentDate As Date

'=============================================================
'
'   METHODS
'
'=============================================================
Public Sub AddMarketState(MS As MarketState)
    If Not pMarketStates.Exists(MS.SnapshotDate) Then
        pMarketStates.Add MS.SnapshotDate, MS
    End If
End Sub

Public Function GetMarketStateFromDate(SnapshotDate As Date) As MarketState
    
    ValidateInDate SnapshotDate, "GetMarketStateFromDate"
         
    Set GetMarketStateFromDate = pMarketStates.Item(SnapshotDate)

End Function

Public Function GetHistory(length As Integer) As MarketState()
    
    If (length >= pMarketStates.Count) Then
        MsgBox "Not enough data, currenty there is " + pMarketStates.Count + " market states"
        Exit Function
    End If
    
    Dim res() As MarketState: ReDim res(length)
    Dim keys() As MarketState: keys = pMarketStates.keys
    Dim v As Variant, i As Integer
    
    For Each v In keys
        If i > 0 And i <= length Then
            res(i - 1) = keys(i)
        End If
    Next
    
    GetHistory = res
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets currency rate for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentCcyRate(rateName As CCY) As Double
    GetCurrentCcyRate = GetMarketStateFromDate(pCurrentDate).GetCcyRate(rateName)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets curve for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentCurve(curveName As String) As Curve
     Set GetCurrentCurve = GetMarketStateFromDate(pCurrentDate).GetCurve(curveName)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets MarketState object for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentMarketState() As MarketState
    Set GetCurrentMarketState = GetMarketStateFromDate(pCurrentDate)
End Function

'=============================================================
'
'   MARKET SHIFTS
'
'=============================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift Currency Rate by given multiplier
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftFXRate(inCcy As CCY, multiplier As Double, Optional inDate As Date)
    
    ValidateInDate inDate, "ShiftFXRate"
    
    GetMarketStateFromDate(inDate).ShiftFXRate inCcy, multiplier
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift all curves in parallel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftMarketParallel(shift As Double, Optional inDate As Date)
        
    ValidateInDate inDate, "ShiftMarketParallel"
    GetMarketStateFromDate(inDate).ShiftMarketParallel shift
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift single curve in parallel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftCurveParallel(curveName As String, shift As Double, Optional inDate As Date)
        
    ValidateInDate inDate, "ShiftCurveParallel"
    
    Dim myCurve As Curve: Set myCurve = Types.CastToMarketState(pMarketStates.Item(inDate)).GetCurve(curveName)
    
    myCurve.ShiftParallel shift
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift single rate on a given tenor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftCurveOnTenor(curveName As String, inTenor As Period, shift As Double, Optional inDate As Date)
    
    ValidateInDate inDate, "ShiftCurveOnTenor"
    
    Dim myCurve As Curve: Set myCurve = Types.CastToMarketState(pMarketStates.Item(inDate)).GetCurve(curveName)
    
    myCurve.ShiftOnTenor inTenor, shift

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cloning
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As MarketStateProvider
    Dim v As Variant
    Dim myClone As MarketStateProvider: Set myClone = Factory.CreateMarketStateProvider(pCurrentDate)
    
    For Each v In pMarketStates.keys
        myClone.AddMarketState Types.CastToMarketState(pMarketStates.Item(v)).Clone
    Next
    
    Set Clone = myClone
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialization, should be used only in factory module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialize(currentDate As Date)
    pCurrentDate = currentDate
End Sub

'=============================================================
'
'   STRUCTRUAL
'
'=============================================================
Private Sub Class_Initialize()
    Set pMarketStates = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set pMarketStates = Nothing
End Sub

Private Function ValidateInDate(ByRef inDate As Date, Optional methodName As String = "ValidateInDate")
    
    If inDate = 0 Then
        inDate = pCurrentDate
    End If
    
    If Not pMarketStates.Exists(inDate) Then
        Err.Raise vbObjectError + 513, "MarketStateProvider::" + methodName, "There is no data for given date: " + CStr(inDate)
    End If

End Function
