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
Option Base 1
'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pCurrentDate As Date
Private pCurrentMarketState As MarketState
Private pHistory As Scripting.Dictionary

'=============================================================
'
'   METHODS
'
'=============================================================
Public Sub AddMarketState(MS As MarketState)
    If MS.SnapshotDate = pCurrentDate Then
       
       Set pCurrentMarketState = MS
    
    ElseIf MS.SnapshotDate < pCurrentDate Then
        
        If Not pHistory.Exists(MS.SnapshotDate) Then
            pHistory.Add MS.SnapshotDate, MS
        End If
    
    Else
        Err.Raise vbObject + 514, "MarketStateProvider::AddMarketState", "Opperation not supported: MS.SnapshotDate > pCurrentDate"
    End If
End Sub

Public Function GetMarketStateFromHistory(SnapshotDate As Date) As MarketState
    
    ValidateInDate SnapshotDate, "GetMarketStateFromHistory"
         
    Set GetMarketStateFromHistory = pHistory.Item(SnapshotDate)

End Function

Public Function GetHistory(length As Integer) As MarketState()
    
    If (length >= pHistory.Count) Then
        MsgBox "Not enough data, currenty there is " + pHistory.Count + " market states"
        Exit Function
    End If
    
    Dim res() As MarketState: ReDim res(1 To length)
    Dim K As Variant, i As Integer: i = 2
    
    Set res(1) = pCurrentMarketState
    
    For Each K In pHistory.keys
        If i <= length Then
            Set res(i) = pHistory.Item(K)
        End If
        i = i + 1
    Next
    
    GetHistory = res
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets currency rate for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentCcyRate(rateName As CCY) As Double
    GetCurrentCcyRate = pCurrentMarketState.GetCcyRate(rateName)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets curve for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentCurve(curveName As String) As Curve
     Set GetCurrentCurve = pCurrentMarketState.GetCurve(curveName)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets MarketState object for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentMarketState() As MarketState
    Set GetCurrentMarketState = pCurrentMarketState
End Function

'=============================================================
'
'   MARKET SHIFTS FOR RISK MEASURES
'
'=============================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Return clone with Shift Currency Rate by given multiplier
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShiftFXRate(inCcy As CCY, multiplier As Double) As MarketStateProvider
        
    Set ShiftFXRate = Me.Clone
        ShiftFXRate.GetCurrentMarketState.ShiftFXRate inCcy, multiplier
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift all curves in parallel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShiftMarketParallel(shift As Double) As MarketStateProvider
            
    Set ShiftMarketParallel = Me.Clone
        ShiftMarketParallel.GetCurrentMarketState.ShiftMarketParallel shift
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift single curve in parallel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShiftCurveParallel(curveName As String, shift As Double) As MarketStateProvider
    
    Set ShiftCurveParallel = Me.Clone
        ShiftCurveParallel.GetCurrentMarketState.GetCurve(curveName).ShiftParallel shift
            
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift all curves in on a given tenor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShiftMarketOnTenor(inTenor As Period, shift As Double) As MarketStateProvider
            
    Set ShiftMarketOnTenor = Me.Clone
        ShiftMarketOnTenor.GetCurrentMarketState.ShiftMarketOnTenor inTenor, shift
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift single rate on a given tenor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShiftCurveOnTenor(curveName As String, inTenor As Period, shift As Double) As MarketStateProvider
    
    Set ShiftCurveOnTenor = Me.Clone
        ShiftCurveOnTenor.GetCurrentMarketState.GetCurve(curveName).ShiftOnTenor inTenor, shift

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cloning
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As MarketStateProvider
    
    Set Clone = Factory.CreateMarketStateProvider(pCurrentDate, pCurrentMarketState.Clone, pHistory)
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialization, should be used only in factory module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialize(inCurrentDate As Date, Optional inMarketState As MarketState, _
                                             Optional inHistory As Scripting.Dictionary)
    
    pCurrentDate = inCurrentDate
    
    If Not inMarketState Is Nothing Then
        Set pCurrentMarketState = inMarketState
    End If
    
    If Not inHistory Is Nothing Then
        Set pHistory = inHistory
    End If
    
End Sub

'=============================================================
'
'   STRUCTRUAL
'
'=============================================================
Private Sub Class_Initialize()
    Set pHistory = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set pHistory = Nothing
End Sub

Private Function ValidateInDate(ByRef inDate As Date, Optional methodName As String = "ValidateInDate")
    
    If inDate = 0 Then
        inDate = pCurrentDate
    End If
    
    If Not pHistory.Exists(inDate) Then
        Err.Raise vbObjectError + 513, "MarketStateProvider::" + methodName, "There is no data for given date: " + CStr(inDate)
    End If

End Function
