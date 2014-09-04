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
Public Function GetMarketStateFromDate(SnapshotDate As Date) As MarketState
    If pMarketStates.Exists(SnapshotDate) Then
         Set GetMarketStateFromDate = pMarketStates.Item(SnapshotDate)
    Else
         Err.Raise vbObject + 513, "MarketStateProvider::GetMarketStateFromDate", "There is no data for given date: " + CStr(SnapshotDate)
         GetMarketStateFromDate = xlErrNA
    End If

End Function

Public Sub AddMarketState(MS As MarketState)
    If Not pMarketStates.Exists(MS.SnapshotDate) Then
        pMarketStates.Add MS.SnapshotDate, MS
    End If
End Sub

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

Public Sub Initialize(currentDate As Date)
    pCurrentDate = currentDate
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets currency rate for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentCcyRate(rateName As String) As Double
    GetCurrentCcyRate = pMarketStates.Item(pCurrentDate).GetCcyRate(rateName)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets curve for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentCurve(curveName As String) As Curve
     Set GetCurrentCurve = pMarketStates.Item(pCurrentDate).GetCurve(curveName)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets MarketState object for current Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentMarketState() As MarketState
    Set GetCurrentMarketState = GetMarketStateFromDate(pCurrentDate)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cloning
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As MarketStateProvider
    Dim v As Variant
    Dim myClone As MarketStateProvider: Set myClone = Factory.MarketStateProvider(pCurrentDate)
    
    For Each v In pMarketStates.keys
        myClone.AddMarketState Types.CastToMarketState(pMarketStates.Item(v)).Clone
    Next
    
    Set Clone = myClone
    
End Function

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


