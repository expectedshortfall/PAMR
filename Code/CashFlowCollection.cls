VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashFlowCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim myCustomCollection As Collection

Public Sub Add(ByVal Item As Object)
    myCustomCollection.Add Item
End Sub

Public Sub AddCashFlow(ParamArray arr() As Variant)
    Dim v As Variant
    For Each v In arr
        myCustomCollection.Add v
    Next
End Sub

Public Sub AddCashFlows(ByVal cfs As CashFlowCollection)
    Dim cf As CashFlow
    For Each cf In cfs
        myCustomCollection.Add cf
    Next
End Sub

Public Sub Remove(index As Long)
    myCustomCollection.Remove (index)
End Sub

Public Sub Clear()
    While myCustomCollection.Count <> 0
        myCustomCollection.Remove (myCustomCollection.Count)
    Wend
End Sub

Public Property Get Item(index As Long) As CashFlow
    Set Item = myCustomCollection.Item(index)
End Property

Public Property Get Count() As Long
    Count = myCustomCollection.Count
End Property

Public Function FilterByCCY(str As String) As CashFlowCollection
    Dim cfs As CashFlowCollection, cf As CashFlow
    Set cfs = New CashFlowCollection

    For Each cf In Me
        If cf.CCY = str Then cfs.Add cf
    Next

    Set FilterByCCY = cfs
End Function

Public Function FilterByOrigin(str As String) As CashFlowCollection
    Dim cfs As CashFlowCollection, cf As CashFlow
    Set cfs = New CashFlowCollection

    For Each cfs In Me
        If cf.Origin Like str Then cfs.Add cf
    Next

    Set FilterByOrigin = cfs
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = myCustomCollection.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set myCustomCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set myCustomCollection = Nothing
End Sub

