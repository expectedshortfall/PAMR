VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstrumentCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim myCustomCollection As Collection


Public Sub Add(ByVal Item As Object)
    myCustomCollection.Add Item
End Sub

Public Sub AddInstruments(ByVal cfs As InstrumentCollection)
    Dim inst As IInstrument
    For Each inst In cfs
        myCustomCollection.Add inst
    Next
End Sub

Public Sub Remove(index As Long)
    myCustomCollection.Remove (index)
End Sub

Public Property Get Item(index As Long) As IInstrument
    Set Item = myCustomCollection.Item(index)
End Property

Public Property Get Count() As Long
    Count = myCustomCollection.Count
End Property

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
