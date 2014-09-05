VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRiskMeasureCollection"
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
Dim myCustomCollection As Collection

'=============================================================
'
'   METHODS
'
'=============================================================
Public Sub Add(ByRef Item As IRiskMeasure)
    myCustomCollection.Add Item
End Sub

Public Sub Remove(index As Long)
    myCustomCollection.Remove (index)
End Sub

Public Sub Clear()
    While myCustomCollection.Count <> 0
        myCustomCollection.Remove (myCustomCollection.Count)
    Wend
End Sub

Public Property Get Item(index As Long) As IRiskMeasure
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

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    Set myCustomCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set myCustomCollection = Nothing
End Sub
