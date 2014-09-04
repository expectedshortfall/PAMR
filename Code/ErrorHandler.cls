VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Handle_error()
    Dim msg As String
    Dim ErrorNum As Long: ErrorNum = Err.Number - vbObjectError
    
    If ErrorNum > 0 And ErrorNum < 65535 Then
        msg = "Application defined error" & vbCrLf & _
                "AppErr is: " & ErrorNum & vbCrLf & _
                "Source: " & Err.Source & vbCrLf & _
                "Description: " & Err.Description
    Else
        msg = "Visual Basic error" & vbCrLf & _
                "Number is: " & Err.Number & vbCrLf & _
                "Source: " & Err.Source & vbCrLf & _
                "Description: " & Err.Description
    End If
    Call MsgBox(msg, , "Object Error")
End Sub

