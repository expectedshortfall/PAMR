Attribute VB_Name = "Main"
Option Explicit

Public Sub GenerateReport()

    On Error GoTo Error_Handling
        
        Dim err_handler As ErrorHandler:    Set err_handler = New ErrorHandler
        Dim the_app As AppObjWrapper:        Set the_app = New AppObjWrapper
        Call the_app.Run
    
Clear_Up:
        Set the_app = Nothing
        Set err_handler = Nothing
        Exit Sub
    
Error_Handling:
        Call err_handler.Handle_error
        Resume Clear_Up

End Sub

