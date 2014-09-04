VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppObjWrapper"
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
Private data_in As InputManager
Private data_out As OutputManager
Private app As RiskManager
Private watch As StopWatch

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Public Sub Run()

    Call watch.Start_Timer
    Call app.Run
    Call watch.Stop_Timer

    data_out.RegisterValues app
    data_out.Elapsed_Time = watch.Elapsed_Time
        
End Sub


'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    
    Set data_in = New InputManager
    Set data_out = New OutputManager
    Set app = New RiskManager:
    
    Call data_out.SetValues(data_in)
    Call app.SetValues(data_in)
    
    Set watch = New StopWatch
       
End Sub

Private Sub Class_Terminate()

    Set watch = Nothing
    Set app = Nothing
    Set data_out = Nothing
    Set data_in = Nothing

End Sub



