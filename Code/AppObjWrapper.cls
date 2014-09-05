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
Private IM As InputManager
Private OM As OutputManager
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

    OM.RegisterValues app
    OM.ElapsedTime = watch.Elapsed_Time
        
End Sub


'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    
    Set watch = New StopWatch
    Set OM = New OutputManager
    Set app = New RiskManager:
    
    ' We would like to know how long does the data loading take
    watch.Start_Timer
    Set IM = New InputManager
    watch.Stop_Timer
    
    Call OM.SetValues(IM): OM.DataLoadingTime = watch.Elapsed_Time: watch.Reset_Timer
    Call app.SetValues(IM)
       
End Sub

Private Sub Class_Terminate()

    Set watch = Nothing
    Set app = Nothing
    Set OM = Nothing
    Set IM = Nothing

End Sub



