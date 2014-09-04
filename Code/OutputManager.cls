VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputManager"
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
    
Private elapsed_time_ As Double

Friend Sub SetValues(im As InputManager)
End Sub

Friend Sub RegisterValues(rm As RiskManager)
End Sub

Friend Property Let Elapsed_Time(et As Double)
    elapsed_time_ = et
End Property


'=============================================================
'
'   STRUCUTURAL
'
'=============================================================
Private Sub Class_Initialize()
    
    Worksheets("RAPORT").Range("outputArea").ClearContents
    Worksheets("StressTest").Range("sensivityArea").ClearContents
    
End Sub

Private Sub Class_Terminate()
    Call Write_out_data
End Sub

Private Sub Write_out_data()
    Worksheets("RAPORT").Range("elapsedTime").Value = elapsed_time_
End Sub
