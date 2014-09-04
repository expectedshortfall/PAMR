VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
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
Private Timer_running_ As Boolean
Private current_begin_time_ As Double

'=============================================================
'
'   METHODS
'
'=============================================================
Friend Sub Start_Timer()

    If Not Timer_running_ Then
        current_begin_time_ = Timer
        Timer_running_ = True
    End If

End Sub

Friend Sub Stop_Timer()

    If Timer_running_ Then
        elapsed_time_ = elapsed_time_ + Timer - current_begin_time_
        Timer_running_ = False
    End If

End Sub

Friend Sub Reset_Timer()

    elapsed_time_ = 0#
    Timer_running_ = False
    
End Sub

Friend Property Get Elapsed_Time() As Double

    Dim Extra_time As Double
    
    If Timer_running_ Then
        Extra_time = Timer - current_begin_time_
    Else
        Extra_time = 0#
    End If
    
    Elapsed_Time = elapsed_time_ + Extra_time

End Property

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    
    elapsed_time_ = 0#
    Timer_running_ = False

End Sub
