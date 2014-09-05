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
Private pElapsed_time_ As Double
Private pData_loading_time_ As Double
Private pCalculationOutcomes As Collection

Friend Sub SetValues(IM As InputManager)
End Sub

Friend Sub RegisterValues(RM As RiskManager)
    RM.RegisterValues pCalculationOutcomes
End Sub

Friend Property Let ElapsedTime(et As Double)
    pElapsed_time_ = et
End Property

Friend Property Let DataLoadingTime(dlt As Double)
    pData_loading_time_ = dlt
End Property

'=============================================================
'
'   STRUCUTURAL
'
'=============================================================
Private Sub Class_Initialize()
    
    Set pCalculationOutcomes = New Collection
    Call ClearWorksheets
    
End Sub

Private Sub Class_Terminate()
    Call Write_out_data
    Call ClearMemory
End Sub

Private Sub Write_out_data()
    
    With Worksheets("RAPORT")
        .Range("elapsedTime").Value = pElapsed_time_
        .Range("dataLoadingTime").Value = pData_loading_time_
    
        Dim r As Range: Set r = .Range("outputAreaStart")
        Dim v As Variant
        For Each v In pCalculationOutcomes
            r.Value = v
            Set r = r.Offset(1, 0)
        Next v
    End With

End Sub

Private Sub ClearMemory()
    Set pCalculationOutcomes = Nothing
End Sub

Private Sub ClearWorksheets()
    
    With Worksheets("RAPORT")
        .Range("outputArea").ClearContents
        .Range("elapsedTime").ClearContents
        .Range("dataLoadingTime").ClearContents
    End With
    
    Worksheets("StressTest").Range("sensivityArea").ClearContents

End Sub
