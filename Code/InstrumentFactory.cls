VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstrumentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function LoadInstruments() As InstrumentCollection
    
    Dim wb As Workbook, ws As Worksheet
    Dim instruments As InstrumentCollection: Set instruments = New InstrumentCollection
    
    Application.ScreenUpdating = False
    
        Set wb = Workbooks.Open(ThisWorkbook.Path + "\deals.xls", True, True)
        For Each ws In wb.Worksheets
            With ws
                CallByName Me, ("Get" + .Name), VbMethod, instruments, ws
            End With
        Next
        
        wb.Close False: Set wb = Nothing
    
    Application.ScreenUpdating = True
    
    Set LoadInstruments = instruments

End Function
    

Public Sub GetIRS_CIRS(ByRef instruments As InstrumentCollection, ws As Worksheet)
    
    instruments.Add Factory.CreateIRS_CIRS()
        
    'Debug.Print "GetIRS_CIRS"

End Sub

Public Sub GetFRA(ByRef instruments As InstrumentCollection, ws As Worksheet)
    Dim i As Integer: i = 2
    
    With ws
        Do Until .Cells(i, 1).Value = ""
        
        Dim discountCurveName As String: discountCurveName = .Cells(i, 5)
        Dim forwardCurveName As String: forwardCurveName = IIf(.Cells(i, 5) = "PLN", "PL3", .Cells(i, 5))
        Dim dcc As DayCountConvention: dcc = IIf(.Cells(i, 5) = "PLN", cAct365, cAct360)
        
        instruments.Add Factory.CreateFRA(.Cells(i, 1).Value, .Cells(i, 2).Value, _
                                          .Cells(i, 3).Value, .Cells(i, 4).Value, _
                                          Types.StringToCCY(.Cells(i, 5).Value), _
                                          Types.StringToPosition(.Cells(i, 6).Value), _
                                          .Cells(i, 7), .Cells(i, 8), _
                                          Factory.CreateRateManager(discountCurveName, forwardCurveName, dcc))
        i = i + 1
    Loop
    End With
    
    Debug.Print "Loading FRAs has finished successfully"

End Sub

Public Sub GetFXSpot(ByRef instruments As InstrumentCollection, ws As Worksheet)
    'Debug.Print "GetFXSpot"
End Sub

Public Sub GetFXSwap(ByRef instruments As InstrumentCollection, ws As Worksheet)
    'Debug.Print "GetFXSwap"
End Sub

Public Sub GetFXOption(ByRef instruments As InstrumentCollection, ws As Worksheet)
    'Debug.Print "GetFXOption"
End Sub

Public Sub GetFWD_NDF(ByRef instruments As InstrumentCollection, ws As Worksheet)
    'Debug.Print "GetFWD_NDF"
End Sub

Public Sub GetBond(ByRef instruments As InstrumentCollection, ws As Worksheet)
    'Debug.Print "GetBond"
End Sub

Public Sub GetFBond(ByRef instruments As InstrumentCollection, ws As Worksheet)
    'Debug.Print "GetFBond"
End Sub


