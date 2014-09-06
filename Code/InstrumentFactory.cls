VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstrumentFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function LoadInstruments() As InstrumentCollection
    
    Dim wb As Workbook, ws As Worksheet
    Dim instruments As InstrumentCollection: Set instruments = New InstrumentCollection
    
    Application.ScreenUpdating = False
    
        Set wb = Workbooks.Open(ThisWorkbook.Path + "\deals.xls", True, True)
        For Each ws In wb.Worksheets
            Dim i As Integer: i = 2
            With ws
                Do Until .Cells(i, 1).Value = ""
                    CallByName Me, ("Get" + .Name), VbMethod, instruments, .Cells(i, 1)
                    i = i + 1
                Loop
            End With
        Next
        
        wb.Close False: Set wb = Nothing
    
    Application.ScreenUpdating = True
    
    Set LoadInstruments = instruments

End Function
    

Public Sub GetIRS_CIRS(ByRef instruments As InstrumentCollection, r As Range)

    instruments.Add Factory.CreateIRS_CIRS()

End Sub

Public Sub GetFRA(ByRef instruments As InstrumentCollection, r As Range)
        
    With r
        
        Dim discountCurveName As String: discountCurveName = .Offset(0, 4)
        Dim forwardCurveName As String: forwardCurveName = IIf(.Offset(0, 4) = "PLN", "PL3", .Offset(0, 4))
        Dim dcc As DayCountConvention: dcc = IIf(.Offset(0, 4) = "PLN", cAct365, cAct360)
        
        instruments.Add Factory.CreateFRA(.Offset(0, 0).Value, .Offset(0, 1).Value, _
                                          .Offset(0, 2).Value, .Offset(0, 3).Value, _
                                          Types.StringToCCY(.Offset(0, 4).Value), _
                                          Types.StringToPosition(.Offset(0, 5).Value), _
                                          .Offset(0, 6), .Offset(0, 7), _
                                          Factory.CreateRateManager(discountCurveName, forwardCurveName, dcc))
        
    End With
    
    Debug.Print "Loading FRAs has finished successfully"
    
End Sub

Public Sub GetFXSpot(ByRef instruments As InstrumentCollection, r As Range)
    
    With r

        Dim dcName1 As String: dcName1 = IIf(.Offset(0, 4) = "PLN", "PLX", .Offset(0, 4))     'discount curve name
        Dim fcName1 As String: fcName1 = IIf(.Offset(0, 4) = "PLN", "PLX", .Offset(0, 4))     'forward curve name
        Dim dcc1 As DayCountConvention: dcc1 = IIf(.Offset(0, 4) = "PLN", cAct365, cAct360)  'day count convention

        Dim dcName2 As String: dcName2 = IIf(.Offset(0, 4) = "PLN", "PLX", .Offset(0, 6))
        Dim fcName2 As String: fcName2 = IIf(.Offset(0, 4) = "PLN", "PLX", .Offset(0, 6))
        Dim dcc2 As DayCountConvention: dcc2 = IIf(.Offset(0, 6) = "PLN", cAct365, cAct360)

        instruments.Add Factory.CreateFXSpot(.Offset(0, 0).Value, .Offset(0, 1).Value, StringToPosition(.Offset(0, 2).Value), _
                .Offset(0, 3).Value, StringToCCY(.Offset(0, 4).Value), Factory.CreateRateManager(dcName1, fcName1, dcc1), _
                .Offset(0, 5).Value, StringToCCY(.Offset(0, 6).Value), Factory.CreateRateManager(dcName2, fcName2, dcc2))

    End With
    
    Debug.Print "Loading FXSpots has finished successfully"

End Sub

Public Sub GetFXSwap(ByRef instruments As InstrumentCollection, r As Range)
    'Debug.Print "GetFXSwap"
End Sub

Public Sub GetFXOption(ByRef instruments As InstrumentCollection, r As Range)

    With r
        
        Dim CCY1 As String: CCY1 = Left(Trim(.Offset(0, 5).Value), 3)
        Dim dcc1 As DayCountConvention: dcc1 = IIf(CCY1 = "PLN", cAct365, cAct360)
        
        Dim CCY2 As String: CCY2 = Left(Trim(.Offset(0, 9).Value), 3)
        Dim dcc2 As DayCountConvention: dcc2 = IIf(CCY2 = "PLN", cAct365, cAct360)
        
        instruments.Add Factory.CreateFXOption(.Offset(0, 0).Value, .Offset(0, 1).Value, .Offset(0, 2).Value, StringToPosition(.Offset(0, 3).Value), _
                .Offset(0, 4).Value, StringToCCY(CCY1), .Offset(0, 6).Value, .Offset(0, 8).Value, _
                 Factory.CreateRateManager(CCY1, CCY1, dcc1), _
                 Factory.CreateRateManager(CCY2, CCY2, dcc2))
 
    End With

    Debug.Print "Loading FXOptions has finished successfully"

End Sub

Public Sub GetFWD_NDF(ByRef instruments As InstrumentCollection, r As Range)
    
    With r
    
        Dim CCY1 As String: CCY1 = .Offset(0, 4).Value
        Dim dcName1 As String: dcName1 = IIf(CCY1 = "PLN", "PLX", CCY1)             'discount curve name
        Dim fcName1 As String: fcName1 = dcName1                                    'forward curve name
        Dim dcc1 As DayCountConvention: dcc1 = IIf(CCY1 = "PLN", cAct365, cAct360)
        
        Dim CCY2 As String: CCY2 = .Offset(0, 5).Value
        Dim dcName2 As String: dcName2 = IIf(CCY2 = "PLN", "PLX", CCY2)             'discount curve name
        Dim fcName2 As String: fcName2 = dcName2                                    'forward curve name
        Dim dcc2 As DayCountConvention: dcc2 = IIf(CCY2 = "PLN", cAct365, cAct360)
        
        instruments.Add Factory.CreateFWD_NDF(.Offset(0, 0), .Offset(0, 1), StringToPosition(.Offset(0, 2)), _
                .Offset(0, 3), StringToCCY(CCY1), StringToCCY(CCY2), .Offset(0, 6), .Offset(0, 7), StringToDelivery(.Offset(0, 8)), _
                Factory.CreateRateManager(dcName1, fcName1, dcc1), _
                Factory.CreateRateManager(dcName2, fcName2, dcc2))
    
    End With
        
    Debug.Print "Loading FWD_NDF has finished successfully"

End Sub

Public Sub GetBond(ByRef instruments As InstrumentCollection, r As Range)
    'Debug.Print "GetBond"
End Sub

Public Sub GetFBond(ByRef instruments As InstrumentCollection, r As Range)
    'Debug.Print "GetFBond"
End Sub


