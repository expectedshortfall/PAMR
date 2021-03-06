VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mBPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRiskMeasure

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Dim pValue As Double
Dim pName As String
Dim pCCY As CCY

'=============================================================
'
'   METHODS
'
'=============================================================
'=============================================================
'
'   IMPLEMENTING IRiskMeasure INTERFACE
'
'=============================================================
Public Function IRiskMeasure_Calculate(ds As IDataSource)
    
    Dim MSP As MarketStateProvider: Set MSP = ds.MarketStateProvider
    Dim p As Portfolio: Set p = ds.Portfolio
    
    Dim shiftedMSP As MarketStateProvider
    
    If VarType(ds.GetDataFromCache("BPV")) = vbNull Then
        Set shiftedMSP = MSP.ShiftMarketParallel(1)
        ds.AddDataToCache "BPV", shiftedMSP
    Else
        Set shiftedMSP = ds.GetDataFromCache("BPV")
    End If
    
    Dim currentValue As Double: currentValue = p.GetNPVByCCY(ds.ValueDate, MSP, pCCY)
    Dim shiftedValue As Double: shiftedValue = p.GetNPVByCCY(ds.ValueDate, shiftedMSP, pCCY)
    
    pValue = (shiftedValue - currentValue)

End Function

Public Sub IRiskMeasure_Initialize(param As Variant)
    
    If Not VarType(param) = vbString Then
        Err.Raise vbObject + 513, "BPV::IRiskMeasure_Initialize", "Param is not a string"
    End If
    
    pCCY = Types.StringToCCY(CStr(param))
    pName = pName + param

End Sub

Public Property Get IRiskMeasure_Name() As String
    IRiskMeasure_Name = pName
End Property

Public Property Get IRiskMeasure_Value() As Double
    IRiskMeasure_Value = pValue
End Property

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    pName = "BPV_"
End Sub
