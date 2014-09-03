VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Curve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private pName As String
Private pCurveDate As Date
Private pRates As Scripting.Dictionary
Private pMaxIndex As Integer

'=============================================================
'
'   PROPERTIES
'
'=============================================================
Public Property Get Name()
    Name = pName
End Property

'=============================================================
'
'   METHODS
'
'=============================================================
Public Sub Initialize(curveDate As Date, Name As String)
    pCurveDate = curveDate
    pName = Name
End Sub

Public Sub AddRate(tenor As Period, rate As Double)
    Dim tenorDate As Date: tenorDate = CurveTenorDate(pCurveDate, tenor)
    
    If Not pRates.Exists(tenorDate) Then
        pRates.Add tenorDate, rate
    End If
    
    pMaxIndex = pRates.Count - 1

End Sub

Public Function GetRate(rateDate As Date) As Double
    If pRates.Exists(rateDate) Then
        GetRate = pRates.Item(rateDate)
    Else
        GetRate = GetInterpolatedRate(rateDate)
    End If
End Function

Public Function GetRateForTenor(tenor As Period) As Double
    GetRateForTenor = GetRate(CurveTenorDate(pCurveDate, tenor))
End Function

Public Function GetForwardRate(valuationDate As Date, startDate As Date, endDate As Date, dcc As DayCountConvention) As Double
        
    Dim d1 As Double, d2 As Double
    Dim r1 As Double, r2 As Double
            
    d1 = YearFrac(valuationDate, startDate, dcc)
    d2 = YearFrac(valuationDate, endDate, dcc)
    r1 = GetRate(startDate)
    r2 = GetRate(endDate)
    
    'Rates are continuously compounded
    GetForwardRate = (r2 * d2 - r1 * d1) / (d2 - d1)

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift entire curve in parralell
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftParallel(shift As Double)
    Dim v As Variant
    
    For Each v In pRates.keys
       pRates.Item(v) = pRates.Item(v) + shift
    Next v

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Shift single rate on a given tenor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShiftRateOnTenor(inTenor As Period, shift As Double)
    Dim tenorDate As Date: tenorDate = CurveTenorDate(pCurveDate, inTenor)

    If Not pRates.Exists(tenorDate) Then
        Err.Raise vbObjectError + 513, "Curve::ShiftRateOnTenor", "Tenor not present in curve: " + pName
    Else
       pRates.Item(tenorDate) = pRates.Item(tenorDate) + shift
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Human readable string representation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
    Dim v As Variant
    Dim res As String: res = pName + " as of " + CStr(pCurveDate) + vbNewLine
    
    For Each v In pRates.keys
        res = res + CStr(v) + " " + CStr(pRates(v)) + vbNewLine
    Next
    ToString = res
End Function


'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    Set pRates = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set pRates = Nothing
End Sub

'=============================================================
'
'   PRIVATE FUNCTIONS
'
'=============================================================
Private Function GetInterpolatedRate(rateDate As Date) As Double
    
    Dim dates() As Variant: dates = pRates.keys
    Dim res As Double
    
    If rateDate < pCurveDate Then
        res = -1
    ElseIf rateDate < dates(0) Then
        res = pRates.Item(dates(0))
    ElseIf rateDate > dates(pMaxIndex) Then
        res = pRates.Item(dates(pMaxIndex))
    Else
        Dim i As Integer: i = 0
        
        While rateDate > dates(i)
            i = i + 1
        Wend

        res = Interpolate(rateDate, CDate(dates(i - 1)), CDate(dates(i)), _
                          pRates.Item(dates(i - 1)), pRates.Item(dates(i)))
    End If
    
    GetInterpolatedRate = res
    
End Function

Private Function Interpolate(t As Date, T_i As Date, T_j As Date, r_i As Double, r_j As Double)
    Interpolate = r_i + (t - T_i) / (T_j - T_i) * (r_j - r_i)
End Function

