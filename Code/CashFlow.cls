VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pOrigin As String
Private pCCY As String
Private pAmount As Double
Private pDiscountFactor As DiscountFactor
Private pDueDate As Date

'=============================================================
'
'   PROPERTIES
'
'=============================================================

''''''''''''''''''''''
' Origin property
''''''''''''''''''''''
Public Property Get Origin() As String
    Origin = pOrigin
End Property

''''''''''''''''''''''
' Currency property
''''''''''''''''''''''
Public Property Get CCY() As String
    CCY = pCCY
End Property


''''''''''''''''''''''
' Origin property
''''''''''''''''''''''
Public Property Get Amount() As Double
    Amount = pAmount
End Property

''''''''''''''''''''''
' CashFlow Date property
''''''''''''''''''''''
Public Property Get DueDate() As Date
    DueDate = pDueDate
End Property

'=============================================================
'
'   METHODS
'
'=============================================================

Public Sub Initialize(inOrigin As String, inCCY As String, _
                      inValue As Double, inValueDate As Date, _
                      inDF As DiscountFactor)

    pOrigin = inOrigin
    pCCY = inCCY
    pAmount = inValue
    pDueDate = inValueDate
    
    Set pDiscountFactor = inDF
    
End Sub

Public Function GetNPV(valueDate As Date, M As MarketState) As Double
    Dim res As Double
    
    If valueDate <= Me.DueDate Then
        res = Me.Amount * pDiscountFactor.Calculate(valueDate, Me.DueDate, M)
    End If
        
    GetNPV = res
    
End Function

Public Function GetNPVinPLN(valueDate As Date, M As MarketState) As Double
    Dim conversionRate As Double: conversionRate = 1
    
    If Not pCCY = "PLN" Then
        conversionRate = M.GetCcyRate(pCCY + "PLN")
    End If
    
    GetNPVinPLN = conversionRate * GetNPV(valueDate, M)
    
End Function





