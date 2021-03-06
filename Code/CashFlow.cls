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

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pOrigin As Origin
Private pCCY As CCY
Private pAmount As Double
Private pDiscountFactor As DiscountFactor
Private pDueDate As Date

'=============================================================
'
'   METHODS
'
'=============================================================
Public Function GetNPV(ValueDate As Date, M As MarketState) As Double
    Dim res As Double
    
    If ValueDate <= Me.DueDate Then
        res = Me.Amount * pDiscountFactor.Calculate(ValueDate, Me.DueDate, M)
    Else
        res = Me.Amount
    End If
        
    GetNPV = res
    
End Function

Public Function GetNPVinPLN(ValueDate As Date, M As MarketState) As Double
    
    GetNPVinPLN = M.GetCcyRate(pCCY) * GetNPV(ValueDate, M)
    
End Function

'=============================================================
'
'   PROPERTIES
'
'=============================================================
''''''''''''''''''''''
' Origin property
''''''''''''''''''''''
Public Property Get Origin() As Origin
    Origin = pOrigin
End Property

''''''''''''''''''''''
' Currency property
''''''''''''''''''''''
Public Property Get GetCCY() As CCY
    GetCCY = pCCY
End Property


''''''''''''''''''''''
' Amount property
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialization, should be used only in factory module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialize(inOrigin As Origin, inCcy As CCY, _
                      inValue As Double, inValueDate As Date, _
                      inDF As DiscountFactor)

    pOrigin = inOrigin
    pCCY = inCcy
    pAmount = inValue
    pDueDate = inValueDate
    
    Set pDiscountFactor = inDF
    
End Sub



