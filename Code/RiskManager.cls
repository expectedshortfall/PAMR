VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IDataSource

'=============================================================
'
'   PRIVATE VARIABLES
'
'=============================================================
Private pPortfolio As Portfolio
Private pMSP As MarketStateProvider
Private pVD As Date
Private pRiskMeasureCollection As IRiskMeasureCollection

'=============================================================
'
'   IMPLEMENTING IDATASOURCE INTERFACE
'
'=============================================================
Public Property Get IDataSource_ValueDate() As Date
    IDataSource_ValueDate = pVD
End Property

Public Property Get IDataSource_MarketStateProvider() As MarketStateProvider
    Set IDataSource_MarketStateProvider = pMSP
End Property

Public Property Get IDataSource_Portfolio() As Portfolio
    Set IDataSource_Portfolio = pPortfolio
End Property

'=============================================================
'
'   METHODS
'
'=============================================================
Friend Sub SetValues(ByRef IM As InputManager)
    
    Set pPortfolio = IM.GetPortfolio
    Set pMSP = IM.GetMarketStateProvider
        pVD = IM.GetValuationDate
        
    Call IM.LoadRiskMeasures(Me)

End Sub

Friend Sub AddRiskMeasure(ByRef RM As IRiskMeasure)
    pRiskMeasureCollection.Add RM
End Sub

Friend Sub Run()
    Dim RM As IRiskMeasure
    
    For Each RM In pRiskMeasureCollection
        RM.Calculate Me
    Next RM

End Sub

Friend Sub RegisterValues(ByRef outcomes As Collection)
    Dim v As IRiskMeasure
    For Each v In pRiskMeasureCollection
        outcomes.Add v.Value
    Next v
End Sub

'=============================================================
'
'   STRUCTURAL
'
'=============================================================
Private Sub Class_Initialize()
    Set pRiskMeasureCollection = New IRiskMeasureCollection
End Sub


Private Sub Class_Terminate()
    Set pPortfolio = Nothing
    Set pMSP = Nothing
    Set pRiskMeasureCollection = Nothing
End Sub
