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
Private pMSP_BPV As MarketStateProvider
Private pVD As Date
Private pRiskMeasureCollection As IRiskMeasureCollection
Private pCache As Scripting.Dictionary

'=============================================================
'
'   IMPLEMENTING IDATASOURCE INTERFACE
'
'=============================================================
Public Property Get IDataSource_ValueDate() As Date
    IDataSource_ValueDate = pVD
End Property

Public Property Get IDataSource_MarketStateProvider() As MarketStateProvider
    If pMSP Is Nothing Then
        Err.Raise vbObject + 513, "RiskManager::MarketStateProvider", "MarketStateProvider is set to nothing"
    End If
    
    Set IDataSource_MarketStateProvider = pMSP

End Property

Public Function IDataSource_GetDataFromCache(key As String) As Variant
    
    If pCache.Exists(key) Then
        If VarType(pCache.Item(key)) = vbObject Then
            Set IDataSource_GetDataFromCache = pCache.Item(key)
        Else
            IDataSource_GetDataFromCache = pCache.Item(key)
        End If
    Else
        IDataSource_GetDataFromCache = Null
    End If

End Function

Public Sub IDataSource_AddDataToCache(key As String, data As Variant)
    If Not pCache.Exists(key) Then
        pCache.Add key, data
    End If
End Sub

Public Property Get IDataSource_Portfolio() As Portfolio
    
    If pPortfolio Is Nothing Then
        Err.Raise vbObject + 513, "RiskManager::Portfolio", "Portfolio is set to Nothing"
    End If
    
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
    Set pCache = New Scripting.Dictionary
End Sub


Private Sub Class_Terminate()
    Set pPortfolio = Nothing
    Set pMSP = Nothing
    Set pRiskMeasureCollection = Nothing
    Set pMSP_BPV = Nothing
    Set pCache = Nothing
End Sub
