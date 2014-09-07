VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================
'
'   PUBLIC IINTERFACE INTERFACE
'
'=============================================================
Public Property Get ValueDate() As Date
End Property

Public Property Get MarketStateProvider() As MarketStateProvider
End Property

Public Function GetDataFromCache(riskMeasureName As String) As Variant
End Function

Public Property Get MarketStateProviderShiftedInParallelBy1Bp() As MarketStateProvider
End Property

Public Property Get Portfolio() As Portfolio
End Property



