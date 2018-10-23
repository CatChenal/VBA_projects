VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmPortfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================================================================
'
' CLSMDL sfrmPortfolio Nov-20-02 14:45
' Prev: May-23-02 14:05
'
'================================================================================
Private Sub Form_Current()
  With Me
    lngCurrentComp = !lngCompID
    lngCurrentForecast = !lngForecastID
    
    .Parent!cbxSelComp = lngCurrentComp
    .Parent!cbxSelForecast.Requery
    .Parent!cbxSelForecast = lngCurrentForecast
    .Parent!lblDefBud.Visible = Nz(.Parent!cbxSelForecast.Column(3), 0)
    .Parent!txtDefaultFRCDesc = GetCoDefaultBudgetDesc(lngCurrentComp)
  End With
End Sub

Private Sub Form_Load()
  Call SyncSubRecord("lngCompID", Me.Parent!cbxSelComp, "lngForecastID", lngCurrentForecast)
  DoCmd.Hourglass False
End Sub
