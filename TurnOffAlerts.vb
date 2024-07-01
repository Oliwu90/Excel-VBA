Sub TurnOffAlerts()
With Application
            .DisplayAlerts = False
            .ScreenUpdating = False
            .DisplayStatusBar = False
            .EnableEvents = False
End With
End Sub
