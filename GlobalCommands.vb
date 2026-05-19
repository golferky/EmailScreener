Public Module GlobalCommands
    Public Sub LOGIT(sMess As String, Optional bTrbl As Boolean = False)
        ' This points to the code inside your LeagueContext class
        Context.Instance.LOGIT(sMess, bTrbl)
    End Sub
End Module