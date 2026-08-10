Imports System.IO

Public NotInheritable Class AppConfiguration
    Private Sub New()
    End Sub

    Public Shared Function GetEnvironment(name As String, Optional defaultValue As String = "") As String
        Dim value = Environment.GetEnvironmentVariable(name)
        If String.IsNullOrWhiteSpace(value) Then Return defaultValue
        Return value.Trim()
    End Function

    Public Shared ReadOnly Property SqlitePath As String
        Get
            Dim configuredPath = GetEnvironment("EMAILSCREENER_SQLITE_PATH")
            If Not String.IsNullOrWhiteSpace(configuredPath) Then Return configuredPath

            Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            Return Path.Combine(appData, "EmailScreener", "EmailScreener.db")
        End Get
    End Property

    Public Shared ReadOnly Property SqlConnectionString As String
        Get
            Return GetEnvironment("EMAILSCREENER_SQL_CONNECTION")
        End Get
    End Property

    Public Shared Function BuildAccount(name As String,
                                        imapHost As String,
                                        userVariable As String,
                                        passwordVariable As String) As String
        Return String.Join(",", name, imapHost, GetEnvironment(userVariable), GetEnvironment(passwordVariable))
    End Function

    Public Shared Function GetSmtpHost(imapHost As String) As String
        Select Case imapHost.ToLowerInvariant()
            Case "imap.gmail.com"
                Return "smtp.gmail.com"
            Case "imap.mail.yahoo.com"
                Return "smtp.mail.yahoo.com"
            Case Else
                Return GetEnvironment("EMAILSCREENER_SMTP_HOST")
        End Select
    End Function
End Class
