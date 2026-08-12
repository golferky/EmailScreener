Imports System.IO

Public NotInheritable Class AppConfiguration
    Private Sub New()
    End Sub

    Public Shared Function GetEnvironment(name As String, Optional defaultValue As String = "") As String
        If String.IsNullOrWhiteSpace(name) Then Return defaultValue
        Dim value = Environment.GetEnvironmentVariable(name)
        If String.IsNullOrWhiteSpace(value) Then Return defaultValue
        Return value.Trim()
    End Function

    Public Shared Sub SetUserEnvironment(name As String, value As String)
        If String.IsNullOrWhiteSpace(name) Then Throw New ArgumentException("An environment-variable name is required.", NameOf(name))
        Environment.SetEnvironmentVariable(name, value, EnvironmentVariableTarget.Process)
        If OperatingSystem.IsWindows() Then
            Environment.SetEnvironmentVariable(name, value, EnvironmentVariableTarget.User)
        End If
    End Sub

    Public Shared Function BuildPasswordVariable(displayName As String) As String
        Dim safeName = System.Text.RegularExpressions.Regex.Replace(displayName.Trim().ToUpperInvariant(), "[^A-Z0-9]+", "_").Trim("_"c)
        If String.IsNullOrWhiteSpace(safeName) Then safeName = "MAIL_ACCOUNT"
        Return $"EMAILSCREENER_{safeName}_APP_PASSWORD"
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

    Public Shared Function GetAccountEmailAddress(imapHost As String, userName As String) As String
        Dim address = If(userName, String.Empty).Trim()
        If address.Contains("@"c) Then Return address
        Select Case If(imapHost, String.Empty).ToLowerInvariant()
            Case "imap.gmail.com"
                Return $"{address}@gmail.com"
            Case "imap.mail.yahoo.com"
                Return $"{address}@yahoo.com"
            Case Else
                Return address
        End Select
    End Function
End Class
