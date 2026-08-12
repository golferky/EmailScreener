Imports Microsoft.Data.Sqlite

Public Class MailAccount
    Public Property Id As Integer
    Public Property DisplayName As String
    Public Property ImapHost As String
    Public Property SmtpHost As String
    Public Property UserName As String
    Public Property PasswordEnvironmentVariable As String
    Public Property Enabled As Boolean

    Public ReadOnly Property PasswordStatus As String
        Get
            Return If(String.IsNullOrWhiteSpace(AppConfiguration.GetEnvironment(PasswordEnvironmentVariable)), "Missing", "Configured")
        End Get
    End Property

    Public Function GetPassword() As String
        Return AppConfiguration.GetEnvironment(PasswordEnvironmentVariable)
    End Function
End Class

Public Class MailAccountRepository
    Private ReadOnly connectionString As String

    Public Sub New(databasePath As String)
        Dim directory = IO.Path.GetDirectoryName(databasePath)
        If Not String.IsNullOrWhiteSpace(directory) AndAlso Not IO.Directory.Exists(directory) Then
            IO.Directory.CreateDirectory(directory)
        End If
        connectionString = $"Data Source={databasePath};"
        EnsureSchemaAndDefaults()
    End Sub

    Private Sub EnsureSchemaAndDefaults()
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Dim sql = "
                CREATE TABLE IF NOT EXISTS MailAccounts (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    DisplayName TEXT NOT NULL COLLATE NOCASE UNIQUE,
                    ImapHost TEXT NOT NULL,
                    SmtpHost TEXT NOT NULL,
                    UserName TEXT NOT NULL,
                    PasswordEnvironmentVariable TEXT NOT NULL,
                    Enabled INTEGER NOT NULL DEFAULT 1
                );"
            Using command As New SqliteCommand(sql, connection)
                command.ExecuteNonQuery()
            End Using

            Using countCommand As New SqliteCommand("SELECT COUNT(*) FROM MailAccounts", connection)
                If CLng(countCommand.ExecuteScalar()) = 0 Then
                    InsertDefault(connection, "Yahoo", "imap.mail.yahoo.com", "smtp.mail.yahoo.com",
                                  AppConfiguration.GetEnvironment("EMAILSCREENER_YAHOO_USER"), "EMAILSCREENER_YAHOO_APP_PASSWORD")
                    InsertDefault(connection, "Gmail", "imap.gmail.com", "smtp.gmail.com",
                                  AppConfiguration.GetEnvironment("EMAILSCREENER_GMAIL_USER"), "EMAILSCREENER_GMAIL_APP_PASSWORD")
                End If
            End Using

            Using migration As New SqliteCommand("UPDATE ForwardedEmails SET AccountName='Gmail' WHERE AccountName='*Gmail'", connection)
                Try
                    migration.ExecuteNonQuery()
                Catch ex As SqliteException
                    ' Forwarding tables are created separately during startup.
                End Try
            End Using
        End Using
    End Sub

    Private Shared Sub InsertDefault(connection As SqliteConnection,
                                     displayName As String,
                                     imapHost As String,
                                     smtpHost As String,
                                     userName As String,
                                     passwordVariable As String)
        Dim sql = "INSERT INTO MailAccounts (DisplayName, ImapHost, SmtpHost, UserName, PasswordEnvironmentVariable, Enabled) " &
                  "VALUES (@name, @imap, @smtp, @user, @passwordVariable, 1)"
        Using command As New SqliteCommand(sql, connection)
            command.Parameters.AddWithValue("@name", displayName)
            command.Parameters.AddWithValue("@imap", imapHost)
            command.Parameters.AddWithValue("@smtp", smtpHost)
            command.Parameters.AddWithValue("@user", userName)
            command.Parameters.AddWithValue("@passwordVariable", passwordVariable)
            command.ExecuteNonQuery()
        End Using
    End Sub

    Public Function LoadAccounts(Optional includeDisabled As Boolean = True) As List(Of MailAccount)
        Dim accounts As New List(Of MailAccount)
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Dim sql = "SELECT Id, DisplayName, ImapHost, SmtpHost, UserName, PasswordEnvironmentVariable, Enabled FROM MailAccounts"
            If Not includeDisabled Then sql &= " WHERE Enabled=1"
            sql &= " ORDER BY DisplayName"
            Using command As New SqliteCommand(sql, connection)
                Using reader = command.ExecuteReader()
                    While reader.Read()
                        accounts.Add(New MailAccount With {
                            .Id = reader.GetInt32(0),
                            .DisplayName = reader.GetString(1),
                            .ImapHost = reader.GetString(2),
                            .SmtpHost = reader.GetString(3),
                            .UserName = reader.GetString(4),
                            .PasswordEnvironmentVariable = reader.GetString(5),
                            .Enabled = reader.GetInt32(6) <> 0
                        })
                    End While
                End Using
            End Using
        End Using
        Return accounts
    End Function

    Public Function SaveAccount(account As MailAccount) As Integer
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            If account.Id = 0 Then
                Dim insertSql = "INSERT INTO MailAccounts (DisplayName, ImapHost, SmtpHost, UserName, PasswordEnvironmentVariable, Enabled) " &
                                "VALUES (@name, @imap, @smtp, @user, @passwordVariable, @enabled); SELECT last_insert_rowid();"
                Using command As New SqliteCommand(insertSql, connection)
                    AddParameters(command, account)
                    Return CInt(CLng(command.ExecuteScalar()))
                End Using
            End If

            Dim updateSql = "UPDATE MailAccounts SET DisplayName=@name, ImapHost=@imap, SmtpHost=@smtp, UserName=@user, " &
                            "PasswordEnvironmentVariable=@passwordVariable, Enabled=@enabled WHERE Id=@id"
            Using command As New SqliteCommand(updateSql, connection)
                AddParameters(command, account)
                command.Parameters.AddWithValue("@id", account.Id)
                command.ExecuteNonQuery()
            End Using
            Return account.Id
        End Using
    End Function

    Private Shared Sub AddParameters(command As SqliteCommand, account As MailAccount)
        command.Parameters.AddWithValue("@name", account.DisplayName.Trim())
        command.Parameters.AddWithValue("@imap", account.ImapHost.Trim())
        command.Parameters.AddWithValue("@smtp", account.SmtpHost.Trim())
        command.Parameters.AddWithValue("@user", account.UserName.Trim())
        command.Parameters.AddWithValue("@passwordVariable", account.PasswordEnvironmentVariable.Trim())
        command.Parameters.AddWithValue("@enabled", If(account.Enabled, 1, 0))
    End Sub

    Public Sub DeleteAccount(id As Integer)
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Using command As New SqliteCommand("DELETE FROM MailAccounts WHERE Id=@id", connection)
                command.Parameters.AddWithValue("@id", id)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class
