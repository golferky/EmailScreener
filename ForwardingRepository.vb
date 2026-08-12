Imports Microsoft.Data.Sqlite

Public Class ForwardingRule
    Public Property Id As Integer
    Public Property MatchType As String
    Public Property MatchValue As String
    Public Property Destination As String
    Public Property Enabled As Boolean
End Class

Public Class ForwardingRepository
    Private ReadOnly connectionString As String

    Public Sub New(databasePath As String)
        Dim directory = IO.Path.GetDirectoryName(databasePath)
        If Not String.IsNullOrWhiteSpace(directory) AndAlso Not IO.Directory.Exists(directory) Then
            IO.Directory.CreateDirectory(directory)
        End If
        connectionString = $"Data Source={databasePath};"
        EnsureSchema()
    End Sub

    Private Sub EnsureSchema()
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Dim sql = "
                CREATE TABLE IF NOT EXISTS ForwardingRules (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    MatchType TEXT NOT NULL,
                    MatchValue TEXT NOT NULL,
                    Destination TEXT NOT NULL,
                    Enabled INTEGER NOT NULL DEFAULT 1
                );
                CREATE TABLE IF NOT EXISTS ForwardedEmails (
                    AccountName TEXT NOT NULL,
                    Folder TEXT NOT NULL,
                    Uid INTEGER NOT NULL,
                    Destination TEXT NOT NULL,
                    ForwardedAt TEXT NOT NULL,
                    PRIMARY KEY (AccountName, Folder, Uid, Destination)
                );
                CREATE TABLE IF NOT EXISTS AppSettings (
                    SettingKey TEXT PRIMARY KEY,
                    SettingValue TEXT NOT NULL
                );"
            Using command As New SqliteCommand(sql, connection)
                command.ExecuteNonQuery()
            End Using
            Using migration As New SqliteCommand("UPDATE ForwardingRules SET MatchType='Name or email contains' WHERE MatchType='Sender name'", connection)
                migration.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function LoadRules() As List(Of ForwardingRule)
        Dim rules As New List(Of ForwardingRule)
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Using command As New SqliteCommand("SELECT Id, MatchType, MatchValue, Destination, Enabled FROM ForwardingRules ORDER BY MatchType, MatchValue", connection)
                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rules.Add(New ForwardingRule With {
                            .Id = reader.GetInt32(0),
                            .MatchType = reader.GetString(1),
                            .MatchValue = reader.GetString(2),
                            .Destination = reader.GetString(3),
                            .Enabled = reader.GetInt32(4) <> 0
                        })
                    End While
                End Using
            End Using
        End Using
        Return rules
    End Function

    Public Function SaveRule(rule As ForwardingRule) As Integer
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            If rule.Id = 0 Then
                Dim sql = "INSERT INTO ForwardingRules (MatchType, MatchValue, Destination, Enabled) VALUES (@type, @value, @destination, @enabled); SELECT last_insert_rowid();"
                Using command As New SqliteCommand(sql, connection)
                    AddRuleParameters(command, rule)
                    Return CInt(CLng(command.ExecuteScalar()))
                End Using
            End If

            Dim updateSql = "UPDATE ForwardingRules SET MatchType=@type, MatchValue=@value, Destination=@destination, Enabled=@enabled WHERE Id=@id"
            Using command As New SqliteCommand(updateSql, connection)
                AddRuleParameters(command, rule)
                command.Parameters.AddWithValue("@id", rule.Id)
                command.ExecuteNonQuery()
            End Using
            Return rule.Id
        End Using
    End Function

    Private Shared Sub AddRuleParameters(command As SqliteCommand, rule As ForwardingRule)
        command.Parameters.AddWithValue("@type", rule.MatchType)
        command.Parameters.AddWithValue("@value", rule.MatchValue.Trim())
        command.Parameters.AddWithValue("@destination", rule.Destination.Trim())
        command.Parameters.AddWithValue("@enabled", If(rule.Enabled, 1, 0))
    End Sub

    Public Sub DeleteRule(id As Integer)
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Using command As New SqliteCommand("DELETE FROM ForwardingRules WHERE Id=@id", connection)
                command.Parameters.AddWithValue("@id", id)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function IsForwarded(accountName As String, folder As String, uid As Integer, destination As String) As Boolean
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Dim sql = "SELECT 1 FROM ForwardedEmails WHERE AccountName=@account AND Folder=@folder AND Uid=@uid AND Destination=@destination LIMIT 1"
            Using command As New SqliteCommand(sql, connection)
                command.Parameters.AddWithValue("@account", accountName)
                command.Parameters.AddWithValue("@folder", folder)
                command.Parameters.AddWithValue("@uid", uid)
                command.Parameters.AddWithValue("@destination", destination)
                Return command.ExecuteScalar() IsNot Nothing
            End Using
        End Using
    End Function

    Public Sub RecordForwarded(accountName As String, folder As String, uid As Integer, destination As String)
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Dim sql = "INSERT OR IGNORE INTO ForwardedEmails (AccountName, Folder, Uid, Destination, ForwardedAt) VALUES (@account, @folder, @uid, @destination, @forwardedAt)"
            Using command As New SqliteCommand(sql, connection)
                command.Parameters.AddWithValue("@account", accountName)
                command.Parameters.AddWithValue("@folder", folder)
                command.Parameters.AddWithValue("@uid", uid)
                command.Parameters.AddWithValue("@destination", destination)
                command.Parameters.AddWithValue("@forwardedAt", DateTimeOffset.UtcNow.ToString("O"))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function GetAutoForwardEnabled() As Boolean
        Return String.Equals(GetSetting("AutoForwardEnabled"), "true", StringComparison.OrdinalIgnoreCase)
    End Function

    Public Sub SetAutoForwardEnabled(enabled As Boolean)
        SetSetting("AutoForwardEnabled", If(enabled, "true", "false"))
    End Sub

    Public Function GetSetting(key As String, Optional defaultValue As String = "") As String
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Using command As New SqliteCommand("SELECT SettingValue FROM AppSettings WHERE SettingKey=@key", connection)
                command.Parameters.AddWithValue("@key", key)
                Dim value = TryCast(command.ExecuteScalar(), String)
                If value Is Nothing Then Return defaultValue
                Return value
            End Using
        End Using
    End Function

    Public Sub SetSetting(key As String, value As String)
        Using connection As New SqliteConnection(connectionString)
            connection.Open()
            Dim sql = "INSERT INTO AppSettings (SettingKey, SettingValue) VALUES (@key, @value) ON CONFLICT(SettingKey) DO UPDATE SET SettingValue=excluded.SettingValue"
            Using command As New SqliteCommand(sql, connection)
                command.Parameters.AddWithValue("@key", key)
                command.Parameters.AddWithValue("@value", value)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class
