Imports System.IO
Imports System.Text
Imports System.Threading
Imports MailKit
Imports MailKit.Net.Imap
Imports MailKit.Search
Imports MailKit.Security
Imports Microsoft.Data.Sqlite
Imports Microsoft.Data.SqlClient
Public Class eUtilities

    Friend myprocessarray As New ArrayList
    Private myProcess As Process
    Public iLogitCounter As Integer = 0
    Public sLogPath As String = "C:\Log\"
    Public sEmailsCSV As String = $"C:\GarysEmails.CSV"
    Public sDomainsCSV As String = $"C:\GarysDomains.CSV"
    Public dtDomains As DataTable
    Public dtEmails As DataTable
    'Public dtMarkedRead As DataTable
    Public client As ImapClient
    Public sThisEmailServiceaName As String
    Public lEmailClients As List(Of String)
    Public sThisEmailService As String = ""
    Public sThisEmailUser As String = ""
    Public sThisEmailPassword As String = ""
    Public personal As IMailFolder
    Public Junk As IMailFolder
    Public subfolders As Object
    'field seperated by commas
    '1-Name
    '2-Format
    '   O-Date with Offset
    '   I-Integer
    '   D-Date with time YYYY-MM-DD HH:MM:SS
    '3-Key
    '4-Visible
    '5-Updateable
    Public sDomainFields As String = "
Name,,Y,,|
Trusted,,,,|
TotalEmails,I,,,|
LatestDate,O,K,,|
Senders,,,,"
    Public sEmailFields As String =
"
DateReceived,O,Y,,|
From,,Y,,|
Domain,,,N,|
Sender,,,N,|
Subject,,,,|
Uid,I,,N,|
Spam,B,,,Y|
Read,,,,|
PhishLevel,,,,Y|
Folder,,,N,|
DateAdded,D,,N,
"

    Public sYahoo As String = AppConfiguration.BuildAccount("Yahoo", "imap.mail.yahoo.com", "EMAILSCREENER_YAHOO_USER", "EMAILSCREENER_YAHOO_APP_PASSWORD")
    Public sGmail As String = AppConfiguration.BuildAccount("*Gmail", "imap.gmail.com", "EMAILSCREENER_GMAIL_USER", "EMAILSCREENER_GMAIL_APP_PASSWORD")
    Public conn As SqliteConnection
    Public sqlConn As String = AppConfiguration.SqlConnectionString
    Public sqlitePath As String = AppConfiguration.SqlitePath
    Public sqliteConnStr As String = $"Data Source={AppConfiguration.SqlitePath};"

    Sub New()
        lEmailClients = New List(Of String)
        lEmailClients.Add(sYahoo)
        lEmailClients.Add(sGmail)
    End Sub

    Sub openconn()
        If Not System.IO.File.Exists(sqlitePath) Then
            LOGIT("SQLite DB not found — skipping openconn")
            Exit Sub
        End If
        conn = New SqliteConnection(sqliteConnStr)
        Try
            conn.Open()
            Console.WriteLine("Connection opened successfully.")
        Catch ex As Exception
            Console.WriteLine("Connection failed: " & ex.Message)
        End Try
    End Sub
    Function UpdateIt(uidid As Integer, folder As String, message As MimeKit.MimeMessage, Optional spam As Boolean = False) As Boolean
        Try
            Dim utf8Encoding As New System.Text.UTF8Encoding(True)
            'Dim encodedString() As Byte
            Debug.Print(message.Subject)
            Dim subjectText As String = If(message.Subject, String.Empty)
            Dim encodedString As Byte() = utf8Encoding.GetBytes(subjectText)
            'Debug.Print($"Date:{message.Date} From:{message.From} Subject: {message.Subject}")
            'Dim x As String = message.From.ToString
            'If x.Contains("The Home Depot") Then
            '    Debug.Print("")
            'End If
            If spam Then
                If Not encodedString.Contains(173) Then
                    Return False
                End If
            End If
            Dim drow As DataRow = Nothing
            If dtEmails IsNot Nothing Then
                Dim sKey() As Object = {message.Date, message.From}
                drow = dtEmails.Rows.Find(sKey)
            End If
            If drow Is Nothing Then
                drow = dtEmails.NewRow
                drow("DateReceived") = message.Date
                '20251126-from is sometimes empty
                Dim sdomain As String = ""
                If message.From.Count > 0 Then
                    sdomain = message.From.ToString.Substring(message.From.ToString.IndexOf("@") + 1).Replace(">", "")
                    If UBound(sdomain.Split(" ")) > 0 Then
                        sdomain = sdomain.Split(" ")(1).Replace("<", "")
                    End If
                    drow("From") = message.From
                    drow("Sender") = message.From.ToString.Substring(0, message.From.ToString.IndexOf("@")).Replace("<", "")
                Else
                    drow("From") = ""
                End If

                drow("Domain") = sdomain
                drow("Subject") = message.Subject
                drow("EmailClient") = Emails.cmbEmailClients.SelectedItem
                'drow("TextBody") = message.TextBody
                drow("uid") = uidid
                If encodedString.Contains(173) Then
                    drow("Spam") = True
                Else
                    drow("Spam") = False
                End If
                drow("Read") = False
                drow("PhishLevel") = ""
                drow("Folder") = folder
                'drow("Headers") = message.Headers
                'drow("HtmlBody") = message.HtmlBody
                drow("DateAdded") = Now()
                dtEmails.Rows.Add(drow)
            End If
            UpdateEmailData(drow)
            If drow("Spam") = True Then
                'MoveToFolder(message)
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            LOGIT($"UpdateIt failed for UID {uidid}: {ex.Message}", True)
        End Try
        Return False
    End Function
    Function Connect() As IImapClient
        Try

            If String.IsNullOrWhiteSpace(sThisEmailUser) OrElse String.IsNullOrWhiteSpace(sThisEmailPassword) Then
                Throw New InvalidOperationException("Email credentials are not configured. See README.md for the required EMAILSCREENER environment variables.")
            End If

            'Dim oFiles() As IO.FileInfo
            'Dim oDirectory As New IO.DirectoryInfo("Y:\")
            'oFiles = oDirectory.GetFiles("*.csv")
            'EmailSetup()
            'BuildTables()
            'UpdateEmailTbl.ShowNamespaces(sThisEmailService, sThisEmailUser, sThisEmailPassword)

            If Not Directory.Exists(sLogPath) Then
                Directory.CreateDirectory(sLogPath)
            End If

            If client IsNot Nothing Then
                client.Dispose()
            End If

            Dim imapLogFile = $"{sLogPath}imap_{Process.GetCurrentProcess().Id}_{DateTime.Now:yyyyMMdd_HHmmss_fff}.log"
            client = New ImapClient(New ProtocolLogger(imapLogFile))
            'setup yahoo Generate App Password called Emailscreener and copy that password and paste here
            client.Connect(sThisEmailService, 993, SecureSocketOptions.SslOnConnect)
            Return client '$"Connected to {sThisEmailService} for User {sThisEmailUser}"
            'tsStatusText.Text = $"Connecting to {sThisEmailService} for User {sThisEmailUser}"
            'Application.DoEvents()
            'Debug.Print($"{client}-{client.IsConnected}")

        Catch ex As Exception
            MessageBox.Show($"{ex.Message}", "Connect Failed", MessageBoxButtons.OK)
            Return Nothing '$"Connect Failed {sThisEmailService} for User {sThisEmailUser}"
        End Try

    End Function
    Function Authenticate() As String
        'tsStatusText.Text = $"Authenticating User {sThisEmailUser}"
        'Application.DoEvents()
        Try
            If client Is Nothing OrElse Not client.IsConnected Then
                Return "Authentication failed: the email server is not connected."
            End If
            client.Authenticate(sThisEmailUser, sThisEmailPassword)
            Return $"Authenticated User {sThisEmailUser}"
        Catch ex As Exception
            Return $"Authentication failed for {sThisEmailUser}: {ex.Message}"
        End Try

    End Function
    Function OpenInbox(Optional readwrite As Boolean = False) As String
        Try
            If client Is Nothing OrElse Not client.IsConnected Then
                Return "Inbox open failed: the email server is not connected."
            End If
            If Not client.IsAuthenticated Then
                Return "Inbox open failed: the email account is not authenticated."
            End If
            If readwrite Then
                client.Inbox.Open(FolderAccess.ReadWrite)
            Else
                client.Inbox.Open(FolderAccess.ReadOnly)
            End If

            Return $"Inbox Opened for {client.Inbox.Access}"
        Catch ex As Exception
            Return $"Inbox open failed: {ex.Message}"
        End Try
        'tsStatusText.Text = $"Opening {client.Inbox.FullName}"
        'Application.DoEvents()

    End Function
    Sub SetupFolders()
        'Dim folder As IMailFolder = client.GetFolder(New FolderNamespace("/", "INBOX"))
        'folder.Open(FolderAccess.ReadOnly)
        'Dim subs = folder.GetSubfolders
        personal = client.GetFolder(client.PersonalNamespaces(0))
        subfolders = personal.GetSubfolders()
    End Sub
    Sub CreateFolder(client As ImapClient, sfolder As String)
        Dim personal = client.GetFolder(client.PersonalNamespaces(0))
        personal.Create(sfolder, True)
    End Sub
    Sub MoveToFolder(suid As Integer, folder As IMailFolder)
        'uid,destination folder
        folder.MoveTo(suid, folder)
    End Sub
    Sub MoveTojunk(client As ImapClient)
        MoveToPersonalFolder(client, "Junk")
    End Sub
    Sub MoveToPersonalFolder(client As ImapClient, destinationFolder As String)
        For i As Integer = 0 To client.Inbox.Count - 1
            Dim message = client.Inbox.GetMessage(i)
            'client.Inbox.AddFlags(i, MessageFlags.Seen, True)
            For Each subfolder In subfolders
                If subfolder.Name = destinationFolder Then
                    client.Inbox.MoveTo(i, subfolder)
                End If
            Next
        Next
    End Sub
    Public Sub ShowNamespaces(sThisEmailService, sThisEmailUser, sThisEmailPassword)
        Using client = New ImapClient(New ProtocolLogger("imap.log"))
            client.Connect(sThisEmailService, 993, SecureSocketOptions.SslOnConnect)
            client.Authenticate(sThisEmailUser, sThisEmailPassword)
            Debug.Print("Personal namespaces:")

            For Each ns In client.PersonalNamespaces
                Debug.Print($"* \{ns.Path}\ \{ns.DirectorySeparator}\")
            Next

            Debug.Print("")
            Debug.Print("Shared namespaces:")

            For Each ns In client.SharedNamespaces
                Debug.Print($"* \{ns.Path}\ \{ns.DirectorySeparator}\")
            Next

            Debug.Print("")
            Debug.Print("Other namespaces:")

            For Each ns In client.OtherNamespaces
                Debug.Print($"* \{ns.Path}\ \{ns.DirectorySeparator}\")
            Next

            Debug.Print("")
            Dim personal = client.GetFolder(client.PersonalNamespaces(0))
            Dim subfolders = personal.GetSubfolders()
            'EmailScreener.subfolders = personal.GetSubfolders()
            Debug.Print("The list of folders that are direct children of the first personmal namespace:")

            For Each folder In subfolders
                Debug.Print($"* {folder.Name}")
            Next

            client.Disconnect(True)
        End Using
    End Sub

#Region "Utilities"
    Public Sub DataTable2CSV(ByVal table As DataTable, ByVal filename As String,
ByVal sepChar As String, Optional appendcsv As Boolean = False)
        filename = "C:\Users\gary_\Documents\GarysEmails.csv"
        Dim bfilefound As Boolean = False
        If File.Exists(filename) Then
            bfilefound = True
        End If
        Using writer As New System.IO.StreamWriter("C:\Users\gary_\Documents\GarysEmails.csv", appendcsv)
            Try
                ' first write a line with the columns name
                Dim sep As String = ""
                Dim builder As New System.Text.StringBuilder
                If Not bfilefound Then
                    For Each col As DataColumn In table.Columns
                        If col.ColumnName = "TextBody" Then Exit For
                        builder.Append(sep).Append(col.ColumnName)
                        sep = sepChar
                    Next
                    writer.WriteLine(builder.ToString())
                End If
                ' then write all the rows
                Dim i = 0
                For Each row As DataRow In table.Rows
                    If i = 10 Then Exit For
                    If row.RowState <> DataRowState.Deleted Then
                        sep = ""
                        builder = New System.Text.StringBuilder
                        For Each col As DataColumn In table.Columns
                            If col.ColumnName = "TextBody" Then Exit For
                            builder.Append(sep).Append($"{Chr(34)}{row(col.ColumnName)}{Chr(34)}")
                            sep = sepChar
                        Next
                        Dim x = builder.ToString()
                        writer.WriteLine(builder.ToString())
                    End If
                Next
            Catch ex As Exception

            End Try
        End Using

    End Sub
    Public Function CSV2DataTable(ByVal dt As DataTable, strFileName As String) As Boolean

        dt.Reset()
        dt.Dispose()
        CSV2DataTable = False
        Dim line As String, dlinecnt As Double
        Dim aRow As DataRow

        Try
            LOGIT($"{strFileName} - file To format")
            Using mystream = New StreamReader(strFileName)
                Do
                    'read a line from the csv
                    line = mystream.ReadLine()
                    If line Is Nothing Then
                        Exit Do
                    ElseIf line.Replace(",", "") = "" Then
                        Exit Do
                    End If
                    dlinecnt += 1
                    If line Is Nothing Then Exit Do

                    'build a string array of scores using comma delimited
                    Dim sAry As String() = Split(line.Trim(","), ",")
                    'if this is the first line, it is a header so save each column header and mark the numeric ones 

                    If dlinecnt = 1 Then
                        If dt.Columns.Count = 0 Then
                            For i = 0 To sAry.Count - 1
                                Dim dc = New DataColumn(sAry(i))
                                'If sAry(i) = "Team" Or sAry(i).Contains("#") Or sAry(i).Contains("$") Then
                                '    dc.DataType = System.Type.GetType("System.Int16")
                                '    'ElseIf sAry(i) = "" Then
                                '    '    dc.DataType = System.Type.GetType("System.Boolean")
                                'Else
                                '    dc.DataType = System.Type.GetType("System.String")
                                'End If
                                ''LOGIT($"field({sAry(i)}) - {dc.DataType}")
                                dt.Columns.Add(dc)
                            Next
                        End If
                        Continue Do
                    End If

                    aRow = dt.NewRow
                    For i = 0 To sAry.Count - 1
                        If sAry(i).Trim <> "" Then
                            Try
                                aRow(i) = sAry(i)
                            Catch ex As Exception
                                Dim msg As New StringBuilder
                                msg.AppendLine(String.Format("Bad field in file-{0}", strFileName))
                                msg.AppendLine(String.Format("Player {0}", sAry(4)))
                                msg.AppendLine(String.Format("Date {0}", sAry(5)))
                                msg.AppendLine(String.Format("Row {0}", dlinecnt))
                                msg.AppendLine(String.Format("Column-{0}", dt.Columns(i).ColumnName))
                                msg.AppendLine(String.Format("Value {0}", sAry(i)))
                                aRow(i) = DBNull.Value
                                LOGIT(msg.ToString)
                                'MsgBox(msg.ToString)
                                'MsgBox(String.Format("bad field in file-{2}{4}Row {3}{4}Column-{0}{4}value {1}", dt.Columns(i).ColumnName, sAry(i), strFileName, dlinecnt, vbCrLf))
                                Debug.Flush()
                            End Try
                        Else
                            'If strFileName.Contains("Payments") Then
                            '    aRow(i) = " "
                            'End If

                        End If
                    Next
                    dt.Rows.Add(aRow)

                    'Try
                    '    dt.Rows.Add(aRow)
                    'Catch ex As Exception
                    '    If Debugger.IsAttached Then LOGIT("")

                    'End Try
                Loop

            End Using

            CSV2DataTable = True
            dt.PrimaryKey = New DataColumn() {dt.Columns("Name")}

            'If strFileName.Contains("Players") Or strFileName.Contains("Courses") Then
            '    dt.PrimaryKey = New DataColumn() {dt.Columns("Name")}
            'ElseIf strFileName.Contains("Schedule") Then
            '    dt.PrimaryKey = New DataColumn() {dt.Columns(0)}
            'ElseIf strFileName.Contains("Scores") Then
            '    dt.PrimaryKey = New DataColumn() {dt.Columns("Player"), dt.Columns("Date")}
            'ElseIf strFileName.Contains("Payments") Then
            'Dim xdata As New List(Of String)
            'xdata = New List(Of String)
            'For Each row As DataRow In dt.Rows
            '    If xdata.Contains(row(Constants.Player) & row(Constants.datecon) & row("Desc") & row(Constants.Detail) & row(Constants.Comment)) Then
            '        Debug.Print($"Duplicate {row(Constants.Player) & row(Constants.datecon) & row("Desc") & row(Constants.Detail) & row(Constants.Comment)Then}")
            '    End If
            '    xdata.Add(row(Constants.Player) & row(Constants.datecon) & row("Desc") & row(Constants.Detail) & row(Constants.Comment))
            'Next

            'dt.PrimaryKey = New DataColumn() {dt.Columns(Constants.Player), dt.Columns(Constants.datecon), dt.Columns(Constants.Desc), dt.Columns(Constants.Detail)}
            'End If

        Catch ex As Exception
            If ex.Message.Contains("being used by another process") Then
                'Dim strFile As String = "c:\windows\system32\msi.dll"
                Dim a As ArrayList = getFilesInUse(strFileName)

                'For Each p As Process In‼ a
                '    If p.ProcessName.Contains("wps") Or p.ProcessName.Contains("excel") Or p.ProcessName.Contains("notepad") Then
                '        LOGIT(p.ProcessName)
                '    End If
                'Next

                'sFileInUseMessage = ex.Message
                MsgBox(String.Format("file {0} In use, try later", strFileName))
            ElseIf ex.Message.Contains("These columns don't currently have unique values.") Then
                Dim dv = New DataView(dt)
                dv.Sort = "Name"
                Dim sprv = "", icnt = 0
                For Each row In dv
                    'If icnt = 0 Then sprv = row("Name")
                    If row("Name") <> sprv Then
                        sprv = row("Name")
                        Continue For
                    End If
                    MsgBox(String.Format("Duplicate Error in table {0} - {1}{2}Fix file{3} and try again", dt.TableName, sprv, vbCrLf, strFileName))
                    LOGIT(row("Name") & " - Dup")
                Next
            Else
                MsgBox(String.Format("Error {0} row {1}", strFileName, dlinecnt) & vbCrLf & ex.Message & vbCrLf & ex.StackTrace)
            End If
            '    MsgBox(String.Format("file {0} In use, try later", strFileName))
            'End If
            'LOGIT("")
        Finally

        End Try

    End Function
    Public Function TSV2DataTable(dt As DataTable, strFileName As String) As Boolean
        TSV2DataTable = False
        Dim line As String
        Dim dlinecnt As Double
        Try
            Using mystream As New StreamReader(strFileName, System.Text.Encoding.UTF8)
                Do
                    line = mystream.ReadLine()
                    If line Is Nothing Then Exit Do

                    ' Skip empty lines
                    If line.Trim() = "" Then Continue Do

                    dlinecnt += 1
                    Dim sAry As List(Of String) = ParseCSVLine(line)

                    ' First line = headers
                    If dlinecnt = 1 Then
                        If dt.Columns.Count = 0 Then
                            For Each col In sAry
                                dt.Columns.Add(New DataColumn(col.Trim()))
                            Next
                        End If
                        Continue Do
                    End If

                    ' Skip rows where all meaningful columns are empty
                    Dim hasData As Boolean = False
                    For i As Integer = 1 To Math.Min(5, sAry.Count - 1)
                        If sAry(i).Trim() <> "" Then
                            hasData = True
                            Exit For
                        End If
                    Next
                    If Not hasData Then Continue Do

                    Dim aRow = dt.NewRow()
                    For i As Integer = 0 To Math.Min(sAry.Count, dt.Columns.Count) - 1
                        aRow(i) = sAry(i).Trim()
                    Next
                    dt.Rows.Add(aRow)

                    If dlinecnt >= 13 AndAlso dlinecnt <= 16 Then
                        LOGIT($"Row {dlinecnt}: split count={sAry.Count}")
                        For i As Integer = 0 To Math.Min(sAry.Count - 1, 5)
                            LOGIT($"  col[{i}] = [{sAry(i)}]")
                        Next
                    End If

                Loop
            End Using
            TSV2DataTable = True
            LOGIT($"TSV2DataTable: loaded {dt.Rows.Count} rows")
        Catch ex As Exception
            LOGIT($"TSV2DataTable Error: {ex.Message} row {dlinecnt}")
        End Try
    End Function

    Private Function ParseCSVLine(line As String) As List(Of String)
        Dim fields As New List(Of String)
        Dim field As New System.Text.StringBuilder
        Dim inQuotes As Boolean = False

        For Each c As Char In line
            If c = """"c Then
                inQuotes = Not inQuotes
            ElseIf c = ","c AndAlso Not inQuotes Then
                fields.Add(field.ToString())
                field.Clear()
            Else
                field.Append(c)
            End If
        Next
        fields.Add(field.ToString())
        Return fields
    End Function
    Public Sub LOGIT(ByVal sMess As String)
        LOGIT(sMess, False)
    End Sub
    Public Sub LOGIT(ByVal sMess As String, Optional ByVal bLogToFile As Boolean = False)

        Try
            'If sMess.Contains("match for Tom Jennings") Then
            '    LOGIT("")
            'End If
            iLogitCounter += 1
            If Not bLogToFile Then
                If Debugger.IsAttached Then
                    Debug.WriteLine($"{iLogitCounter.ToString.PadLeft(5)} {DateTime.Now:yyyy-MM-dd HH:mm:ss} | {sMess}")
                    Exit Sub
                End If
            End If
            If Not Directory.Exists($"{sLogPath}") Then Directory.CreateDirectory($"{sLogPath}")

            Using swLog As New StreamWriter($"{sLogPath}{My.Computer.Name}_{Emails.sPgm}_{DateTime.Now:yyyyMMdd_HH}.log", True)
                If sMess IsNot Nothing Then
                    For Each stmpmess As String In sMess.Split(CChar(vbCrLf))
                        swLog.WriteLine($"{iLogitCounter.ToString.PadLeft(5)} {DateTime.Now:yyyy-MM-dd HH:mm:ss} | {sMess}")
                    Next
                End If
                swLog.Close()
            End Using
        Catch ex As Exception
            'MsgBox("Error " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
    Private Function getFilesInUse(ByVal strFile As String) As ArrayList

        myprocessarray.Clear()
        Dim processes As Process() = Process.GetProcesses()
        Dim i As Integer = 0

        For i = 0 To processes.GetUpperBound(0) - 1
            myProcess = processes(i)
            'Dim x = myProcess.StandardOutput.Read
            'LOGIT(myProcess.ProcessName & "-" & myProcess.MainWindowTitle)
            'If myProcess.ProcessName.ToLower.Contains("wps") Then
            '    Continue For
            'End If
            'Dim myprocess1 As New Process
            ''Program you want to launch execute etc.
            'myprocess1.StartInfo.FileName = "c:\windows\system32\openfiles.exe"
            'myprocess1.StartInfo.Arguments = "/query /s " + strFile + " /v"

            ''This is important. Since this is a windows service it will run
            ''even though no one is logged in.
            ''Therefore there is not desktop available so you better
            ''not show any windows dialogs
            'myprocess1.StartInfo.UseShellExecute = False
            'myprocess1.StartInfo.CreateNoWindow = True
            ''We want to redirect the output from the openfiles call to the program
            ''Since there won't be any window to display it in
            'myprocess1.StartInfo.RedirectStandardOutput = True

            'Dim tmpstr2 As String = String.Empty
            'Dim values(6) As Object 'This storeds the fields from openfiles
            'Dim values2(0) As Object 'This is the current date
            'values2(0) = DateTime.Now
            'Dim cnt As Integer = 0

            'Do
            '    tmpstr2 = myprocess1.StandardOutput.ReadLine
            '    ' Add some text to the file.
            '    If Not (tmpstr2 Is Nothing) Then
            '        cnt += 1
            '        'The output is fixed length
            '        If cnt > 5 Then
            '            values(0) = tmpstr2.Substring(0, 15).Trim 'Host name
            '            values(1) = tmpstr2.Substring(16, 8).Trim 'ID
            '            values(2) = tmpstr2.Substring(25, 20).Trim 'accessed by
            '            values(3) = tmpstr2.Substring(46, 10).Trim 'type
            '            values(4) = tmpstr2.Substring(57, 10).Trim 'locks
            '            values(5) = tmpstr2.Substring(68, 15).Trim 'open mode
            '            values(6) = tmpstr2.Substring(84) 'open file
            '        End If
            '    End If

            'Loop Until tmpstr2 Is Nothing

            If myProcess.Threads.Count > 0 Then

                Try
                    Dim modules As ProcessModuleCollection = myProcess.Modules
                    Dim j As Integer = 0

                    For j = 0 To modules.Count - 1

                        If (modules(j).FileName.ToLower().CompareTo(strFile.ToLower()) = 0) Then
                            myprocessarray.Add(myProcess)
                            Exit For
                        End If
                    Next

                Catch exception As Exception
                End Try
            End If
        Next

        Return myprocessarray
    End Function
    Sub getdbInfo()
        Dim sqlText As String = "SELECT * FROM Emails LIMIT 100"
        dtEmails = New DataTable()
        Using sqliteConn As New SqliteConnection(sqliteConnStr)
            sqliteConn.Open()
            Using cmd As New SqliteCommand(sqlText, sqliteConn)
                Using dr As SqliteDataReader = cmd.ExecuteReader()
                    dtEmails.Load(dr)
                End Using
            End Using
        End Using
        dtEmails.PrimaryKey = New DataColumn() {dtEmails.Columns("DateReceived"), dtEmails.Columns("From")}
    End Sub
    Function UpdateEmailData(drow As DataRow) As Long
        Dim sql = ""
        Dim rowsAffected As Long = 0
        Try
            sql = $"SELECT uid FROM emails WHERE [from] = '{drow("from").ToString.Replace("'", "''")}' AND datereceived = '{drow("datereceived")}'"
            Dim cmd As SQLiteCommand = New SQLiteCommand(sql, conn)
            cmd.CommandTimeout = 300

            Dim alreadyExists As Boolean = False
            Using dr As SQLiteDataReader = cmd.ExecuteReader()
                If dr.HasRows Then
                    dr.Read()
                    Debug.Print($"Already in database {dr(0)}")
                    alreadyExists = True
                End If
            End Using

            ' Only insert if not already there
            If Not alreadyExists Then
                Dim values As String = ""
                sql = "INSERT INTO Emails ("
                For Each col As DataColumn In dtEmails.Columns
                    If col.ColumnName = "Id" Then Continue For
                    sql &= col.ColumnName & ","
                    values &= $"'{drow(col.ColumnName).ToString.Replace("'", "''")}',"
                Next
                sql = sql.Substring(0, sql.LastIndexOf(",")) & ") VALUES ("
                values = values.Substring(0, values.LastIndexOf(",")) & ")"
                sql = sql.Replace("From,", "[From],").Replace("Read,", "[Read],")
                sql &= $"{vbCrLf}{values}"

                Dim insertCmd As New SQLiteCommand(sql, conn)
                insertCmd.CommandTimeout = 300
                rowsAffected = insertCmd.ExecuteNonQuery()
            End If

        Catch ex As Exception
            Debug.Print($"UpdateEmailData error: {sql} {ex.Message}")
        End Try
        Return rowsAffected
    End Function
    Sub CreateTableOnDBServer(tableName As String)
        Dim sqlText As String = "" '= "select * From Emails order by uid desc"
        sqlText = addsql(sqlText, $"SET ANS_NULLS ON")
        sqlText = addsql(sqlText, $"SET QUOTED_IDENTIFIER ON")
        sqlText = addsql(sqlText, $"DROP TABLE OF EXISTS [dbo].[{tableName}]")
        sqlText = addsql(sqlText, $"CREATE TABLE [dbo].[{tableName}]")

        sqlText &= $""
        Dim MaxColLen As New Dictionary(Of String, Integer)
        For Each dc As DataColumn In dtEmails.Columns
            MaxColLen.Add(dc.ColumnName, 0)
            If dc.DataType Is GetType(Boolean) Then
                MaxColLen(dc.ColumnName) = 5
                Continue For
            End If
            For Each dr As DataRow In dtEmails.Rows
                If dr.Field(Of String)(dc.ColumnName).Length > MaxColLen(dc.ColumnName) Then
                    MaxColLen(dc.ColumnName) = dr.Field(Of String)(dc.ColumnName).Length
                End If
            Next
        Next
        Dim i = 1
        For Each col In MaxColLen
            Dim scoltype = "varchar"
            Dim iVal = 200
            If col.Value > 1 Then iVal = col.Value
            If col.Key.ToLower.Contains("date") Then iVal = 50
            If col.Key.ToLower.Contains("date") Then iVal = 50
            'If col.Key.ToLower = "request restore" Or col.Key.ToLower = "setup env" Then
            '    scoltype = "bit "
            'Else
            scoltype = $"varchar({iVal})"
            'End If

            If i <> MaxColLen.Count Then
                sqlText = addsql(sqlText, $"[{col.Key}] {scoltype},")
            Else
                sqlText = addsql(sqlText, $"[{col.Key}] {scoltype})")
            End If
            i += 1
            'x = col.Key
            'xx = col.Value
        Next

    End Sub
    Function addsql(basesql As String, sql As String, Optional bskipVBCRLF As Boolean = False) As String
        addsql = basesql
        addsql &= $"{sql}"
        If Not bskipVBCRLF Then addsql &= vbCrLf
        Return addsql
    End Function
    Sub BuildTables()
        getdbInfo()
        dtEmails = New DataTable
        dtEmails.TableName = "Emails"
        dtEmails = BuildDTColumns(dtEmails, sEmailFields.Replace(vbCrLf, ""))

        ' Load from SQLite instead of SQL Server
        Using sqliteConn As New SqliteConnection("Data Source=C:\GarysEmails.db;")
            sqliteConn.Open()
            Using cmd As New SqliteCommand("SELECT * FROM Emails", sqliteConn)
                Using dr As SqliteDataReader = cmd.ExecuteReader()
                    dtEmails.Load(dr)
                End Using
            End Using
        End Using

        ' Remove duplicates
        Dim sprevuid As Double = 0
        Dim idups As Long = 0
        Dim rowsToDelete As New List(Of DataRow)
        For Each row As DataRow In dtEmails.Rows
            If CDbl(row("Uid")) = sprevuid Then
                idups += 1
                rowsToDelete.Add(row)
            Else
                sprevuid = CDbl(row("Uid"))
            End If
        Next
        For Each row As DataRow In rowsToDelete
            row.Delete()
        Next
        dtEmails.AcceptChanges()

        Debug.Print($"dups {idups}")
        dtEmails.PrimaryKey = New DataColumn() {dtEmails.Columns("Uid")}
        'sqlText = "select * From leagueparms"
        'da = New SqlDataAdapter(sqlText, sqlConn)
        'da.Fill(oHelper.dsLeague, "LeagueParms")

        If File.Exists(sDomainsCSV) Then
            CSV2DataTable(dtDomains, sDomainsCSV)
        Else
            dtDomains = New DataTable
            dtDomains.TableName = "Domains"
            dtDomains = BuildDTColumns(dtDomains, sDomainFields.Replace(vbCrLf, ""))
        End If

        If File.Exists(sEmailsCSV) Then
            CSV2DataTable(dtEmails, sEmailsCSV)
        Else
            dtEmails = New DataTable
            dtEmails.TableName = "Emails"
            dtEmails = BuildDTColumns(dtEmails, sEmailFields.Replace(vbCrLf, ""))
        End If

        For Each row In dtEmails.Rows
            'Dim 
        Next
    End Sub
    Function BuildDTColumns(InputDT As DataTable, sCols As String) As DataTable
        'option 1 - field name
        'option 2 - field format
        'option 3 - primary key field
        'option 4 - visible to grid
        'option 5 - Writeable to grid
        Dim dt = InputDT
        Dim dc = New DataColumn("")
        Dim dcPKey() As DataColumn = Nothing
        Dim dcNewKey As DataColumn
        Dim IPkeyRC As Integer = 0

        Dim colname As String = ""
        Dim fieldformat As String = ""
        Dim primaryKey As String = ""
        Dim visibleingrid As String = ""
        Dim Writeable As String = ""

        For Each scol In sCols.Split("|")

            Dim sColOptions() As String = scol.Split(",")
            If sColOptions.Count <> 5 Then
                MessageBox.Show($"{dt.TableName}-{sColOptions(0)}", "Not enough options, ending", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                End
            End If
            colname = sColOptions(0)
            fieldformat = sColOptions(1)
            primaryKey = sColOptions(2)
            visibleingrid = sColOptions(3)
            Writeable = sColOptions(4)

            dc = New DataColumn(colname)
            'option 2 - field format
            Select Case fieldformat
                Case "O"
                    dc.DataType = Type.GetType("System.DateTimeOffset")
                Case "D"
                    dc.DataType = Type.GetType("System.DateTime")
                Case "B"
                    dc.DataType = Type.GetType("System.Boolean")
                Case "I"
                    dc.DataType = Type.GetType("System.Int32")
            End Select
            'option 3 - primary key field
            If primaryKey = "Y" Then
                dcNewKey = New DataColumn(scol.Split(",")(0))
                dt.Columns.Add(dcNewKey)
                ReDim Preserve dcPKey(IPkeyRC)
                dcPKey(IPkeyRC) = dcNewKey
                IPkeyRC += 1
            Else
                dt.Columns.Add(dc)
            End If
        Next
        dt.PrimaryKey = dcPKey
        Return dt
    End Function

#End Region

#Region "Mail Utilities"

    'Opens client, gets access to inbox as read only
    Sub readmessages()
        Using client = New ImapClient(New ProtocolLogger("imap.log"))
            Dim j As Integer = 1
l1:
            client.Connect(sThisEmailService, 993, True)
            client.Authenticate(sThisEmailUser, sThisEmailPassword)
            Dim inbox = client.Inbox
            inbox.Open(FolderAccess.[ReadOnly])
            Dim uids = client.Inbox.Search(SearchQuery.NotSeen)
            'GetFlags(inbox)
            Console.WriteLine("Total messages:  {0}", inbox.Count)
            Console.WriteLine("Recent messages: {0}", inbox.Recent)
            Debug.Print("Total messages: {0}", inbox.Count)
            Debug.Print("Recent messages: {0}", inbox.Recent)
            Debug.Print("You have {0} unread message(s).", uids.Count)

            For i As Integer = 0 To uids.Count - 1
                Dim message = inbox.GetMessage(uids(i))
                Dim sdomain = message.From.ToString.Substring(message.From.ToString.IndexOf("@") + 1).Replace(">", "")
                UpdatedtTable(sdomain, message)

                'Console.WriteLine("Subject: {0}", message.Subject)

                'this will mark each unread email as read
                'client.Inbox.AddFlags(New UniqueId() {uids(i)}, MessageFlags.Seen, True)
                'Console.WriteLine("You have {0} unread message(s).", uids.Count - i)
                'Debug.Print("You have {0} unread message(s).", uids.Count - i)

                If j Mod 500 = 0 Then
                    client.Disconnect(True)
                    Console.WriteLine("Disconnected")
                    Thread.Sleep(10)
                    GoTo l1
                End If
                j += 1
            Next

            'Using sw As StreamWriter()
            'GetAllUnread(inbox)
            'End Using
            client.Disconnect(True)
        End Using
    End Sub
    Sub GetAllUnread(inbox As IMailFolder)
        For i As Integer = 0 To inbox.Count - 1
            Dim message = inbox.GetMessage(i)
            'Console.WriteLine("Subject: {0}", message.Subject)
            Debug.Print($"Date:{message.Date} From:{message.From}{vbCrLf}Subject: {message.Subject}")
            If i = 10 Then
                Dim x = ""
            End If
        Next
    End Sub
    Sub UpdatedtTable(sdomain As String, message As MimeKit.MimeMessage)
        UpdatedtTable(sdomain, message, False)
    End Sub
    Sub UpdatedtTable(sdomain As String, message As MimeKit.MimeMessage, Optional bTrusted As Boolean = False)
        Try

            'Dim x As String = message.From.ToString
            'If x.Contains("The Home Depot") Then
            '    Debug.Print("")
            'End If
            'Debug.Print($"Date:{message.Date} From:{message.From} Subject: {message.Subject}")
            Dim drow As DataRow
            drow = dtDomains.Rows.Find(sdomain)
            If drow Is Nothing Then
                drow = dtDomains.NewRow
                If bTrusted Then
                    drow("Trusted") = True
                Else
                    drow("Trusted") = False
                End If
                'Dim mbr = MsgBox($"Keep messages from this domain{vbCrLf}{vbCrLf}{sdomain}?", MsgBoxStyle.YesNo)
                'If mbr = MsgBoxResult.Yes Then
                '    drow("Trusted") = True
                'Else
                '    drow("Trusted") = False
                'End If
                drow("Name") = sdomain
                drow("TotalEmails") = 1
                drow("LatestDate") = message.Date
                Dim sender = message.From.ToString.Substring(0, message.From.ToString.IndexOf("@")).Replace("<", "")
                drow("Senders") &= sender
                dtDomains.Rows.Add(drow)
            Else
                drow("TotalEmails") += 1
                If message.Date > drow("LatestDate") Then
                    drow("LatestDate") = message.Date
                End If
                Dim sender As String = message.From.ToString.Substring(0, message.From.ToString.IndexOf("@")).Replace("<", "")
                If Not drow("Senders").ToString.Contains(sender) Then
                    drow("Senders") &= sender & vbCrLf
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub EmailSetup()

    End Sub
    '    Sub CountEmails()
    '        Try
    '            Dim j = 1
    '            'Using client = New ImapClient(New ProtocolLogger("imap.log"))
    '            Using client = New ImapClient()
    '                EmailSetup()
    'l1:
    '                tsStatusText.Text = $"Connecting To {sThisEmailService} For User {sThisEmailUser}"
    '                Application.DoEvents()
    '                client.Connect(sThisEmailService, 993, True)
    '                'setup yahoo Generate App Password called Emailscreener and copy that password and paste here
    '                tsStatusText.Text = $"Authenticating User {sThisEmailUser}"
    '                Application.DoEvents()
    '                client.Authenticate(sThisEmailUser, sThisEmailPassword)
    '                Dim inbox = client.Inbox
    '                tsStatusText.Text = $"Opening Inbox {inbox.FullName}"
    '                Application.DoEvents()
    '                inbox.Open(FolderAccess.[ReadOnly])
    '                tsStatusText.Text = $"Total Emails {inbox.Count}"
    '                tsStatusBar.Value = 0
    '                tsStatusBar.Maximum = inbox.Count - 1
    '                For i = 0 To inbox.Count - 1
    '                    LOGIT($"{i} Of {inbox.Count} ", True)
    '                    tsStatusText.Text = $"Getting email message {i} Of {inbox.Count}"
    '                    Application.DoEvents()
    '                    Dim message = inbox.GetMessage(i)
    '                    Dim domain = message.From.ToString.Substring(message.From.ToString.IndexOf("@") + 1).Replace(">", "")
    '                    Dim sender = message.From.ToString.Substring(0, message.From.ToString.IndexOf("@")).Replace("<", "")
    '                    UpdatedtTable(domain, message)
    '                    'UpdateEmailTbl.UpdateIt(dtEmails, message)
    '                    If j Mod 500 = 0 Then
    '                        client.Disconnect(True)
    '                        Console.WriteLine("Disconnected")
    '                        Thread.Sleep(10)
    '                        GoTo l1
    '                    End If
    '                    j += 1
    '                    'tsStatusText.Text = $"Saved And counted {i} email Of {inbox.Count}"
    '                    tsStatusBar.Value += 1
    '                    Application.DoEvents()
    '                Next
    '                client.Disconnect(True)
    '            End Using

    '        Catch ex As Exception

    '        End Try
    '        Dim x = ""
    '    End Sub
    Private Sub GetFlags(inbox As IMailFolder)
        Dim info = inbox.Fetch({4442}, MessageSummaryItems.Flags Or MessageSummaryItems.Flags)

        If info(0).Flags.Value.HasFlag(MessageFlags.Flagged) Then
        End If

        If info(0).Flags.Value.HasFlag(MessageFlags.Draft) Then
        End If

        If info(0).GMailLabels.Contains("Important") Then
        End If
    End Sub
    Sub SendEmails()

        Dim rs As New Resizer
        '        Dim semailfile = oHelper.sFilePath & "\" & Main.cbLeagues.Text.Substring(Main.cbLeagues.Text.IndexOf("(") + 1, 4) & "_" &
        'Main.cbLeagues.Text.Substring(0, Main.cbLeagues.Text.IndexOf("(") - 1) & "_Schedule.csv"
        Dim semailfile As String = ""
        Dim sfn = semailfile

        Dim ToAddresses As New List(Of String)

        Dim mbr = MsgBox(String.Format("are you ready to send emails to {0} players?", ToAddresses.Count), MsgBoxStyle.YesNo)
        If mbr <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        'Dim attachs() As String = {"d:\temp_Excell226.xlsx", "d:\temp_Excell224.xlsx", "d:\temp_Excell225.xlsx"}
        '2020_Hugh's_Schedule - Hugh's (2020)

        Dim attachs() As String = {semailfile}
        Dim subject As String = "Revised Schedule"
        Dim body As String = "see attached schedule"
        Dim bresult = False
        'If ToAddresses.Count > 0 Then
        '    bresult = GGmail.SendMail(ToAddresses, subject, body, attachs)
        '    If bresult Then
        '        MsgBox("mails sent successfully", MsgBoxStyle.Information)
        '    Else
        '        MsgBox(GGmail.ErrorText, MsgBoxStyle.Critical)
        '    End If
        'End If

    End Sub
    Sub GetGmail()
        'Using client = New ImapClient(New ProtocolLogger("imap.log"))
        '    client.Connect("imap.gmail.com", 993, True)
        '    client.Authenticate(sThisEmailUser, sThisEmailPassword)
        '    Dim inbox = client.Inbox
        '    inbox.Open(FolderAccess.[ReadOnly])
        '    Console.WriteLine("Total messages: {0}", inbox.Count)
        '    Console.WriteLine("Recent messages: {0}", inbox.Recent)

        '    For i As Integer = 0 To inbox.Count - 1
        '        Dim message = inbox.GetMessage(i)
        '        Console.WriteLine("Subject: {0}", message.Subject)
        '        Debug.Print($"From:{message.From}{vbCrLf}Subject: {message.Subject}")

        '        If i = 10 Then
        '            Dim x=""
        '        End If
        '    Next

        '    client.Disconnect(True)
        'End Using
    End Sub
#End Region
    Sub ConvertCSVToSQLite(csvPath As String)
        Try
            ' Create folder if it doesn't exist
            If Not System.IO.Directory.Exists("C:\EmailScreener") Then
                System.IO.Directory.CreateDirectory("C:\EmailScreener")
            End If


            ' Read CSV into DataTable
            Dim dt As New DataTable
            LOGIT($"Reading CSV from {csvPath}")
            Try

                TSV2DataTable(dt, csvPath)
            Catch ex As Exception

            End Try
            LOGIT($"Read {dt.Rows.Count} rows from CSV")

            ' Create SQLite DB and table
            Using sqliteConn As New SqliteConnection($"Data Source={sqlitePath};")
                sqliteConn.Open()

                Dim createSql As String = "
                CREATE TABLE IF NOT EXISTS Emails (
                    Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    DateReceived TEXT,
                    [From]      TEXT,
                    Domain      TEXT,
                    Sender      TEXT,
                    Subject     TEXT,
                    EmailClient TEXT,
                    Uid         INTEGER,
                    Spam        INTEGER DEFAULT 0,
                    [Read]      INTEGER DEFAULT 0,
                    PhishLevel  TEXT,
                    Folder      TEXT,
                    DateAdded   TEXT
                )"

                Using cmd As New SqliteCommand(createSql, sqliteConn)
                    cmd.ExecuteNonQuery()
                End Using

                ' Insert rows
                Using trans = sqliteConn.BeginTransaction()
                    Dim insertSql As String = "
                    INSERT INTO Emails 
                    (DateReceived, [From], Domain, Sender, Subject, EmailClient, Uid, Spam, [Read], PhishLevel, Folder, DateAdded)
                    VALUES
                    (@DateReceived, @From, @Domain, @Sender, @Subject, @EmailClient, @Uid, @Spam, @Read, @PhishLevel, @Folder, @DateAdded)"

                    Dim i As Integer = 0
                    For Each row As DataRow In dt.Rows
                        Using cmd As New SqliteCommand(insertSql, sqliteConn)
                            cmd.Transaction = trans
                            cmd.Parameters.AddWithValue("@DateReceived", If(row.Table.Columns.Contains("DateReceived"), row("DateReceived").ToString(), ""))
                            cmd.Parameters.AddWithValue("@From", If(row.Table.Columns.Contains("From"), row("From").ToString(), ""))
                            cmd.Parameters.AddWithValue("@Domain", If(row.Table.Columns.Contains("Domain"), row("Domain").ToString(), ""))
                            cmd.Parameters.AddWithValue("@Sender", If(row.Table.Columns.Contains("Sender"), row("Sender").ToString(), ""))
                            cmd.Parameters.AddWithValue("@Subject", If(row.Table.Columns.Contains("Subject"), row("Subject").ToString(), ""))
                            cmd.Parameters.AddWithValue("@EmailClient", If(row.Table.Columns.Contains("EmailClient"), row("EmailClient").ToString(), ""))
                            cmd.Parameters.AddWithValue("@Uid", If(row.Table.Columns.Contains("Uid") AndAlso IsNumeric(row("Uid")), CInt(row("Uid")), 0))
                            cmd.Parameters.AddWithValue("@Spam", If(row.Table.Columns.Contains("Spam") AndAlso row("Spam").ToString() = "True", 1, 0))
                            cmd.Parameters.AddWithValue("@Read", If(row.Table.Columns.Contains("Read") AndAlso row("Read").ToString() = "True", 1, 0))
                            cmd.Parameters.AddWithValue("@PhishLevel", If(row.Table.Columns.Contains("PhishLevel"), row("PhishLevel").ToString(), ""))
                            cmd.Parameters.AddWithValue("@Folder", If(row.Table.Columns.Contains("Folder"), row("Folder").ToString(), ""))
                            cmd.Parameters.AddWithValue("@DateAdded", If(row.Table.Columns.Contains("DateAdded"), row("DateAdded").ToString(), ""))
                            cmd.ExecuteNonQuery()
                        End Using
                        i += 1
                        If i Mod 1000 = 0 Then LOGIT($"Inserted {i} of {dt.Rows.Count}")
                    Next
                    trans.Commit()
                End Using

                LOGIT($"Migration complete — {dt.Rows.Count} rows inserted into {sqlitePath}")
            End Using

        Catch ex As Exception
            LOGIT($"ConvertCSVToSQLite Error: {ex.Message}")
        End Try
    End Sub

End Class
