Imports System.Threading
Imports MailKit
Imports MailKit.Net.Imap
Imports MailKit.Search
Imports Microsoft.Data.SqlClient
Imports Microsoft.Data.Sqlite
Public Class Emails

    Dim rs As New Resizer
    Dim uids As IEnumerable(Of UniqueId)
    Dim client As ImapClient
    Dim dtEmails As DataTable
    Public sPgm As String = System.Diagnostics.Process.GetCurrentProcess().ProcessName
    Dim eUtil As New eUtilities
    Private WithEvents keepAliveTimer As New Windows.Forms.Timer With {.Interval = 240000}
    Private keepAliveBusy As Boolean = False

    Private Sub Emails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadscreen()
        ToolStrip.Visible = False
        Me.Text = $"Form {Me.CompanyName}-{My.Computer.Name}, Resolution {Screen.PrimaryScreen.Bounds.Width} x {Screen.PrimaryScreen.Bounds.Height}, Menu {Me.Width} x {Me.Height}, Grid {dgvEmails.Width} x {dgvEmails.Height}"

        For Each sclient As String In eUtil.lEmailClients
            cmbEmailClients.Items.Add(sclient.Split(",")(0))
        Next
        cmbEmailClients.SelectedIndex = 0

        For Each eclient As String In eUtil.lEmailClients
            If eclient.Split(",")(0) = cmbEmailClients.SelectedItem Then
                eUtil.sThisEmailService = eclient.Split(",")(1)
                eUtil.sThisEmailUser = eclient.Split(",")(2)
                eUtil.sThisEmailPassword = eclient.Split(",")(3)
                Exit For
            End If
        Next
        If System.IO.File.Exists(eUtil.sqlitePath) Then
            eUtil.openconn()
            eUtil.getdbInfo()
            dtEmails = eUtil.dtEmails
            dgvEmails.DataSource = dtEmails
        Else
            MessageBox.Show("Database not found. Please click Convert button to migrate from SQL Server.", "No Database", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        ' Load data from SQLite
        dtEmails = eUtil.dtEmails
        dgvEmails.DataSource = dtEmails
        dgvEmails.AllowUserToAddRows = False
        dgvEmails.RowHeadersVisible = False
        dgvEmails.AllowUserToResizeColumns = True

        For Each col As DataGridViewColumn In dgvEmails.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Dim colname As String = ""
            Dim visibleingrid As String = ""
            Dim Writeable As String = ""
            For Each fld As String In eUtil.sEmailFields.Split("|")
                colname = fld.Split(",")(0).Replace(vbCrLf, "")
                If colname = col.Name Then
                    visibleingrid = fld.Split(",")(3)
                    Writeable = fld.Split(",")(4)
                    If visibleingrid = "N" Then col.Visible = False
                    If Writeable <> "Y" Then col.ReadOnly = True
                    Exit For
                End If
            Next
        Next

        Me.Show()
        tbMailClient.Text = cmbEmailClients.SelectedItem
    End Sub
    Sub loadscreen()
        Dim allScreens = Screen.AllScreens
        Dim Current_Screen As Screen = Screen.FromControl(Me)
        If Current_Screen.Primary Then
            Dim HCenter = Current_Screen.Bounds.Left +
        (((Current_Screen.Bounds.Right - Current_Screen.Bounds.Left) / 2) - ((Me.Width) / 2))
            Dim VCenter = (Current_Screen.Bounds.Bottom / 2) - ((Me.Height) / 2)
            Me.StartPosition = FormStartPosition.Manual
            Me.Location = New Point(HCenter, VCenter)
        Else
            Me.StartPosition = FormStartPosition.CenterScreen
        End If
    End Sub
    Sub GetUnread()
        keepAliveBusy = True
        ToolStrip.Visible = True
        Dim imessagecount = 0
        btnSave.Visible = False
        cbSelectAll.Checked = False

        'Using client As New ImapClient'
        Me.Show()
        Try

            Dim j As Integer = 1
l1:
            If Not IsClientConnected() Then
                client = eUtil.Connect()
                Application.DoEvents()
                eUtil.Authenticate()
                Application.DoEvents()
            End If

            Try
                tsStatusText.Text = eUtil.OpenInbox
                Application.DoEvents()
                tsStatusText.Text = $"Opening {client.Inbox.FullName}"
                Application.DoEvents()
                client.Inbox.Open(FolderAccess.ReadOnly)
                '(Email, folder)
                'inbox.Search(SearchQuery.All).Reverse().Take(50);,
                uids = client.Inbox.Search(SearchQuery.NotSeen).Reverse
                Dim olduids = client.Inbox.Search(SearchQuery.DeliveredAfter(Now.AddYears(-1)))
                Dim oldmessage = client.Inbox.GetMessage(olduids(1))
                oldmessage = client.Inbox.GetMessage(olduids(olduids.Count - 1))
                imessagecount = uids.Count
                'Console.WriteLine("You have {0} unread message(s).", uids.Count)
                'Debug.Print($"You have {uids.Count} unread message(s).")
                tsProgressBar.Value = 0
                tsProgressBar.Maximum = uids.Count
                tbUnread.Text = uids.Count
                tbSpam.Text = 0
                If imessagecount = 0 Then
                    tsStatusText.Text = $"{Now()}-No messages to process"
                    MessageBox.Show($"No messages to process", $"{sPgm}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    keepAliveBusy = False
                    Exit Sub
                End If
                For i As Integer = 0 To uids.Count - 1
                    If i = 0 Then btnSave.Visible = True
                    'resize each col in grid
                    'If i = 0 Then
                    '    For Each col As DataGridViewColumn In dgvEmails.Columns
                    '        If col.Name = "Date" Then col.Width = 100
                    '        If col.Name = "From" Then col.Visible = False
                    '        If col.Name = "Subject" Then col.Width = 300
                    '    Next
                    'End If
                    Dim uids1 = uids(i)
                    Dim message = client.Inbox.GetMessage(uids1)
                    Dim sdomain = message.From.ToString.Substring(message.From.ToString.IndexOf("@") + 1).Replace(">", "")
                    'Dim mbr = MsgBox($"Add {sdomain} to trusted domains amd mark Message {message.Subject} from {message.From}", MsgBoxStyle.YesNo, "EmailScreener")
                    'If mbr = MsgBoxResult.Yes Then
                    '    UpdatedtTable(sdomain, message, True)
                    'Else
                    '    UpdatedtTable(sdomain, message)
                    'End If
                    If eUtil.UpdateIt(uids1.Id, client.Inbox.FullName, message, If(cbSpam.Checked, True, False)) Then
                        tbSpam.Text += 1
                    End If
                    'Console.WriteLine("You have {0} unread message(s).", uids.Count - i)
                    'Debug.Print("You have {0} unread message(s).", uids.Count - i)
                    If dtEmails.Rows.Count = 0 Then Continue For
                    If j Mod 500 = 0 Then
                        SafeDisconnectClient()
                        Console.WriteLine("Disconnected")
                        Thread.Sleep(10)
                        GoTo l1
                    End If
                    tsProgressBar.Value += 1
                    tsStatusText.Text = $"{i + 1} of {uids.Count} {message.Date} {message.From} {message.Subject}"
                    Application.DoEvents()
                    'If Debugger.IsAttached Then
                    '    If i = 10 Then Exit For
                    'End If
                Next
                'client.Disconnect(True)

            Catch e As Exception
                SafeDisconnectClient()
                Console.WriteLine("Error")
                GoTo l1
            End Try

            j += 1
            If imessagecount > 500 Then

            End If
        Catch ex As Exception
            SafeDisconnectClient()
        Finally
            keepAliveBusy = False
        End Try
        'End Using
        eUtil.LOGIT($"Total Messages Marked: {imessagecount}")
        Dim sImportemails As String() = {"glemker@fuse.net".ToUpper, "glemker@fuse.net".ToUpper}
        For Each row As DataGridViewRow In dgvEmails.Rows
            Dim sfrom = row.Cells("From").Value.ToString.ToUpper
            If sfrom.Contains("LEMKER") Then
                Dim xx = ""
            End If
            If row.Cells("Spam").Value = True Then
                row.DefaultCellStyle.BackColor = Color.LightPink
            ElseIf sImportemails.Contains(row.Cells("From").Value.ToString.ToUpper) Then
                row.DefaultCellStyle.BackColor = Color.LightGreen
            Else
                row.DefaultCellStyle.BackColor = DefaultBackColor
            End If
        Next
        'dgvEmails.Columns("Spam").Visible = False
    End Sub
    Sub getJBulk()
        Dim sthisfolder As IMailFolder = Nothing
        sthisfolder = client.GetFolder(SpecialFolder.Junk)
        sthisfolder.Open(FolderAccess.ReadOnly)
        If cbFolders.SelectedItem = "Junk" Then

        End If

        tsProgressBar.Value = 0
        tsProgressBar.Maximum = sthisfolder.Count
        tbUnread.Text = sthisfolder.Count
        tbSpam.Text = 0
        For i = 0 To sthisfolder.Count - 1
            Dim thisfoldersmessage = sthisfolder.GetMessage(i)
            'updateTable(sthisfolder.GetMessage(i))
            tsProgressBar.Value += 1
            tsStatusText.Text = $"Folder {sthisfolder.Name}:{i + 1} of {sthisfolder.Count} {thisfoldersmessage.Date} {thisfoldersmessage.From} {thisfoldersmessage.Subject}"
            Application.DoEvents()
        Next

    End Sub
    Sub updateTable(message As MimeKit.MimeMessage)
        Try

            Dim sdomain = message.From.ToString.Substring(message.From.ToString.IndexOf("@") + 1).Replace(">", "")
            'Dim mbr = MsgBox($"Add {sdomain} to trusted domains amd mark Message {message.Subject} from {message.From}", MsgBoxStyle.YesNo, "EmailScreener")
            'If mbr = MsgBoxResult.Yes Then
            '    UpdatedtTable(sdomain, message, True)
            'Else
            '    UpdatedtTable(sdomain, message)
            'End If
            'If UpdateEmailTbl.UpdateIt(dtEmails, uid, message, If(cbSpam.Checked, True, False)) Then
            '    tbSpam.Text += 1
            'End If
            eUtil.LOGIT($"{message.Date}|{message.From}|{message.Subject}")
        Catch ex As Exception

        End Try

    End Sub
    Sub UpdatetsStatusText(sMessage)
        tsStatusText.Text = sMessage
        Application.DoEvents()
    End Sub

    Private Sub Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        'rs.ResizeAllControls(Me)
        'Me.Text = String.Format("Form {7}-{0}, Resolution {1} x {2}, Menu {3} x {4}, Grid {5} x {6}", My.Computer.Name, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Width, Me.Width, Me.Height, dgvEmails.Width, dgvEmails.Height, Me)
        Me.Text = $"Form {Me.Name}-{My.Computer.Name}, Resolution {Screen.PrimaryScreen.Bounds.Width} x {Screen.PrimaryScreen.Bounds.Height}, Menu {Me.Width} x {Me.Height}, Grid {dgvEmails.Width} x {dgvEmails.Height}"
    End Sub

    'Private Sub dgvEmails_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvEmails.CellDoubleClick
    '    If e.RowIndex < 0 Then
    '        Exit Sub
    '    End If
    '    Dim row As DataGridViewRow = dgvEmails.Rows(e.RowIndex)
    '    row.Cells("Read").Value = "Y"
    '    row.Selected = True
    'End Sub

    Private Sub btnMarkRead_Click(sender As Object, e As EventArgs) Handles btnMarkRead.Click
        If IsClientConnected() Then
            client.Inbox.Open(FolderAccess.ReadWrite)
        Else
            MessageBox.Show($"Email Client isnt connected, Quitting", "Warning", MessageBoxButtons.OK)
            Exit Sub
        End If
        If dgvEmails.SelectedRows.Count = 0 Then
            MessageBox.Show($"No rows selected, Quitting", "Warning", MessageBoxButtons.OK)
            Exit Sub
        End If
        tsProgressBar.Maximum = dgvEmails.SelectedRows.Count
        tsProgressBar.Value = 1
        UpdatetsStatusText(eUtil.OpenInbox(True))
        For i = 0 To dgvEmails.Rows.Count - 1
            If dgvEmails.Rows(i).Selected Then
                For Each uid In uids
                    If uid.Id = dtEmails.Rows(i)("uid") Then
                        client.Inbox.AddFlags(uid, MessageFlags.Seen, True)
                        dtEmails.Rows(i)("Read") = True
                        ' eUtil.UpdateIt(uid.Id,)
                        dgvEmails.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
                        tsStatusText.Text = $"{uid.Id} Marked as Read"
                        If tsProgressBar.Value < tsProgressBar.Maximum Then
                            tsProgressBar.Value += 1
                            Application.DoEvents()
                        End If
                        Exit For
                    End If
                Next
            End If
        Next
        tsStatusText.Text = $"{dgvEmails.Rows.Count} Marked as Read"
        Application.DoEvents()
        If eUtil.conn.State = ConnectionState.Open Then
            eUtil.conn.Close()
        End If

    End Sub

    Private Sub dgvEmails_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvEmails.CellClick
        If e.RowIndex < 0 Then Exit Sub
        'Dim dgc = TryCast(dgvEmails.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewCheckBoxCell)
        'If dgc IsNot Nothing Then Exit Sub
        Dim col = sender
        If col.currentcell.OwningColumn.Name = "Spam" Then
            Dim mbr = MsgBox($"Changing email to Spam {dgvEmails.Rows(e.RowIndex).Cells("Domain").Value}-{dgvEmails.Rows(e.RowIndex).Cells("Sender").Value} ", vbOKCancel)
            If mbr.Ok Then
                'col.currentcell.value = DBNull.Value
                col.currentcell.value = True
            End If
            'dgvEmails.Refresh()
            Dim x = ""
        End If
    End Sub
    Private Sub dgvEmails_SortCompare(sender As Object, e As DataGridViewSortCompareEventArgs) Handles dgvEmails.SortCompare
        eUtil.LOGIT("Entering " & Reflection.MethodBase.GetCurrentMethod.Name)
        'If e.Column.Index <> 0 Then
        '    Return
        'End If
        Try
            Dim sc1 = e.CellValue1
            Dim sc2 = e.CellValue2
            If IsNumeric(sc1) And IsNumeric(sc2) Then
                e.SortResult = If(CInt(sc1) < CInt(sc2), -1, 1)
            Else
                e.SortResult = If(CStr(sc1) < CStr(sc2), -1, 1)
            End If

            e.Handled = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnRefreshGrid_Click(sender As Object, e As EventArgs) Handles btnRefreshGrid.Click
        btnSave.Visible = True
        If dtEmails IsNot Nothing Then
            dtEmails.Rows.Clear()
        End If
        dgvEmails.Refresh()
        ToolStrip.Visible = True
        GetUnread()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        eUtil.DataTable2CSV(dtEmails, eUtil.sEmailsCSV, ",", True)
        btnSave.Visible = False
    End Sub

    Private Sub Emails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If dgvEmails.Rows.Count > 0 And btnSave.Visible = True Then
            Dim mbr = MsgBox($"You sure you want to exit without saving emails?", MsgBoxStyle.YesNo, "EmailScreener")
            If mbr = MsgBoxResult.Yes Then
                Exit Sub
            Else
                e.Cancel = True
                Exit Sub
            End If

        End If
        If eUtil.conn IsNot Nothing AndAlso eUtil.conn.State = ConnectionState.Open Then
            eUtil.conn.Close()
        End If
        keepAliveTimer.Stop()
        SafeDisconnectClient()
    End Sub

    Private Sub cbSelectAll_CheckedChanged(sender As Object, e As EventArgs) Handles cbSelectAll.CheckedChanged
        If cbSelectAll.Checked Then
            'Dim mbr = MessageBox.Show($"You sure you want to mark these {dgvEmails.Rows.Count} as read?", "Warning", MessageBoxButtons.YesNo)
            'If mbr = Windows.Forms.DialogResult.Yes Then
            'If client.IsConnected Then
            '    client.Inbox.Open(FolderAccess.ReadWrite)
            For Each row As DataGridViewRow In dgvEmails.Rows
                row.Selected = True
            Next
            'End If
            'End If
        Else
            For Each row As DataGridViewRow In dgvEmails.Rows
                row.Selected = False
            Next
        End If
    End Sub

    Private Sub btnCountEmails_Click(sender As Object, e As EventArgs) Handles btnCountEmails.Click
        'eUtil.CountEmails()
    End Sub

    Private Sub cbFolders_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbFolders.SelectedIndexChanged

    End Sub

    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        'If refreshing a new connection, then disconnect an existing connection if its already connected

        ToolStrip.Visible = True
        If client IsNot Nothing Then
            If IsClientConnected() Then
                SafeDisconnectClient()
                client = Nothing
            End If
        End If
        client = eUtil.Connect
        If IsClientConnected() Then
            UpdatetsStatusText($"Connected to {eUtil.sThisEmailService} for User {eUtil.sThisEmailUser}")
        Else
            UpdatetsStatusText($"Connected to {eUtil.sThisEmailService} for User {eUtil.sThisEmailUser}")
            tsStatusText.BackColor = Color.Red
            Exit Sub
        End If
        ToolStrip.Visible = True

        UpdatetsStatusText(eUtil.Authenticate)
        UpdatetsStatusText(eUtil.OpenInbox)
        eUtil.SetupFolders()
        If IsClientConnected() Then
            tbMailClient.BackColor = Color.LightGreen
            btnConnect.Visible = False
            StartKeepAlive()
            btnRefreshGrid_Click(sender, e)
        Else
            tbMailClient.BackColor = Color.Red
        End If

    End Sub

    Private Sub StartKeepAlive()
        keepAliveTimer.Stop()
        keepAliveTimer.Start()
        UpdatetsStatusText($"Keep alive started for {eUtil.sThisEmailUser}")
    End Sub

    Private Sub keepAliveTimer_Tick(sender As Object, e As EventArgs) Handles keepAliveTimer.Tick
        If keepAliveBusy Then Exit Sub
        keepAliveBusy = True

        Try
            If Not IsClientConnected() Then
                client = eUtil.Connect()
                If Not IsClientConnected() Then
                    UpdatetsStatusText($"Keep alive reconnect failed at {Now:t}")
                    Exit Sub
                End If
            End If

            If Not client.IsAuthenticated Then
                UpdatetsStatusText(eUtil.Authenticate())
            End If

            client.NoOp()
            UpdatetsStatusText($"Keep alive sent at {Now:t}")
        Catch ex As Exception
            UpdatetsStatusText($"Keep alive failed at {Now:t}: {ex.Message}")
        Finally
            keepAliveBusy = False
        End Try
    End Sub

    Private Function IsClientConnected() As Boolean
        Try
            Return client IsNot Nothing AndAlso client.IsConnected
        Catch ex As ObjectDisposedException
            client = Nothing
            Return False
        End Try
    End Function

    Private Sub SafeDisconnectClient()
        If client Is Nothing Then Exit Sub

        Try
            If IsClientConnected() Then
                client.Disconnect(True)
            End If
        Catch ex As ObjectDisposedException
            ' Already closed by MailKit or a reconnect attempt.
        Catch ex As Exception
            eUtil.LOGIT($"Disconnect failed: {ex.Message}")
        Finally
            client = Nothing
        End Try
    End Sub

    Private Sub cmbEmailClients_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEmailClients.SelectedIndexChanged
        Dim x = ""
        If tbMailClient.Text <> "" Then
            If cmbEmailClients.SelectedItem <> tbMailClient.Text Then
                If IsClientConnected() Then
                    SafeDisconnectClient()
                    btnConnect.Visible = True
                End If
            End If
        End If

    End Sub

    Private Sub dgvEmails_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvEmails.CellEnter
        Dim dgc As DataGridViewCell = sender.currentrow.cells(e.ColumnIndex)
        'checkSkinsCtps(sender,e)

        If dgc.ReadOnly Then
            SendKeys.Send("{tab}")
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        eUtil.ConvertCSVToSQLite("\\GarysNas\Backup\Emails.csv")
        MessageBox.Show("Migration complete!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class
