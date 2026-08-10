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
    Private WithEvents keepAliveTimer As New System.Windows.Forms.Timer With {.Interval = 240000}
    Private keepAliveBusy As Boolean = False
    Private forwardingRepository As ForwardingRepository
    Private forwardingService As New ForwardingService()
    Private forwardingRules As New List(Of ForwardingRule)
    Private forwardingUiLoading As Boolean
    Private selectedForwardingRuleId As Integer
    Private forwardingTabs As TabControl
    Private forwardingGrid As DataGridView
    Private WithEvents cbAutoForward As CheckBox
    Private WithEvents cmbForwardMatchType As ComboBox
    Private WithEvents txtForwardMatchValue As TextBox
    Private WithEvents txtForwardDestination As TextBox
    Private WithEvents cbForwardRuleEnabled As CheckBox
    Private WithEvents btnSaveForwardRule As Button
    Private WithEvents btnDeleteForwardRule As Button
    Private WithEvents btnUseSelectedSender As Button
    Private WithEvents forwardSearchSince As DateTimePicker
    Private WithEvents btnScanForwarding As Button
    Private WithEvents btnForwardPreviewed As Button
    Private forwardingPreviewGrid As DataGridView
    Private forwardingPreviewMessages As New Dictionary(Of Integer, MimeKit.MimeMessage)
    Private forwardingPreviewRule As ForwardingRule

    Private Sub Emails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadscreen()
        InitializeForwardingTab()
        ToolStrip.Visible = False
        Me.Text = $"Form {Me.CompanyName}-{My.Computer.Name}, Resolution {Screen.PrimaryScreen.Bounds.Width} x {Screen.PrimaryScreen.Bounds.Height}, Menu {Me.Width} x {Me.Height}, Grid {dgvEmails.Width} x {dgvEmails.Height}"

        For Each sclient As String In eUtil.lEmailClients
            cmbEmailClients.Items.Add(sclient.Split(",")(0))
        Next
        cmbEmailClients.SelectedIndex = 0
        ApplySelectedEmailClientConfiguration()
        If System.IO.File.Exists(eUtil.sqlitePath) Then
            eUtil.openconn()
            eUtil.getdbInfo()
            dtEmails = eUtil.dtEmails
            dgvEmails.DataSource = dtEmails
            InitializeForwardingData()
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
                    ProcessAutoForward(message, uids1.Id, client.Inbox.FullName)
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

    Private Sub InitializeForwardingTab()
        If forwardingTabs IsNot Nothing Then Exit Sub

        Dim emailTab As New TabPage("Email screening") With {.AutoScroll = True}
        Dim forwardTab As New TabPage("Auto forwarding")
        Dim controlsToMove As New List(Of Control)
        For Each control As Control In Controls
            If control IsNot ToolStrip Then controlsToMove.Add(control)
        Next
        For Each control In controlsToMove
            emailTab.Controls.Add(control)
        Next

        forwardingTabs = New TabControl With {.Dock = DockStyle.Fill}
        forwardingTabs.TabPages.Add(emailTab)
        forwardingTabs.TabPages.Add(forwardTab)
        Controls.Add(forwardingTabs)
        forwardingTabs.BringToFront()
        ToolStrip.BringToFront()

        Dim instructions As New Label With {
            .AutoSize = False,
            .Location = New Point(20, 15),
            .Size = New Size(1200, 38),
            .Text = "Automatically forward matching unread mail, or scan matching inbox mail since a chosen date. Duplicate sends are prevented."
        }
        cbAutoForward = New CheckBox With {
            .AutoSize = True,
            .Location = New Point(20, 58),
            .Text = "Enable automatic forwarding while screening unread mail"
        }
        Dim searchSinceLabel As New Label With {.AutoSize = True, .Location = New Point(540, 60), .Text = "Search since"}
        forwardSearchSince = New DateTimePicker With {
            .Format = DateTimePickerFormat.Short,
            .Location = New Point(625, 56),
            .Size = New Size(125, 23),
            .Value = Today.AddDays(-7)
        }
        btnScanForwarding = New Button With {.Location = New Point(765, 54), .Size = New Size(145, 28), .Text = "Preview matches"}
        btnForwardPreviewed = New Button With {.Location = New Point(925, 54), .Size = New Size(145, 28), .Text = "Forward previewed", .Enabled = False}
        cmbForwardMatchType = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(20, 112),
            .Size = New Size(130, 23)
        }
        cmbForwardMatchType.Items.AddRange(New Object() {"Sender", "Sender name", "Domain"})
        cmbForwardMatchType.SelectedIndex = 0
        txtForwardMatchValue = New TextBox With {.Location = New Point(165, 112), .Size = New Size(290, 23)}
        txtForwardDestination = New TextBox With {.Location = New Point(470, 112), .Size = New Size(290, 23)}
        cbForwardRuleEnabled = New CheckBox With {.AutoSize = True, .Location = New Point(775, 114), .Text = "Rule enabled", .Checked = True}
        btnSaveForwardRule = New Button With {.Location = New Point(885, 109), .Size = New Size(105, 28), .Text = "Add rule"}
        btnDeleteForwardRule = New Button With {.Location = New Point(1000, 109), .Size = New Size(105, 28), .Text = "Delete rule", .Enabled = False}
        btnUseSelectedSender = New Button With {.Location = New Point(1115, 109), .Size = New Size(150, 28), .Text = "Use selected sender"}

        Dim matchLabel As New Label With {.AutoSize = True, .Location = New Point(20, 91), .Text = "Match type"}
        Dim valueLabel As New Label With {.AutoSize = True, .Location = New Point(165, 91), .Text = "Sender email, name, or domain"}
        Dim destinationLabel As New Label With {.AutoSize = True, .Location = New Point(470, 91), .Text = "Destination Gmail address"}

        forwardingGrid = New DataGridView With {
            .Location = New Point(20, 155),
            .Size = New Size(1245, 165),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .ReadOnly = True,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False
        }
        AddHandler forwardingGrid.CellClick, AddressOf forwardingGrid_CellClick

        Dim previewLabel As New Label With {.AutoSize = True, .Location = New Point(20, 338), .Text = "Preview — no messages are sent until you click Forward previewed"}
        forwardingPreviewGrid = New DataGridView With {
            .Location = New Point(20, 360),
            .Size = New Size(1245, 295),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .ReadOnly = True,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = True
        }

        forwardTab.Controls.AddRange(New Control() {
            instructions, cbAutoForward, searchSinceLabel, forwardSearchSince, btnScanForwarding, btnForwardPreviewed,
            matchLabel, cmbForwardMatchType, valueLabel, txtForwardMatchValue,
            destinationLabel, txtForwardDestination, cbForwardRuleEnabled, btnSaveForwardRule,
            btnDeleteForwardRule, btnUseSelectedSender, forwardingGrid, previewLabel, forwardingPreviewGrid
        })
    End Sub

    Private Sub InitializeForwardingData()
        Try
            forwardingRepository = New ForwardingRepository(eUtil.sqlitePath)
            forwardingUiLoading = True
            cbAutoForward.Checked = forwardingRepository.GetAutoForwardEnabled()
            Dim savedDate = forwardingRepository.GetSetting("ForwardSearchSinceDate")
            Dim parsedDate As DateTime
            If DateTime.TryParse(savedDate, parsedDate) Then forwardSearchSince.Value = parsedDate
            RefreshForwardingRules()
        Catch ex As Exception
            cbAutoForward.Enabled = False
            UpdatetsStatusText($"Forwarding setup failed: {ex.Message}")
        Finally
            forwardingUiLoading = False
        End Try
    End Sub

    Private Sub RefreshForwardingRules()
        If forwardingRepository Is Nothing Then Exit Sub
        forwardingRules = forwardingRepository.LoadRules()
        forwardingGrid.DataSource = Nothing
        forwardingGrid.DataSource = forwardingRules
        If forwardingGrid.Columns.Contains("Id") Then forwardingGrid.Columns("Id").Visible = False
    End Sub

    Private Function ProcessAutoForward(message As MimeKit.MimeMessage,
                                        uid As Integer,
                                        folder As String,
                                        Optional force As Boolean = False,
                                        Optional rulesOverride As IEnumerable(Of ForwardingRule) = Nothing) As Integer
        If forwardingRepository Is Nothing OrElse (Not force AndAlso Not cbAutoForward.Checked) Then Return 0
        Try
            Dim smtpHost = AppConfiguration.GetSmtpHost(eUtil.sThisEmailService)
            Dim rulesToApply = If(rulesOverride, forwardingRules)
            Dim count = forwardingService.ForwardMatching(
                message,
                uid,
                cmbEmailClients.SelectedItem.ToString(),
                folder,
                smtpHost,
                eUtil.sThisEmailUser,
                eUtil.sThisEmailPassword,
                rulesToApply,
                forwardingRepository)
            If count > 0 Then eUtil.LOGIT($"Auto-forwarded UID {uid} to {count} destination(s)")
            Return count
        Catch ex As Exception
            eUtil.LOGIT($"Auto-forward failed for UID {uid}: {ex.Message}", True)
            UpdatetsStatusText($"Auto-forward failed for UID {uid}: {ex.Message}")
            Return 0
        End Try
    End Function

    Private Sub cbAutoForward_CheckedChanged(sender As Object, e As EventArgs) Handles cbAutoForward.CheckedChanged
        If forwardingUiLoading OrElse forwardingRepository Is Nothing Then Exit Sub
        forwardingRepository.SetAutoForwardEnabled(cbAutoForward.Checked)
    End Sub

    Private Sub forwardSearchSince_ValueChanged(sender As Object, e As EventArgs) Handles forwardSearchSince.ValueChanged
        If forwardingUiLoading OrElse forwardingRepository Is Nothing Then Exit Sub
        forwardingRepository.SetSetting("ForwardSearchSinceDate", forwardSearchSince.Value.Date.ToString("yyyy-MM-dd"))
    End Sub

    Private Sub btnScanForwarding_Click(sender As Object, e As EventArgs) Handles btnScanForwarding.Click
        If forwardingRepository Is Nothing Then Exit Sub
        If Not ValidateMatchCriteria() Then Exit Sub

        forwardingPreviewRule = New ForwardingRule With {
            .MatchType = cmbForwardMatchType.Text,
            .MatchValue = txtForwardMatchValue.Text.Trim(),
            .Enabled = True
        }

        ApplySelectedEmailClientConfiguration()
        If Not IsClientConnected() Then client = eUtil.Connect()
        If Not IsClientConnected() Then Exit Sub
        If Not client.IsAuthenticated Then eUtil.Authenticate()
        If Not client.IsAuthenticated Then
            MessageBox.Show("Email authentication failed.", "Scan inbox", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        tbMailClient.Text = cmbEmailClients.SelectedItem.ToString()
        tbMailClient.BackColor = Color.LightGreen
        If Not client.Inbox.IsOpen Then client.Inbox.Open(FolderAccess.ReadOnly)

        Dim sinceDate = forwardSearchSince.Value.Date
        forwardingRepository.SetSetting("ForwardSearchSinceDate", sinceDate.ToString("yyyy-MM-dd"))
        Dim serverQuery = BuildForwardingSearchQuery(forwardingPreviewRule, sinceDate)
        Dim scanUids = client.Inbox.Search(serverQuery).Reverse().ToList()
        If scanUids.Count = 0 Then
            MessageBox.Show($"No likely matches found in the inbox since {sinceDate:d}.", "Preview matches", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        btnScanForwarding.Enabled = False
        btnForwardPreviewed.Enabled = False
        forwardingPreviewMessages.Clear()
        Dim previewTable As New DataTable()
        previewTable.Columns.Add("Sender name")
        previewTable.Columns.Add("Email address")
        previewTable.Columns.Add("Message count", GetType(Integer))
        previewTable.Columns.Add("Oldest", GetType(DateTimeOffset))
        previewTable.Columns.Add("Newest", GetType(DateTimeOffset))
        Dim summaryRows As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
        tsProgressBar.Value = 0
        tsProgressBar.Maximum = Math.Max(1, scanUids.Count)
        ToolStrip.Visible = True
        Try
            For Each uid In scanUids
                Dim message = client.Inbox.GetMessage(uid)
                If ForwardingService.RuleMatches(forwardingPreviewRule, message) Then
                    forwardingPreviewMessages(CInt(uid.Id)) = message
                    Dim mailbox = message.From.Mailboxes.FirstOrDefault()
                    If mailbox IsNot Nothing Then
                        Dim address = mailbox.Address
                        If Not summaryRows.ContainsKey(address) Then
                            Dim summaryRow = previewTable.NewRow()
                            summaryRow("Sender name") = If(String.IsNullOrWhiteSpace(mailbox.Name), "(no display name)", mailbox.Name)
                            summaryRow("Email address") = address
                            summaryRow("Message count") = 1
                            summaryRow("Oldest") = message.Date
                            summaryRow("Newest") = message.Date
                            previewTable.Rows.Add(summaryRow)
                            summaryRows(address) = summaryRow
                        Else
                            Dim summaryRow = summaryRows(address)
                            summaryRow("Message count") = CInt(summaryRow("Message count")) + 1
                            If message.Date < DirectCast(summaryRow("Oldest"), DateTimeOffset) Then summaryRow("Oldest") = message.Date
                            If message.Date > DirectCast(summaryRow("Newest"), DateTimeOffset) Then summaryRow("Newest") = message.Date
                        End If
                    End If
                End If
                tsProgressBar.Value += 1
                UpdatetsStatusText($"Previewing {tsProgressBar.Value} of {scanUids.Count}: {message.From} — {message.Subject}")
            Next
            forwardingPreviewGrid.DataSource = previewTable
            btnForwardPreviewed.Enabled = forwardingPreviewMessages.Count > 0
            MessageBox.Show($"Preview complete. Found {forwardingPreviewMessages.Count} matching message(s) from {summaryRows.Count} email address(es). Nothing was sent.",
                            "Preview matches", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            eUtil.LOGIT($"Forwarding preview failed: {ex.Message}", True)
            MessageBox.Show($"Preview failed: {ex.Message}", "Preview matches", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnScanForwarding.Enabled = True
        End Try
    End Sub

    Private Shared Function BuildForwardingSearchQuery(rule As ForwardingRule, sinceDate As DateTime) As SearchQuery
        Dim query As SearchQuery = SearchQuery.DeliveredAfter(sinceDate.AddDays(-1))
        If String.Equals(rule.MatchType, "Sender name", StringComparison.OrdinalIgnoreCase) Then
            Dim nameTokens = System.Text.RegularExpressions.Regex.Split(rule.MatchValue.Trim(), "\W+").
                Where(Function(token) Not String.IsNullOrWhiteSpace(token))
            For Each token In nameTokens
                query = query.And(SearchQuery.FromContains(token))
            Next
        Else
            query = query.And(SearchQuery.FromContains(rule.MatchValue.Trim().TrimStart("@"c)))
        End If
        Return query
    End Function

    Private Sub btnForwardPreviewed_Click(sender As Object, e As EventArgs) Handles btnForwardPreviewed.Click
        If forwardingRepository Is Nothing OrElse forwardingPreviewRule Is Nothing OrElse forwardingPreviewMessages.Count = 0 Then Exit Sub
        Dim destination = txtForwardDestination.Text.Trim()
        If Not ValidateGmailDestination(destination) Then Exit Sub
        If MessageBox.Show($"Forward all {forwardingPreviewMessages.Count} previewed message(s) to {destination}?",
                           "Forward previewed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub

        Dim previewSendRule As New ForwardingRule With {
            .MatchType = forwardingPreviewRule.MatchType,
            .MatchValue = forwardingPreviewRule.MatchValue,
            .Destination = destination,
            .Enabled = True
        }
        Dim previewRules As New List(Of ForwardingRule) From {previewSendRule}
        btnForwardPreviewed.Enabled = False
        tsProgressBar.Value = 0
        tsProgressBar.Maximum = Math.Max(1, forwardingPreviewMessages.Count)
        Dim forwardedCount = 0
        Try
            For Each previewMessage In forwardingPreviewMessages
                forwardedCount += ProcessAutoForward(previewMessage.Value, previewMessage.Key, client.Inbox.FullName, True, previewRules)
                tsProgressBar.Value += 1
                UpdatetsStatusText($"Forwarding preview {tsProgressBar.Value} of {forwardingPreviewMessages.Count}")
            Next
            MessageBox.Show($"Forwarding complete. {forwardedCount} message(s) sent; previously sent messages were skipped.",
                            "Forward previewed", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            eUtil.LOGIT($"Forward preview failed: {ex.Message}", True)
            MessageBox.Show($"Forwarding failed: {ex.Message}", "Forward previewed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnForwardPreviewed.Enabled = True
        End Try
    End Sub

    Private Sub btnSaveForwardRule_Click(sender As Object, e As EventArgs) Handles btnSaveForwardRule.Click
        If forwardingRepository Is Nothing Then Exit Sub
        If Not ValidateMatchCriteria() Then Exit Sub
        Dim matchValue = txtForwardMatchValue.Text.Trim()
        Dim destination = txtForwardDestination.Text.Trim()
        If Not ValidateGmailDestination(destination) Then Exit Sub

        Dim rule As New ForwardingRule With {
            .Id = selectedForwardingRuleId,
            .MatchType = cmbForwardMatchType.Text,
            .MatchValue = matchValue,
            .Destination = destination,
            .Enabled = cbForwardRuleEnabled.Checked
        }
        selectedForwardingRuleId = forwardingRepository.SaveRule(rule)
        RefreshForwardingRules()
        btnSaveForwardRule.Text = "Update rule"
        btnDeleteForwardRule.Enabled = True
    End Sub

    Private Sub btnDeleteForwardRule_Click(sender As Object, e As EventArgs) Handles btnDeleteForwardRule.Click
        If forwardingRepository Is Nothing OrElse selectedForwardingRuleId = 0 Then Exit Sub
        If MessageBox.Show("Delete the selected forwarding rule?", "Forwarding rule", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub
        forwardingRepository.DeleteRule(selectedForwardingRuleId)
        ClearForwardingEditor()
        RefreshForwardingRules()
    End Sub

    Private Sub btnUseSelectedSender_Click(sender As Object, e As EventArgs) Handles btnUseSelectedSender.Click
        If dgvEmails.CurrentRow Is Nothing OrElse Not dgvEmails.Columns.Contains("From") Then
            MessageBox.Show("Select an email on the Email screening tab first.", "Forwarding rule", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        txtForwardMatchValue.Text = ExtractEmailAddress(Convert.ToString(dgvEmails.CurrentRow.Cells("From").Value))
        cmbForwardMatchType.SelectedItem = "Sender"
    End Sub

    Private Sub forwardingGrid_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        Dim rule = TryCast(forwardingGrid.Rows(e.RowIndex).DataBoundItem, ForwardingRule)
        If rule Is Nothing Then Exit Sub
        selectedForwardingRuleId = rule.Id
        cmbForwardMatchType.SelectedItem = rule.MatchType
        txtForwardMatchValue.Text = rule.MatchValue
        txtForwardDestination.Text = rule.Destination
        cbForwardRuleEnabled.Checked = rule.Enabled
        btnSaveForwardRule.Text = "Update rule"
        btnDeleteForwardRule.Enabled = True
    End Sub

    Private Sub ClearForwardingEditor()
        selectedForwardingRuleId = 0
        cmbForwardMatchType.SelectedIndex = 0
        txtForwardMatchValue.Clear()
        txtForwardDestination.Clear()
        cbForwardRuleEnabled.Checked = True
        btnSaveForwardRule.Text = "Add rule"
        btnDeleteForwardRule.Enabled = False
    End Sub

    Private Shared Function IsValidEmailAddress(value As String) As Boolean
        Try
            Dim address As New System.Net.Mail.MailAddress(value)
            Return String.Equals(address.Address, value, StringComparison.OrdinalIgnoreCase)
        Catch ex As FormatException
            Return False
        End Try
    End Function

    Private Function ValidateMatchCriteria() As Boolean
        Dim matchValue = txtForwardMatchValue.Text.Trim()
        If String.IsNullOrWhiteSpace(matchValue) Then
            MessageBox.Show("Enter a sender email address, sender name, or domain.", "Forwarding rule", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        If cmbForwardMatchType.Text = "Sender" AndAlso Not IsValidEmailAddress(matchValue) Then
            If matchValue.Contains(" "c) Then
                cmbForwardMatchType.SelectedItem = "Sender name"
            Else
                MessageBox.Show("A Sender rule requires a complete email address. Choose Sender name to search by a person's name.",
                                "Forwarding rule", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
        End If
        Return True
    End Function

    Private Shared Function ValidateGmailDestination(destination As String) As Boolean
        If Not IsValidEmailAddress(destination) Then
            MessageBox.Show("Enter a valid destination Gmail address.", "Forwarding rule", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If Not destination.EndsWith("@gmail.com", StringComparison.OrdinalIgnoreCase) AndAlso
           Not destination.EndsWith("@googlemail.com", StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("The forwarding destination must be a Gmail address.", "Forwarding rule", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        Return True
    End Function

    Private Shared Function ExtractEmailAddress(value As String) As String
        Dim match = System.Text.RegularExpressions.Regex.Match(value, "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If match.Success Then Return match.Value
        Return value.Trim()
    End Function

    Private Sub Emails_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
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
            If mbr = MsgBoxResult.Ok Then
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
            tbMailClient.Text = cmbEmailClients.SelectedItem.ToString()
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
        If tbMailClient.Text <> "" Then
            If cmbEmailClients.SelectedItem <> tbMailClient.Text Then
                If IsClientConnected() Then
                    SafeDisconnectClient()
                    btnConnect.Visible = True
                End If
            End If
        End If
        ApplySelectedEmailClientConfiguration()
    End Sub

    Private Sub ApplySelectedEmailClientConfiguration()
        If cmbEmailClients.SelectedItem Is Nothing Then Exit Sub
        For Each emailClient As String In eUtil.lEmailClients
            Dim parts = emailClient.Split(","c)
            If parts.Length >= 4 AndAlso parts(0) = cmbEmailClients.SelectedItem.ToString() Then
                eUtil.sThisEmailService = parts(1)
                eUtil.sThisEmailUser = parts(2)
                eUtil.sThisEmailPassword = parts(3)
                Exit For
            End If
        Next
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
