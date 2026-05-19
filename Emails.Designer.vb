<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Emails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        dgvEmails = New DataGridView()
        ToolStrip = New ToolStrip()
        tsProgressBar = New ToolStripProgressBar()
        tsStatusText = New ToolStripLabel()
        cbSelectAll = New CheckBox()
        cbSpam = New CheckBox()
        btnRefreshGrid = New Button()
        btnSave = New Button()
        GroupBox2 = New GroupBox()
        Label2 = New Label()
        tbSpam = New TextBox()
        Label1 = New Label()
        tbUnread = New TextBox()
        btnMarkRead = New Button()
        cbRead = New CheckBox()
        cbUnRead = New CheckBox()
        tbMailClient = New TextBox()
        Label3 = New Label()
        cmbEmailClients = New ComboBox()
        Label4 = New Label()
        btnConnect = New Button()
        gbFuture = New GroupBox()
        cbFolders = New ComboBox()
        btnCountEmails = New Button()
        Label5 = New Label()
        CheckBox1 = New CheckBox()
        gbFilter = New GroupBox()
        Button1 = New Button()
        CType(dgvEmails, ComponentModel.ISupportInitialize).BeginInit()
        ToolStrip.SuspendLayout()
        GroupBox2.SuspendLayout()
        gbFuture.SuspendLayout()
        gbFilter.SuspendLayout()
        SuspendLayout()
        ' 
        ' dgvEmails
        ' 
        dgvEmails.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvEmails.Location = New Point(28, 140)
        dgvEmails.Name = "dgvEmails"
        dgvEmails.RowTemplate.Height = 25
        dgvEmails.Size = New Size(1235, 418)
        dgvEmails.TabIndex = 0
        ' 
        ' ToolStrip
        ' 
        ToolStrip.Dock = DockStyle.Bottom
        ToolStrip.Items.AddRange(New ToolStripItem() {tsProgressBar, tsStatusText})
        ToolStrip.Location = New Point(0, 710)
        ToolStrip.Name = "ToolStrip"
        ToolStrip.Size = New Size(1322, 25)
        ToolStrip.TabIndex = 1
        ToolStrip.Text = "ToolStrip1"
        ' 
        ' tsProgressBar
        ' 
        tsProgressBar.Name = "tsProgressBar"
        tsProgressBar.Size = New Size(100, 22)
        ' 
        ' tsStatusText
        ' 
        tsStatusText.Name = "tsStatusText"
        tsStatusText.Size = New Size(88, 22)
        tsStatusText.Text = "ToolStripLabel1"
        ' 
        ' cbSelectAll
        ' 
        cbSelectAll.AutoSize = True
        cbSelectAll.Location = New Point(209, 643)
        cbSelectAll.Name = "cbSelectAll"
        cbSelectAll.Size = New Size(99, 19)
        cbSelectAll.TabIndex = 3
        cbSelectAll.Text = "Mark All Read"
        cbSelectAll.UseVisualStyleBackColor = True
        ' 
        ' cbSpam
        ' 
        cbSpam.AutoSize = True
        cbSpam.Location = New Point(209, 666)
        cbSpam.Name = "cbSpam"
        cbSpam.Size = New Size(84, 19)
        cbSpam.TabIndex = 11
        cbSpam.Text = "Spam Only"
        cbSpam.UseVisualStyleBackColor = True
        ' 
        ' btnRefreshGrid
        ' 
        btnRefreshGrid.Location = New Point(61, 42)
        btnRefreshGrid.Name = "btnRefreshGrid"
        btnRefreshGrid.Size = New Size(118, 23)
        btnRefreshGrid.TabIndex = 12
        btnRefreshGrid.Text = "Refresh Grid"
        btnRefreshGrid.UseVisualStyleBackColor = True
        ' 
        ' btnSave
        ' 
        btnSave.Location = New Point(61, 614)
        btnSave.Name = "btnSave"
        btnSave.Size = New Size(118, 23)
        btnSave.TabIndex = 13
        btnSave.Text = "Save Grid to CSV"
        btnSave.UseVisualStyleBackColor = True
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(Label2)
        GroupBox2.Controls.Add(tbSpam)
        GroupBox2.Controls.Add(Label1)
        GroupBox2.Controls.Add(tbUnread)
        GroupBox2.Location = New Point(367, 602)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(351, 96)
        GroupBox2.TabIndex = 14
        GroupBox2.TabStop = False
        GroupBox2.Text = "Totals"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(185, 28)
        Label2.Name = "Label2"
        Label2.Size = New Size(66, 15)
        Label2.TabIndex = 12
        Label2.Text = "Total Spam"
        ' 
        ' tbSpam
        ' 
        tbSpam.Location = New Point(185, 49)
        tbSpam.Name = "tbSpam"
        tbSpam.ReadOnly = True
        tbSpam.Size = New Size(100, 23)
        tbSpam.TabIndex = 11
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(66, 29)
        Label1.Name = "Label1"
        Label1.Size = New Size(74, 15)
        Label1.TabIndex = 10
        Label1.Text = "Total Unread"
        ' 
        ' tbUnread
        ' 
        tbUnread.Location = New Point(66, 50)
        tbUnread.Name = "tbUnread"
        tbUnread.ReadOnly = True
        tbUnread.Size = New Size(100, 23)
        tbUnread.TabIndex = 9
        ' 
        ' btnMarkRead
        ' 
        btnMarkRead.Location = New Point(209, 614)
        btnMarkRead.Name = "btnMarkRead"
        btnMarkRead.Size = New Size(118, 23)
        btnMarkRead.TabIndex = 2
        btnMarkRead.Text = "Mark Selected Read"
        btnMarkRead.UseVisualStyleBackColor = True
        ' 
        ' cbRead
        ' 
        cbRead.AutoSize = True
        cbRead.Checked = True
        cbRead.CheckState = CheckState.Checked
        cbRead.Location = New Point(196, 22)
        cbRead.Name = "cbRead"
        cbRead.Size = New Size(97, 19)
        cbRead.TabIndex = 19
        cbRead.Text = "Retrieve Read"
        cbRead.UseVisualStyleBackColor = True
        ' 
        ' cbUnRead
        ' 
        cbUnRead.AutoSize = True
        cbUnRead.Location = New Point(196, 47)
        cbUnRead.Name = "cbUnRead"
        cbUnRead.Size = New Size(112, 19)
        cbUnRead.TabIndex = 20
        cbUnRead.Text = "Retrieve UnRead"
        cbUnRead.UseVisualStyleBackColor = True
        ' 
        ' tbMailClient
        ' 
        tbMailClient.Location = New Point(378, 43)
        tbMailClient.Name = "tbMailClient"
        tbMailClient.Size = New Size(100, 23)
        tbMailClient.TabIndex = 21
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(378, 26)
        Label3.Name = "Label3"
        Label3.Size = New Size(64, 15)
        Label3.TabIndex = 22
        Label3.Text = "Mail Client"
        ' 
        ' cmbEmailClients
        ' 
        cmbEmailClients.FormattingEnabled = True
        cmbEmailClients.Location = New Point(496, 42)
        cmbEmailClients.Name = "cmbEmailClients"
        cmbEmailClients.Size = New Size(121, 23)
        cmbEmailClients.TabIndex = 24
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(496, 24)
        Label4.Name = "Label4"
        Label4.Size = New Size(81, 15)
        Label4.TabIndex = 23
        Label4.Text = "Email Services"
        ' 
        ' btnConnect
        ' 
        btnConnect.Location = New Point(623, 42)
        btnConnect.Name = "btnConnect"
        btnConnect.Size = New Size(118, 23)
        btnConnect.TabIndex = 25
        btnConnect.Text = "Connect to Email"
        btnConnect.UseVisualStyleBackColor = True
        ' 
        ' gbFuture
        ' 
        gbFuture.Controls.Add(cbFolders)
        gbFuture.Controls.Add(btnCountEmails)
        gbFuture.Controls.Add(Label5)
        gbFuture.Location = New Point(810, 12)
        gbFuture.Name = "gbFuture"
        gbFuture.Size = New Size(200, 100)
        gbFuture.TabIndex = 26
        gbFuture.TabStop = False
        gbFuture.Text = "Future Enhancements"
        gbFuture.Visible = False
        ' 
        ' cbFolders
        ' 
        cbFolders.FormattingEnabled = True
        cbFolders.Location = New Point(15, 37)
        cbFolders.Name = "cbFolders"
        cbFolders.Size = New Size(121, 23)
        cbFolders.TabIndex = 6
        ' 
        ' btnCountEmails
        ' 
        btnCountEmails.Location = New Point(15, 66)
        btnCountEmails.Name = "btnCountEmails"
        btnCountEmails.Size = New Size(121, 23)
        btnCountEmails.TabIndex = 3
        btnCountEmails.Text = "Count Emails"
        btnCountEmails.UseVisualStyleBackColor = True
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(15, 19)
        Label5.Name = "Label5"
        Label5.Size = New Size(77, 15)
        Label5.TabIndex = 5
        Label5.Text = "Email Folders"
        ' 
        ' CheckBox1
        ' 
        CheckBox1.AutoSize = True
        CheckBox1.Location = New Point(6, 22)
        CheckBox1.Name = "CheckBox1"
        CheckBox1.Size = New Size(79, 19)
        CheckBox1.TabIndex = 27
        CheckBox1.Text = "Important"
        CheckBox1.UseVisualStyleBackColor = True
        ' 
        ' gbFilter
        ' 
        gbFilter.Controls.Add(CheckBox1)
        gbFilter.Location = New Point(196, 72)
        gbFilter.Name = "gbFilter"
        gbFilter.Size = New Size(245, 62)
        gbFilter.TabIndex = 28
        gbFilter.TabStop = False
        gbFilter.Text = "Filter"
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(61, 91)
        Button1.Name = "Button1"
        Button1.Size = New Size(118, 23)
        Button1.TabIndex = 29
        Button1.Text = "Convert to SQLite"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Emails
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1322, 735)
        Controls.Add(Button1)
        Controls.Add(gbFilter)
        Controls.Add(gbFuture)
        Controls.Add(btnConnect)
        Controls.Add(cmbEmailClients)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(tbMailClient)
        Controls.Add(cbUnRead)
        Controls.Add(cbRead)
        Controls.Add(btnSave)
        Controls.Add(btnMarkRead)
        Controls.Add(GroupBox2)
        Controls.Add(cbSelectAll)
        Controls.Add(btnRefreshGrid)
        Controls.Add(cbSpam)
        Controls.Add(ToolStrip)
        Controls.Add(dgvEmails)
        Name = "Emails"
        Text = "Emails"
        CType(dgvEmails, ComponentModel.ISupportInitialize).EndInit()
        ToolStrip.ResumeLayout(False)
        ToolStrip.PerformLayout()
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        gbFuture.ResumeLayout(False)
        gbFuture.PerformLayout()
        gbFilter.ResumeLayout(False)
        gbFilter.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents dgvEmails As DataGridView
    Friend WithEvents ToolStrip As ToolStrip
    Friend WithEvents tsStatusText As ToolStripLabel
    Friend WithEvents tsProgressBar As ToolStripProgressBar
    Friend WithEvents cbSelectAll As CheckBox
    Friend WithEvents cbSpam As CheckBox
    Friend WithEvents btnRefreshGrid As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents tbSpam As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbUnread As TextBox
    Friend WithEvents btnMarkRead As Button
    Friend WithEvents cbRead As CheckBox
    Friend WithEvents cbUnRead As CheckBox
    Friend WithEvents tbMailClient As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents cmbEmailClients As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents btnConnect As Button
    Friend WithEvents gbFuture As GroupBox
    Friend WithEvents cbFolders As ComboBox
    Friend WithEvents btnCountEmails As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents gbFilter As GroupBox
    Friend WithEvents Button1 As Button
End Class
