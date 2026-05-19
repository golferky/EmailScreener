<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreateEmailFile
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
        ToolStrip = New ToolStrip()
        tsProgressBar = New ToolStripProgressBar()
        tsStatusText = New ToolStripLabel()
        DataGridView1 = New DataGridView()
        ToolStrip.SuspendLayout()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' ToolStrip
        ' 
        ToolStrip.Dock = DockStyle.Bottom
        ToolStrip.Items.AddRange(New ToolStripItem() {tsProgressBar, tsStatusText})
        ToolStrip.Location = New Point(0, 425)
        ToolStrip.Name = "ToolStrip"
        ToolStrip.Size = New Size(1185, 25)
        ToolStrip.TabIndex = 2
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
        tsStatusText.Size = New Size(87, 22)
        tsStatusText.Text = "ToolStripLabel1"
        ' 
        ' DataGridView1
        ' 
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Location = New Point(12, 111)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.RowTemplate.Height = 25
        DataGridView1.Size = New Size(1145, 268)
        DataGridView1.TabIndex = 3
        ' 
        ' CreateEmailFile
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1185, 450)
        Controls.Add(DataGridView1)
        Controls.Add(ToolStrip)
        Name = "CreateEmailFile"
        Text = "CreateEmailFile"
        ToolStrip.ResumeLayout(False)
        ToolStrip.PerformLayout()
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents ToolStrip As ToolStrip
    Friend WithEvents tsProgressBar As ToolStripProgressBar
    Friend WithEvents tsStatusText As ToolStripLabel
    Friend WithEvents DataGridView1 As DataGridView
End Class
