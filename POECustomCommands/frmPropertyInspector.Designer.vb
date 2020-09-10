<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPropertyInspector
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
        Me.lvProperties = New System.Windows.Forms.ListView()
        Me.Column1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Column2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CheckBoxLongFormat = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'lvProperties
        '
        Me.lvProperties.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Column1, Me.Column2})
        Me.lvProperties.Location = New System.Drawing.Point(-1, 30)
        Me.lvProperties.Name = "lvProperties"
        Me.lvProperties.Size = New System.Drawing.Size(322, 791)
        Me.lvProperties.TabIndex = 0
        Me.lvProperties.UseCompatibleStateImageBehavior = False
        Me.lvProperties.View = System.Windows.Forms.View.Details
        '
        'Column1
        '
        Me.Column1.Text = ""
        Me.Column1.Width = 136
        '
        'Column2
        '
        Me.Column2.Text = ""
        Me.Column2.Width = 182
        '
        'CheckBoxLongFormat
        '
        Me.CheckBoxLongFormat.AutoSize = True
        Me.CheckBoxLongFormat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CheckBoxLongFormat.Location = New System.Drawing.Point(232, 7)
        Me.CheckBoxLongFormat.Name = "CheckBoxLongFormat"
        Me.CheckBoxLongFormat.Size = New System.Drawing.Size(75, 17)
        Me.CheckBoxLongFormat.TabIndex = 1
        Me.CheckBoxLongFormat.Text = "long format"
        Me.CheckBoxLongFormat.UseVisualStyleBackColor = True
        '
        'frmPropertyInspector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(322, 833)
        Me.Controls.Add(Me.CheckBoxLongFormat)
        Me.Controls.Add(Me.lvProperties)
        Me.Name = "frmPropertyInspector"
        Me.Text = "POE Property Inspector"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lvProperties As Windows.Forms.ListView
    Friend WithEvents Column1 As Windows.Forms.ColumnHeader
    Friend WithEvents Column2 As Windows.Forms.ColumnHeader
    Friend WithEvents CheckBoxLongFormat As Windows.Forms.CheckBox
End Class
