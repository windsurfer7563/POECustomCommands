<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDefineWorksace
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
        Me.lvObjs = New System.Windows.Forms.ListView()
        Me.OutPut = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'lvObjs
        '
        Me.lvObjs.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.OutPut})
        Me.lvObjs.GridLines = True
        Me.lvObjs.Location = New System.Drawing.Point(44, 132)
        Me.lvObjs.Name = "lvObjs"
        Me.lvObjs.Size = New System.Drawing.Size(373, 170)
        Me.lvObjs.TabIndex = 3
        Me.lvObjs.UseCompatibleStateImageBehavior = False
        Me.lvObjs.View = System.Windows.Forms.View.Details
        '
        'OutPut
        '
        Me.OutPut.Text = "OutPut"
        Me.OutPut.Width = 177
        '
        'frmDefineWorksace
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(773, 314)
        Me.Controls.Add(Me.lvObjs)
        Me.Name = "frmDefineWorksace"
        Me.Text = "frmDefineWorksace"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lvObjs As Windows.Forms.ListView
    Friend WithEvents OutPut As Windows.Forms.ColumnHeader
End Class
