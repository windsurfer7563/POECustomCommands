﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGetObjectsOnDwg
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
        Me.treeViewDeliverableRoot = New System.Windows.Forms.TreeView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'treeViewDeliverableRoot
        '
        Me.treeViewDeliverableRoot.Location = New System.Drawing.Point(12, 12)
        Me.treeViewDeliverableRoot.Name = "treeViewDeliverableRoot"
        Me.treeViewDeliverableRoot.Size = New System.Drawing.Size(297, 415)
        Me.treeViewDeliverableRoot.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(323, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Create CSV"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmGetObjectsOnDwg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(410, 439)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.treeViewDeliverableRoot)
        Me.Name = "frmGetObjectsOnDwg"
        Me.Text = "Export Drawing Data"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents treeViewDeliverableRoot As Windows.Forms.TreeView
    Friend WithEvents Button1 As Windows.Forms.Button
End Class
