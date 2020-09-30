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
        Me.CheckBoxLongFormat = New System.Windows.Forms.CheckBox()
        Me.lvProperties = New System.Windows.Forms.DataGridView()
        Me.PropertyName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PropertyValue = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ButtonColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.lvProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CheckBoxLongFormat
        '
        Me.CheckBoxLongFormat.AutoSize = True
        Me.CheckBoxLongFormat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CheckBoxLongFormat.Location = New System.Drawing.Point(221, 12)
        Me.CheckBoxLongFormat.Name = "CheckBoxLongFormat"
        Me.CheckBoxLongFormat.Size = New System.Drawing.Size(75, 17)
        Me.CheckBoxLongFormat.TabIndex = 1
        Me.CheckBoxLongFormat.Text = "long format"
        Me.CheckBoxLongFormat.UseVisualStyleBackColor = True
        '
        'lvProperties
        '
        Me.lvProperties.AllowUserToAddRows = False
        Me.lvProperties.AllowUserToDeleteRows = False
        Me.lvProperties.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.lvProperties.BackgroundColor = System.Drawing.SystemColors.Window
        Me.lvProperties.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.lvProperties.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PropertyName, Me.PropertyValue, Me.ButtonColumn})
        Me.lvProperties.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lvProperties.Location = New System.Drawing.Point(0, 42)
        Me.lvProperties.MultiSelect = False
        Me.lvProperties.Name = "lvProperties"
        Me.lvProperties.ReadOnly = True
        Me.lvProperties.RowHeadersVisible = False
        Me.lvProperties.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.lvProperties.Size = New System.Drawing.Size(326, 791)
        Me.lvProperties.TabIndex = 2
        '
        'PropertyName
        '
        Me.PropertyName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PropertyName.HeaderText = "Property"
        Me.PropertyName.MinimumWidth = 137
        Me.PropertyName.Name = "PropertyName"
        Me.PropertyName.ReadOnly = True
        Me.PropertyName.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.PropertyName.Width = 137
        '
        'PropertyValue
        '
        Me.PropertyValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.PropertyValue.HeaderText = "Value"
        Me.PropertyValue.MinimumWidth = 150
        Me.PropertyValue.Name = "PropertyValue"
        Me.PropertyValue.ReadOnly = True
        Me.PropertyValue.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.PropertyValue.Width = 150
        '
        'ButtonColumn
        '
        Me.ButtonColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ButtonColumn.HeaderText = ""
        Me.ButtonColumn.MinimumWidth = 20
        Me.ButtonColumn.Name = "ButtonColumn"
        Me.ButtonColumn.ReadOnly = True
        Me.ButtonColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.ButtonColumn.Width = 20
        '
        'frmPropertyInspector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(326, 833)
        Me.Controls.Add(Me.lvProperties)
        Me.Controls.Add(Me.CheckBoxLongFormat)
        Me.Name = "frmPropertyInspector"
        Me.Text = "POE Property Inspector"
        Me.TopMost = True
        CType(Me.lvProperties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CheckBoxLongFormat As Windows.Forms.CheckBox
    Friend WithEvents lvProperties As Windows.Forms.DataGridView
    Friend WithEvents PropertyName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PropertyValue As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ButtonColumn As Windows.Forms.DataGridViewTextBoxColumn
End Class
