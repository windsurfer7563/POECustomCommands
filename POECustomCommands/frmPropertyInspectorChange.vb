
Public Class frmPropertyInspectorChange
    Inherits System.Windows.Forms.Form

    Private propertyType As PropertyTypes
    Private values As List(Of String)
    Private currPropValue As String
    Private textBox1 As System.Windows.Forms.TextBox
    Private comboBox1 As System.Windows.Forms.ComboBox
    Private WithEvents ButtonOK As Windows.Forms.Button
    Private WithEvents ButtonCancel As Windows.Forms.Button
    Public result As String


    Sub New(propType As PropertyTypes, currValue As String, Optional values As List(Of String) = Nothing)
        MyBase.New()
        Me.propertyType = propType
        Me.currPropValue = currValue
        Me.values = values
        Me.Text = "Edit Property"
        InitializeComponent()

    End Sub


    Private Sub InitializeComponent()

        If propertyType = PropertyTypes.text Or propertyType = PropertyTypes.distance Then
            Me.textBox1 = New System.Windows.Forms.TextBox()
            Me.textBox1.Location = New System.Drawing.Point(10, 20)
            Me.textBox1.Size = New System.Drawing.Size(280, 20)
            Me.textBox1.TabIndex = 0
            Me.textBox1.Text = currPropValue
        End If

        If propertyType = PropertyTypes.codelist Then
            Me.comboBox1 = New System.Windows.Forms.ComboBox()
            Me.comboBox1.DropDownWidth = 280
            Me.comboBox1.Location = New System.Drawing.Point(10, 20)
            Me.comboBox1.Size = New System.Drawing.Size(280, 20)
            Me.comboBox1.TabIndex = 0
            If values.Count > 0 Then
                For i As Integer = 0 To values.Count - 1
                    Me.comboBox1.Items.Add(values.Item(i))
                Next
                Me.comboBox1.Text = currPropValue
            End If

        End If


            Me.ButtonOK = New System.Windows.Forms.Button()
        Me.ButtonCancel = New System.Windows.Forms.Button()

        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(84, 136)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(70, 22)
        Me.ButtonOK.TabIndex = 1
        Me.ButtonOK.Text = "OK"
        Me.ButtonOK.UseVisualStyleBackColor = True
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(176, 136)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(66, 22)
        Me.ButtonCancel.TabIndex = 2
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'frmPropertyInspectorChange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(328, 170)
        Me.ControlBox = False
        If propertyType = PropertyTypes.text Or propertyType = PropertyTypes.distance Then
            Me.Controls.Add(Me.textBox1)
        Else
            Me.Controls.Add(Me.comboBox1)
        End If

        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)

        Me.Name = "frmPropertyInspectorChange"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmPropertyInspectorChange"
        Me.ResumeLayout(False)
    End Sub

    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click
        DialogResult = Windows.Forms.DialogResult.OK
        If propertyType = PropertyTypes.text Or propertyType = PropertyTypes.distance Then
            Me.result = textBox1.Text
        Else
            Me.result = comboBox1.Text
        End If
        Me.Close()
    End Sub

    Private Sub frmPropertyInspectorChange_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

End Class
