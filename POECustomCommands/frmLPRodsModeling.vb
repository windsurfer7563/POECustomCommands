Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Windows.Forms
Imports System.Windows

Public Class frmLPRodsModeling
    Public Event OnFinish()
    Public Event OnClose()
    Public Event OnNameChanged()
    Public Event MoveToNewPosition(ByVal Ps As Position)
    Public Event OnPositionChanged(ByVal Ps As Position)
    Public Event OnNewPositionComputed(ByVal Ps As Position)
    Private m_oUOM As UOMManager
    Private m_Pos1 As New Position
    Private m_Pos2 As New Position
    Private m_Pos3 As New Position
    Private m_NPos1 As New Position
    Private m_NPos2 As New Position
    Private m_Pos As New Position


    Public Property EqpPosition1() As Position
        Get
            Return m_Pos1
        End Get
        Set(ByVal value As Position)
            m_Pos1 = value
            txtX1.Text = FormatDistanceValue(m_Pos1.X)
            txtY1.Text = FormatDistanceValue(m_Pos1.Y)
            txtZ1.Text = FormatDistanceValue(m_Pos1.Z)
        End Set
    End Property
    Public Property EqpPosition2() As Position
        Get
            Return m_Pos2
        End Get
        Set(ByVal value As Position)
            m_Pos2 = value
            txtX2.Text = FormatDistanceValue(m_Pos2.X)
            txtY2.Text = FormatDistanceValue(m_Pos2.Y)
            txtZ2.Text = FormatDistanceValue(m_Pos2.Z)
        End Set
    End Property
    Public Property EqpPosition3() As Position
        Get
            Return m_Pos3
        End Get
        Set(ByVal value As Position)
            m_Pos3 = value
            txtX3.Text = FormatDistanceValue(m_Pos3.X)
            txtY3.Text = FormatDistanceValue(m_Pos3.Y)
            txtZ3.Text = FormatDistanceValue(m_Pos3.Z)
        End Set
    End Property

    Public Property CalcPosition1() As Position
        Get
            Return m_NPos1
        End Get
        Set(ByVal value As Position)
            m_NPos1 = value
            txtNX1.Text = FormatDistanceValue(m_NPos1.X)
            txtNY1.Text = FormatDistanceValue(m_NPos1.Y)
            txtNZ1.Text = FormatDistanceValue(m_NPos1.Z)
        End Set
    End Property
    Public Property CalcPosition2() As Position
        Get
            Return m_NPos2
        End Get
        Set(ByVal value As Position)
            m_NPos2 = value
            txtNX2.Text = FormatDistanceValue(m_NPos2.X)
            txtNY2.Text = FormatDistanceValue(m_NPos2.Y)
            txtNZ2.Text = FormatDistanceValue(m_NPos2.Z)
        End Set
    End Property

    Public Property EqpName() As String
        Get
            Return txtName.Text
        End Get
        Set(ByVal value As String)
            txtName.Text = value
        End Set
    End Property

    Public Property EqpName1() As String
        Get
            Return txtName1.Text
        End Get
        Set(ByVal value As String)
            txtName1.Text = value
        End Set
    End Property
    Public Property EqpName2() As String
        Get
            Return txtName2.Text
        End Get
        Set(ByVal value As String)
            txtName2.Text = value
        End Set
    End Property
    Public Property EqpName3() As String
        Get
            Return txtName3.Text
        End Get
        Set(ByVal value As String)
            txtName3.Text = value
        End Set
    End Property


    Private Sub btnFinish_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinish.Click
        RaiseEvent OnFinish()
    End Sub

    Private Sub txtName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            RaiseEvent OnNameChanged()
        End If
    End Sub

    Private Sub txtName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.LostFocus
        RaiseEvent OnNameChanged()
    End Sub

    'called whenever a coordinate value is changed.
    'It will try to parse the value, and if it fails,
    'it will show the value in RED
    Private Function CheckCoordinate(ByVal sender As Object, ByRef dValue As Double) As Boolean

        Dim oTextBox As Windows.Forms.TextBox = sender
        Dim bResult As Boolean = True

        Try
            dValue = m_oUOM.ParseUnit(UnitType.Distance, oTextBox.Text)
        Catch ex As Exception
            bResult = False
            oTextBox.ForeColor = Drawing.Color.Red
        End Try

        Return bResult
    End Function

    'comment Above two methods and uncomment this one.
    Private Function FormatDistanceValue(ByVal dValue) As String
        Return MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dValue)
    End Function

    Private Sub frmLPRodsModeling_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        RaiseEvent OnNewPositionComputed(m_Pos)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RaiseEvent MoveToNewPosition(m_Pos1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RaiseEvent OnClose()
    End Sub

    Private Sub frmLPRodsModeling_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True

    End Sub
End Class
