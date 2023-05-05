Option Strict On
Option Explicit On

Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Equipment.Middle
Imports System.Globalization

'import this to avoid 'long names' for mouse button
'constants in OnMouseMove and OnMouseDown
Imports Ingr.SP3D.Common.Client.Services.GraphicViewManager.GraphicViewEventArgs

Public Class LPRodsModeling
    'Inherits BaseGraphicCommand ' we need Graphic Events ok
    'Inherits BaseModalCommand
    Inherits BaseStepCommand


    Private WithEvents m_frmLPRodsModeling As frmLPRodsModeling ' our Form
    Private m_oEqp1 As Equipment
    Private m_oEqp2 As Equipment
    Private m_oEqp3 As Equipment
    Private n_Eqp As Integer
    Private m_oEqp As Equipment ' The Selected Equipment we are dealing with
    Private m_oTransactionMgr As ClientTransactionManager
    Private m_oUOM As UOMManager
    ' To make ourself Modal if no Eqp was selected before starting the command.
    Private m_ModalState As Boolean

    ' Our Implementation of Logic on Command Start.

    Public Overrides Sub OnStart(ByVal commandID As Integer, ByVal argument As Object)

        m_oTransactionMgr = ClientServiceProvider.TransactionMgr
        m_ModalState = False
        StepCount = 2

        '1st step 
        Steps(0).Prompt = "Select two or three Equipments. Press UpArrow Key when ready to specify positions."
        Steps(0).LocateBehavior = StepDefinition.LocateBehaviors.Auto
        Steps(0).StepFilter = New StepFilter ' Required in v2009.1
        Steps(0).StepFilter.AddInterface("IJSmartEquipment")
        Steps(0).UseStepHiliterForSelection = True ' We use Step Hiliter
        Steps(0).HiliteSelected = True
        Steps(0).SelectHiliterColor = ColorConstants.RGBYellow
        Steps(0).SelectHiliterWeight = 1
        Steps(0).MaximumSelectable = 3


        Steps(1).Prompt = "Select New Equipment to Move to calculated position."
        Steps(1).LocateBehavior = StepDefinition.LocateBehaviors.Auto
        Steps(1).StepFilter = New StepFilter ' Required in v2009.1
        Steps(1).StepFilter.AddInterface("IJSmartEquipment")
        Steps(1).UseStepHiliterForSelection = True ' We use Step Hiliter
        Steps(1).HiliteSelected = True
        Steps(1).SelectHiliterColor = ColorConstants.RGBRed
        Steps(1).SelectHiliterWeight = 1
        Steps(1).MaximumSelectable = 1


        ClientServiceProvider.TransactionMgr.Abort()
        CurrentStepIndex = 0 'set active step to 0
        Enabled = True

        'create and display form
        m_frmLPRodsModeling = New frmLPRodsModeling
        m_frmLPRodsModeling.Show() 'display as Modeless



    End Sub


    Protected Overrides Sub OnKeyDown(ByVal e As Ingr.SP3D.Common.Client.BaseGraphicCommand.KeyEventArgs)
        If e.KeyValue = System.Windows.Forms.Keys.Up Then 'Key For setting "Position" Step
            Dim nSelected = Steps(0).SelectedBusinessObjects.Count
            If (CurrentStepIndex = 0) And Not (nSelected = 0) Then
                Step1_Finisched()
            End If
        End If
    End Sub

    Private Sub Step1_Finisched()
        'check if we have exactly one equipment in the SelectSet
        n_Eqp = Steps(0).SelectedBusinessObjects.Count

        If n_Eqp <> 2 And n_Eqp <> 3 Then
            MsgBox("Select Two or Three Equipments")
            Exit Sub
        End If

        'get selected equipment
        m_oEqp1 = CType(Steps(0).SelectedBusinessObjects.Item(0), Equipment)
        m_oEqp2 = CType(Steps(0).SelectedBusinessObjects.Item(1), Equipment)
        If n_Eqp = 3 Then
            m_oEqp3 = CType(Steps(0).SelectedBusinessObjects.Item(2), Equipment)
        End If

        'initialize form Properties
        m_frmLPRodsModeling.EqpName1 = m_oEqp1.Name
        m_frmLPRodsModeling.EqpPosition1 = m_oEqp1.Origin
        m_frmLPRodsModeling.EqpName2 = m_oEqp2.Name
        m_frmLPRodsModeling.EqpPosition2 = m_oEqp2.Origin
        If n_Eqp = 3 Then
            m_frmLPRodsModeling.EqpName3 = m_oEqp3.Name
            m_frmLPRodsModeling.EqpPosition3 = m_oEqp3.Origin
        Else
            Dim NPos3 As Position = New Position(0, 0, 0)
            m_frmLPRodsModeling.EqpPosition3 = NPos3
        End If
        CalculateNewPosition()

        CurrentStepIndex = 1 ' goto "Select Position" step

    End Sub

    Protected Overrides Sub OnStepSelectionChanged(stepDef As StepDefinition, isSelected As Boolean, businessObj As BusinessObject, position As Position)
        If CurrentStepIndex = 1 Then
            m_oEqp = CType(businessObj, Equipment)
            m_frmLPRodsModeling.EqpName = m_oEqp.Name
        End If
        MyBase.OnStepSelectionChanged(stepDef, isSelected, businessObj, position)
    End Sub


    'Our implementation of Modal Property - Defaults False, but if no Eqp was 
    ' selected before Starting, we want to portray as 'Modal=True', so that
    ' System will Stop us right after OnStart exits.
    Public Overrides ReadOnly Property Modal() As Boolean
        Get
            Return m_ModalState
        End Get
    End Property



    ' Our Implementation of Logic on Command Stop.
    Public Overrides Sub OnStop()
        'Take care of the situation that form is not initiated.
        'See OnStart implementation, in case one eqp wasnt selected
        'OnStart() shows a message and exits without shows form.
        If (m_frmLPRodsModeling IsNot Nothing) Then m_frmLPRodsModeling.Close()
        m_oTransactionMgr.Abort() ' dont leave a Transaction

    End Sub

    ' Our Implementation of Logic on Command Suspend. ' Hide Our form
    Public Overrides Sub OnSuspend()
        m_frmLPRodsModeling.Hide()
    End Sub

    ' Our Implementation of Logic on Command Resume. Show Our form
    Public Overrides Sub OnResume()
        m_frmLPRodsModeling.Show()
    End Sub


    ' Implementation of our Form Events - OnFinish - Commit and Stop Cmd.
    Private Sub m_frmLPRodsModeling_OnFinish() Handles m_frmLPRodsModeling.OnFinish
        m_oTransactionMgr.Commit("Custom Equipment Modification")
        Steps(0).SelectionInfos.Clear()
        Steps(1).SelectionInfos.Clear()
        CurrentStepIndex = 0
        UpdateStep(CurrentStepIndex)

    End Sub

    Protected Overrides Sub OnMouseDown(ByVal view As Ingr.SP3D.Common.Client.Services.GraphicView, ByVal e As Ingr.SP3D.Common.Client.Services.GraphicViewManager.GraphicViewEventArgs, ByVal position As Ingr.SP3D.Common.Middle.Position)
        If (e.Button = MouseButtons.Right) Then
            StopCommand()
        Else
            If (e.Button = MouseButtons.Left) Then
                MyBase.OnMouseDown(view, e, position)

            End If
        End If
    End Sub

    Private Sub MoveToNewPosition(ByVal oPos As Ingr.SP3D.Common.Middle.Position) Handles m_frmLPRodsModeling.MoveToNewPosition
        If m_oEqp.Name <> m_frmLPRodsModeling.EqpName Then
            m_oEqp.SetUserDefinedName(m_frmLPRodsModeling.EqpName)
        End If

        If m_frmLPRodsModeling.RadioButton1.Checked Then
            m_oEqp.Origin = m_frmLPRodsModeling.CalcPosition1
            m_oTransactionMgr.Compute()
        Else
            m_oEqp.Origin = m_frmLPRodsModeling.CalcPosition2
            m_oTransactionMgr.Compute()
        End If
    End Sub

    Private Sub OnClose() Handles m_frmLPRodsModeling.OnClose
        StopCommand()
    End Sub

    Private Sub CalculateNewPosition()
        Dim sb As System.Text.StringBuilder = New Text.StringBuilder()
        sb.Append(m_frmLPRodsModeling.EqpPosition1.X.ToString()).Append(" ")
        sb.Append(m_frmLPRodsModeling.EqpPosition1.Y.ToString()).Append(" ")
        sb.Append(m_frmLPRodsModeling.EqpPosition1.Z.ToString()).Append(" ")
        sb.Append(m_frmLPRodsModeling.EqpPosition2.X.ToString()).Append(" ")
        sb.Append(m_frmLPRodsModeling.EqpPosition2.Y.ToString()).Append(" ")
        sb.Append(m_frmLPRodsModeling.EqpPosition2.Z.ToString()).Append(" ")
        If n_Eqp = 3 Then
            sb.Append(m_frmLPRodsModeling.EqpPosition3.X.ToString()).Append(" ")
            sb.Append(m_frmLPRodsModeling.EqpPosition3.Y.ToString()).Append(" ")
            sb.Append(m_frmLPRodsModeling.EqpPosition3.Z.ToString()).Append(" ")
        End If

        Dim oProcess As New Process()
        Dim oStartInfo As New ProcessStartInfo(GetSymbolShare() + "\CustomCommands\sphereintersections.exe", sb.ToString())


        oStartInfo.UseShellExecute = False
        oStartInfo.RedirectStandardOutput = True
        oProcess.StartInfo = oStartInfo
        oProcess.Start()

        Dim sOutput As String
        Using oStreamReader As System.IO.StreamReader = oProcess.StandardOutput
            sOutput = oStreamReader.ReadToEnd()
        End Using
        Dim strArray = sOutput.Split(New String() {" "}, StringSplitOptions.None)
        Dim x1 As Double = GetDouble(strArray(0))
        Dim y1 As Double = GetDouble(strArray(1))
        Dim z1 As Double = GetDouble(strArray(2))
        Dim x2 As Double = GetDouble(strArray(3))
        Dim y2 As Double = GetDouble(strArray(4))
        Dim z2 As Double = GetDouble(strArray(5))
        Dim NPos1 As Position = New Position(x1, y1, z1)
        Dim NPos2 As Position = New Position(x2, y2, z2)
        m_frmLPRodsModeling.CalcPosition1 = NPos1
        m_frmLPRodsModeling.CalcPosition2 = NPos2


    End Sub

    Public Shared Function GetDouble(ByVal doublestring As String) As Double
        Dim retval As Double
        Dim sep As String = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator

        Double.TryParse(Replace(Replace(doublestring, ".", sep), ",", sep), retval)
        Return retval
    End Function

    Public Function GetSymbolShare() As String
        Dim oModel = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel
        Return oModel.PlantCatalog.SymbolShare.ToString()
    End Function
End Class


