Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports System.Windows.Forms


Public Class StrNameRuleAssignment
    Inherits BaseModalCommand
    Private WithEvents m_oTrnxMgr As ClientTransactionManager
    Private WithEvents m_MyForm As frmNameRuleAssignment
    'Private m_ProgressDisplay As ProgressDisplayService


    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)

        m_oTrnxMgr = ClientServiceProvider.TransactionMgr
        'm_ProgressDisplay = ClientServiceProvider.ProgressDisplayService


        Dim oSelectSet As SelectSet = ClientServiceProvider.SelectSet
        If (oSelectSet.SelectedObjects.Count = 0) Then
            MsgBox("No Objects Selected")
        Else
            'm_ProgressDisplay.EnableAssocProgressDisplay = True


            m_MyForm = New frmNameRuleAssignment
            Application.EnableVisualStyles()
            m_MyForm.Label1.Text = "Updating " + Str(oSelectSet.SelectedObjects.Count) + " names..."
            m_MyForm.ProgressBar1.Minimum = 0
            m_MyForm.ProgressBar1.Maximum = oSelectSet.SelectedObjects.Count
            m_MyForm.Show()
            Dim stepNumber As Integer = 1

            Dim selectedBOCs As BOCollection = oSelectSet.SelectedObjects


            For Each oObj As BusinessObject In selectedBOCs
                SetNameRule(oObj, stepNumber)
                stepNumber += 1
                'm_MyForm.ProgressBar1.PerformStep()
            Next
            m_oTrnxMgr.Compute()
            m_oTrnxMgr.Commit("Assign POE NameRule")
            oSelectSet.Clear()

            MsgBox("NameRules updated")

        End If
    End Sub

    Private Sub m_Commited() Handles m_oTrnxMgr.PostCommit
        m_MyForm.ProgressBar1.PerformStep()

    End Sub


    Public Overrides Sub OnStop()
        m_oTrnxMgr.Abort()
        m_MyForm.Close()

    End Sub
    Public Sub SetNameRule(oObj As BusinessObject, stepNumber As Integer)
        'MsgBox(oObj.ClassInfo.Name)
        Dim mPart = Nothing
        Dim s3dClass As String
        Dim ruleName As String

        s3dClass = oObj.ClassInfo.Name

        If s3dClass = "CSPSMemberPartPrismatic" Or s3dClass = "CSPSMemberPartCurve" Then
            ruleName = "POEMemberPartTypeNameRule"
            mPart = DirectCast(oObj, MemberPart)
        ElseIf s3dClass = "CSPSMemberSystemLinear" Or s3dClass = "CSPSMemberSystemCurve" Then
            ruleName = "MemberSystemTypeNameRule"
            mPart = DirectCast(oObj, MemberSystem)
        ElseIf s3dClass = "CSPSSlabEntity" Then
            ruleName = "DefaultNameRule"
            mPart = DirectCast(oObj, Slab)
        Else
            m_MyForm.ProgressBar1.PerformStep()
            Exit Sub
        End If
        'CSPSSlabEntity
        'CSPSFooting
        'CSPSFootingComponent
        Try
            Dim oPOENameRule = GetNameRuleObj(s3dClass, ruleName)

            If (oPOENameRule IsNot Nothing) Then
                mPart.ActiveNameRule = oPOENameRule

                If (stepNumber Mod 10) = 0 Then
                    m_oTrnxMgr.Compute()
                    m_oTrnxMgr.Commit("Assign POE NameRule")
                End If

            End If
        Catch ex As Exception
            DebugLogMessage(ex.Message)
        End Try



    End Sub

    Private Function GetNameRuleObj(s3dClass As String, ruleName As String) As BusinessObject
        'Lets find available Name Rules for class and pick 'POENameRule' if it exists.
        ' For this, we use CatalogHelper to find out available Name rules on an object Class type.
        ' See GenericNameRules sheet. 
        Dim oCatalogHelper As New CatalogBaseHelper
        Dim oPOENameRule As BusinessObject = Nothing

        For Each oNameRule As BusinessObject In oCatalogHelper.GetNamingRules(s3dClass)
            'If we got our name rule of interest, then we exit this loop
            Dim oPVS As PropertyValueString = oNameRule.GetPropertyValue("IJDNameRuleHolder", "Name")
            If oPVS.PropValue = ruleName Then
                oPOENameRule = oNameRule
                Exit For
            End If
        Next oNameRule
        Return oPOENameRule
    End Function



End Class
