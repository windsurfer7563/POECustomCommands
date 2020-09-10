Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Client
Imports System.Collections.ObjectModel



Public Class frmDefineWorksace
    Private Sub frmDefineWorksace_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim m_oRoot As HierarchiesRoot
        Dim oFilter As New Filter()
        'Dim oFilter2 As New Filter()
        Dim oFilter2 As New Filter("TestLabFilter2", "My Filters")
        Dim oWorkingcoll As BOCollection
        'Dim oSystems As Collection(Of BusinessObject)

        oWorkingcoll = ClientServiceProvider.WorkingSet.GetObjectsByFilter(oFilter, ClientServiceProvider.WorkingSet.ActiveConnection)


        'Dim oObjs = New ReadOnlyCollection(Of BusinessObject)(oWorkingcoll)

        'oFilter2.Definition.AddHierarchy(HierarchyTypes.System, oObjs, True)

        'ClientServiceProvider.TransactionMgr.Commit("Commit Filter")
        'm_oRoot = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.RootSystem
        Dim oModel As Model = TryCast(ClientServiceProvider.WorkingSet.ActiveConnection, Model)
        m_oRoot = oModel.RootSystem


        Dim ws = ClientServiceProvider.WorkingSet

        Dim oPreferences As Preferences

        ' Get preferences from ClientServiceProvider.
        oPreferences = ClientServiceProvider.Preferences
        Dim pk As SessionPreferenceKeys

        'Dim oValueManager As ValueManager
        ' Get ValueManager from ClientServiceProvider.
        'oValueManager = ClientServiceProvider.ValueMgr

        'For i = 1 To oValueManager.Count
        ' MsgBox(oValueManager.GetKeyName(i))
        'Next


        Dim oSysChilds As ReadOnlyCollection(Of ISystemChild)
        oSysChilds = m_oRoot.SystemChildren


        Dim oBOs = New Collection(Of BusinessObject)

        For Each bo In oSysChilds
            oBOs.Add(bo)
        Next

        Dim oBOReadOnly = New ReadOnlyCollection(Of BusinessObject)(oBOs)


        oFilter2.Definition.AddHierarchy(HierarchyTypes.System, oBOReadOnly, False)

        'ClientServiceProvider.TransactionMgr.Commit("Commit HierarchyFilter")

        Dim oWorkSpaceDef As WorkspaceDefinitionCriteria

        oWorkSpaceDef = New WorkspaceDefinitionCriteria(oFilter2)

        'Dim oFS = New FilterSelector()
        'oFS.ShowDialog()
        'MsgBox(oFS.SelectedModelFilters(0).ToString())

        Ingr.SP3D.Common.Client.Services.ClientServiceProvider.CommandMgr.StartCommand("GSCADWrkSpDefCmd.WrkSpDefCmd", CommandManager.CommandPriority.Normal, vbNull)

        'ClientServiceProvider.WorkingSet.Refresh(WorkingSet.UpdateQuery.IncrementalUpdate)



        Call OutputFilterMembers(oSysChilds, oWorkingcoll)
    End Sub

    Private Sub OutputFilterMembers(oObjs As ReadOnlyCollection(Of ISystemChild), oWorkingCol As BOCollection)

        For Each oObj As BusinessObject In oObjs
            If oWorkingCol.Contains(oObj) = True Then
                lvObjs.Items.Add("Inner " & oObj.ClassInfo.Name & " - " & oObj.ToString())
            End If
            lvObjs.Items.Add(oObj.ClassInfo.Name & " - " & oObj.ToString())
            lvObjs.Select()

        Next

    End Sub

    Private Sub lvObjs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvObjs.SelectedIndexChanged

    End Sub
End Class