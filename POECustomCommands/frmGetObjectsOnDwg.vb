Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Drawings.Middle
Imports Ingr.SP3D.Drawings.Middle.Services
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle.Services.Hidden
Imports System.Data
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text


Public Class frmGetObjectsOnDwg

    Private csv_entries As List(Of CSV_Entry)
    Public Class CSV_Entry
        Public Oid As String = ""
        Public ClassInfo As String = ""
        Public ItemTag As String = ""
        Public NormTag As String = ""
        Public SAP_GUID As String = ""

    End Class
    Private Sub frmGetObjectsOnDwg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        csv_entries = New List(Of CSV_Entry)
        FillTreeView()
    End Sub

    Private Sub FillTreeView()
        Dim plantName As String = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.Name
        Dim bookParent As IDeliverableParent = EntityHelper.GetBook()
        Dim parentNode As TreeNode = New TreeNode(plantName)
        'parentNode.ContextMenuStrip = ContextMenuStripFolder;

        treeViewDeliverableRoot.BeginUpdate()
        treeViewDeliverableRoot.Nodes.Clear()
        treeViewDeliverableRoot.Nodes.Add(parentNode)

        'Get the childrens collection And verify that newly created child exist
        Dim childrenBOs As IReadOnlyCollection(Of IDeliverableChild) = bookParent.DeliverableComponents

        If childrenBOs.Count > 0 Then
            For Each oChild As BusinessObject In childrenBOs
                Dim treeComponent As TreeNode = New TreeNode(oChild.ToString())
                treeComponent.Name = oChild.ToString()
                parentNode.Nodes.Add(treeComponent)
                AddFirstChild(treeComponent)
            Next
        End If

        treeViewDeliverableRoot.EndUpdate()
        treeViewDeliverableRoot.Sort()

        If Not IsNothing(treeViewDeliverableRoot.TopNode) Then
            treeViewDeliverableRoot.TopNode.Expand()
        End If
    End Sub

    'Adds children to Treeview
    Sub AddChildren(Node As TreeNode, oCol As IReadOnlyCollection(Of IDeliverableChild))

        For Each oBo As BusinessObject In oCol
            Dim IsNodeExists As Boolean = False
            Dim oTempNode As TreeNode = Nothing

            For Each ochild As TreeNode In Node.Nodes
                If oBo.ToString() = ochild.Text Then
                    oTempNode = ochild
                    'If (oBo.SupportsInterface("IJDwgImportedFolder")) Then
                    'oTempNode.ContextMenuStrip = contextMenuStripImportedFolder
                    'ElseIf (oBo.SupportsInterface("IJDwgImportedChildFolder")) Then
                    'oTempNode.ContextMenuStrip = contextMenuStripChildImportedFolder;
                    IsNodeExists = True
                    Exit For
                End If
            Next

            If IsNodeExists = False Then
                Dim key As String = Node.Name + "\\" + oBo.ToString()
                oTempNode = Node.Nodes.Add(key, oBo.ToString())

                'If oBo.SupportsInterface("IJDwgImportedFolder") Then
                'oTempNode.ContextMenuStrip = contextMenuStripImportedFolder;
                'ElseIf (oBo.SupportsInterface("IJDwgImportedChildFolder")) Then
                '   oTempNode.ContextMenuStrip = contextMenuStripChildImportedFolder;
            End If
            AddFirstChild(oTempNode)
        Next

    End Sub

    'Adds first child to TreeNode
    Sub AddFirstChild(Node As TreeNode)

        Dim strPath As String = Node.Name
        Dim oDocParent As IDeliverableParent = TryCast(GetObjectByPath(strPath), IDeliverableParent)
        If Not IsNothing(oDocParent) Then

            Dim oDocCol As IReadOnlyCollection(Of IDeliverableChild) = oDocParent.Deliverables
            Dim oComponentCol As IReadOnlyCollection(Of IDeliverableChild) = oDocParent.DeliverableComponents
            Dim oBO As BusinessObject

            If (oComponentCol.Count > 0) Then
                oBO = TryCast(oComponentCol.First(), BusinessObject)
            ElseIf (oDocCol.Count > 0) Then
                oBO = TryCast(oDocCol.First(), BusinessObject)
            Else
                Return
            End If

            For Each ochild As TreeNode In Node.Nodes
                If oBO.ToString() = ochild.Text Then
                    Return
                End If
            Next
            If Not IsNothing(oBO) Then
                Dim key As String = Node.Name + "\\" + oBO.ToString()
                Node.Nodes.Add(key, oBO.ToString())
            End If

        End If
    End Sub

    Private Function GetObjectByPath(path As String) As BusinessObject
        Dim newPath As String = path
        Dim plantName As String = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.Name

        If path.Contains(plantName) Then
            newPath = path.Remove(0, plantName.Count())
        End If
        Return EntityHelper.GetObjectByPath(newPath)
    End Function

    Private Sub treeViewDeliverableRoot_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles treeViewDeliverableRoot.AfterExpand

        Dim strNodeName As String = e.Node.Text
        Dim strPath As String = e.Node.Name
        Dim eNode As TreeNode = TryCast(e.Node, TreeNode)
        treeViewDeliverableRoot.SelectedNode = eNode
        Dim oDocParent As IDeliverableParent = GetObjectByPath(strPath)
        Dim oDocCol As IReadOnlyCollection(Of IDeliverableChild) = oDocParent.Deliverables
        Dim oComponentCol As IReadOnlyCollection(Of IDeliverableChild) = oDocParent.DeliverableComponents
        AddChildren(eNode, oComponentCol)
        AddChildren(eNode, oDocCol)
        treeViewDeliverableRoot.EndUpdate()
        treeViewDeliverableRoot.SelectedNode = Nothing

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Cursor.Current = Cursors.WaitCursor

        Dim oSelectedNode As TreeNode = treeViewDeliverableRoot.SelectedNode
        Dim strNodePath As String = oSelectedNode.FullPath
        Dim oEntity As DeliverableBase = Nothing
        csv_entries.Clear()

        Try
            oEntity = TryCast(GetObjectByPath(strNodePath), DeliverableBase)
        Catch
            MsgBox("Selected item is not DelivarableObject")
            Exit Sub
        End Try

        If oEntity.ClassInfo.Name <> "CDrawingSheet" Then
            MsgBox("Selected item is not Drawing Sheet")
            Exit Sub
        End If


        Dim DataT As DataTable = GetDataTable(oEntity)

        Dim oid As String
        Dim csv_entry = Nothing
        For Each row As DataRow In DataT.Rows
            oid = row.Item("oBO").ToString()
            oid = "{" + oid + "}"
            csv_entry = ProcessOneOID(oid)
            If Not csv_entry Is Nothing Then csv_entries.Add(csv_entry)
        Next

        Dim FilePath As String = GetCSVFilePath(oEntity.ToString())
        Create_CSV(csv_entries, FilePath)
        Cursor.Current = Cursors.Default
        MsgBox("CSV file " + FilePath + " created.")

    End Sub


    Private Function GetDataTable(oEntity As DeliverableBase)
        Dim sql_t As String = "Select J.oid, J1.ItemName As DrawingName,J5.oid as oBO, J2.OidDestination As ViewOid, J3.ItemName As ViewName, j6.ItemName As ObjectName from JDDWGSheet2 J
                                Join JNamedItem J1 on J1.Oid=J.Oid
                                Join XSheetHasViews J2 on J2.OidDestination=J1.Oid
                                Join JNamedItem J3 on J3.Oid=J2.OidOrigin
                                Join XDrawingMap j4 on j4.oidOrigin = j3.Oid
                                Join JDObject j5 on j5.Oid = j4.oidDestination
                                Join JNamedItem j6 on j6.Oid = j5.Oid
                                where J1.ItemName = '" + oEntity.ToString() +
                                "' order by DrawingName, viewname"

        Dim oConn As SP3DConnection = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantReport
        Dim oConn2 As New SQLDBConnection
        oConn2.OpenSQLConnection(oConn.Server, oConn.Name)
        Dim DataT As DataTable = oConn2.ExecuteSelectQuery(sql_t)
        Return DataT
    End Function

    Private Function ProcessOneOID(oid As String) As CSV_Entry

        Dim oConn As SP3DConnection = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel
        Dim oBo As BusinessObject = Nothing

        Try
            oBo = oConn.WrapSP3DBO(oConn.GetBOMonikerFromDbIdentifier(oid))
        Catch
            Return Nothing
        End Try

        If oBo = Nothing Then
            Return Nothing
        End If

        Dim csv_entry As CSV_Entry = New CSV_Entry
        csv_entry.Oid = oid
        Dim ItemTag As String = oBo.ToString()

        csv_entry.ItemTag = CleanItemTag(ItemTag)

        csv_entry.ClassInfo = oBo.ClassInfo.DisplayName

        If oBo.SupportsInterface("IUPipePipeOMV_Attribute") = True Then
            csv_entry.SAP_GUID = oBo.GetPropertyValue("IUPipePipeOMV_Attribute", "SAP_GUID").ToString()
            csv_entry.NormTag = oBo.GetPropertyValue("IUPipePipeOMV_Attribute", "NormTag").ToString()

        ElseIf oBo.SupportsInterface("IUPipePartOMV_Attribute") = True Then
            csv_entry.SAP_GUID = oBo.GetPropertyValue("IUPipePartOMV_Attribute", "SAP_GUID").ToString()
            csv_entry.NormTag = oBo.GetPropertyValue("IUPipePartOMV_Attribute", "NormTag").ToString()
        ElseIf oBo.SupportsInterface("IUPipSpecOMV_Attribute") = True Then
            csv_entry.SAP_GUID = oBo.GetPropertyValue("IUPipSpecOMV_Attribute", "SAP_GUID").ToString()
            csv_entry.NormTag = oBo.GetPropertyValue("IUPipSpecOMV_Attribute", "NormTag").ToString()
        ElseIf oBo.SupportsInterface("IUPipeInstrOMV_Attribute") = True Then
            csv_entry.SAP_GUID = oBo.GetPropertyValue("IUPipeInstrOMV_Attribute", "SAP_GUID").ToString()
            csv_entry.NormTag = oBo.GetPropertyValue("IUPipeInstrOMV_Attribute", "NormTag").ToString()
        ElseIf oBo.SupportsInterface("IUAEQP_OMV_Attribute") = True Then
            csv_entry.SAP_GUID = oBo.GetPropertyValue("IUAEQP_OMV_Attribute", "SAP_GUID").ToString()
            csv_entry.NormTag = oBo.GetPropertyValue("IUAEQP_OMV_Attribute", "NormTag").ToString()
        ElseIf oBo.SupportsInterface("IUPipSuppOMV_Attribute") = True Then
            csv_entry.SAP_GUID = oBo.GetPropertyValue("IUPipSuppOMV_Attribute", "SAP_GUID").ToString()
            csv_entry.NormTag = oBo.GetPropertyValue("IUPipSuppOMV_Attribute", "NormTag").ToString()
        End If

        Return csv_entry
    End Function

    Public Function CleanItemTag(ItemTag As String)
        Return ItemTag.Replace("?", "-")
    End Function

    Public Function GetCSVFilePath(DrawingNumber As String, Optional suffix As String = "") As String
        Dim FileName As String = DrawingNumber + suffix + ".csv"
        Dim FilePath As String = Path.Combine({Path.GetTempPath(), FileName})
        Return FilePath
    End Function

    Public Sub Create_CSV(csv_entries As List(Of CSV_Entry), FilePath As String)

        Dim sb As StringBuilder = New StringBuilder()
        Dim header_columns() As String = {"Oid", "ClassInfo", "ItemTag", "NormTag", "SAP_GUID"}
        sb.AppendLine(String.Join(";", header_columns))

        Dim line As String = ""
        For Each e As CSV_Entry In csv_entries
            line = ""
            For Each col_name As String In header_columns
                Try
                    line = line + e.GetType().GetField(col_name).GetValue(e) + ";"
                Catch
                    line = line + ";"
                End Try
            Next
            sb.AppendLine(line)
        Next

        File.WriteAllText(FilePath, sb.ToString())
        sb.Clear()

    End Sub







End Class

