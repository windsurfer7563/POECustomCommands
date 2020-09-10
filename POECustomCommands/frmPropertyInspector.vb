
Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports System.Windows.Forms


Public Class frmPropertyInspector

    Private WithEvents m_oEVSelectSet As SelectSet
    Private currentBOClass As String
    Private current_bo As BusinessObject


    Private Sub frmPropertyInspector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        m_oEVSelectSet = ClientServiceProvider.SelectSet
        lvProperties.Width = Me.Width - 10
        lvProperties.Height = Me.Height - 40
        lvProperties.ShowGroups = True
        current_bo = Nothing

    End Sub


    Private Sub frmPropertyInspector_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        m_oEVSelectSet = Nothing
    End Sub

    Private Sub m_oEVSelectSet_SelectionChanged(eventSpecifics As MultiCasterEventArgs) Handles m_oEVSelectSet.SelectionChanged
        If m_oEVSelectSet.SelectedObjects.Count = 0 Then Exit Sub
        If m_oEVSelectSet.SelectedObjects.Item(0) = current_bo Then Exit Sub
        Call ProcessSelection(m_oEVSelectSet.SelectedObjects.Item(0))

    End Sub

    Private Sub ProcessSelection(bo As BusinessObject)
        current_bo = bo
        currentBOClass = current_bo.ClassInfo.DisplayName

        lvProperties.Items.Clear()

        GetGeneralProperties(current_bo)

        Select Case currentBOClass
            Case "Equipment"
                Call GetEqpProperties(current_bo)
            Case "Shape"
                Call GetShapeProperties(current_bo)
            Case "Pipe Part"
                Call GetPipeProperties(current_bo)
            Case "Pipe Component"
                Call GetPipePartProperties(current_bo)
            Case "Pipe Nozzle"
                Call GetPipeNozzleProperties(current_bo)
            Case "Pipe Run"
                Call GetPipeRunData(current_bo)
                Dim oPipeLine = getPipelineFromPart(current_bo)
                Call GetPipeLineData(oPipeLine)
            Case "Pipeline System"
                Call GetPipeLineProperties(current_bo)
            Case "Pipe Specialty Item"
                Call GetPipeSpecialtyProperties(current_bo)
            Case "Pipe Instrument"
                Call GetPipeSpecialtyProperties(current_bo)
            Case "Distribution Connection"
                Call GetConnectionData(current_bo, 1)
            Case "Member Part Prismatic"
                Call GetMemberpartProperies(current_bo)
            Case "Member System Prismatic"
                Call GetMemberSystemProperties(current_bo)
            Case "Slab"
                Call GetSlabProperties(current_bo)

        End Select


        If currentBOClass.Contains("Feature") Then
            Call GetPipeFeatureProperties(current_bo)

        End If

        If CheckBoxLongFormat.Checked = True Then
            Call GetNotes(current_bo)
        End If


    End Sub
    Private Sub GetGeneralProperties(businessObject As BusinessObject)
        Dim group As New ListViewGroup("General", HorizontalAlignment.Left)
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Class", currentBOClass, group))
        lvProperties.Items.Add(CreateNewLI("Name", businessObject.ToString(), group))
        lvProperties.Items.Add(CreateNewLI("ApprovalStatus", businessObject.ApprovalStatus.ToString(), group))

        Call GetConstructionInfoData(businessObject, group)

        If CheckBoxLongFormat.Checked = True Then
            lvProperties.Items.Add(CreateNewLI("UserLastModified", businessObject.UserLastModified.ToString(), group))
            lvProperties.Items.Add(CreateNewLI("DateLastModified", businessObject.DateLastModified.ToString(), group))
            lvProperties.Items.Add(CreateNewLI("UserCreated", businessObject.UserCreated.ToString(), group))
        End If

        If currentBOClass.Contains("Feature") Then
            businessObject = getPartFromFeature(businessObject)
        End If

        If businessObject.SupportsInterface("IJMtoInfo") = True Then
            lvProperties.Items.Add(CreateNewLI("ReportingRequirements", businessObject.GetPropertyValue("IJMTOInfo", "ReportingRequirements").ToString(), group))
        End If

    End Sub

    Private Sub GetConstructionInfoData(businessObject As BusinessObject, group As ListViewGroup)

        If currentBOClass.Contains("Feature") Then
            businessObject = getPartFromFeature(businessObject)
        End If

        If businessObject.SupportsInterface("IJConstructionInfo") = True Then
            lvProperties.Items.Add(CreateNewLI("ConstructionRequirement", businessObject.GetPropertyValue("IJConstructionInfo", "ConstructionRequirement").ToString(), group))
        End If

        If businessObject.SupportsInterface("IJConstructionInfo") = True Then
            lvProperties.Items.Add(CreateNewLI("ConstructionType", businessObject.GetPropertyValue("IJConstructionInfo", "ConstructionType").ToString(), group))
        End If

    End Sub
    Private Sub GetEqpProperties(businessObject As BusinessObject)
        Dim pos As Position
        Dim c1, c2, c3

        Dim equipment_group As New ListViewGroup("Equipment", HorizontalAlignment.Left)
        lvProperties.Groups.Add(equipment_group)

        lvProperties.Items.Add(CreateNewLI("Name", businessObject.ToString(), equipment_group))
        pos = GetEqpPosition(businessObject)
        lvProperties.Items.Add(CreateNewLI("X", FormatDistanceValue(pos.X), equipment_group))
        lvProperties.Items.Add(CreateNewLI("Y", FormatDistanceValue(pos.Y), equipment_group))
        lvProperties.Items.Add(CreateNewLI("Z", FormatDistanceValue(pos.Z), equipment_group))

        c1 = businessObject.GetRelationship("SOtoSI_R", "toSI_ORIG").TargetObjects.Item(0).GetPropertyValue("IJEquipmentPart", "ProcessEqTypes0")
        c2 = businessObject.GetRelationship("SOtoSI_R", "toSI_ORIG").TargetObjects.Item(0).GetPropertyValue("IJEquipmentPart", "ProcessEqTypes0")
        c3 = businessObject.GetRelationship("SOtoSI_R", "toSI_ORIG").TargetObjects.Item(0).GetPropertyValue("IJEquipmentPart", "ProcessEqTypes0")

        lvProperties.Items.Add(CreateNewLI("Classification1", c1.ToString(), equipment_group))
        lvProperties.Items.Add(CreateNewLI("Classification2", c2.ToString(), equipment_group))
        lvProperties.Items.Add(CreateNewLI("Classification3", c3.ToString(), equipment_group))

    End Sub
    Private Sub GetPipeNozzleProperties(oNozzle As PipeNozzle)
        Dim oConn As Ingr.SP3D.Route.Middle.Connection

        Dim group As New ListViewGroup("Nozzle", HorizontalAlignment.Left)
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Name", oNozzle.ToString(), group))
        lvProperties.Items.Add(CreateNewLI("Length", FormatDistanceValue(oNozzle.Length), group))

        Call GetPortData(oNozzle)

        If CheckBoxLongFormat.Checked = True Then
            Dim RC As RelationCollection
            RC = oNozzle.GetRelationship("FlowPorts", "DistribConnectionObj")
            If RC.TargetObjects.Count > 0 Then

                oConn = RC.TargetObjects(0)
                Call GetConnectionData(oConn, 1)
            End If
        End If

            Dim oEquipment = oNozzle.SystemParent
        Call GetEqpProperties(oEquipment)

    End Sub

    Private Sub GetShapeProperties(oShape As GenericShape)
        Dim group As New ListViewGroup("Shape")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Name", oShape.Name, group))

        If oShape.SupportsInterface("IJUACylinder") = True Then
            lvProperties.Items.Add(CreateNewLI("A", oShape.GetPropertyValue("IJUACylinder", "A").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("B", oShape.GetPropertyValue("IJUACylinder", "B").ToString(), group))
            GoTo eqp
        End If

        If oShape.SupportsInterface("IJUACone") = True Then
            lvProperties.Items.Add(CreateNewLI("A", oShape.GetPropertyValue("IJUACone", "A").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("B", oShape.GetPropertyValue("IJUACone", "B").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("C", oShape.GetPropertyValue("IJUACone", "c").ToString(), group))
            GoTo eqp
        End If
        If oShape.SupportsInterface("IJUAEccentricCone") = True Then
            lvProperties.Items.Add(CreateNewLI("A", oShape.GetPropertyValue("IJUAEccentricCone", "A").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("B", oShape.GetPropertyValue("IJUAEccentricCone", "B").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("C", oShape.GetPropertyValue("IJUAEccentricCone", "c").ToString(), group))
            GoTo eqp
        End If

        If oShape.SupportsInterface("IJUARectSolid") = True Then
            lvProperties.Items.Add(CreateNewLI("A", oShape.GetPropertyValue("IJUARectSolid", "A").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("B", oShape.GetPropertyValue("IJUARectSolid", "B").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("C", oShape.GetPropertyValue("IJUARectSolid", "C").ToString(), group))
            GoTo eqp
        End If
        If oShape.SupportsInterface("IJUASphere") = True Then
            lvProperties.Items.Add(CreateNewLI("A", oShape.GetPropertyValue("IJUASphere", "A").ToString(), group))
            GoTo eqp
        End If

        If oShape.SupportsInterface("IJUASemiElliptical") = True Then
            lvProperties.Items.Add(CreateNewLI("A", oShape.GetPropertyValue("IJUASemiElliptical", "A").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("B", oShape.GetPropertyValue("IJUASemiElliptical", "B").ToString(), group))
        End If
eqp:

        Dim oEquipment = oShape.SystemParent
        Call GetEqpProperties(oEquipment)
    End Sub

    Private Sub GetPipeProperties(businessObject As BusinessObject)
        Dim p1
        Dim group As New ListViewGroup("Pipe")
        lvProperties.Groups.Add(group)

        lvProperties.Items.Add(CreateNewLI("Length", businessObject.GetPropertyValue("IJRteStockPartOccur", "Length").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("CutLength", businessObject.GetPropertyValue("IJRteStockPartOccur", "CutLength").ToString(), group))
        p1 = businessObject.GetRelationship("madeFrom", "part").TargetObjects.Item(0)
        lvProperties.Items.Add(CreateNewLI("CommodityCode", p1.GetPropertyValue("IJDPipeComponent", "IndustryCommodityCode").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("FirstSizeSchedule", p1.GetPropertyValue("IJDPipeComponent", "FirstSizeSchedule").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("SecondSizeSchedule", p1.GetPropertyValue("IJDPipeComponent", "SecondSizeSchedule").ToString(), group))

        If CheckBoxLongFormat.Checked = True Then
            Call GetDistribPorts(businessObject)
            Call GetBoltingData(businessObject)
        End If
        Dim oPipeRun = GetPipeRun(businessObject)
        Call GetPipeRunData(oPipeRun)
        Dim oPipeLine = getPipelineFromPart(oPipeRun)
        Call GetPipeLineData(oPipeLine)

    End Sub



    Private Sub GetPipePartProperties(businessObject As BusinessObject)
        Dim p1
        Dim group As New ListViewGroup("Pipe Part")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("ShortCode", businessObject.GetPropertyValue("IJRtePartData", "ShortCode").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("OptionCode", businessObject.GetPropertyValue("IJRtePartData", "OptionCode").ToString(), group))
        p1 = businessObject.GetRelationship("madeFrom", "part").TargetObjects.Item(0)
        lvProperties.Items.Add(CreateNewLI("CommodityType", p1.GetPropertyValue("IJDPipeComponent", "CommodityType").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("CommodityCode", p1.GetPropertyValue("IJDPipeComponent", "IndustryCommodityCode").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("FirstSizeSchedule", p1.GetPropertyValue("IJDPipeComponent", "FirstSizeSchedule").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("SecondSizeSchedule", p1.GetPropertyValue("IJDPipeComponent", "SecondSizeSchedule").ToString(), group))
        If CheckBoxLongFormat.Checked = True Then
            Call GetDistribPorts(businessObject)
            Call GetBoltingData(businessObject)
        End If
        Dim oPipeRun = GetPipeRun(businessObject)
        Call GetPipeRunData(oPipeRun)
        Dim oPipeLine = getPipelineFromPart(oPipeRun)
        Call GetPipeLineData(oPipeLine)

    End Sub


    Private Sub GetPipeSpecialtyProperties(businessObject As BusinessObject)
        Dim group As New ListViewGroup("Pipe Specialty")
        lvProperties.Groups.Add(group)
        If businessObject.SupportsInterface("IJFaceToCenter") Then
            lvProperties.Items.Add(CreateNewLI("FacetoCenter", businessObject.GetPropertyValue("IJFaceToCenter", "FacetoCenter").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("Face1toCenter", businessObject.GetPropertyValue("IJFaceToCenter", "Face1toCenter").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("Face2toCenter", businessObject.GetPropertyValue("IJFaceToCenter", "Face2toCenter").ToString(), group))
        End If
        If businessObject.SupportsInterface("IJUAMucFlange") Then
            lvProperties.Items.Add(CreateNewLI("Length", businessObject.GetPropertyValue("IJUAMucFlange", "Length").ToString(), group))
            lvProperties.Items.Add(CreateNewLI("FlangeThk", businessObject.GetPropertyValue("IJUAMucFlange", "FlangeThk").ToString(), group))
        End If

        Dim group2 = New ListViewGroup("Material Control Data")
        lvProperties.Groups.Add(group2)

        Dim oRC As RelationCollection
        oRC = businessObject.GetRelationship("PartOccToMaterialControlData", "MaterialControlData")
        Dim oGenMat = oRC.TargetObjects.Item(0)

        lvProperties.Items.Add(CreateNewLI("ContractorCommodityCode", oGenMat.GetPropertyValue("IJGenericMaterialControlData", "ContractorCommodityCode").ToString(), group2))
        lvProperties.Items.Add(CreateNewLI("Description", oGenMat.GetPropertyValue("IJGenericMaterialControlData", "ShortMaterialDescription").ToString(), group2))

        oRC = oGenMat.GetRelationship("DefinesMaterialControlDataForComponent", "Component")

        Dim oComp = oRC.TargetObjects.Item(0)
        lvProperties.Items.Add(CreateNewLI("PartNumber", oComp.GetPropertyValue("IJDPart", "PartNumber").ToString(), group2))

        If CheckBoxLongFormat.Checked = True Then
            Call GetDistribPorts(businessObject)
            Call GetBoltingData(businessObject)
        End If
        Dim oPipeRun = GetPipeRun(businessObject)
        Call GetPipeRunData(oPipeRun)
        Dim oPipeLine = getPipelineFromPart(oPipeRun)
        Call GetPipeLineData(oPipeLine)

    End Sub

    Private Sub GetDistribPorts(businessObject As BusinessObject)
        Dim oPorts As RelationCollection
        oPorts = businessObject.GetRelationship("DistribPorts", "DistribPort")
        Dim i = 1
        For Each oPort In oPorts.TargetObjects
            Call GetPortData(oPort)
        Next
    End Sub

    Private Sub GetPortData(oPort As BusinessObject)
        Dim group As New ListViewGroup("Port " + oPort.GetPropertyValue("IJDPipePOrt", "PortIndex").ToString())
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("NPD", oPort.GetPropertyValue("IJDPipePOrt", "NPD").ToString() & " " & oPort.GetPropertyValue("IJDPipePOrt", "NpdUnitType").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("Pressure", oPort.GetPropertyValue("IJDPipePOrt", "PressureRating").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("Term.Class", oPort.GetPropertyValue("IJDPipePOrt", "TerminationClass").ToString(), group))
        lvProperties.Items.Add(CreateNewLI("EndPrep", FormatCodeListValue(oPort.GetPropertyValue("IJDPipePOrt", "EndPreparation")), group))
        lvProperties.Items.Add(CreateNewLI("EndPrep Descr.", GetCodelistedValue(oPort.GetPropertyValue("IJDPipePOrt", "EndPreparation"), short_format:=False), group))
        lvProperties.Items.Add(CreateNewLI("EndStd", FormatCodeListValue(oPort.GetPropertyValue("IJDPipePOrt", "EndStandard")), group))
        lvProperties.Items.Add(CreateNewLI("EndPractice", oPort.GetPropertyValue("IJDPipePOrt", "EndPractice").ToString(), group))

    End Sub

    Private Function GetPipeRun(businessObject As BusinessObject)
        Dim oPipeRun As PipeRun
        Dim PiperunCol As RelationCollection = businessObject.GetRelationship("OwnsParts", "Owner")    'IJRtePathGenPart(Owner) -> OwnsParts(Run to Part) -> IJDesignParent(Parts),
        oPipeRun = PiperunCol.TargetObjects.Item(0)
        Return oPipeRun
    End Function
    Private Sub GetPipeRunData(oPipeRun As PipeRun)
        Dim group As New ListViewGroup("PipeRun")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Run Name", oPipeRun.ToString(), group))
        lvProperties.Items.Add(CreateNewLI("Piping Spec", oPipeRun.Specification.SpecificationName, group))
        Dim oPV As PropertyValueCodelist = oPipeRun.GetPropertyValue("IJRteInsulation", "Purpose")
        lvProperties.Items.Add(CreateNewLI("InsulationPurpose", oPV.PropertyInfo.CodeListInfo.GetCodelistItem(oPV.PropValue).ShortDisplayName, group))
        oPV = oPipeRun.GetPropertyValue("IJRteInsulation", "Material")
        lvProperties.Items.Add(CreateNewLI("InsulationMaterial", oPV.PropertyInfo.CodeListInfo.GetCodelistItem(oPV.PropValue).ShortDisplayName, group))

        Dim dInsulationThickness As Double = DirectCast(oPipeRun.GetPropertyValue("IJRteInsulation", "Thickness"), PropertyValueDouble).PropValue
        lvProperties.Items.Add(CreateNewLI("InsulationThickness", CStr(dInsulationThickness * 1000.0), group))

    End Sub


    Private Sub GetPipeLineProperties(oPipeLine As Pipeline)
        Call GetPipeLineData(oPipeLine)
    End Sub

    Private Sub GetPipeLineData(oPipeLine As Pipeline)
        Dim group As New ListViewGroup("Pieline")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Line Name", oPipeLine.ToString(), group))
        lvProperties.Items.Add(CreateNewLI("Fluid System", GetCodelistedValue(oPipeLine.GetPropertyValue("IJPipelineSystem", "FluidSystem")), group))
        lvProperties.Items.Add(CreateNewLI("Fluid Code", GetCodelistedValue(oPipeLine.GetPropertyValue("IJPipelineSystem", "FluidCode")), group))
    End Sub

    Private Sub GetPipeFeatureProperties(current_bo)
        Call GetPipeFeatureData(current_bo)
        Dim oPart = getPartFromFeature(current_bo)

        Select Case oPart.ClassInfo.DisplayName
            Case "Pipe Part"
                Call GetPipeProperties(oPart)
            Case "Pipe Component"
                Call GetPipePartProperties(oPart)
            Case "Pipe Specialty Item"
                Call GetPipeSpecialtyProperties(oPart)
            Case "Pipe Instrument"
                Call GetPipeSpecialtyProperties(oPart)

        End Select

        If CheckBoxLongFormat.Checked = True Then
            Call GetNotes(oPart)
        End If

    End Sub

    Private Sub GetPipeFeatureData(oPipeFeature As BusinessObject)
        Dim group As New ListViewGroup("Feature")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("OptionCode", oPipeFeature.GetPropertyValue("IJRtePipePathFeat", "CommodityOption").ToString(), group))
        Dim Tag = oPipeFeature.GetPropertyValue("IJRtePipePathFeat", "Tag").ToString()
        If Tag <> "" Then lvProperties.Items.Add(CreateNewLI("Tag", oPipeFeature.GetPropertyValue("IJRtePipePathFeat", "Tag").ToString(), group))

        Dim oPV As PropertyValueCodelist = oPipeFeature.GetPropertyValue("IJRteInsulation", "Purpose")
        lvProperties.Items.Add(CreateNewLI("Insulation Purpose", oPV.PropertyInfo.CodeListInfo.GetCodelistItem(oPV.PropValue).ShortDisplayName, group))
        oPV = oPipeFeature.GetPropertyValue("IJRteInsulation", "Material")
        lvProperties.Items.Add(CreateNewLI("Insulation Material", oPV.PropertyInfo.CodeListInfo.GetCodelistItem(oPV.PropValue).ShortDisplayName, group))

        Dim dInsulationThickness As Double = DirectCast(oPipeFeature.GetPropertyValue("IJRteInsulation", "Thickness"), PropertyValueDouble).PropValue
        lvProperties.Items.Add(CreateNewLI("Insulation Thickness", CStr(dInsulationThickness * 1000.0), group))
    End Sub

    Private Function getFeatureFromPart(oBo As BusinessObject) As BusinessObject
        Dim RC As RelationCollection = Nothing
        RC = oBo.GetRelationship("PathGeneratedParts", "DefiningFeature") 'IJRtePathGenPart(DefiningFeature) -> PathGeneratedParts(Feature to Part) -> IJRtePathFeat(PathDefinedPart),
        If RC IsNot Nothing AndAlso RC.TargetObjects.Count > 0 Then
            Return RC.TargetObjects.Item(0)
        Else
            Return Nothing
        End If
    End Function

    Private Function getPartFromFeature(oBo As BusinessObject) As BusinessObject
        Dim RC As RelationCollection = Nothing
        RC = oBo.GetRelationship("PathGeneratedParts", "PathDefinedPart") 'IJRtePathGenPart(DefiningFeature) -> PathGeneratedParts(Feature to Part) -> IJRtePathFeat(PathDefinedPart),
        If RC IsNot Nothing AndAlso RC.TargetObjects.Count > 0 Then
            Return RC.TargetObjects.Item(0)
        Else
            Return Nothing
        End If
    End Function

    Private Function getPipelineFromPart(oBo As BusinessObject) As BusinessObject
        Dim RC As RelationCollection = Nothing
        RC = oBo.GetRelationship("SystemHierarchy", "SystemParent")
        If RC IsNot Nothing AndAlso RC.TargetObjects.Count > 0 Then
            Return RC.TargetObjects.Item(0)
        Else
            Return Nothing
        End If
    End Function

#If comment Then
    Private Sub GetBoltingData(businessObject As BusinessObject)
        Dim oImpliedItems As RelationCollection
        Dim oImpliedMatingParts As RelationCollection
        Dim oBoltPart As BusinessObject
        Dim oGasketPart As BusinessObject

        oImpliedItems = businessObject.GetRelationship("OwnsImpliedItems", "ImpliedItem")

        For Each oItem In oImpliedItems.TargetObjects
            If oItem.SupportsInterface("IJRteBolt") Then
                Dim group As New ListViewGroup("Bolt Set")
                lvProperties.Groups.Add(group)
                oImpliedMatingParts = oItem.GetRelationship("ImpliedMatingParts", "UsedImpliedPart")
                oBoltPart = oImpliedMatingParts.TargetObjects(0)
                lvProperties.Items.Add(CreateNewLI("Bolt CommodityCode", oBoltPart.GetPropertyValue("IJBolt", "IndustryCommodityCode").ToString(), group))
                lvProperties.Items.Add(CreateNewLI("BoltType", oBoltPart.GetPropertyValue("IJBolt", "BoltType").ToString(), group))

                lvProperties.Items.Add(CreateNewLI("CalculatedLength", oItem.GetPropertyValue("IJRteBolt", "CalculatedLength").ToString(), group))
                lvProperties.Items.Add(CreateNewLI("RoundedLength", oItem.GetPropertyValue("IJRteBolt", "RoundedLength").ToString(), group))
                lvProperties.Items.Add(CreateNewLI("Quantity", oItem.GetPropertyValue("IJRteBolt", "Quantity").ToString(), group))
                lvProperties.Items.Add(CreateNewLI("Diameter", oItem.GetPropertyValue("IJRteBolt", "Diameter").ToString(), group))
            End If

            If oItem.SupportsInterface("IJRteGasket") Then
                Dim group As New ListViewGroup("Gasket")
                lvProperties.Groups.Add(group)
                oImpliedMatingParts = oItem.GetRelationship("ImpliedMatingParts", "UsedImpliedPart")
                oGasketPart = oImpliedMatingParts.TargetObjects(0)
                lvProperties.Items.Add(CreateNewLI("Gasket CommodityCode", oGasketPart.GetPropertyValue("IJGasket", "IndustryCommodityCode").ToString(), group))
                lvProperties.Items.Add(CreateNewLI("Gasket Thikness", oGasketPart.GetPropertyValue("IJGasket", "ThicknessFor3DModel").ToString(), group))
            End If

        Next

    End Sub

#End If

    Private Sub GetBoltingData(businessObject As BusinessObject)
        Dim oConnections As RelationCollection
        oConnections = businessObject.GetRelationship("RelConnectionAndPartOcc", "Connection")
        Dim i = 1
        For Each oConn As Ingr.SP3D.Route.Middle.Connection In oConnections.TargetObjects
            Call GetConnectionData(oConn, i)
            i = i + 1
        Next

    End Sub

    Private Sub GetConnectionData(oConn As Ingr.SP3D.Route.Middle.Connection, connIndex As Integer)
        Dim group As New ListViewGroup("Connection " + connIndex.ToString())
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Type", oConn.GetPropertyValue("IJDistribConnection", "ConnectionType").ToString(), group))
        Dim i = 1
        For Each oPart In oConn.ConnectionParts
            lvProperties.Items.Add(CreateNewLI("Part " + i.ToString(), oPart.ToString(), group))
            i = i + 1
        Next

        Call GetConnectionItems(oConn, group)

    End Sub

    Private Sub GetConnectionItems(oConn As Ingr.SP3D.Route.Middle.Connection, group As ListViewGroup)
        Dim oImpliedMatingParts As RelationCollection
        Dim oBoltPart As BusinessObject
        Dim oGasketPart As BusinessObject
        For Each oItem In oConn.ConnectionItems
            Try
                If oItem.SupportsInterface("IJRteBolt") Then
                    oImpliedMatingParts = oItem.GetRelationship("ImpliedMatingParts", "UsedImpliedPart")
                    oBoltPart = oImpliedMatingParts.TargetObjects(0)
                    lvProperties.Items.Add(CreateNewLI("Bolt CommodityCode", oBoltPart.GetPropertyValue("IJBolt", "IndustryCommodityCode").ToString(), group))
                    lvProperties.Items.Add(CreateNewLI("BoltType", oBoltPart.GetPropertyValue("IJBolt", "BoltType").ToString(), group))

                    lvProperties.Items.Add(CreateNewLI("CalculatedLength", oItem.GetPropertyValue("IJRteBolt", "CalculatedLength").ToString(), group))
                    lvProperties.Items.Add(CreateNewLI("RoundedLength", oItem.GetPropertyValue("IJRteBolt", "RoundedLength").ToString(), group))
                    lvProperties.Items.Add(CreateNewLI("Quantity", oItem.GetPropertyValue("IJRteBolt", "Quantity").ToString(), group))
                    lvProperties.Items.Add(CreateNewLI("Diameter", oItem.GetPropertyValue("IJRteBolt", "Diameter").ToString(), group))
                End If

                If oItem.SupportsInterface("IJRteGasket") Then
                    oImpliedMatingParts = oItem.GetRelationship("ImpliedMatingParts", "UsedImpliedPart")
                    oGasketPart = oImpliedMatingParts.TargetObjects(0)
                    lvProperties.Items.Add(CreateNewLI("Gasket CommodityCode", oGasketPart.GetPropertyValue("IJGasket", "IndustryCommodityCode").ToString(), group))
                    lvProperties.Items.Add(CreateNewLI("Gasket Thikness", oGasketPart.GetPropertyValue("IJGasket", "ThicknessFor3DModel").ToString(), group))
                End If

            Catch

            End Try



        Next

    End Sub

    Private Sub GetNotes(oBO As BusinessObject)
        Dim group As New ListViewGroup("Notes")
        lvProperties.Groups.Add(group)
        Dim RC As RelationCollection
        RC = oBO.GetRelationship("ContainsNote", "GeneralNote")
        If RC.TargetObjects.Count > 0 Then
            For Each note As Note In RC.TargetObjects
                lvProperties.Items.Add(CreateNewLI(note.Name, note.Text, group))
            Next
        End If

    End Sub

    Private Function FormatCodeListValue(oPV As PropertyValueCodelist)
        Return oPV.ToString() & " (" & oPV.PropValue.ToString() & ")"
    End Function


    Private Function GetCodelistedValue(oPV As PropertyValueCodelist, Optional short_format As Boolean = True)
        If short_format = True Then
            Return oPV.PropertyInfo.CodeListInfo.GetCodelistItem(oPV.PropValue).ShortDisplayName
        Else
            Return oPV.PropertyInfo.CodeListInfo.GetCodelistItem(oPV.PropValue).DisplayName
        End If

    End Function

    Private Function GetEqpPosition(businessObject As BusinessObject)
        Dim oEquipment As Equipment
        oEquipment = CType(businessObject, Equipment)
        Return oEquipment.Origin
    End Function

    Private Sub GetMemberpartProperies(oMemberPart As MemberPart)
        Call GetMemberPartPrismaticData(oMemberPart)
        Dim oMemberSystem As MemberSystem = GetMemberSystemFromPart(oMemberPart)
        Call GetMemberSystemData(oMemberSystem)

    End Sub


    Private Sub GetMemberPartPrismaticData(oMemberPart As MemberPart)
        Dim group As New ListViewGroup("Member Part")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Name", oMemberPart.Name, group))
        lvProperties.Items.Add(CreateNewLI("Type", GetCodelistedValue(oMemberPart.GetPropertyValue("ISPSMemberType", "Type")), group))
        lvProperties.Items.Add(CreateNewLI("CutLength", FormatDistanceValue(oMemberPart.CutLength), group))
        Dim group2 = New ListViewGroup("CrossSection")
        lvProperties.Groups.Add(group2)
        Dim crossSect As Ingr.SP3D.ReferenceData.Middle.CrossSection
        crossSect = oMemberPart.CrossSection

        lvProperties.Items.Add(CreateNewLI("Name", oMemberPart.CrossSection.ToString(), group2))
        lvProperties.Items.Add(CreateNewLI("SectionStandard", crossSect.CrossSectionStandard.ToString(), group2))
        'lvProperties.Items.Add(CreateNewLI("SectionType", crossSect.CrossSectionType.ToString(), group2))
        lvProperties.Items.Add(CreateNewLI("SectionClass", crossSect.CrossSectionClass.ToString(), group2))

    End Sub

    Private Function GetMemberSystemFromPart(oMemberPart As MemberPart)
        Dim RC As RelationCollection = Nothing
        Dim oMemberSystem As MemberSystem
        oMemberSystem = oMemberPart.MemberSystem
        Return oMemberSystem

    End Function


    Private Sub GetMemberSystemProperties(oMemberSystem As MemberSystem)
        Call GetMemberSystemData(oMemberSystem)
        Dim oMemberParts = GetMemberPartsFromSystem(oMemberSystem)
        For Each oPart As BusinessObject In oMemberParts
            If oPart.SupportsInterface("ISPSMemberType") Then
                Call GetMemberPartPrismaticData(oPart)
            End If
        Next

    End Sub

    Private Function GetMemberPartsFromSystem(oMemberSystem As MemberSystem)
        Return oMemberSystem.SystemChildren

    End Function

    Private Sub GetMemberSystemData(oMemberSystem As MemberSystem)
        Dim group As New ListViewGroup("Member System")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("System Name", oMemberSystem.ToString(), group))
        lvProperties.Items.Add(CreateNewLI("System Type", GetCodelistedValue(oMemberSystem.GetPropertyValue("ISPSMemberType", "Type")), group))
        lvProperties.Items.Add(CreateNewLI("Length", oMemberSystem.GetPropertyValue("IJLine", "Length").ToString(), group))

    End Sub

    Private Sub GetSlabProperties(oSlab As Slab)
        Dim group As New ListViewGroup("Slab")
        lvProperties.Groups.Add(group)
        lvProperties.Items.Add(CreateNewLI("Priority", GetCodelistedValue(oSlab.GetPropertyValue("IJUASlabGeneralType", "Priority")), group))
        lvProperties.Items.Add(CreateNewLI("Composition", oSlab.Composition.ToString(), group))
    End Sub

    Private Function CreateNewLI(name As String, value As String, Optional group As ListViewGroup = Nothing)
        Dim str(2) As String
        str(0) = name
        str(1) = value

        Dim listItem As New ListViewItem(str)
        If Not group Is Nothing Then
            listItem.Group = group
        End If
        Return listItem

    End Function



    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        lvProperties.Width = Me.Width - 10
        lvProperties.Height = Me.Height - 40

    End Sub


    Private Function FormatDistanceValue(ByVal dValue) As String
        Return MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dValue)
    End Function

    Private Sub CheckBoxLongFormat_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLongFormat.CheckedChanged
        CheckBoxLongFormat.Invalidate()
        CheckBoxLongFormat.Update()
        Application.DoEvents()
        If current_bo <> Nothing Then
            Call ProcessSelection(current_bo)
        End If
    End Sub



    Private Sub frmPropertyInspector_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    Private Sub frmPropertyInspector_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If m_oEVSelectSet.SelectedObjects.Count <> 0 Then
            Call ProcessSelection(m_oEVSelectSet.SelectedObjects.Item(0))
        End If
    End Sub
End Class
