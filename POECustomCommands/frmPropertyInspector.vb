Option Strict Off


Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services

Imports System.Windows.Forms
Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Collections.Generic



Public Class frmPropertyInspector
    Private m_oTransactionMgr As ClientTransactionManager
    Private WithEvents m_oEVSelectSet As SelectSet
    Private currentBOClass As String
    Private current_bo As BusinessObject
    Private GroupStyle As DataGridViewCellStyle
    Private m_UOM As UOMManager
    Private changeables As Dictionary(Of String, ChangeableEntity)
    Private cHelper As CatalogBaseHelper


    Private Sub frmPropertyInspector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        m_oEVSelectSet = ClientServiceProvider.SelectSet
        m_oTransactionMgr = ClientServiceProvider.TransactionMgr
        m_UOM = MiddleServiceProvider.UOMMgr
        cHelper = New CatalogBaseHelper

        'lvProperties.Width = Me.Width - 30
        'lvProperties.Height = Me.Height - 40
        lvProperties.CellBorderStyle = DataGridViewCellBorderStyle.Single
        lvProperties.ColumnCount = 3

        GroupStyle = New DataGridViewCellStyle()
        GroupStyle.Font = New Font(lvProperties.DefaultCellStyle.Font.FontFamily, 9, FontStyle.Bold)
        GroupStyle.ForeColor = Color.Blue

        changeables = New Dictionary(Of String, ChangeableEntity)

        current_bo = Nothing
    End Sub
    Private Sub lvProperties_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles lvProperties.CellPainting
        If e.ColumnIndex = 2 And e.RowIndex > -1 Then
            e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None
        End If
        If e.ColumnIndex = 1 And e.RowIndex > -1 Then
            e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None
        End If
    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        'lvProperties.Width = Me.Width - 20
        'lvProperties.Height = Me.Height - 50
        'lvProperties.Columns(0).MinimumWidth = (Me.Width - 20) * 0.3
        'lvProperties.Columns(1).MinimumWidth = (Me.Width - 20) * 0.7


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
        changeables.Clear()
        lvProperties.Rows.Clear()

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

        AddGroup("General")
        CreateNewLI("Class", currentBOClass)

        AddInterfaceTypeProperty(businessObject, "IJNamedItem", "Name", PropertyTypes.text, listItemName:="Name", changeable:=True)

        AddInterfaceTypeProperty(businessObject, "IJDObject", "ApprovalStatus", PropertyTypes.codelist, changeable:=True)

        Call GetConstructionInfoData(businessObject)

        If CheckBoxLongFormat.Checked = True Then
            CreateNewLI("UserLastModified", businessObject.UserLastModified.ToString())
            CreateNewLI("DateLastModified", businessObject.DateLastModified.ToString())
            CreateNewLI("UserCreated", businessObject.UserCreated.ToString())
        End If

        If currentBOClass.Contains("Feature") Then
            businessObject = getPartFromFeature(businessObject)
        End If

        If businessObject.SupportsInterface("IJMtoInfo") = True Then
            AddInterfaceTypeProperty(businessObject, "IJMTOInfo", "ReportingRequirements", PropertyTypes.codelist, changeable:=True)
        End If

    End Sub

    Private Sub GetConstructionInfoData(businessObject As BusinessObject)

        If currentBOClass.Contains("Feature") Then
            businessObject = getPartFromFeature(businessObject)
        End If

        If businessObject.SupportsInterface("IJConstructionInfo") = True Then
            AddInterfaceTypeProperty(businessObject, "IJConstructionInfo", "ConstructionRequirement", PropertyTypes.codelist, changeable:=True)
            AddInterfaceTypeProperty(businessObject, "IJConstructionInfo", "ConstructionType", PropertyTypes.codelist, changeable:=True, parentCodeListPropName:="ConstructionRequirement")
        End If
    End Sub

    Private Sub GetEqpProperties(businessObject As BusinessObject)
        Dim pos As Position
        Dim c1, c2, c3

        AddGroup("Equipment")

        CreateNewLI("Name", businessObject.ToString())
        pos = GetEqpPosition(businessObject)
        CreateNewLI("X", FormatDistanceValue(pos.X))
        CreateNewLI("Y", FormatDistanceValue(pos.Y))
        CreateNewLI("Z", FormatDistanceValue(pos.Z))
        'getting cataliog item from equipment
        Dim cBO As BusinessObject = businessObject.GetRelationship("SOtoSI_R", "toSI_ORIG").TargetObjects.Item(0)
        AddInterfaceTypeProperty(cBO, "IJEquipmentPart", "ProcessEqTypes0",
                                 PropertyTypes.codelist, listItemName:="Classification0", changeable:=True)
        AddInterfaceTypeProperty(cBO, "IJEquipmentPart", "ProcessEqTypes1",
                                 PropertyTypes.codelist, listItemName:="Classification1", changeable:=True, parentCodeListPropName:="ProcessEqTypes0")
        AddInterfaceTypeProperty(cBO, "IJEquipmentPart", "ProcessEqTypes2",
                                 PropertyTypes.codelist, listItemName:="Classification2", changeable:=True, parentCodeListPropName:="ProcessEqTypes1")
        AddInterfaceTypeProperty(cBO, "IJEquipmentPart", "ProcessEqTypes3",
                                 PropertyTypes.codelist, listItemName:="Classification3", changeable:=True, parentCodeListPropName:="ProcessEqTypes2")
        AddInterfaceTypeProperty(cBO, "IJEquipmentPart", "ProcessEqTypes4",
                                 PropertyTypes.codelist, listItemName:="Classification4", changeable:=True, parentCodeListPropName:="ProcessEqTypes3")
        AddInterfaceTypeProperty(cBO, "IJEquipmentPart", "ProcessEqTypes5",
                                 PropertyTypes.codelist, listItemName:="Classification5", changeable:=True, parentCodeListPropName:="ProcessEqTypes4")
        AddInterfaceTypeProperty(cBO, "IJEquipmentPart", "ProcessEqTypes6",
                                 PropertyTypes.codelist, listItemName:="Classification6", changeable:=True, parentCodeListPropName:="ProcessEqTypes5")



    End Sub
    Private Sub GetPipeNozzleProperties(oNozzle As PipeNozzle)
        Dim oConn As Ingr.SP3D.Route.Middle.Connection

        AddGroup("Nozzle")
        CreateNewLI("Name", oNozzle.ToString())
        CreateNewLI("Length", FormatDistanceValue(oNozzle.Length))

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

        AddGroup("Shape")

        AddInterfaceTypeProperty(oShape, "IJNamedItem", "Name",
                                 PropertyTypes.codelist, listItemName:="Shape Name", changeable:=True)

        If oShape.SupportsInterface("IJUACylinder") = True Then
            CreateNewLI("A", oShape.GetPropertyValue("IJUACylinder", "A").ToString())
            CreateNewLI("B", oShape.GetPropertyValue("IJUACylinder", "B").ToString())
            GoTo eqp
        End If

        If oShape.SupportsInterface("IJUACone") = True Then
            CreateNewLI("A", oShape.GetPropertyValue("IJUACone", "A").ToString())
            CreateNewLI("B", oShape.GetPropertyValue("IJUACone", "B").ToString())
            CreateNewLI("C", oShape.GetPropertyValue("IJUACone", "c").ToString())
            GoTo eqp
        End If
        If oShape.SupportsInterface("IJUAEccentricCone") = True Then
            CreateNewLI("A", oShape.GetPropertyValue("IJUAEccentricCone", "A").ToString())
            CreateNewLI("B", oShape.GetPropertyValue("IJUAEccentricCone", "B").ToString())
            CreateNewLI("C", oShape.GetPropertyValue("IJUAEccentricCone", "c").ToString())
            GoTo eqp
        End If

        If oShape.SupportsInterface("IJUARectSolid") = True Then
            CreateNewLI("A", oShape.GetPropertyValue("IJUARectSolid", "A").ToString())
            CreateNewLI("B", oShape.GetPropertyValue("IJUARectSolid", "B").ToString())
            CreateNewLI("C", oShape.GetPropertyValue("IJUARectSolid", "C").ToString())
            GoTo eqp
        End If
        If oShape.SupportsInterface("IJUASphere") = True Then
            CreateNewLI("A", oShape.GetPropertyValue("IJUASphere", "A").ToString())
            GoTo eqp
        End If

        If oShape.SupportsInterface("IJUASemiElliptical") = True Then
            CreateNewLI("A", oShape.GetPropertyValue("IJUASemiElliptical", "A").ToString())
            CreateNewLI("B", oShape.GetPropertyValue("IJUASemiElliptical", "B").ToString())
        End If
eqp:

        Dim oEquipment = oShape.SystemParent
        Call GetEqpProperties(oEquipment)
    End Sub

    Private Sub GetPipeProperties(businessObject As BusinessObject)
        Dim p1
        AddGroup("Pipe")

        AddInterfaceTypeProperty(businessObject, "IJRteStockPartOccur", "Length", PropertyTypes.text, changeable:=False)
        AddInterfaceTypeProperty(businessObject, "IJRteStockPartOccur", "CutLength", PropertyTypes.text, changeable:=False)

        p1 = businessObject.GetRelationship("madeFrom", "part").TargetObjects.Item(0)
        AddInterfaceTypeProperty(p1, "IJDPipeComponent", "IndustryCommodityCode", PropertyTypes.text, changeable:=False)
        AddInterfaceTypeProperty(p1, "IJDPipeComponent", "FirstSizeSchedule", PropertyTypes.text, changeable:=False)
        AddInterfaceTypeProperty(p1, "IJDPipeComponent", "SecondSizeSchedule", PropertyTypes.text, changeable:=False)

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
        AddGroup("Pipe Part")
        AddInterfaceTypeProperty(businessObject, "IJRtePartData", "ShortCode", PropertyTypes.text, changeable:=False)

        AddInterfaceTypeProperty(businessObject, "IJRtePartData", "OptionCode", PropertyTypes.codelist, changeable:=False)

        p1 = businessObject.GetRelationship("madeFrom", "part").TargetObjects.Item(0)
        AddInterfaceTypeProperty(p1, "IJDPipeComponent", "CommodityType", PropertyTypes.text, changeable:=False)
        AddInterfaceTypeProperty(p1, "IJDPipeComponent", "IndustryCommodityCode", PropertyTypes.text, listItemName:="CommodityCode", changeable:=False)
        AddInterfaceTypeProperty(p1, "IJDPipeComponent", "FirstSizeSchedule", PropertyTypes.text, changeable:=False)
        AddInterfaceTypeProperty(p1, "IJDPipeComponent", "SecondSizeSchedule", PropertyTypes.text, changeable:=False)

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
        AddGroup("Pipe Specialty")

        AddInterfaceTypeProperty(businessObject, "IJFaceToFace", "FacetoFace", PropertyTypes.distance, changeable:=True)
        AddInterfaceTypeProperty(businessObject, "IJFaceToCenter", "FacetoCenter", PropertyTypes.distance, changeable:=True)
        AddInterfaceTypeProperty(businessObject, "IJFaceToCenter", "Face1toCenter", PropertyTypes.distance, changeable:=True)
        AddInterfaceTypeProperty(businessObject, "IJFaceToCenter", "Face2toCenter", PropertyTypes.distance, changeable:=True)

        If businessObject.SupportsInterface("IJUAMucFlange") Then
            AddInterfaceTypeProperty(businessObject, "IJUAMucFlange", "Length", PropertyTypes.text, changeable:=True)
            AddInterfaceTypeProperty(businessObject, "IJUAMucFlange", "FlangeThk", PropertyTypes.text, changeable:=True)
        End If

        AddGroup("Material Control Data")

        Dim oRC As RelationCollection
        oRC = businessObject.GetRelationship("PartOccToMaterialControlData", "MaterialControlData")
        Dim oGenMat = oRC.TargetObjects.Item(0)
        AddInterfaceTypeProperty(oGenMat, "IJGenericMaterialControlData", "ContractorCommodityCode", PropertyTypes.text, changeable:=True)
        AddInterfaceTypeProperty(oGenMat, "IJGenericMaterialControlData", "ShortMaterialDescription", PropertyTypes.text, listItemName:="Description", changeable:=True)
        AddInterfaceTypeProperty(oGenMat, "IJGenericMaterialControlData", "BoltingRequirements", PropertyTypes.codelist, changeable:=True)
        AddInterfaceTypeProperty(oGenMat, "IJGenericMaterialControlData", "GasketRequirements", PropertyTypes.codelist, changeable:=True)
        AddInterfaceTypeProperty(oGenMat, "IJGenericMaterialControlData", "SubstCapScrewCntrCommodityCode", PropertyTypes.text, changeable:=True)

        AddInterfaceTypeProperty(oGenMat, "IJValveOperatorInfo", "ValveOperatorClass", PropertyTypes.codelist, listItemName:="Valve Operator Class", changeable:=True)
        AddInterfaceTypeProperty(oGenMat, "IJValveOperatorInfo", "ValveOperatorType", PropertyTypes.codelist, changeable:=True, listItemName:="Valve Operator Type", parentCodeListPropName:="ValveOperatorClass")


        oRC = oGenMat.GetRelationship("DefinesMaterialControlDataForComponent", "Component")
        Dim oComp = oRC.TargetObjects.Item(0)
        AddInterfaceTypeProperty(oComp, "IJDPart", "PartNumber", PropertyTypes.text, changeable:=False)

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
        AddGroup("Port " + oPort.GetPropertyValue("IJDPipePOrt", "PortIndex").ToString())

        CreateNewLI("NPD", oPort.GetPropertyValue("IJDPipePOrt", "NPD").ToString() & " " & oPort.GetPropertyValue("IJDPipePOrt", "NpdUnitType").ToString())
        CreateNewLI("Pressure", oPort.GetPropertyValue("IJDPipePOrt", "PressureRating").ToString())
        CreateNewLI("Term.Class", oPort.GetPropertyValue("IJDPipePOrt", "TerminationClass").ToString())
        CreateNewLI("EndPrep", FormatCodeListValue(oPort.GetPropertyValue("IJDPipePOrt", "EndPreparation")))
        CreateNewLI("EndPrep Descr.", GetCodelistedValue(oPort.GetPropertyValue("IJDPipePOrt", "EndPreparation"), short_format:=False))
        CreateNewLI("EndStd", FormatCodeListValue(oPort.GetPropertyValue("IJDPipePOrt", "EndStandard")))
        CreateNewLI("EndPractice", oPort.GetPropertyValue("IJDPipePOrt", "EndPractice").ToString())

    End Sub

    Private Function GetPipeRun(businessObject As BusinessObject)
        Dim oPipeRun As PipeRun
        Dim PiperunCol As RelationCollection = businessObject.GetRelationship("OwnsParts", "Owner")    'IJRtePathGenPart(Owner) -> OwnsParts(Run to Part) -> IJDesignParent(Parts),
        oPipeRun = PiperunCol.TargetObjects.Item(0)
        Return oPipeRun
    End Function
    Private Sub GetPipeRunData(oPipeRun As PipeRun)
        AddGroup("PipeRun")
        CreateNewLI("Run Name", oPipeRun.ToString())
        CreateNewLI("Piping Spec", oPipeRun.Specification.SpecificationName)

        AddInterfaceTypeProperty(oPipeRun, "IJRteInsulation", "Purpose", PropertyTypes.insulationSpec, listItemName:="Insulation Spec", changeable:=True)

        AddInterfaceTypeProperty(oPipeRun, "IJRteInsulation", "Purpose", PropertyTypes.codelist, listItemName:="Insulation Purpose", changeable:=True)
        AddInterfaceTypeProperty(oPipeRun, "IJRteInsulation", "Material", PropertyTypes.codelist, listItemName:="Insulation Material", changeable:=True)
        AddInterfaceTypeProperty(oPipeRun, "IJRteInsulation", "Thickness", PropertyTypes.codelist, listItemName:="Insulation Thikness", changeable:=True)


        Dim oInsulSpec As PipeInsulationSpec = Nothing
        Dim oPipeLine As Pipeline = oPipeRun.SystemParent
        Dim oPipelineSpecs As IAllowableSpecs = oPipeLine
        'Dim oSpecCollection As ReadOnlyCollection(Of SpecificationBase) = oPipelineSpecs.AllowableSpecs
        'Dim oSpec As BusinessObject
        'For Each oSpec In oPipelineSpecs.AllowableSpecs
        ' If oSpec.SupportsInterface("IJPipeInsulationSpec") Then
        ' oInsulSpec = oSpec
        ' MsgBox(oInsulSpec.SpecificationName)
        ' End If
        'Next
        'MsgBox(oPipeRun.InsulationSpec.ToString())


    End Sub


    Private Sub GetPipeLineProperties(oPipeLine As Pipeline)
        Call GetPipeLineData(oPipeLine)
    End Sub

    Private Sub GetPipeLineData(oPipeLine As Pipeline)
        AddGroup("Pipeline")
        CreateNewLI("Line Name", oPipeLine.ToString())
        AddInterfaceTypeProperty(oPipeLine, "IJPipelineSystem", "FluidSystem", PropertyTypes.codelist, listItemName:="Fluid System", changeable:=True)
        AddInterfaceTypeProperty(oPipeLine, "IJPipelineSystem", "FluidCode", PropertyTypes.codelist, listItemName:="Fluid Code", changeable:=True, parentCodeListPropName:="FluidSystem")
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
        AddGroup("Feature")
        AddInterfaceTypeProperty(oPipeFeature, "IJRtePipePathFeat", "CommodityOption", PropertyTypes.codelist, listItemName:="Option Code", changeable:=False)
        'AddInterfaceTypeProperty(oPipeFeature, "IJRtePipePathFeat", "Tag", PropertyTypes.text, changeable:=True)
        AddInterfaceTypeProperty(oPipeFeature, "IJRteInsulation", "Purpose", PropertyTypes.insulationSpec, listItemName:="Insulation Spec of Feat.", changeable:=True)
        AddInterfaceTypeProperty(oPipeFeature, "IJRteInsulation", "Purpose", PropertyTypes.codelist, listItemName:="Insulation Purpose of Feat.", changeable:=True)
        AddInterfaceTypeProperty(oPipeFeature, "IJRteInsulation", "Material", PropertyTypes.codelist, listItemName:="Insulation Material of Feat.", changeable:=True)
        AddInterfaceTypeProperty(oPipeFeature, "IJRteInsulation", "Thickness", PropertyTypes.codelist, listItemName:="Insulation Thikness of Feat.", changeable:=True)

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
        AddGroup("Connection " + connIndex.ToString())
        CreateNewLI("Type", oConn.GetPropertyValue("IJDistribConnection", "ConnectionType").ToString())
        Dim i = 1
        For Each oPart In oConn.ConnectionParts
            CreateNewLI("Part " + i.ToString(), oPart.ToString())
            i = i + 1
        Next

        Call GetConnectionItems(oConn)

    End Sub

    Private Sub GetConnectionItems(oConn As Ingr.SP3D.Route.Middle.Connection)
        Dim oImpliedMatingParts As RelationCollection
        Dim oBoltPart As BusinessObject
        Dim oGasketPart As BusinessObject
        For Each oItem In oConn.ConnectionItems
            Try
                If oItem.SupportsInterface("IJRteBolt") Then
                    oImpliedMatingParts = oItem.GetRelationship("ImpliedMatingParts", "UsedImpliedPart")
                    oBoltPart = oImpliedMatingParts.TargetObjects(0)
                    CreateNewLI("Bolt CommodityCode", oBoltPart.GetPropertyValue("IJBolt", "IndustryCommodityCode").ToString())
                    CreateNewLI("BoltType", oBoltPart.GetPropertyValue("IJBolt", "BoltType").ToString())

                    CreateNewLI("CalculatedLength", oItem.GetPropertyValue("IJRteBolt", "CalculatedLength").ToString())
                    CreateNewLI("RoundedLength", oItem.GetPropertyValue("IJRteBolt", "RoundedLength").ToString())
                    CreateNewLI("Quantity", oItem.GetPropertyValue("IJRteBolt", "Quantity").ToString())
                    CreateNewLI("Diameter", oItem.GetPropertyValue("IJRteBolt", "Diameter").ToString())
                End If

                If oItem.SupportsInterface("IJRteGasket") Then
                    oImpliedMatingParts = oItem.GetRelationship("ImpliedMatingParts", "UsedImpliedPart")
                    oGasketPart = oImpliedMatingParts.TargetObjects(0)
                    CreateNewLI("Gasket CommodityCode", oGasketPart.GetPropertyValue("IJGasket", "IndustryCommodityCode").ToString())
                    CreateNewLI("Gasket Thikness", oGasketPart.GetPropertyValue("IJGasket", "ThicknessFor3DModel").ToString())
                End If

            Catch

            End Try

        Next

    End Sub

    Private Sub GetNotes(oBO As BusinessObject)
        Dim group As New ListViewGroup("Notes")
        AddGroup("Notes")
        Dim RC As RelationCollection
        RC = oBO.GetRelationship("ContainsNote", "GeneralNote")
        If RC.TargetObjects.Count > 0 Then
            Dim i As Integer = 1
            For Each note As Note In RC.TargetObjects
                AddInterfaceTypeProperty(note, "IJGeneralNote", "Name", PropertyTypes.text, listItemName:="Note " + CStr(i) + " Name", changeable:=True)
                AddInterfaceTypeProperty(note, "IJGeneralNote", "Text", PropertyTypes.text, listItemName:="Note " + CStr(i) + " Text", changeable:=True)
                AddInterfaceTypeProperty(note, "IJGeneralNote", "Purpose", PropertyTypes.codelist, listItemName:="Note " + CStr(i) + " Purpose", changeable:=True)
                i = i + 1
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
        AddGroup("Member Part")
        CreateNewLI("Name", oMemberPart.Name)
        CreateNewLI("Type", GetCodelistedValue(oMemberPart.GetPropertyValue("ISPSMemberType", "Type")))
        CreateNewLI("CutLength", FormatDistanceValue(oMemberPart.CutLength))
        Dim group2 = New ListViewGroup("CrossSection")
        'lvProperties.Groups.Add(group2)
        Dim crossSect As Ingr.SP3D.ReferenceData.Middle.CrossSection
        crossSect = oMemberPart.CrossSection

        CreateNewLI("Name", oMemberPart.CrossSection.ToString())
        CreateNewLI("SectionStandard", crossSect.CrossSectionStandard.ToString())
        'CreateNewLI("SectionType", crossSect.CrossSectionType.ToString())
        CreateNewLI("SectionClass", crossSect.CrossSectionClass.ToString())

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
        AddGroup("Member System")
        CreateNewLI("System Name", oMemberSystem.ToString())
        CreateNewLI("System Type", GetCodelistedValue(oMemberSystem.GetPropertyValue("ISPSMemberType", "Type")))
        CreateNewLI("Length", oMemberSystem.GetPropertyValue("IJLine", "Length").ToString())

    End Sub

    Private Sub GetSlabProperties(oSlab As Slab)
        Dim group As New ListViewGroup("Slab")
        AddGroup("Slab")
        CreateNewLI("Priority", GetCodelistedValue(oSlab.GetPropertyValue("IJUASlabGeneralType", "Priority")))
        CreateNewLI("Composition", oSlab.Composition.ToString())
    End Sub

    Private Sub AddGroup(group As String)
        Dim row(2) As String
        Dim idx As Integer
        row(0) = group
        idx = lvProperties.Rows.Add(row)
        lvProperties.Rows(idx).DefaultCellStyle = GroupStyle
    End Sub

    Private Sub AddInterfaceTypeProperty(ByRef oBO As BusinessObject, ifaceName As String, propName As String, propType As PropertyTypes,
                                         Optional parentCodeListPropName As String = "", Optional listItemName As String = Nothing, Optional changeable As Boolean = False)
        Dim codelistInfo = Nothing

        If listItemName Is Nothing Then listItemName = propName
        If Not oBO.SupportsInterface(ifaceName) Then Exit Sub

        Dim pv As PropertyValue = oBO.GetPropertyValue(ifaceName, propName)

        Dim txtProperty As String = ""

        If propType = PropertyTypes.codelist Then
            codelistInfo = pv.PropertyInfo.CodeListInfo
        End If


        If propType = PropertyTypes.insulationSpec Then
            Dim pvcl As PropertyValueCodelist = pv
            If pvcl.PropValue = -1 Then
                txtProperty = "Not Insulated"
            Else
                txtProperty = "User Defined"
            End If
        Else
            If pv.ToString <> "" Then
                If propType = PropertyTypes.insulationSpec Then
                    Dim pvcl As PropertyValueCodelist = pv
                    If pvcl.PropValue = -1 Then
                        txtProperty = "Not Insulated"
                    Else
                        txtProperty = "User Defined"
                    End If
                Else

                    If pv.PropertyInfo.UOMType = UnitType.Undefined Then
                        txtProperty = pv.ToString()
                    Else
                        txtProperty = m_UOM.FormatUnit(pv)
                    End If
                End If
            End If
        End If

        CreateNewLI(listItemName, txtProperty, changeable:=changeable)
        If changeable = True And Not changeables.ContainsKey(listItemName) Then
            changeables.Add(listItemName,
                            New ChangeableEntity(oBO,
                                                 ifaceName:=ifaceName,
                                                 propName:=propName,
                                                 propType:=propType,
                                                 clInfo:=codelistInfo,
                                                 pclPropName:=parentCodeListPropName))
        End If
    End Sub



    Private Sub CreateNewLI(name As String, value As String, Optional changeable As Boolean = False)
        Dim row(2) As String
        Dim idx As Integer
        row(0) = name
        row(1) = value
        idx = lvProperties.Rows.Add(row)

        If changeable = True Then
            Dim btnCell = New DataGridViewButtonCell()
            lvProperties.Rows(idx).Cells(2) = btnCell
            'lvProperties.Rows(idx).Cells(2).Value = ".."

        End If

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


    Private Sub frmPropertyInspector_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If m_oEVSelectSet.SelectedObjects.Count <> 0 Then
            Call ProcessSelection(m_oEVSelectSet.SelectedObjects.Item(0))
        End If
    End Sub

    Private Sub lvProperties_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles lvProperties.CellContentClick

        If e.RowIndex <= 0 Or e.ColumnIndex <> 2 Then
            Exit Sub
        End If
        Dim propertyName As String = lvProperties(0, e.RowIndex).Value
        Dim currPropertyValue As String = lvProperties(1, e.RowIndex).Value

        Dim changeable As ChangeableEntity = Nothing
        changeable = changeables.Item(propertyName)
        If changeable Is Nothing Then Exit Sub


        If propertyName = "Insulation Material" Then
            Dim oInsulated As IPipeInsulation = changeable.item
            If oInsulated.InsulationPurpose = -1 Then
                MsgBox("Please select Insulation Purpose At first")
                Exit Sub
            End If
        End If


        If propertyName = "Insulation Thikness" Then
            Dim oInsulated As IPipeInsulation = changeable.item
            If oInsulated.InsulationMaterial = -1 Then
                MsgBox("Please select Insulation Material At first")
                Exit Sub
            End If
        End If

        Dim values As List(Of String) = Nothing
        If changeable.propertyType = PropertyTypes.codelist Or changeable.propertyType = PropertyTypes.insulationSpec Then
            values = GetPossiblePropertyValues(changeable)
            If values.Count() = 0 Then
                MsgBox("No possible values to change", vbCritical)
                Exit Sub
            End If
        End If

        Dim namingRule As String = Nothing
        If propertyName = "Name" Then
            namingRule = GetNamingRule(changeable.item)
        End If


        Dim NewValue As String
        If propertyName = "Insulation Thikness" Then
            NewValue = GetNewValue(PropertyTypes.codelist, currPropertyValue, values)
        Else
            NewValue = GetNewValue(changeable.propertyType, currPropertyValue, values, namingRule:=namingRule)
        End If

        If NewValue <> "" And NewValue <> currPropertyValue Then
            Try
                changeable.changeProperty(NewValue, m_UOM)
                If propertyName = "Insulation Thikness" Then
                    m_oTransactionMgr.Commit("Change " + propertyName)
                    Dim oInsulated As IPipeInsulation = changeable.item
                    Dim currentBOClass1 = changeable.item.ClassInfo.DisplayName
                    If currentBOClass1 = "Pipe Run" Then
                        Dim oPipeRun As PipeRun = changeable.item
                        oPipeRun.SetUserDefinedInsulation(oInsulated.InsulationPurpose, oInsulated.InsulationMaterial, oInsulated.InsulationThickness, 273)

                    Else
                        Dim oFeature As IPipePathFeature = changeable.item
                        oFeature.SetUserDefinedInsulation(oInsulated.InsulationPurpose, oInsulated.InsulationMaterial, oInsulated.InsulationThickness, 273)
                        ProcessSelection(current_bo)
                    End If
                End If

                m_oTransactionMgr.Commit("Change " + propertyName)
                lvProperties(1, e.RowIndex).Value = NewValue
                If changeable.propertyType = PropertyTypes.insulationSpec Then
                    If NewValue = "Not Insulated" Then
                        ProcessSelection(current_bo)
                    End If
                End If
            Catch
                If propertyName = "Insulation Thikness" Then
                    MsgBox("Ivalid combination Of Purpose, Material, thickness", vbCritical)
                Else
                    MsgBox("Failed to change property. Is object in Working status?", vbCritical)
                End If
            End Try
        End If
    End Sub

    Private Function GetNamingRule(oBo As BusinessObject)
        Dim oRelationColl As RelationCollection = oBo.GetRelationship("NamedEntity", "entityAE")
        Dim sNameRule As String
        If oRelationColl.TargetObjects.Count <> 0 Then
            sNameRule = oRelationColl.TargetObjects(0).GetRelationship("EntityNamingRule", "namingRule").TargetObjects(0).GetPropertyValue("IJDNameRuleHolder", "Name").ToString()
        Else
            sNameRule = "User Defined"
        End If
        Return sNameRule
    End Function


    Private Function GetPossiblePropertyValues(ByRef changeable As ChangeableEntity)
        Dim values = New List(Of String)
        If changeable.propertyType = PropertyTypes.insulationSpec Then
            values.Add("Not Insulated")
            values.Add("User Defined")

        ElseIf changeable.interfaceName = "IJRteInsulation" And changeable.propertyName = "Material" Then
            For Each material As InsulationMaterial In cHelper.GetInsulationMaterials(InsulationMaterialTypes.Piping)
                values.Add(changeable.codelistInfo.GetCodelistItem(material.MaterialType).ToString())
            Next
        ElseIf changeable.interfaceName = "IJRteInsulation" And changeable.propertyName = "Thickness" Then
            Dim oInsulated As IPipeInsulation = changeable.item
            For Each material As InsulationMaterial In cHelper.GetInsulationMaterials(InsulationMaterialTypes.Piping)
                If material.MaterialType = oInsulated.InsulationMaterial Then
                    For Each thk As Double In material.AllowableThicknesses(InsulationMaterialTypes.Piping)
                        values.Add(FormatDistanceValue(thk))
                    Next
                End If
            Next
        Else
            If changeable.parentCodeListPropName <> "" Then

                Dim pv As PropertyValueCodelist = changeable.item.GetPropertyValue(changeable.interfaceName, changeable.parentCodeListPropName)

                If Not changeable.codelistInfo.Parent.GetChildCodelistMembers(pv.PropValue) Is Nothing Then
                    For Each cl As CodelistItem In changeable.codelistInfo.Parent.GetChildCodelistMembers(pv.PropValue).Values
                        values.Add(cl.ShortDisplayName)
                    Next
                End If
            Else
                    For Each cl As CodelistItem In changeable.codelistInfo.CodelistMembers
                    values.Add(cl.ShortDisplayName)
                Next
            End If

        End If

            Return values
    End Function

    Private Function GetNewValue(propertyType As String, currValue As String, Optional values As List(Of String) = Nothing, Optional namingRule As String = Nothing)
        Dim m_frmChange As frmPropertyInspectorChange = New frmPropertyInspectorChange(propertyType, currValue, values, namingRule)
        Dim result As DialogResult = m_frmChange.ShowDialog(Me)
        If result = Windows.Forms.DialogResult.OK Then
            Return m_frmChange.result
        End If
        Return ""
    End Function

    Private Function FormatDistanceValue(ByVal val As Double) As String
        Return m_UOM.FormatUnit(UnitType.Distance, val)
    End Function


End Class

Public Class ChangeableEntity
    Public item As BusinessObject
    Public interfaceName As String = Nothing
    Public propertyName As String = Nothing
    Public propertyType As String = Nothing
    Public codelistInfo As CodelistInformation
    Public parentCodeListPropName As String = ""

    Public Sub New(ByRef newItem As BusinessObject,
                                ByVal Optional propType As PropertyTypes = PropertyTypes.text,
                                ByVal Optional ifaceName As String = Nothing,
                                ByVal Optional propName As String = Nothing,
                                ByVal Optional clInfo As CodelistInformation = Nothing,
                                ByVal Optional pclPropName As String = "")
        item = newItem
        propertyType = propType
        interfaceName = ifaceName
        propertyName = propName
        codelistInfo = clInfo
        parentCodeListPropName = pclPropName


    End Sub

    Public Sub changeProperty(ByVal newValue, ByRef m_UOM)
        If propertyName = "Name" And item.SupportsInterface("IJNamedItem") Then
            Dim namedItem As INamedItem = item
            namedItem.SetUserDefinedName(newValue)
        ElseIf Not interfaceName Is Nothing Then
            If propertyType = PropertyTypes.insulationSpec Then
                If newValue = "Not Insulated" Then
                    item.SetPropertyValue(-1, "IJRteInsulation", "Purpose")
                    item.SetPropertyValue(-1, "IJRteInsulation", "Material")
                    item.SetPropertyValue(False, "IJRteInsulation", "IsInsulated")
                End If

            ElseIf propertyType = PropertyTypes.codelist Then
                    If interfaceName = "IJRteInsulation" And propertyName = "Thickness" Then
                        Dim pv As PropertyValueDouble
                        pv = item.GetPropertyValue(interfaceName, propertyName)
                        pv.PropValue = m_UOM.ParseUnit(UnitType.Distance, newValue)
                        item.SetPropertyValue(pv.PropValue, interfaceName, propertyName)
                    Else
                        Dim oCL As CodelistItem
                        oCL = codelistInfo.GetCodelistItem(newValue)
                        item.SetPropertyValue(oCL, interfaceName, propertyName)
                    End If
                ElseIf propertyType = PropertyTypes.distance Then
                    Dim pv As PropertyValueDouble
                    pv = item.GetPropertyValue(interfaceName, propertyName)
                    pv.PropValue = m_UOM.ParseUnit(UnitType.Distance, newValue)
                    item.SetPropertyValue(pv.PropValue, interfaceName, propertyName)
                Else
                    item.SetPropertyValue(newValue, interfaceName, propertyName)
                End If
            End If
    End Sub


End Class
Public Enum PropertyTypes
    text = 1
    codelist = 2
    distance = 3
    insulationSpec = 4
End Enum

Public Enum InsulationSpecs
    NotInsulated = -1
    UserDefined = 1
End Enum