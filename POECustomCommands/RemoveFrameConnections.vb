Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Structure.Middle


Public Class RemoveFrameConnections
    Inherits BaseModalCommand
    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)
        MyBase.OnStart(instanceId, argument)

        Dim oSelectSet = ClientServiceProvider.SelectSet
        Dim oTransactionMgr = ClientServiceProvider.TransactionMgr

        Dim oHiliter As GraphicViewHiliter
        oHiliter = New GraphicViewHiliter
        oHiliter.Color = ColorConstants.RGBBlue
        oHiliter.Weight = 3
        'oHiliter.LinePattern = HiliterBase.HiliterLinePattern.Dotted


        If oSelectSet.Count = 0 Then MsgBox("No objects Selected") : Exit Sub

        Dim oObj As BusinessObject = ClientServiceProvider.SelectSet.SelectedObjects(0)
        'MsgBox(oObj.ClassInfo.DisplayName)

        If Not TypeOf oObj Is MemberSystem Then
            MsgBox("The command could be used only for Member System ") : Exit Sub
            Exit Sub
        End If

        Dim oMemberSystem As MemberSystem = oObj

        Dim endC_1 As FrameConnection = oMemberSystem.FrameConnection(MemberAxisEnd.Start)
        Dim endC_2 As FrameConnection = oMemberSystem.FrameConnection(MemberAxisEnd.End)

        If endC_1.PartName <> "Unsupported" Then
            oHiliter.HilitedObjects.Add(endC_1)
            endC_1.Delete()
            'MsgBox("C1 deleted")
        End If
        If endC_2.PartName <> "Unsupported" Then
            oHiliter.HilitedObjects.Add(endC_2)
            endC_2.Delete()

            ' MsgBox("C2 deleted")
        End If


        Dim connectedObjs = oMemberSystem.GetConnectedObjects()
        'Dim BOCClass2 As String
        For Each o In connectedObjs
            'MsgBox(o.ToString())
            'BOCClass2 = o.ClassInfo.DisplayName

            If TypeOf o Is MemberSystem Then
                Dim oMemberSystem2 As MemberSystem = o
                Dim endC_1_2 As FrameConnection = oMemberSystem2.FrameConnection(MemberAxisEnd.Start)
                Dim endC_2_2 As FrameConnection = oMemberSystem2.FrameConnection(MemberAxisEnd.End)

                If endC_1_2.PartName <> "Unsupported" And endC_1_2.GetRelatedObjects().Count() <> 0 Then
                    For Each ro In endC_1_2.GetRelatedObjects()
                        If ro.ObjectID = oObj.ObjectID Then
                            oHiliter.HilitedObjects.Add(endC_1_2)
                            endC_1_2.Delete()
                            'MsgBox("C1_2 deleted")
                        End If
                    Next
                End If
                If endC_2_2.PartName <> "Unsupported" And endC_2_2.GetRelatedObjects().Count() <> 0 Then
                    For Each ro In endC_2_2.GetRelatedObjects()
                        If ro.ObjectID = oObj.ObjectID Then
                            oHiliter.HilitedObjects.Add(endC_2_2)
                            endC_2_2.Delete()
                            'MsgBox("C2_2 deleted")
                        End If
                    Next
                End If
            End If



        Next

        'oTransactionMgr.Compute()
        Dim result As MsgBoxResult = MsgBox("Hiligted connection will be removed. Confirm?", vbYesNo,
                              "Remove Frame Connection")

        If result = vbYes Then
            oTransactionMgr.Commit("Delete Connections")
        Else
            oTransactionMgr.Abort()
        End If




    End Sub

End Class
