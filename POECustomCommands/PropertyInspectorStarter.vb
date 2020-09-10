Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services

Public Class PropertyInspectorStarter
    Inherits BaseModalCommand

    Public Overrides Sub OnStart(ByVal commandID As Integer, ByVal argument As Object)
        Dim iAsstID As Integer

        StartAssistant("POECustomCommands,POECustomCommands.PropertyInspector", iAsstID)
        ClientServiceProvider.ValueMgr.Add("CmdAsst-MyAsst1", iAsstID)

    End Sub

End Class
