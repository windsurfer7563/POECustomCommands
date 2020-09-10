Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services.GraphicViewManager.GraphicViewEventArgs
Imports System.Windows.Forms


Public Class PropertyInspector
    Inherits BaseGraphicCommand


    Private WithEvents m_oEVMyForm As frmPropertyInspector

    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)
        Application.EnableVisualStyles()
        'Create form
        m_oEVMyForm = New frmPropertyInspector
        m_oEVMyForm.Show()

    End Sub


End Class
'C:\S3D\LPModeling\POECustomCommands\POECustomCommands\bin\Debug\POECustomCommands,POECustomCommands.PropertyInspector