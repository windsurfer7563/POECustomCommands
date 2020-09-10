Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services.GraphicViewManager.GraphicViewEventArgs
Imports System.Windows.Forms




Public Class DefineWorkspace
    Inherits BaseGraphicCommand

    Private WithEvents m_oEVMyForm As frmDefineWorksace

    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)

        'Create form
        m_oEVMyForm = New frmDefineWorksace
        Application.EnableVisualStyles()
        m_oEVMyForm.Show()

    End Sub

End Class