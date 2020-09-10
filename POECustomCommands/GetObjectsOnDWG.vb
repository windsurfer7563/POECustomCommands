Imports Ingr.SP3D.Common.Client
Imports System.Windows.Forms


Public Class GetObjectsOnDWG
    Inherits BaseModalCommand

    Private WithEvents m_oEVMyForm As frmGetObjectsOnDwg

    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)

        'Create form
        m_oEVMyForm = New frmGetObjectsOnDwg
        Application.EnableVisualStyles()
        m_oEVMyForm.ShowDialog()


    End Sub


End Class
