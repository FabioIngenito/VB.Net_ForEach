''' <summary>
''' VB .NET - Manipulando o DataGridView (Manipulating the DataGridView)
''' http://www.macoratti.net/07/06/vbn5_mdg.htm
''' </summary>
Public Class ClsVetor
    Private vetor As String

    Public Sub New(ByVal value As String)
        vetor = value
    End Sub

    Public Property nome() As String

        Get
            Return vetor
        End Get

        Set(value As String)
            vetor = value
        End Set

    End Property

End Class
