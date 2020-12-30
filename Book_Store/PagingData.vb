Imports System

Namespace Book_Store
    Public Class PagingData
        ' Methods
        Public Sub New(ByVal page As String, ByVal enabled As Boolean)
            Me.page = page
            Me.enabled = enabled
        End Sub


        ' Properties
        Public ReadOnly Property Enabled As Boolean
            Get
                Return Me.enabled
            End Get
        End Property

        Public ReadOnly Property Page As String
            Get
                Return Me.page
            End Get
        End Property


        ' Fields
        Private enabled As Boolean
        Private page As String
    End Class
End Namespace

