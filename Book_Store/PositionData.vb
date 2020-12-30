Imports System

Namespace Book_Store
    Public Class PositionData
        ' Methods
        Public Sub New(ByVal Name As String, ByVal CategoryID As String)
            Me.name = Name
            Me.CatID = CategoryID
        End Sub


        ' Properties
        Public ReadOnly Property CategoryID As String
            Get
                Return Me.CatID
            End Get
        End Property

        Public ReadOnly Property Name As String
            Get
                Return Me.name
            End Get
        End Property


        ' Fields
        Private CatID As String
        Private name As String
    End Class
End Namespace

