﻿Imports ASP
Imports System
Imports System.Web.Profile
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls

Namespace Book_Store
    Public Class Footer
        Inherits UserControl
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Protected Sub Footer_Show()
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Footer_Show
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Public Sub ValidateNumeric(ByVal source As Object, ByVal args As ServerValidateEventArgs)
            Try 
                Decimal.Parse(args.Value)
                args.IsValid = True
            Catch obj1 As Object
                args.IsValid = False
            End Try
        End Sub


        ' Properties
        Protected ReadOnly Property ApplicationInstance As global_asax
            Get
                Return DirectCast(Me.Context.ApplicationInstance, global_asax)
            End Get
        End Property

        Protected ReadOnly Property Profile As DefaultProfile
            Get
                Return DirectCast(Me.Context.Profile, DefaultProfile)
            End Get
        End Property


        ' Fields
        Protected Footer_Field1 As HyperLink
        Protected Footer_Field2 As HyperLink
        Protected Footer_Field3 As HyperLink
        Protected Footer_Field4 As HyperLink
        Protected Footer_Field5 As HyperLink
        Protected Footer_FormAction As String = ".aspx?"
        Protected Footer_holder As HtmlTable
        Protected Utility As CCUtility
    End Class
End Namespace

