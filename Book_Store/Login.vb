Imports ASP
Imports System
Imports System.Web.Profile
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls

Namespace Book_Store
    Public Class Login
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub Login_login_Click(ByVal Src As Object, ByVal E As EventArgs)
            If Me.Login_logged Then
                Me.Login_logged = False
                Me.Session.Item("UserID") = 0
                Me.Session.Item("UserRights") = 0
                Me.Login_Show
            ElseIf (Convert.ToInt32(Me.Utility.Dlookup("members", "count(*)", String.Concat(New String() { "member_login ='", Me.Login_name.Text, "' and member_password='", CCUtility.Quote(Me.Login_password.Text), "'" }))) > 0) Then
                Me.Login_message.Visible = False
                Me.Session.Item("UserID") = Convert.ToInt32(Me.Utility.Dlookup("members", "member_id", String.Concat(New String() { "member_login ='", Me.Login_name.Text, "' and member_password='", CCUtility.Quote(Me.Login_password.Text), "'" })))
                Me.Login_logged = True
                Me.Session.Item("UserRights") = Convert.ToInt32(Me.Utility.Dlookup("members", "member_level", String.Concat(New String() { "member_login ='", Me.Login_name.Text, "' and member_password='", CCUtility.Quote(Me.Login_password.Text), "'" })))
                Dim param As String = Me.Utility.GetParam("querystring")
                Dim str2 As String = Me.Utility.GetParam("ret_page")
                If (Not str2.Equals(MyBase.Request.ServerVariables.Item("SCRIPT_NAME")) AndAlso (str2.Length > 0)) Then
                    MyBase.Response.Redirect((str2 & "?" & param))
                Else
                    MyBase.Response.Redirect(Me.Login_FormAction)
                End If
            Else
                Me.Login_message.Visible = True
            End If
        End Sub

        Private Sub Login_Show()
            If Me.Login_logged Then
                Me.Login_login.Text = "Logout"
                Me.Login_trpassword.Visible = False
                Me.Login_trname.Visible = False
                Me.Login_labelname.Visible = True
                Me.Login_labelname.Text = (Me.Utility.Dlookup("members", "member_login", ("member_id=" & Me.Session.Item("UserID"))) & "&nbsp;&nbsp;&nbsp;")
            Else
                Me.Login_login.Text = "Login"
                Me.Login_trpassword.Visible = True
                Me.Login_trname.Visible = True
                Me.Login_labelname.Visible = False
            End If
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Login_login.Click, New EventHandler(AddressOf Me.Login_login_Click)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            If ((Not Me.Session.Item("UserID") Is Nothing) AndAlso (Short.Parse(Me.Session.Item("UserID").ToString) > 0)) Then
                Me.Login_logged = True
            End If
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Login_Show
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
        Protected Footer As Footer
        Protected Header As Header
        Protected Login_FormAction As String = "ShoppingCart.aspx?"
        Protected Login_holder As HtmlTable
        Protected Login_labelname As Label
        Protected Login_logged As Boolean
        Protected Login_login As Button
        Protected Login_message As Label
        Protected Login_name As TextBox
        Protected Login_password As TextBox
        Protected Login_querystring As HtmlInputHidden
        Protected Login_ret_page As HtmlInputHidden
        Protected Login_trname As HtmlTableRow
        Protected Login_trpassword As HtmlTableRow
        Protected LoginForm_Title As Label
        Protected Utility As CCUtility
    End Class
End Namespace

