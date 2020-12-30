Imports ASP
Imports System
Imports System.Web.Profile
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls

Namespace Book_Store
    Public Class AdvSearch
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Search_search_button.Click, New EventHandler(AddressOf Me.Search_search_Click)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Search_Show
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Search_search_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (String.Concat(New String() { Me.Search_FormAction, "name=", Me.Search_name.Text, "&author=", Me.Search_author.Text, "&category_id=", Me.Search_category_id.SelectedItem.Value, "&pricemin=", Me.Search_pricemin.Text, "&pricemax=", Me.Search_pricemax.Text, "&" }) & "")
            MyBase.Response.Redirect(url)
        End Sub

        Private Sub Search_Show()
            Me.Utility.buildListBox(Me.Search_category_id.Items, "select category_id,name from categories order by 2", "category_id", "name", "All", "")
            Dim param As String = Me.Utility.GetParam("name")
            Me.Search_name.Text = param
            param = Me.Utility.GetParam("author")
            Me.Search_author.Text = param
            param = Me.Utility.GetParam("category_id")
            Try 
                Me.Search_category_id.SelectedIndex = Me.Search_category_id.Items.IndexOf(Me.Search_category_id.Items.FindByValue(param))
            Catch obj1 As Object
            End Try
            param = Me.Utility.GetParam("pricemin")
            Me.Search_pricemin.Text = param
            param = Me.Utility.GetParam("pricemax")
            Me.Search_pricemax.Text = param
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
        Protected Search_author As TextBox
        Protected Search_category_id As DropDownList
        Protected Search_FormAction As String = "Books.aspx?"
        Protected Search_holder As HtmlTable
        Protected Search_name As TextBox
        Protected Search_pricemax As TextBox
        Protected Search_pricemin As TextBox
        Protected Search_search_button As Button
        Protected SearchForm_Title As Label
        Protected Utility As CCUtility
    End Class
End Namespace

