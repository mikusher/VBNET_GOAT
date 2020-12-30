Imports ASP
Imports System
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.Profile
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls

Namespace Book_Store
    Public Class CategoriesGrid
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub Categories_Bind()
            Me.Categories_Repeater.DataSource = Me.Categories_CreateDataSource
            Me.Categories_Repeater.DataBind
        End Sub

        Private Function Categories_CreateDataSource() As ICollection
            Me.Categories_sSQL = ""
            Me.Categories_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            str2 = " order by c.name Asc"
            If ((Me.Utility.GetParam("FormCategories_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("FormCategories_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            New StringDictionary
            Me.Categories_sSQL = "select [c].[category_id] as c_category_id, [c].[name] as c_name  from [categories] c "
            Me.Categories_sSQL = (Me.Categories_sSQL & str & str2)
            If (Me.Categories_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.Categories_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.Categories_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.Categories_sCountSQL = Me.Categories_sSQL.Replace(Me.Categories_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.Categories_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.Categories_sCountSQL = Me.Categories_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.Categories_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_Categories_curpage - 1) * 20), 20, "Categories")
            Dim command As New OleDbCommand(Me.Categories_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.Categories_Pager.MaxPage = IIf(((num3 Mod 20) > 0), ((num3 / 20) + 1), (num3 / 20))
            Dim flag As Boolean = (Me.Categories_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Categories_no_records.Visible = True
                flag = False
            Else
                Me.Categories_no_records.Visible = False
                flag = flag
            End If
            Me.Categories_Pager.Visible = flag
            Return view
        End Function

        Private Sub Categories_insert_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (Me.Categories_FormAction & "")
            MyBase.Response.Redirect(url)
        End Sub

        Protected Sub Categories_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_Categories_curpage = CurrPage
            Me.Categories_Bind
        End Sub

        Public Sub Categories_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Protected Sub Categories_SortChange(ByVal Src As Object, ByVal E As EventArgs)
            If ((Me.ViewState.Item("SortColumn") Is Nothing) OrElse (Me.ViewState.Item("SortColumn").ToString <> DirectCast(Src, LinkButton).CommandArgument)) Then
                Me.ViewState.Item("SortColumn") = DirectCast(Src, LinkButton).CommandArgument
                Me.ViewState.Item("SortDir") = "ASC"
            Else
                Me.ViewState.Item("SortDir") = IIf((Me.ViewState.Item("SortDir").ToString = "ASC"), "DESC", "ASC")
            End If
            Me.Categories_Bind
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Categories_insert.Click, New EventHandler(AddressOf Me.Categories_insert_Click)
            AddHandler Me.Categories_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.Categories_pager_navigate_completed)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Categories_Bind
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
        Protected Categories_Column_name As LinkButton
        Protected Categories_CountPage As Integer
        Protected Categories_FormAction As String = "CategoriesRecord.aspx?"
        Protected Categories_holder As HtmlTable
        Protected Categories_insert As LinkButton
        Protected Categories_no_records As HtmlTableRow
        Private Const Categories_PAGENUM As Integer = 20
        Protected Categories_Pager As CCPager
        Protected Categories_Repeater As Repeater
        Protected Categories_sCountSQL As String
        Protected Categories_sSQL As String
        Protected CategoriesForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected i_Categories_curpage As Integer = 1
        Protected Utility As CCUtility
    End Class
End Namespace

