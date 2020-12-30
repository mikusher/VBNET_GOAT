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
    Public Class EditorialCatGrid
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub editorial_categories_Bind()
            Me.editorial_categories_Repeater.DataSource = Me.editorial_categories_CreateDataSource
            Me.editorial_categories_Repeater.DataBind
        End Sub

        Private Function editorial_categories_CreateDataSource() As ICollection
            Me.editorial_categories_sSQL = ""
            Me.editorial_categories_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            str2 = " order by e.editorial_cat_name Asc"
            If ((Me.Utility.GetParam("Formeditorial_categories_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("Formeditorial_categories_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            New StringDictionary
            Me.editorial_categories_sSQL = "select [e].[editorial_cat_id] as e_editorial_cat_id, [e].[editorial_cat_name] as e_editorial_cat_name  from [editorial_categories] e "
            Me.editorial_categories_sSQL = (Me.editorial_categories_sSQL & str & str2)
            If (Me.editorial_categories_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.editorial_categories_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.editorial_categories_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.editorial_categories_sCountSQL = Me.editorial_categories_sSQL.Replace(Me.editorial_categories_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.editorial_categories_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.editorial_categories_sCountSQL = Me.editorial_categories_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.editorial_categories_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_editorial_categories_curpage - 1) * 20), 20, "editorial_categories")
            Dim command As New OleDbCommand(Me.editorial_categories_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.editorial_categories_Pager.MaxPage = IIf(((num3 Mod 20) > 0), ((num3 / 20) + 1), (num3 / 20))
            Dim flag As Boolean = (Me.editorial_categories_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.editorial_categories_no_records.Visible = True
                flag = False
            Else
                Me.editorial_categories_no_records.Visible = False
                flag = flag
            End If
            Me.editorial_categories_Pager.Visible = flag
            Return view
        End Function

        Private Sub editorial_categories_insert_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (Me.editorial_categories_FormAction & "")
            MyBase.Response.Redirect(url)
        End Sub

        Protected Sub editorial_categories_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_editorial_categories_curpage = CurrPage
            Me.editorial_categories_Bind
        End Sub

        Public Sub editorial_categories_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Protected Sub editorial_categories_SortChange(ByVal Src As Object, ByVal E As EventArgs)
            If ((Me.ViewState.Item("SortColumn") Is Nothing) OrElse (Me.ViewState.Item("SortColumn").ToString <> DirectCast(Src, LinkButton).CommandArgument)) Then
                Me.ViewState.Item("SortColumn") = DirectCast(Src, LinkButton).CommandArgument
                Me.ViewState.Item("SortDir") = "ASC"
            Else
                Me.ViewState.Item("SortDir") = IIf((Me.ViewState.Item("SortDir").ToString = "ASC"), "DESC", "ASC")
            End If
            Me.editorial_categories_Bind
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.editorial_categories_insert.Click, New EventHandler(AddressOf Me.editorial_categories_insert_Click)
            AddHandler Me.editorial_categories_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.editorial_categories_pager_navigate_completed)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_editorial_categories_editorial_cat_id.Value = Me.Utility.GetParam("editorial_cat_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.editorial_categories_Bind
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
        Protected editorial_categories_Column_editorial_cat_name As LinkButton
        Protected editorial_categories_CountPage As Integer
        Protected editorial_categories_FormAction As String = "EditorialCatRecord.aspx?"
        Protected editorial_categories_holder As HtmlTable
        Protected editorial_categories_insert As LinkButton
        Protected editorial_categories_no_records As HtmlTableRow
        Private Const editorial_categories_PAGENUM As Integer = 20
        Protected editorial_categories_Pager As CCPager
        Protected editorial_categories_Repeater As Repeater
        Protected editorial_categories_sCountSQL As String
        Protected editorial_categories_sSQL As String
        Protected editorial_categoriesForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected i_editorial_categories_curpage As Integer = 1
        Protected p_editorial_categories_editorial_cat_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

