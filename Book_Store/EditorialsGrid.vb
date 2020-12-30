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
    Public Class EditorialsGrid
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub editorials_Bind()
            Me.editorials_Repeater.DataSource = Me.editorials_CreateDataSource
            Me.editorials_Repeater.DataBind
        End Sub

        Private Function editorials_CreateDataSource() As ICollection
            Me.editorials_sSQL = ""
            Me.editorials_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            str2 = " order by e.article_title Asc"
            If ((Me.Utility.GetParam("Formeditorials_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("Formeditorials_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            New StringDictionary
            Me.editorials_sSQL = "select [e].[article_id] as e_article_id, [e].[article_title] as e_article_title, [e].[editorial_cat_id] as e_editorial_cat_id, [e].[item_id] as e_item_id, [e1].[editorial_cat_id] as e1_editorial_cat_id, [e1].[editorial_cat_name] as e1_editorial_cat_name, [i].[item_id] as i_item_id, [i].[name] as i_name  from [editorials] e, [editorial_categories] e1, [items] i where [e1].[editorial_cat_id]=e.[editorial_cat_id] and [i].[item_id]=e.[item_id]  "
            Me.editorials_sSQL = (Me.editorials_sSQL & str & str2)
            If (Me.editorials_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.editorials_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.editorials_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.editorials_sCountSQL = Me.editorials_sSQL.Replace(Me.editorials_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.editorials_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.editorials_sCountSQL = Me.editorials_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.editorials_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_editorials_curpage - 1) * 20), 20, "editorials")
            Dim command As New OleDbCommand(Me.editorials_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.editorials_Pager.MaxPage = IIf(((num3 Mod 20) > 0), ((num3 / 20) + 1), (num3 / 20))
            Dim flag As Boolean = (Me.editorials_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.editorials_no_records.Visible = True
                flag = False
            Else
                Me.editorials_no_records.Visible = False
                flag = flag
            End If
            Me.editorials_Pager.Visible = flag
            Return view
        End Function

        Private Sub editorials_insert_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (Me.editorials_FormAction & "")
            MyBase.Response.Redirect(url)
        End Sub

        Protected Sub editorials_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_editorials_curpage = CurrPage
            Me.editorials_Bind
        End Sub

        Public Sub editorials_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Protected Sub editorials_SortChange(ByVal Src As Object, ByVal E As EventArgs)
            If ((Me.ViewState.Item("SortColumn") Is Nothing) OrElse (Me.ViewState.Item("SortColumn").ToString <> DirectCast(Src, LinkButton).CommandArgument)) Then
                Me.ViewState.Item("SortColumn") = DirectCast(Src, LinkButton).CommandArgument
                Me.ViewState.Item("SortDir") = "ASC"
            Else
                Me.ViewState.Item("SortDir") = IIf((Me.ViewState.Item("SortDir").ToString = "ASC"), "DESC", "ASC")
            End If
            Me.editorials_Bind
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.editorials_insert.Click, New EventHandler(AddressOf Me.editorials_insert_Click)
            AddHandler Me.editorials_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.editorials_pager_navigate_completed)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_editorials_article_id.Value = Me.Utility.GetParam("article_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.editorials_Bind
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
        Protected editorials_Column_article_title As LinkButton
        Protected editorials_Column_editorial_cat_id As LinkButton
        Protected editorials_Column_item_id As LinkButton
        Protected editorials_CountPage As Integer
        Protected editorials_FormAction As String = "EditorialsRecord.aspx?"
        Protected editorials_holder As HtmlTable
        Protected editorials_insert As LinkButton
        Protected editorials_no_records As HtmlTableRow
        Private Const editorials_PAGENUM As Integer = 20
        Protected editorials_Pager As CCPager
        Protected editorials_Repeater As Repeater
        Protected editorials_sCountSQL As String
        Protected editorials_sSQL As String
        Protected editorialsForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected i_editorials_curpage As Integer = 1
        Protected p_editorials_article_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

