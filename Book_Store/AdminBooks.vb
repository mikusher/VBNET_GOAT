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
    Public Class AdminBooks
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub Items_Bind()
            Me.Items_Repeater.DataSource = Me.Items_CreateDataSource
            Me.Items_Repeater.DataBind
        End Sub

        Private Function Items_CreateDataSource() As ICollection
            Me.Items_sSQL = ""
            Me.Items_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            Dim flag As Boolean = False
            If ((Me.Utility.GetParam("FormItems_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("FormItems_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("category_id") Then
                Dim param As String = Me.Utility.GetParam("category_id")
                If (Me.Utility.IsNumeric(Nothing, param) AndAlso (param.Length > 0)) Then
                    param = CCUtility.ToSQL(param, FieldTypes.Number)
                Else
                    param = ""
                End If
                dictionary.Add("category_id", param)
            End If
            If Not dictionary.ContainsKey("is_recommended") Then
                Dim str4 As String = Me.Utility.GetParam("is_recommended")
                If (Me.Utility.IsNumeric(Nothing, str4) AndAlso (str4.Length > 0)) Then
                    str4 = CCUtility.ToSQL(str4, FieldTypes.Number)
                Else
                    str4 = ""
                End If
                dictionary.Add("is_recommended", str4)
            End If
            If (dictionary.Item("category_id").Length > 0) Then
                flag = True
                str = (str & "i.[category_id]=" & dictionary.Item("category_id"))
            End If
            If (dictionary.Item("is_recommended").Length > 0) Then
                If (str.Length > 0) Then
                    str = (str & " and ")
                End If
                flag = True
                str = (str & "i.[is_recommended]=" & dictionary.Item("is_recommended"))
            End If
            If flag Then
                str = (" AND (" & str & ")")
            End If
            Me.Items_sSQL = "select [i].[author] as i_author, [i].[category_id] as i_category_id, [i].[is_recommended] as i_is_recommended, [i].[item_id] as i_item_id, [i].[name] as i_name, [i].[price] as i_price, [c].[category_id] as c_category_id, [c].[name] as c_name  from [items] i, [categories] c where [c].[category_id]=i.[category_id]  "
            Me.Items_sSQL = (Me.Items_sSQL & str & str2)
            If (Me.Items_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.Items_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.Items_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.Items_sCountSQL = Me.Items_sSQL.Replace(Me.Items_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.Items_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.Items_sCountSQL = Me.Items_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.Items_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_Items_curpage - 1) * 20), 20, "Items")
            Dim command As New OleDbCommand(Me.Items_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.Items_Pager.MaxPage = IIf(((num3 Mod 20) > 0), ((num3 / 20) + 1), (num3 / 20))
            Dim flag2 As Boolean = (Me.Items_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Items_no_records.Visible = True
                flag2 = False
            Else
                Me.Items_no_records.Visible = False
                flag2 = flag2
            End If
            Me.Items_Pager.Visible = flag2
            Return view
        End Function

        Private Sub Items_insert_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = String.Concat(New String() { Me.Items_FormAction, "category_id=", MyBase.Server.UrlEncode(Me.Utility.GetParam("category_id")), "&is_recommended=", MyBase.Server.UrlEncode(Me.Utility.GetParam("is_recommended")), "&" })
            MyBase.Response.Redirect(url)
        End Sub

        Protected Sub Items_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_Items_curpage = CurrPage
            Me.Items_Bind
        End Sub

        Public Sub Items_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Protected Sub Items_SortChange(ByVal Src As Object, ByVal E As EventArgs)
            If ((Me.ViewState.Item("SortColumn") Is Nothing) OrElse (Me.ViewState.Item("SortColumn").ToString <> DirectCast(Src, LinkButton).CommandArgument)) Then
                Me.ViewState.Item("SortColumn") = DirectCast(Src, LinkButton).CommandArgument
                Me.ViewState.Item("SortDir") = "ASC"
            Else
                Me.ViewState.Item("SortDir") = IIf((Me.ViewState.Item("SortDir").ToString = "ASC"), "DESC", "ASC")
            End If
            Me.Items_Bind
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Search_search_button.Click, New EventHandler(AddressOf Me.Search_search_Click)
            AddHandler Me.Items_insert.Click, New EventHandler(AddressOf Me.Items_insert_Click)
            AddHandler Me.Items_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.Items_pager_navigate_completed)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Search_Show
            Me.Items_Bind
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Search_search_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (String.Concat(New String() { Me.Search_FormAction, "category_id=", Me.Search_category_id.SelectedItem.Value, "&is_recommended=", Me.Search_is_recommended.SelectedItem.Value, "&" }) & "")
            MyBase.Response.Redirect(url)
        End Sub

        Private Sub Search_Show()
            Me.Utility.buildListBox(Me.Search_category_id.Items, "select category_id,name from categories order by 2", "category_id", "name", "All", "")
            Me.Utility.buildListBox(Me.Search_is_recommended.Items, Me.Search_is_recommended_lov, Nothing, "")
            Dim param As String = Me.Utility.GetParam("category_id")
            Try 
                Me.Search_category_id.SelectedIndex = Me.Search_category_id.Items.IndexOf(Me.Search_category_id.Items.FindByValue(param))
            Catch obj1 As Object
            End Try
            param = Me.Utility.GetParam("is_recommended")
            Try 
                Me.Search_is_recommended.SelectedIndex = Me.Search_is_recommended.Items.IndexOf(Me.Search_is_recommended.Items.FindByValue(param))
            Catch obj2 As Object
            End Try
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
        Protected i_Items_curpage As Integer = 1
        Protected Items_Column_author As LinkButton
        Protected Items_Column_category_id As LinkButton
        Protected Items_Column_Field1 As Label
        Protected Items_Column_is_recommended As LinkButton
        Protected Items_Column_name As LinkButton
        Protected Items_Column_price As LinkButton
        Protected Items_CountPage As Integer
        Protected Items_FormAction As String = "BookMaint.aspx?"
        Protected Items_holder As HtmlTable
        Protected Items_insert As LinkButton
        Protected Items_is_recommended_lov As String() = "0;No;1;Yes".Split(New Char() { ";"c })
        Protected Items_no_records As HtmlTableRow
        Private Const Items_PAGENUM As Integer = 20
        Protected Items_Pager As CCPager
        Protected Items_Repeater As Repeater
        Protected Items_sCountSQL As String
        Protected Items_sSQL As String
        Protected ItemsForm_Title As Label
        Protected Search_category_id As DropDownList
        Protected Search_FormAction As String = "AdminBooks.aspx?"
        Protected Search_holder As HtmlTable
        Protected Search_is_recommended As DropDownList
        Protected Search_is_recommended_lov As String() = ";All;0;No;1;Yes".Split(New Char() { ";"c })
        Protected Search_search_button As Button
        Protected Utility As CCUtility
    End Class
End Namespace

