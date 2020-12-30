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
    Public Class [Default]
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Protected Sub AdvMenu_Show()
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
            New StringDictionary
            Me.Categories_sSQL = "select [c].[category_id] as c_category_id, [c].[name] as c_name  from [categories] c "
            Me.Categories_sSQL = (Me.Categories_sSQL & str & str2)
            Dim adapter As New OleDbDataAdapter(Me.Categories_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "Categories")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Categories_no_records.Visible = True
                Return view
            End If
            Me.Categories_no_records.Visible = False
            Return view
        End Function

        Public Sub Categories_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub New_Bind()
            Me.New_Repeater.DataSource = Me.New_CreateDataSource
            Me.New_Repeater.DataBind
        End Sub

        Private Function New_CreateDataSource() As ICollection
            Me.New_sSQL = ""
            Me.New_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            New StringDictionary
            str = " WHERE editorial_cat_id=2"
            Me.New_sSQL = "select [e].[article_desc] as e_article_desc, [e].[article_title] as e_article_title, [e].[item_id] as e_item_id  from [editorials] e "
            Me.New_sSQL = (Me.New_sSQL & str & str2)
            Dim adapter As New OleDbDataAdapter(Me.New_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "New")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.New_no_records.Visible = True
                Return view
            End If
            Me.New_no_records.Visible = False
            Return view
        End Function

        Public Sub New_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
            If ((e.Item.ItemType = ListItemType.Item) OrElse (e.Item.ItemType = ListItemType.AlternatingItem)) Then
                DirectCast(e.Item.FindControl("New_article_title"), HyperLink).Text = ("<b>" & DirectCast(e.Item.DataItem, DataRowView).Item("e_article_title").ToString & "<b>")
                DirectCast(e.Item.FindControl("New_article_desc"), Label).Text = ("<img align=""left"" border=""0"" src=""" & Me.Utility.Dlookup("items", "image_url", ("item_id=" & DirectCast(e.Item.DataItem, DataRowView).Item("e_item_id").ToString)) & """>" & DirectCast(e.Item.DataItem, DataRowView).Item("e_article_desc").ToString)
            End If
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Search_search_button.Click, New EventHandler(AddressOf Me.Search_search_Click)
            AddHandler Me.Recommended_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.Recommended_pager_navigate_completed)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Search_Show
            Me.AdvMenu_Show
            Me.Recommended_Bind
            Me.What_Bind
            Me.Categories_Bind
            Me.New_Bind
            Me.Weekly_Bind
            Me.Specials_Bind
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Recommended_Bind()
            Me.Recommended_Repeater.DataSource = Me.Recommended_CreateDataSource
            Me.Recommended_Repeater.DataBind
        End Sub

        Private Function Recommended_CreateDataSource() As ICollection
            Me.Recommended_sSQL = ""
            Me.Recommended_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            New StringDictionary
            str = " WHERE is_recommended=1"
            Me.Recommended_sSQL = "select [i].[author] as i_author, [i].[image_url] as i_image_url, [i].[item_id] as i_item_id, [i].[name] as i_name, [i].[price] as i_price  from [items] i "
            Me.Recommended_sSQL = (Me.Recommended_sSQL & str & str2)
            If (Me.Recommended_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.Recommended_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.Recommended_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.Recommended_sCountSQL = Me.Recommended_sSQL.Replace(Me.Recommended_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.Recommended_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.Recommended_sCountSQL = Me.Recommended_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.Recommended_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_Recommended_curpage - 1) * 5), 5, "Recommended")
            Dim command As New OleDbCommand(Me.Recommended_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.Recommended_Pager.MaxPage = IIf(((num3 Mod 5) > 0), ((num3 / 5) + 1), (num3 / 5))
            Dim flag As Boolean = (Me.Recommended_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Recommended_no_records.Visible = True
                flag = False
            Else
                Me.Recommended_no_records.Visible = False
                flag = flag
            End If
            Me.Recommended_Pager.Visible = flag
            Return view
        End Function

        Protected Sub Recommended_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_Recommended_curpage = CurrPage
            Me.Recommended_Bind
        End Sub

        Public Sub Recommended_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
            If ((e.Item.ItemType = ListItemType.Item) OrElse (e.Item.ItemType = ListItemType.AlternatingItem)) Then
                DirectCast(e.Item.FindControl("Recommended_name"), HyperLink).Text = String.Concat(New String() { "<img border=""0"" src=""", DirectCast(e.Item.DataItem, DataRowView).Item("i_image_url").ToString, """></td><td valign=""top""><table width=""100%"" style=""width:100%""><tr><td style=""background-color: #FFFFFF; border-style: inset; border-width: 0""><font style=""font-size: 10pt; color: #CE7E00; font-weight: bold""><b>", DirectCast(e.Item.DataItem, DataRowView).Item("i_name").ToString, "</b>" })
            End If
        End Sub

        Private Sub Search_search_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (String.Concat(New String() { Me.Search_FormAction, "category_id=", Me.Search_category_id.SelectedItem.Value, "&name=", Me.Search_name.Text, "&" }) & "")
            MyBase.Response.Redirect(url)
        End Sub

        Private Sub Search_Show()
            Me.Utility.buildListBox(Me.Search_category_id.Items, "select category_id,name from categories order by 2", "category_id", "name", "All", "")
            Dim param As String = Me.Utility.GetParam("category_id")
            Try 
                Me.Search_category_id.SelectedIndex = Me.Search_category_id.Items.IndexOf(Me.Search_category_id.Items.FindByValue(param))
            Catch obj1 As Object
            End Try
            param = Me.Utility.GetParam("name")
            Me.Search_name.Text = param
        End Sub

        Private Sub Specials_Bind()
            Me.Specials_Repeater.DataSource = Me.Specials_CreateDataSource
            Me.Specials_Repeater.DataBind
        End Sub

        Private Function Specials_CreateDataSource() As ICollection
            Me.Specials_sSQL = ""
            Me.Specials_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            New StringDictionary
            str = " WHERE editorial_cat_id=4"
            Me.Specials_sSQL = "select [e].[article_desc] as e_article_desc, [e].[article_title] as e_article_title  from [editorials] e "
            Me.Specials_sSQL = (Me.Specials_sSQL & str & str2)
            Dim adapter As New OleDbDataAdapter(Me.Specials_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "Specials")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Specials_no_records.Visible = True
                Return view
            End If
            Me.Specials_no_records.Visible = False
            Return view
        End Function

        Public Sub Specials_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Public Sub ValidateNumeric(ByVal source As Object, ByVal args As ServerValidateEventArgs)
            Try 
                Decimal.Parse(args.Value)
                args.IsValid = True
            Catch obj1 As Object
                args.IsValid = False
            End Try
        End Sub

        Private Sub Weekly_Bind()
            Me.Weekly_Repeater.DataSource = Me.Weekly_CreateDataSource
            Me.Weekly_Repeater.DataBind
        End Sub

        Private Function Weekly_CreateDataSource() As ICollection
            Me.Weekly_sSQL = ""
            Me.Weekly_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            New StringDictionary
            str = " WHERE editorial_cat_id=3"
            Me.Weekly_sSQL = "select [e].[article_desc] as e_article_desc, [e].[article_title] as e_article_title, [e].[item_id] as e_item_id  from [editorials] e "
            Me.Weekly_sSQL = (Me.Weekly_sSQL & str & str2)
            Dim adapter As New OleDbDataAdapter(Me.Weekly_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "Weekly")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Weekly_no_records.Visible = True
                Return view
            End If
            Me.Weekly_no_records.Visible = False
            Return view
        End Function

        Public Sub Weekly_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
            If ((e.Item.ItemType = ListItemType.Item) OrElse (e.Item.ItemType = ListItemType.AlternatingItem)) Then
                DirectCast(e.Item.FindControl("Weekly_article_title"), HyperLink).Text = ("<b>" & DirectCast(e.Item.DataItem, DataRowView).Item("e_article_title").ToString & "<b>")
                DirectCast(e.Item.FindControl("Weekly_article_desc"), Label).Text = ("<img align=""left"" border=""0"" src=""" & Me.Utility.Dlookup("items", "image_url", ("item_id=" & DirectCast(e.Item.DataItem, DataRowView).Item("e_item_id").ToString)) & """>" & DirectCast(e.Item.DataItem, DataRowView).Item("e_article_desc").ToString)
            End If
        End Sub

        Private Sub What_Bind()
            Me.What_Repeater.DataSource = Me.What_CreateDataSource
            Me.What_Repeater.DataBind
        End Sub

        Private Function What_CreateDataSource() As ICollection
            Me.What_sSQL = ""
            Me.What_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            New StringDictionary
            str = " WHERE editorial_cat_id=1"
            Me.What_sSQL = "select [e].[article_desc] as e_article_desc, [e].[article_title] as e_article_title, [e].[item_id] as e_item_id  from [editorials] e "
            Me.What_sSQL = (Me.What_sSQL & str & str2)
            Dim adapter As New OleDbDataAdapter(Me.What_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "What")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.What_no_records.Visible = True
                Return view
            End If
            Me.What_no_records.Visible = False
            Return view
        End Function

        Public Sub What_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
            If ((e.Item.ItemType = ListItemType.Item) OrElse (e.Item.ItemType = ListItemType.AlternatingItem)) Then
                DirectCast(e.Item.FindControl("What_article_title"), HyperLink).Text = ("<b>" & DirectCast(e.Item.DataItem, DataRowView).Item("e_article_title").ToString & "<b>")
                DirectCast(e.Item.FindControl("What_article_desc"), Label).Text = ("<img align=""left"" border=""0"" src=""" & Me.Utility.Dlookup("items", "image_url", ("item_id=" & DirectCast(e.Item.DataItem, DataRowView).Item("e_item_id").ToString)) & """>" & DirectCast(e.Item.DataItem, DataRowView).Item("e_article_desc").ToString)
            End If
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
        Protected AdvMenu_Field1 As HyperLink
        Protected AdvMenu_FormAction As String = ".aspx?"
        Protected AdvMenu_holder As HtmlTable
        Protected AdvMenuForm_Title As Label
        Protected Categories_Column_name As Label
        Protected Categories_CountPage As Integer
        Protected Categories_FormAction As String = ".aspx?"
        Protected Categories_holder As HtmlTable
        Protected Categories_no_records As HtmlTableRow
        Private Const Categories_PAGENUM As Integer = 20
        Protected Categories_Repeater As Repeater
        Protected Categories_sCountSQL As String
        Protected Categories_sSQL As String
        Protected CategoriesForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected i_Categories_curpage As Integer = 1
        Protected i_New_curpage As Integer = 1
        Protected i_Recommended_curpage As Integer = 1
        Protected i_Specials_curpage As Integer = 1
        Protected i_Weekly_curpage As Integer = 1
        Protected i_What_curpage As Integer = 1
        Protected New_CountPage As Integer
        Protected New_FormAction As String = ".aspx?"
        Protected New_holder As HtmlTable
        Protected New_no_records As HtmlTableRow
        Private Const New_PAGENUM As Integer = 20
        Protected New_Repeater As Repeater
        Protected New_sCountSQL As String
        Protected New_sSQL As String
        Protected NewForm_Title As Label
        Protected Recommended_CountPage As Integer
        Protected Recommended_FormAction As String = ".aspx?"
        Protected Recommended_holder As HtmlTable
        Protected Recommended_no_records As HtmlTableRow
        Private Const Recommended_PAGENUM As Integer = 5
        Protected Recommended_Pager As CCPager
        Protected Recommended_Repeater As Repeater
        Protected Recommended_sCountSQL As String
        Protected Recommended_sSQL As String
        Protected RecommendedForm_Title As Label
        Protected Search_category_id As DropDownList
        Protected Search_FormAction As String = "Books.aspx?"
        Protected Search_holder As HtmlTable
        Protected Search_name As TextBox
        Protected Search_search_button As Button
        Protected SearchForm_Title As Label
        Protected Specials_CountPage As Integer
        Protected Specials_FormAction As String = ".aspx?"
        Protected Specials_holder As HtmlTable
        Protected Specials_no_records As HtmlTableRow
        Private Const Specials_PAGENUM As Integer = 20
        Protected Specials_Repeater As Repeater
        Protected Specials_sCountSQL As String
        Protected Specials_sSQL As String
        Protected SpecialsForm_Title As Label
        Protected Utility As CCUtility
        Protected Weekly_CountPage As Integer
        Protected Weekly_FormAction As String = ".aspx?"
        Protected Weekly_holder As HtmlTable
        Protected Weekly_no_records As HtmlTableRow
        Private Const Weekly_PAGENUM As Integer = 20
        Protected Weekly_Repeater As Repeater
        Protected Weekly_sCountSQL As String
        Protected Weekly_sSQL As String
        Protected WeeklyForm_Title As Label
        Protected What_CountPage As Integer
        Protected What_FormAction As String = ".aspx?"
        Protected What_holder As HtmlTable
        Protected What_no_records As HtmlTableRow
        Private Const What_PAGENUM As Integer = 20
        Protected What_Repeater As Repeater
        Protected What_sCountSQL As String
        Protected What_sSQL As String
        Protected WhatForm_Title As Label
    End Class
End Namespace

