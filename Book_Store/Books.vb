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
    Public Class Books
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Protected Sub AdvMenu_Show()
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Results_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.Results_pager_navigate_completed)
            AddHandler Me.Search_search_button.Click, New EventHandler(AddressOf Me.Search_search_Click)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            If Not MyBase.IsPostBack Then
                Me.p_Total_item_id.Value = Me.Utility.GetParam("item_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Results_Bind
            Me.Search_Show
            Me.AdvMenu_Show
            Me.Total_Bind
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Results_Bind()
            Me.Results_Repeater.DataSource = Me.Results_CreateDataSource
            Me.Results_Repeater.DataBind
        End Sub

        Private Function Results_CreateDataSource() As ICollection
            Me.Results_sSQL = ""
            Me.Results_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            Dim flag As Boolean = False
            str2 = " order by i.name Asc"
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("author") Then
                Dim param As String = Me.Utility.GetParam("author")
                dictionary.Add("author", param)
            End If
            If Not dictionary.ContainsKey("category_id") Then
                Dim str4 As String = Me.Utility.GetParam("category_id")
                If (Me.Utility.IsNumeric(Nothing, str4) AndAlso (str4.Length > 0)) Then
                    str4 = CCUtility.ToSQL(str4, FieldTypes.Number)
                Else
                    str4 = ""
                End If
                dictionary.Add("category_id", str4)
            End If
            If Not dictionary.ContainsKey("name") Then
                Dim str5 As String = Me.Utility.GetParam("name")
                dictionary.Add("name", str5)
            End If
            If Not dictionary.ContainsKey("pricemax") Then
                Dim str6 As String = Me.Utility.GetParam("pricemax")
                If (Me.Utility.IsNumeric(Nothing, str6) AndAlso (str6.Length > 0)) Then
                    str6 = CCUtility.ToSQL(str6, FieldTypes.Number)
                Else
                    str6 = ""
                End If
                dictionary.Add("pricemax", str6)
            End If
            If Not dictionary.ContainsKey("pricemin") Then
                Dim str7 As String = Me.Utility.GetParam("pricemin")
                If (Me.Utility.IsNumeric(Nothing, str7) AndAlso (str7.Length > 0)) Then
                    str7 = CCUtility.ToSQL(str7, FieldTypes.Number)
                Else
                    str7 = ""
                End If
                dictionary.Add("pricemin", str7)
            End If
            If (dictionary.Item("author").Length > 0) Then
                flag = True
                str = (str & "i.[author] like '%" & dictionary.Item("author").Replace("'", "''") & "%'")
            End If
            If (dictionary.Item("category_id").Length > 0) Then
                If (str.Length > 0) Then
                    str = (str & " and ")
                End If
                flag = True
                str = (str & "i.[category_id]=" & dictionary.Item("category_id"))
            End If
            If (dictionary.Item("name").Length > 0) Then
                If (str.Length > 0) Then
                    str = (str & " and ")
                End If
                flag = True
                str = (str & "i.[name] like '%" & dictionary.Item("name").Replace("'", "''") & "%'")
            End If
            If (dictionary.Item("pricemax").Length > 0) Then
                If (str.Length > 0) Then
                    str = (str & " and ")
                End If
                flag = True
                str = (str & "i.[price]<" & dictionary.Item("pricemax"))
            End If
            If (dictionary.Item("pricemin").Length > 0) Then
                If (str.Length > 0) Then
                    str = (str & " and ")
                End If
                flag = True
                str = (str & "i.[price]>" & dictionary.Item("pricemin"))
            End If
            If flag Then
                str = (" AND (" & str & ")")
            End If
            Me.Results_sSQL = "select [i].[author] as i_author, [i].[category_id] as i_category_id, [i].[image_url] as i_image_url, [i].[item_id] as i_item_id, [i].[name] as i_name, [i].[price] as i_price, [c].[category_id] as c_category_id, [c].[name] as c_name  from [items] i, [categories] c where [c].[category_id]=i.[category_id]  "
            Me.Results_sSQL = (Me.Results_sSQL & str & str2)
            If (Me.Results_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.Results_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.Results_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.Results_sCountSQL = Me.Results_sSQL.Replace(Me.Results_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.Results_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.Results_sCountSQL = Me.Results_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.Results_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_Results_curpage - 1) * 20), 20, "Results")
            Dim command As New OleDbCommand(Me.Results_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.Results_Pager.MaxPage = IIf(((num3 Mod 20) > 0), ((num3 / 20) + 1), (num3 / 20))
            Dim flag2 As Boolean = (Me.Results_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Results_no_records.Visible = True
                flag2 = False
            Else
                Me.Results_no_records.Visible = False
                flag2 = flag2
            End If
            Me.Results_Pager.Visible = flag2
            Return view
        End Function

        Protected Sub Results_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_Results_curpage = CurrPage
            Me.Results_Bind
        End Sub

        Public Sub Results_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
            If ((e.Item.ItemType = ListItemType.Item) OrElse (e.Item.ItemType = ListItemType.AlternatingItem)) Then
                DirectCast(e.Item.FindControl("Results_name"), HyperLink).Text = String.Concat(New String() { "<img border=""0"" src=""", DirectCast(e.Item.DataItem, DataRowView).Item("i_image_url").ToString, """></td><td valign=""top"" width=""100%""><table ><tr><td style=""background-color: #FFFFFF; border-style: inset; border-width: 0""><font style=""font-size: 10pt; color: #CE7E00; font-weight: bold""><b>", DirectCast(e.Item.DataItem, DataRowView).Item("i_name").ToString, "</b>" })
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

        Private Sub Total_Bind()
            Me.Total_Repeater.DataSource = Me.Total_CreateDataSource
            Me.Total_Repeater.DataBind
        End Sub

        Private Function Total_CreateDataSource() As ICollection
            Me.Total_sSQL = ""
            Me.Total_sCountSQL = ""
            Dim sWhere As String = ""
            Dim sOrder As String = ""
            Dim flag As Boolean = False
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("author") Then
                Dim param As String = Me.Utility.GetParam("author")
                dictionary.Add("author", param)
            End If
            If Not dictionary.ContainsKey("category_id") Then
                Dim str4 As String = Me.Utility.GetParam("category_id")
                If (Me.Utility.IsNumeric(Nothing, str4) AndAlso (str4.Length > 0)) Then
                    str4 = CCUtility.ToSQL(str4, FieldTypes.Number)
                Else
                    str4 = ""
                End If
                dictionary.Add("category_id", str4)
            End If
            If Not dictionary.ContainsKey("name") Then
                Dim str5 As String = Me.Utility.GetParam("name")
                dictionary.Add("name", str5)
            End If
            If Not dictionary.ContainsKey("pricemax") Then
                Dim str6 As String = Me.Utility.GetParam("pricemax")
                If (Me.Utility.IsNumeric(Nothing, str6) AndAlso (str6.Length > 0)) Then
                    str6 = CCUtility.ToSQL(str6, FieldTypes.Number)
                Else
                    str6 = ""
                End If
                dictionary.Add("pricemax", str6)
            End If
            If Not dictionary.ContainsKey("pricemin") Then
                Dim str7 As String = Me.Utility.GetParam("pricemin")
                If (Me.Utility.IsNumeric(Nothing, str7) AndAlso (str7.Length > 0)) Then
                    str7 = CCUtility.ToSQL(str7, FieldTypes.Number)
                Else
                    str7 = ""
                End If
                dictionary.Add("pricemin", str7)
            End If
            If (dictionary.Item("author").Length > 0) Then
                flag = True
                sWhere = (sWhere & "i.[author] like '%" & dictionary.Item("author").Replace("'", "''") & "%'")
            End If
            If (dictionary.Item("category_id").Length > 0) Then
                If (sWhere.Length > 0) Then
                    sWhere = (sWhere & " and ")
                End If
                flag = True
                sWhere = (sWhere & "i.[category_id]=" & dictionary.Item("category_id"))
            End If
            If (dictionary.Item("name").Length > 0) Then
                If (sWhere.Length > 0) Then
                    sWhere = (sWhere & " and ")
                End If
                flag = True
                sWhere = (sWhere & "i.[name] like '%" & dictionary.Item("name").Replace("'", "''") & "%'")
            End If
            If (dictionary.Item("pricemax").Length > 0) Then
                If (sWhere.Length > 0) Then
                    sWhere = (sWhere & " and ")
                End If
                flag = True
                sWhere = (sWhere & "i.[price]<=" & dictionary.Item("pricemax"))
            End If
            If (dictionary.Item("pricemin").Length > 0) Then
                If (sWhere.Length > 0) Then
                    sWhere = (sWhere & " and ")
                End If
                flag = True
                sWhere = (sWhere & "i.[price]>=" & dictionary.Item("pricemin"))
            End If
            If flag Then
                sWhere = (" WHERE (" & sWhere & ")")
            End If
            Me.Total_sSQL = "select [i].[author] as i_author, [i].[category_id] as i_category_id, [i].[item_id] as i_item_id, [i].[name] as i_name, [i].[price] as i_price  from [items] i "
            Me.Total_Open((sWhere), (sOrder))
            Me.Total_sSQL = (Me.Total_sSQL & sWhere & sOrder)
            Dim adapter As New OleDbDataAdapter(Me.Total_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "Total")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Total_no_records.Visible = True
                Return view
            End If
            Me.Total_no_records.Visible = False
            Return view
        End Function

        Public Sub Total_Open(ByRef sWhere As String, ByRef sOrder As String)
            Me.Total_sSQL = "select count(item_id) as i_item_id from items as i"
        End Sub

        Public Sub Total_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
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
        Protected AdvMenu_Field1 As HyperLink
        Protected AdvMenu_FormAction As String = ".aspx?"
        Protected AdvMenu_holder As HtmlTable
        Protected Footer As Footer
        Protected Header As Header
        Protected i_Results_curpage As Integer = 1
        Protected i_Total_curpage As Integer = 1
        Protected p_Total_item_id As HtmlInputHidden
        Protected Results_CountPage As Integer
        Protected Results_FormAction As String = ".aspx?"
        Protected Results_holder As HtmlTable
        Protected Results_no_records As HtmlTableRow
        Private Const Results_PAGENUM As Integer = 20
        Protected Results_Pager As CCPager
        Protected Results_Repeater As Repeater
        Protected Results_sCountSQL As String
        Protected Results_sSQL As String
        Protected ResultsForm_Title As Label
        Protected Search_category_id As DropDownList
        Protected Search_FormAction As String = "Books.aspx?"
        Protected Search_holder As HtmlTable
        Protected Search_name As TextBox
        Protected Search_search_button As Button
        Protected Total_CountPage As Integer
        Protected Total_FormAction As String = ".aspx?"
        Protected Total_holder As HtmlTable
        Protected Total_no_records As HtmlTableRow
        Private Const Total_PAGENUM As Integer = 20
        Protected Total_Repeater As Repeater
        Protected Total_sCountSQL As String
        Protected Total_sSQL As String
        Protected TotalForm_Title As Label
        Protected Utility As CCUtility
    End Class
End Namespace

