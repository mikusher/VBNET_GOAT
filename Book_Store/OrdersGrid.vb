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
    Public Class OrdersGrid
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub Orders_Bind()
            Me.Orders_Repeater.DataSource = Me.Orders_CreateDataSource
            Me.Orders_Repeater.DataBind
        End Sub

        Private Function Orders_CreateDataSource() As ICollection
            Me.Orders_sSQL = ""
            Me.Orders_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            Dim flag As Boolean = False
            str2 = " order by o.order_id Asc"
            If ((Me.Utility.GetParam("FormOrders_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("FormOrders_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("item_id") Then
                Dim param As String = Me.Utility.GetParam("item_id")
                If (Me.Utility.IsNumeric(Nothing, param) AndAlso (param.Length > 0)) Then
                    param = CCUtility.ToSQL(param, FieldTypes.Number)
                Else
                    param = ""
                End If
                dictionary.Add("item_id", param)
            End If
            If Not dictionary.ContainsKey("member_id") Then
                Dim str4 As String = Me.Utility.GetParam("member_id")
                If (Me.Utility.IsNumeric(Nothing, str4) AndAlso (str4.Length > 0)) Then
                    str4 = CCUtility.ToSQL(str4, FieldTypes.Number)
                Else
                    str4 = ""
                End If
                dictionary.Add("member_id", str4)
            End If
            If (dictionary.Item("item_id").Length > 0) Then
                flag = True
                str = (str & "o.[item_id]=" & dictionary.Item("item_id"))
            End If
            If (dictionary.Item("member_id").Length > 0) Then
                If (str.Length > 0) Then
                    str = (str & " and ")
                End If
                flag = True
                str = (str & "o.[member_id]=" & dictionary.Item("member_id"))
            End If
            If flag Then
                str = (" AND (" & str & ")")
            End If
            Me.Orders_sSQL = "select [o].[item_id] as o_item_id, [o].[member_id] as o_member_id, [o].[order_id] as o_order_id, [o].[quantity] as o_quantity, [i].[item_id] as i_item_id, [i].[name] as i_name, [m].[member_id] as m_member_id, [m].[member_login] as m_member_login  from [orders] o, [items] i, [members] m where [i].[item_id]=o.[item_id] and [m].[member_id]=o.[member_id]  "
            Me.Orders_sSQL = (Me.Orders_sSQL & str & str2)
            If (Me.Orders_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.Orders_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.Orders_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.Orders_sCountSQL = Me.Orders_sSQL.Replace(Me.Orders_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.Orders_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.Orders_sCountSQL = Me.Orders_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.Orders_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_Orders_curpage - 1) * 20), 20, "Orders")
            Dim command As New OleDbCommand(Me.Orders_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.Orders_Pager.MaxPage = IIf(((num3 Mod 20) > 0), ((num3 / 20) + 1), (num3 / 20))
            Dim flag2 As Boolean = (Me.Orders_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Orders_no_records.Visible = True
                flag2 = False
            Else
                Me.Orders_no_records.Visible = False
                flag2 = flag2
            End If
            Me.Orders_Pager.Visible = flag2
            Return view
        End Function

        Private Sub Orders_insert_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = String.Concat(New String() { Me.Orders_FormAction, "item_id=", MyBase.Server.UrlEncode(Me.Utility.GetParam("item_id")), "&member_id=", MyBase.Server.UrlEncode(Me.Utility.GetParam("member_id")), "&" })
            MyBase.Response.Redirect(url)
        End Sub

        Protected Sub Orders_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_Orders_curpage = CurrPage
            Me.Orders_Bind
        End Sub

        Public Sub Orders_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Protected Sub Orders_SortChange(ByVal Src As Object, ByVal E As EventArgs)
            If ((Me.ViewState.Item("SortColumn") Is Nothing) OrElse (Me.ViewState.Item("SortColumn").ToString <> DirectCast(Src, LinkButton).CommandArgument)) Then
                Me.ViewState.Item("SortColumn") = DirectCast(Src, LinkButton).CommandArgument
                Me.ViewState.Item("SortDir") = "ASC"
            Else
                Me.ViewState.Item("SortDir") = IIf((Me.ViewState.Item("SortDir").ToString = "ASC"), "DESC", "ASC")
            End If
            Me.Orders_Bind
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Search_search_button.Click, New EventHandler(AddressOf Me.Search_search_Click)
            AddHandler Me.Orders_insert.Click, New EventHandler(AddressOf Me.Orders_insert_Click)
            AddHandler Me.Orders_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.Orders_pager_navigate_completed)
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
            Me.Orders_Bind
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Search_search_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (String.Concat(New String() { Me.Search_FormAction, "item_id=", Me.Search_item_id.SelectedItem.Value, "&member_id=", Me.Search_member_id.SelectedItem.Value, "&" }) & "")
            MyBase.Response.Redirect(url)
        End Sub

        Private Sub Search_Show()
            Me.Utility.buildListBox(Me.Search_item_id.Items, "select item_id,name from items order by 2", "item_id", "name", "All", "")
            Me.Utility.buildListBox(Me.Search_member_id.Items, "select member_id,member_login from members order by 2", "member_id", "member_login", "All", "")
            Dim param As String = Me.Utility.GetParam("item_id")
            Try 
                Me.Search_item_id.SelectedIndex = Me.Search_item_id.Items.IndexOf(Me.Search_item_id.Items.FindByValue(param))
            Catch obj1 As Object
            End Try
            param = Me.Utility.GetParam("member_id")
            Try 
                Me.Search_member_id.SelectedIndex = Me.Search_member_id.Items.IndexOf(Me.Search_member_id.Items.FindByValue(param))
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
        Protected i_Orders_curpage As Integer = 1
        Protected Orders_Column_Field1 As Label
        Protected Orders_Column_item_id As LinkButton
        Protected Orders_Column_member_id As LinkButton
        Protected Orders_Column_quantity As LinkButton
        Protected Orders_CountPage As Integer
        Protected Orders_FormAction As String = "OrdersRecord.aspx?"
        Protected Orders_holder As HtmlTable
        Protected Orders_insert As LinkButton
        Protected Orders_no_records As HtmlTableRow
        Private Const Orders_PAGENUM As Integer = 20
        Protected Orders_Pager As CCPager
        Protected Orders_Repeater As Repeater
        Protected Orders_sCountSQL As String
        Protected Orders_sSQL As String
        Protected OrdersForm_Title As Label
        Protected Search_FormAction As String = "OrdersGrid.aspx?"
        Protected Search_holder As HtmlTable
        Protected Search_item_id As DropDownList
        Protected Search_member_id As DropDownList
        Protected Search_search_button As Button
        Protected Utility As CCUtility
    End Class
End Namespace

