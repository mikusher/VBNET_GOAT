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
    Public Class MembersInfo
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
            Dim flag As Boolean = True
            Dim flag2 As Boolean = False
            If ((Me.Utility.GetParam("FormOrders_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("FormOrders_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("member_id") Then
                Dim param As String = Me.Utility.GetParam("member_id")
                If (Me.Utility.IsNumeric(Nothing, param) AndAlso (param.Length > 0)) Then
                    param = CCUtility.ToSQL(param, FieldTypes.Number)
                Else
                    param = ""
                End If
                dictionary.Add("member_id", param)
            End If
            If (dictionary.Item("member_id").Length > 0) Then
                flag2 = True
                str = (str & "o.[member_id]=" & dictionary.Item("member_id"))
            Else
                flag = False
            End If
            If flag2 Then
                str = (" AND (" & str & ")")
            End If
            Me.Orders_sSQL = "select [o].[item_id] as o_item_id, [o].[member_id] as o_member_id, [o].[order_id] as o_order_id, [o].[quantity] as o_quantity, [i].[item_id] as i_item_id, [i].[name] as i_name  from [orders] o, [items] i where [i].[item_id]=o.[item_id]  "
            Me.Orders_sSQL = (Me.Orders_sSQL & str & str2)
            If Not flag Then
                Me.Orders_no_records.Visible = True
                Return Nothing
            End If
            Dim adapter As New OleDbDataAdapter(Me.Orders_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "Orders")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Orders_no_records.Visible = True
                Return view
            End If
            Me.Orders_no_records.Visible = False
            Return view
        End Function

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
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_Record_member_id.Value = Me.Utility.GetParam("member_id")
                Me.p_Orders_order_id.Value = Me.Utility.GetParam("order_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Record_Show
            Me.Orders_Bind
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Record_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Sub Record_Show()
            Dim flag As Boolean = True
            If (Me.p_Record_member_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Record_member_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from members where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Record") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Record_member_id.Value = CCUtility.GetValue(row, "member_id")
                    Me.Record_member_login.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "member_login").ToString)
                    Me.Record_member_login.NavigateUrl = ("MembersRecord.aspx?member_id=" & MyBase.Server.UrlEncode(CCUtility.GetValue(row, "member_id").ToString) & "&")
                    Me.Record_member_level.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "member_level").ToString)
                    Me.Record_name.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "first_name").ToString)
                    Me.Record_last_name.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "last_name").ToString)
                    Me.Record_email.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "email").ToString)
                    Me.Record_phone.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "phone").ToString)
                    Me.Record_address.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "address").ToString)
                    Me.Record_notes.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "notes").ToString)
                    flag = False
                End If
            End If
        End Sub

        Private Function Record_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Record_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Record") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Record_ValidationSummary.Text = (Me.Record_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Record_ValidationSummary.Visible = Not flag
            Return flag
        End Function

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
        Protected Orders_Column_item_id As LinkButton
        Protected Orders_Column_order_id As LinkButton
        Protected Orders_Column_quantity As LinkButton
        Protected Orders_CountPage As Integer
        Protected Orders_FormAction As String = "AdminMenu.aspx?"
        Protected Orders_holder As HtmlTable
        Protected Orders_no_records As HtmlTableRow
        Private Const Orders_PAGENUM As Integer = 20
        Protected Orders_Repeater As Repeater
        Protected Orders_sCountSQL As String
        Protected Orders_sSQL As String
        Protected OrdersForm_Title As Label
        Protected p_Orders_order_id As HtmlInputHidden
        Protected p_Record_member_id As HtmlInputHidden
        Protected Record_address As Label
        Protected Record_email As Label
        Protected Record_FormAction As String = "AdminMenu.aspx?"
        Protected Record_holder As HtmlTable
        Protected Record_last_name As Label
        Protected Record_member_id As HtmlInputHidden
        Protected Record_member_level As Label
        Protected Record_member_login As HyperLink
        Protected Record_name As Label
        Protected Record_notes As Label
        Protected Record_phone As Label
        Protected Record_ValidationSummary As Label
        Protected RecordForm_Title As Label
        Protected Utility As CCUtility
    End Class
End Namespace

