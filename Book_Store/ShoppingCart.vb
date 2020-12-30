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
    Public Class ShoppingCart
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
            Dim flag As Boolean = True
            Dim flag2 As Boolean = False
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("UserID") Then
                Dim str3 As String
                Dim obj2 As Object = Me.Session.Item("UserID")
                If (obj2 Is Nothing) Then
                    str3 = ""
                Else
                    str3 = obj2.ToString
                End If
                If (Me.Utility.IsNumeric(Nothing, str3) AndAlso (str3.Length > 0)) Then
                    str3 = CCUtility.ToSQL(str3, FieldTypes.Number)
                Else
                    str3 = ""
                End If
                dictionary.Add("UserID", str3)
            End If
            If (dictionary.Item("UserID").Length > 0) Then
                flag2 = True
                str = (str & "[member_id]=" & dictionary.Item("UserID"))
            Else
                flag = False
            End If
            If flag2 Then
                str = (" AND (" & str & ")")
            End If
            Me.Items_sSQL = "SELECT order_id, name, price, quantity, member_id, quantity*price as sub_total FROM items, orders WHERE orders.item_id=items.item_id"
            str2 = " ORDER BY order_id"
            Me.Items_sSQL = (Me.Items_sSQL & str & str2)
            If Not flag Then
                Me.Items_no_records.Visible = True
                Return Nothing
            End If
            Dim adapter As New OleDbDataAdapter(Me.Items_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "Items")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Items_no_records.Visible = True
                Return view
            End If
            Me.Items_no_records.Visible = False
            Return view
        End Function

        Public Sub Items_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Private Sub Member_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Sub Member_Show()
            Dim flag As Boolean = True
            If (Me.p_Member_member_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Member_member_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from members where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Member") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Member_member_id.Value = CCUtility.GetValue(row, "member_id")
                    Me.Member_member_login.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "member_login").ToString)
                    Me.Member_member_login.NavigateUrl = "MyInfo.aspx?"
                    Me.Member_name.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "first_name").ToString)
                    Me.Member_last_name.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "last_name").ToString)
                    Me.Member_address.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "address").ToString)
                    Me.Member_email.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "email").ToString)
                    Me.Member_phone.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "phone").ToString)
                    flag = False
                End If
            End If
            If flag Then
                Dim str3 As String
                If (Not Me.Session.Item("UserID") Is Nothing) Then
                    str3 = Me.Session.Item("UserID").ToString
                Else
                    str3 = ""
                End If
                Me.Member_member_id.Value = str3
            End If
        End Sub

        Private Function Member_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Member_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Member") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Member_ValidationSummary.Text = (Me.Member_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Member_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(1)
            If Not MyBase.IsPostBack Then
                Me.p_Items_order_id.Value = Me.Utility.GetParam("order_id")
                If (Not Me.Session.Item("UserID") Is Nothing) Then
                    Me.p_Member_member_id.Value = Me.Session.Item("UserID").ToString
                Else
                    Me.p_Member_member_id.Value = ""
                End If
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Items_Bind
            Me.Total_Bind
            Me.Member_Show
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Total_Bind()
            Me.Total_Repeater.DataSource = Me.Total_CreateDataSource
            Me.Total_Repeater.DataBind
        End Sub

        Private Function Total_CreateDataSource() As ICollection
            Me.Total_sSQL = ""
            Me.Total_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            Dim flag As Boolean = True
            Dim flag2 As Boolean = False
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("UserID") Then
                Dim str3 As String
                Dim obj2 As Object = Me.Session.Item("UserID")
                If (obj2 Is Nothing) Then
                    str3 = ""
                Else
                    str3 = obj2.ToString
                End If
                If (Me.Utility.IsNumeric(Nothing, str3) AndAlso (str3.Length > 0)) Then
                    str3 = CCUtility.ToSQL(str3, FieldTypes.Number)
                Else
                    str3 = ""
                End If
                dictionary.Add("UserID", str3)
            End If
            If (dictionary.Item("UserID").Length > 0) Then
                flag2 = True
                str = (str & "[member_id]=" & dictionary.Item("UserID"))
            Else
                flag = False
            End If
            If flag2 Then
                str = (" AND (" & str & ")")
            End If
            Me.Total_sSQL = "SELECT member_id, sum(quantity*price) as sub_total FROM items, orders WHERE orders.item_id=items.item_id"
            str2 = " GROUP BY member_id"
            Me.Total_sSQL = (Me.Total_sSQL & str & str2)
            If Not flag Then
                Me.Total_no_records.Visible = True
                Return Nothing
            End If
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
        Protected Footer As Footer
        Protected Header As Header
        Protected i_Items_curpage As Integer = 1
        Protected i_Total_curpage As Integer = 1
        Protected Items_Column_item_id As Label
        Protected Items_Column_order_id As Label
        Protected Items_Column_price As Label
        Protected Items_Column_quantity As Label
        Protected Items_Column_sub_total As Label
        Protected Items_CountPage As Integer
        Protected Items_FormAction As String = ".aspx?"
        Protected Items_holder As HtmlTable
        Protected Items_no_records As HtmlTableRow
        Private Const Items_PAGENUM As Integer = 20
        Protected Items_Repeater As Repeater
        Protected Items_sCountSQL As String
        Protected Items_sSQL As String
        Protected ItemsForm_Title As Label
        Protected Member_address As Label
        Protected Member_email As Label
        Protected Member_FormAction As String = "AdminMenu.aspx?"
        Protected Member_holder As HtmlTable
        Protected Member_last_name As Label
        Protected Member_member_id As HtmlInputHidden
        Protected Member_member_login As HyperLink
        Protected Member_name As Label
        Protected Member_phone As Label
        Protected Member_ValidationSummary As Label
        Protected MemberForm_Title As Label
        Protected p_Items_order_id As HtmlInputHidden
        Protected p_Member_member_id As HtmlInputHidden
        Protected Total_Column_sub_total As Label
        Protected Total_CountPage As Integer
        Protected Total_FormAction As String = "Default.aspx?"
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

