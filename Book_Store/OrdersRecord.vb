Imports ASP
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.Profile
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls

Namespace Book_Store
    Public Class OrdersRecord
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub Orders_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.Orders_insert_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.Orders_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("delete") > 0) Then
                flag = Me.Orders_delete_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.Orders_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect(String.Concat(New String() { Me.Orders_FormAction, "item_id=", MyBase.Server.UrlEncode(Me.Utility.GetParam("item_id")), "&member_id=", MyBase.Server.UrlEncode(Me.Utility.GetParam("member_id")), "&" }))
            End If
        End Sub

        Private Sub Orders_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function Orders_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function Orders_delete_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            If (Me.p_Orders_order_id.Value.Length > 0) Then
                str = (str & "order_id=" & CCUtility.ToSQL(Me.p_Orders_order_id.Value, FieldTypes.Number))
            End If
            Dim sQL As String = ("delete from orders where " & str)
            Me.Orders_BeforeSQLExecute(sQL, "Delete")
            Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
            Try 
                command.ExecuteNonQuery
            Catch exception As Exception
                Me.Orders_ValidationSummary.Text = (Me.Orders_ValidationSummary.Text & exception.Message)
                Me.Orders_ValidationSummary.Visible = True
                Return False
            End Try
            Return True
        End Function

        Private Function Orders_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Orders_Validate
            Dim str2 As String = CCUtility.ToSQL(Me.Utility.GetParam("Orders_member_id"), FieldTypes.Number)
            Dim str3 As String = CCUtility.ToSQL(Me.Utility.GetParam("Orders_item_id"), FieldTypes.Number)
            Dim str4 As String = CCUtility.ToSQL(Me.Utility.GetParam("Orders_quantity"), FieldTypes.Number)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = String.Concat(New String() { "insert into orders (member_id,item_id,quantity) values (", str2, ",", str3, ",", str4, ")" })
                End If
                Me.Orders_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Orders_ValidationSummary.Text = (Me.Orders_ValidationSummary.Text & exception.Message)
                    Me.Orders_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub Orders_Show()
            Me.Utility.buildListBox(Me.Orders_member_id.Items, "select member_id,member_login from members order by 2", "member_id", "member_login", Nothing, "")
            Me.Utility.buildListBox(Me.Orders_item_id.Items, "select item_id,name from items order by 2", "item_id", "name", Nothing, "")
            Dim flag As Boolean = True
            If (Me.p_Orders_order_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "order_id=" & CCUtility.ToSQL(Me.p_Orders_order_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from orders where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Orders") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Orders_order_id.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "order_id").ToString)
                    Dim str3 As String = CCUtility.GetValue(row, "member_id")
                    Try 
                        Me.Orders_member_id.SelectedIndex = Me.Orders_member_id.Items.IndexOf(Me.Orders_member_id.Items.FindByValue(str3))
                    Catch obj1 As Object
                    End Try
                    Dim str4 As String = CCUtility.GetValue(row, "item_id")
                    Try 
                        Me.Orders_item_id.SelectedIndex = Me.Orders_item_id.Items.IndexOf(Me.Orders_item_id.Items.FindByValue(str4))
                    Catch obj2 As Object
                    End Try
                    Me.Orders_quantity.Text = CCUtility.GetValue(row, "quantity")
                    Me.Orders_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("order_id")
                Me.Orders_order_id.Text = param
                param = Me.Utility.GetParam("member_id")
                Try 
                    Me.Orders_member_id.SelectedIndex = Me.Orders_member_id.Items.IndexOf(Me.Orders_member_id.Items.FindByValue(param))
                Catch obj3 As Object
                End Try
                param = Me.Utility.GetParam("item_id")
                Try 
                    Me.Orders_item_id.SelectedIndex = Me.Orders_item_id.Items.IndexOf(Me.Orders_item_id.Items.FindByValue(param))
                Catch obj4 As Object
                End Try
                Me.Orders_delete.Visible = False
                Me.Orders_update.Visible = False
            End If
        End Sub

        Private Function Orders_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Orders_Validate
            If flag Then
                If (Me.p_Orders_order_id.Value.Length > 0) Then
                    str = (str & "order_id=" & CCUtility.ToSQL(Me.p_Orders_order_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (String.Concat(New String() { "update orders set [member_id]=", CCUtility.ToSQL(Me.Utility.GetParam("Orders_member_id"), FieldTypes.Number), ",[item_id]=", CCUtility.ToSQL(Me.Utility.GetParam("Orders_item_id"), FieldTypes.Number), ",[quantity]=", CCUtility.ToSQL(Me.Utility.GetParam("Orders_quantity"), FieldTypes.Number) }) & " where " & str)
                Me.Orders_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Orders_ValidationSummary.Text = (Me.Orders_ValidationSummary.Text & exception.Message)
                    Me.Orders_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function Orders_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Orders_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Orders") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Orders_ValidationSummary.Text = (Me.Orders_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Orders_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Orders_insert.ServerClick, New EventHandler(AddressOf Me.Orders_Action)
            AddHandler Me.Orders_update.ServerClick, New EventHandler(AddressOf Me.Orders_Action)
            AddHandler Me.Orders_delete.ServerClick, New EventHandler(AddressOf Me.Orders_Action)
            AddHandler Me.Orders_cancel.ServerClick, New EventHandler(AddressOf Me.Orders_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_Orders_order_id.Value = Me.Utility.GetParam("order_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Orders_Show
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
        Protected Footer As Footer
        Protected Header As Header
        Protected Orders_cancel As HtmlInputButton
        Protected Orders_delete As HtmlInputButton
        Protected Orders_FormAction As String = "OrdersGrid.aspx?"
        Protected Orders_holder As HtmlTable
        Protected Orders_insert As HtmlInputButton
        Protected Orders_item_id As DropDownList
        Protected Orders_item_id_Validator_Num As CustomValidator
        Protected Orders_item_id_Validator_Req As RequiredFieldValidator
        Protected Orders_member_id As DropDownList
        Protected Orders_member_id_Validator_Num As CustomValidator
        Protected Orders_member_id_Validator_Req As RequiredFieldValidator
        Protected Orders_order_id As Label
        Protected Orders_quantity As TextBox
        Protected Orders_quantity_Validator_Num As CustomValidator
        Protected Orders_quantity_Validator_Req As RequiredFieldValidator
        Protected Orders_update As HtmlInputButton
        Protected Orders_ValidationSummary As Label
        Protected OrdersForm_Title As Label
        Protected p_Orders_order_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

