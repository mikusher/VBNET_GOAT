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
    Public Class ShoppingCartRecord
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.ShoppingCartRecord_update.ServerClick, New EventHandler(AddressOf Me.ShoppingCartRecord_Action)
            AddHandler Me.ShoppingCartRecord_delete.ServerClick, New EventHandler(AddressOf Me.ShoppingCartRecord_Action)
            AddHandler Me.ShoppingCartRecord_cancel.ServerClick, New EventHandler(AddressOf Me.ShoppingCartRecord_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(1)
            If Not MyBase.IsPostBack Then
                Me.p_ShoppingCartRecord_order_id.Value = Me.Utility.GetParam("order_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.ShoppingCartRecord_Show
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub ShoppingCartRecord_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.ShoppingCartRecord_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("delete") > 0) Then
                flag = Me.ShoppingCartRecord_delete_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.ShoppingCartRecord_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.ShoppingCartRecord_FormAction & ""))
            End If
        End Sub

        Private Sub ShoppingCartRecord_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function ShoppingCartRecord_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function ShoppingCartRecord_delete_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            If (Me.p_ShoppingCartRecord_order_id.Value.Length > 0) Then
                str = (str & "order_id=" & CCUtility.ToSQL(Me.p_ShoppingCartRecord_order_id.Value, FieldTypes.Number))
            End If
            Dim sQL As String = ("delete from orders where " & str)
            Me.ShoppingCartRecord_BeforeSQLExecute(sQL, "Delete")
            Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
            Try 
                command.ExecuteNonQuery
            Catch exception As Exception
                Me.ShoppingCartRecord_ValidationSummary.Text = (Me.ShoppingCartRecord_ValidationSummary.Text & exception.Message)
                Me.ShoppingCartRecord_ValidationSummary.Visible = True
                Return False
            End Try
            Return True
        End Function

        Private Sub ShoppingCartRecord_Show()
            Dim flag As Boolean = True
            If (Me.p_ShoppingCartRecord_order_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "order_id=" & CCUtility.ToSQL(Me.p_ShoppingCartRecord_order_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from orders where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "ShoppingCartRecord") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.ShoppingCartRecord_order_id.Value = CCUtility.GetValue(row, "order_id")
                    Me.ShoppingCartRecord_member_id.Value = CCUtility.GetValue(row, "member_id")
                    Me.ShoppingCartRecord_item_id.Text = MyBase.Server.HtmlEncode(Me.Utility.Dlookup("items", "name", ("item_id=" & CCUtility.ToSQL(CCUtility.GetValue(row, "item_id"), FieldTypes.Number))).ToString)
                    Me.ShoppingCartRecord_quantity.Text = CCUtility.GetValue(row, "quantity")
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
                Me.ShoppingCartRecord_member_id.Value = str3
                Me.ShoppingCartRecord_delete.Visible = False
                Me.ShoppingCartRecord_update.Visible = False
            End If
        End Sub

        Private Function ShoppingCartRecord_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.ShoppingCartRecord_Validate
            If flag Then
                If (Me.p_ShoppingCartRecord_order_id.Value.Length > 0) Then
                    str = (str & "order_id=" & CCUtility.ToSQL(Me.p_ShoppingCartRecord_order_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (("update orders set [member_id]=" & CCUtility.ToSQL(Me.Utility.GetParam("ShoppingCartRecord_member_id"), FieldTypes.Number) & ",[quantity]=" & CCUtility.ToSQL(Me.Utility.GetParam("ShoppingCartRecord_quantity"), FieldTypes.Number)) & " where " & str)
                Me.ShoppingCartRecord_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.ShoppingCartRecord_ValidationSummary.Text = (Me.ShoppingCartRecord_ValidationSummary.Text & exception.Message)
                    Me.ShoppingCartRecord_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function ShoppingCartRecord_Validate() As Boolean
            Dim flag As Boolean = True
            Me.ShoppingCartRecord_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("ShoppingCartRecord") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.ShoppingCartRecord_ValidationSummary.Text = (Me.ShoppingCartRecord_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.ShoppingCartRecord_ValidationSummary.Visible = Not flag
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
        Protected p_ShoppingCartRecord_order_id As HtmlInputHidden
        Protected ShoppingCartRecord_cancel As HtmlInputButton
        Protected ShoppingCartRecord_delete As HtmlInputButton
        Protected ShoppingCartRecord_FormAction As String = "ShoppingCart.aspx?"
        Protected ShoppingCartRecord_holder As HtmlTable
        Protected ShoppingCartRecord_item_id As Label
        Protected ShoppingCartRecord_member_id As HtmlInputHidden
        Protected ShoppingCartRecord_order_id As HtmlInputHidden
        Protected ShoppingCartRecord_quantity As TextBox
        Protected ShoppingCartRecord_quantity_Validator_Num As CustomValidator
        Protected ShoppingCartRecord_quantity_Validator_Req As RequiredFieldValidator
        Protected ShoppingCartRecord_update As HtmlInputButton
        Protected ShoppingCartRecord_ValidationSummary As Label
        Protected ShoppingCartRecordForm_Title As Label
        Protected Utility As CCUtility
    End Class
End Namespace

