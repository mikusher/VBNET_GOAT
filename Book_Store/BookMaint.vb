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
    Public Class BookMaint
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub Book_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.Book_insert_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.Book_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("delete") > 0) Then
                flag = Me.Book_delete_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.Book_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.Book_FormAction & "category_id=" & MyBase.Server.UrlEncode(Me.Utility.GetParam("category_id")) & "&"))
            End If
        End Sub

        Private Sub Book_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function Book_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function Book_delete_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            If (Me.p_Book_item_id.Value.Length > 0) Then
                str = (str & "item_id=" & CCUtility.ToSQL(Me.p_Book_item_id.Value, FieldTypes.Number))
            End If
            Dim sQL As String = ("delete from items where " & str)
            Me.Book_BeforeSQLExecute(sQL, "Delete")
            Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
            Try 
                command.ExecuteNonQuery
            Catch exception As Exception
                Me.Book_ValidationSummary.Text = (Me.Book_ValidationSummary.Text & exception.Message)
                Me.Book_ValidationSummary.Visible = True
                Return False
            End Try
            Return True
        End Function

        Private Function Book_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Book_Validate
            Dim str2 As String = CCUtility.ToSQL(Me.Utility.GetParam("Book_name"), FieldTypes.Text)
            Dim str3 As String = CCUtility.ToSQL(Me.Utility.GetParam("Book_author"), FieldTypes.Text)
            Dim str4 As String = CCUtility.ToSQL(Me.Utility.GetParam("Book_category_id"), FieldTypes.Number)
            Dim str5 As String = CCUtility.ToSQL(Me.Utility.GetParam("Book_price"), FieldTypes.Number)
            Dim str6 As String = CCUtility.ToSQL(Me.Utility.GetParam("Book_product_url"), FieldTypes.Text)
            Dim str7 As String = CCUtility.ToSQL(Me.Utility.GetParam("Book_image_url"), FieldTypes.Text)
            Dim str8 As String = CCUtility.ToSQL(Me.Utility.GetParam("Book_notes"), FieldTypes.Text)
            Dim str9 As String = CCUtility.getCheckBoxValue(Me.Utility.GetParam("Book_is_recommended"), "1", "0", FieldTypes.Number)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = String.Concat(New String() { "insert into items (name,author,category_id,price,product_url,image_url,notes,is_recommended) values (", str2, ",", str3, ",", str4, ",", str5, ",", str6, ",", str7, ",", str8, ",", str9, ")" })
                End If
                Me.Book_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Book_ValidationSummary.Text = (Me.Book_ValidationSummary.Text & exception.Message)
                    Me.Book_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub Book_Show()
            Me.Utility.buildListBox(Me.Book_category_id.Items, "select category_id,name from categories order by 2", "category_id", "name", Nothing, "")
            Dim flag As Boolean = True
            If (Me.p_Book_item_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "item_id=" & CCUtility.ToSQL(Me.p_Book_item_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from items where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Book") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Book_item_id.Value = CCUtility.GetValue(row, "item_id")
                    Me.Book_name.Text = CCUtility.GetValue(row, "name")
                    Me.Book_author.Text = CCUtility.GetValue(row, "author")
                    Dim str3 As String = CCUtility.GetValue(row, "category_id")
                    Try 
                        Me.Book_category_id.SelectedIndex = Me.Book_category_id.Items.IndexOf(Me.Book_category_id.Items.FindByValue(str3))
                    Catch obj1 As Object
                    End Try
                    Me.Book_price.Text = CCUtility.GetValue(row, "price")
                    Me.Book_product_url.Text = CCUtility.GetValue(row, "product_url")
                    Me.Book_image_url.Text = CCUtility.GetValue(row, "image_url")
                    Me.Book_notes.Text = CCUtility.GetValue(row, "notes")
                    Me.Book_is_recommended.Checked = CCUtility.GetValue(row, "is_recommended").ToLower.Equals("1".ToLower)
                    Me.Book_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("item_id")
                Me.Book_item_id.Value = param
                param = Me.Utility.GetParam("category_id")
                Try 
                    Me.Book_category_id.SelectedIndex = Me.Book_category_id.Items.IndexOf(Me.Book_category_id.Items.FindByValue(param))
                Catch obj2 As Object
                End Try
                Me.Book_delete.Visible = False
                Me.Book_update.Visible = False
            End If
        End Sub

        Private Function Book_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Book_Validate
            If flag Then
                If (Me.p_Book_item_id.Value.Length > 0) Then
                    str = (str & "item_id=" & CCUtility.ToSQL(Me.p_Book_item_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (String.Concat(New String() { "update items set [name]=", CCUtility.ToSQL(Me.Utility.GetParam("Book_name"), FieldTypes.Text), ",[author]=", CCUtility.ToSQL(Me.Utility.GetParam("Book_author"), FieldTypes.Text), ",[category_id]=", CCUtility.ToSQL(Me.Utility.GetParam("Book_category_id"), FieldTypes.Number), ",[price]=", CCUtility.ToSQL(Me.Utility.GetParam("Book_price"), FieldTypes.Number), ",[product_url]=", CCUtility.ToSQL(Me.Utility.GetParam("Book_product_url"), FieldTypes.Text), ",[image_url]=", CCUtility.ToSQL(Me.Utility.GetParam("Book_image_url"), FieldTypes.Text), ",[notes]=", CCUtility.ToSQL(Me.Utility.GetParam("Book_notes"), FieldTypes.Text), ",[is_recommended]=", CCUtility.getCheckBoxValue(Me.Utility.GetParam("Book_is_recommended"), "1", "0", FieldTypes.Number) }) & " where " & str)
                Me.Book_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Book_ValidationSummary.Text = (Me.Book_ValidationSummary.Text & exception.Message)
                    Me.Book_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function Book_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Book_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Book") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Book_ValidationSummary.Text = (Me.Book_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Book_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Book_insert.ServerClick, New EventHandler(AddressOf Me.Book_Action)
            AddHandler Me.Book_update.ServerClick, New EventHandler(AddressOf Me.Book_Action)
            AddHandler Me.Book_delete.ServerClick, New EventHandler(AddressOf Me.Book_Action)
            AddHandler Me.Book_cancel.ServerClick, New EventHandler(AddressOf Me.Book_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_Book_item_id.Value = Me.Utility.GetParam("item_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Book_Show
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
        Protected Book_author As TextBox
        Protected Book_cancel As HtmlInputButton
        Protected Book_category_id As DropDownList
        Protected Book_category_id_Validator_Num As CustomValidator
        Protected Book_category_id_Validator_Req As RequiredFieldValidator
        Protected Book_delete As HtmlInputButton
        Protected Book_FormAction As String = "AdminBooks.aspx?"
        Protected Book_holder As HtmlTable
        Protected Book_image_url As TextBox
        Protected Book_insert As HtmlInputButton
        Protected Book_is_recommended As CheckBox
        Protected Book_item_id As HtmlInputHidden
        Protected Book_name As TextBox
        Protected Book_name_Validator_Req As RequiredFieldValidator
        Protected Book_notes As TextBox
        Protected Book_price As TextBox
        Protected Book_price_Validator_Num As CustomValidator
        Protected Book_price_Validator_Req As RequiredFieldValidator
        Protected Book_product_url As TextBox
        Protected Book_update As HtmlInputButton
        Protected Book_ValidationSummary As Label
        Protected BookForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected p_Book_item_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

