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
    Public Class EditorialsRecord
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub editorials_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.editorials_insert_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.editorials_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("delete") > 0) Then
                flag = Me.editorials_delete_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.editorials_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.editorials_FormAction & ""))
            End If
        End Sub

        Private Sub editorials_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function editorials_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function editorials_delete_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            If (Me.p_editorials_article_id.Value.Length > 0) Then
                str = (str & "article_id=" & CCUtility.ToSQL(Me.p_editorials_article_id.Value, FieldTypes.Number))
            End If
            Dim sQL As String = ("delete from editorials where " & str)
            Me.editorials_BeforeSQLExecute(sQL, "Delete")
            Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
            Try 
                command.ExecuteNonQuery
            Catch exception As Exception
                Me.editorials_ValidationSummary.Text = (Me.editorials_ValidationSummary.Text & exception.Message)
                Me.editorials_ValidationSummary.Visible = True
                Return False
            End Try
            Return True
        End Function

        Private Function editorials_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.editorials_Validate
            Dim str2 As String = CCUtility.ToSQL(Me.Utility.GetParam("editorials_article_desc"), FieldTypes.Text)
            Dim str3 As String = CCUtility.ToSQL(Me.Utility.GetParam("editorials_article_title"), FieldTypes.Text)
            Dim str4 As String = CCUtility.ToSQL(Me.Utility.GetParam("editorials_editorial_cat_id"), FieldTypes.Number)
            Dim str5 As String = CCUtility.ToSQL(Me.Utility.GetParam("editorials_item_id"), FieldTypes.Number)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = String.Concat(New String() { "insert into editorials (article_desc,article_title,editorial_cat_id,item_id) values (", str2, ",", str3, ",", str4, ",", str5, ")" })
                End If
                Me.editorials_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.editorials_ValidationSummary.Text = (Me.editorials_ValidationSummary.Text & exception.Message)
                    Me.editorials_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub editorials_Show()
            Me.Utility.buildListBox(Me.editorials_editorial_cat_id.Items, "select editorial_cat_id,editorial_cat_name from editorial_categories order by 2", "editorial_cat_id", "editorial_cat_name", Nothing, "")
            Me.Utility.buildListBox(Me.editorials_item_id.Items, "select item_id,name from items order by 2", "item_id", "name", "", "")
            Dim flag As Boolean = True
            If (Me.p_editorials_article_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "article_id=" & CCUtility.ToSQL(Me.p_editorials_article_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from editorials where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "editorials") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.editorials_article_id.Value = CCUtility.GetValue(row, "article_id")
                    Me.editorials_article_desc.Text = CCUtility.GetValue(row, "article_desc")
                    Me.editorials_article_title.Text = CCUtility.GetValue(row, "article_title")
                    Dim str3 As String = CCUtility.GetValue(row, "editorial_cat_id")
                    Try 
                        Me.editorials_editorial_cat_id.SelectedIndex = Me.editorials_editorial_cat_id.Items.IndexOf(Me.editorials_editorial_cat_id.Items.FindByValue(str3))
                    Catch obj1 As Object
                    End Try
                    Dim str4 As String = CCUtility.GetValue(row, "item_id")
                    Try 
                        Me.editorials_item_id.SelectedIndex = Me.editorials_item_id.Items.IndexOf(Me.editorials_item_id.Items.FindByValue(str4))
                    Catch obj2 As Object
                    End Try
                    Me.editorials_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("article_id")
                Me.editorials_article_id.Value = param
                Me.editorials_delete.Visible = False
                Me.editorials_update.Visible = False
            End If
        End Sub

        Private Function editorials_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.editorials_Validate
            If flag Then
                If (Me.p_editorials_article_id.Value.Length > 0) Then
                    str = (str & "article_id=" & CCUtility.ToSQL(Me.p_editorials_article_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (String.Concat(New String() { "update editorials set [article_desc]=", CCUtility.ToSQL(Me.Utility.GetParam("editorials_article_desc"), FieldTypes.Text), ",[article_title]=", CCUtility.ToSQL(Me.Utility.GetParam("editorials_article_title"), FieldTypes.Text), ",[editorial_cat_id]=", CCUtility.ToSQL(Me.Utility.GetParam("editorials_editorial_cat_id"), FieldTypes.Number), ",[item_id]=", CCUtility.ToSQL(Me.Utility.GetParam("editorials_item_id"), FieldTypes.Number) }) & " where " & str)
                Me.editorials_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.editorials_ValidationSummary.Text = (Me.editorials_ValidationSummary.Text & exception.Message)
                    Me.editorials_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function editorials_Validate() As Boolean
            Dim flag As Boolean = True
            Me.editorials_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("editorials") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.editorials_ValidationSummary.Text = (Me.editorials_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.editorials_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.editorials_insert.ServerClick, New EventHandler(AddressOf Me.editorials_Action)
            AddHandler Me.editorials_update.ServerClick, New EventHandler(AddressOf Me.editorials_Action)
            AddHandler Me.editorials_delete.ServerClick, New EventHandler(AddressOf Me.editorials_Action)
            AddHandler Me.editorials_cancel.ServerClick, New EventHandler(AddressOf Me.editorials_Action)
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
            Me.editorials_Show
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
        Protected editorials_article_desc As TextBox
        Protected editorials_article_id As HtmlInputHidden
        Protected editorials_article_title As TextBox
        Protected editorials_cancel As HtmlInputButton
        Protected editorials_delete As HtmlInputButton
        Protected editorials_editorial_cat_id As DropDownList
        Protected editorials_editorial_cat_id_Validator_Num As CustomValidator
        Protected editorials_editorial_cat_id_Validator_Req As RequiredFieldValidator
        Protected editorials_FormAction As String = "EditorialsGrid.aspx?"
        Protected editorials_holder As HtmlTable
        Protected editorials_insert As HtmlInputButton
        Protected editorials_item_id As DropDownList
        Protected editorials_item_id_Validator_Num As CustomValidator
        Protected editorials_update As HtmlInputButton
        Protected editorials_ValidationSummary As Label
        Protected editorialsForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected p_editorials_article_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

