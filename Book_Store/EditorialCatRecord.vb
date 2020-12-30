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
    Public Class EditorialCatRecord
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub editorial_categories_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.editorial_categories_insert_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.editorial_categories_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("delete") > 0) Then
                flag = Me.editorial_categories_delete_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.editorial_categories_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.editorial_categories_FormAction & ""))
            End If
        End Sub

        Private Sub editorial_categories_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function editorial_categories_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function editorial_categories_delete_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            If (Me.p_editorial_categories_editorial_cat_id.Value.Length > 0) Then
                str = (str & "editorial_cat_id=" & CCUtility.ToSQL(Me.p_editorial_categories_editorial_cat_id.Value, FieldTypes.Number))
            End If
            Dim sQL As String = ("delete from editorial_categories where " & str)
            Me.editorial_categories_BeforeSQLExecute(sQL, "Delete")
            Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
            Try 
                command.ExecuteNonQuery
            Catch exception As Exception
                Me.editorial_categories_ValidationSummary.Text = (Me.editorial_categories_ValidationSummary.Text & exception.Message)
                Me.editorial_categories_ValidationSummary.Visible = True
                Return False
            End Try
            Return True
        End Function

        Private Function editorial_categories_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.editorial_categories_Validate
            Dim str2 As String = CCUtility.ToSQL(Me.Utility.GetParam("editorial_categories_editorial_cat_name"), FieldTypes.Text)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = ("insert into editorial_categories (editorial_cat_name) values (" & str2 & ")")
                End If
                Me.editorial_categories_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.editorial_categories_ValidationSummary.Text = (Me.editorial_categories_ValidationSummary.Text & exception.Message)
                    Me.editorial_categories_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub editorial_categories_Show()
            Dim flag As Boolean = True
            If (Me.p_editorial_categories_editorial_cat_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "editorial_cat_id=" & CCUtility.ToSQL(Me.p_editorial_categories_editorial_cat_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from editorial_categories where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "editorial_categories") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.editorial_categories_editorial_cat_id.Value = CCUtility.GetValue(row, "editorial_cat_id")
                    Me.editorial_categories_editorial_cat_name.Text = CCUtility.GetValue(row, "editorial_cat_name")
                    Me.editorial_categories_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("editorial_cat_id")
                Me.editorial_categories_editorial_cat_id.Value = param
                Me.editorial_categories_delete.Visible = False
                Me.editorial_categories_update.Visible = False
            End If
        End Sub

        Private Function editorial_categories_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.editorial_categories_Validate
            If flag Then
                If (Me.p_editorial_categories_editorial_cat_id.Value.Length > 0) Then
                    str = (str & "editorial_cat_id=" & CCUtility.ToSQL(Me.p_editorial_categories_editorial_cat_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (("update editorial_categories set [editorial_cat_name]=" & CCUtility.ToSQL(Me.Utility.GetParam("editorial_categories_editorial_cat_name"), FieldTypes.Text)) & " where " & str)
                Me.editorial_categories_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.editorial_categories_ValidationSummary.Text = (Me.editorial_categories_ValidationSummary.Text & exception.Message)
                    Me.editorial_categories_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function editorial_categories_Validate() As Boolean
            Dim flag As Boolean = True
            Me.editorial_categories_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("editorial_categories") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.editorial_categories_ValidationSummary.Text = (Me.editorial_categories_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.editorial_categories_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.editorial_categories_insert.ServerClick, New EventHandler(AddressOf Me.editorial_categories_Action)
            AddHandler Me.editorial_categories_update.ServerClick, New EventHandler(AddressOf Me.editorial_categories_Action)
            AddHandler Me.editorial_categories_delete.ServerClick, New EventHandler(AddressOf Me.editorial_categories_Action)
            AddHandler Me.editorial_categories_cancel.ServerClick, New EventHandler(AddressOf Me.editorial_categories_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_editorial_categories_editorial_cat_id.Value = Me.Utility.GetParam("editorial_cat_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.editorial_categories_Show
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
        Protected editorial_categories_cancel As HtmlInputButton
        Protected editorial_categories_delete As HtmlInputButton
        Protected editorial_categories_editorial_cat_id As HtmlInputHidden
        Protected editorial_categories_editorial_cat_name As TextBox
        Protected editorial_categories_FormAction As String = "EditorialCatGrid.aspx?"
        Protected editorial_categories_holder As HtmlTable
        Protected editorial_categories_insert As HtmlInputButton
        Protected editorial_categories_update As HtmlInputButton
        Protected editorial_categories_ValidationSummary As Label
        Protected editorial_categoriesForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected p_editorial_categories_editorial_cat_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

