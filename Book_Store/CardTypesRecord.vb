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
    Public Class CardTypesRecord
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub CardTypes_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.CardTypes_insert_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.CardTypes_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("delete") > 0) Then
                flag = Me.CardTypes_delete_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.CardTypes_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.CardTypes_FormAction & ""))
            End If
        End Sub

        Private Sub CardTypes_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function CardTypes_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function CardTypes_delete_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            If (Me.p_CardTypes_card_type_id.Value.Length > 0) Then
                str = (str & "card_type_id=" & CCUtility.ToSQL(Me.p_CardTypes_card_type_id.Value, FieldTypes.Number))
            End If
            Dim sQL As String = ("delete from card_types where " & str)
            Me.CardTypes_BeforeSQLExecute(sQL, "Delete")
            Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
            Try 
                command.ExecuteNonQuery
            Catch exception As Exception
                Me.CardTypes_ValidationSummary.Text = (Me.CardTypes_ValidationSummary.Text & exception.Message)
                Me.CardTypes_ValidationSummary.Visible = True
                Return False
            End Try
            Return True
        End Function

        Private Function CardTypes_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.CardTypes_Validate
            Dim str2 As String = CCUtility.ToSQL(Me.Utility.GetParam("CardTypes_name"), FieldTypes.Text)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = ("insert into card_types (name) values (" & str2 & ")")
                End If
                Me.CardTypes_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.CardTypes_ValidationSummary.Text = (Me.CardTypes_ValidationSummary.Text & exception.Message)
                    Me.CardTypes_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub CardTypes_Show()
            Dim flag As Boolean = True
            If (Me.p_CardTypes_card_type_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "card_type_id=" & CCUtility.ToSQL(Me.p_CardTypes_card_type_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from card_types where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "CardTypes") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.CardTypes_card_type_id.Value = CCUtility.GetValue(row, "card_type_id")
                    Me.CardTypes_name.Text = CCUtility.GetValue(row, "name")
                    Me.CardTypes_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("card_type_id")
                Me.CardTypes_card_type_id.Value = param
                Me.CardTypes_delete.Visible = False
                Me.CardTypes_update.Visible = False
            End If
        End Sub

        Private Function CardTypes_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.CardTypes_Validate
            If flag Then
                If (Me.p_CardTypes_card_type_id.Value.Length > 0) Then
                    str = (str & "card_type_id=" & CCUtility.ToSQL(Me.p_CardTypes_card_type_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (("update card_types set [name]=" & CCUtility.ToSQL(Me.Utility.GetParam("CardTypes_name"), FieldTypes.Text)) & " where " & str)
                Me.CardTypes_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.CardTypes_ValidationSummary.Text = (Me.CardTypes_ValidationSummary.Text & exception.Message)
                    Me.CardTypes_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function CardTypes_Validate() As Boolean
            Dim flag As Boolean = True
            Me.CardTypes_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("CardTypes") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.CardTypes_ValidationSummary.Text = (Me.CardTypes_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.CardTypes_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.CardTypes_insert.ServerClick, New EventHandler(AddressOf Me.CardTypes_Action)
            AddHandler Me.CardTypes_update.ServerClick, New EventHandler(AddressOf Me.CardTypes_Action)
            AddHandler Me.CardTypes_delete.ServerClick, New EventHandler(AddressOf Me.CardTypes_Action)
            AddHandler Me.CardTypes_cancel.ServerClick, New EventHandler(AddressOf Me.CardTypes_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_CardTypes_card_type_id.Value = Me.Utility.GetParam("card_type_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.CardTypes_Show
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
        Protected CardTypes_cancel As HtmlInputButton
        Protected CardTypes_card_type_id As HtmlInputHidden
        Protected CardTypes_delete As HtmlInputButton
        Protected CardTypes_FormAction As String = "CardTypesGrid.aspx?"
        Protected CardTypes_holder As HtmlTable
        Protected CardTypes_insert As HtmlInputButton
        Protected CardTypes_name As TextBox
        Protected CardTypes_name_Validator_Req As RequiredFieldValidator
        Protected CardTypes_update As HtmlInputButton
        Protected CardTypes_ValidationSummary As Label
        Protected CardTypesForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected p_CardTypes_card_type_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

