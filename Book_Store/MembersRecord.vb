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
    Public Class MembersRecord
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub Members_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.Members_insert_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.Members_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("delete") > 0) Then
                flag = Me.Members_delete_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.Members_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.Members_FormAction & "member_login=" & MyBase.Server.UrlEncode(Me.Utility.GetParam("member_login")) & "&"))
            End If
        End Sub

        Private Sub Members_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function Members_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function Members_delete_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            If (Me.p_Members_member_id.Value.Length > 0) Then
                str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Members_member_id.Value, FieldTypes.Number))
            End If
            Dim sQL As String = ("delete from members where " & str)
            Me.Members_BeforeSQLExecute(sQL, "Delete")
            Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
            Try 
                command.ExecuteNonQuery
            Catch exception As Exception
                Me.Members_ValidationSummary.Text = (Me.Members_ValidationSummary.Text & exception.Message)
                Me.Members_ValidationSummary.Visible = True
                Return False
            End Try
            Return True
        End Function

        Private Function Members_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Members_Validate
            If (Me.Utility.DlookupInt("members", "count(*)", ("member_login=" & CCUtility.ToSQL(Me.Utility.GetParam("Members_member_login"), FieldTypes.Text))) <> 0) Then
                Me.Members_ValidationSummary.Visible = True
                Me.Members_ValidationSummary.Text = (Me.Members_ValidationSummary.Text & "The value in field Login* is already in database.<br>")
                flag = False
            End If
            Dim str2 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_member_login"), FieldTypes.Text)
            Dim str3 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_member_password"), FieldTypes.Text)
            Dim str4 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_member_level"), FieldTypes.Number)
            Dim str5 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_name"), FieldTypes.Text)
            Dim str6 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_last_name"), FieldTypes.Text)
            Dim str7 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_email"), FieldTypes.Text)
            Dim str8 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_phone"), FieldTypes.Text)
            Dim str9 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_address"), FieldTypes.Text)
            Dim str10 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_notes"), FieldTypes.Text)
            Dim str11 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_card_type_id"), FieldTypes.Number)
            Dim str12 As String = CCUtility.ToSQL(Me.Utility.GetParam("Members_card_number"), FieldTypes.Text)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = String.Concat(New String() { "insert into members (member_login,member_password,member_level,first_name,last_name,email,phone,address,notes,card_type_id,card_number) values (", str2, ",", str3, ",", str4, ",", str5, ",", str6, ",", str7, ",", str8, ",", str9, ",", str10, ",", str11, ",", str12, ")" })
                End If
                Me.Members_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Members_ValidationSummary.Text = (Me.Members_ValidationSummary.Text & exception.Message)
                    Me.Members_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub Members_Show()
            Me.Utility.buildListBox(Me.Members_member_level.Items, Me.Members_member_level_lov, Nothing, "")
            Me.Utility.buildListBox(Me.Members_card_type_id.Items, "select card_type_id,name from card_types order by 2", "card_type_id", "name", "", "")
            Dim flag As Boolean = True
            If (Me.p_Members_member_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Members_member_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from members where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Members") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Members_member_id.Value = CCUtility.GetValue(row, "member_id")
                    Me.Members_member_login.Text = CCUtility.GetValue(row, "member_login")
                    Me.Members_member_password.Text = CCUtility.GetValue(row, "member_password")
                    Dim str3 As String = CCUtility.GetValue(row, "member_level")
                    Try 
                        Me.Members_member_level.SelectedIndex = Me.Members_member_level.Items.IndexOf(Me.Members_member_level.Items.FindByValue(str3))
                    Catch obj1 As Object
                    End Try
                    Me.Members_name.Text = CCUtility.GetValue(row, "first_name")
                    Me.Members_last_name.Text = CCUtility.GetValue(row, "last_name")
                    Me.Members_email.Text = CCUtility.GetValue(row, "email")
                    Me.Members_phone.Text = CCUtility.GetValue(row, "phone")
                    Me.Members_address.Text = CCUtility.GetValue(row, "address")
                    Me.Members_notes.Text = CCUtility.GetValue(row, "notes")
                    Dim str4 As String = CCUtility.GetValue(row, "card_type_id")
                    Try 
                        Me.Members_card_type_id.SelectedIndex = Me.Members_card_type_id.Items.IndexOf(Me.Members_card_type_id.Items.FindByValue(str4))
                    Catch obj2 As Object
                    End Try
                    Me.Members_card_number.Text = CCUtility.GetValue(row, "card_number")
                    Me.Members_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("member_id")
                Me.Members_member_id.Value = param
                param = Me.Utility.GetParam("member_login")
                Me.Members_member_login.Text = param
                Me.Members_delete.Visible = False
                Me.Members_update.Visible = False
            End If
        End Sub

        Private Function Members_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Members_Validate
            If flag Then
                If (Me.p_Members_member_id.Value.Length > 0) Then
                    str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Members_member_id.Value, FieldTypes.Number))
                End If
                If (Me.Utility.DlookupInt("members", "count(*)", String.Concat(New String() { "member_login=", CCUtility.ToSQL(Me.Utility.GetParam("Members_member_login"), FieldTypes.Text), " and not(", str, ")" })) <> 0) Then
                    Me.Members_ValidationSummary.Visible = True
                    Me.Members_ValidationSummary.Text = (Me.Members_ValidationSummary.Text & "The value in field Login* is already in database.<br>")
                    flag = False
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (String.Concat(New String() { "update members set [member_login]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_member_login"), FieldTypes.Text), ",[member_password]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_member_password"), FieldTypes.Text), ",[member_level]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_member_level"), FieldTypes.Number), ",[first_name]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_name"), FieldTypes.Text), ",[last_name]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_last_name"), FieldTypes.Text), ",[email]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_email"), FieldTypes.Text), ",[phone]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_phone"), FieldTypes.Text), ",[address]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_address"), FieldTypes.Text), ",[notes]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_notes"), FieldTypes.Text), ",[card_type_id]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_card_type_id"), FieldTypes.Number), ",[card_number]=", CCUtility.ToSQL(Me.Utility.GetParam("Members_card_number"), FieldTypes.Text) }) & " where " & str)
                Me.Members_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Members_ValidationSummary.Text = (Me.Members_ValidationSummary.Text & exception.Message)
                    Me.Members_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function Members_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Members_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Members") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Members_ValidationSummary.Text = (Me.Members_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Members_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Members_insert.ServerClick, New EventHandler(AddressOf Me.Members_Action)
            AddHandler Me.Members_update.ServerClick, New EventHandler(AddressOf Me.Members_Action)
            AddHandler Me.Members_delete.ServerClick, New EventHandler(AddressOf Me.Members_Action)
            AddHandler Me.Members_cancel.ServerClick, New EventHandler(AddressOf Me.Members_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.p_Members_member_id.Value = Me.Utility.GetParam("member_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Members_Show
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
        Protected Members_address As TextBox
        Protected Members_cancel As HtmlInputButton
        Protected Members_card_number As TextBox
        Protected Members_card_type_id As DropDownList
        Protected Members_card_type_id_Validator_Num As CustomValidator
        Protected Members_delete As HtmlInputButton
        Protected Members_email As TextBox
        Protected Members_email_Validator_Req As RequiredFieldValidator
        Protected Members_FormAction As String = "MembersGrid.aspx?"
        Protected Members_holder As HtmlTable
        Protected Members_insert As HtmlInputButton
        Protected Members_last_name As TextBox
        Protected Members_last_name_Validator_Req As RequiredFieldValidator
        Protected Members_member_id As HtmlInputHidden
        Protected Members_member_level As DropDownList
        Protected Members_member_level_lov As String() = "1;Member;2;Administrator".Split(New Char() { ";"c })
        Protected Members_member_level_Validator_Num As CustomValidator
        Protected Members_member_level_Validator_Req As RequiredFieldValidator
        Protected Members_member_login As TextBox
        Protected Members_member_login_Validator_Req As RequiredFieldValidator
        Protected Members_member_password As TextBox
        Protected Members_member_password_Validator_Req As RequiredFieldValidator
        Protected Members_name As TextBox
        Protected Members_name_Validator_Req As RequiredFieldValidator
        Protected Members_notes As TextBox
        Protected Members_phone As TextBox
        Protected Members_update As HtmlInputButton
        Protected Members_ValidationSummary As Label
        Protected MembersForm_Title As Label
        Protected p_Members_member_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

