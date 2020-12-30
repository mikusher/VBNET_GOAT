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
    Public Class Registration
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
            AddHandler Me.Reg_insert.ServerClick, New EventHandler(AddressOf Me.Reg_Action)
            AddHandler Me.Reg_update.ServerClick, New EventHandler(AddressOf Me.Reg_Action)
            AddHandler Me.Reg_cancel.ServerClick, New EventHandler(AddressOf Me.Reg_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            If Not MyBase.IsPostBack Then
                If (Not Me.Session.Item("UserID") Is Nothing) Then
                    Me.p_Reg_member_id.Value = Me.Session.Item("UserID").ToString
                Else
                    Me.p_Reg_member_id.Value = ""
                End If
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Reg_Show
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Reg_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.Reg_insert_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.Reg_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.Reg_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.Reg_FormAction & ""))
            End If
        End Sub

        Private Sub Reg_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function Reg_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Function Reg_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Reg_Validate
            If (Me.Utility.DlookupInt("members", "count(*)", ("member_login=" & CCUtility.ToSQL(Me.Utility.GetParam("Reg_member_login"), FieldTypes.Text))) <> 0) Then
                Me.Reg_ValidationSummary.Visible = True
                Me.Reg_ValidationSummary.Text = (Me.Reg_ValidationSummary.Text & "The value in field Login* is already in database.<br>")
                flag = False
            End If
            If (Me.Reg_member_password.Text <> Me.Reg_member_password2.Text) Then
                Me.Reg_ValidationSummary.Text = (Me.Reg_ValidationSummary.Text & "Password and Confirm Password fields don't match<br>")
                Me.Reg_ValidationSummary.Visible = True
                flag = False
            End If
            Dim str2 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_member_login"), FieldTypes.Text)
            Dim str3 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_member_password"), FieldTypes.Text)
            Dim str4 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_first_name"), FieldTypes.Text)
            Dim str5 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_last_name"), FieldTypes.Text)
            Dim str6 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_email"), FieldTypes.Text)
            Dim str7 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_address"), FieldTypes.Text)
            Dim str8 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_phone"), FieldTypes.Text)
            Dim str9 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_card_type_id"), FieldTypes.Number)
            Dim str10 As String = CCUtility.ToSQL(Me.Utility.GetParam("Reg_card_number"), FieldTypes.Text)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = String.Concat(New String() { "insert into members (member_login,member_password,first_name,last_name,email,address,phone,card_type_id,card_number) values (", str2, ",", str3, ",", str4, ",", str5, ",", str6, ",", str7, ",", str8, ",", str9, ",", str10, ")" })
                End If
                Me.Reg_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Reg_ValidationSummary.Text = (Me.Reg_ValidationSummary.Text & exception.Message)
                    Me.Reg_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub Reg_Show()
            Me.Utility.buildListBox(Me.Reg_card_type_id.Items, "select card_type_id,name from card_types order by 2", "card_type_id", "name", "", "")
            Dim flag As Boolean = True
            If (Me.p_Reg_member_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Reg_member_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from members where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Reg") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Reg_member_id.Value = CCUtility.GetValue(row, "member_id")
                    Me.Reg_member_login.Text = CCUtility.GetValue(row, "member_login")
                    Me.Reg_member_password.Text = CCUtility.GetValue(row, "member_password")
                    Me.Reg_first_name.Text = CCUtility.GetValue(row, "first_name")
                    Me.Reg_last_name.Text = CCUtility.GetValue(row, "last_name")
                    Me.Reg_email.Text = CCUtility.GetValue(row, "email")
                    Me.Reg_address.Text = CCUtility.GetValue(row, "address")
                    Me.Reg_phone.Text = CCUtility.GetValue(row, "phone")
                    Dim str3 As String = CCUtility.GetValue(row, "card_type_id")
                    Try 
                        Me.Reg_card_type_id.SelectedIndex = Me.Reg_card_type_id.Items.IndexOf(Me.Reg_card_type_id.Items.FindByValue(str3))
                    Catch obj1 As Object
                    End Try
                    Me.Reg_card_number.Text = CCUtility.GetValue(row, "card_number")
                    Me.Reg_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim str4 As String
                If (Not Me.Session.Item("UserID") Is Nothing) Then
                    str4 = Me.Session.Item("UserID").ToString
                Else
                    str4 = ""
                End If
                Me.Reg_member_id.Value = str4
                Me.Reg_update.Visible = False
            End If
        End Sub

        Private Function Reg_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Reg_Validate
            If flag Then
                If (Me.p_Reg_member_id.Value.Length > 0) Then
                    str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Reg_member_id.Value, FieldTypes.Number))
                End If
                If (Me.Utility.DlookupInt("members", "count(*)", String.Concat(New String() { "member_login=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_member_login"), FieldTypes.Text), " and not(", str, ")" })) <> 0) Then
                    Me.Reg_ValidationSummary.Visible = True
                    Me.Reg_ValidationSummary.Text = (Me.Reg_ValidationSummary.Text & "The value in field Login* is already in database.<br>")
                    flag = False
                End If
                If (Me.Reg_member_password.Text <> Me.Reg_member_password2.Text) Then
                    Me.Reg_ValidationSummary.Text = (Me.Reg_ValidationSummary.Text & "Password and Confirm Password fields don't match<br>")
                    Me.Reg_ValidationSummary.Visible = True
                    flag = False
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (String.Concat(New String() { "update members set [member_login]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_member_login"), FieldTypes.Text), ",[member_password]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_member_password"), FieldTypes.Text), ",[first_name]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_first_name"), FieldTypes.Text), ",[last_name]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_last_name"), FieldTypes.Text), ",[email]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_email"), FieldTypes.Text), ",[address]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_address"), FieldTypes.Text), ",[phone]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_phone"), FieldTypes.Text), ",[card_type_id]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_card_type_id"), FieldTypes.Number), ",[card_number]=", CCUtility.ToSQL(Me.Utility.GetParam("Reg_card_number"), FieldTypes.Text) }) & " where " & str)
                Me.Reg_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Reg_ValidationSummary.Text = (Me.Reg_ValidationSummary.Text & exception.Message)
                    Me.Reg_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function Reg_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Reg_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Reg") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Reg_ValidationSummary.Text = (Me.Reg_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Reg_ValidationSummary.Visible = Not flag
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
        Protected p_Reg_member_id As HtmlInputHidden
        Protected Reg_address As TextBox
        Protected Reg_cancel As HtmlInputButton
        Protected Reg_card_number As TextBox
        Protected Reg_card_type_id As DropDownList
        Protected Reg_card_type_id_Validator_Num As CustomValidator
        Protected Reg_email As TextBox
        Protected Reg_email_Validator_Req As RequiredFieldValidator
        Protected Reg_first_name As TextBox
        Protected Reg_first_name_Validator_Req As RequiredFieldValidator
        Protected Reg_FormAction As String = "Default.aspx?"
        Protected Reg_holder As HtmlTable
        Protected Reg_insert As HtmlInputButton
        Protected Reg_last_name As TextBox
        Protected Reg_last_name_Validator_Req As RequiredFieldValidator
        Protected Reg_member_id As HtmlInputHidden
        Protected Reg_member_login As TextBox
        Protected Reg_member_login_Validator_Req As RequiredFieldValidator
        Protected Reg_member_password As TextBox
        Protected Reg_member_password_Validator_Req As RequiredFieldValidator
        Protected Reg_member_password2 As TextBox
        Protected Reg_member_password2_Validator_Req As RequiredFieldValidator
        Protected Reg_phone As TextBox
        Protected Reg_update As HtmlInputButton
        Protected Reg_ValidationSummary As Label
        Protected RegForm_Title As Label
        Protected Utility As CCUtility
    End Class
End Namespace

