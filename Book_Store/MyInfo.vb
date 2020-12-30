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
    Public Class MyInfo
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub Form_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.Form_update_Click(Src, E)
            End If
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("cancel") > 0) Then
                flag = Me.Form_cancel_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.Form_FormAction & ""))
            End If
        End Sub

        Private Sub Form_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function Form_cancel_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Return True
        End Function

        Private Sub Form_Show()
            Me.Utility.buildListBox(Me.Form_card_type_id.Items, "select card_type_id,name from card_types order by 2", "card_type_id", "name", "", "")
            Dim flag As Boolean = True
            If (Me.p_Form_member_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Form_member_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from members where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Form") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Form_member_id.Value = CCUtility.GetValue(row, "member_id")
                    Me.Form_member_login.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "member_login").ToString)
                    Me.Form_member_password.Text = CCUtility.GetValue(row, "member_password")
                    Me.Form_name.Text = CCUtility.GetValue(row, "first_name")
                    Me.Form_last_name.Text = CCUtility.GetValue(row, "last_name")
                    Me.Form_email.Text = CCUtility.GetValue(row, "email")
                    Me.Form_address.Text = CCUtility.GetValue(row, "address")
                    Me.Form_phone.Text = CCUtility.GetValue(row, "phone")
                    Me.Form_notes.Text = CCUtility.GetValue(row, "notes")
                    Dim str3 As String = CCUtility.GetValue(row, "card_type_id")
                    Try 
                        Me.Form_card_type_id.SelectedIndex = Me.Form_card_type_id.Items.IndexOf(Me.Form_card_type_id.Items.FindByValue(str3))
                    Catch obj1 As Object
                    End Try
                    Me.Form_card_number.Text = CCUtility.GetValue(row, "card_number")
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
                Me.Form_member_id.Value = str4
                Me.Form_update.Visible = False
            End If
        End Sub

        Private Function Form_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Form_Validate
            If flag Then
                If (Me.p_Form_member_id.Value.Length > 0) Then
                    str = (str & "member_id=" & CCUtility.ToSQL(Me.p_Form_member_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (String.Concat(New String() { "update members set [member_password]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_member_password"), FieldTypes.Text), ",[first_name]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_name"), FieldTypes.Text), ",[last_name]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_last_name"), FieldTypes.Text), ",[email]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_email"), FieldTypes.Text), ",[address]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_address"), FieldTypes.Text), ",[phone]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_phone"), FieldTypes.Text), ",[notes]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_notes"), FieldTypes.Text), ",[card_type_id]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_card_type_id"), FieldTypes.Number), ",[card_number]=", CCUtility.ToSQL(Me.Utility.GetParam("Form_card_number"), FieldTypes.Text) }) & " where " & str)
                Me.Form_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Form_ValidationSummary.Text = (Me.Form_ValidationSummary.Text & exception.Message)
                    Me.Form_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function Form_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Form_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Form") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Form_ValidationSummary.Text = (Me.Form_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Form_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Form_update.ServerClick, New EventHandler(AddressOf Me.Form_Action)
            AddHandler Me.Form_cancel.ServerClick, New EventHandler(AddressOf Me.Form_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(1)
            If Not MyBase.IsPostBack Then
                If (Not Me.Session.Item("UserID") Is Nothing) Then
                    Me.p_Form_member_id.Value = Me.Session.Item("UserID").ToString
                Else
                    Me.p_Form_member_id.Value = ""
                End If
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Form_Show
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
        Protected Form_address As TextBox
        Protected Form_cancel As HtmlInputButton
        Protected Form_card_number As TextBox
        Protected Form_card_type_id As DropDownList
        Protected Form_card_type_id_Validator_Num As CustomValidator
        Protected Form_email As TextBox
        Protected Form_email_Validator_Req As RequiredFieldValidator
        Protected Form_FormAction As String = "ShoppingCart.aspx?"
        Protected Form_holder As HtmlTable
        Protected Form_last_name As TextBox
        Protected Form_last_name_Validator_Req As RequiredFieldValidator
        Protected Form_member_id As HtmlInputHidden
        Protected Form_member_login As Label
        Protected Form_member_password As TextBox
        Protected Form_member_password_Validator_Req As RequiredFieldValidator
        Protected Form_name As TextBox
        Protected Form_name_Validator_Req As RequiredFieldValidator
        Protected Form_notes As TextBox
        Protected Form_phone As TextBox
        Protected Form_update As HtmlInputButton
        Protected Form_ValidationSummary As Label
        Protected FormForm_Title As Label
        Protected Header As Header
        Protected p_Form_member_id As HtmlInputHidden
        Protected Utility As CCUtility
    End Class
End Namespace

