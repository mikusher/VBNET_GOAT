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
    Public Class MembersGrid
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub Members_Bind()
            Me.Members_Repeater.DataSource = Me.Members_CreateDataSource
            Me.Members_Repeater.DataBind
        End Sub

        Private Function Members_CreateDataSource() As ICollection
            Me.Members_sSQL = ""
            Me.Members_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            Dim flag As Boolean = False
            str2 = " order by m.member_login Asc"
            If ((Me.Utility.GetParam("FormMembers_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("FormMembers_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            Dim dictionary As New StringDictionary
            If Not dictionary.ContainsKey("name") Then
                Dim param As String = Me.Utility.GetParam("name")
                dictionary.Add("name", param)
            End If
            If Not dictionary.ContainsKey("name") Then
                Dim str4 As String = Me.Utility.GetParam("name")
                dictionary.Add("name", str4)
            End If
            If Not dictionary.ContainsKey("name") Then
                Dim str5 As String = Me.Utility.GetParam("name")
                dictionary.Add("name", str5)
            End If
            If (dictionary.Item("name").Length > 0) Then
                flag = True
                str = String.Concat(New String() { "m.[member_login] like '%", dictionary.Item("name").Replace("'", "''"), "%' or m.[first_name] like '%", dictionary.Item("name").Replace("'", "''"), "%' or m.[last_name] like '%", dictionary.Item("name").Replace("'", "''"), "%'" })
            End If
            If flag Then
                str = (" WHERE (" & str & ")")
            End If
            Me.Members_sSQL = "select [m].[first_name] as m_first_name, [m].[last_name] as m_last_name, [m].[member_id] as m_member_id, [m].[member_level] as m_member_level, [m].[member_login] as m_member_login  from [members] m "
            Me.Members_sSQL = (Me.Members_sSQL & str & str2)
            If (Me.Members_sCountSQL.Length = 0) Then
                Dim index As Integer = Me.Members_sSQL.ToLower.IndexOf("select ")
                Dim num2 As Integer = (Me.Members_sSQL.ToLower.LastIndexOf(" from ") - 1)
                Me.Members_sCountSQL = Me.Members_sSQL.Replace(Me.Members_sSQL.Substring((index + 7), (num2 - 6)), " count(*) ")
                index = Me.Members_sCountSQL.ToLower.IndexOf(" order by")
                If (index > 1) Then
                    Me.Members_sCountSQL = Me.Members_sCountSQL.Substring(0, index)
                End If
            End If
            Dim adapter As New OleDbDataAdapter(Me.Members_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, ((Me.i_Members_curpage - 1) * 20), 20, "Members")
            Dim command As New OleDbCommand(Me.Members_sCountSQL, Me.Utility.Connection)
            Dim num3 As Integer = CInt(command.ExecuteScalar)
            Me.Members_Pager.MaxPage = IIf(((num3 Mod 20) > 0), ((num3 / 20) + 1), (num3 / 20))
            Dim flag2 As Boolean = (Me.Members_Pager.MaxPage <> 1)
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.Members_no_records.Visible = True
                flag2 = False
            Else
                Me.Members_no_records.Visible = False
                flag2 = flag2
            End If
            Me.Members_Pager.Visible = flag2
            Return view
        End Function

        Private Sub Members_insert_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (Me.Members_FormAction & "name=" & MyBase.Server.UrlEncode(Me.Utility.GetParam("name")) & "&")
            MyBase.Response.Redirect(url)
        End Sub

        Protected Sub Members_pager_navigate_completed(ByVal Src As Object, ByVal CurrPage As Integer)
            Me.i_Members_curpage = CurrPage
            Me.Members_Bind
        End Sub

        Public Sub Members_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Protected Sub Members_SortChange(ByVal Src As Object, ByVal E As EventArgs)
            If ((Me.ViewState.Item("SortColumn") Is Nothing) OrElse (Me.ViewState.Item("SortColumn").ToString <> DirectCast(Src, LinkButton).CommandArgument)) Then
                Me.ViewState.Item("SortColumn") = DirectCast(Src, LinkButton).CommandArgument
                Me.ViewState.Item("SortDir") = "ASC"
            Else
                Me.ViewState.Item("SortDir") = IIf((Me.ViewState.Item("SortDir").ToString = "ASC"), "DESC", "ASC")
            End If
            Me.Members_Bind
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Search_search_button.Click, New EventHandler(AddressOf Me.Search_search_Click)
            AddHandler Me.Members_insert.Click, New EventHandler(AddressOf Me.Members_insert_Click)
            AddHandler Me.Members_Pager.NavigateCompleted, New NavigateCompletedHandler(AddressOf Me.Members_pager_navigate_completed)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Search_Show
            Me.Members_Bind
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Search_search_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = ((Me.Search_FormAction & "name=" & Me.Search_name.Text & "&") & "")
            MyBase.Response.Redirect(url)
        End Sub

        Private Sub Search_Show()
            Dim param As String = Me.Utility.GetParam("name")
            Me.Search_name.Text = param
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
        Protected i_Members_curpage As Integer = 1
        Protected Members_Column_last_name As LinkButton
        Protected Members_Column_member_level As LinkButton
        Protected Members_Column_member_login As LinkButton
        Protected Members_Column_name As LinkButton
        Protected Members_CountPage As Integer
        Protected Members_FormAction As String = "MembersRecord.aspx?"
        Protected Members_holder As HtmlTable
        Protected Members_insert As LinkButton
        Protected Members_member_level_lov As String() = "1;Member;2;Administrator".Split(New Char() { ";"c })
        Protected Members_no_records As HtmlTableRow
        Private Const Members_PAGENUM As Integer = 20
        Protected Members_Pager As CCPager
        Protected Members_Repeater As Repeater
        Protected Members_sCountSQL As String
        Protected Members_sSQL As String
        Protected MembersForm_Title As Label
        Protected Search_FormAction As String = "MembersGrid.aspx?"
        Protected Search_holder As HtmlTable
        Protected Search_name As TextBox
        Protected Search_search_button As Button
        Protected Utility As CCUtility
    End Class
End Namespace

