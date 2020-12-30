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
    Public Class CardTypesGrid
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub CardTypes_Bind()
            Me.CardTypes_Repeater.DataSource = Me.CardTypes_CreateDataSource
            Me.CardTypes_Repeater.DataBind
        End Sub

        Private Function CardTypes_CreateDataSource() As ICollection
            Me.CardTypes_sSQL = ""
            Me.CardTypes_sCountSQL = ""
            Dim str As String = ""
            Dim str2 As String = ""
            str2 = " order by c.name Asc"
            If ((Me.Utility.GetParam("FormCardTypes_Sorting").Length > 0) AndAlso Not MyBase.IsPostBack) Then
                Me.ViewState.Item("SortColumn") = Me.Utility.GetParam("FormCardTypes_Sorting")
                Me.ViewState.Item("SortDir") = "ASC"
            End If
            If (Not Me.ViewState.Item("SortColumn") Is Nothing) Then
                str2 = (" ORDER BY " & Me.ViewState.Item("SortColumn").ToString & " " & Me.ViewState.Item("SortDir").ToString)
            End If
            New StringDictionary
            Me.CardTypes_sSQL = "select [c].[card_type_id] as c_card_type_id, [c].[name] as c_name  from [card_types] c "
            Me.CardTypes_sSQL = (Me.CardTypes_sSQL & str & str2)
            Dim adapter As New OleDbDataAdapter(Me.CardTypes_sSQL, Me.Utility.Connection)
            Dim dataSet As New DataSet
            adapter.Fill(dataSet, 0, 20, "CardTypes")
            Dim view As New DataView(dataSet.Tables.Item(0))
            If (dataSet.Tables.Item(0).Rows.Count = 0) Then
                Me.CardTypes_no_records.Visible = True
                Return view
            End If
            Me.CardTypes_no_records.Visible = False
            Return view
        End Function

        Private Sub CardTypes_insert_Click(ByVal Src As Object, ByVal E As EventArgs)
            Dim url As String = (Me.CardTypes_FormAction & "")
            MyBase.Response.Redirect(url)
        End Sub

        Public Sub CardTypes_Repeater_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        End Sub

        Protected Sub CardTypes_SortChange(ByVal Src As Object, ByVal E As EventArgs)
            If ((Me.ViewState.Item("SortColumn") Is Nothing) OrElse (Me.ViewState.Item("SortColumn").ToString <> DirectCast(Src, LinkButton).CommandArgument)) Then
                Me.ViewState.Item("SortColumn") = DirectCast(Src, LinkButton).CommandArgument
                Me.ViewState.Item("SortDir") = "ASC"
            Else
                Me.ViewState.Item("SortDir") = IIf((Me.ViewState.Item("SortDir").ToString = "ASC"), "DESC", "ASC")
            End If
            Me.CardTypes_Bind
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.CardTypes_insert.Click, New EventHandler(AddressOf Me.CardTypes_insert_Click)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(2)
            If Not MyBase.IsPostBack Then
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.CardTypes_Bind
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
        Protected CardTypes_Column_name As LinkButton
        Protected CardTypes_CountPage As Integer
        Protected CardTypes_FormAction As String = "CardTypesRecord.aspx?"
        Protected CardTypes_holder As HtmlTable
        Protected CardTypes_insert As LinkButton
        Protected CardTypes_no_records As HtmlTableRow
        Private Const CardTypes_PAGENUM As Integer = 20
        Protected CardTypes_Repeater As Repeater
        Protected CardTypes_sCountSQL As String
        Protected CardTypes_sSQL As String
        Protected CardTypesForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected i_CardTypes_curpage As Integer = 1
        Protected Utility As CCUtility
    End Class
End Namespace

