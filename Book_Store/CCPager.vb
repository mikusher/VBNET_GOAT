Imports ASP
Imports System
Imports System.Collections
Imports System.Runtime.CompilerServices
Imports System.Web.Profile
Imports System.Web.UI
Imports System.Web.UI.WebControls

Namespace Book_Store
    Public Class CCPager
        Inherits UserControl
        ' Events
        Public Event NavigateCompleted As NavigateCompletedHandler

        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub First_Create()
            If (Me.CurrentPage = 1) Then
                Me.lblFirst.Visible = True
                Me.lnkFirst.Visible = False
            Else
                Me.lblFirst.Visible = False
                Me.lnkFirst.Visible = True
            End If
        End Sub

        Private Sub InitializeComponent()
        End Sub

        Private Sub Last_Create()
            If (Me.CurrentPage = Me.MaxPage) Then
                Me.lblLast.Visible = True
                Me.lnkLast.Visible = False
            Else
                Me.lblLast.Visible = False
                Me.lnkLast.Visible = True
            End If
        End Sub

        Protected Sub lnkFirst_Click(ByVal sender As Object, ByVal e As EventArgs)
            Me.CurrentPage = 1
            Me.ViewState.Item("CurrPage") = Me.CurrentPage
            Me.Refresh_Components
        End Sub

        Protected Sub lnkLast_Click(ByVal sender As Object, ByVal e As EventArgs)
            Me.CurrentPage = Me.MaxPage
            Me.ViewState.Item("CurrPage") = Me.CurrentPage
            Me.Refresh_Components
        End Sub

        Protected Sub lnkNext_Click(ByVal sender As Object, ByVal e As EventArgs)
            Me.CurrentPage += 1
            Me.ViewState.Item("CurrPage") = Me.CurrentPage
            Me.Refresh_Components
        End Sub

        Protected Sub lnkPrev_Click(ByVal sender As Object, ByVal e As EventArgs)
            Me.CurrentPage -= 1
            Me.ViewState.Item("CurrPage") = Me.CurrentPage
            Me.Refresh_Components
        End Sub

        Private Sub Next_Create()
            If (Me.CurrentPage = Me.MaxPage) Then
                Me.lblNext.Visible = True
                Me.lnkNext.Visible = False
            Else
                Me.lblNext.Visible = False
                Me.lnkNext.Visible = True
            End If
        End Sub

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            If Not MyBase.IsPostBack Then
                If (Not Me.CssClass Is Nothing) Then
                    Try 
                        Me.lnkFirst.CssClass = Me.CssClass
                        Me.lnkLast.CssClass = Me.CssClass
                        Me.lnkPrev.CssClass = Me.CssClass
                        Me.lnkNext.CssClass = Me.CssClass
                        Me.lblPrev.CssClass = Me.CssClass
                        Me.lblNext.CssClass = Me.CssClass
                        Me.lblLast.CssClass = Me.CssClass
                        Me.lblFirst.CssClass = Me.CssClass
                        Me.lblShowTotal.CssClass = Me.CssClass
                    Catch obj1 As Object
                    End Try
                End If
                If (Not Me.style Is Nothing) Then
                    Dim strArray As String() = Me.style.Split(New Char() { ";"c })
                    Dim i As Integer
                    For i = 0 To strArray.Length - 1
                        Try 
                            Dim key As String = strArray(i).Split(New Char() { ":"c })(0).Trim
                            Dim str2 As String = strArray(i).Split(New Char() { ":"c })(1).Trim
                            Me.lnkFirst.Style.Add(key, str2)
                            Me.lnkLast.Style.Add(key, str2)
                            Me.lnkPrev.Style.Add(key, str2)
                            Me.lnkNext.Style.Add(key, str2)
                            Me.lblPrev.Style.Add(key, str2)
                            Me.lblNext.Style.Add(key, str2)
                            Me.lblLast.Style.Add(key, str2)
                            Me.lblFirst.Style.Add(key, str2)
                            Me.lblShowTotal.Style.Add(key, str2)
                        Catch obj2 As Object
                        End Try
                    Next i
                End If
                Me.ViewState.Item("MaxPage") = Me.MaxPage
                Me.lnkFirst.Text = Me.ShowFirstCaption
                Me.lnkPrev.Text = Me.ShowPrevCaption
                Me.lnkNext.Text = Me.ShowNextCaption
                Me.lnkLast.Text = Me.ShowLastCaption
                Me.lblPrev.Text = Me.ShowPrevCaption
                Me.lblNext.Text = Me.ShowNextCaption
                Me.lblLast.Text = Me.ShowLastCaption
                Me.lblFirst.Text = Me.ShowFirstCaption
                Me.lblShowTotal.Text = (Me.ShowTotalString & " "c & Me.MaxPage.ToString)
                Me.Refresh_Components
            Else
                Me.MaxPage = CInt(Me.ViewState.Item("MaxPage"))
                Try 
                    Me.CurrentPage = CInt(Me.ViewState.Item("CurrPage"))
                Catch obj3 As Object
                    Me.CurrentPage = 1
                End Try
            End If
        End Sub

        Private Sub Prev_Create()
            If (Me.CurrentPage = 1) Then
                Me.lblPrev.Visible = True
                Me.lnkPrev.Visible = False
            Else
                Me.lblPrev.Visible = False
                Me.lnkPrev.Visible = True
            End If
        End Sub

        Private Sub Refresh_Components()
            If Me.ShowFirst Then
                Me.First_Create
            End If
            If Me.ShowLast Then
                Me.Last_Create
            End If
            If Me.ShowPrev Then
                Me.Prev_Create
            End If
            If Me.ShowNext Then
                Me.Next_Create
            End If
            Me.Show_Pager
            Me.NavigateCompleted.Invoke(Me, Me.CurrentPage)
        End Sub

        Protected Sub Repeater1_ItemCommand(ByVal Sender As Object, ByVal e As RepeaterCommandEventArgs)
            If (Me.PagerStyle <> PagerStyleEnum.Moved) Then
                Me.CurrentPage = Short.Parse(e.CommandArgument.ToString)
                Me.ViewState.Item("CurrPage") = Me.CurrentPage
                Me.Refresh_Components
            Else
                Dim s As String = e.CommandArgument.ToString
                Me.PagerStartPressed = s.StartsWith("<")
                Me.PagerEndPressed = s.EndsWith(">")
                Me.CurrentPage = IIf(Me.PagerStartPressed, Short.Parse(s.TrimStart(New Char() { "<"c })), IIf(Me.PagerEndPressed, Short.Parse(s.ToString.TrimEnd(New Char() { ">"c })), Short.Parse(s)))
                Me.ViewState.Item("CurrPage") = Me.CurrentPage
                Me.Refresh_Components
            End If
        End Sub

        Protected Sub Repeater1_ItemDataBound(ByVal source As Object, ByVal e As RepeaterItemEventArgs)
            If ((e.Item.ItemType = ListItemType.Item) OrElse (e.Item.ItemType = ListItemType.AlternatingItem)) Then
                If (Not Me.CssClass Is Nothing) Then
                    Try 
                        DirectCast(e.Item.FindControl("LinkButton1"), LinkButton).CssClass = Me.CssClass
                    Catch obj1 As Object
                    End Try
                End If
                If (Not Me.style Is Nothing) Then
                    Dim strArray As String() = Me.style.Split(New Char() { ";"c })
                    Dim i As Integer
                    For i = 0 To strArray.Length - 1
                        Try 
                            Dim key As String = strArray(i).Split(New Char() { ":"c })(0).Trim
                            Dim str2 As String = strArray(i).Split(New Char() { ":"c })(1).Trim
                            DirectCast(e.Item.FindControl("LinkButton1"), LinkButton).Style.Add(key, str2)
                        Catch obj2 As Object
                        End Try
                    Next i
                End If
            End If
        End Sub

        Private Sub Show_Pager()
            Dim num As Integer
            Dim list As New ArrayList
            Select Case Me.PagerStyle
                Case PagerStyleEnum.NoPager
                    Me.Repeater1.Visible = False
                    Me.Current.Visible = False
                    Me.OpenPar.Visible = False
                    Me.ClosePar.Visible = False
                    Me.lblShowTotal.Visible = False
                    Return
                Case PagerStyleEnum.PageNumberOnly
                    Me.Repeater1.Visible = False
                    Me.Current.Visible = True
                    Me.Current.Text = Me.CurrentPage.ToString
                    Me.lblShowTotal.Visible = Me.ShowTotal
                    Return
                Case PagerStyleEnum.Centered
                    Me.Repeater1.Visible = True
                    Me.Current.Visible = False
                    num = IIf((Me.CurrentPage <= (Me.NumberOfPages / 2)), 1, (Me.CurrentPage - (Me.NumberOfPages / 2)))
                    num = IIf((num >= ((Me.MaxPage - Me.NumberOfPages) + 1)), ((Me.MaxPage - Me.NumberOfPages) + 1), num)
                    num = IIf((num <= 0), 1, num)
                    Dim i As Integer = num
                    Do While ((i < (num + Me.NumberOfPages)) AndAlso (i <= Me.MaxPage))
                        list.Add(New PagingData(i.ToString, (Me.CurrentPage <> i)))
                        i += 1
                    Loop
                    Me.Repeater1.DataSource = list
                    Me.Repeater1.DataBind
                    Me.lblShowTotal.Visible = Me.ShowTotal
                    Return
                Case PagerStyleEnum.Moved
                    num = ((((Me.CurrentPage - 1) / Me.NumberOfPages) * Me.NumberOfPages) + 1)
                    If (num <> 1) Then
                        Dim num4 As Integer = (num - 1)
                        list.Add(New PagingData(("<"c & num4.ToString), True))
                    End If
                    Dim j As Integer = num
                    Do While ((j < (num + Me.NumberOfPages)) AndAlso (j <= Me.MaxPage))
                        list.Add(New PagingData(j.ToString, (Me.CurrentPage <> j)))
                        j += 1
                    Loop
                    If ((num + Me.NumberOfPages) <= Me.MaxPage) Then
                        Dim num5 As Integer = (num + Me.NumberOfPages)
                        list.Add(New PagingData((num5.ToString & ">"c), True))
                    End If
                    Me.Repeater1.DataSource = list
                    Me.Repeater1.DataBind
                    Me.Repeater1.Visible = True
                    Me.Current.Visible = False
                    Me.lblShowTotal.Visible = Me.ShowTotal
                    Return
            End Select
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
        Protected ClosePar As Label
        Public CssClass As String
        Protected Current As Label
        Public CurrentPage As Integer = 1
        Protected lblFirst As Label
        Protected lblLast As Label
        Protected lblNext As Label
        Protected lblPrev As Label
        Protected lblShowTotal As Label
        Protected lnkFirst As LinkButton
        Protected lnkLast As LinkButton
        Protected lnkNext As LinkButton
        Protected lnkPrev As LinkButton
        Public MaxPage As Integer
        Public NumberOfPages As Integer
        Protected OpenPar As Label
        Protected PagerEndPressed As Boolean
        Protected PagerStartPressed As Boolean
        Public PagerStyle As PagerStyleEnum
        Protected Repeater1 As Repeater
        Public ShowFirst As Boolean
        Public ShowFirstCaption As String
        Public ShowLast As Boolean
        Public ShowLastCaption As String
        Public ShowNext As Boolean
        Public ShowNextCaption As String
        Public ShowPrev As Boolean
        Public ShowPrevCaption As String
        Public ShowTotal As Boolean
        Public ShowTotalString As String
        Public style As String

        ' Nested Types
        Public Enum PagerStyleEnum
            ' Fields
            Centered = 2
            Moved = 3
            NoPager = 0
            PageNumberOnly = 1
        End Enum
    End Class
End Namespace

