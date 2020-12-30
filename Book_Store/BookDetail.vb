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
    Public Class BookDetail
        Inherits Page
        Implements IRequiresSessionState
        ' Methods
        Public Sub New()
            AddHandler MyBase.Init, New EventHandler(AddressOf Me.Page_Init)
        End Sub

        Private Sub Detail_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Sub Detail_Show()
            Dim flag As Boolean = True
            If (Me.p_Detail_itaem_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "item_id=" & CCUtility.ToSQL(Me.p_Detail_item_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from items where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Detail") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Detail_item_id.Value = CCUtility.GetValue(row, "item_id")
                    Me.Detail_name.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "name").ToString)
                    Me.Detail_author.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "author").ToString)
                    Me.Detail_category_id.Text = MyBase.Server.HtmlEncode(Me.Utility.Dlookup("categories", "name", ("category_id=" & CCUtility.ToSQL(CCUtility.GetValue(row, "category_id"), FieldTypes.Number))).ToString)
                    Me.Detail_price.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "price").ToString)
                    Me.Detail_image_url.Text = CCUtility.GetValue(row, "image_url")
                    Me.Detail_image_url.NavigateUrl = (CCUtility.GetValue(row, "product_url") & "")
                    Me.Detail_notes.Text = CCUtility.GetValue(row, "notes")
                    Me.Detail_product_url.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "product_url").ToString)
                    Me.Detail_product_url.NavigateUrl = (CCUtility.GetValue(row, "product_url") & "")
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("item_id")
                Me.Detail_item_id.Value = param
            End If
            Me.Detail_image_url.ImageUrl = Me.Detail_image_url.Text
            Me.Detail_product_url.Text = "Review this book on Amazon.com"
        End Sub

        Private Function Detail_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Detail_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Detail") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Detail_ValidationSummary.Text = (Me.Detail_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Detail_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Private Sub InitializeComponent()
        End Sub

        Private Sub Order_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("insert") > 0) Then
                flag = Me.Order_insert_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.Order_FormAction & ""))
            End If
        End Sub

        Private Sub Order_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Function Order_insert_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Order_Validate
            Dim str2 As String = CCUtility.ToSQL(Me.Session.Item("UserID").ToString, FieldTypes.Number)
            Dim str3 As String = CCUtility.ToSQL(Me.Utility.GetParam("Order_quantity"), FieldTypes.Number)
            Dim str4 As String = CCUtility.ToSQL(Me.Utility.GetParam("Order_item_id"), FieldTypes.Number)
            If flag Then
                If (sQL.Length = 0) Then
                    sQL = String.Concat(New String() { "insert into orders ([member_id],quantity,item_id) values (", str2, ",", str3, ",", str4, ")" })
                End If
                Me.Order_BeforeSQLExecute(sQL, "Insert")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Order_ValidationSummary.Text = (Me.Order_ValidationSummary.Text & exception.Message)
                    Me.Order_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Sub Order_Show()
            Dim flag As Boolean = True
            If (Me.p_Order_order_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "order_id=" & CCUtility.ToSQL(Me.p_Order_order_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from orders where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Order") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Order_order_id.Value = CCUtility.GetValue(row, "order_id")
                    Me.Order_quantity.Text = CCUtility.GetValue(row, "quantity")
                    Me.Order_item_id.Value = CCUtility.GetValue(row, "item_id")
                    Me.Order_insert.Visible = False
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("item_id")
                Me.Order_item_id.Value = param
            End If
        End Sub

        Private Function Order_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Order_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Order") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Order_ValidationSummary.Text = (Me.Order_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Order_ValidationSummary.Visible = Not flag
            Return flag
        End Function

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            Me.InitializeComponent
            AddHandler Me.Order_insert.ServerClick, New EventHandler(AddressOf Me.Order_Action)
            AddHandler Me.Rating_update.ServerClick, New EventHandler(AddressOf Me.Rating_Action)
        End Sub

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Me.Utility = New CCUtility(Me)
            Me.Utility.CheckSecurity(1)
            If Not MyBase.IsPostBack Then
                Me.p_Detail_item_id.Value = Me.Utility.GetParam("item_id")
                Me.p_Order_order_id.Value = Me.Utility.GetParam("order_id")
                Me.p_Rating_item_id.Value = Me.Utility.GetParam("item_id")
                Me.Page_Show(sender, e)
            End If
        End Sub

        Protected Sub Page_Show(ByVal sender As Object, ByVal e As EventArgs)
            Me.Detail_Show
            Me.Order_Show
            Me.Rating_Show
        End Sub

        Protected Sub Page_Unload(ByVal sender As Object, ByVal e As EventArgs)
            If (Not Me.Utility Is Nothing) Then
                Me.Utility.DBClose
            End If
        End Sub

        Private Sub Rating_Action(ByVal Src As Object, ByVal E As EventArgs)
            Dim flag As Boolean = True
            If (DirectCast(Src, HtmlInputButton).ID.IndexOf("update") > 0) Then
                flag = Me.Rating_update_Click(Src, E)
            End If
            If flag Then
                MyBase.Response.Redirect((Me.Rating_FormAction & "item_id=" & MyBase.Server.UrlEncode(Me.Utility.GetParam("item_id")) & "&"))
            End If
        End Sub

        Private Sub Rating_BeforeSQLExecute(ByVal SQL As String, ByVal Action As String)
        End Sub

        Private Sub Rating_Show()
            Me.Utility.buildListBox(Me.Rating_rating.Items, Me.Rating_rating_lov, Nothing, "")
            Dim flag As Boolean = True
            If (Me.p_Rating_item_id.Value.Length > 0) Then
                Dim str As String = ""
                str = (str & "item_id=" & CCUtility.ToSQL(Me.p_Rating_item_id.Value, FieldTypes.Number))
                Dim adapter As New OleDbDataAdapter(("select * from items where " & str), Me.Utility.Connection)
                Dim dataSet As New DataSet
                If (adapter.Fill(dataSet, 0, 1, "Rating") > 0) Then
                    Dim row As DataRow = dataSet.Tables.Item(0).Rows.Item(0)
                    Me.Rating_item_id.Value = CCUtility.GetValue(row, "item_id")
                    Me.Rating_rating_view.Text = CCUtility.GetValue(row, "rating")
                    Me.Rating_rating_count_view.Text = MyBase.Server.HtmlEncode(CCUtility.GetValue(row, "rating_count").ToString)
                    Dim str3 As String = CCUtility.GetValue(row, "rating")
                    Try 
                        Me.Rating_rating.SelectedIndex = Me.Rating_rating.Items.IndexOf(Me.Rating_rating.Items.FindByValue(str3))
                    Catch obj1 As Object
                    End Try
                    Me.Rating_rating_count.Value = CCUtility.GetValue(row, "rating_count")
                    flag = False
                End If
            End If
            If flag Then
                Dim param As String = Me.Utility.GetParam("item_id")
                Me.Rating_item_id.Value = param
                Me.Rating_update.Visible = False
            End If
            If (Short.Parse(Me.Rating_rating_view.Text) = 0) Then
                Me.Rating_rating_view.Text = "Not rated yet"
                Me.Rating_rating_count_view.Text = ""
            Else
                Me.Rating_rating_view.Text = ("<img src=images/" & Math.Round(CDbl((Double.Parse(Me.Rating_rating.SelectedItem.Value) / Double.Parse(Me.Rating_rating_count.Value)))) & "stars.gif>")
            End If
        End Sub

        Private Function Rating_update_Click(ByVal Src As Object, ByVal E As EventArgs) As Boolean
            Dim str As String = ""
            Dim sQL As String = ""
            Dim flag As Boolean = Me.Rating_Validate
            If flag Then
                If (Me.p_Rating_item_id.Value.Length > 0) Then
                    str = (str & "item_id=" & CCUtility.ToSQL(Me.p_Rating_item_id.Value, FieldTypes.Number))
                End If
                If Not flag Then
                    Return flag
                End If
                sQL = (("update items set [rating]=" & CCUtility.ToSQL(Me.Utility.GetParam("Rating_rating"), FieldTypes.Number) & ",[rating_count]=" & CCUtility.ToSQL(Me.Utility.GetParam("Rating_rating_count"), FieldTypes.Number)) & " where " & str)
                sQL = ("update items set rating=rating+" & Me.Rating_rating.SelectedItem.Value & ", rating_count=rating_count+1 where item_id=" & Me.Rating_item_id.Value)
                Me.Rating_BeforeSQLExecute(sQL, "Update")
                Dim command As New OleDbCommand(sQL, Me.Utility.Connection)
                Try 
                    command.ExecuteNonQuery
                Catch exception As Exception
                    Me.Rating_ValidationSummary.Text = (Me.Rating_ValidationSummary.Text & exception.Message)
                    Me.Rating_ValidationSummary.Visible = True
                    Return False
                End Try
            End If
            Return flag
        End Function

        Private Function Rating_Validate() As Boolean
            Dim flag As Boolean = True
            Me.Rating_ValidationSummary.Text = ""
            Dim i As Integer
            For i = 0 To Me.Page.Validators.Count - 1
                If (DirectCast(Me.Page.Validators.Item(i), BaseValidator).ID.ToString.StartsWith("Rating") AndAlso Not Me.Page.Validators.Item(i).IsValid) Then
                    Me.Rating_ValidationSummary.Text = (Me.Rating_ValidationSummary.Text & Me.Page.Validators.Item(i).ErrorMessage & "<br>")
                    flag = False
                End If
            Next i
            Me.Rating_ValidationSummary.Visible = Not flag
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
        Protected Detail_author As Label
        Protected Detail_category_id As Label
        Protected Detail_FormAction As String = "ShoppingCart.aspx?"
        Protected Detail_holder As HtmlTable
        Protected Detail_image_url As HyperLink
        Protected Detail_item_id As HtmlInputHidden
        Protected Detail_name As Label
        Protected Detail_notes As Label
        Protected Detail_price As Label
        Protected Detail_product_url As HyperLink
        Protected Detail_ValidationSummary As Label
        Protected DetailForm_Title As Label
        Protected Footer As Footer
        Protected Header As Header
        Protected Order_FormAction As String = "ShoppingCart.aspx?"
        Protected Order_holder As HtmlTable
        Protected Order_insert As HtmlInputButton
        Protected Order_item_id As HtmlInputHidden
        Protected Order_order_id As HtmlInputHidden
        Protected Order_quantity As TextBox
        Protected Order_quantity_Validator_Num As CustomValidator
        Protected Order_quantity_Validator_Req As RequiredFieldValidator
        Protected Order_ValidationSummary As Label
        Protected OrderForm_Title As Label
        Protected p_Detail_item_id As HtmlInputHidden
        Protected p_Order_order_id As HtmlInputHidden
        Protected p_Rating_item_id As HtmlInputHidden
        Protected Rating_FormAction As String = "BookDetail.aspx?"
        Protected Rating_holder As HtmlTable
        Protected Rating_item_id As HtmlInputHidden
        Protected Rating_rating As DropDownList
        Protected Rating_rating_count As HtmlInputHidden
        Protected Rating_rating_count_view As Label
        Protected Rating_rating_lov As String() = "1;Deficient;2;Regular;3;Good;4;Very Good;5;Excellent".Split(New Char() { ";"c })
        Protected Rating_rating_Validator_Num As CustomValidator
        Protected Rating_rating_Validator_Req As RequiredFieldValidator
        Protected Rating_rating_view As Label
        Protected Rating_update As HtmlInputButton
        Protected Rating_ValidationSummary As Label
        Protected RatingForm_Title As Label
        Protected Utility As CCUtility
    End Class
End Namespace