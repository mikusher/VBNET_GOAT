Imports System
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI.WebControls

Namespace Book_Store
    Public Class CCUtility
        ' Methods
        Public Sub New(ByVal parent As Object)
            Me.DBOpen
        End Sub

        Public Sub buildListBox(ByVal Items As ListItemCollection, ByVal values As String(), ByVal CustomInitialDisplayValue As String, ByVal CustomInitialSubmitValue As String)
            Items.Clear
            If (Not CustomInitialDisplayValue Is Nothing) Then
                Items.Add(New ListItem(CustomInitialDisplayValue, CustomInitialSubmitValue))
            End If
            Dim i As Integer = 0
            Do While (i < values.Length)
                Items.Add(New ListItem(values((i + 1)), values(i)))
                i = (i + 2)
            Loop
        End Sub

        Public Function buildListBox(ByVal sSQL As String, ByVal sId As String, ByVal sTitle As String, ByVal CustomInitialDisplayValue As String, ByVal CustomInitialSubmitValue As String) As ICollection
            Dim adapter As New OleDbDataAdapter(sSQL, Me.Connection)
            Dim dataSet As New DataSet
            dataSet.Tables.Add("lookup")
            Dim column As New DataColumn
            column.DataType = Type.GetType("System.String")
            column.ColumnName = sId
            dataSet.Tables.Item(0).Columns.Add(column)
            column = New DataColumn
            column.DataType = Type.GetType("System.String")
            column.ColumnName = sTitle
            dataSet.Tables.Item(0).Columns.Add(column)
            If (Not CustomInitialDisplayValue Is Nothing) Then
                Dim row As DataRow = dataSet.Tables.Item(0).NewRow
                row.Item(0) = CustomInitialSubmitValue
                row.Item(1) = CustomInitialDisplayValue
                dataSet.Tables.Item(0).Rows.Add(row)
            End If
            adapter.Fill(dataSet, "lookup")
            Return New DataView(dataSet.Tables.Item(0))
        End Function

        Public Sub buildListBox(ByVal Items As ListItemCollection, ByVal sSQL As String, ByVal sId As String, ByVal sTitle As String, ByVal CustomInitialDisplayValue As String, ByVal CustomInitialSubmitValue As String)
            Items.Clear
            Dim reader As OleDbDataReader = New OleDbCommand(sSQL, Me.Connection).ExecuteReader
            If (Not CustomInitialDisplayValue Is Nothing) Then
                Items.Add(New ListItem(CustomInitialDisplayValue, CustomInitialSubmitValue))
            End If
            Do While reader.Read
                If ((sId = "") AndAlso (sTitle = "")) Then
                    Items.Add(New ListItem(reader.Item(1).ToString, reader.Item(0).ToString))
                Else
                    Items.Add(New ListItem(reader.Item(sTitle).ToString, reader.Item(sId).ToString))
                End If
            Loop
            reader.Close
        End Sub

        Public Sub CheckSecurity(ByVal iLevel As Integer)
            If ((Me.Session.Item("UserID") Is Nothing) OrElse (Me.Session.Item("UserID").ToString.Length = 0)) Then
                Me.Response.Redirect(("Login.aspx?QueryString=" & Me.Server.UrlEncode(Me.Request.ServerVariables.Item("QUERY_STRING")) & "&ret_page=" & Me.Server.UrlEncode(Me.Request.ServerVariables.Item("SCRIPT_NAME"))))
            ElseIf (Short.Parse(Me.Session.Item("UserRights").ToString) < iLevel) Then
                Me.Response.Redirect(("Login.aspx?QueryString=" & Me.Server.UrlEncode(Me.Request.ServerVariables.Item("QUERY_STRING")) & "&ret_page=" & Me.Server.UrlEncode(Me.Request.ServerVariables.Item("SCRIPT_NAME"))))
            End If
        End Sub

        Public Sub DBClose()
            Me.Connection.Close
        End Sub

        Public Sub DBOpen()
            Dim connectionString As String = ConfigurationSettings.AppSettings.Item("sBook_StoreDBConnectionString")
            Me.Connection = New OleDbConnection(connectionString)
            Me.Connection.Open
        End Sub

        Public Function Dlookup(ByVal table As String, ByVal field As String, ByVal sWhere As String) As String
            Dim str2 As String
            Dim reader As OleDbDataReader = New OleDbCommand(String.Concat(New String() { "SELECT ", field, " FROM ", table, " WHERE ", sWhere }), Me.Connection).ExecuteReader(CommandBehavior.SingleRow)
            If reader.Read Then
                str2 = reader.Item(0).ToString
                If (str2 Is Nothing) Then
                    str2 = ""
                End If
            Else
                str2 = ""
            End If
            reader.Close
            Return str2
        End Function

        Public Function DlookupInt(ByVal table As String, ByVal field As String, ByVal sWhere As String) As Integer
            Dim reader As OleDbDataReader = New OleDbCommand(String.Concat(New String() { "SELECT ", field, " FROM ", table, " WHERE ", sWhere }), Me.Connection).ExecuteReader(CommandBehavior.SingleRow)
            Dim num As Integer = -1
            If reader.Read Then
                num = reader.GetInt32(0)
            End If
            reader.Close
            Return num
        End Function

        Public Sub Execute(ByVal sSQL As String)
            New OleDbCommand(sSQL, Me.Connection).ExecuteNonQuery
        End Sub

        Public Function FillDataSet(ByVal sSQL As String) As DataSet
            Dim sset As New DataSet
            New OleDbDataAdapter(sSQL, Me.Connection)
            Return [sset]
        End Function

        Public Function FillDataSet(ByVal sSQL As String, ByRef ds As DataSet) As Integer
            Dim adapter As New OleDbDataAdapter(sSQL, Me.Connection)
            Return adapter.Fill(ds, "Table")
        End Function

        Public Function FillDataSet(ByVal sSQL As String, ByRef ds As DataSet, ByVal start As Integer, ByVal count As Integer) As Integer
            Dim adapter As New OleDbDataAdapter(sSQL, Me.Connection)
            Return adapter.Fill(ds, start, count, "Table")
        End Function

        Public Shared Function getCheckBoxValue(ByVal sVal As String, ByVal CheckedValue As String, ByVal UnCheckedValue As String, ByVal Type As FieldTypes) As String
            If (sVal.Length = 0) Then
                Return CCUtility.ToSQL(UnCheckedValue, Type)
            End If
            Return CCUtility.ToSQL(CheckedValue, Type)
        End Function

        Public Function GetParam(ByVal ParamName As String) As String
            Dim str As String = Me.Request.QueryString.Item(ParamName)
            If (str Is Nothing) Then
                str = Me.Request.Form.Item(ParamName)
            End If
            If (str Is Nothing) Then
                Return ""
            End If
            Return str
        End Function

        Public Shared Function GetValFromLOV(ByVal val As String, ByVal arr As String()) As String
            Dim str As String = ""
            If ((arr.Length Mod 2) = 0) Then
                Dim index As Integer = Array.IndexOf(Of String)(arr, val)
                str = IIf((index = -1), "", arr((index + 1)))
            End If
            Return str
        End Function

        Public Shared Function GetValue(ByVal row As DataRow, ByVal field As String) As String
            If (row.Item(field).ToString Is Nothing) Then
                Return ""
            End If
            Return row.Item(field).ToString
        End Function

        Public Function IsNumeric(ByVal source As Object, ByVal value As String) As Boolean
            Try 
                Convert.ToDecimal(value)
                Return True
            Catch obj1 As Object
                Return False
            End Try
        End Function

        Public Shared Function Quote(ByVal Param As String) As String
            If ((Not Param Is Nothing) AndAlso (Param.Length <> 0)) Then
                Return Param.Replace("'", "''")
            End If
            Return ""
        End Function

        Public Shared Function ToSQL(ByVal Param As String, ByVal Type As FieldTypes) As String
            If ((Param Is Nothing) OrElse (Param.Length = 0)) Then
                Return "Null"
            End If
            Dim str As String = CCUtility.Quote(Param)
            If (Type = FieldTypes.Number) Then
                Return str.Replace(","c, "."c)
            End If
            Return ("'" & str & "'")
        End Function


        ' Fields
        Public Connection As OleDbConnection
        Protected Request As HttpRequest = HttpContext.Current.Request
        Protected Response As HttpResponse = HttpContext.Current.Response
        Protected Server As HttpServerUtility = HttpContext.Current.Server
        Protected Session As HttpSessionState = HttpContext.Current.Session
    End Class
End Namespace
