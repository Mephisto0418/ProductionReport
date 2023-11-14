Imports System.Data.SqlClient

Module mdlVar
    Public sqlconnMQL03 As New SqlConnection("Data Source=10.44.66.105;Initial Catalog=H3_Systematic;User ID=pptcim;Password=cim945")
    Public Function SQL_Select(ByVal QueryString As String, ByRef SQLConn As SqlConnection)
        Dim sqldaQuery As New SqlDataAdapter
        sqldaQuery.SelectCommand = SQLConn.CreateCommand
        sqldaQuery.SelectCommand.CommandText = QueryString
        If QueryString.Contains(";") = False Then
            Dim dtQuery As New DataTable
            If SQLConn.State = ConnectionState.Closed Then SQLConn.Open()
            sqldaQuery.Fill(dtQuery)
            If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
            Return dtQuery
        Else
            Dim dsQuery As New DataSet
            If SQLConn.State = ConnectionState.Closed Then SQLConn.Open()
            sqldaQuery.Fill(dsQuery)
            If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
            Return dsQuery
        End If
    End Function

End Module
