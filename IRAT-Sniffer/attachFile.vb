Imports System.Data.SqlClient
Public Class attachFile
    Public attID As Integer
    Public attName As String
    Public attPath As String
    Private connStr As String = "Data Source=irat-srv-acc01;Initial Catalog=irl_sd;Integrated Security=false;user id=sa;password=7mmT@XAy"

    Sub New()

    End Sub
    Public Sub addew(ticketID As Integer)
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            ' Dim lastAgentID As Integer = 0
            Dim sqlcommand As New SqlCommand("insert into tbl_files (linkid,filename,filepath) values (" & ticketID & ",'" & attName & "','" & attPath & "')", sqlcon)

            sqlcommand.ExecuteNonQuery()

        Finally
            sqlcon.Dispose()

        End Try
    End Sub
End Class
