Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class groups
    Public groupid As Integer
    Public groupName As String
    Public membersByID As New List(Of Integer)
    Public membersTotal As Integer
    Private connStr As String = "Data Source=irat-srv-acc01;Initial Catalog=irl_sd;Integrated Security=false;user id=sa;password=7mmT@XAy"

    Sub New()

    End Sub
    Sub New(grID As Integer)
        groupid = grID

        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()

            Dim sql As String
            sql = "select t.incidentgroup,t.id,count(tm.agentid) as members from tbl_incidentgroups t left outer join tblGroupMapping tm on t.id=tm.groupid where t.id=" & grID & " group by t.id,t.incidentgroup order by t.incidentgroup ASC"

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While dbRead.Read
                groupid = dbRead("id")
                groupName = dbRead("incidentgroup")
                '  getMembers()
            End While
            dbRead.Close()

        Finally
            sqlcon.Dispose()

        End Try
    End Sub
    Public Function getMembers() As List(Of agent)
        Dim sqlcon As New SqlConnection(connStr)

        getMembers = New List(Of agent)

        Try
            sqlcon.Open()

            Dim sql As String
            sql = "select * from tblGroupMapping where groupid=" & groupid


            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While dbRead.Read
                Dim tmp As New agent(dbRead("agentid"))
                getMembers.Add(tmp)
            End While

        Finally
            sqlcon.Dispose()
        End Try
    End Function
    Public Function getGroups() As List(Of groups)
        Dim sqlcon As New SqlConnection(connStr)

        getGroups = New List(Of groups)

        Try
            sqlcon.Open()

            Dim sql As String
            sql = "select t.incidentgroup,t.id,count(tm.agentid) as members from tbl_incidentgroups t left outer join tblGroupMapping tm on t.id=tm.groupid group by t.id,t.incidentgroup  order by t.incidentgroup ASC"

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While dbRead.Read
                Dim tmp As New groups
                With tmp
                    .groupid = dbRead("id")
                    .groupName = dbRead("incidentgroup")
                    .membersTotal = dbRead("members")
                End With
                getGroups.Add(tmp)
            End While
            dbRead.Close()

        Finally
            sqlcon.Dispose()

        End Try
    End Function
    Public Sub addNew()
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim sqlcommand As New SqlCommand("addGroup", sqlcon)
            sqlcommand.CommandType = Data.CommandType.StoredProcedure

            sqlcommand.Parameters.AddWithValue("@id", groupid)
            sqlcommand.Parameters.AddWithValue("@name", groupName)

            Dim param As New SqlParameter
            param.ParameterName = "returnV"
            param.Direction = Data.ParameterDirection.ReturnValue

            sqlcommand.Parameters.Add(param)

            sqlcommand.ExecuteNonQuery()

            ' Dim returnV As Integer
            groupid = sqlcommand.Parameters("returnV").Value

        Finally
            sqlcon.Dispose()

            For Each aID In membersByID
                addToGroup(aID)
            Next
        End Try

    End Sub

    Sub addToGroup(agentID As Integer)
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim sqlcommand As New SqlCommand("addToGroup", sqlcon)
            sqlcommand.CommandType = Data.CommandType.StoredProcedure

            sqlcommand.Parameters.AddWithValue("@groupid", groupid)
            sqlcommand.Parameters.AddWithValue("@agentid", agentID)

            Dim param As New SqlParameter
            param.ParameterName = "returnV"
            param.Direction = Data.ParameterDirection.ReturnValue

            sqlcommand.Parameters.Add(param)

            sqlcommand.ExecuteNonQuery()

            ' Dim returnV As Integer
            ' groupid = sqlcommand.Parameters("returnV").Value

        Finally
            sqlcon.Dispose()

        End Try
    End Sub
End Class
