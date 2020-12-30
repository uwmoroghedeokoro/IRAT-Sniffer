Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.VisualBasic

Public Class agent
    Public id As Integer
    Public firstn As String
    Public lastn As String
    Public email As String
    Public pwd As String
    Public role As String
    Public scope As String
    Public associate_group As Integer
    Public groupsByID As New List(Of Integer)
    Public myStats As New reportStats
    Public loggedIn As Boolean
    Public fullname As String
    Public firstLogin As Boolean
    Private connStr As String = "Data Source=irat-srv-acc01;Initial Catalog=irl_sd;Integrated Security=false;user id=sa;password=7mmT@XAy"


    Sub New()

    End Sub

    Sub New(emaild As String, fname As String, lname As String)
        email = emaild
        firstn = fname
        lastn = lname
        fullname = fname & " " & lname
        Dim sqlcon As New SqlConnection(connStr)


        Try
            sqlcon.Open()

            Dim sql As String
            sql = "select * From tbl_Agents where email ='" & emaild & "'"

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader()

            Dim found As Boolean = False

            While dbRead.Read
                id = dbRead("id")
                firstn = dbRead("firstn")
                pwd = dbRead("pwd")
                lastn = dbRead("lastn")
                email = dbRead("email")
                role = dbRead("role")
                scope = dbRead("scope")
                fullname = dbRead("firstn") & " " & dbRead("lastn")
                firstLogin = dbRead("firstLogin")
                found = True
            End While
            dbRead.Close()

            If found = False Then
                addNew()
            End If

        Finally
            sqlcon.Dispose()

        End Try

    End Sub
    Sub New(aid As Integer)
        id = aid

        Dim sqlcon As New SqlConnection(connStr)


        Try
            sqlcon.Open()

            Dim sql As String
            sql = "select * From tbl_Agents where id=" & id

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader()

            While dbRead.Read
                id = dbRead("id")
                firstn = dbRead("firstn")
                lastn = dbRead("lastn")
                email = dbRead("email")
                role = dbRead("role")
                scope = dbRead("scope")
                fullname = dbRead("firstn") & " " & dbRead("lastn")
                firstLogin = dbRead("firstLogin")
            End While
            dbRead.Close()

            'load my groups
            sqlcommand = New SqlCommand("select groupid from tblGroupMapping where agentid=" & id, sqlcon)
            dbRead = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While dbRead.Read
                groupsByID.Add(dbRead("groupid"))
            End While
            dbRead.Close()
        Finally
            sqlcon.Dispose()

        End Try
    End Sub
    Public Sub addNew()
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim sendI As Boolean = False
            Dim sendW As Boolean = False
            Dim sqlcommand As New SqlCommand("addAgent", sqlcon)
            sqlcommand.CommandType = Data.CommandType.StoredProcedure
            If (id < 1 And pwd = "") Then
                pwd = generatePWD()
                sendI = True
            ElseIf (id < 1 And pwd <> "") Then
                sendW = True
            End If
            sqlcommand.Parameters.AddWithValue("@firstn", firstn)
            sqlcommand.Parameters.AddWithValue("@lastn", lastn)
            sqlcommand.Parameters.AddWithValue("@email", email)
            sqlcommand.Parameters.AddWithValue("@pwd", pwd)
            sqlcommand.Parameters.AddWithValue("@role", role)
            sqlcommand.Parameters.AddWithValue("@id", id)
            sqlcommand.Parameters.AddWithValue("@scope", scope)
            sqlcommand.Parameters.AddWithValue("@groupid", associate_group)

            Dim param As New SqlParameter
            param.ParameterName = "returnV"
            param.Direction = Data.ParameterDirection.ReturnValue

            sqlcommand.Parameters.Add(param)

            sqlcommand.ExecuteNonQuery()

            Dim returnV As Integer
            id = sqlcommand.Parameters("returnV").Value

            For Each aID In groupsByID
                addToGroup(aID)
            Next
            If sendI Then
                sendInvite()
            End If
            If sendW Then
                ' sendWelcome()
            End If

        Finally
            sqlcon.Dispose()

        End Try
    End Sub
    Public Sub generateStats(frDate As DateTime, toDate As DateTime)
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim sql As String
            ''get total assigned
            sql = "select count(*) as count From tbl_tickets where linkid=-1 and agentid=" & id & " and lastUpdated >='" & frDate & "' and lastUpdated <='" & toDate & "'"

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader()
            While dbRead.Read
                myStats.assinged = IIf(IsDBNull(dbRead("count")), 0, CType(dbRead("count"), Integer))
            End While
            dbRead.Close()

            ''get resolved
            sql = "select count(*) as count From tbl_tickets where linkid=-1 and resolvedby=" & id & " and lastUpdated >='" & frDate & "' and lastUpdated <='" & toDate & "'"
            sqlcommand = New SqlCommand(sql, sqlcon)
            dbRead = sqlcommand.ExecuteReader()
            While dbRead.Read
                myStats.resolved = IIf(IsDBNull(dbRead("count")), 0, CType(dbRead("count"), Integer))
            End While
            dbRead.Close()

            ''get responses
            sql = "select count(*) as count From tbl_tickets where updatedby=" & id & " and linkid <> -1 and lastUpdated >='" & frDate & "' and lastUpdated <='" & toDate & "'"
            sqlcommand = New SqlCommand(sql, sqlcon)
            dbRead = sqlcommand.ExecuteReader()
            While dbRead.Read
                myStats.responses = IIf(IsDBNull(dbRead("count")), 0, CType(dbRead("count"), Integer))
            End While
            dbRead.Close()

            ''calc avg resolution time
            sql = "SELECT AVG(CAST(CAST(CAST(resolveDate AS float) - CAST(created AS float) AS int) * 24 + DATEPART(hh, resolveDate - created) AS float)) AS avgres FROM tbl_tickets where updatedby=" & id & " and linkid = -1 and (CONVERT(varchar(10), created, 120) >= cast('" & frDate & "' as datetime) and CONVERT(varchar(10), created, 120) <= cast('" & toDate & "' as datetime))"
            sqlcommand = New SqlCommand(sql, sqlcon)
            dbRead = sqlcommand.ExecuteReader()
            While dbRead.Read
                myStats.avgFirstRespond = IIf(IsDBNull(dbRead("avgres")), "-", dbRead("avgres") & " hours")
            End While
            dbRead.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub updatePWD(pwd As String)

        Dim sqlcon As New SqlConnection(connStr)


        Try
            sqlcon.Open()

            Dim sql As String

            sql = "update tbl_agents set firstlogin='false',pwd='" & pwd & "' where id=" & id

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            sqlcommand.ExecuteNonQuery()

        Finally
            sqlcon.Dispose()

        End Try

    End Sub
    Sub New(emaild As String, pwd As String)

        Dim sqlcon As New SqlConnection(connStr)


        Try
            sqlcon.Open()

            Dim sql As String
            sql = "select * from tbl_agents where email='" & emaild & "' and pwd='" & pwd & "' COLLATE SQL_Latin1_General_CP1_CS_AS"

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader()

            While dbRead.Read
                loggedIn = True
                id = dbRead("id")
                firstn = dbRead("firstn")
                lastn = dbRead("lastn")
                email = dbRead("email")
                role = dbRead("role")
                scope = dbRead("scope")
                fullname = dbRead("firstn") & " " & dbRead("lastn")
                firstLogin = dbRead("firstLogin")
            End While
            dbRead.Close()


        Finally
            sqlcon.Dispose()

        End Try
    End Sub
    Function generatePWD() As String
        Dim s As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Dim r As New Random
        Dim sb As New StringBuilder
        For i As Integer = 1 To 8
            Dim idx As Integer = r.Next(0, 35)
            sb.Append(s.Substring(idx, 1))
        Next
        generatePWD = (sb.ToString())
        ' Console.ReadKey()
    End Function
    Public Function getAgents(Optional aType As String = "all") As List(Of agent)
        Dim sqlcon As New SqlConnection(connStr)

        getAgents = New List(Of agent)

        Try
            sqlcon.Open()

            Dim sql As String
            sql = "select * From tbl_Agents order by firstn ASC"

            If aType = "agent" Then
                sql = "select * From tbl_Agents where role='agent' order by firstn ASC"
            ElseIf aType = "administrator" Then
                sql = "select * From tbl_Agents where role='administrator' order by firstn ASC"
            End If

            If scope = "restricted" Then
                sql = "select * from tbl_agents where id=" & id
            End If

            Dim sqlcommand As New SqlCommand(sql, sqlcon)
            Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While dbRead.Read
                Dim tmp As New agent
                With tmp
                    .id = dbRead("id")
                    .firstn = dbRead("firstn")
                    .lastn = dbRead("lastn")
                    .email = dbRead("email")
                    .role = dbRead("role")
                    .scope = dbRead("scope")
                    .fullname = dbRead("firstn") & " " & dbRead("lastn")
                End With
                getAgents.Add(tmp)
            End While
            dbRead.Close()

        Finally
            sqlcon.Dispose()

        End Try
    End Function
    Structure reportStats
        Dim assinged As Integer
        Dim resolved As Integer
        Dim responses As Integer
        Dim avgFirstRespond As String
    End Structure

    Private Sub sendInvite()
        Try
            Dim Message As New Net.Mail.MailMessage
            ' Dim agt = New agent(assigner)
            'Dim smtp1 As New Net.Mail.SmtpClient("relay.jangosmtp.net", 587)
            'Dim myCredential As New Net.NetworkCredential("uwmoroghe", "Godz4ever")
            Dim smtp1 As New Net.Mail.SmtpClient("10.206.100.111", 25)
            Dim myCredential As New Net.NetworkCredential("iratdesk", "k33p1ts1mpl3@")

            Message.From = New Net.Mail.MailAddress("iratdesk@islandroutes.com", "Island Routes Helpdesk")
            Message.To.Add(New Net.Mail.MailAddress((email)))



            Message.Subject = "Welcome to the Island Routes Service Desk"
            Message.IsBodyHtml = True
            smtp1.EnableSsl = False
            smtp1.Port = 25
            Message.Body = "<font style='font-family:arial;font-size:10pt'>"
            Message.Body &= "Hi " & firstn & ",<br><br>You have been invited to join the Island Routes Service Desk. The Service Desk is your workspace for creating tickets for specified requests.<br><br>Use the below temporary password to access your account. <br><br>Temporary password: " & pwd & "<br><br><a href='http://irat-srv-acc01/'>Access IRAT Service Desk</a><br><br>Regards,<br>Island Routes Service Desk"
            Message.Body &= "</font>"
            smtp1.Credentials = myCredential

            smtp1.Send(Message)
        Catch ex As Exception

        End Try
    End Sub
    Sub addToGroup(groupid As Integer)
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim sqlcommand As New SqlCommand("addToGroup", sqlcon)
            sqlcommand.CommandType = Data.CommandType.StoredProcedure

            sqlcommand.Parameters.AddWithValue("@groupid", groupid)
            sqlcommand.Parameters.AddWithValue("@agentid", id)

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
    Structure agentRole
        Dim roleid As Integer
        Dim role As String
    End Structure


    Structure loginResult
        Dim agentid As Integer
        Dim loggedin As Boolean
        Dim role As String
        Dim scope As String
    End Structure
End Class
