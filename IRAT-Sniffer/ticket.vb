Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Microsoft.VisualBasic

Public Class ticket
    Public ticketid As Integer
    Public subject As String
    Public description As String
    Public filepath As String
    Public fileName As String
    Public linkid As Integer
    Public requester As New agent
    Public assignedAgent As New agent
    Public teamID As Integer
    Public ticketNo As String
    Public incidentType As New bigClass.incidentType
    Public status As String
    Public priority As String
    Public created As DateTime
    Public isNew As Boolean
    Public dueDate As DateTime = Today.ToShortDateString
    Public lastUpdated As DateTime
    Public updatedby As New agent
    Public closedby As New agent
    Public resolvedby As New agent
    Public reassign As Boolean = False
    Public budget As String
    Public dueDateOverride As Boolean = False
    Public isEditted As Boolean = False
    Public CCs As String
    Public isUpdate As Boolean = False
    Private connStr As String = "Data Source=irat-srv-acc01;Initial Catalog=irl_sd;Integrated Security=false;user id=sa;password=7mmT@XAy"

    Sub New()

    End Sub
    Sub New(tid As Integer)
        ticketid = tid

        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim lastAgentID As Integer = 0
            Dim sqlcommand As New SqlCommand("select * from view_tickets where id=" & ticketid, sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            While dbread.Read
                subject = dbread("subject")
                teamID = dbread("teamID")
                assignedAgent.firstn = dbread("firstn")
                assignedAgent.lastn = dbread("lastn")
                assignedAgent.id = dbread("agentid")
                description = dbread("description")
                filepath = dbread("filepath")
                fileName = dbread("fileName")
                incidentType.name = dbread("incidentType")
                incidentType.typeid = dbread("incident_Type")
                linkid = dbread("linkid")
                priority = dbread("priority")
                status = dbread("status")
                ticketid = dbread("id")
                created = dbread("created")
                dueDate = dbread("dueDate")
                lastUpdated = dbread("lastupdated")
                budget = dbread("budget")
                updatedby.firstn = dbread("updatedByfirstn")
                updatedby.lastn = dbread("updatedBylastn")
                updatedby.id = dbread("updatedby")
                requester = New agent(CType(dbread("reqid"), Integer))
                isEditted = dbread("editted")
                resolvedby = New agent(CType(dbread("resolvedby"), Integer))
                closedby = New agent(CType(dbread("closedby"), Integer))
            End While
            dbread.Close()

            ''reset open flag

            sqlcommand = New SqlCommand("update tbl_tickets set [new]='false' where id=" & ticketid, sqlcon)
            sqlcommand.ExecuteNonQuery()

        Finally
            sqlcon.Dispose()

        End Try
    End Sub
    Public Function trail() As List(Of ticket)
        trail = New List(Of ticket)
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim lastAgentID As Integer = 0
            Dim sqlcommand As New SqlCommand("select * from view_tickets where linkid=" & ticketid & " order by created DESC", sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            While dbread.Read
                Dim tTix As New ticket
                With tTix
                    .subject = dbread("subject")
                    .assignedAgent.firstn = dbread("firstn")
                    .assignedAgent.lastn = dbread("lastn")
                    .updatedby.firstn = dbread("updatedByfirstn")
                    .updatedby.lastn = dbread("updatedBylastn")
                    .updatedby.id = dbread("updatedby")
                    .assignedAgent.id = dbread("agentID")
                    .description = dbread("description")
                    .filepath = dbread("filepath")
                    .fileName = dbread("fileName")
                    .incidentType.name = dbread("incidentType")
                    .incidentType.typeid = dbread("incident_Type")
                    .linkid = dbread("linkid")
                    .priority = dbread("priority")
                    .status = dbread("status")
                    .ticketid = dbread("id")
                    .created = dbread("created")
                    .dueDate = dbread("dueDate")
                    .lastUpdated = CType(dbread("lastupdated"), DateTime)
                End With
                trail.Add(tTix)
            End While
            dbread.Close()

        Finally
            sqlcon.Dispose()

        End Try

    End Function

    Public Function getMyStats(Optional agentid As Integer = -1, Optional teamid As Integer = -1) As tixStat
        getMyStats = New tixStat
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim lastAgentID As Integer = 0

            Dim sql As String = "select count(*) as cnt from view_tickets where linkid=-1 and status='pending'"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            Dim sqlcommand As New SqlCommand(sql, sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            While dbread.Read
                getMyStats.unresolve = dbread("cnt")
            End While
            dbread.Close()

            'overdue
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and duedate < '" & Now & "' and (status <> 'closed' and status <> 'resolved')"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getMyStats.overdue = dbread("cnt")
            End While
            dbread.Close()

            'due today
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and duedate = '" & Today.ToShortDateString & "' and (status <> 'closed' and status <> 'resolved')"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getMyStats.duetoday = dbread("cnt")
            End While
            dbread.Close()


            'open
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and new='true'"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getMyStats.open = dbread("cnt")
            End While
            dbread.Close()

            'unassigned
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and agentid=-1"
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getMyStats.unassigned = dbread("cnt")
            End While
            dbread.Close()
        Finally
            sqlcon.Dispose()

        End Try

    End Function

    Public Function getTrendingStats(Optional agentid As Integer = -1, Optional teamid As Integer = -1) As tixStat
        getTrendingStats = New tixStat
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim lastAgentID As Integer = 0

            Dim sql As String = "select count(*) as cnt from view_tickets where linkid=-1 and datepart(month,lastupdated)=" & Today.Month & " and datepart(year,lastupdated)=" & Today.Year & " and status='pending'"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            Dim sqlcommand As New SqlCommand(sql, sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            While dbread.Read
                getTrendingStats.unresolve = dbread("cnt")
            End While
            dbread.Close()

            'overdue
            sql = "select count(*) as cnt from view_tickets where linkid=-1  and datepart(month,lastupdated)=" & Today.Month & " and datepart(year,lastupdated)=" & Today.Year & " and duedate < '" & Now & "' and (status <> 'closed' and status <> 'resolved')"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getTrendingStats.overdue = dbread("cnt")
            End While
            dbread.Close()

            'due today
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and  datepart(month,lastupdated)=" & Today.Month & " and datepart(year,lastupdated)=" & Today.Year & " and duedate = '" & Today.ToShortDateString & "' and (status <> 'closed' and status <> 'resolved')"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getTrendingStats.duetoday = dbread("cnt")
            End While
            dbread.Close()


            'open
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and datepart(month,lastupdated)=" & Today.Month & " and datepart(year,lastupdated)=" & Today.Year & " and new='true'"
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getTrendingStats.open = dbread("cnt")
            End While
            dbread.Close()

            'close
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and datepart(month,lastupdated)=" & Today.Month & " and datepart(year,lastupdated)=" & Today.Year & " and (status='closed')"
            If agentid > 0 Then
                sql &= " and closedby=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getTrendingStats.closed = dbread("cnt")
            End While
            dbread.Close()

            'resolve
            sql = "select count(*) as cnt from view_tickets where linkid=-1 and datepart(month,lastupdated)=" & Today.Month & " and datepart(year,lastupdated)=" & Today.Year & " and (status='resolved')"
            If agentid > 0 Then
                sql &= " and resolvedby=" & agentid
            End If
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            sqlcommand = New SqlCommand(sql, sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                getTrendingStats.resolved = dbread("cnt")
            End While
            dbread.Close()
        Finally
            sqlcon.Dispose()

        End Try

    End Function

    Public Function viewQueue(Optional teamid As Integer = 0, Optional agentid As Integer = -1, Optional created As String = "", Optional dueby As String = "", Optional status As String = "open", Optional priority As String = "", Optional typez As Integer = -1) As List(Of ticket)
        Dim sqlcon As New SqlConnection(connStr)
        viewQueue = New List(Of ticket)

        Try
            sqlcon.Open()
            Dim sql As String = "select * from view_tickets where linkid=-1 "
            If teamid > 0 Then
                sql &= " and teamid=" & teamid
            End If

            If agentid > 0 Then
                sql &= " and (agentid=" & agentid & " or reqid=" & agentid & ")"
            End If
            If status <> "" Then
                sql &= " and status='" & status & "' "
            Else
                sql &= " and (lower(status) not like 'closed' and lower(status) not like 'resolved') "
            End If
            If priority <> "" Then
                sql &= " and priority='" & priority & "' "
            End If
            If typez > 0 Then
                sql &= " and incident_type=" & typez
            End If

            'filter created
            If created = "last7" Then
                sql &= " and (cast(CONVERT(varchar(10), created, 120) as datetime) >='" & Today.AddDays(-7).ToShortDateString & "' and cast(CONVERT(varchar(10), created, 120) as datetime) <='" & Today.ToShortDateString & "')"
            ElseIf created = "last30" Then
                sql &= " and (cast(CONVERT(varchar(10), created, 120) as datetime) >='" & Today.AddDays(-30).ToShortDateString & "' and cast(CONVERT(varchar(10), created, 120) as datetime) <= CONVERT(varchar(10), '" & Today.ToShortDateString & "', 120))"
            ElseIf created = "last60" Then
                sql &= " and (cast(CONVERT(varchar(10), created, 120) as datetime) >='" & Today.AddDays(-60).ToShortDateString & "' and cast(CONVERT(varchar(10), created, 120) as datetime) <='" & Today.ToShortDateString & "')"
            ElseIf created = "last120" Then
                sql &= " and (cast(CONVERT(varchar(10), created, 120) as datetime) >='" & Today.AddDays(-120).ToShortDateString & "' and cast(CONVERT(varchar(10), created, 120) as datetime) <='" & Today.ToShortDateString & "')"
            Else
                sql &= " and (cast(CONVERT(varchar(10), created, 120) as datetime) >='" & Today.AddDays(-30).ToShortDateString & "' and cast(CONVERT(varchar(10), created, 120) as datetime) <='" & Today.ToShortDateString & "')"
            End If

            'filter due by
            If dueby = "overdue" Then
                sql &= " and (duedate <='" & Now.ToShortDateString & "') and (status <> 'closed' and status <> 'resolved') "
            ElseIf dueby = "today" Then
                sql &= " and (duedate ='" & Today.ToShortDateString & "') and (status <> 'closed' and status <> 'resolved') "
            ElseIf dueby = "tomorrow" Then
                sql &= " and (duedate ='" & Today.AddDays(1).ToShortDateString & "') and (status <> 'closed' and status <> 'resolved') "
            ElseIf dueby = "nextweek" Then
                sql &= " and (duedate >='" & Today.ToShortDateString & "') and (duedate <='" & Today.AddDays(7).ToShortDateString & "') and (status <> 'closed' and status <> 'resolved') "
            ElseIf dueby = "thismonth" Then
                sql &= " and (datepart(month,duedate) ='" & Today.Month & "') and (datepart(year,duedate) ='" & Today.Year & "') and (status <> 'closed' and status <> 'resolved') "
            ElseIf dueby = "nextmonth" Then
                sql &= " and (datepart(month,duedate) ='" & Today.AddMonths(1).Month & "') and (datepart(year,duedate) ='" & Today.AddMonths(1).Year & "') and (status <> 'closed' and status <> 'resolved') "
            End If


            sql &= " order by lastupdated DESC"

            Dim lastAgentID As Integer = 0
            Dim sqlcommand As New SqlCommand(sql, sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            While dbread.Read
                Dim tTix As New ticket
                With tTix
                    .subject = dbread("subject")
                    .assignedAgent.firstn = dbread("firstn")
                    .assignedAgent.lastn = dbread("lastn")
                    .assignedAgent.id = dbread("agentID")
                    .updatedby.firstn = dbread("updatedByfirstn")
                    .updatedby.lastn = dbread("updatedBylastn")
                    .updatedby.id = dbread("updatedby")
                    .description = dbread("description")
                    .filepath = dbread("filepath")
                    .fileName = dbread("fileName")
                    .incidentType.name = dbread("incidentType")
                    .incidentType.typeid = dbread("incident_Type")
                    .linkid = dbread("linkid")
                    .priority = dbread("priority")
                    .status = dbread("status")
                    .ticketid = dbread("id")
                    .created = dbread("created")
                    .dueDate = dbread("dueDate")
                    .lastUpdated = CType(dbread("lastupdated"), DateTime)
                    .requester.id = dbread("reqid")
                    .requester.firstn = dbread("req_first")
                    .isEditted = dbread("editted")
                    .requester.lastn = dbread("req_last")
                End With
                viewQueue.Add(tTix)
            End While
            dbread.Close()

        Finally
            sqlcon.Dispose()

        End Try

    End Function

    Public Function viewQueueAttention(Optional agentid As Integer = -1, Optional teams As List(Of Integer) = Nothing) As List(Of ticket)
        Dim sqlcon As New SqlConnection(connStr)
        viewQueueAttention = New List(Of ticket)

        Try
            sqlcon.Open()
            Dim sql As String = "select * from view_tickets where linkid=-1 and (status='open' or new='true') "
            If agentid > 0 Then
                sql &= " and agentid=" & agentid
            End If

            If teams IsNot Nothing Then
                For Each tid In teams
                    sql &= " or teamid=" & tid
                Next
                '  sql As String = "select * from view_tickets where linkid=-1 and (status='open' or new='true') "
            End If

            sql &= " order by lastupdated DESC"

            Dim lastAgentID As Integer = 0
            Dim sqlcommand As New SqlCommand(sql, sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            While dbread.Read
                Dim tTix As New ticket
                With tTix
                    .subject = dbread("subject")
                    .assignedAgent.firstn = dbread("firstn")
                    .assignedAgent.lastn = dbread("lastn")
                    .assignedAgent.id = dbread("agentID")
                    .updatedby.firstn = dbread("updatedByfirstn")
                    .updatedby.lastn = dbread("updatedBylastn")
                    .updatedby.id = dbread("updatedby")
                    .description = dbread("description")
                    .filepath = dbread("filepath")
                    .fileName = dbread("fileName")
                    .incidentType.name = dbread("incidentType")
                    .incidentType.typeid = dbread("incident_Type")
                    .linkid = dbread("linkid")
                    .priority = dbread("priority")
                    .status = dbread("status")
                    .ticketid = dbread("id")
                    .created = dbread("created")
                    .dueDate = dbread("dueDate")
                    .lastUpdated = CType(dbread("lastupdated"), DateTime)
                    .requester.id = dbread("reqid")
                    .requester.firstn = dbread("req_first")
                    .requester.lastn = dbread("req_last")
                End With
                viewQueueAttention.Add(tTix)
            End While
            dbread.Close()

        Finally
            sqlcon.Dispose()

        End Try

    End Function

    Sub newTicket()
        Dim sqlcon As New SqlConnection(connStr)


        sqlcon.Open()
        Dim sqlcommand As New SqlCommand("newTicket", sqlcon)
        sqlcommand.CommandType = Data.CommandType.StoredProcedure

        sqlcommand.Parameters.AddWithValue("@ticketid", ticketid)
        sqlcommand.Parameters.AddWithValue("@subject", subject)
        sqlcommand.Parameters.AddWithValue("@descr", description)
        sqlcommand.Parameters.AddWithValue("@filepath", filepath)
        sqlcommand.Parameters.AddWithValue("@fileName", fileName)
        sqlcommand.Parameters.AddWithValue("@linkid", linkid)
        sqlcommand.Parameters.AddWithValue("@requester", requester.id)
        ' assignedAgent.id = -1
        sqlcommand.Parameters.AddWithValue("@assigned", assignedAgent.id)
        sqlcommand.Parameters.AddWithValue("@teamid", teamID)
        sqlcommand.Parameters.AddWithValue("@incidenttype", incidentType.typeid)
        sqlcommand.Parameters.AddWithValue("@status", status)
        sqlcommand.Parameters.AddWithValue("@priority", priority)
        dueDate = IIf((dueDateOverride = False), calculateDueDate, dueDate)
        sqlcommand.Parameters.AddWithValue("@dueDate", dueDate)
        sqlcommand.Parameters.AddWithValue("@updatedby", updatedby.id)
        sqlcommand.Parameters.AddWithValue("@budget", budget)

        Dim param As New SqlParameter
        param.ParameterName = "returnV"
        param.Direction = Data.ParameterDirection.ReturnValue

        sqlcommand.Parameters.Add(param)

        sqlcommand.ExecuteNonQuery()

        ticketid = sqlcommand.Parameters("returnV").Value
        If isUpdate = False Then
            If assignedAgent.id > 0 Then
                '   sendNotify()
            ElseIf teamID > 0 Then
                sendRequesterNotify()
                TeamNotify()
            End If

            '' 'if open 
        ElseIf isUpdate = True Then
            'ticket still open and is being updated with a comment
            If assignedAgent.id > 0 And (updatedby.id = requester.id) Then
                'is assigned to an agent, and update sent by requester... notify agent
                sendUpdateNotify("single", assignedAgent.id)
            ElseIf assignedAgent.id < 0 And (updatedby.id = requester.id) Then
                'notify entire team
                sendUpdateNotify("team", teamID)
            Else
                'notify requester
                sendUpdateNotify("single", requester.id)
            End If
        Else

            'sendUpdateNotify()
        End If


        sqlcon.Dispose()


    End Sub
    Private Sub sendCCNotify()
        Try
            Dim Message As New Net.Mail.MailMessage
            Dim smtp1 As New Net.Mail.SmtpClient("10.206.100.111", 25)
            Dim myCredential As New Net.NetworkCredential("iratdesk", "k33p1ts1mpl3@")

            Message.From = New Net.Mail.MailAddress("iratdesk@islandroutes.com", "Island Routes Helpdesk")

            Dim _toAddrs() = CCs.Split(",")
            For Each _toAddr In _toAddrs
                Dim agt = New agent(CType(_toAddr, Integer))
                If Not agt.email.Trim = "" Then
                    Message.To.Add(New Net.Mail.MailAddress((agt.email.Trim)))
                End If
            Next

            Message.Subject = "You have been CC'd - [#" & ticketid & "] " & subject
            Message.IsBodyHtml = True
            smtp1.EnableSsl = True
            smtp1.Port = 587
            Message.Body = "<font style='font-family:arial;font-size:10pt'>"
            Message.Body = "Hello,<br><br>You have been CC'd on this request. Please follow the link below to view the ticket.<br><br><b>Ticket ID:</b> " & ticketid & "<br><b>Subject:</b> " & subject & "<br><b>Requester: </b>" & requester.fullname & "<br><b>Priority: </b>" & priority & "<br><b>Type: </b>" & New bigClass.incidentType(incidentType.typeid).name & IIf(budget <> "", "<br><b>Budget: </b>" & budget, "") & "<br><br>" & description & "<br><br>You can check the status or reply to this ticket at: <br><a href='http://iratdesk.corp.islandroutes.com'>Access IRAT Service Desk</a><br><br><br>Regards,"
            Message.Body &= "</font>"
            '  smtp1.UseDefaultCredentials = False
            smtp1.Credentials = myCredential

            smtp1.Send(Message)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub sendRequesterNotify()
        Try
            Dim Message As New Net.Mail.MailMessage
            Dim agt = New agent(requester.id)
            Dim smtp1 As New Net.Mail.SmtpClient("10.206.100.111", 25)
            Dim myCredential As New Net.NetworkCredential("iratdesk", "k33p1ts1mpl3@")
            Message.From = New Net.Mail.MailAddress("iratdesk@islandroutes.com", "Island Routes Helpdesk")
            Dim teamS = New groups(teamID)
            Message.To.Add(New Net.Mail.MailAddress((agt.email)))

            Message.Subject = "Ticket Assigned - [#" & ticketid & "] " & subject
            Message.IsBodyHtml = True
            smtp1.EnableSsl = False
            smtp1.Port = 25
            Message.Body = "<font style='font-family:arial;font-size:10pt'>"
            Message.Body &= "Dear " & agt.fullname & ",<br><br>Thank you for your request. This is an automated response confirming the receipt of your ticket. For your records, the details of the ticket are listed below. When replying, please make sure that the ticket ID is kept in the subject line to ensure that your replies are tracked appropriately.<br><br><b>Ticket ID:</b> " & ticketid & "<br><b>Subject:</b> " & subject & "<br><b>Team Assigned: </b>" & teamS.groupName & "<br><b>Priority: </b>" & priority & "<br><b>Type: </b>" & New bigClass.incidentType(incidentType.typeid).name & IIf(budget <> "", "<br><b>Budget: </b>" & budget, "") & "<br><br>" & description & "..............................................................................................................................................................<br>To view the status of the ticket or add comments, please visit the link below <a href='http://iratdesk.corp.islandroutes.com/'>Access IRAT Service Desk</a><br><br><br>Regards,"
            Message.Body &= "</font>"
            '  smtp1.UseDefaultCredentials = False
            smtp1.Credentials = myCredential

            smtp1.Send(Message)
        Catch ex As Exception

        End Try
    End Sub
    Public Sub TeamReminder(grpID As Integer)
        Try
            Dim teamS = New groups(grpID)

            Dim remB As String = ""
            remB &= "<table style='width:100%;border-spacing:5px;border-collapse:seperate'>"
            remB &= "<tr style='font-size:10pt;font-family:arial;font-weight:bold;background-color:yellow'><td>Ticket #</td><td>Subject</td><td>Requester</td><td>Assigned To</td><td>Last Updated</td>"
            Dim isEmpty = True
            For Each tix In viewQueue(grpID)
                remB &= "<tr style='font-size:10pt;font-family:arial;font-weight:normal'><td>" & tix.ticketid & "</td><td>" & tix.subject & "</td><td>" & tix.requester.firstn & " " & tix.requester.lastn & "</td><td>" & tix.assignedAgent.firstn & " " & tix.assignedAgent.lastn & "</td><td>" & tix.lastUpdated.ToLongDateString & "</td>"
                isEmpty = False
            Next
            remB &= "</table>"

            If Not isEmpty Then
                For Each agt In teamS.getMembers

                    Dim Message As New Net.Mail.MailMessage

                    Dim smtp1 As New Net.Mail.SmtpClient("10.206.100.111", 25)
                    Dim myCredential As New Net.NetworkCredential("iratdesk", "k33p1ts1mpl3@")

                    Message.From = New Net.Mail.MailAddress("iratdesk@islandroutes.com", "Island Routes Helpdesk")

                    Message.Subject = "Outstanding Tickets for Team '" & teamS.groupName & "'"
                    Message.IsBodyHtml = True
                    smtp1.EnableSsl = False
                    smtp1.Port = 25
                    Message.Body = "<font style='font-family:arial;font-size:10pt'>"

                    smtp1.Credentials = myCredential
                    Message.To.Add(New Net.Mail.MailAddress((agt.email)))
                    Message.Body &= "Dear " & agt.fullname & ",<br><br>There are outstanding tickets still in queue for the team '" & teamS.groupName & "', of which you are a member. Please visit your queue and provide necessary updates for the below tickets.<br><br>"
                    Message.Body &= remB
                    '<b>Ticket ID:</b> " & ticketid & "<br><b>Subject:</b> " & subject & "<br><b>Requester: </b>" & requester.fullname & "<br><b>Priority: </b><br><br>.................................................................................................................................." & description & 
                    Message.Body &= ".........................................................................................................................................................<br>You can check the status Or reply to this ticket at: <br><a href='http://iratdesk.corp.islandroutes.com/'>Access IRAT Service Desk</a><br><br><br>Regards,"
                    Message.Body &= "</font>"
                    smtp1.Send(Message)
                Next
                '  smtp1.UseDefaultCredentials = False
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub TeamNotify()
        Try
            Dim teamS = New groups(teamID)

            For Each agt In teamS.getMembers

                Dim Message As New Net.Mail.MailMessage

                Dim smtp1 As New Net.Mail.SmtpClient("10.206.100.111", 25)
                Dim myCredential As New Net.NetworkCredential("iratdesk", "k33p1ts1mpl3@")

                Message.From = New Net.Mail.MailAddress("iratdesk@islandroutes.com", "Island Routes Helpdesk")

                Message.Subject = "Team '" & teamS.groupName & "' assignment - [#" & ticketid & "] " & subject
                Message.IsBodyHtml = True
                smtp1.EnableSsl = False
                smtp1.Port = 25
                Message.Body = "<font style='font-family:arial;font-size:10pt'>"

                smtp1.Credentials = myCredential
                Message.To.Add(New Net.Mail.MailAddress((agt.email)))
                Message.Body &= "Dear " & agt.fullname & ",<br><br>A new ticket has been assigned to the team '" & teamS.groupName & "', of which you are a member. Please follow the link below to view the ticket.<br><br><b>Ticket ID:</b> " & ticketid & "<br><b>Subject:</b> " & subject & "<br><b>Requester: </b>" & requester.fullname & "<br><b>Priority: </b><br><br>.................................................................................................................................." & description & ".........................................................................................................................................................<br>You can check the status or reply to this ticket at: <br><a href='http://iratdesk.corp.islandroutes.com/'>Access IRAT Service Desk</a><br><br><br>Regards,"
                Message.Body &= "</font>"
                smtp1.Send(Message)
            Next
            '  smtp1.UseDefaultCredentials = False


        Catch ex As Exception

        End Try
    End Sub
    Private Sub sendUpdateNotify(entity As String, entityid As Integer)
        Try
            Dim Message As New Net.Mail.MailMessage


            Dim smtp1 As New Net.Mail.SmtpClient("10.206.100.111", 25)
            Dim myCredential As New Net.NetworkCredential("iratdesk", "k33p1ts1mpl3@")

            Message.From = New Net.Mail.MailAddress("iratdesk@islandroutes.com", "Island Routes Helpdesk")

            If entity = "single" Then
                Dim ent As New agent(entityid)
                Message.To.Add(New Net.Mail.MailAddress((ent.email)))

                Message.Subject = "Ticket Updated - [#" & linkid & "] " & subject & " - " & System.Globalization.CultureInfo.CurrentUICulture.TextInfo.ToTitleCase(status)
                Message.IsBodyHtml = True
                smtp1.EnableSsl = False
                smtp1.Port = 25
                Message.Body = "<font style='font-family:arial;font-size:10pt'>"
                Message.Body &= "Dear " & ent.fullname & ",<br><br>Your ticket has been updated. Please follow the link below to view the ticket.<br><br><b>Ticket ID:</b> " & linkid & "<br><b>Subject:</b> " & subject & "<br><b>Reply From: </b>" & updatedby.fullname & "<br><b>Status:</b> " & System.Globalization.CultureInfo.CurrentUICulture.TextInfo.ToTitleCase(status) & "<br><br>" & description & "<br><br>You can check the status or reply to this ticket at: <br><a href='http://iratdesk.corp.islandroutes.com'>Access IRAT Service Desk</a><br><br><br>Regards,"
                Message.Body &= "</font>"
            Else
                Dim teamS = New groups(entityid)

                For Each agt In teamS.getMembers

                    Message.Subject = "Team '" & teamS.groupName & "' Ticket Updated - [#" & ticketid & "] " & subject
                    Message.IsBodyHtml = True
                    smtp1.EnableSsl = False
                    smtp1.Port = 25
                    Message.Body = "<font style='font-family:arial;font-size:10pt'>"
                    Message.To.Add(New Net.Mail.MailAddress((agt.email)))
                    Message.Body &= "Dear " & agt.fullname & ",<br><br>Your team ticket has been updated. Please follow the link below to view the ticket.<br><br><b>Ticket ID:</b> " & linkid & "<br><b>Subject:</b> " & subject & "<br><b>Reply From: </b>" & updatedby.fullname & "<br><b>Status:</b> " & System.Globalization.CultureInfo.CurrentUICulture.TextInfo.ToTitleCase(status) & "<br><br>" & description & "<br><br>You can check the status or reply to this ticket at: <br><a href='http://iratdesk.corp.islandroutes.com'>Access IRAT Service Desk</a><br><br><br>Regards,"
                    Message.Body &= "</font>"

                Next
            End If
            '  smtp1.UseDefaultCredentials = False
            smtp1.Credentials = myCredential

            smtp1.Send(Message)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub sendReAssignNotify()
        Try
            Dim Message As New Net.Mail.MailMessage
            Dim agt = New agent(assigner)
            'Dim smtp1 As New Net.Mail.SmtpClient("relay.jangosmtp.net", 587)
            'Dim myCredential As New Net.NetworkCredential("uwmoroghe", "Godz4ever")
            Dim smtp1 As New Net.Mail.SmtpClient("smtp.gmail.com", 587)
            Dim myCredential As New Net.NetworkCredential("islandroutes.sd", "Isl@ndR0ut3s")

            Message.From = New Net.Mail.MailAddress("islandroutes.sd@gmail.com", "Island Routes SD")
            Message.To.Add(New Net.Mail.MailAddress((agt.email)))

            Message.Subject = "Ticket Reassigned - [#" & linkid & "] " & subject & " - " & System.Globalization.CultureInfo.CurrentUICulture.TextInfo.ToTitleCase(status)
            Message.IsBodyHtml = True
            smtp1.EnableSsl = True
            smtp1.Port = 587
            Message.Body = "<font style='font-family:arial;font-size:10pt'>"
            Message.Body &= "Dear " & agt.fullname & ",<br><br>This ticket was reassigned to you by " & updatedby.fullname & ". Please follow the link below To view the ticket.<br><br><b>Ticket ID:</b> " & ticketid & "<br><b>Subject:</b> " & subject & "<br><b>Status:</b> " & System.Globalization.CultureInfo.CurrentUICulture.TextInfo.ToTitleCase(status) & "<br><br>" & description & "<br><br>You can check the status or reply to this ticket at: <br><a href='http://irat-srv-acc01.sri.sandals.com/'>Access IRAT Service Desk</a><br><br><br>Regards,"
            Message.Body &= "</font>"
            '  smtp1.UseDefaultCredentials = False
            smtp1.Credentials = myCredential

            'smtp1.Host = "relay.jangosmtp.net"
            smtp1.Send(Message)
        Catch ex As Exception

        End Try
    End Sub

    Sub updateTicket(Optional isEdit As Boolean = False)
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim sqlcommand As New SqlCommand("newTicket", sqlcon)
            sqlcommand.CommandType = Data.CommandType.StoredProcedure

            sqlcommand.Parameters.AddWithValue("@ticketid", ticketid)
            sqlcommand.Parameters.AddWithValue("@subject", subject)
            sqlcommand.Parameters.AddWithValue("@descr", description)
            sqlcommand.Parameters.AddWithValue("@filepath", filepath)
            sqlcommand.Parameters.AddWithValue("@linkid", linkid)
            sqlcommand.Parameters.AddWithValue("@requester", requester.id)
            sqlcommand.Parameters.AddWithValue("@teamid", teamID)
            sqlcommand.Parameters.AddWithValue("@assigned", assigner)
            sqlcommand.Parameters.AddWithValue("@incidenttype", incidentType.typeid)
            sqlcommand.Parameters.AddWithValue("@status", status)
            sqlcommand.Parameters.AddWithValue("@priority", priority)
            sqlcommand.Parameters.AddWithValue("@updatedby", updatedby.id)
            sqlcommand.Parameters.AddWithValue("@budget", budget)
            sqlcommand.Parameters.AddWithValue("@isEdit", isEdit)

            Dim param As New SqlParameter
            param.ParameterName = "returnV"
            param.Direction = Data.ParameterDirection.ReturnValue

            sqlcommand.Parameters.Add(param)

            sqlcommand.ExecuteNonQuery()

            If reassign Then
                sendReAssignNotify()
            End If
            'ticketid = sqlcommand.Parameters("returnV").Value

        Finally
            sqlcon.Dispose()

        End Try
    End Sub
    Private Function calculateDueDate() As DateTime
        calculateDueDate = Today.AddDays(2)

        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim lastAgentID As Integer = 0
            Dim sqlcommand As New SqlCommand("Select resolve from tbl_sla where priority='" & priority.ToLower & "'", sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            Dim sla() As String
            Dim inc As Integer = 0
            Dim elap As String = ""

            sla = "|".Split("|")

            While dbread.Read
                sla = CType(dbread("resolve"), String).Split("|")
            End While

            If sla(0) <> "" Then
                inc = sla(0)
                elap = sla(1)

                If elap = "Mins" Then
                    calculateDueDate = DateTime.Now.AddMinutes(inc)
                ElseIf elap = "Hrs" Then
                    calculateDueDate = DateTime.Now.AddHours(inc)
                ElseIf elap = "Days" Then
                    calculateDueDate = DateTime.Now.AddDays(inc)
                ElseIf elap = "Mons" Then
                    calculateDueDate = Today.AddMonths(inc)
                End If
            End If

        Finally
            sqlcon.Dispose()

        End Try

    End Function
    Private Function assigner() As Integer
        assigner = 0

        Dim sqlcon As New SqlConnection(connStr)

        If assignedAgent.id < 1 Then
            Try
                sqlcon.Open()
                Dim lastAgentID As Integer = 0
                Dim sqlcommand As New SqlCommand("select * from tbl_incidentMapping where incidentID=" & incidentType.typeid, sqlcon)

                Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
                While dbread.Read
                    If dbread("objectType") = "group" Then
                        assigner = -1 'assignWho(dbread("objectID"))
                        teamID = dbread("objectID")
                    Else
                        assigner = dbread("objectID")
                    End If
                End While

            Finally
                sqlcon.Dispose()
            End Try
        End If

        If assigner = 0 And assignedAgent.id < 0 Then
            assigner = -1 ' assignWho(teamID)
        ElseIf assignedAgent.id > 0 Then
            assigner = assignedAgent.id
        End If
    End Function
    Private Function assignWho(team_id As Integer) As Integer

        ''run round robin against team
        Dim sqlcon As New SqlConnection(connStr)

        Try
            sqlcon.Open()
            Dim lastAgentID As Integer = 0
            Dim sqlcommand As New SqlCommand("select agentID from tbl_tickets where teamID=" & team_id, sqlcon)

            Dim dbread As SqlDataReader = sqlcommand.ExecuteReader
            While dbread.Read
                lastAgentID = dbread("agentID")
            End While
            dbread.Close()

            '' get next agent from team
            Dim nextAgentID As Integer = 0
            sqlcommand = New SqlCommand("select top 1 agentID from tblGroupMapping where groupid=" & teamID & " and agentid <>" & lastAgentID & " and agentid <> " & requester.id & " and agentID IN(select id from tbl_agents) ORDER BY NEWID()", sqlcon)

            dbread = sqlcommand.ExecuteReader
            While dbread.Read
                nextAgentID = dbread("agentID")
            End While
            dbread.Close()

            If nextAgentID = 0 Then
                sqlcommand = New SqlCommand("select top 1 agentID from tblgroupMapping where groupid=" & team_id & " and agentID IN(select id from tbl_agents) order by agentid asc", sqlcon)

                dbread = sqlcommand.ExecuteReader
                While dbread.Read
                    nextAgentID = dbread("agentID")
                End While
                dbread.Close()
            End If

            assignWho = nextAgentID
        Finally
            sqlcon.Dispose()

        End Try
    End Function

    Structure tixStat
        Dim unresolve As Integer
        Dim overdue As Integer
        Dim duetoday As Integer
        Dim open As Integer
        Dim closed As Integer
        Dim resolved As Integer
        Dim unassigned As Integer
    End Structure
End Class
