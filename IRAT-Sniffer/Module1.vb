Imports System.IO
Imports System.Threading
Imports EAGetMail

Module Module1
    Dim curpath As String = Directory.GetCurrentDirectory()
    Dim mailbox As String = String.Format("{0}\\inbox", Directory.GetCurrentDirectory())
    Dim connString As String = "Data Source=irat-srv-acc01;Initial Catalog=irl_sd;Integrated Security=false;user id=sa;password=7mmT@XAy"

    Dim aTimer As New System.Timers.Timer
    Dim remTimer As New System.Timers.Timer
    Dim mailBoxes As New List(Of mBox)

    Sub Main()
        mailBoxes.Add(New mBox("systems-support@islandroutes.com", "Systems-support", "k33p1ts1mpl3@", 16))
        mailBoxes.Add(New mBox("it-support@islandroutes.com", "it-support", "k33p1ts1mpl3@", 3))
        mailBoxes.Add(New mBox("iratdesk@islandroutes.com", "iratdesk", "k33p1ts1mpl3@", 3))
        mailBoxes.Add(New mBox("marketing-support@islandroutes.com", "marketing-support", "k33p1ts1mpl3@", 4))
        mailBoxes.Add(New mBox("products-support@islandroutes.com", "products-support", "k33p1ts1mpl3@", 6))
        mailBoxes.Add(New mBox("hr-support@islandroutes.com", "hr-support", "k33p1ts1mpl3@", 15))

        Console.WriteLine("================================\r\rIRAT Desk Inbox Sniffer\r\r================================")



        TimerCall()
        aTimer.AutoReset = True
        aTimer.Interval = 10000 '10 seconds
        AddHandler aTimer.Elapsed, AddressOf TimerCall
        aTimer.Start()
        '  Dim t As New Timer(TimerCall, Nothing, 0, 2000)

        'REMINDER TIMER
        RemTimerCall()
        remTimer.AutoReset = True
        remTimer.Interval = 1000 * 60 * 60
        AddHandler remTimer.Elapsed, AddressOf RemTimerCall
        remTimer.Start()
        Console.ReadLine()
    End Sub

    Private Sub RemTimerCall()
        If Date.Now.Hour = 9 Or Date.Now.Hour = 16 Then
            Dim Grp As New groups
            For Each tGrp In Grp.getGroups
                Console.WriteLine("Checking Group: " & tGrp.groupName)
                Dim tix As New ticket
                tix.TeamReminder(tGrp.groupid)
            Next
        End If
    End Sub

    Private Sub TimerCall()

        For Each mb In mailBoxes

            'let's check each mailbox

            Dim oServer As New MailServer("10.206.100.111",
                           mb.emailid, mb.pwd, ServerProtocol.Pop3)
            Dim oClient As New MailClient("TryIt")
            oServer.SSLConnection = True
            oServer.User = mb.uname
            oServer.Password = mb.pwd
            oServer.Port = 995

            Dim mdC = mb.emailid

            Try

                oClient.Connect(oServer)
                Console.WriteLine("Connected to: " & mb.uname)
                Dim infos As MailInfo() = oClient.GetMailInfos()
                For i As Integer = 0 To infos.Length - 1

                    Dim info As MailInfo = infos(i)
                    ' Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",info.Index, info.Size, info.UIDL)

                    ' Receive email from POP3 server
                    Dim oMail As Mail = oClient.GetMail(info)

                    Console.WriteLine("From: {0}", oMail.From.ToString)
                    Console.WriteLine("Subject: {0}\r\n", oMail.Subject)

                    '
                    Dim requester As New agent(oMail.From.Address, oMail.From.Name.First, oMail.From.Name.Last)

                    Console.WriteLine("Requester: " & requester.fullname)


                    'Get Subject and determine if new ticket or update
                    Dim subj As String = oMail.Subject.Split("(Trial Version)")(0)
                    ' Dim tes As String = "red meat - [#222] last part"
                    Dim inx As Integer = subj.IndexOf("[#")
                    Dim ticid As String = ""
                    Dim ticket_found As Boolean = False

                    If inx > 0 Then
                        Dim linx As Integer = subj.IndexOf("]")
                        ticid = subj.Substring(inx + 2, linx - (inx + 2))
                        ticket_found = True
                    End If

                    If ticket_found = False Then
                        'Let's create the new ticket
                        Dim newTicketThread As New Thread(Sub() new_Ticket(requester, mb.teamid, oMail.Subject.Split("(Trial Version)")(0), oMail.HtmlBody, oMail.Attachments))
                        With newTicketThread
                            .IsBackground = True
                            .Start()
                        End With
                    Else
                        Dim replyTicketThrea As New Thread(Sub() reply_ticket(ticid, requester, oMail.HtmlBody))
                        With replyTicketThrea
                            .IsBackground = True
                            .Start()
                        End With
                    End If


                    '  


                    ' Mark email as deleted from POP3 server.
                    oClient.Delete(info)
                Next


                ' Quit And purge emails marked as deleted from POP3 server.


            Catch ep As Exception

                Console.WriteLine(ep.Message)

            Finally
                oClient.Quit()

            End Try


        Next


    End Sub

    Private Sub reply_ticket(ticketid As Integer, updater As agent, descr As String)
        Dim vTix As New ticket(ticketid)

        Try
            Dim newTicket As New ticket
            Dim teamID As Integer
            Dim agentID As Integer

            vTix.updatedby.id = updater.id
            vTix.updateTicket()

            With newTicket
                .subject = vTix.subject
                .teamID = vTix.teamID
                .assignedAgent.id = vTix.assignedAgent.id
                .updatedby = updater
                .budget = vTix.budget
                .description = descr
                .filepath = ""
                .isUpdate = True
                .fileName = ""
                .incidentType.typeid = vTix.incidentType.typeid
                .status = vTix.status
                .priority = vTix.priority
                .requester = vTix.requester
                .linkid = vTix.ticketid
                .newTicket()
            End With
            ' Response.Redirect("frmQueue.aspx")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub new_Ticket(req As agent, teamid As String, subj As String, descr As String, atts() As Attachment)
        Try
            Dim agentID As Integer
            Dim newTicket As New ticket

            '     ' agentID = IIf(Request.Form("dl_agent") > 0, Request.Form("dl_agent"), -1)

            With newTicket
                .subject = subj
                .teamID = teamid
                .CCs = ""
                .assignedAgent.id = -1
                .description = descr
                .filepath = ""
                .fileName = ""
                .isUpdate = False
                .budget = ""
                .incidentType.typeid = New bigClass.incidentType("Misc", teamid).typeid
                .status = "Open"
                .priority = "Medium"
                .updatedby.id = req.id
                .requester = req
                .linkid = -1
                .newTicket()
            End With

            '  Dim atts() As Attachment = oMail.Attachments

            Dim tempFolder As String = "C:\inetpub\wwwroot\attach"
            Dim count As Integer = atts.Length
            If (Not System.IO.Directory.Exists(tempFolder)) Then
                System.IO.Directory.CreateDirectory(tempFolder)
            End If


            For i As Integer = 0 To count - 1
                Dim attFile As New attachFile
                Dim att As Attachment = atts(i)
                Dim fname As String = Guid.NewGuid.ToString & System.IO.Path.GetExtension(att.Name)
                Dim attname As String = String.Format("{0}\{1}", tempFolder, fname)
                att.SaveAs(attname, True)
                attFile.attName = att.Name
                attFile.attPath = fname
                attFile.addew(newTicket.ticketid)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Class mBox

        Sub New(emailID, username, password, team)

            emailID = emailID
            uname = username
            pwd = password
            teamid = team
        End Sub

        Public emailid As String
        Public uname As String
        Public pwd As String
        Public teamid As Integer
    End Class

End Module
