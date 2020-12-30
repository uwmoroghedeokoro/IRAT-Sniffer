Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class bigClass

    Public Class incidentType
        Public typeid As Integer
        Public name As String
        Public parentGrp As New incidentGroup
        Public prefObject As New pObject
        Private connStr As String = "Data Source=irat-srv-acc01;Initial Catalog=irl_sd;Integrated Security=false;user id=sa;password=7mmT@XAy"


        Sub New()

        End Sub
        Function incidentGroups() As List(Of incidentGroup)
            Dim sqlcon As New SqlConnection(connStr)

            incidentGroups = New List(Of incidentGroup)

            Try
                sqlcon.Open()

                Dim sql As String
                sql = "select * from tbl_incidentGroups order by incidentGroup ASC"

                Dim sqlcommand As New SqlCommand(sql, sqlcon)
                Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While dbRead.Read
                    Dim tmp As New incidentGroup
                    With tmp
                        .grpID = dbRead("id")
                        .grpName = dbRead("incidentGroup")
                    End With
                    incidentGroups.Add(tmp)
                End While
                dbRead.Close()

            Finally
                sqlcon.Dispose()

            End Try
        End Function
        Sub New(tID As Integer)
            typeid = tID

            Dim sqlcon As New SqlConnection(connStr)

            Try
                sqlcon.Open()

                Dim sql As String
                sql = "select * from tbl_incidentTypes where id=" & typeid

                Dim sqlcommand As New SqlCommand(sql, sqlcon)
                Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While dbRead.Read
                    typeid = dbRead("id")
                    name = dbRead("incidentType")
                End While
                dbRead.Close()

            Finally
                sqlcon.Dispose()

            End Try
        End Sub
        Sub New(typen As String, teamID As String)
            ' typeid = tID

            Dim sqlcon As New SqlConnection(connStr)

            Try
                sqlcon.Open()

                Dim sql As String
                sql = "select * from tbl_incidentTypes where parentid=" & teamID & " and incidentType ='Misc'"

                Dim sqlcommand As New SqlCommand(sql, sqlcon)
                Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While dbRead.Read
                    typeid = dbRead("id")
                    name = dbRead("incidentType")
                End While
                dbRead.Close()

            Finally
                sqlcon.Dispose()

            End Try
        End Sub
        Public Function getTypes() As List(Of incidentType)
            Dim sqlcon As New SqlConnection(connStr)

            getTypes = New List(Of incidentType)

            Try
                sqlcon.Open()

                Dim sql As String
                sql = "select n.*,i.objectid,i.objectType from tbl_incidentTypes n left outer join tbl_incidentMapping i on i.incidentid=n.id   order by n.incidentType ASC"

                Dim sqlcommand As New SqlCommand(sql, sqlcon)
                Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While dbRead.Read
                    Dim tmp As New incidentType
                    With tmp
                        .typeid = dbRead("id")
                        .name = dbRead("incidentType")
                        .prefObject.objectid = IIf(IsDBNull(dbRead("objectid")), -1, dbRead("objectid"))
                        .prefObject.objtype = IIf(IsDBNull(dbRead("objectType")), "", dbRead("objectType"))
                    End With
                    getTypes.Add(tmp)
                End While
                dbRead.Close()

            Finally
                sqlcon.Dispose()

            End Try
        End Function

        Public Function getTypesByParent(parentID As Integer) As List(Of incidentType)
            Dim sqlcon As New SqlConnection(connStr)

            getTypesByParent = New List(Of incidentType)

            Try
                sqlcon.Open()

                Dim sql As String
                sql = "select * from tbl_IncidentTypes where parentID=" & parentID & " order by incidentType ASC"

                Dim sqlcommand As New SqlCommand(sql, sqlcon)
                Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While dbRead.Read
                    Dim tmp As New incidentType
                    With tmp
                        .typeid = dbRead("id")
                        .name = dbRead("incidentType")
                    End With
                    getTypesByParent.Add(tmp)
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
                Dim sqlcommand As New SqlCommand("addIncType", sqlcon)
                sqlcommand.CommandType = Data.CommandType.StoredProcedure

                sqlcommand.Parameters.AddWithValue("@id", typeid)
                sqlcommand.Parameters.AddWithValue("@name", name)
                sqlcommand.Parameters.AddWithValue("@parentid", parentGrp.grpID)

                Dim param As New SqlParameter
                param.ParameterName = "returnV"
                param.Direction = Data.ParameterDirection.ReturnValue

                sqlcommand.Parameters.Add(param)

                sqlcommand.ExecuteNonQuery()

                ' Dim returnV As Integer
                typeid = sqlcommand.Parameters("returnV").Value

            Finally
                sqlcon.Dispose()

                ' For Each aID In membersByID
                '  addToGroup(aID)
                ' Next
            End Try

        End Sub

        Public Sub mapTeam(incType As Integer, teamID As Integer, objType As String)
            Dim sqlcon As New SqlConnection(connStr)

            Try
                sqlcon.Open()
                Dim sqlcommand As New SqlCommand("mapIncTeam", sqlcon)
                sqlcommand.CommandType = Data.CommandType.StoredProcedure

                sqlcommand.Parameters.AddWithValue("@incID", incType)
                sqlcommand.Parameters.AddWithValue("@teamID", teamID)
                sqlcommand.Parameters.AddWithValue("@objType", objType)

                Dim param As New SqlParameter
                param.ParameterName = "returnV"
                param.Direction = Data.ParameterDirection.ReturnValue

                sqlcommand.Parameters.Add(param)

                sqlcommand.ExecuteNonQuery()

                ' Dim returnV As Integer
                typeid = sqlcommand.Parameters("returnV").Value

            Finally
                sqlcon.Dispose()

                ' For Each aID In membersByID
                '  addToGroup(aID)
                ' Next
            End Try

        End Sub
    End Class

    Structure pObject
        Dim objectid As Integer
        Dim objtype As String
    End Structure
    Structure incidentGroup
        Dim grpID As Integer
        Dim grpName As String
    End Structure
    Public Class SLA
        Public priority As String = "|"
        Public response As String = "|"
        Public resolve As String = "|"
        Private connStr As String = "Data Source=irat-srv-acc01;Initial Catalog=irl_sd;Integrated Security=false;user id=sa;password=7mmT@XAy"


        Sub New()

        End Sub

        Function sla_find(priorty As String) As SLA
            sla_find = New SLA
            sla_find.priority = priorty


            Dim sqlcon As New SqlConnection(connstr)

            Try
                sqlcon.Open()

                Dim sql As String
                sql = "select * from tbl_sla where priority='" & priorty & "'"

                Dim sqlcommand As New SqlCommand(sql, sqlcon)
                Dim dbRead As SqlDataReader = sqlcommand.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While dbRead.Read
                    sla_find.resolve = IIf(IsDBNull(dbRead("resolve")), "|", dbRead("resolve"))
                    sla_find.response = IIf(IsDBNull(dbRead("respond")), "|", dbRead("respond"))
                End While
                dbRead.Close()

            Finally
                sqlcon.Dispose()

            End Try
        End Function

        Public Sub update_sla(priority As String, opt As String, sla As String)
            Dim sqlcon As New SqlConnection(connstr)

            Try
                sqlcon.Open()
                Dim sqlcommand As New SqlCommand("update_sla", sqlcon)
                sqlcommand.CommandType = Data.CommandType.StoredProcedure

                sqlcommand.Parameters.AddWithValue("@priority", priority)
                sqlcommand.Parameters.AddWithValue("@opt", opt)
                sqlcommand.Parameters.AddWithValue("@sla", sla)

                Dim param As New SqlParameter
                param.ParameterName = "returnV"
                param.Direction = Data.ParameterDirection.ReturnValue

                sqlcommand.Parameters.Add(param)

                sqlcommand.ExecuteNonQuery()

                ' Dim returnV As Integer
                '   typeid = sqlcommand.Parameters("returnV").Value

            Finally
                sqlcon.Dispose()

                ' For Each aID In membersByID
                '  addToGroup(aID)
                ' Next
            End Try

        End Sub
    End Class
End Class
