Imports System.Data.OleDb
Imports System.IO

Public Class clsDataHandler
    Dim strConn As String = ""

    Public Sub New(ByVal strAppType As String)

        If Not strAppType = "" Then
            InitialiseConnection(strAppType)
        End If
    End Sub

    Private Sub InitialiseConnection(ByVal strAppType As String)

        Dim strFullPath As String = ""
        Dim strContents As String
        Dim objReader As StreamReader

        Try
            If strAppType = "B" Or strAppType = "BE" Or strAppType = "M" Then
                strFullPath = Application.StartupPath & "..\TaxcomB.ini" 'after make exe run this
            ElseIf strAppType = "P" Or strAppType = "CP30" Then
                strFullPath = Application.StartupPath & "..\TaxcomP.ini" 'after make exe run this
            End If
            objReader = New StreamReader(strFullPath)
            strContents = objReader.ReadToEnd()
            objReader.Close()
            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strContents

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Function GetDataReader(ByVal strQuery As String) As OleDbDataReader
        Dim dr As OleDbDataReader
        Dim cmd As New OleDbCommand
        With cmd
            ' Create a Connection object
            .Connection = New OleDbConnection(strConn)
            .Connection.Open()
            .CommandText = strQuery
            dr = .ExecuteReader(CommandBehavior.CloseConnection)
        End With
        'If Not dr.HasRows Or dr.RecordsAffected <= 0 Then
        '    dr = Nothing
        'End If
        Return dr
    End Function

    Public Function GetDataReader1(ByVal strQuery As String) As OleDbDataReader
        Dim dr As OleDbDataReader
        Dim cmd As New OleDbCommand
        With cmd
            ' Create a Connection object
            InitialiseConnection("B")
            .Connection = New OleDbConnection(strConn)
            .Connection.Open()
            .CommandText = strQuery
            dr = .ExecuteReader(CommandBehavior.CloseConnection)
        End With
        Return dr
    End Function

    Public Function GetData(ByVal strQuery As String) As DataSet

        Dim ds As New DataSet
        Dim dataConnection As New OleDbConnection
        dataConnection.ConnectionString = strConn
        Try
            Dim cmd As New OleDbCommand(strQuery, dataConnection)
            'If prmOleDb IsNot Nothing Then
            'For Each prmOle As OleDbParameter In prmOleDb
            'If prmOle IsNot Nothing Then cmd.Parameters.Add(prmOle)
            '    Next
            'End If
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds)

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If dataConnection.State = ConnectionState.Open Then dataConnection.Close()
        End Try

        Return ds
    End Function

    Public Function GetData(ByVal strQuery As String, ByVal ParamArray prmOleDb As IDataParameter()) As DataSet

        Dim ds As New DataSet
        Dim dataConnection As New OleDbConnection
        dataConnection.ConnectionString = strConn
        Try
            Dim cmd As New OleDbCommand(strQuery, dataConnection)

            If prmOleDb IsNot Nothing Then
                For Each prmOle As OleDbParameter In prmOleDb
                    If prmOle IsNot Nothing Then cmd.Parameters.Add(prmOle)
                Next
            End If
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds)

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If dataConnection.State = ConnectionState.Open Then dataConnection.Close()
        End Try

        Return ds
    End Function

    Public Function Execute(ByVal strSQL As String) As Integer
        Dim objConn As New OleDbConnection(strConn)
        Dim cmd As OleDbCommand
        Dim intAffectedRow As Integer

        Try
            cmd = New OleDbCommand
            cmd.Connection = objConn
            objConn.Open()
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        Finally
            objConn.Close()
        End Try

        Return intAffectedRow
    End Function

    Public ReadOnly Property OledbConnection()
        Get
            Return strConn
        End Get
    End Property
End Class
