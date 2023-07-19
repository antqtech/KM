#Region " ※※※　　　　  修改記錄　　    　　※※※"
'!--------------------------------------------------------------------------------------------
'! ComponentName : PISCP.DB.classBaseDB
'!
'! FileName : classBaseDB.vb
'!
'! Author : Steven
'! 
'! Environment : O.S. - Windows 7 Professional
'!               Tool - Visual Studio .Net 2013
'!
'! Description : base connect DB class
'!                  & include SqlTransaction
'!
'! Modification Log :
'! Vers   Date             By               Notes
'! -----  ---------------  ---------------  ------------------------------------------------
'! V1.0   2019/09/10       Steven		    Create
'! -----  ---------------  ---------------  ------------------------------------------------
#End Region
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Namespace PISCP.DB
    Public Class ClassBaseDB
        'Private mySystemLog As New SystemLog
        Private myConnection As New SqlConnection
        Private myConnectionString As String
        Private myTransaction As SqlTransaction
        Private myTransactionFlag As Boolean = False


#Region "  Property Attribute  "

        Public Property ConnectionString() As String
            Get
                Return Me.myConnectionString
            End Get

            Set(ByVal Value As String)
                Me.myConnectionString = Value
                Me.Connection = New SqlConnection(Me.myConnectionString)
            End Set
        End Property

        Public Property Connection() As SqlConnection
            Get
                Return Me.myConnection
            End Get

            Set(ByVal Value As SqlConnection)
                Me.myConnection = Value
            End Set
        End Property

        Public ReadOnly Property TransactionFlag() As Boolean
            Get
                Return Me.myTransactionFlag
            End Get
        End Property
#End Region

#Region "  Transaction  "
        Public Sub BeginTransaction(ByVal strTransName As String)
            If Me.myConnection.State = ConnectionState.Closed Then Me.myConnection.Open()
            Me.myTransaction = Me.myConnection.BeginTransaction(strTransName)
            Me.myTransactionFlag = True
        End Sub

        Public Sub BeginTransaction()
            If Me.myConnection.State = ConnectionState.Closed Then Me.myConnection.Open()
            Me.myTransaction = Me.myConnection.BeginTransaction()
            Me.myTransactionFlag = True
        End Sub

        Public Sub Commit()
            If Me.myTransactionFlag Then
                Me.myTransaction.Commit()
            End If
        End Sub

        Public Sub Rollback(ByVal strTransName As String)
            If Me.myTransactionFlag Then
                Me.myTransaction.Rollback(strTransName)
            End If
        End Sub

        Public Sub Rollback()
            If Me.myTransactionFlag Then
                Me.myTransaction.Rollback()
            End If
        End Sub
#End Region

        'Modify Friend To Public
        Public Function BuildQueryCommand(ByVal storedProcName As String, ByVal parameters() As SqlParameter) As SqlCommand
            Dim command As SqlCommand
            Dim parameter As SqlParameter

            If Me.myTransactionFlag Then
                If Me.myTransaction Is Nothing Then
                    Me.myTransactionFlag = False
                ElseIf Me.myTransaction.Connection Is Nothing Then
                    Me.myTransactionFlag = False
                End If
            End If

            If Me.myTransactionFlag Then
                command = New SqlCommand(storedProcName, Me.myTransaction.Connection)
                command.Transaction = Me.myTransaction
            Else
                command = New SqlCommand(storedProcName, Me.myConnection)
            End If
            command.CommandTimeout = Me.myConnection.ConnectionTimeout
            command.CommandType = CommandType.StoredProcedure
            For Each parameter In parameters
                command.Parameters.Add(parameter)
            Next parameter

            Return command
        End Function

        Friend Function BuildIntCommand(ByVal storedProcName As String, ByVal parameters() As SqlParameter) As SqlCommand
            Dim command As SqlCommand

            command = BuildQueryCommand(storedProcName, parameters)
            command.Parameters.Add(New SqlParameter("ReturnValue", SqlDbType.Int, 4, ParameterDirection.ReturnValue, False, 0, 0, String.Empty, DataRowVersion.Default, 0))

            Return command
        End Function

        Public Function RunProcedureAdapter(ByVal storedProcName As String, ByVal parameters() As SqlParameter) As SqlDataAdapter
            Dim returnAdapter As New SqlDataAdapter

            returnAdapter.SelectCommand = BuildIntCommand(storedProcName, parameters)

            Return returnAdapter
        End Function

        Public Function RunProcedure(ByVal storedProcName As String, ByVal parameters() As SqlParameter, ByVal tableName As String) As DataSet
            Dim dataSet As New DataSet
            Dim sqlDA As New SqlDataAdapter

            Try
                sqlDA.SelectCommand = BuildQueryCommand(storedProcName, parameters)
                If sqlDA.SelectCommand.Connection.State = ConnectionState.Closed Then sqlDA.SelectCommand.Connection.Open()
                sqlDA.Fill(dataSet, tableName)
            Finally
                If Not Me.myTransactionFlag Then
                    sqlDA.SelectCommand.Connection.Close()
                    sqlDA.SelectCommand.Dispose()
                    sqlDA.Dispose()
                End If
            End Try

            Return dataSet
        End Function

        Public Sub RunProcedureInsert(ByVal storedProcName As String, ByVal parameters() As SqlParameter)
            Dim command As New SqlCommand

            Try
                command = BuildIntCommand(storedProcName, parameters)
                If command.Connection.State = ConnectionState.Closed Then command.Connection.Open()
                command.ExecuteNonQuery()
            Finally
                If Not Me.myTransactionFlag Then
                    command.Connection.Close()
                    command.Dispose()
                End If
            End Try
        End Sub

    End Class
End Namespace
