#Region " ※※※　　　　  修改記錄　　    　　※※※"
'!--------------------------------------------------------------------------------------------
'! ComponentName : PISCP.DB.classDB
'!
'! FileName : classDB.vb
'!
'! Author : Steven
'! 
'! Environment : O.S. - Windows 7 Professional
'!               Tool - Visual Studio .Net 2013
'!
'! Description : connect DB class
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
    Public Class ClassDB
        Inherits ClassBaseDB

        Public htParam As New Hashtable

        Public Sub New()
            Me.classDB()
        End Sub

        Public Sub New(ByVal strConfigConnStr As String)
            Me.initDB(strConfigConnStr)
        End Sub

        Private Sub classDB()
            Me.initDB("DB_PISCP")
        End Sub

        Private Sub initDB(ByVal strConfigConnStr As String)
            Me.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strConfigConnStr)
            If ConnectionString Is Nothing Then
                Me.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings(strConfigConnStr).ConnectionString
            End If
            Me.Connection = New SqlConnection(ConnectionString)
        End Sub

#Region "   Public method "
        Public Overloads Function RunProcedure(ByVal storedProcName As String, ByVal parameters() As SqlParameter) As SqlDataReader
            Dim returnReader As SqlDataReader
            Dim classNewDB As New ClassBaseDB
            Dim command As New SqlCommand

            classNewDB.ConnectionString = Me.ConnectionString
            command = classNewDB.BuildIntCommand(storedProcName, parameters)
            If command.Connection.State = ConnectionState.Closed Then command.Connection.Open()
            returnReader = command.ExecuteReader()

            Return returnReader
        End Function

        Public Function RunProcedure2(ByVal storedProcNameA As String, ByVal parametersA() As SqlParameter, ByVal storedProcNameB As String, ByVal parametersB() As SqlParameter) As DataSet
            Dim oDataSet As New DataSet
            Dim sqlDAA As New SqlDataAdapter
            Dim sqlDAB As New SqlDataAdapter

            Try
                sqlDAA.SelectCommand = BuildQueryCommand(storedProcNameA, parametersA)
                sqlDAA.SelectCommand.Connection.Open()
                sqlDAA.Fill(oDataSet, storedProcNameA)
                sqlDAA.SelectCommand.Connection.Close()

                sqlDAB.SelectCommand = BuildQueryCommand(storedProcNameB, parametersB)
                sqlDAB.SelectCommand.Connection.Open()
                sqlDAB.Fill(oDataSet, storedProcNameB)
                sqlDAB.SelectCommand.Connection.Close()
            Finally
                sqlDAA.SelectCommand.Dispose()
                sqlDAB.SelectCommand.Dispose()
                sqlDAA.Dispose()
                sqlDAB.Dispose()
            End Try

            Return oDataSet
        End Function

        Public Sub OpenConnection()
            If Connection.ConnectionString = "" Then
                classDB()
            End If

            If Not Connection.State = ConnectionState.Open Then
                Connection.Open()
            End If
        End Sub

        Public Sub CloseConnection()
            If Connection.State = ConnectionState.Open Then
                Connection.Close()
                Connection.Dispose()
                classDB()
            End If
        End Sub
#End Region

#Region "   htParam method "
        ' <summary>
        ' 依據htParam生成SQL參數
        ' </summary>
        ' <param name="i">生成的SQL參數的起始Index</param>
        ' <returns></returns>
        Public Function getSpParams(ByVal i As Integer) As SqlParameter()
            Dim spParams(Me.htParam.Count + i - 1) As SqlParameter
            Dim item As DictionaryEntry

            For Each item In Me.htParam
                spParams(i) = New SqlParameter(item.Key.ToString(), item.Value)
                If item.Value Is Nothing Then ' Or item.Value.ToString().Trim().Equals(String.Empty) Then
                    spParams(i).Value = DBNull.Value
                Else
                    '防 sql injection 攻擊
                    Dim strTemp As String = item.Value.ToString()
                    If strTemp.IndexOf("--") >= 0 Or strTemp.IndexOf("/*") >= 0 Then
                        spParams(i).Value = item.Value.ToString().Replace("--", "").Replace("/*", "")
                    End If
                End If
                i = i + 1
            Next item

            Return spParams
        End Function

        ' <summary>
        ' 依據傳入的hashtable生成SQL參數
        ' </summary>
        ' <param name="i">生成的SQL參數的起始Index</param>
        ' <param name="htParam">hashtable</param>
        ' <returns></returns>
        Public Function getSpParams(ByVal i As Integer, ByVal htParam As Hashtable) As SqlParameter()
            Dim spParams(Me.htParam.Count + i - 1) As SqlParameter
            Dim item As DictionaryEntry

            For Each item In Me.htParam
                spParams(i) = New SqlParameter(item.Key.ToString(), item.Value)
                If item.Value Is Nothing Then 'Or item.Value.ToString().Equals(String.Empty) Then
                    spParams(i).Value = DBNull.Value
                Else
                    '防 sql injection 攻擊
                    Dim strTemp As String = item.Value.ToString()
                    If strTemp.IndexOf("--") >= 0 Or strTemp.IndexOf("/*") >= 0 Then
                        spParams(i).Value = item.Value.ToString().Replace("--", "").Replace("/*", "")
                    End If
                End If
                i = i + 1
            Next item

            Return spParams
        End Function
#End Region

    End Class
End Namespace