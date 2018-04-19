Imports System.Data.OleDb
Imports System.Data.SqlClient

Namespace com.vasilchenko.DBAccessConnection
    Module DBConnection
        Public Function GetOleDBDataTable(strSQLQuery As String, strSourcePath As String) As DataTable
            Dim strSQLConnectionParameters As String
            Dim objOleDbCommand As OleDbCommand
            Dim objDataTable As New DataTable

            strSQLConnectionParameters = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strSourcePath & ";"
            Using objOleDbConnection As New OleDbConnection(strSQLConnectionParameters)
                Try
                    objOleDbConnection.Open()
                    objOleDbCommand = New OleDbCommand(strSQLQuery, objOleDbConnection)
                    objDataTable.Load(objOleDbCommand.ExecuteReader)

                    Return objDataTable
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Throw New ArgumentNullException
                End Try
            End Using
        End Function

        Public Function GetSQLDBDataTable(strSQLQuery As String, strSQLConnectionParameters As String) As DataTable
            Dim objDataTable As New DataTable

            Using objSQLDbConnection As New SqlConnection(strSQLConnectionParameters)
                Try
                    objSQLDbConnection.Open()
                    Dim objSQLDbCommand = New SqlCommand(strSQLQuery, objSQLDbConnection)
                    objDataTable.Load(objSQLDbCommand.ExecuteReader)

                    Return objDataTable
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Throw New ArgumentNullException
                End Try
            End Using
        End Function

    End Module
End Namespace
