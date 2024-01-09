
Imports System.Data
Imports System.Data.OleDb
'Imports System.Data
'Imports System.Reflection.Metadata.Ecma335

Public Class AccessAPI

    Private _ConnectionString As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source="
    Private _FilePath As String

    Public Property FilePath() As String
        Get
            Return _FilePath
        End Get
        Set(ByVal value As String)
            _FilePath = value
        End Set
    End Property

    Public Function GetTableData(ByVal sql As String) As DataTable
        'SQL作成

        If sql <> "" And _FilePath <> "" Then
            Dim resultDt As New DataTable
            'Dim sql = New System.Text.StringBuilder()
            'sql.AppendLine("SELECT")
            'sql.AppendLine("  *")
            'sql.AppendLine("FROM T_HANTEI_RIREKI")
            'sql.AppendLine(" WHERE JUTAKU_NO = '230898'")

            'Access接続準備
            Dim command As New OleDbCommand
            Dim da As New OleDbDataAdapter
            Dim cnAccess As OleDbConnection = New OleDbConnection
            'cnAccess.ConnectionString = My.Settings.AccessCon
            'cnAccess.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=W:\判定管理_DATA.accdb"
            cnAccess.ConnectionString = _ConnectionString + _FilePath

            'Access接続開始
            cnAccess.Open()

            Try

                command.Connection = cnAccess
                command.CommandText = sql.ToString
                da.SelectCommand = command

                'SQL実行 結果をデータテーブルに格納
                da.Fill(resultDt)


            Catch ex As Exception
                Throw
            Finally
                command.Dispose()
                da.Dispose()
                cnAccess.Close()
                GetTableData = resultDt
            End Try
        Else
            GetTableData = Nothing
        End If
    End Function

End Class
