'====================================================================================================
'
'   Accessファイルに接続し、データを読み取りコンソールに表示するプログラム
'       
'       ConnectAccessFile.exe 引数1 引数2
'
'           引数1：判定管理_DATA.accdbのPath
'           引数2：判定受付番号
'
'           出力：データテーブルのリスト（カンマ区切りテキスト）

'                   JUTAKU_NO, JUTAKU_NO_EDA, MemberID, HA_Data_KUBUN, HA_Number, HA_Select_Date, HA_JISSI_Date
'                   220841,1,155,1,2,20221104,20221108
'                   220841,1,155,1,2,20221104,20221109
'                   220841,1,157,1,1,20221104,20221110
'                   220841,2,157,1,2,20221104,20221108
'                   220841,2,157,1,2,20221104,20221109
'                   220841,2,155,1,1,20221104,20221110
'
'====================================================================================================


Imports System
Imports System.Data
Imports System.Data.OleDb

Module Program
    Sub Main(args As String())
        'Console.WriteLine("Hello World!")

        Dim API1 As New AccessAPI
        'API1.FilePath = "W:\判定管理_DATA.accdb"
        Dim Path As String = args(0)
        Dim gbrcNo As String = args(1)

        API1.FilePath = Path

        Dim sql As String = "SELECT * FROM T_HANTEI_RIREKI WHERE JUTAKU_NO = '" + gbrcNo + "'"

        Dim resultDt As DataTable = API1.GetTableData(sql)

        With resultDt

            Dim rn As Integer = .Rows.Count
            If rn > 0 Then
                Dim itemName As String() = {"JUTAKU_NO", "JUTAKU_NO_EDA", "MemberID", "HA_Data_KUBUN", "HA_Number", "HA_Select_Date", "HA_JISSI_Date"}
                Dim n0 As Integer = itemName.Length
                Dim Ans As String = ""
                For j As Integer = 0 To n0 - 1
                    Ans += itemName(j)
                    If j < n0 - 1 Then
                        Ans += ","
                    Else
                        Ans += vbCrLf
                    End If
                Next

                For i As Integer = 0 To rn - 1
                    'Ans += .Rows(i).Item("JUTAKU_NO").ToString + ","
                    'Ans += .Rows(i).Item("JUTAKU_NO_EDA").ToString + ","
                    'Ans += .Rows(i).Item("MemberID").ToString + ","
                    'Ans += .Rows(i).Item("HA_Data_KUBUN").ToString + ","
                    'Ans += .Rows(i).Item("HA_Number").ToString + ","
                    'Ans += .Rows(i).Item("HA_Select_Date").ToString + ","
                    'Ans += .Rows(i).Item("HA_JISSI_Date").ToString + vbCrLf
                    For j As Integer = 0 To n0 - 1
                        Ans += .Rows(i).Item(itemName(j)).ToString
                        If j < n0 - 1 Then
                            Ans += ","
                        Else
                            Ans += vbCrLf
                        End If
                    Next
                Next
                Console.WriteLine(Ans)
            Else
                Console.WriteLine("Nothing")
            End If
        End With

    End Sub
End Module
