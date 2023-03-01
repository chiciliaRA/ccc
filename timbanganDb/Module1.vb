
Imports MySql.Data.MySqlClient
Module Module1
    Public conn As MySqlConnection
    Public cmd As MySqlCommand
    Public rd As MySqlDataReader
    Dim conf As String = "Server=localhost; user id=root; password=; database=timbangan"
    Sub koneksi()
        Try
            conn = New MySqlConnection(conf)
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Module
