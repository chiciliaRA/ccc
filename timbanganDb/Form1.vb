Imports MySql.Data.MySqlClient
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call koneksi()
        Call tampil_harian()
        Call cb_Supplier()
        Call tampil_pegawai()
        Call tampil_Op()

    End Sub
    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        FlowLayoutPanel1.Visible = False
        FlowLayoutPanel2.Visible = True
    End Sub

    Private Sub lbchangeOp_Click(sender As Object, e As EventArgs) Handles lbchangeOp.Click
        FlowLayoutPanel2.Visible = False
        FlowLayoutPanel1.Visible = True
    End Sub

    Private Sub btLoginOp_Click(sender As Object, e As EventArgs) Handles btLoginOp.Click
        Call koneksi()
        cmd = New MySqlCommand("select password from user where kode_op= '" + TextBox1.Text + "' ", conn)
        rd = cmd.ExecuteReader
        Do While rd.Read = True
            If rd(0) = TextBox2.Text Then
                PanelDashOp.BringToFront()
                PictureBox1.Enabled = False
            Else
                MessageBox.Show("ID atau Password salah")
            End If
        Loop

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        sidebar.Visible = False
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        sidebar.Visible = True
        sidebar.BringToFront()
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            Call koneksi()
            Call tampil_harian()
        End If
        If ComboBox1.SelectedIndex = 1 Then
            Call koneksi()
            Call tampil_operator()
        End If
        If ComboBox1.SelectedIndex = 2 Then
            Call koneksi()
            Call tampil_pg()
        End If

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox1.SelectedIndex = 0 Then
            Call koneksi()
            Call cari_tgl()
        End If
        If ComboBox1.SelectedIndex = 1 Then
            Call koneksi()
            Call cari_op()
        End If
        If ComboBox1.SelectedIndex = 2 Then
            Call koneksi()
            Call cari_pg()
        End If
    End Sub

    Sub tampil_harian()
        Call koneksi()
        cmd = New MySqlCommand("select * from operasi_harian", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6), rd(7), rd(8))

        Loop
    End Sub
    Sub tampil_operator()
        Call koneksi()
        cmd = New MySqlCommand("select * from operasi_harian order by kode_op", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))

        Loop
    End Sub
    Sub cb_Supplier()

        Call koneksi()
        cmd = New MySqlCommand("select namaSup from supplier", conn)
        rd = cmd.ExecuteReader
        ComboBox2.Items.Clear()
        Do While rd.Read = True
            ComboBox2.Items.Add(rd(0))
        Loop
    End Sub
    Sub tampil_pg()
        Call koneksi()
        cmd = New MySqlCommand("select * from operasi_harian order by kode_pg", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))

        Loop
    End Sub
    Sub cari_pg()
        Call koneksi()
        cmd = New MySqlCommand("select * from operasi_harian where kode_pg = '" + tbCari.Text + "' ", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))

        Loop
    End Sub
    Sub cari_op()
        Call koneksi()
        cmd = New MySqlCommand("select * from operasi_harian where kode_op = '" + tbCari.Text + "' ", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))

        Loop
    End Sub
    Sub cari_tgl()
        Call koneksi()
        cmd = New MySqlCommand("select * from operasi_harian where tanggal = '" + tbCari.Text + "' ", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))

        Loop
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SaveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim xlaapp As Microsoft.Office.Interop.Excel.Application
            Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
            Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim misvalue As Object = System.Reflection.Missing.Value
            Dim i As Integer
            Dim j As Integer

            xlaapp = New Microsoft.Office.Interop.Excel.Application
            xlworkbook = xlaapp.Workbooks.Add(misvalue)
            xlworksheet = xlworkbook.Sheets("sheet1")

            For i = 0 To DataGridView1.RowCount - 2
                For j = 0 To DataGridView1.ColumnCount - 1
                    For k As Integer = 1 To DataGridView1.Columns.Count
                        xlworksheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                        xlworksheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()


                    Next
                Next
            Next
            xlworksheet.SaveAs(SaveFileDialog1.FileName)
            xlworkbook.Close()
            xlaapp.Quit()

            releaseobject(xlaapp)
            releaseobject(xlworkbook)
            releaseobject(xlworksheet)

            MessageBox.Show("Proses Export Berhasil", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

    End Sub

    Private Sub releaseobject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing

        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call koneksi()
        Call inputhasilKerja()
    End Sub
    Sub inputhasilKerja()
        TextBox7.Text = TextBox5.Text
        If PKB.Checked Then
            TextBox8.Text = PKB.Text
            PKK.Checked = False
            PKKK.Checked = False
            NON.Checked = False
        End If
        If PKK.Checked Then
            TextBox8.Text = PKK.Text
            PKB.Checked = False
            PKKK.Checked = False
            NON.Checked = False
        End If
        If PKKK.Checked Then
            TextBox8.Text = PKKK.Text
            PKB.Checked = False
            PKK.Checked = False
            NON.Checked = False
        End If
        If NON.Checked Then
            TextBox8.Text = NON.Text
            PKB.Checked = False
            PKKK.Checked = False
            PKK.Checked = False
        End If
        Call koneksi()
        cmd = New MySqlCommand("select nama_pg from pegawai where kode_pg = '" + TextBox5.Text + "' ", conn)
        rd = cmd.ExecuteReader
        Do While rd.Read = True
            TextBox9.Text = (rd(0))
        Loop
        cmd = New MySqlCommand("select nama from produk where kode_pd = '" + TextBox8.Text + "' ", conn)
        rd = cmd.ExecuteReader
        Do While rd.Read = True
            TextBox10.Text = (rd(0))
        Loop
    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        PanelDashOp.BringToFront()

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        Panel1.BringToFront()

    End Sub
    Sub tampil_pegawai()
        Call koneksi()
        cmd = New MySqlCommand("select * from pegawai", conn)
        rd = cmd.ExecuteReader
        DataGridView2.Rows.Clear()
        Do While rd.Read = True
            DataGridView2.Rows.Add(rd(0), rd(1), rd(2))

        Loop
    End Sub
    Sub tampil_Op()
        Call koneksi()
        cmd = New MySqlCommand("select * from user", conn)
        rd = cmd.ExecuteReader
        DataGridView3.Rows.Clear()
        Do While rd.Read = True
            DataGridView3.Rows.Add(rd(0), rd(1), rd(2), rd(3))

        Loop
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        SaveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim xlaapp As Microsoft.Office.Interop.Excel.Application
            Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
            Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim misvalue As Object = System.Reflection.Missing.Value
            Dim i As Integer
            Dim j As Integer

            xlaapp = New Microsoft.Office.Interop.Excel.Application
            xlworkbook = xlaapp.Workbooks.Add(misvalue)
            xlworksheet = xlworkbook.Sheets("sheet1")

            For i = 0 To DataGridView2.RowCount - 2
                For j = 0 To DataGridView2.ColumnCount - 1
                    For k As Integer = 1 To DataGridView1.Columns.Count
                        xlworksheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                        xlworksheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()


                    Next
                Next
            Next
            xlworksheet.SaveAs(SaveFileDialog1.FileName)
            xlworkbook.Close()
            xlaapp.Quit()

            releaseobject(xlaapp)
            releaseobject(xlworkbook)
            releaseobject(xlworksheet)

            MessageBox.Show("Proses Export Berhasil", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call koneksi()
        cmd = New MySqlCommand("select * from pegawai where kode_pg = '" + TextBox13.Text + "' ", conn)
        rd = cmd.ExecuteReader
        DataGridView2.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2))

        Loop
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        PanelOp.BringToFront()

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        sidebar.Visible = True
        sidebar.BringToFront()

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        sidebar.Visible = True
        sidebar.BringToFront()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        sidebar.Visible = True
        sidebar.BringToFront()
    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click
        PanelPegawai.BringToFront()

    End Sub

    Private Sub PKB_CheckedChanged(sender As Object, e As EventArgs) Handles PKB.CheckedChanged
        If PKB.Checked Then
            PKK.Checked = False
            PKKK.Checked = False
            NON.Checked = False
        End If

    End Sub

    Private Sub PKKK_CheckedChanged(sender As Object, e As EventArgs) Handles PKKK.CheckedChanged

        If PKKK.Checked Then
            PKB.Checked = False
            PKK.Checked = False
            NON.Checked = False
        End If

    End Sub

    Private Sub PKK_CheckedChanged(sender As Object, e As EventArgs) Handles PKK.CheckedChanged

        If PKK.Checked Then
            PKB.Checked = False
            PKKK.Checked = False
            NON.Checked = False
        End If
    End Sub

    Private Sub NON_CheckedChanged(sender As Object, e As EventArgs) Handles NON.CheckedChanged

        If NON.Checked Then
            PKB.Checked = False
            PKKK.Checked = False
            PKK.Checked = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call koneksi()
        cmd = New MySqlCommand("select password from user where kode_op= '" + TextBox3.Text + "' ", conn)
        rd = cmd.ExecuteReader
        Do While rd.Read = True
            If rd(0) = TextBox4.Text Then
                PanelDashOp.BringToFront()
                PictureBox1.Enabled = True
            Else
                MessageBox.Show("ID atau Password salah")
            End If
        Loop
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Call koneksi()
            cmd = New MySqlCommand("update pegawai set '" + TextBox8.Text + "' = '" + TextBox1.Text + "'where kode_pg = '" + TextBox7.Text + "' ", conn)
            rd = cmd.ExecuteReader
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call koneksi()
        cmd = New MySqlCommand("insert into pegawai (kode_pg, nama_pg, telp) values ('" + TextBox6.Text + "','" + TextBox16.Text + "','" + TextBox17.Text + "')", conn)
        rd = cmd.ExecuteReader
        MessageBox.Show("Data berhasil disimpan")
        TextBox6.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        PanelDashOp.BringToFront()

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        PanelTambahPekerja.BringToFront()

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Call koneksi()
        cmd = New MySqlCommand("update user set nama='" + TextBox19.Text + "', password= '" + TextBox20.Text + "',status= '" + TextBox21.Text + "'where kode_op= '" + TextBox18.Text + "'", conn)
        rd = cmd.ExecuteReader
        MessageBox.Show("Data berhasil disimpan")
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        FlowLayoutPanel12.Visible = True
        FlowLayoutPanel12.BringToFront()
        Button11.Visible = True
        Button14.Visible = False

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Call koneksi()
        cmd = New MySqlCommand("insert into user (kode_op, nama, password,status) values ('" + TextBox18.Text + "','" + TextBox19.Text + "','" + TextBox20.Text + "','" + TextBox21.Text + "')", conn)
        rd = cmd.ExecuteReader
        MessageBox.Show("Data berhasil disimpan")
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        FlowLayoutPanel12.Visible = True
        FlowLayoutPanel12.BringToFront()
        Button11.Visible = False
        Button14.Visible = True
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        FlowLayoutPanel12.Visible = False
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        PanelTambahPekerja.BringToFront()

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        sidebar.Visible = True
        sidebar.BringToFront()
    End Sub
End Class
