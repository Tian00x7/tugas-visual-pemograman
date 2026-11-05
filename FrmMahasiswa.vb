Option Strict On
Option Explicit On

Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Data.Sqlite

Namespace ProyekMahasiswa


    Public Class FrmMahasiswa
        Inherits Form

        Private txtNIM, txtNama, txtJurusan, txtAlamat, txtTelepon, txtCari As TextBox
        Private btnCari, btnSemua, btnTambah, btnSimpan, btnEdit, btnHapus, btnBatal, btnKeluar As Button
        Private btnAwal, btnSebelum, btnBerikut, btnAkhir As Button
        Private lblRecord As Label
        Private grid As DataGridView

        Private dt As DataTable
        Private pos As Integer = 0
        Private modeAksi As String = ""      ' "" = normal, "TAMBAH", "EDIT"
        Private sinkron As Boolean = False   ' mencegah event saling memanggil

        Public Sub New()
            BuatUI()
        End Sub

        ' Membuat tampilan

        Private Function Buat(Of T As {Control, New})(teks As String, x As Integer, y As Integer,
                                                      w As Integer, h As Integer) As T
            Dim c As New T()
            c.Text = teks
            c.SetBounds(x, y, w, h)
            Controls.Add(c)
            Return c
        End Function

        Private Sub BuatUI()
            Text = "Data Mahasiswa - ADO.NET dan DataTable (VB.NET)"
            FormBorderStyle = FormBorderStyle.FixedSingle
            MaximizeBox = False
            StartPosition = FormStartPosition.CenterScreen
            ClientSize = New Size(820, 620)
            Font = New Font("Segoe UI", 9.0F)

            Dim judul As Label = Buat(Of Label)("DATA MAHASISWA  (Database, DataTable dan Navigasi Record)", 20, 10, 780, 32)
            judul.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)

            Buat(Of Label)("NIM", 20, 63, 90, 23)
            txtNIM = Buat(Of TextBox)("", 120, 60, 150, 27)
            txtNIM.MaxLength = 10

            Buat(Of Label)("Nama", 20, 98, 90, 23)
            txtNama = Buat(Of TextBox)("", 120, 95, 320, 27)
            txtNama.MaxLength = 50

            Buat(Of Label)("Jurusan", 20, 133, 90, 23)
            txtJurusan = Buat(Of TextBox)("", 120, 130, 320, 27)
            txtJurusan.MaxLength = 30

            Buat(Of Label)("Alamat", 20, 168, 90, 23)
            txtAlamat = Buat(Of TextBox)("", 120, 165, 500, 27)
            txtAlamat.MaxLength = 100

            Buat(Of Label)("Telepon", 20, 203, 90, 23)
            txtTelepon = Buat(Of TextBox)("", 120, 200, 200, 27)
            txtTelepon.MaxLength = 15

            Buat(Of Label)("Cari NIM/Nama :", 20, 253, 110, 23)
            txtCari = Buat(Of TextBox)("", 135, 250, 250, 27)
            btnCari = Buat(Of Button)("Cari", 395, 248, 80, 30)
            btnSemua = Buat(Of Button)("Tampilkan Semua", 485, 248, 130, 30)

            btnTambah = Buat(Of Button)("&Tambah", 20, 295, 90, 32)
            btnSimpan = Buat(Of Button)("&Simpan", 120, 295, 90, 32)
            btnEdit = Buat(Of Button)("&Edit", 220, 295, 90, 32)
            btnHapus = Buat(Of Button)("&Hapus", 320, 295, 90, 32)
            btnBatal = Buat(Of Button)("&Batal", 420, 295, 90, 32)
            btnKeluar = Buat(Of Button)("&Keluar", 520, 295, 90, 32)

            btnAwal = Buat(Of Button)("|<", 20, 340, 50, 30)
            btnSebelum = Buat(Of Button)("<", 75, 340, 50, 30)
            btnBerikut = Buat(Of Button)(">", 130, 340, 50, 30)
            btnAkhir = Buat(Of Button)(">|", 185, 340, 50, 30)
            lblRecord = Buat(Of Label)("Record 0 dari 0", 255, 346, 250, 23)

            grid = New DataGridView()
            grid.SetBounds(20, 385, 780, 215)
            grid.ReadOnly = True
            grid.AllowUserToAddRows = False
            grid.AllowUserToDeleteRows = False
            grid.MultiSelect = False
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            grid.RowHeadersVisible = False
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            Controls.Add(grid)

            AddHandler btnCari.Click, AddressOf BtnCari_Click
            AddHandler btnSemua.Click, AddressOf BtnSemua_Click
            AddHandler btnTambah.Click, AddressOf BtnTambah_Click
            AddHandler btnSimpan.Click, AddressOf BtnSimpan_Click
            AddHandler btnEdit.Click, AddressOf BtnEdit_Click
            AddHandler btnHapus.Click, AddressOf BtnHapus_Click
            AddHandler btnBatal.Click, AddressOf BtnBatal_Click
            AddHandler btnKeluar.Click, AddressOf BtnKeluar_Click
            AddHandler btnAwal.Click, AddressOf BtnAwal_Click
            AddHandler btnSebelum.Click, AddressOf BtnSebelum_Click
            AddHandler btnBerikut.Click, AddressOf BtnBerikut_Click
            AddHandler btnAkhir.Click, AddressOf BtnAkhir_Click
            AddHandler txtCari.KeyDown, AddressOf TxtCari_KeyDown
            AddHandler grid.SelectionChanged, AddressOf Grid_SelectionChanged
        End Sub


        ' Form dibuka
        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)
            Try
                BuatDatabaseJikaBelum()
                modeAksi = ""
                AturMode(False)
                MuatData("")
            Catch ex As Exception
                MessageBox.Show("Gagal membuka database:" & vbCrLf & ex.Message,
                                "Kesalahan", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        ' Memuat data + menampilkan record
        Private Sub MuatData(kataKunci As String, Optional pindahKeNim As String = "")
            dt = Ambil(kataKunci)
            pos = 0
            If pindahKeNim.Length > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    If dt.Rows(i)("nim").ToString() = pindahKeNim Then
                        pos = i
                        Exit For
                    End If
                Next
            End If

            sinkron = True
            grid.DataSource = dt
            If grid.Columns.Count = 5 Then
                grid.Columns(0).HeaderText = "NIM"
                grid.Columns(1).HeaderText = "Nama"
                grid.Columns(2).HeaderText = "Jurusan"
                grid.Columns(3).HeaderText = "Alamat"
                grid.Columns(4).HeaderText = "Telepon"
            End If
            sinkron = False

            Tampil()
        End Sub

        Private Sub Tampil()
            sinkron = True
            Try
                If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                    KosongkanKotak()
                    lblRecord.Text = "Tidak ada data"
                    Return
                End If

                If pos < 0 Then pos = 0
                If pos > dt.Rows.Count - 1 Then pos = dt.Rows.Count - 1

                Dim r As DataRow = dt.Rows(pos)
                txtNIM.Text = r("nim").ToString()
                txtNama.Text = r("nama").ToString()
                txtJurusan.Text = r("jurusan").ToString()
                txtAlamat.Text = r("alamat").ToString()
                txtTelepon.Text = r("telepon").ToString()
                lblRecord.Text = "Record " & (pos + 1) & " dari " & dt.Rows.Count

                grid.ClearSelection()
                grid.Rows(pos).Selected = True
                grid.CurrentCell = grid.Rows(pos).Cells(0)
            Finally
                sinkron = False
            End Try
        End Sub

        Private Sub KosongkanKotak()
            txtNIM.Clear()
            txtNama.Clear()
            txtJurusan.Clear()
            txtAlamat.Clear()
            txtTelepon.Clear()
        End Sub

        Private Sub Grid_SelectionChanged(sender As Object, e As EventArgs)
            If sinkron OrElse modeAksi <> "" OrElse dt Is Nothing Then Return
            If grid.CurrentRow Is Nothing Then Return
            pos = grid.CurrentRow.Index
            Tampil()
        End Sub

        ' Mengatur tombol/TextBox sesuai mode

        Private Sub AturMode(editing As Boolean)
            txtNIM.ReadOnly = Not (editing AndAlso modeAksi = "TAMBAH")  ' NIM = kunci, tidak diubah saat Edit
            txtNama.ReadOnly = Not editing
            txtJurusan.ReadOnly = Not editing
            txtAlamat.ReadOnly = Not editing
            txtTelepon.ReadOnly = Not editing

            btnSimpan.Enabled = editing
            btnBatal.Enabled = editing

            btnTambah.Enabled = Not editing
            btnEdit.Enabled = Not editing
            btnHapus.Enabled = Not editing
            btnCari.Enabled = Not editing
            btnSemua.Enabled = Not editing
            txtCari.Enabled = Not editing
            btnAwal.Enabled = Not editing
            btnSebelum.Enabled = Not editing
            btnBerikut.Enabled = Not editing
            btnAkhir.Enabled = Not editing
            grid.Enabled = Not editing
        End Sub

        ' Tombol CRUD

        Private Sub BtnTambah_Click(sender As Object, e As EventArgs)
            modeAksi = "TAMBAH"
            AturMode(True)
            KosongkanKotak()
            lblRecord.Text = "Record baru"
            txtNIM.Focus()
        End Sub

        Private Sub BtnEdit_Click(sender As Object, e As EventArgs)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                MessageBox.Show("Belum ada data yang bisa diedit.", "Info",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            modeAksi = "EDIT"
            AturMode(True)
            txtNama.Focus()
        End Sub

        Private Sub BtnSimpan_Click(sender As Object, e As EventArgs)
            If txtNIM.Text.Trim().Length = 0 OrElse txtNama.Text.Trim().Length = 0 Then
                MessageBox.Show("NIM dan Nama wajib diisi!", "Validasi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim nim As String = txtNIM.Text.Trim()
            Try
                If modeAksi = "TAMBAH" Then
                    Tambah(nim, txtNama.Text.Trim(), txtJurusan.Text.Trim(),
                           txtAlamat.Text.Trim(), txtTelepon.Text.Trim())
                Else
                    Ubah(nim, txtNama.Text.Trim(), txtJurusan.Text.Trim(),
                         txtAlamat.Text.Trim(), txtTelepon.Text.Trim())
                End If
            Catch ex As SqliteException When ex.SqliteErrorCode = 19   ' kunci ganda
                MessageBox.Show("NIM " & nim & " sudah terdaftar!", "Validasi",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtNIM.Focus()
                Return
            Catch ex As Exception
                MessageBox.Show("Gagal menyimpan data:" & vbCrLf & ex.Message, "Kesalahan",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            MessageBox.Show("Data berhasil disimpan.", "Sukses",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            modeAksi = ""
            AturMode(False)
            txtCari.Clear()
            MuatData("", nim)
        End Sub

        Private Sub BtnHapus_Click(sender As Object, e As EventArgs)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                MessageBox.Show("Tidak ada data yang bisa dihapus.", "Info",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim nim As String = dt.Rows(pos)("nim").ToString()
            Dim nama As String = dt.Rows(pos)("nama").ToString()
            If MessageBox.Show("Yakin menghapus data " & nama & " ?", "Konfirmasi Hapus",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    ModDatabase.Hapus(nim)
                    MuatData(txtCari.Text.Trim())
                    MessageBox.Show("Data berhasil dihapus.", "Sukses",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Gagal menghapus data:" & vbCrLf & ex.Message, "Kesalahan",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Sub

        Private Sub BtnBatal_Click(sender As Object, e As EventArgs)
            modeAksi = ""
            AturMode(False)
            Tampil()
        End Sub

        Private Sub BtnKeluar_Click(sender As Object, e As EventArgs)
            Close()
        End Sub


        ' Pencarian

        Private Sub BtnCari_Click(sender As Object, e As EventArgs)
            MuatData(txtCari.Text.Trim())
            If dt.Rows.Count = 0 Then
                MessageBox.Show("Data tidak ditemukan.", "Pencarian",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub BtnSemua_Click(sender As Object, e As EventArgs)
            txtCari.Clear()
            MuatData("")
        End Sub

        Private Sub TxtCari_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                BtnCari_Click(sender, EventArgs.Empty)
            End If
        End Sub

        ' Navigasi record

        Private Sub PindahPos(posBaru As Integer)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
            pos = Math.Max(0, Math.Min(posBaru, dt.Rows.Count - 1))
            Tampil()
        End Sub

        Private Sub BtnAwal_Click(sender As Object, e As EventArgs)
            PindahPos(0)
        End Sub

        Private Sub BtnSebelum_Click(sender As Object, e As EventArgs)
            PindahPos(pos - 1)
        End Sub

        Private Sub BtnBerikut_Click(sender As Object, e As EventArgs)
            PindahPos(pos + 1)
        End Sub

        Private Sub BtnAkhir_Click(sender As Object, e As EventArgs)
            If dt IsNot Nothing Then PindahPos(dt.Rows.Count - 1)
        End Sub

    End Class

End Namespace
