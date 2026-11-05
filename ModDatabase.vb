Option Strict On
Option Explicit On

Imports System.Data
Imports System.IO
Imports Microsoft.Data.Sqlite

Namespace ProyekMahasiswa

    Public Module ModDatabase

        Private Const NamaDb As String = "dbKampus.db"

        Public Function PathDB() As String
            Return Path.Combine(AppContext.BaseDirectory, NamaDb)
        End Function

        Private Function Koneksi() As SqliteConnection
            Return New SqliteConnection("Data Source=" & PathDB())
        End Function

        Public Sub BuatDatabaseJikaBelum()
            Using con As SqliteConnection = Koneksi()
                con.Open()

                Using cmd As SqliteCommand = con.CreateCommand()
                    cmd.CommandText =
                        "CREATE TABLE IF NOT EXISTS mahasiswa (" &
                        "nim TEXT PRIMARY KEY, " &
                        "nama TEXT NOT NULL, " &
                        "jurusan TEXT, " &
                        "alamat TEXT, " &
                        "telepon TEXT)"
                    cmd.ExecuteNonQuery()
                End Using

                Dim jumlah As Long
                Using cmd As SqliteCommand = con.CreateCommand()
                    cmd.CommandText = "SELECT COUNT(*) FROM mahasiswa"
                    jumlah = Convert.ToInt64(cmd.ExecuteScalar())
                End Using

                If jumlah = 0 Then
                    TambahContoh(con, "2025001", "Andi Pratama", "Teknik Informatika", "Jl. Merdeka No. 10, Medan", "081234567890")
                    TambahContoh(con, "2025002", "Siti Aminah", "Sistem Informasi", "Jl. Sudirman No. 25, Medan", "081298765432")
                    TambahContoh(con, "2025003", "Budi Santoso", "Teknik Informatika", "Jl. Diponegoro No. 5, Medan", "085611223344")
                    TambahContoh(con, "2025004", "Dewi Lestari", "Manajemen Informatika", "Jl. Ahmad Yani No. 8, Binjai", "087755443322")
                    TambahContoh(con, "2025005", "Rudi Hartono", "Sistem Informasi", "Jl. Gatot Subroto No. 12, Medan", "082133445566")
                End If
            End Using
        End Sub

        Private Sub TambahContoh(con As SqliteConnection, nim As String, nama As String,
                                 jurusan As String, alamat As String, telepon As String)
            Using cmd As SqliteCommand = con.CreateCommand()
                cmd.CommandText = "INSERT INTO mahasiswa (nim, nama, jurusan, alamat, telepon) " &
                                  "VALUES (@nim, @nama, @jurusan, @alamat, @telepon)"
                IsiParameter(cmd, nim, nama, jurusan, alamat, telepon)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub IsiParameter(cmd As SqliteCommand, nim As String, nama As String,
                                 jurusan As String, alamat As String, telepon As String)
            cmd.Parameters.AddWithValue("@nim", nim)
            cmd.Parameters.AddWithValue("@nama", nama)
            cmd.Parameters.AddWithValue("@jurusan", jurusan)
            cmd.Parameters.AddWithValue("@alamat", alamat)
            cmd.Parameters.AddWithValue("@telepon", telepon)
        End Sub


        Public Function Ambil(kataKunci As String) As DataTable
            Dim dt As New DataTable()
            Using con As SqliteConnection = Koneksi()
                con.Open()
                Using cmd As SqliteCommand = con.CreateCommand()
                    Dim sql As String = "SELECT nim, nama, jurusan, alamat, telepon FROM mahasiswa"
                    If kataKunci.Length > 0 Then
                        sql &= " WHERE nim LIKE @k OR nama LIKE @k"
                        cmd.Parameters.AddWithValue("@k", "%" & kataKunci & "%")
                    End If
                    cmd.CommandText = sql & " ORDER BY nim"
                    Using reader As SqliteDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using
            End Using
            Return dt
        End Function

        Public Sub Tambah(nim As String, nama As String, jurusan As String, alamat As String, telepon As String)
            Using con As SqliteConnection = Koneksi()
                con.Open()
                TambahContoh(con, nim, nama, jurusan, alamat, telepon)
            End Using
        End Sub

        Public Sub Ubah(nim As String, nama As String, jurusan As String, alamat As String, telepon As String)
            Using con As SqliteConnection = Koneksi()
                con.Open()
                Using cmd As SqliteCommand = con.CreateCommand()
                    cmd.CommandText = "UPDATE mahasiswa SET nama=@nama, jurusan=@jurusan, " &
                                      "alamat=@alamat, telepon=@telepon WHERE nim=@nim"
                    IsiParameter(cmd, nim, nama, jurusan, alamat, telepon)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub Hapus(nim As String)
            Using con As SqliteConnection = Koneksi()
                con.Open()
                Using cmd As SqliteCommand = con.CreateCommand()
                    cmd.CommandText = "DELETE FROM mahasiswa WHERE nim=@nim"
                    cmd.Parameters.AddWithValue("@nim", nim)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

    End Module

End Namespace
