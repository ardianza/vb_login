Public Class Form1

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        'Membuat koneksi database
        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\data\users.accdb")

        'Membuka koneksi ke database
        conn.Open()

        'Command untuk memeriksa data login
        Dim cmd As New OleDb.OleDbCommand("SELECT * FROM users WHERE username = @username AND password = @password", conn)

        'Mengisi Parameter pada command
        cmd.Parameters.AddWithValue("@username", txtUsername.Text)
        cmd.Parameters.AddWithValue("@password", txtPassword.Text)

        'Menjalankan command dan menyimpan hasil nya dalam data reader
        Dim dr As OleDb.OleDbDataReader = cmd.ExecuteReader()

        'Cek data di database
        If dr.Read() Then
            'Jika username dan password ada di database
            MessageBox.Show("Login Berhasil")
        Else
            'Jika username dan password salah atau tidak ada si dtabase
            lblError.Text = "Username atau Password salah"
        End If

        'Menutup koneksi ke Database
        dr.Close()
        cmd.Dispose()
        conn.Close()

    End Sub
End Class
