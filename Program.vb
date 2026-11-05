Option Strict On
Option Explicit On

Imports System.Windows.Forms

Namespace ProyekMahasiswa

    Public Module Program

        <STAThread()>
        Public Sub Main()
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Application.Run(New FrmMahasiswa())
        End Sub

    End Module

End Namespace
