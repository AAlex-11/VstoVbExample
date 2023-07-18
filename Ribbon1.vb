Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Dim WithEvents Bg1 As BackgroundWorker
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Debug.WriteLine("Ribbon1_Load")
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Debug.WriteLine("BackgroundWorker started")
        Bg1 = New BackgroundWorker
        Bg1.RunWorkerAsync()
    End Sub

    Private Sub Bg1_DoWork(sender As Object, e As DoWorkEventArgs) Handles Bg1.DoWork
        Application.DoEvents()
        Module1.Start()
    End Sub

    Private Sub Bg1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles Bg1.RunWorkerCompleted
        Module1.Finish()
    End Sub
End Class
