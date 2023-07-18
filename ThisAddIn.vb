Imports System.Diagnostics

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Debug.WriteLine("ThisAddIn_Startup")
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Debug.WriteLine("ThisAddIn_Shutdown")
    End Sub

End Class
