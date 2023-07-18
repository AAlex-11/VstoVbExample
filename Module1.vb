Imports System.Diagnostics
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Module Module1

    Public Sub Start()
        Try
            Debug.WriteLine("Start")
            Dim Num As Integer = CInt(Globals.Ribbons.Ribbon1.EditBox1.Text)
            Dim Wb As Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
            Dim Ws As Worksheet = Wb.ActiveSheet
            Dim Range1 As Range = Ws.Range($"A1:Z{Num}")
            Dim i As Integer = 0
            Dim RND As New Random(255)
            For Each One As Range In Range1.Cells
                One.Interior.Color = RGB(RND.Next Mod 255, RND.Next Mod 255, RND.Next Mod 255)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Finish()
        Debug.WriteLine("Finish")
    End Sub

End Module
