Imports System
Imports System.IO
Imports System.IO.File
Imports System.Net.WebRequestMethods
Imports Microsoft.Office.Interop.Word

Public Class templateFiller

    Dim templateWord As Microsoft.Office.Interop.Word.Application
    Dim documentoWord As Microsoft.Office.Interop.Word.Document

    Private Sub templateFiller_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim templateWord As Microsoft.Office.Interop.Word.Application
            Dim documentoWord As Microsoft.Office.Interop.Word.Document
        Catch ex As Exception
            MsgBox("No se ha podido iniciar Microsoft Word", vbCritical)
        End Try
    End Sub

    Private Sub btnCargar_Click(sender As Object, e As EventArgs) Handles btnCargar.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            templateWord = New Microsoft.Office.Interop.Word.Application()
            documentoWord = templateWord.Documents.Add(OpenFileDialog1.FileName)

            If documentoWord.Bookmarks.Count = 0 Then
                MsgBox("La plantilla seleccionada no contiene marcadores para ser sustituidos por valores." + Environment.NewLine + "Por favor, seleccione una plantilla con marcadores.", vbExclamation)
            Else
                For i = 1 To documentoWord.Bookmarks.Count
                    Me.DataGridView1.Rows.Add(documentoWord.Bookmarks(i).Name)
                Next

                txtTemplate.Text = OpenFileDialog1.FileName

                documentoWord.Close(WdSaveOptions.wdDoNotSaveChanges)
                templateWord.Quit()
            End If

        End If
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Try

            templateWord = New Microsoft.Office.Interop.Word.Application()
            documentoWord = templateWord.Documents.Add(OpenFileDialog1.FileName)

            For fila = 0 To DataGridView1.Rows.Count - 1

                For marcador = 1 To documentoWord.Bookmarks.Count
                    If documentoWord.Bookmarks(marcador).Name = DataGridView1.Item(0, fila).Value.ToString Then
                        documentoWord.Bookmarks(marcador).Range.Text = DataGridView1.Item(1, fila).Value.ToString
                        Exit For
                    End If
                Next

            Next

            If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                documentoWord.SaveAs(SaveFileDialog1.FileName, 16)
                MsgBox("El documento se ha guardado correctamente", vbInformation)
                For fila = 0 To DataGridView1.Rows.Count - 1
                    DataGridView1.Item(1, fila).Value = ""
                Next
            Else
                MsgBox("No se ha podido guardar el documento.", vbCritical)
            End If
            

        Catch ex2 As Exception

        Finally
            documentoWord.Close(WdSaveOptions.wdDoNotSaveChanges)
            templateWord.Quit()
        End Try

    End Sub

    Private Sub templateFiller_Closing(sender As Object, e As EventArgs) Handles MyBase.Closing
        Try
            documentoWord.Close(WdSaveOptions.wdDoNotSaveChanges)
            templateWord.Quit()
        Catch ex As Exception

        End Try
    End Sub
End Class
