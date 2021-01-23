Imports System.IO

Public Class ArquivoOS

    Public Shared Sub Fatura(ByVal diretorio As String, ByVal nomeArquivo As String, ByVal nomearquivopdf As String, ByVal local As Integer, ByVal estadual As Integer, ByVal nacional As Integer, ByVal lote As Integer, ByVal codpost As String, ByVal codadmin As String)
        Try
            Dim arquivo As StreamWriter
            If Not Directory.Exists(diretorio) Then
                Directory.CreateDirectory(diretorio)
            End If
            If File.Exists(diretorio & nomeArquivo) Then
                arquivo = New StreamWriter(diretorio & nomeArquivo, True)
            Else
                arquivo = New StreamWriter(diretorio & nomeArquivo)
            End If
            arquivo.WriteLine("A;" & nomearquivopdf & ";" & local & ";" & estadual & ";" & nacional & ";1;N;" & lote & ";0;;" & (local + estadual + nacional) * 2 & ";" & (local + estadual + nacional) & ";" & codpost & ";" & codadmin)
            arquivo.Flush()
            arquivo.Close()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

End Class
