Imports Marpress.Interfaces
Imports System.IO
Imports System.Reflection
Imports Marpress.Padrao

Public Class ArquivoFatura
    Implements IArquivo

    Private arquivo As StreamReader

    Private _linhas As List(Of ICIF)
    Public Property Linhas() As List(Of ICIF) Implements IArquivo.Linhas
        Get
            Return _linhas
        End Get
        Set(ByVal value As List(Of ICIF))
            _linhas = value
        End Set
    End Property

    Public Sub New(ByVal nomeArquivoEntrada As String, ByVal label As Label)
        Try
            arquivo = New StreamReader(nomeArquivoEntrada, System.Text.Encoding.Default)
            Dim linhas As New List(Of ICIF)
            linhas = lerArquivo(label)
            _linhas = linhas
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function lerArquivo(ByVal label As Label) As List(Of ICIF)
        Dim linhasArquivo As New List(Of ICIF)
        Try
            Dim fatura As New Layout
            Dim linha As String
            Dim id As String = ""
            Dim qtdeLinhas As Integer = 0
            Dim qtdeCarnes As Integer = 0
            Dim texto As String = label.Text

            While arquivo.Peek > -1
                qtdeLinhas += 1
                linha = arquivo.ReadLine
                Dim campo As String() = linha.Split(";")
                If id <> campo(9) Then
                    fatura = New Layout
                    qtdeCarnes += 1
                    fatura.CarregaRegistros(linha)
                    linhasArquivo.Add(fatura)
                    id = campo(9)
                Else
                    fatura.CarregaParcelas(linha)
                End If
                'MsgBox("" & fatura.Codigo & " " & fatura.Nome)
                label.Text = texto & qtdeLinhas & vbCrLf & "Carnês... " & qtdeCarnes
                If linha Is Nothing Then
                    Exit While
                End If
            End While

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return linhasArquivo

    End Function

End Class
