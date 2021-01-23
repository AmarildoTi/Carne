Imports System.IO
Imports Marpress.Funcoes
Imports Marpress.FAC
Imports Marpress.Interfaces

Public Class Processamento

    Shared Sub processa(ByVal diretorio As String, ByVal arquivoEntrada As String, ByVal tipofac As Tipo, ByVal contratofac As Contrato, ByVal label As Label, ByVal dataFAC As Date, ByVal producao As Boolean, ByVal ordemdeserviço As String)
        Try
            Dim dia As String = String.Format("{0:ddMMyyyy}", Now.Date)
            Dim nomearquivo As String = Left(arquivoEntrada, arquivoEntrada.Length - 4)

            label.Text = "Lendo o arquivo..."
            Dim arquivo As New ArquivoFatura(diretorio & arquivoEntrada, label)
            Dim cont As Integer = 0
            Dim sequencia As Integer = 0
            Dim total As Integer = 0
            Dim lote As Integer = 0
            Dim imagemcapa As String = ""
            Dim imageminstrucoes As String = ""
            Dim imagemparcela As String = ""
      
            processaFAC(contratofac, "Sehal", arquivo, producao, dataFAC, label)

            File.Delete(diretorio & nomearquivo & ".sem")
            File.Delete(diretorio & nomearquivo & "_99999" & ".dev")
            File.Delete(diretorio & nomearquivo & "_" & arquivo.Linhas(0).CIF.TipoRegistro & ".err")
            File.Delete(diretorio & nomearquivo & "_" & arquivo.Linhas(0).CIF.TipoRegistro & ".rel")
            File.Delete(diretorio & nomearquivo & "_" & arquivo.Linhas(0).CIF.TipoRegistro & ".os")

            Dim local As Integer = 0
            Dim estadual As Integer = 0
            Dim nacional As Integer = 0

            'Dim selecao As IEnumerable(Of ICIF) = arquivo.Linhas.OrderBy(Function(arqNot) arqNot.TipoRegistro)

            'arquivo.Linhas = selecao.ToList

            label.Text += vbCrLf & "Processando..."
            Dim texto As String = label.Text & vbCrLf
            For Each linha As Layout In arquivo.Linhas
                total += 1
                If linha.CIF.TipoRegistro = "CepErrado" Then
                    ArquivoSaida.escrever(diretorio, nomearquivo & ".err", linha.MontarLinha)
                ElseIf linha.CIF.TipoRegistro = "SemCodigoDeBarras" Then
                    ArquivoSaida.escrever(diretorio, nomearquivo & ".sem", linha.MontarLinha)
                Else
                    sequencia += 1
                    If sequencia = 1 Then
                        ArquivoRelatorio.escreverCabecalho(diretorio, nomearquivo & "_" & linha.CIF.TipoRegistro & ".REL", linha, ordemdeserviço)
                    End If
                    ArquivoRelatorio.escreverDetalhe(diretorio, nomearquivo & "_" & linha.CIF.TipoRegistro & ".REL", linha, sequencia)
                    lote = linha.CIF.CodigoCIF.Substring(10, 5)
                    cont += 1
                    Application.DoEvents()
                    Devolucao.Criar(diretorio, nomearquivo & "_" & lote & ".DEV", linha)
                    linha.IDProcessamento = cont.ToString.PadLeft(7, "0")
                    linha.ModeloFatura = New ModeloFatura(DirectCast(linha, Layout))
                    If SomenteNumeros(linha.Destinatario.Endereco.CEP).PadLeft(8, "0").Substring(0, 1) = "0" Then
                        local += 1
                    ElseIf SomenteNumeros(linha.Destinatario.Endereco.CEP).PadLeft(8, "0").Substring(0, 1) = "1" Then
                        estadual += 1
                    Else
                        nacional += 1
                    End If
                    label.Text = texto & "Definindo os modelos ... " & cont
                    End If
            Next

            ArquivoRelatorio.escreverFim(diretorio, nomearquivo & "_" & "Sehal" & ".REL", local, estadual, nacional)

            Dim selecao As IEnumerable(Of ICIF) = From linhas In arquivo.Linhas Where linhas.CIF.TipoRegistro <> "CepErrado" And linhas.CIF.TipoRegistro <> "SemCodigoDeBarras" Order By linhas.CIF.TipoRegistro, linhas.CIF.CodigoCIF

            arquivo.Linhas = selecao.ToList
            ArquivoPDF.criar(diretorio, nomearquivo, arquivo, imagemcapa, imageminstrucoes, imagemparcela, label)

            If cont > 0 Then
                If producao Then
                    File.Move(diretorio & nomearquivo & "_" & lote & ".DEV", "U:\ArqDev\" & nomearquivo & "_" & lote & ".DEV")
                End If
            End If

            label.Text &= vbCrLf & "Arquivo Processado!!!"
            MessageBox.Show("Arquivo Processado!!!", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
