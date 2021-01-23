Imports System.IO
Imports Marpress.Funcoes

Public Class Devolucao

    Shared Sub Criar(ByVal diretorio As String, ByVal nomeArquivo As String, ByVal fatura As Layout)
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
            With fatura
                Dim linha As String = String.Format("NOMECLI|{0}" & _
                                                    "#CIF|{1}" & _
                                                    "#CPF|{2}" & _
                                                    "#RG|{3}" & _
                                                    "#DOCUMENTO|{4}" & _
                                                    "#CEDENTE|{5}" & _
                                                    "#NOME|{6}" & _
                                                    "#ENDERECO|{7}" & _
                                                    "#NUMERO|{8}" & _
                                                    "#COMPLEMENTO|{9}" & _
                                                    "#BAIRRO|{10}" & _
                                                    "#CIDADE|{11}" & _
                                                    "#ESTADO|{12}" & _
                                                    "#CEP|{13}" & _
                                                    "#VALOR|{14}" & _
                                                    "#VENCIMENTO|{15}" & _
                                                    "#NNUMERO|{16}", _
                                                    .Remetente.Apelido, _
                                                    .CIF.CodigoCIF.Trim, _
                                                    .Destinatario.Documento.Trim, _
                                                    " ", _
                                                    .CodigoCliente.Trim, _
                                                    .Boleto.Beneficiario.Nome.Trim, _
                                                    .Destinatario.Nome.Trim, _
                                                    .Destinatario.Endereco.Logradouro.Trim, _
                                                    .Destinatario.Endereco.Numero, _
                                                    .Destinatario.Endereco.Complemento, _
                                                    .Destinatario.Endereco.Bairro.Trim, _
                                                    .Destinatario.Endereco.Cidade.Trim, _
                                                    .Destinatario.Endereco.Estado, _
                                                    .Destinatario.Endereco.CEP.Trim, _
                                                    .Boleto.Parcelas(0).Valor.Trim.Replace("+", ""), _
                                                    .Boleto.Parcelas(0).Vencimento.Trim.Substring(0, 10), _
                                                    .Boleto.Parcelas(0).NossoNumero.Trim)
                arquivo.WriteLine(linha)
            End With
            arquivo.Flush()
            arquivo.Close()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

End Class