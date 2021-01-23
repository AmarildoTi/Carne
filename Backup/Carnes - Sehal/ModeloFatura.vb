Imports Marpress.Funcoes
Imports Marpress.FichaCompensacao.Itau

Public Class ModeloFatura

    Private _conteudocapa As New List(Of String)
    Public Property ConteudoCapa() As List(Of String)
        Get
            Return _conteudocapa
        End Get
        Set(ByVal value As List(Of String))
            _conteudocapa = value
        End Set
    End Property


    Private _conteudoinstrucoes As New List(Of String)
    Public Property ConteudoInstrucoes() As List(Of String)
        Get
            Return _conteudoinstrucoes
        End Get
        Set(ByVal value As List(Of String))
            _conteudoinstrucoes = value
        End Set
    End Property


    Private _conteudoparcela As New List(Of String)
    Public Property ConteudoParcela() As List(Of String)
        Get
            Return _conteudoparcela
        End Get
        Set(ByVal value As List(Of String))
            _conteudoparcela = value
        End Set
    End Property

    Public Sub New(ByVal fatura As Layout)
        With fatura

            'Comeco Lamina Capa do Carne *****************************************************************************************
            _conteudocapa.Add("Arial|10|0|" & .Destinatario.Nome.Trim)
            _conteudocapa.Add("Arial|10|0|" & .Destinatario.Endereco.Logradouro.Trim)
            _conteudocapa.Add("Arial|10|0|" & .Destinatario.Endereco.Bairro.Trim)
            _conteudocapa.Add("Arial|10|0|" & .Destinatario.Endereco.Cidade.Trim & "  " & .Destinatario.Endereco.Estado)
            _conteudocapa.Add("Arial|10|0|" & SomenteNumeros(.Destinatario.Endereco.CEP).Insert(5, "-"))
            _conteudocapa.Add("Arial|10|0|" & .CIF.CodigoCIF)
            _conteudocapa.Add("Arial|6|0|" & .IDProcessamento)
            'Final Lamina Capa do Carne ******************************************************************************************

            'Começo Lamina Instrucoes ********************************************************************************************
            _conteudoinstrucoes.Add("Arial|6|0|" & .IDProcessamento)
            'Final Lamina Instrucoes *********************************************************************************************

            'Começo Lamina Parcelas do Carne *************************************************************************************
            For i As Integer = 0 To .Boleto.Parcelas.Count - 1

                Dim boletoItau As New Itau(Marpress.FichaCompensacao.Itau.Itau.TipoCarteira.Normal, .Boleto.Carteira, 0, fatura.Agencia, fatura.NumeroConta, .Boleto.Parcelas(i).NossoNumero, .Boleto.Parcelas(i).Valor, .Boleto.Parcelas(i).Vencimento.Substring(0, 10))

                _conteudoparcela.Add("Arial|11|1|" & boletoItau.LinhaDigitavel)
                _conteudoparcela.Add("Arial|9|0|" & .Boleto.LocalDePagamento)

                If i = 0 Then
                    _conteudoparcela.Add("Arial|9|0| ÚNICA")
                Else
                    _conteudoparcela.Add("Arial|9|0|" & CStr(i).PadLeft(2, "0") & "/12")
                End If

                _conteudoparcela.Add("Arial|10|1|" & .Boleto.Parcelas(i).Vencimento.Substring(0, 10))
                _conteudoparcela.Add("Arial|6|0|" & .Boleto.Beneficiario.Nome & Space(1) & .Boleto.Beneficiario.Endereco.Logradouro & Space(1) & .Boleto.Beneficiario.Endereco.Numero & Space(1) & .Boleto.Beneficiario.Endereco.Complemento & Space(1) & .Boleto.Beneficiario.Endereco.Bairro & Space(1) & .Boleto.Beneficiario.Endereco.CEP.Insert(5, "-") & Space(1) & .Boleto.Beneficiario.Endereco.Cidade & Space(1) & .Boleto.Beneficiario.Endereco.Estado & Space(1) & "CNPJ:" & Space(1) & .Boleto.Beneficiario.Documento)
                _conteudoparcela.Add("Arial|9|0|" & .Boleto.AgenciaCodigoBeneficiario)
                _conteudoparcela.Add("Arial|9|0|" & .Boleto.DataProcessamento)
                _conteudoparcela.Add("Arial|9|0|" & .Boleto.Parcelas(i).NumeroDocumento)
                _conteudoparcela.Add("Arial|9|0|" & boletoItau.NossoNumero)
                _conteudoparcela.Add("Arial|8|0| DM")
                _conteudoparcela.Add("Arial|8|0| N")
                _conteudoparcela.Add("Arial|8|0|" & .Boleto.Carteira)
                _conteudoparcela.Add("Arial|9|0| R$")
                _conteudoparcela.Add("Arial|9|1|" & EditaDois(Convert.ToDouble(.Boleto.Parcelas(i).Valor)))
                For Each instrucao In .Boleto.Parcelas(i).Instrucoes
                    _conteudoparcela.Add("Arial|8|0|" & instrucao)
                Next
                _conteudoparcela.Add("Arial|8|1|" & .Destinatario.Nome & Space(1) & "CNPJ:" & Space(1) & .Destinatario.Documento.Trim.Insert(2, ".").Insert(6, ".").Insert(10, "/").Insert(15, "-"))
                _conteudoparcela.Add("Arial|8|0|" & .Destinatario.Endereco.Logradouro.Trim & "  " & .Destinatario.Endereco.Bairro)
                _conteudoparcela.Add("Arial|8|0|" & SomenteNumeros(.Destinatario.Endereco.CEP).Insert(5, "-") & "  " & .Destinatario.Endereco.Cidade.Trim & "  " & .Destinatario.Endereco.Estado)
                _conteudoparcela.Add(boletoItau.CodigoDeBarras)
                _conteudoparcela.Add("Arial|6|0|" & .IDProcessamento)
                'Final Lamina Parcelas do Carne **********************************************************************************
            Next
        End With
    End Sub

End Class