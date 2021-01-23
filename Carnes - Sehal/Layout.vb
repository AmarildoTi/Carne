Imports Marpress.Interfaces
Imports Marpress.Padrao.Fatura
Imports Marpress.Padrao
Imports Marpress.Funcoes

Public Class Layout
    Inherits Fatura

    ' ********************** Começo Propriedades do Boleto *************************

    Private _agencia As String
    Public Property Agencia() As String
        Get
            Return _agencia
        End Get
        Set(ByVal value As String)
            _agencia = value
        End Set
    End Property


    Private _numeroconta As String
    Public Property NumeroConta() As String
        Get
            Return _numeroconta
        End Get
        Set(ByVal value As String)
            _numeroconta = value
        End Set
    End Property


    ' ********************** Começo Propriedades do Contador *************************
    Private _nomecontador As String
    Public Property NomeContador() As String
        Get
            Return _nomecontador
        End Get
        Set(ByVal value As String)
            _nomecontador = value
        End Set
    End Property


    Private _enderecocontador As String
    Public Property EnderecoContador() As String
        Get
            Return _enderecocontador
        End Get
        Set(ByVal value As String)
            _enderecocontador = value
        End Set
    End Property


    Private _numerocontador As String
    Public Property NumeroContador() As String
        Get
            Return _numerocontador
        End Get
        Set(ByVal value As String)
            _numerocontador = value
        End Set
    End Property


    Private _complementocontador As String
    Public Property ComplementoContador() As String
        Get
            Return _complementocontador
        End Get
        Set(ByVal value As String)
            _complementocontador = value
        End Set
    End Property


    Private _bairrocontador As String
    Public Property BairroContador() As String
        Get
            Return _bairrocontador
        End Get
        Set(ByVal value As String)
            _bairrocontador = value
        End Set
    End Property


    Private _cidadecontador As String
    Public Property CidadeContador() As String
        Get
            Return _cidadecontador
        End Get
        Set(ByVal value As String)
            _cidadecontador = value
        End Set
    End Property


    Private _estadocontador As String
    Public Property EstadoContador() As String
        Get
            Return _estadocontador
        End Get
        Set(ByVal value As String)
            _estadocontador = value
        End Set
    End Property


    Private _cepcontador As String
    Public Property CepContador() As String
        Get
            Return _cepcontador
        End Get
        Set(ByVal value As String)
            _cepcontador = value
        End Set
    End Property

    ' ********************** Começo Propriedades do Contador *************************

    Private _modelofatura As ModeloFatura
    Public Property ModeloFatura() As ModeloFatura
        Get
            Return _modelofatura
        End Get
        Set(ByVal value As ModeloFatura)
            _modelofatura = value
        End Set
    End Property

    Private _idprocessamento As String
    Public Property IDProcessamento() As String
        Get
            Return _idprocessamento
        End Get
        Set(ByVal value As String)
            _idprocessamento = value
        End Set
    End Property

    Public Sub CarregaRegistros(ByVal linha As String)
        Try
            Dim campo As String() = linha.Split(";")
            Dim Data_hoje As Date = Date.Now

            ' Dados do Remetente / Beneficiario
            CIF.TipoRegistro = "Sehal"
            Remetente.Apelido = "Sehal"
            Remetente.Nome = campo(0)
            Remetente.Endereco.Logradouro = campo(1)
            Remetente.Endereco.Numero = campo(2)
            Remetente.Endereco.Complemento = campo(3)
            Remetente.Endereco.Bairro = campo(4)
            Remetente.Endereco.Cidade = campo(5)
            Remetente.Endereco.Estado = campo(6)
            Remetente.Endereco.CEP = campo(7)
            Remetente.Documento = campo(8)

            ' Dados do Cliente
            CodigoCliente = campo(9)
            Destinatario.Nome = campo(10)
            Destinatario.Endereco.Logradouro = campo(11)
            Destinatario.Endereco.Numero = campo(12)
            Destinatario.Endereco.Complemento = campo(13)
            Destinatario.Endereco.Bairro = campo(14)
            Destinatario.Endereco.Cidade = campo(15)
            Destinatario.Endereco.Estado = campo(16)
            Destinatario.Endereco.CEP = campo(17)
            Destinatario.Documento = campo(18)

            ' Dados do Contador
            NomeContador = campo(29)
            EnderecoContador = campo(30)
            NumeroContador = campo(31)
            ComplementoContador = campo(32)
            BairroContador = campo(33)
            CidadeContador = campo(34)
            EstadoContador = campo(35)
            CepContador = campo(36)

            ' Dados das Parcelas
            Boleto.Beneficiario.Nome = campo(0)
            Boleto.Beneficiario.Endereco.Logradouro = campo(1)
            Boleto.Beneficiario.Endereco.Numero = campo(2)
            Boleto.Beneficiario.Endereco.Complemento = campo(3)
            Boleto.Beneficiario.Endereco.Bairro = campo(4)
            Boleto.Beneficiario.Endereco.Cidade = campo(5)
            Boleto.Beneficiario.Endereco.Estado = campo(6)
            Boleto.Beneficiario.Endereco.CEP = campo(7)
            Boleto.Beneficiario.Documento = campo(8)
            Boleto.AgenciaCodigoBeneficiario = campo(23) & "/" & campo(24)
            Boleto.Carteira = campo(25)
            Boleto.LocalDePagamento = campo(37)
            Agencia = campo(23)
            NumeroConta = campo(24).Substring(0, campo(24).Trim.Length - 2)
            Boleto.DataProcessamento = Data_hoje.Day.ToString.PadLeft(2, "0") & "/" & Data_hoje.Month.ToString.PadLeft(2, "0") & "/" & Data_hoje.Year


            CarregaParcelas(linha)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub CarregaParcelas(ByVal linha As String)
        Try
            Dim campo As String() = linha.Split(";")
            Boleto.Parcelas.Add(New Parcela)

            Boleto.Parcelas(Boleto.Parcelas.Count - 1).NumeroDocumento = campo(21)
            Boleto.Parcelas(Boleto.Parcelas.Count - 1).NossoNumero = campo(21)
            Boleto.Parcelas(Boleto.Parcelas.Count - 1).Vencimento = campo(19)
            Boleto.Parcelas(Boleto.Parcelas.Count - 1).Valor = campo(20)

            For i As Integer = 38 To 43
                Boleto.Parcelas(Boleto.Parcelas.Count - 1).Instrucoes.Add(campo(i))
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Function MontarLinha() As String
        Dim linha As String = "01"
        'Try
        '    linha += _tipoFatura
        '    linha += _correio
        '    linha += _codigoAssociado
        '    linha += _destinario.Nome
        '    linha += _destinario.Endereco.Logradouro
        '    linha += _destinario.Endereco.Bairro
        '    linha += _destinario.Endereco.CEP
        '    linha += _destinario.Endereco.Cidade
        '    linha += _destinario.Endereco.Estado
        '    linha += _titulo
        '    linha += _dataemissao
        '    linha += _dataVencimento
        '    linha += _identificadortitulo
        '    linha += _banco
        '    linha += _agencia
        '    linha += _conta
        '    linha += _contadv
        '    linha += _carteira
        '    linha += _especie
        '    linha += _mora
        '    linha += _desconto
        '    linha += _datadesconto
        '    linha += _abatimento
        '    linha += _valorDocumento
        '    linha += _mensagem(0)
        '    linha += _mensagem(1)
        '    linha += _mensagem(2)
        '    linha += _mensagem(3)
        '    linha += _quantidade
        '    linha += _linha
        '    linha += _demonstrativo(0)
        '    linha += _destinario.Documento
        '    linha += _reservado
        '    linha += _sequencia
        'Catch ex As Exception
        'MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        Return linha
    End Function

End Class
