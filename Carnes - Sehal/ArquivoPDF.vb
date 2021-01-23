Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
Imports iTextSharp.text.BaseColor
Imports Marpress.Funcoes


Public Class ArquivoPDF

    Public Shared Sub criar(ByVal diretorio As String, ByVal nomearquivo As String, ByVal faturas As ArquivoFatura, ByVal imagemcapa As String, ByVal imageminstrucoes As String, ByVal imagemparcela As String, ByVal label As Label)
        Try
            If Not Directory.Exists(diretorio) Then
                Directory.CreateDirectory(diretorio)
            End If
            For i As Integer = 1 To 1

                Dim processamento As String = "Sehal"

                If i = 1 Then
                    imagemcapa = "sehal_capa.pdf"
                    imageminstrucoes = "sehal_instrucoes.pdf"
                    imagemparcela = "sehal_parcela.pdf"
                End If

                Dim impressao As New List(Of Layout)
                Dim local As Integer = 0
                Dim estadual As Integer = 0
                Dim nacional As Integer = 0
                Dim cont As Integer = 0
                Dim cont2 As Integer = 0
                Dim indice As Integer = 0
                Dim nomearquivosaida As String = ""
                Dim lote As String = ""
                Dim arq As Integer = 1
                Dim texto As String = label.Text & vbCrLf
                Dim documento As Document
                Dim escritor As PdfWriter
                Dim codpostagem As String = ""
                Dim codadministrativo As String = ""

                For Each fatura As Layout In faturas.Linhas
                    If fatura.CIF.TipoRegistro = processamento Then
                        cont += 1
                        cont2 += 1
                        indice += 1
                        codpostagem = fatura.CIF.CodigoPostagem
                        codadministrativo = fatura.CIF.CodigoAdministrativo
                        If cont2 > 500 Then
                            cont2 = 1
                            arq += 1
                            documento.Close()
                            escritor.Close()
                            If cont > 1 Then ArquivoOS.Fatura(diretorio, nomearquivo & "_" & Right(processamento, 3) & ".os", nomearquivosaida & ".pdf", local, estadual, nacional, lote, codpostagem, codadministrativo)
                            local = 0
                            estadual = 0
                            nacional = 0
                        End If
                        nomearquivosaida = nomearquivo & "_" & Right(fatura.CIF.TipoRegistro, 5) & "_" & arq.ToString.PadLeft(3, "0")
                        If SomenteNumeros(fatura.Destinatario.Endereco.CEP).PadLeft(8, "0").Substring(0, 1) = "0" Then
                            local += 1
                        ElseIf SomenteNumeros(fatura.Destinatario.Endereco.CEP).PadLeft(8, "0").Substring(0, 1) = "1" Then
                            estadual += 1
                        Else
                            nacional += 1
                        End If
                        lote = fatura.CIF.CodigoCIF.Substring(10, 5)
                        label.Text = texto & "Gerando os PDF's (" & nomearquivosaida & ") ..." & cont

                        DirectCast(fatura, Layout).ModeloFatura.ConteudoCapa(6) &= Space(2) & "Arq:" & nomearquivosaida
                        DirectCast(fatura, Layout).ModeloFatura.ConteudoInstrucoes(0) &= Space(2) & "Arq:" & nomearquivosaida
                        DirectCast(fatura, Layout).ModeloFatura.ConteudoParcela(24) &= Space(2) & "Arq:" & nomearquivosaida

                        If cont2 = 1 Then
                            Dim pdfnovo As String = diretorio & nomearquivosaida
                            Dim fs As New FileStream(pdfnovo & ".PDF", FileMode.Create, FileAccess.Write, FileShare.None)
                            documento = New Document(PageSize.A4)
                            escritor = PdfWriter.GetInstance(documento, fs)
                        End If
                        Application.DoEvents()

                        documento.Open()

                        Dim cb As PdfContentByte = escritor.DirectContent
                        cb.SetColorFill(BaseColor.BLACK)

                        FontFactory.Register("C:\Windows\Fonts\arial.ttf")
                        FontFactory.Register("C:\Windows\Fonts\arialbd.ttf")
                        FontFactory.Register("C:\Windows\Fonts\COUR.TTF")

                        impressao.Add(fatura)
                        If indice = 3 Then
                            criarCapaSEHAL(escritor, imagemcapa, cb, impressao)
                            documento.NewPage()
                            criarInstrucoesSEHAL(escritor, imageminstrucoes, cb, impressao)
                            documento.NewPage()
                            criarParcelaSEHAL(escritor, imagemparcela, cb, impressao, documento)
                            documento.NewPage()
                            impressao = New List(Of Layout)
                            indice = 0
                        End If
                    End If
                Next
                If cont > 0 Then
                    ArquivoOS.Fatura(diretorio, nomearquivo & "_" & Right(processamento, 5) & ".os", nomearquivosaida & ".pdf", local, estadual, nacional, lote, codpostagem, codadministrativo)
                End If
                If Not documento Is Nothing Then
                    documento.Close()
                    escritor.Close()
                End If
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Shared Sub criarCapaSEHAL(ByVal escritor As PdfWriter, ByVal imagem As String, ByVal cb As PdfContentByte, ByVal faturas As List(Of Layout))
        Try

            Dim rotacao As Double = 0
            Dim begin As Integer = 0

            If imagem.Trim <> "" Then
                Dim img As New PdfReader(diretorioPDF & imagem)
                Dim pagina As PdfImportedPage = escritor.GetImportedPage(img, 1)
                cb.AddTemplate(pagina, 1, 0, 0, 1, 0, 0)
            End If

            For Each fatura In faturas
                Dim ceppostnet As New BarcodePostnet

                ceppostnet.Code = SomenteNumeros(fatura.Destinatario.Endereco.CEP).PadLeft(8, "0")

                Dim datamatrix As New BarcodeDatamatrix
                datamatrix.Generate(DataMatrix2D(fatura))

                Dim codigocif As New Barcode128

                With fatura.ModeloFatura
                    inserirImagem(cb, datamatrix.CreateImage(), 42, 68.5 + begin, 0, 100, rotacao)
                    inserirImagem(cb, ceppostnet.CreateImageWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK), 59, 53.5 + begin, 0, 100, rotacao)
                    escreverFrase(cb, 59, 59.5 + begin, 195, 1, 5, Element.ALIGN_LEFT, 0, rotacao, .ConteudoCapa(0), BaseColor.BLACK)
                    escreverFrase(cb, 59, 63 + begin, 195, 1, 5, Element.ALIGN_LEFT, 0, rotacao, .ConteudoCapa(1), BaseColor.BLACK)
                    escreverFrase(cb, 59, 66.5 + begin, 195, 1, 5, Element.ALIGN_LEFT, 0, rotacao, .ConteudoCapa(2), BaseColor.BLACK)
                    escreverFrase(cb, 59, 70 + begin, 195, 1, 5, Element.ALIGN_LEFT, 0, rotacao, .ConteudoCapa(3), BaseColor.BLACK)
                    escreverFrase(cb, 59, 73.5 + begin, 195, 1, 5, Element.ALIGN_LEFT, 0, rotacao, .ConteudoCapa(4), BaseColor.BLACK)

                    escreverFrase(cb, 15, 93 + begin, 195, 1, 5, Element.ALIGN_CENTER, 0, rotacao, .ConteudoCapa(5), BaseColor.BLACK)
                    codigocif.Code = .ConteudoCapa(5).Replace("Arial|10|0|", "")
                    codigocif.Font = Nothing
                    codigocif.BarHeight = Utilities.MillimetersToPoints(10)
                    inserirImagem(cb, codigocif.CreateImageWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK), 73.5, 90 + begin, 0, 100, rotacao)

                    escreverFrase(cb, 5, 45.5 + begin, 0, 1, 5, Element.ALIGN_RIGHT, 0, 90, .ConteudoCapa(6), BaseColor.BLACK)
                    begin += 99
                End With
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Shared Sub criarInstrucoesSEHAL(ByVal escritor As PdfWriter, ByVal imagem As String, ByVal cb As PdfContentByte, ByVal faturas As List(Of Layout))
        Try

            Dim begin As Integer = 0

            If imagem.Trim <> "" Then
                Dim img As New PdfReader(diretorioPDF & imagem)
                Dim pagina As PdfImportedPage = escritor.GetImportedPage(img, 1)
                cb.AddTemplate(pagina, 1, 0, 0, 1, 0, 0)
            End If

            For Each fatura In faturas
                With fatura.ModeloFatura
                    escreverFrase(cb, 5, 45.5 + begin, 0, 1, 5, Element.ALIGN_RIGHT, 0, 90, .ConteudoInstrucoes(0), BaseColor.BLACK)
                    begin += 99
                End With
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Private Shared Sub criarParcelaSEHAL(ByVal escritor As PdfWriter, ByVal imagem As String, ByVal cb As PdfContentByte, ByVal faturas As List(Of Layout), ByVal documento As Document)
        Try
            Dim begin As Double = 0.0
            Dim posicao As Double = 0.0
            Dim ind As Integer = 0

            For i As Integer = 0 To faturas(0).Boleto.Parcelas.Count - 1

                If imagem.Trim <> "" Then
                    Dim img As New PdfReader(diretorioPDF & imagem)
                    Dim pagina As PdfImportedPage = escritor.GetImportedPage(img, 1)
                    cb.AddTemplate(pagina, 1, 0, 0, 1, 0, 0)
                End If

                For Each fatura In faturas

                    Dim code2de5 As New BarcodeInter25
                    Dim y As Integer = 0

                
                    With fatura.ModeloFatura

                        'Começo Recibo do Sacado **********************************************************************************************
                        escreverFrase(cb, 15, 81 + begin, 50, 1, 3, Element.ALIGN_LEFT, 1, 0, .ConteudoParcela(ind + 20), BaseColor.BLACK)
                        escreverFrase(cb, 42, 43.5 + begin, 10, 1, 3, Element.ALIGN_RIGHT, 0, 0, .ConteudoParcela(ind + 13), BaseColor.BLACK)
                        escreverFrase(cb, 26, 25 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 3), BaseColor.BLACK)
                        escreverFrase(cb, 24, 31.5 + begin, 200, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 5), BaseColor.BLACK)
                        escreverFrase(cb, 30, 18 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 7), BaseColor.BLACK)
                        escreverFrase(cb, 20, 37.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 8), BaseColor.BLACK)
                        'Final Recibo do Sacado **********************************************************************************************


                        'Começo Ficha de Compensação ******************************************************************************************
                        escreverFrase(cb, 93, 10 + begin, 215, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 0), BaseColor.BLACK)
                        escreverFrase(cb, 55, 18 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 1), BaseColor.BLACK)
                        escreverFrase(cb, 151, 18 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 2), BaseColor.BLACK)
                        escreverFrase(cb, 177, 18 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 3), BaseColor.BLACK)
                        escreverFrase(cb, 68, 22 + begin, 165, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 4), BaseColor.BLACK)
                        escreverFrase(cb, 175, 25 + begin, 200, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 5), BaseColor.BLACK)
                        escreverFrase(cb, 58, 31.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 6), BaseColor.BLACK)
                        escreverFrase(cb, 145, 31.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 6), BaseColor.BLACK)
                        escreverFrase(cb, 90, 31.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 7), BaseColor.BLACK)
                        escreverFrase(cb, 171, 31.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 8), BaseColor.BLACK)
                        escreverFrase(cb, 124.5, 31.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 9), BaseColor.BLACK)
                        escreverFrase(cb, 136.5, 31.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 10), BaseColor.BLACK)
                        escreverFrase(cb, 83, 37.5 + begin, 115, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 11), BaseColor.BLACK)
                        escreverFrase(cb, 91.5, 37.5 + begin, 125, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 12), BaseColor.BLACK)
                        escreverFrase(cb, 100, 37.5 + begin, 194, 1, 3, Element.ALIGN_RIGHT, 0, 0, .ConteudoParcela(ind + 13), BaseColor.BLACK)

                        y = 50 + begin
                        For j As Integer = ind + 14 To ind + 19
                            escreverFrase(cb, 55, y, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(j), BaseColor.BLACK)
                            y += 3
                        Next

                        escreverFrase(cb, 64, 72 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 20), BaseColor.BLACK)
                        escreverFrase(cb, 64, 75 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 21), BaseColor.BLACK)
                        escreverFrase(cb, 64, 77.5 + begin, 195, 1, 3, Element.ALIGN_LEFT, 0, 0, .ConteudoParcela(ind + 22), BaseColor.BLACK)

                        code2de5.Code = .ConteudoParcela(ind + 23)
                        code2de5.BarHeight = 36
                        code2de5.Font = Nothing
                        code2de5.N = 2.5
                        cb.AddTemplate(code2de5.CreateTemplateWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK), Utilities.MillimetersToPoints(64), Utilities.MillimetersToPoints(297 - (94 + begin)))

                        escreverFrase(cb, 5, 45.5 + begin, 0, 1, 5, Element.ALIGN_RIGHT, 0, 90, .ConteudoParcela(24), BaseColor.BLACK)
                        'Final Ficha de Compensação **********************************************************************************************

                        begin += 99
                    End With
                Next
                begin = 0
                documento.NewPage()
                ind += 25
            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

End Class

