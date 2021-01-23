Imports Marpress
Imports Marpress.Funcoes

Public Class FormPrincipal

    Private diretorio As String
    Private nomeArquivo As String

    Private Sub FormPrincipal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CarregarFormulario()
    End Sub

    Private Sub CarregarFormulario()
        ComboBoxContratoFAC.DataSource = FAC.Enumeradores.Listar(GetType(FAC.Contrato))
        ComboBoxContratoFAC.ValueMember = "Key"
        ComboBoxContratoFAC.DisplayMember = "Value"

        ComboBoxProducao.DataSource = FAC.Enumeradores.Listar(GetType(FAC.Processamento))
        ComboBoxProducao.ValueMember = "Key"
        ComboBoxProducao.DisplayMember = "Value"

        ComboBoxTipoFAC.DataSource = FAC.Enumeradores.Listar(GetType(FAC.Tipo))
        ComboBoxTipoFAC.ValueMember = "Key"
        ComboBoxTipoFAC.DisplayMember = "Value"

        Dim data = Date.Now
        data = DateAdd(DateInterval.Day, 2, data)
        If data.DayOfWeek = DayOfWeek.Saturday Then
            data = DateAdd(DateInterval.Day, 2, data)
        ElseIf data.DayOfWeek = DayOfWeek.Sunday Then
            data = DateAdd(DateInterval.Day, 1, data)
        End If
        DateTimePickerPostagem.Value = data
        DateTimePickerPostagem.MinDate = Date.Now
        DateTimePickerPostagem.MaxDate = DateAdd(DateInterval.Day, 30, Date.Now)
    End Sub

    Private Sub ButtonArquivo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonArquivo.Click
        Dim ofd As New OpenFileDialog
        Dim resp As DialogResult
        ofd.Multiselect = False
        resp = ofd.ShowDialog()
        If resp <> Windows.Forms.DialogResult.Cancel Then
            TextBoxArquivo.Text = ofd.FileName
            diretorio = ofd.FileName.Replace(ofd.SafeFileName, "")
            nomeArquivo = ofd.SafeFileName
        End If
    End Sub

    Private Sub ButtonProcessar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonProcessar.Click
        If TextBoxArquivo.Text.Trim = "" Then
            MessageBox.Show("Selecione um arquivo para ser processado", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ButtonArquivo.Focus()
            Exit Sub
        End If
        If TextBoxOrdemDeServiço.Text.Trim = "" Then
            MessageBox.Show("Favor informar o número da Ordem de Serviço", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBoxOrdemDeServiço.Focus()
            Exit Sub
        End If
        DateTimePickerPostagem.Value = validaDataPostagem(DateTimePickerPostagem.Value)
        Processamento.processa(diretorio, nomeArquivo, ComboBoxTipoFAC.SelectedValue, ComboBoxContratoFAC.SelectedValue, LabelProcessamento, DateTimePickerPostagem.Value, ComboBoxProducao.SelectedValue, TextBoxOrdemDeServiço.Text)
    End Sub

End Class
