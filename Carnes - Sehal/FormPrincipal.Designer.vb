<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPrincipal
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormPrincipal))
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBoxArquivo = New System.Windows.Forms.TextBox
        Me.ButtonArquivo = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.ComboBoxProducao = New System.Windows.Forms.ComboBox
        Me.DateTimePickerPostagem = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.ButtonProcessar = New System.Windows.Forms.Button
        Me.ComboBoxContratoFAC = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.LabelProcessamento = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ComboBoxTipoFAC = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBoxOrdemDeServiço = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 16)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Arquivo:"
        '
        'TextBoxArquivo
        '
        Me.TextBoxArquivo.Location = New System.Drawing.Point(12, 31)
        Me.TextBoxArquivo.Name = "TextBoxArquivo"
        Me.TextBoxArquivo.ReadOnly = True
        Me.TextBoxArquivo.Size = New System.Drawing.Size(454, 22)
        Me.TextBoxArquivo.TabIndex = 4
        '
        'ButtonArquivo
        '
        Me.ButtonArquivo.Image = CType(resources.GetObject("ButtonArquivo.Image"), System.Drawing.Image)
        Me.ButtonArquivo.Location = New System.Drawing.Point(472, 9)
        Me.ButtonArquivo.Name = "ButtonArquivo"
        Me.ButtonArquivo.Size = New System.Drawing.Size(44, 44)
        Me.ButtonArquivo.TabIndex = 5
        Me.ButtonArquivo.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(12, 102)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(141, 16)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Tipo Processamento:"
        '
        'ComboBoxProducao
        '
        Me.ComboBoxProducao.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxProducao.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxProducao.FormattingEnabled = True
        Me.ComboBoxProducao.Location = New System.Drawing.Point(12, 119)
        Me.ComboBoxProducao.Name = "ComboBoxProducao"
        Me.ComboBoxProducao.Size = New System.Drawing.Size(138, 24)
        Me.ComboBoxProducao.TabIndex = 17
        '
        'DateTimePickerPostagem
        '
        Me.DateTimePickerPostagem.Location = New System.Drawing.Point(156, 121)
        Me.DateTimePickerPostagem.Name = "DateTimePickerPostagem"
        Me.DateTimePickerPostagem.Size = New System.Drawing.Size(310, 22)
        Me.DateTimePickerPostagem.TabIndex = 16
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(156, 102)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(108, 16)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Data Postagem:"
        '
        'ButtonProcessar
        '
        Me.ButtonProcessar.BackColor = System.Drawing.Color.Transparent
        Me.ButtonProcessar.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonProcessar.Location = New System.Drawing.Point(317, 149)
        Me.ButtonProcessar.Name = "ButtonProcessar"
        Me.ButtonProcessar.Size = New System.Drawing.Size(149, 44)
        Me.ButtonProcessar.TabIndex = 14
        Me.ButtonProcessar.Text = "Processar Arquivo"
        Me.ButtonProcessar.UseVisualStyleBackColor = False
        '
        'ComboBoxContratoFAC
        '
        Me.ComboBoxContratoFAC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxContratoFAC.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxContratoFAC.FormattingEnabled = True
        Me.ComboBoxContratoFAC.Location = New System.Drawing.Point(156, 75)
        Me.ComboBoxContratoFAC.Name = "ComboBoxContratoFAC"
        Me.ComboBoxContratoFAC.Size = New System.Drawing.Size(310, 24)
        Me.ComboBoxContratoFAC.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(156, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(95, 16)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Contrato FAC:"
        '
        'LabelProcessamento
        '
        Me.LabelProcessamento.Location = New System.Drawing.Point(12, 196)
        Me.LabelProcessamento.Name = "LabelProcessamento"
        Me.LabelProcessamento.Size = New System.Drawing.Size(504, 258)
        Me.LabelProcessamento.TabIndex = 21
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 16)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Tipo FAC"
        '
        'ComboBoxTipoFAC
        '
        Me.ComboBoxTipoFAC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxTipoFAC.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxTipoFAC.FormattingEnabled = True
        Me.ComboBoxTipoFAC.Location = New System.Drawing.Point(12, 75)
        Me.ComboBoxTipoFAC.Name = "ComboBoxTipoFAC"
        Me.ComboBoxTipoFAC.Size = New System.Drawing.Size(138, 24)
        Me.ComboBoxTipoFAC.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 146)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(127, 16)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Ordem de Serviço:"
        '
        'TextBoxOrdemDeServiço
        '
        Me.TextBoxOrdemDeServiço.Location = New System.Drawing.Point(12, 165)
        Me.TextBoxOrdemDeServiço.Name = "TextBoxOrdemDeServiço"
        Me.TextBoxOrdemDeServiço.Size = New System.Drawing.Size(138, 22)
        Me.TextBoxOrdemDeServiço.TabIndex = 25
        '
        'FormPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(527, 463)
        Me.Controls.Add(Me.TextBoxOrdemDeServiço)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxTipoFAC)
        Me.Controls.Add(Me.LabelProcessamento)
        Me.Controls.Add(Me.ComboBoxContratoFAC)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.ComboBoxProducao)
        Me.Controls.Add(Me.DateTimePickerPostagem)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.ButtonProcessar)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBoxArquivo)
        Me.Controls.Add(Me.ButtonArquivo)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "FormPrincipal"
        Me.ShowInTaskbar = False
        Me.Text = "Carnes - Sehal"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBoxArquivo As System.Windows.Forms.TextBox
    Friend WithEvents ButtonArquivo As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxProducao As System.Windows.Forms.ComboBox
    Friend WithEvents DateTimePickerPostagem As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ButtonProcessar As System.Windows.Forms.Button
    Friend WithEvents ComboBoxContratoFAC As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents LabelProcessamento As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxTipoFAC As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBoxOrdemDeServiço As System.Windows.Forms.TextBox

End Class
