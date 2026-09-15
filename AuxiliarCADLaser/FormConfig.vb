Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class FormConfig
    Inherits Form

    Private txtProjetos As TextBox
    Private txtCN As TextBox
    Private txtFinalizados As TextBox
    Private txtCharmilles As TextBox

    Public Sub New()
        CriarTela()
        CarregarDados()
    End Sub

    Private Sub CriarTela()
        Me.Text = "Configurações"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ClientSize = New Size(760, 500)
        Me.MinimumSize = New Size(700, 480)

        Me.Controls.Add(New Label() With {
            .Text = "Configurações do Auxiliar",
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .Location = New Point(25, 18),
            .AutoSize = True
        })
        Me.Controls.Add(New Label() With {
            .Text = "Defina os caminhos usados neste computador. Caminhos de rede são aceitos.",
            .Location = New Point(28, 58),
            .AutoSize = True,
            .Tag = "suave"
        })

        txtProjetos = CriarLinha("Pasta raiz dos projetos", 95, AddressOf SelecionarProjetos)
        txtCN = CriarLinha("Pasta Programa CN", 175, AddressOf SelecionarCN)
        txtFinalizados = CriarLinha("Pasta Finalizados", 255, AddressOf SelecionarFinalizados)
        txtCharmilles = CriarLinha("Pasta Laser Charmilles", 335, AddressOf SelecionarCharmilles)

        Dim btnSalvar As New Button() With {
            .Text = "Salvar configurações",
            .Location = New Point(540, 430),
            .Size = New Size(175, 38),
            .Tag = "principal"
        }
        AddHandler btnSalvar.Click, AddressOf Salvar
        Me.Controls.Add(btnSalvar)

        Dim btnCancelar As New Button() With {
            .Text = "Cancelar",
            .Location = New Point(415, 430),
            .Size = New Size(110, 38),
            .Tag = "neutro"
        }
        AddHandler btnCancelar.Click, Sub() Me.Close()
        Me.Controls.Add(btnCancelar)

        TemaVisual.AplicarTema(Me)
    End Sub

    Private Function CriarLinha(titulo As String, y As Integer, acao As EventHandler) As TextBox
        Me.Controls.Add(New Label() With {
            .Text = titulo,
            .Location = New Point(30, y),
            .AutoSize = True
        })

        Dim txt As New TextBox() With {
            .Location = New Point(30, y + 24),
            .Size = New Size(550, 28)
        }
        Me.Controls.Add(txt)

        Dim btn As New Button() With {
            .Text = "Selecionar",
            .Location = New Point(595, y + 22),
            .Size = New Size(120, 32),
            .Tag = "neutro"
        }
        AddHandler btn.Click, acao
        Me.Controls.Add(btn)
        Return txt
    End Function

    Private Sub CarregarDados()
        Dim config As ConfiguracaoAuxiliar = ModuloFuncoes.CarregarConfiguracao()
        txtProjetos.Text = config.PastaProjetos
        txtCN.Text = If(String.IsNullOrWhiteSpace(config.PastaCN), ModuloFuncoes.PASTA_CN_PADRAO, config.PastaCN)
        txtFinalizados.Text = config.PastaFinalizados
        txtCharmilles.Text = config.PastaCharmilles
    End Sub

    Private Sub SelecionarProjetos(sender As Object, e As EventArgs)
        SelecionarParaCampo(txtProjetos, "Selecione a pasta raiz que contém os projetos")
    End Sub

    Private Sub SelecionarCN(sender As Object, e As EventArgs)
        SelecionarParaCampo(txtCN, "Selecione a pasta do Programa CN")
    End Sub

    Private Sub SelecionarFinalizados(sender As Object, e As EventArgs)
        SelecionarParaCampo(txtFinalizados, "Selecione a pasta dos Finalizados")
    End Sub

    Private Sub SelecionarCharmilles(sender As Object, e As EventArgs)
        SelecionarParaCampo(txtCharmilles, "Selecione a pasta do Laser Charmilles")
    End Sub

    Private Sub SelecionarParaCampo(campo As TextBox, titulo As String)
        Dim pasta As String = ModuloFuncoes.SelecionarPasta(titulo, campo.Text)
        If pasta <> "" Then campo.Text = pasta
    End Sub

    Private Sub Salvar(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtProjetos.Text) Then
            MessageBox.Show("Configure a pasta raiz dos projetos.", "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProjetos.Focus()
            Return
        End If

        Dim config As ConfiguracaoAuxiliar = ModuloFuncoes.CarregarConfiguracao()
        config.PastaProjetos = txtProjetos.Text.Trim()
        config.PastaCN = txtCN.Text.Trim()
        config.PastaFinalizados = txtFinalizados.Text.Trim()
        config.PastaCharmilles = txtCharmilles.Text.Trim()
        ModuloFuncoes.SalvarConfiguracao(config)

        MessageBox.Show("Configurações salvas.", "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
