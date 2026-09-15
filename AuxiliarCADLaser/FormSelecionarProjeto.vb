Imports System.Drawing
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class FormSelecionarProjeto
    Inherits Form

    Private lista As ListView
    Private lblCaminho As Label
    Public Property ProjetoSelecionado As ProjetoEncontrado

    Public Sub New(codigo As String, resultados As List(Of ProjetoEncontrado))
        Me.Text = "Projetos encontrados para " & codigo
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ClientSize = New Size(920, 500)
        Me.MinimumSize = New Size(760, 430)
        CriarTela(codigo, If(resultados, New List(Of ProjetoEncontrado)()))
        TemaVisual.AplicarTema(Me)
    End Sub

    Private Sub CriarTela(codigo As String, resultados As List(Of ProjetoEncontrado))
        Me.Controls.Add(New Label() With {
            .Text = "Mais de um projeto possui o código " & codigo,
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .Location = New Point(25, 20),
            .AutoSize = True
        })
        Me.Controls.Add(New Label() With {
            .Text = "Escolha a pasta que deseja utilizar. Nenhum arquivo será alterado nesta etapa.",
            .Location = New Point(28, 60),
            .AutoSize = True,
            .Tag = "suave"
        })

        lista = New ListView() With {
            .Location = New Point(28, 95),
            .Size = New Size(Me.ClientSize.Width - 56, Me.ClientSize.Height - 195),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .View = View.Details,
            .FullRowSelect = True,
            .MultiSelect = False,
            .HideSelection = False,
            .BackColor = TemaVisual.CorCampo,
            .ForeColor = TemaVisual.CorTexto
        }
        lista.Columns.Add("Pasta", 300)
        lista.Columns.Add("Última alteração", 145)
        lista.Columns.Add("Caminho", 600)
        AddHandler lista.SelectedIndexChanged, AddressOf SelecaoMudou
        AddHandler lista.DoubleClick, AddressOf Confirmar
        Me.Controls.Add(lista)

        For Each projeto As ProjetoEncontrado In resultados
            Dim item As New ListViewItem(projeto.NomePasta)
            item.SubItems.Add(projeto.DataAlteracao.ToString("dd/MM/yyyy HH:mm"))
            item.SubItems.Add(projeto.Caminho)
            item.Tag = projeto
            lista.Items.Add(item)
        Next

        lblCaminho = New Label() With {
            .Text = "Selecione uma pasta.",
            .Location = New Point(30, Me.ClientSize.Height - 82),
            .Size = New Size(Me.ClientSize.Width - 360, 45),
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom,
            .Tag = "suave"
        }
        Me.Controls.Add(lblCaminho)

        Dim btnAbrir As New Button() With {
            .Text = "Abrir pasta",
            .Location = New Point(Me.ClientSize.Width - 330, Me.ClientSize.Height - 75),
            .Size = New Size(120, 36),
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Tag = "neutro"
        }
        AddHandler btnAbrir.Click, AddressOf AbrirPasta
        Me.Controls.Add(btnAbrir)

        Dim btnUsar As New Button() With {
            .Text = "Usar projeto",
            .Location = New Point(Me.ClientSize.Width - 195, Me.ClientSize.Height - 75),
            .Size = New Size(165, 36),
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Tag = "principal"
        }
        AddHandler btnUsar.Click, AddressOf Confirmar
        Me.Controls.Add(btnUsar)

        If lista.Items.Count > 0 Then lista.Items(0).Selected = True
    End Sub

    Private Function Selecionado() As ProjetoEncontrado
        If lista.SelectedItems.Count = 0 Then Return Nothing
        Return TryCast(lista.SelectedItems(0).Tag, ProjetoEncontrado)
    End Function

    Private Sub SelecaoMudou(sender As Object, e As EventArgs)
        Dim projeto As ProjetoEncontrado = Selecionado()
        lblCaminho.Text = If(projeto Is Nothing, "Selecione uma pasta.", projeto.Caminho)
    End Sub

    Private Sub AbrirPasta(sender As Object, e As EventArgs)
        Dim projeto As ProjetoEncontrado = Selecionado()
        If projeto Is Nothing Then Return
        Process.Start(New ProcessStartInfo("explorer.exe", "\"" & projeto.Caminho & "\"") With {.UseShellExecute = True})
    End Sub

    Private Sub Confirmar(sender As Object, e As EventArgs)
        Dim projeto As ProjetoEncontrado = Selecionado()
        If projeto Is Nothing Then
            MessageBox.Show("Selecione uma pasta.", "Projeto", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ProjetoSelecionado = projeto
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
