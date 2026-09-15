Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class OpcaoEscolhaAuxiliar
    Public Property Texto As String = ""
    Public Property Valor As String = ""

    Public Overrides Function ToString() As String
        Return Texto
    End Function
End Class

Public Class FormEscolherOpcao
    Inherits Form

    Private lista As ListBox
    Public Property ValorSelecionado As String = ""

    Public Sub New(titulo As String, mensagem As String, opcoes As List(Of OpcaoEscolhaAuxiliar))
        Me.Text = titulo
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ClientSize = New Size(720, 410)
        Me.MinimumSize = New Size(620, 360)
        CriarTela(mensagem, If(opcoes, New List(Of OpcaoEscolhaAuxiliar)()))
        TemaVisual.AplicarTema(Me)
    End Sub

    Private Sub CriarTela(mensagem As String, opcoes As List(Of OpcaoEscolhaAuxiliar))
        Me.Controls.Add(New Label() With {
            .Text = mensagem,
            .Font = New Font("Segoe UI", 15, FontStyle.Bold),
            .Location = New Point(25, 22),
            .Size = New Size(665, 58)
        })

        lista = New ListBox() With {
            .Location = New Point(28, 92),
            .Size = New Size(664, 235),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .HorizontalScrollbar = True
        }
        For Each opcao As OpcaoEscolhaAuxiliar In opcoes
            lista.Items.Add(opcao)
        Next
        AddHandler lista.DoubleClick, AddressOf Confirmar
        Me.Controls.Add(lista)

        Dim btnCancelar As New Button() With {
            .Text = "Cancelar",
            .Location = New Point(Me.ClientSize.Width - 275, Me.ClientSize.Height - 60),
            .Size = New Size(110, 36),
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Tag = "neutro"
        }
        AddHandler btnCancelar.Click, Sub() Me.Close()
        Me.Controls.Add(btnCancelar)

        Dim btnConfirmar As New Button() With {
            .Text = "Selecionar",
            .Location = New Point(Me.ClientSize.Width - 150, Me.ClientSize.Height - 60),
            .Size = New Size(122, 36),
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Tag = "principal"
        }
        AddHandler btnConfirmar.Click, AddressOf Confirmar
        Me.Controls.Add(btnConfirmar)

        If lista.Items.Count > 0 Then lista.SelectedIndex = 0
    End Sub

    Private Sub Confirmar(sender As Object, e As EventArgs)
        Dim opcao As OpcaoEscolhaAuxiliar = TryCast(lista.SelectedItem, OpcaoEscolhaAuxiliar)
        If opcao Is Nothing Then
            MessageBox.Show("Selecione uma opção.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ValorSelecionado = opcao.Valor
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
