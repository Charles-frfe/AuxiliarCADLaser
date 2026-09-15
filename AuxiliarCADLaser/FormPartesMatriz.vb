Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class FormPartesMatriz
    Inherits Form

    Private ReadOnly familia As String
    Private ReadOnly partes As String()
    Private ReadOnly numeros As String
    Private ReadOnly texturasDisponiveis As List(Of String)
    Private ReadOnly campos As New Dictionary(Of String, CheckedListBox)(StringComparer.OrdinalIgnoreCase)

    Public ReadOnly Property MatrizesConfiguradas As New List(Of MatrizAuxiliar)()

    Public Sub New(nomeFamilia As String, codigos As String(), numerosEscala As String, texturas As List(Of String))
        familia = nomeFamilia
        partes = codigos
        numeros = numerosEscala

        If texturas Is Nothing Then
            texturasDisponiveis = New List(Of String)()
        Else
            texturasDisponiveis = New List(Of String)(texturas)
        End If

        Me.Text = "Configurar " & nomeFamilia
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ClientSize = New Size(760, 170 + (partes.Length * 115))
        Me.MinimumSize = New Size(720, Math.Min(Me.ClientSize.Height, 650))
        Me.AutoScroll = True

        CriarTela()
        TemaVisual.AplicarTema(Me)
    End Sub

    Private Sub CriarTela()
        Me.Controls.Add(New Label() With {
            .Text = familia & " — texturas por parte",
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .Location = New Point(25, 18),
            .AutoSize = True
        })

        Me.Controls.Add(New Label() With {
            .Text = "Marque as texturas que pertencem a cada parte. Deixe sem marcar quando a parte não possuir textura.",
            .Location = New Point(28, 56),
            .AutoSize = True,
            .Tag = "suave"
        })

        Dim y As Integer = 90

        For Each parte As String In partes
            Dim lbl As New Label() With {
                .Text = parte,
                .Font = New Font("Segoe UI", 11, FontStyle.Bold),
                .Location = New Point(30, y + 9),
                .Size = New Size(90, 25)
            }

            Me.Controls.Add(lbl)

            Dim chk As New CheckedListBox() With {
                .Location = New Point(125, y),
                .Size = New Size(590, 88),
                .CheckOnClick = True,
                .HorizontalScrollbar = True
            }

            For Each textura As String In texturasDisponiveis
                chk.Items.Add(textura, False)
            Next

            campos.Add(parte, chk)
            Me.Controls.Add(chk)

            y += 115
        Next

        If texturasDisponiveis.Count = 0 Then
            Me.Controls.Add(New Label() With {
                .Text = "Nenhuma textura foi detectada na pasta do modelo.",
                .Location = New Point(125, 90),
                .AutoSize = True,
                .Tag = "suave"
            })
        End If

        Dim btnCancelar As New Button() With {
            .Text = "Cancelar",
            .Location = New Point(Me.ClientSize.Width - 285, y + 5),
            .Size = New Size(110, 36),
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Tag = "neutro"
        }

        AddHandler btnCancelar.Click, Sub() Me.Close()
        Me.Controls.Add(btnCancelar)

        Dim btnConfirmar As New Button() With {
            .Text = "Adicionar partes",
            .Location = New Point(Me.ClientSize.Width - 165, y + 5),
            .Size = New Size(140, 36),
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .Tag = "principal"
        }

        AddHandler btnConfirmar.Click, AddressOf Confirmar
        Me.Controls.Add(btnConfirmar)
    End Sub

    Private Sub Confirmar(sender As Object, e As EventArgs)
        MatrizesConfiguradas.Clear()

        For Each parte As String In partes
            MatrizesConfiguradas.Add(New MatrizAuxiliar() With {
                .FamiliaMatriz = familia,
                .TipoMatriz = parte,
                .TextoTexturas = ObterTexturasMarcadas(campos(parte)),
                .TextoNumeros = numeros
            })
        Next

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Function ObterTexturasMarcadas(lista As CheckedListBox) As String
        If lista Is Nothing OrElse lista.CheckedItems.Count = 0 Then Return ""

        Dim selecionadas As New List(Of String)()

        For Each item As Object In lista.CheckedItems
            If item IsNot Nothing AndAlso item.ToString().Trim() <> "" Then
                selecionadas.Add(item.ToString().Trim().ToUpperInvariant())
            End If
        Next

        Return String.Join(Environment.NewLine, selecionadas)
    End Function
End Class