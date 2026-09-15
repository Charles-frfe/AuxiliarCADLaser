Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Public Module TemaVisual

    Public Function ObterTemaAtual() As String
        If temaAtual = "" Then
            temaAtual = CarregarTemaSalvo()
        End If

        If temaAtual <> "claro" AndAlso temaAtual <> "escuro" Then
            temaAtual = "escuro"
        End If

        Return temaAtual
    End Function

    Public Sub AlternarTema()
        If ObterTemaAtual() = "escuro" Then
            temaAtual = "claro"
        Else
            temaAtual = "escuro"
        End If

        SalvarTema(temaAtual)
    End Sub

    Public Sub DefinirTema(nomeTema As String)
        If String.Equals(nomeTema, "claro", StringComparison.OrdinalIgnoreCase) Then
            temaAtual = "claro"
        Else
            temaAtual = "escuro"
        End If

        SalvarTema(temaAtual)
    End Sub

    Private Function ObterCaminhoTema() As String
        Dim pasta As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CADLaser",
        "Auxiliar"
    )

        Directory.CreateDirectory(pasta)

        Return Path.Combine(pasta, "tema.txt")
    End Function

    Private Function CarregarTemaSalvo() As String
        Try
            Dim arquivo As String = ObterCaminhoTema()

            If File.Exists(arquivo) Then
                Dim valor As String = File.ReadAllText(arquivo, Encoding.UTF8).Trim().ToLowerInvariant()

                If valor = "claro" OrElse valor = "escuro" Then
                    Return valor
                End If
            End If
        Catch
        End Try

        Return "escuro"
    End Function

    Private Sub SalvarTema(nomeTema As String)
        Try
            File.WriteAllText(ObterCaminhoTema(), nomeTema, Encoding.UTF8)
        Catch
        End Try
    End Sub

    Private temaAtual As String = ""

    Private temaCorFundo As Color
    Private temaCorPainel As Color
    Private temaCorTexto As Color
    Private temaCorCaixa As Color
    Private temaCorBotao As Color
    Private temaCorBotaoTexto As Color

    Public ReadOnly CorFundo As Color = Color.FromArgb(24, 24, 27)
    Public ReadOnly CorCard As Color = Color.FromArgb(32, 33, 36)
    Public ReadOnly CorCampo As Color = Color.FromArgb(44, 45, 49)
    Public ReadOnly CorTexto As Color = Color.FromArgb(245, 245, 245)
    Public ReadOnly CorTextoSuave As Color = Color.FromArgb(180, 180, 180)
    Public ReadOnly CorPrincipal As Color = Color.FromArgb(52, 120, 246)
    Public ReadOnly CorNeutro As Color = Color.FromArgb(75, 75, 82)
    Public ReadOnly CorPerigo As Color = Color.FromArgb(180, 70, 70)
    Public ReadOnly CorSucesso As Color = Color.FromArgb(42, 145, 105)
    Public ReadOnly CorSecundario As Color = Color.FromArgb(65, 82, 110)



    Public Sub AplicarTema(controle As Control)
        If controle Is Nothing Then
            Return
        End If

        PrepararCores()

        AplicarTemaControle(controle)

        For Each filho As Control In controle.Controls
            AplicarTema(filho)
        Next
    End Sub

    Private Sub PrepararCores()
        If ObterTemaAtual() = "claro" Then
            temaCorFundo = Color.FromArgb(245, 245, 245)
            temaCorPainel = Color.FromArgb(235, 235, 235)
            temaCorTexto = Color.FromArgb(30, 30, 30)
            temaCorCaixa = Color.White
            temaCorBotao = Color.FromArgb(220, 220, 220)
            temaCorBotaoTexto = Color.FromArgb(20, 20, 20)
        Else
            temaCorFundo = Color.FromArgb(32, 32, 32)
            temaCorPainel = Color.FromArgb(45, 45, 48)
            temaCorTexto = Color.WhiteSmoke
            temaCorCaixa = Color.FromArgb(55, 55, 58)
            temaCorBotao = Color.FromArgb(63, 63, 70)
            temaCorBotaoTexto = Color.WhiteSmoke
        End If
    End Sub

    Private Sub AplicarTemaControle(controle As Control)
        If controle Is Nothing Then
            Return
        End If

        If TypeOf controle Is Form Then
            controle.BackColor = temaCorFundo
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is Panel OrElse TypeOf controle Is GroupBox Then
            controle.BackColor = temaCorPainel
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is TabControl OrElse TypeOf controle Is TabPage Then
            controle.BackColor = temaCorPainel
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is Label OrElse TypeOf controle Is CheckBox OrElse TypeOf controle Is RadioButton Then
            controle.ForeColor = temaCorTexto

            If Not TypeOf controle Is CheckBox AndAlso Not TypeOf controle Is RadioButton Then
                controle.BackColor = Color.Transparent
            End If

        ElseIf TypeOf controle Is TextBox Then
            controle.BackColor = temaCorCaixa
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is RichTextBox Then
            controle.BackColor = temaCorCaixa
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is ComboBox Then
            controle.BackColor = temaCorCaixa
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is ListBox Then
            controle.BackColor = temaCorCaixa
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is CheckedListBox Then
            controle.BackColor = temaCorCaixa
            controle.ForeColor = temaCorTexto

        ElseIf TypeOf controle Is Button Then
            ' Não mexe nos botões antigos do programa.
            ' Só mexe no botão de tema.
            If String.Equals(controle.Name, "btnTema", StringComparison.OrdinalIgnoreCase) Then
                controle.BackColor = temaCorBotao
                controle.ForeColor = temaCorBotaoTexto
                CType(controle, Button).FlatStyle = FlatStyle.Flat
            End If

        Else
            ' Correção de segurança para textos que ficaram brancos no modo claro.
            If ObterTemaAtual() = "claro" Then
                If controle.ForeColor = Color.White OrElse
               controle.ForeColor = Color.WhiteSmoke OrElse
               controle.ForeColor = Color.Gainsboro OrElse
               controle.ForeColor = Color.LightGray Then

                    controle.ForeColor = temaCorTexto
                End If
            Else
                controle.ForeColor = temaCorTexto
            End If
        End If
    End Sub

    Public Sub AplicarControle(controle As Control)
        If TypeOf controle Is Panel Then
            controle.BackColor = CorCard
            controle.ForeColor = CorTexto
        ElseIf TypeOf controle Is Label Then
            controle.BackColor = Color.Transparent
            If controle.Tag IsNot Nothing AndAlso controle.Tag.ToString() = "suave" Then
                controle.ForeColor = CorTextoSuave
            Else
                controle.ForeColor = CorTexto
            End If
        ElseIf TypeOf controle Is TextBox Then
            Dim txt As TextBox = CType(controle, TextBox)
            txt.BackColor = CorCampo
            txt.ForeColor = CorTexto
            txt.BorderStyle = BorderStyle.FixedSingle
        ElseIf TypeOf controle Is ComboBox Then
            Dim cbo As ComboBox = CType(controle, ComboBox)
            cbo.BackColor = CorCampo
            cbo.ForeColor = CorTexto
            cbo.FlatStyle = FlatStyle.Flat
        ElseIf TypeOf controle Is Button Then
            Dim btn As Button = CType(controle, Button)
            Dim estilo As String = "principal"
            If btn.Tag IsNot Nothing Then estilo = btn.Tag.ToString().ToLower()

            Select Case estilo
                Case "neutro"
                    btn.BackColor = CorNeutro
                Case "perigo"
                    btn.BackColor = CorPerigo
                Case "sucesso"
                    btn.BackColor = CorSucesso
                Case "secundario"
                    btn.BackColor = CorSecundario
                Case Else
                    btn.BackColor = CorPrincipal
            End Select
            btn.ForeColor = Color.White
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.Cursor = Cursors.Hand
        ElseIf TypeOf controle Is ListBox Then
            Dim lst As ListBox = CType(controle, ListBox)
            lst.BackColor = CorCampo
            lst.ForeColor = CorTexto
            lst.BorderStyle = BorderStyle.FixedSingle
        End If

        For Each filho As Control In controle.Controls
            AplicarControle(filho)
        Next
    End Sub

End Module
