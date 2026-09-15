Imports System.Diagnostics
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Net

Public Module EmailAuxiliar

    Public Sub AbrirRascunho(codigo As String,
                             nomeModelo As String,
                             componente As String,
                             parteMatriz As String,
                             acabamento As String,
                             observacao As String)

        Dim modeloFormatado As String = FormatarNomeModelo(codigo, nomeModelo)
        Dim assunto As String = "CAD/LASER - ACABAMENTO - " & modeloFormatado
        Dim corpoHtml As String = MontarCorpoHtml(
            modeloFormatado, componente, parteMatriz, acabamento, observacao)

        If AbrirNoOutlookDesktop(assunto, corpoHtml) Then Return

        ' O Outlook Web recebe somente texto pelo deeplink. Mantém o fluxo anterior
        ' como alternativa quando o Outlook clássico/COM não estiver instalado.
        Dim corpoTexto As String = MontarCorpoTexto(
            modeloFormatado, componente, parteMatriz, acabamento, observacao)
        Dim endereco As String =
            "https://outlook.office.com/mail/deeplink/compose" &
            "?subject=" & Uri.EscapeDataString(assunto) &
            "&body=" & Uri.EscapeDataString(corpoTexto)

        Process.Start(New ProcessStartInfo() With {
            .FileName = endereco,
            .UseShellExecute = True
        })
    End Sub

    Private Function AbrirNoOutlookDesktop(assunto As String, corpoHtml As String) As Boolean
        Try
            Dim tipoOutlook As Type = Type.GetTypeFromProgID("Outlook.Application")
            If tipoOutlook Is Nothing Then Return False

            Dim outlook As Object = Activator.CreateInstance(tipoOutlook)
            Dim mensagem As Object = outlook.CreateItem(0)
            mensagem.Subject = assunto
            mensagem.Display(False)
            mensagem.HTMLBody = corpoHtml & mensagem.HTMLBody
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function MontarCorpoHtml(modelo As String,
                                     componente As String,
                                     parteMatriz As String,
                                     acabamento As String,
                                     observacao As String) As String
        Dim html As New StringBuilder()
        html.Append("<div style='font-family:Segoe UI,Arial,sans-serif;font-size:11pt;color:#202124'>")
        html.Append("<p>").Append(Codificar(ObterSaudacao())).Append(".</p>")
        html.Append("<p>Seguem as informações de acabamento:</p>")
        html.Append("<table style='border-collapse:collapse'>")
        AdicionarLinhaHtml(html, "MODELO", modelo)
        AdicionarLinhaHtml(html, "COMPONENTE", Maiusculo(componente))
        AdicionarLinhaHtml(html, "PARTE DA MATRIZ", Maiusculo(parteMatriz))
        AdicionarLinhaHtml(html, "ACABAMENTO", Maiusculo(acabamento))
        html.Append("</table>")

        If Not String.IsNullOrWhiteSpace(observacao) Then
            html.Append("<p style='margin-top:18px;color:#c00000;font-weight:700'>")
            html.Append("OBSERVAÇÃO: ").Append(CodificarComQuebras(Maiusculo(observacao)))
            html.Append("</p>")
        End If

        html.Append("<p>Qualquer dúvida, fico à disposição.</p></div><br>")
        Return html.ToString()
    End Function

    Private Sub AdicionarLinhaHtml(html As StringBuilder, rotulo As String, valor As String)
        html.Append("<tr><td style='padding:3px 14px 3px 0;font-weight:700'>")
        html.Append(Codificar(rotulo)).Append(":</td><td style='padding:3px 0'>")
        html.Append(Codificar(valor)).Append("</td></tr>")
    End Sub

    Private Function MontarCorpoTexto(modelo As String,
                                      componente As String,
                                      parteMatriz As String,
                                      acabamento As String,
                                      observacao As String) As String
        Dim texto As New StringBuilder()
        texto.AppendLine(ObterSaudacao() & ".")
        texto.AppendLine()
        texto.AppendLine("Seguem as informações de acabamento:")
        texto.AppendLine()
        texto.AppendLine("MODELO: " & modelo)
        texto.AppendLine("COMPONENTE: " & Maiusculo(componente))
        texto.AppendLine("PARTE DA MATRIZ: " & Maiusculo(parteMatriz))
        texto.AppendLine("ACABAMENTO: " & Maiusculo(acabamento))
        If Not String.IsNullOrWhiteSpace(observacao) Then
            texto.AppendLine()
            texto.AppendLine("OBSERVAÇÃO: " & Maiusculo(observacao))
        End If
        texto.AppendLine()
        texto.AppendLine("Qualquer dúvida, fico à disposição.")
        Return texto.ToString()
    End Function

    Private Function FormatarNomeModelo(codigo As String, nomePasta As String) As String
        Dim nome As String = If(nomePasta, "").Trim()
        nome = Regex.Replace(nome, "[_-]+", " ")
        nome = Regex.Replace(nome, "\s+", " ").Trim().ToUpperInvariant()

        Dim codigoLimpo As String = If(codigo, "").Trim()
        If codigoLimpo <> "" AndAlso Not nome.StartsWith(codigoLimpo, StringComparison.OrdinalIgnoreCase) Then
            nome = (codigoLimpo & " " & nome).Trim()
        End If
        Return nome
    End Function

    Private Function Maiusculo(valor As String) As String
        Return If(valor, "").Trim().ToUpperInvariant()
    End Function

    Private Function Codificar(valor As String) As String
        Return WebUtility.HtmlEncode(If(valor, ""))
    End Function

    Private Function CodificarComQuebras(valor As String) As String
        Return Codificar(valor).Replace(vbCrLf, "<br>").Replace(vbCr, "<br>").Replace(vbLf, "<br>")
    End Function

    Private Function ObterSaudacao() As String
        Dim hora As Integer = DateTime.Now.Hour
        If hora < 12 Then Return "Bom dia"
        If hora < 18 Then Return "Boa tarde"
        Return "Boa noite"
    End Function

End Module
