Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class TexturaAuxiliar
    Public Property NomeOriginal As String = ""
    Public Property NomePasta As String = ""
    Public Property TemBaseAcabamento As Boolean = False
End Class

Public Class MatrizAuxiliar
    Public Property FamiliaMatriz As String = ""
    Public Property TipoMatriz As String = ""
    Public Property TextoTexturas As String = ""
    Public Property TextoNumeros As String = ""
    Public Property Maquina As String = ""
    Public Property Material As String = ""
    Public Property Eixos As String = ""
    Public Property Projetista As String = ""
    Public Property Acabamento As String = ""
    Public Property ObservacoesAcabamento As String = ""

End Class

Public Class ResultadoGeracao
    Public Property PastaPrincipal As String = ""
    Public Property Mensagens As New List(Of String)
    Public Property PastasCriadas As New List(Of String)
    Public Property PastasJaExistiam As New List(Of String)
End Class

Public Class ResultadoCopiaCharmilles
    Public Property ArquivosCopiados As Integer = 0
    Public Property ArquivosIgnorados As Integer = 0
    Public Property ItensJaExistentes As New List(Of String)
End Class


Public Class ResultadoFinalizacao
    Public Property PastaDestino As String = ""
    Public Property QuantidadeArquivos As Integer = 0
    Public Property CopiaEntreVolumes As Boolean = False
End Class

Public Class ConfiguracaoAuxiliar
    Public Property PastaProjetos As String = ""
    Public Property PastaCN As String = ""
    Public Property PastaFinalizados As String = ""
    Public Property PastaCharmilles As String = ""
    Public Property CaminhoRhino As String = ""
End Class

Public Module ModuloFuncoes

    Public Const PASTA_CN_PADRAO As String = "C:\programa_cn"

    Public Function ObterPastaCN() As String
        Dim configurada As String = CarregarConfiguracao().PastaCN
        If String.IsNullOrWhiteSpace(configurada) Then
            Return PASTA_CN_PADRAO
        End If
        Return configurada.Trim()
    End Function

    Public Function DetectarTipoServicoPeloCaminho(caminho As String) As String
        If String.IsNullOrWhiteSpace(caminho) Then
            Return ""
        End If

        Dim caminhoComparacao As String = RemoverAcentos(caminho).ToLowerInvariant()

        If caminhoComparacao.Contains("piloto") OrElse caminhoComparacao.Contains("pilotos") Then
            Return "piloto"
        ElseIf caminhoComparacao.Contains("colecao") OrElse caminhoComparacao.Contains("escala") Then
            Return "escala"
        ElseIf caminhoComparacao.Contains("reforma") Then
            Return "reforma"
        ElseIf caminhoComparacao.Contains("teste") Then
            Return "teste"
        ElseIf caminhoComparacao.Contains("limpeza") Then
            Return "limpeza"
        ElseIf caminhoComparacao.Contains("retoque") Then
            Return "retoque"
        End If

        Return ""
    End Function

    Public Function DetectarTipoTrabalhoPeloCaminho(caminho As String) As String
        If String.IsNullOrWhiteSpace(caminho) Then Return "novo"
        Dim comparacao As String = RemoverAcentos(caminho).ToLowerInvariant()
        If comparacao.Contains("reforma") Then Return "reforma"
        If comparacao.Contains("retoque") Then Return "retoque"
        If comparacao.Contains("limpeza") Then Return "limpeza"
        Return "novo"
    End Function

    Public Function SelecionarPasta(titulo As String, pastaInicial As String) As String
        Try
            Return SeletorPastaModerno.Selecionar(titulo, pastaInicial)
        Catch
            ' Nunca volta para o FolderBrowserDialog antigo. Se a seleção nativa
            ' de pastas falhar, usa o Explorador moderno permitindo ver arquivos.
            Return SelecionarPastaPeloExplorador(titulo, pastaInicial)
        End Try
    End Function

    Private Function SelecionarPastaPeloExplorador(titulo As String,
                                                    pastaInicial As String) As String
        Using dialogo As New OpenFileDialog()
            dialogo.Title = titulo & " - entre na pasta desejada e clique em Abrir"
            dialogo.Filter = "Todos os arquivos (*.*)|*.*"
            dialogo.AutoUpgradeEnabled = True
            dialogo.CheckFileExists = False
            dialogo.CheckPathExists = True
            dialogo.ValidateNames = False
            dialogo.DereferenceLinks = True
            dialogo.RestoreDirectory = True
            dialogo.FileName = "Selecionar esta pasta"

            If Not String.IsNullOrWhiteSpace(pastaInicial) AndAlso Directory.Exists(pastaInicial) Then
                dialogo.InitialDirectory = pastaInicial
            End If

            If dialogo.ShowDialog() <> DialogResult.OK Then
                Return ""
            End If

            If Directory.Exists(dialogo.FileName) Then
                Return dialogo.FileName
            End If

            Dim pastaEscolhida As String = Path.GetDirectoryName(dialogo.FileName)

            If Not String.IsNullOrWhiteSpace(pastaEscolhida) AndAlso Directory.Exists(pastaEscolhida) Then
                Return pastaEscolhida
            End If

            Return ""
        End Using
    End Function

    Public Function LimparNomePasta(valor As String) As String
        If String.IsNullOrWhiteSpace(valor) Then
            Return ""
        End If

        Dim semAcentos As String = RemoverAcentos(valor.Trim())
        Dim sb As New StringBuilder()

        For Each ch As Char In semAcentos
            If Char.IsLetterOrDigit(ch) Then
                sb.Append(Char.ToLower(ch))
            ElseIf ch = "_"c OrElse ch = "-"c Then
                sb.Append(ch)
            ElseIf Char.IsWhiteSpace(ch) Then
                sb.Append("_")
            Else
                sb.Append("_")
            End If
        Next

        Dim resultado As String = sb.ToString()

        While resultado.Contains("__")
            resultado = resultado.Replace("__", "_")
        End While

        resultado = resultado.Trim("_"c)

        Return resultado
    End Function

    Public Function ObterNomePastaMatriz(tipoMatriz As String) As String
        Dim limpo As String = LimparNomePasta(tipoMatriz)
        Select Case limpo
            Case "lfe", "lfd", "lme", "lmd", "fe", "fd"
                Return limpo.ToUpperInvariant()
            Case "fundod", "fundo_d", "fundo-d"
                Return "FundoD"
            Case "fundoe", "fundo_e", "fundo-e"
                Return "FundoE"
            Case Else
                Return limpo
        End Select
    End Function

    Public Function RemoverAcentos(texto As String) As String
        If texto Is Nothing Then
            Return ""
        End If

        Dim normalizado As String = texto.Normalize(NormalizationForm.FormD)
        Dim sb As New StringBuilder()

        For Each ch As Char In normalizado
            Dim categoria As UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(ch)

            If categoria <> UnicodeCategory.NonSpacingMark Then
                sb.Append(ch)
            End If
        Next

        Return sb.ToString().Normalize(NormalizationForm.FormC)
    End Function

    Public Function ObterPastaTipoServico(tipoServico As String) As String
        Dim tipo As String = ""

        If tipoServico IsNot Nothing Then
            tipo = tipoServico.Trim().ToLower()
        End If

        Select Case tipo
            Case "piloto"
                Return "pilotos"
            Case "escala"
                Return "colecao"
            Case "reforma"
                Return "reformas"
            Case "teste"
                Return "testes"
            Case "limpeza"
                Return "limpezas"
            Case "retoque"
                Return "retoques"
            Case Else
                Return "outros"
        End Select
    End Function

    Public Function ObterNomePastaTrabalho(tipoTrabalho As String) As String
        Dim tipo As String = RemoverAcentos(If(tipoTrabalho, "")).Trim().ToLowerInvariant()
        Select Case tipo
            Case "reforma"
                Return "reforma"
            Case "retoque"
                Return "retoque"
            Case "limpeza"
                Return "limpeza"
            Case Else
                Return ""
        End Select
    End Function

    Public Function ObterPastaModeloCN(nomePastaModelo As String,
                                       tipoServico As String,
                                       tipoTrabalho As String) As String
        Dim pastaModelo As String = Path.Combine(
            ObterPastaCN(),
            ObterPastaTipoServico(tipoServico),
            nomePastaModelo
        )

        If Not String.Equals(If(tipoServico, "").Trim(), "teste", StringComparison.OrdinalIgnoreCase) Then
            Dim trabalho As String = ObterNomePastaTrabalho(tipoTrabalho)
            If trabalho <> "" Then pastaModelo = Path.Combine(pastaModelo, trabalho)
        End If

        Return pastaModelo
    End Function

    Public Function ObterNomePastaModelo(caminhoPastaModelo As String) As String
        If String.IsNullOrWhiteSpace(caminhoPastaModelo) Then
            Return ""
        End If

        If Not Directory.Exists(caminhoPastaModelo) Then
            Return ""
        End If

        Dim info As New DirectoryInfo(caminhoPastaModelo)

        ' Regra atual:
        ' O nome da pasta no CN será igual ao nome da pasta selecionada.
        Return info.Name
    End Function

    Public Function QuebrarNumerosPorVirgula(texto As String) As List(Of String)
        Dim lista As New List(Of String)

        If String.IsNullOrWhiteSpace(texto) Then
            Return lista
        End If

        If texto.Contains(vbCr) OrElse texto.Contains(vbLf) OrElse texto.Contains(";") Then
            Throw New Exception("Separe os números da escala somente por vírgula. Exemplo: 29,30,31,32")
        End If

        ' Regra atual:
        ' Separar somente por vírgula.
        ' Traço NÃO gera intervalo. Exemplo: 29-35 vira uma pasta chamada 29-35.
        Dim partes() As String = texto.Split(","c)

        For Each parte As String In partes
            Dim numeroOriginal As String = parte.Trim()

            If numeroOriginal <> "" Then
                Dim numeroLimpo As String = LimparNomePasta(numeroOriginal)

                If numeroLimpo <> "" AndAlso Not lista.Contains(numeroLimpo) Then
                    lista.Add(numeroLimpo)
                End If
            End If
        Next

        Return lista
    End Function

    Public Function LerTexturas(texto As String) As List(Of TexturaAuxiliar)
        Dim lista As New List(Of TexturaAuxiliar)

        If String.IsNullOrWhiteSpace(texto) Then
            Return lista
        End If

        Dim linhas() As String = texto.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(vbLf)

        For Each linhaOriginal As String In linhas
            Dim linha As String = linhaOriginal.Trim()

            If linha <> "" Then
                Dim temBA As Boolean = DetectarBaseAcabamento(linha)
                Dim nomeSemMarcador As String = RemoverMarcadoresBaseAcabamento(linha)
                Dim nomePasta As String = LimparNomePasta(nomeSemMarcador)

                ' Uma linha pode ser apenas "B/A", "BA", "Base" ou "Acabamento".
                ' Nesse caso ela ainda representa uma textura com as duas etapas.
                If nomePasta = "" AndAlso temBA Then
                    nomePasta = LimparNomePasta(linha)
                End If

                If nomePasta <> "" Then
                    Dim textura As New TexturaAuxiliar()
                    textura.NomeOriginal = linha
                    textura.NomePasta = nomePasta
                    textura.TemBaseAcabamento = temBA

                    lista.Add(textura)
                End If
            End If
        Next

        Return lista
    End Function

    Private Function DetectarBaseAcabamento(texto As String) As Boolean
        If String.IsNullOrWhiteSpace(texto) Then
            Return False
        End If

        Return Regex.IsMatch(
            texto,
            "(?i)(\bB\s*/\s*A\b|\(BA\)|\[BA\]|\bBA\b|\bBASE\b|\bACABAMENTO\b)"
        )
    End Function

    Private Function RemoverMarcadoresBaseAcabamento(texto As String) As String
        If texto Is Nothing Then
            Return ""
        End If

        Dim resultado As String = texto

        resultado = Regex.Replace(resultado, "(?i)\bB\s*/\s*A\b", "")
        resultado = Regex.Replace(resultado, "(?i)\(BA\)", "")
        resultado = Regex.Replace(resultado, "(?i)\[BA\]", "")
        resultado = Regex.Replace(resultado, "(?i)\bBA\b", "")
        resultado = Regex.Replace(resultado, "(?i)\bBASE\s*/\s*ACABAMENTO\b", "")
        resultado = Regex.Replace(resultado, "(?i)\bBASE\s+E\s+ACABAMENTO\b", "")
        resultado = Regex.Replace(resultado, "(?i)\bBASE\b", "")
        resultado = Regex.Replace(resultado, "(?i)\bACABAMENTO\b", "")

        Return resultado.Trim()
    End Function

    Private Sub GarantirPastaCN(caminho As String, resultado As ResultadoGeracao)
        If String.IsNullOrWhiteSpace(caminho) Then
            Return
        End If

        Dim caminhoCompleto As String = Path.GetFullPath(caminho).
        TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        Dim jaRegistradaComoCriada As Boolean =
        resultado.PastasCriadas.Exists(Function(p) String.Equals(p, caminhoCompleto, StringComparison.OrdinalIgnoreCase))

        Dim jaRegistradaComoExistente As Boolean =
        resultado.PastasJaExistiam.Exists(Function(p) String.Equals(p, caminhoCompleto, StringComparison.OrdinalIgnoreCase))

        If Directory.Exists(caminhoCompleto) Then
            If Not jaRegistradaComoCriada AndAlso Not jaRegistradaComoExistente Then
                resultado.PastasJaExistiam.Add(caminhoCompleto)
            End If
        Else
            Directory.CreateDirectory(caminhoCompleto)

            If Not jaRegistradaComoCriada Then
                resultado.PastasCriadas.Add(caminhoCompleto)
            End If
        End If
    End Sub

    Public Function GerarPastasCN(caminhoPastaModelo As String,
                                  tipoServico As String,
                                  tipoTrabalho As String,
                                  componente As String,
                                  matrizes As List(Of MatrizAuxiliar)) As ResultadoGeracao

        If String.IsNullOrWhiteSpace(caminhoPastaModelo) Then
            Throw New Exception("Selecione a pasta do modelo.")
        End If

        If Not Directory.Exists(caminhoPastaModelo) Then
            Throw New Exception("A pasta do modelo não foi encontrada: " & caminhoPastaModelo)
        End If

        Dim nomePastaModelo As String = ObterNomePastaModelo(caminhoPastaModelo)

        If nomePastaModelo = "" Then
            Throw New Exception("Não foi possível identificar o nome da pasta do modelo.")
        End If

        Return GerarPastasCNPorNomeModelo(nomePastaModelo, tipoServico, tipoTrabalho, componente, matrizes)
    End Function

    Public Function GerarPastasCNPorNomeModelo(nomePastaModelo As String,
                                               tipoServico As String,
                                               tipoTrabalho As String,
                                               componente As String,
                                               matrizes As List(Of MatrizAuxiliar)) As ResultadoGeracao

        Dim resultado As New ResultadoGeracao()

        If String.IsNullOrWhiteSpace(nomePastaModelo) Then
            Throw New Exception("Informe ou selecione a pasta do modelo.")
        End If

        If String.IsNullOrWhiteSpace(tipoServico) Then
            Throw New Exception("Informe o tipo de serviço.")
        End If

        If String.IsNullOrWhiteSpace(componente) Then
            Throw New Exception("Informe o componente.")
        End If

        If matrizes Is Nothing OrElse matrizes.Count = 0 Then
            Throw New Exception("Adicione pelo menos uma matriz.")
        End If

        Dim componenteLimpo As String = LimparNomePasta(componente)
        Dim ehEscala As Boolean = tipoServico.Trim().ToLower() = "escala"

        Dim tiposMatrizesValidos As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each matrizContagem As MatrizAuxiliar In matrizes
            If matrizContagem Is Nothing Then Continue For

            Dim tipoContagem As String =
                ObterNomePastaMatriz(matrizContagem.TipoMatriz)

            If tipoContagem <> "" Then
                tiposMatrizesValidos.Add(tipoContagem)
            End If
        Next

        Dim usarLavPenDiretoNoNumero As Boolean = ehEscala AndAlso tiposMatrizesValidos.Count = 1

        If componenteLimpo = "" Then
            Throw New Exception("O nome do componente ficou inválido depois da limpeza.")
        End If

        Dim pastaModeloCN As String = ObterPastaModeloCN(nomePastaModelo, tipoServico, tipoTrabalho)
        GarantirPastaCN(pastaModeloCN, resultado)

        resultado.PastaPrincipal = pastaModeloCN
        resultado.Mensagens.Add("Pasta Programa CN: " & ObterPastaCN())
        resultado.Mensagens.Add("Nome da pasta do modelo: " & nomePastaModelo)
        If ObterNomePastaTrabalho(tipoTrabalho) <> "" Then
            resultado.Mensagens.Add("Tipo do trabalho: " & tipoTrabalho.Trim().ToUpperInvariant())
        End If
        resultado.Mensagens.Add("Pasta criada no CN: " & pastaModeloCN)
        resultado.Mensagens.Add("")

        For Each matriz As MatrizAuxiliar In matrizes
            If matriz Is Nothing Then
                Continue For
            End If

            Dim tipoMatrizLimpo As String = ObterNomePastaMatriz(matriz.TipoMatriz)

            If tipoMatrizLimpo = "" Then
                Throw New Exception("Existe uma matriz sem tipo informado.")
            End If

            Dim texturas As List(Of TexturaAuxiliar) = LerTexturas(matriz.TextoTexturas)

            If ehEscala Then
                Dim numeros As List(Of String) = QuebrarNumerosPorVirgula(matriz.TextoNumeros)

                If numeros.Count = 0 Then
                    Throw New Exception("Informe os números da escala para a matriz: " & matriz.TipoMatriz)
                End If

                For Each numero As String In numeros
                    Dim pastaMatriz As String

                    If usarLavPenDiretoNoNumero Then
                        pastaMatriz = Path.Combine(
                            pastaModeloCN,
componenteLimpo,
numero)
                    Else
                        pastaMatriz = Path.Combine(
                            pastaModeloCN,
componenteLimpo,
numero,
tipoMatrizLimpo)
                    End If

                    CriarEstruturaLavPen(pastaMatriz, texturas, resultado)

                    resultado.Mensagens.Add("Escala criada:")
                    resultado.Mensagens.Add("  Número: " & numero)
                    resultado.Mensagens.Add("  Matriz: " & tipoMatrizLimpo)

                    If usarLavPenDiretoNoNumero Then
                        resultado.Mensagens.Add("  Estrutura: LAV/PEN direto dentro do número")
                    End If

                    resultado.Mensagens.Add("  Pasta:" & pastaMatriz)
                    resultado.Mensagens.Add("")
                Next
            Else
                Dim pastaMatriz As String = Path.Combine(
                    pastaModeloCN,
                    componenteLimpo,
                    tipoMatrizLimpo
                )

                CriarEstruturaLavPen(pastaMatriz, texturas, resultado)

                resultado.Mensagens.Add("Matriz criada:")
                resultado.Mensagens.Add("  Matriz: " & tipoMatrizLimpo)
                resultado.Mensagens.Add("  Pasta: " & pastaMatriz)
                resultado.Mensagens.Add("")
            End If
        Next

        resultado.Mensagens.Add("RESUMO DA GERAÇÃO CN")
        resultado.Mensagens.Add("Pastas criadas:" &
            resultado.PastasCriadas.Count.ToString())
        resultado.Mensagens.Add("Pastas que já existiam:" &
            resultado.PastasJaExistiam.Count.ToString())
        resultado.Mensagens.Add("")

        If resultado.PastasJaExistiam.Count > 0 Then
            resultado.Mensagens.Add("Pastas que já exixtiam:")
            For Each pastaExistente As String In
                    resultado.PastasJaExistiam
                resultado.Mensagens.Add(" " & pastaExistente)
            Next
            resultado.Mensagens.Add("")
        End If
        resultado.Mensagens.Add("Concluído.")
        Return resultado
    End Function

    Private Sub CriarEstruturaLavPen(pastaMatriz As String,
                                 texturas As List(Of TexturaAuxiliar),
                                 resultado As ResultadoGeracao)

        Dim pastaLav As String = Path.Combine(pastaMatriz, "lav")
        Dim pastaPen As String = Path.Combine(pastaMatriz, "pen")

        GarantirPastaCN(pastaLav, resultado)
        GarantirPastaCN(pastaPen, resultado)

        If texturas Is Nothing OrElse texturas.Count = 0 Then
            Return
        End If

        If texturas.Count = 1 Then
            Dim texturaUnica As TexturaAuxiliar = texturas(0)

            If texturaUnica.TemBaseAcabamento Then
                GarantirPastaCN(Path.Combine(pastaLav, "base"), resultado)
                GarantirPastaCN(Path.Combine(pastaLav, "acabamento"), resultado)

                GarantirPastaCN(Path.Combine(pastaPen, "base"), resultado)
                GarantirPastaCN(Path.Combine(pastaPen, "acabamento"), resultado)

                resultado.Mensagens.Add("  Textura única com B/A: base e acabamento criados direto em LAV/PEN.")
            Else
                resultado.Mensagens.Add("  Textura única sem B/A: mantido somente LAV/PEN.")
            End If

            Return
        End If

        For Each textura As TexturaAuxiliar In texturas
            Dim pastaTexturaLav As String = Path.Combine(pastaLav, textura.NomePasta)
            Dim pastaTexturaPen As String = Path.Combine(pastaPen, textura.NomePasta)

            GarantirPastaCN(pastaTexturaLav, resultado)
            GarantirPastaCN(pastaTexturaPen, resultado)

            If textura.TemBaseAcabamento Then
                GarantirPastaCN(Path.Combine(pastaTexturaLav, "base"), resultado)
                GarantirPastaCN(Path.Combine(pastaTexturaLav, "acabamento"), resultado)

                GarantirPastaCN(Path.Combine(pastaTexturaPen, "base"), resultado)
                GarantirPastaCN(Path.Combine(pastaTexturaPen, "acabamento"), resultado)
            End If
        Next
    End Sub

    Public Function AnalisarLavPen(pastaRaiz As String) As List(Of String)
        Dim mensagens As New List(Of String)

        If String.IsNullOrWhiteSpace(pastaRaiz) Then
            mensagens.Add("Nenhuma pasta informada para análise.")
            Return mensagens
        End If

        If Not Directory.Exists(pastaRaiz) Then
            mensagens.Add("Pasta não encontrada: " & pastaRaiz)
            Return mensagens
        End If

        mensagens.Add("Analisando: " & pastaRaiz)
        mensagens.Add("")

        Dim encontrouAlgumaEstrutura As Boolean = False

        Try
            Dim pastasParaAnalisar As New List(Of String)
            pastasParaAnalisar.Add(pastaRaiz)

            For Each pasta As String In Directory.GetDirectories(pastaRaiz, "*", SearchOption.AllDirectories)
                pastasParaAnalisar.Add(pasta)
            Next

            For Each pasta As String In pastasParaAnalisar
                Dim lav As String = Path.Combine(pasta, "lav")
                Dim pen As String = Path.Combine(pasta, "pen")

                If Directory.Exists(lav) OrElse Directory.Exists(pen) Then
                    encontrouAlgumaEstrutura = True

                    mensagens.Add("Pasta:")
                    mensagens.Add(pasta)

                    If Directory.Exists(lav) Then
                        mensagens.Add("  LAV: OK")

                        If ExisteSimulator(lav) Then
                            mensagens.Add("  Simulator em LAV: OK")
                        Else
                            mensagens.Add("  Simulator em LAV: NÃO ENCONTRADO")
                        End If
                    Else
                        mensagens.Add("  LAV: NÃO ENCONTRADO")
                    End If

                    If Directory.Exists(pen) Then
                        mensagens.Add("  PEN: OK")

                        If ExisteSimulator(pen) Then
                            mensagens.Add("  Simulator em PEN: OK")
                        Else
                            mensagens.Add("  Simulator em PEN: NÃO ENCONTRADO")
                        End If
                    Else
                        mensagens.Add("  PEN: NÃO ENCONTRADO")
                    End If

                    mensagens.Add("")
                End If
            Next

            If Not encontrouAlgumaEstrutura Then
                mensagens.Add("Nenhuma estrutura com LAV/PEN foi encontrada.")
            End If

        Catch ex As Exception
            mensagens.Add("Erro durante análise: " & ex.Message)
        End Try

        Return mensagens
    End Function

    Public Function ValidarSimulatorParaEnvio(pastaRaiz As String) As List(Of String)
        Dim problemas As New List(Of String)()

        If String.IsNullOrWhiteSpace(pastaRaiz) OrElse Not Directory.Exists(pastaRaiz) Then
            problemas.Add("A pasta CN de origem não foi encontrada: " & If(pastaRaiz, ""))
            Return problemas
        End If

        Dim estruturas As New List(Of String)()
        estruturas.Add(pastaRaiz)
        estruturas.AddRange(Directory.GetDirectories(pastaRaiz, "*", SearchOption.AllDirectories))

        Dim encontrou As Boolean = False
        For Each pasta As String In estruturas
            Dim lav As String = Path.Combine(pasta, "lav")
            Dim pen As String = Path.Combine(pasta, "pen")
            If Not Directory.Exists(lav) AndAlso Not Directory.Exists(pen) Then Continue For

            encontrou = True
            Dim identificacao As String = pasta.Substring(pastaRaiz.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If identificacao = "" Then identificacao = New DirectoryInfo(pasta).Name

            If Not Directory.Exists(lav) Then
                problemas.Add(identificacao & ": pasta LAV não encontrada.")
            ElseIf Not ExisteSimulator(lav) Then
                problemas.Add(identificacao & ": arquivo Simulator não encontrado em LAV.")
            End If

            If Not Directory.Exists(pen) Then
                problemas.Add(identificacao & ": pasta PEN não encontrada.")
            ElseIf Not ExisteSimulator(pen) Then
                problemas.Add(identificacao & ": arquivo Simulator não encontrado em PEN.")
            End If
        Next

        If Not encontrou Then problemas.Add("Nenhuma estrutura LAV/PEN foi encontrada na pasta CN.")
        Return problemas
    End Function

    Public Function ObterDestinoCharmilles(pastaOrigem As String) As String
        Dim config As ConfiguracaoAuxiliar = CarregarConfiguracao()
        If String.IsNullOrWhiteSpace(config.PastaCharmilles) Then
            Throw New Exception("Configure a pasta Laser Charmilles na engrenagem.")
        End If
        If Not Directory.Exists(config.PastaCharmilles) Then
            Throw New Exception("A pasta Laser Charmilles está inacessível: " & config.PastaCharmilles)
        End If

        Dim raizCN As String = Path.GetFullPath(ObterPastaCN()).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim origem As String = Path.GetFullPath(pastaOrigem).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim prefixo As String = raizCN & Path.DirectorySeparatorChar
        If Not origem.StartsWith(prefixo, StringComparison.OrdinalIgnoreCase) Then
            Throw New Exception("A pasta de origem não está dentro do Programa CN configurado.")
        End If

        Dim relativo As String = origem.Substring(prefixo.Length)
        Dim destino As String = Path.Combine(config.PastaCharmilles, relativo)
        Dim destinoCompleto As String = Path.GetFullPath(destino).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        If String.Equals(destinoCompleto, origem, StringComparison.OrdinalIgnoreCase) OrElse
           destinoCompleto.StartsWith(origem & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then
            Throw New Exception("A pasta Laser Charmilles não pode ficar dentro da pasta de origem.")
        End If
        Return destinoCompleto
    End Function

    Public Function CopiarPastaRecursivamente(origem As String,
                                               destino As String,
                                               sobrescrever As Boolean) As Integer
        If Not Directory.Exists(origem) Then Throw New DirectoryNotFoundException(origem)
        Directory.CreateDirectory(destino)

        For Each pasta As String In Directory.GetDirectories(origem, "*", SearchOption.AllDirectories)
            Dim relativo As String = pasta.Substring(origem.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Directory.CreateDirectory(Path.Combine(destino, relativo))
        Next

        Dim quantidade As Integer = 0
        For Each arquivo As String In Directory.GetFiles(origem, "*", SearchOption.AllDirectories)
            Dim relativo As String = arquivo.Substring(origem.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim arquivoDestino As String = Path.Combine(destino, relativo)
            Directory.CreateDirectory(Path.GetDirectoryName(arquivoDestino))
            File.Copy(arquivo, arquivoDestino, sobrescrever)
            quantidade += 1
        Next
        Return quantidade
    End Function

    Public Function ListarItensJaExistentesNoDestino(origem As String, destino As String) As List(Of String)
        Dim existentes As New List(Of String)()

        If String.IsNullOrWhiteSpace(origem) OrElse Not Directory.Exists(origem) Then
            Return existentes
        End If

        If String.IsNullOrWhiteSpace(destino) Then
            Return existentes
        End If

        If Not Directory.Exists(destino) Then
            Return existentes
        End If

        For Each pasta As String In Directory.GetDirectories(origem, "*", SearchOption.AllDirectories)
            Dim relativo As String = pasta.Substring(origem.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim pastaDestino As String = Path.Combine(destino, relativo)

            If Directory.Exists(pastaDestino) Then
                existentes.Add("Pasta: " & relativo)
            End If
        Next

        For Each arquivo As String In Directory.GetFiles(origem, "*", SearchOption.AllDirectories)
            Dim relativo As String = arquivo.Substring(origem.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim arquivoDestino As String = Path.Combine(destino, relativo)

            If File.Exists(arquivoDestino) Then
                existentes.Add("Arquivo: " & relativo)
            End If
        Next

        existentes.Sort()

        Return existentes
    End Function

    Public Function ListarArquivosNovosParaDestino(origem As String,
                                               destino As String) As List(Of String)

        Dim novos As New List(Of String)()

        If String.IsNullOrWhiteSpace(origem) OrElse Not Directory.Exists(origem) Then
            Return novos
        End If

        For Each arquivo As String In Directory.GetFiles(origem, "*", SearchOption.AllDirectories)

            Dim relativo As String = arquivo.Substring(origem.Length).
            TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

            Dim arquivoDestino As String = Path.Combine(destino, relativo)

            If Not File.Exists(arquivoDestino) Then
                novos.Add(relativo)
            End If
        Next

        novos.Sort()

        Return novos
    End Function

    Public Function CopiarPastaParaCharmillesSemSobrescrever(origem As String, destino As String) As ResultadoCopiaCharmilles
        If Not Directory.Exists(origem) Then
            Throw New DirectoryNotFoundException(origem)
        End If

        Dim resultado As New ResultadoCopiaCharmilles()

        Directory.CreateDirectory(destino)

        For Each pasta As String In Directory.GetDirectories(origem, "*", SearchOption.AllDirectories)
            Dim relativo As String = pasta.Substring(origem.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim pastaDestino As String = Path.Combine(destino, relativo)

            If Directory.Exists(pastaDestino) Then
                resultado.ItensJaExistentes.Add("Pasta: " & relativo)
            Else
                Directory.CreateDirectory(pastaDestino)
            End If
        Next

        For Each arquivo As String In Directory.GetFiles(origem, "*", SearchOption.AllDirectories)
            Dim relativo As String = arquivo.Substring(origem.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim arquivoDestino As String = Path.Combine(destino, relativo)

            Directory.CreateDirectory(Path.GetDirectoryName(arquivoDestino))

            If File.Exists(arquivoDestino) Then
                resultado.ArquivosIgnorados += 1
                resultado.ItensJaExistentes.Add("Arquivo: " & relativo)
            Else
                File.Copy(arquivo, arquivoDestino, False)
                resultado.ArquivosCopiados += 1
            End If
        Next

        resultado.ItensJaExistentes.Sort()

        Return resultado
    End Function

    Public Function LocalizarArquivosBloqueados(pastaRaiz As String) As List(Of String)
        Dim bloqueados As New List(Of String)()
        If String.IsNullOrWhiteSpace(pastaRaiz) OrElse Not Directory.Exists(pastaRaiz) Then
            bloqueados.Add(If(pastaRaiz, "") & " — pasta não encontrada")
            Return bloqueados
        End If

        Try
            For Each arquivo As String In Directory.GetFiles(pastaRaiz, "*", SearchOption.AllDirectories)
                Try
                    Using fluxo As New FileStream(arquivo, FileMode.Open, FileAccess.Read, FileShare.None)
                    End Using
                Catch ex As Exception
                    bloqueados.Add(arquivo & " — " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            bloqueados.Add(pastaRaiz & " — não foi possível verificar todos os arquivos: " & ex.Message)
        End Try
        Return bloqueados
    End Function

    Public Function ObterDestinoFinalizados(pastaProjeto As String) As String
        Dim config As ConfiguracaoAuxiliar = CarregarConfiguracao()

        If String.IsNullOrWhiteSpace(pastaProjeto) Then
            Throw New Exception("Nenhuma pasta de projeto foi selecionada.")
        End If

        If Not Directory.Exists(pastaProjeto) Then
            Throw New Exception("A pasta do projeto não foi encontrada:" & Environment.NewLine & pastaProjeto)
        End If

        If String.IsNullOrWhiteSpace(config.PastaFinalizados) Then
            Throw New Exception("Configure a pasta Finalizados na engrenagem.")
        End If

        If Not Directory.Exists(config.PastaFinalizados) Then
            Throw New Exception("A pasta Finalizados está inacessível:" & Environment.NewLine & config.PastaFinalizados)
        End If

        Dim origem As String = Path.GetFullPath(pastaProjeto).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim pastaFinalizados As String = Path.GetFullPath(config.PastaFinalizados).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        Dim nomeModelo As String = Path.GetFileName(origem)

        If String.IsNullOrWhiteSpace(nomeModelo) Then
            Throw New Exception("Não foi possível identificar o nome da pasta do modelo.")
        End If

        Dim tipoProjeto As String = IdentificarTipoProjetoParaFinalizados(origem)

        If tipoProjeto = "" Then
            Throw New Exception(
            "Não foi possível identificar o tipo do projeto pelo caminho selecionado." &
            Environment.NewLine & Environment.NewLine &
            "Caminho analisado:" &
            Environment.NewLine &
            origem
        )
        End If

        Dim pastaTipoFinalizados As String = EncontrarPastaTipoFinalizados(pastaFinalizados, tipoProjeto)

        If pastaTipoFinalizados = "" Then
            Throw New Exception(
            "Não encontrei dentro de Finalizados uma pasta correspondente ao tipo: " & tipoProjeto &
            Environment.NewLine & Environment.NewLine &
            "Pasta Finalizados:" &
            Environment.NewLine &
            pastaFinalizados
        )
        End If

        Dim destino As String = Path.Combine(pastaTipoFinalizados, nomeModelo)
        destino = Path.GetFullPath(destino).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        If Directory.Exists(destino) OrElse File.Exists(destino) Then
            Throw New Exception(
            "Já existe uma pasta ou arquivo no destino Finalizados:" &
            Environment.NewLine &
            destino
        )
        End If

        Return destino
    End Function

    Private Function IdentificarTipoProjetoParaFinalizados(caminho As String) As String
        If String.IsNullOrWhiteSpace(caminho) Then
            Return ""
        End If

        Dim partes As String() = caminho.Split(
        New Char() {Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar},
        StringSplitOptions.RemoveEmptyEntries
    )

        ' Começa do penúltimo item para ignorar o nome da pasta do modelo.
        For i As Integer = partes.Length - 2 To 0 Step -1
            Dim tipo As String = NormalizarTipoProjetoParaFinalizados(partes(i))

            If tipo <> "" Then
                Return tipo
            End If
        Next

        Return ""
    End Function

    Private Function EncontrarPastaTipoFinalizados(pastaFinalizados As String, tipoProjeto As String) As String
        If String.IsNullOrWhiteSpace(pastaFinalizados) Then
            Return ""
        End If

        If String.IsNullOrWhiteSpace(tipoProjeto) Then
            Return ""
        End If

        If Not Directory.Exists(pastaFinalizados) Then
            Return ""
        End If

        For Each pasta As String In Directory.GetDirectories(pastaFinalizados)
            Dim nomePasta As String = Path.GetFileName(pasta)
            Dim tipoPasta As String = NormalizarTipoProjetoParaFinalizados(nomePasta)

            If tipoPasta = tipoProjeto Then
                Return pasta
            End If
        Next

        Return ""
    End Function

    Private Function NormalizarTipoProjetoParaFinalizados(nome As String) As String
        If String.IsNullOrWhiteSpace(nome) Then
            Return ""
        End If

        Dim limpo As String = ModuloFuncoes.LimparNomePasta(nome)

        Select Case limpo
            Case "colecao", "escala"
                Return "colecao"

            Case "piloto", "pilotos"
                Return "piloto"

            Case "teste", "testes"
                Return "teste"

            Case "reforma", "reformas"
                Return "reforma"

            Case "limpeza", "limpezas"
                Return "limpeza"

            Case "retoque", "retoques"
                Return "retoque"

            Case Else
                Return ""
        End Select
    End Function

    Public Function MoverPastaParaFinalizados(origem As String, destino As String) As ResultadoFinalizacao
        origem = Path.GetFullPath(origem).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        destino = Path.GetFullPath(destino).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        If Not Directory.Exists(origem) Then Throw New DirectoryNotFoundException(origem)
        If Directory.Exists(destino) OrElse File.Exists(destino) Then
            Throw New IOException("O destino já existe: " & destino)
        End If

        Dim resultado As New ResultadoFinalizacao() With {
            .PastaDestino = destino,
            .QuantidadeArquivos = Directory.GetFiles(origem, "*", SearchOption.AllDirectories).Length
        }

        Dim raizOrigem As String = Path.GetPathRoot(origem)
        Dim raizDestino As String = Path.GetPathRoot(destino)
        If String.Equals(raizOrigem, raizDestino, StringComparison.OrdinalIgnoreCase) Then
            Directory.CreateDirectory(Path.GetDirectoryName(destino))
            Directory.Move(origem, destino)
            Return resultado
        End If

        resultado.CopiaEntreVolumes = True
        Dim temporario As String = destino & ".movendo_" & DateTime.Now.ToString("yyyyMMdd_HHmmssfff")
        Try
            CopiarPastaRecursivamente(origem, temporario, False)
            VerificarCopiaCompleta(origem, temporario)
            Directory.CreateDirectory(Path.GetDirectoryName(destino))
            Directory.Move(temporario, destino)

            Try
                Directory.Delete(origem, True)
            Catch exExcluir As Exception
                Throw New Exception(
                    "A pasta foi copiada e conferida em Finalizados, mas a origem não pôde ser removida:" &
                    Environment.NewLine & origem & Environment.NewLine & Environment.NewLine & exExcluir.Message,
                    exExcluir
                )
            End Try
        Catch
            If Directory.Exists(temporario) Then
                Try
                    Directory.Delete(temporario, True)
                Catch
                End Try
            End If
            Throw
        End Try
        Return resultado
    End Function

    Private Sub VerificarCopiaCompleta(origem As String, copia As String)
        Dim arquivosOrigem() As String = Directory.GetFiles(origem, "*", SearchOption.AllDirectories)
        Dim arquivosCopia() As String = Directory.GetFiles(copia, "*", SearchOption.AllDirectories)
        If arquivosOrigem.Length <> arquivosCopia.Length Then
            Throw New IOException("A conferência da cópia falhou: quantidade de arquivos diferente.")
        End If

        For Each arquivoOrigem As String In arquivosOrigem
            Dim relativo As String = arquivoOrigem.Substring(origem.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim arquivoCopia As String = Path.Combine(copia, relativo)
            If Not File.Exists(arquivoCopia) OrElse
               New FileInfo(arquivoOrigem).Length <> New FileInfo(arquivoCopia).Length Then
                Throw New IOException("A conferência da cópia falhou no arquivo: " & arquivoOrigem)
            End If
        Next
    End Sub

    Private Function ExisteSimulator(pasta As String) As Boolean
        Try
            For Each arquivo As String In Directory.GetFiles(pasta, "*.*", SearchOption.AllDirectories)
                Dim nome As String = Path.GetFileName(arquivo).ToLower()

                If nome.Contains("simulator") OrElse nome.Contains("simulador") Then
                    Return True
                End If
            Next
        Catch
            Return False
        End Try

        Return False
    End Function

    Public Sub AbrirPasta(pasta As String)
        If String.IsNullOrWhiteSpace(pasta) Then
            Throw New Exception("Nenhuma pasta informada.")
        End If

        If Not Directory.Exists(pasta) Then
            Throw New Exception("Pasta não encontrada: " & pasta)
        End If

        Process.Start("explorer.exe", pasta)
    End Sub

    Public Function ObterCaminhoArquivoConfig() As String
        Dim pastaConfig As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "CADLaser",
            "Auxiliar"
        )

        If Not Directory.Exists(pastaConfig) Then
            Directory.CreateDirectory(pastaConfig)
        End If

        Return Path.Combine(pastaConfig, "configuracao.txt")
    End Function

    Private Function ObterCaminhoArquivoConfigLegado() As String
        Return Path.Combine(Application.StartupPath, "config", "configuracao.txt")
    End Function

    Public Function CarregarConfiguracao() As ConfiguracaoAuxiliar
        Dim config As New ConfiguracaoAuxiliar()
        Dim arquivo As String = ObterCaminhoArquivoConfig()

        If Not File.Exists(arquivo) Then
            Dim legado As String = ObterCaminhoArquivoConfigLegado()
            If File.Exists(legado) Then
                Try
                    File.Copy(legado, arquivo, False)
                Catch
                    arquivo = legado
                End Try
            End If
        End If

        If Not File.Exists(arquivo) Then
            Return config
        End If

        Try
            For Each linha As String In File.ReadAllLines(arquivo, Encoding.UTF8)
                If linha.StartsWith("PROJETOS=", StringComparison.OrdinalIgnoreCase) Then
                    config.PastaProjetos = linha.Substring("PROJETOS=".Length).Trim()
                ElseIf linha.StartsWith("PROGRAMA_CN=", StringComparison.OrdinalIgnoreCase) Then
                    config.PastaCN = linha.Substring("PROGRAMA_CN=".Length).Trim()
                ElseIf linha.StartsWith("FINALIZADOS=", StringComparison.OrdinalIgnoreCase) Then
                    config.PastaFinalizados = linha.Substring("FINALIZADOS=".Length).Trim()
                ElseIf linha.StartsWith("CHARMILLES=", StringComparison.OrdinalIgnoreCase) Then
                    config.PastaCharmilles = linha.Substring("CHARMILLES=".Length).Trim()
                ElseIf linha.StartsWith("RHINO=", StringComparison.OrdinalIgnoreCase) Then
                    config.CaminhoRhino = linha.Substring("RHINO=".Length).Trim()
                End If
            Next
        Catch
        End Try

        Return config
    End Function

    Public Sub SalvarConfiguracao(config As ConfiguracaoAuxiliar)
        If config Is Nothing Then
            config = New ConfiguracaoAuxiliar()
        End If

        Dim arquivo As String = ObterCaminhoArquivoConfig()
        Dim linhas As New List(Of String)()

        linhas.Add("PROJETOS=" & config.PastaProjetos)
        linhas.Add("PROGRAMA_CN=" & config.PastaCN)
        linhas.Add("FINALIZADOS=" & config.PastaFinalizados)
        linhas.Add("CHARMILLES=" & config.PastaCharmilles)
        linhas.Add("RHINO=" & config.CaminhoRhino)

        File.WriteAllLines(arquivo, linhas.ToArray(), Encoding.UTF8)
    End Sub

End Module
