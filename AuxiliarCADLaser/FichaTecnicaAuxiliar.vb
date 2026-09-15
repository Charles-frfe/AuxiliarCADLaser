Option Strict Off

Imports System.IO
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Module FichaTecnicaAuxiliar

    Public Function GerarFichaTeste(caminhoModeloFicha As String,
                                    pastaDestino As String,
                                    codigo As String,
                                    nomeModelo As String,
                                    tipoServico As String,
                                    componente As String,
                                    matriz As MatrizAuxiliar,
                                    Optional numeroEscala As String = "") As String

        If String.IsNullOrWhiteSpace(caminhoModeloFicha) OrElse Not File.Exists(caminhoModeloFicha) Then
            Throw New Exception("Modelo da ficha técnica não encontrado.")
        End If

        If matriz Is Nothing Then
            Throw New Exception("Nenhuma matriz foi informada para gerar a ficha.")
        End If

        If String.IsNullOrWhiteSpace(pastaDestino) OrElse Not Directory.Exists(pastaDestino) Then
            Throw New Exception("Pasta de destino da ficha não encontrada.")
        End If

        Dim pastaFichas As String = Path.Combine(pastaDestino, "Ficha Técnica")
        Directory.CreateDirectory(pastaFichas)

        Dim extensaoModelo As String =
            Path.GetExtension(caminhoModeloFicha).ToLowerInvariant()

        If extensaoModelo <> ".xlsx" AndAlso extensaoModelo <> ".xlsm" Then
            Throw New Exception("O modelo da ficha técnica precisa ser .xlsx ou .xlsm.")
        End If

        Dim tipoMatrizFicha As String =
            ObterTipoMatrizParaFicha(matriz)

        Dim sufixoNumero As String = ""

        If Not String.IsNullOrWhiteSpace(numeroEscala) Then
            sufixoNumero = "_" & LimparNomeArquivo(numeroEscala)
        End If

        Dim nomeArquivo As String =
            "Ficha_Tecnica_" & LimparNomeArquivo(codigo) & "_" &
            LimparNomeArquivo(tipoMatrizFicha) &
            sufixoNumero &
            extensaoModelo


        Dim caminhoSaida As String = Path.Combine(pastaFichas, nomeArquivo)

        If File.Exists(caminhoSaida) Then
            Dim resposta As DialogResult = MessageBox.Show(
                "A ficha já existe:" & Environment.NewLine &
                caminhoSaida & Environment.NewLine & Environment.NewLine &
                "Deseja substituir?",
                "Ficha Técnica",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            )

            If resposta <> DialogResult.Yes Then
                Return ""
            End If

            File.Delete(caminhoSaida)
        End If

        File.Copy(caminhoModeloFicha, caminhoSaida, True)

        Dim excel As Object = Nothing
        Dim wb As Object = Nothing

        Try
            excel = CreateObject("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False

            wb = excel.Workbooks.Open(caminhoSaida)

            Dim nomeAbaBase As String = ObterAbaBase(tipoMatrizFicha)
            Dim ws As Object = ObterPlanilha(wb, nomeAbaBase)

            If ws Is Nothing Then
                Throw New Exception("A aba base não foi encontrada na ficha: " & nomeAbaBase)
            End If

            ManterSomenteAba(wb, ws)
            ws.Name = LimitarNomeAba(tipoMatrizFicha)

            PreencherFicha(ws, codigo, nomeModelo, tipoServico, componente, matriz, numeroEscala)

            wb.Save()
            wb.Close(False)
            excel.Quit()

            Return caminhoSaida

        Catch
            Try
                If wb IsNot Nothing Then wb.Close(False)
            Catch
            End Try

            Try
                If excel IsNot Nothing Then excel.Quit()
            Catch
            End Try

            Throw
        Finally
            LiberarObjeto(wb)
            LiberarObjeto(excel)
        End Try
    End Function

    Public Function GerarFichaAgrupada(caminhoModeloFicha As String,
                                   pastaDestino As String,
                                   codigo As String,
                                   nomeModelo As String,
                                   tipoServico As String,
                                   componente As String,
                                   matrizesFicha As List(Of MatrizAuxiliar),
                                   Optional numeroEscala As String = "") As String

        If String.IsNullOrWhiteSpace(caminhoModeloFicha) OrElse Not File.Exists(caminhoModeloFicha) Then
            Throw New Exception("Modelo da ficha técnica não encontrado.")
        End If

        If matrizesFicha Is Nothing OrElse matrizesFicha.Count = 0 Then
            Throw New Exception("Nenhuma matriz foi informada para gerar a ficha.")
        End If

        If String.IsNullOrWhiteSpace(pastaDestino) OrElse Not Directory.Exists(pastaDestino) Then
            Throw New Exception("Pasta de destino da ficha não encontrada.")
        End If

        Dim extensaoModelo As String = Path.GetExtension(caminhoModeloFicha).ToLowerInvariant()

        If extensaoModelo <> ".xlsx" AndAlso extensaoModelo <> ".xlsm" Then
            Throw New Exception("O modelo da ficha técnica precisa ser .xlsx ou .xlsm.")
        End If

        Dim pastaFichas As String = Path.Combine(pastaDestino, "Ficha Técnica")
        Directory.CreateDirectory(pastaFichas)

        Dim sufixoNumero As String = ""

        If Not String.IsNullOrWhiteSpace(numeroEscala) Then
            sufixoNumero = "_" & LimparNomeArquivo(numeroEscala)
        End If

        Dim nomeArquivo As String =
        "Ficha_Tecnica_" &
        LimparNomeArquivo(codigo) &
        sufixoNumero &
        extensaoModelo

        Dim caminhoSaida As String = Path.Combine(pastaFichas, nomeArquivo)

        If File.Exists(caminhoSaida) Then
            Dim resposta As DialogResult = MessageBox.Show(
            "A ficha já existe:" & Environment.NewLine &
            caminhoSaida & Environment.NewLine & Environment.NewLine &
            "Deseja substituir?",
            "Ficha Técnica",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )

            If resposta <> DialogResult.Yes Then
                Return ""
            End If

            File.Delete(caminhoSaida)
        End If

        File.Copy(caminhoModeloFicha, caminhoSaida, True)

        Dim excel As Object = Nothing
        Dim wb As Object = Nothing

        Try
            excel = CreateObject("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False

            wb = excel.Workbooks.Open(caminhoSaida)

            Dim nomesTemporarios As New List(Of String)()
            Dim nomesDesejados As New List(Of String)()
            Dim contador As Integer = 1

            For Each matriz As MatrizAuxiliar In matrizesFicha
                If matriz Is Nothing Then Continue For

                Dim tipoMatrizFicha As String = ObterTipoMatrizParaFicha(matriz)
                Dim nomeAbaBase As String = ObterAbaBase(tipoMatrizFicha)
                Dim abaBase As Object = ObterPlanilha(wb, nomeAbaBase)

                If abaBase Is Nothing Then
                    Throw New Exception("A aba base não foi encontrada na ficha: " & nomeAbaBase)
                End If

                abaBase.Copy(After:=wb.Worksheets(wb.Worksheets.Count))

                Dim ws As Object = wb.Worksheets(wb.Worksheets.Count)
                Dim nomeTemporario As String = "__GERADA_" & contador.ToString("000")

                ws.Name = nomeTemporario

                nomesTemporarios.Add(nomeTemporario)
                nomesDesejados.Add(LimitarNomeAba(tipoMatrizFicha))

                PreencherFicha(ws, codigo, nomeModelo, tipoServico, componente, matriz, numeroEscala)

                contador += 1
            Next

            ManterSomenteAbasTemporarias(wb, nomesTemporarios)
            RenomearAbasGeradas(wb, nomesTemporarios, nomesDesejados)

            wb.Save()
            wb.Close(False)
            excel.Quit()

            Return caminhoSaida

        Catch
            Try
                If wb IsNot Nothing Then wb.Close(False)
            Catch
            End Try

            Try
                If excel IsNot Nothing Then excel.Quit()
            Catch
            End Try

            Throw
        Finally
            LiberarObjeto(wb)
            LiberarObjeto(excel)
        End Try
    End Function

    Private Sub ManterSomenteAbasTemporarias(wb As Object, nomesTemporarios As List(Of String))
        For i As Integer = wb.Worksheets.Count To 1 Step -1
            Dim ws As Object = wb.Worksheets(i)
            Dim nome As String = ws.Name.ToString()

            Dim manter As Boolean = nomesTemporarios.Exists(
            Function(item As String)
                Return String.Equals(item, nome, StringComparison.OrdinalIgnoreCase)
            End Function
        )

            If Not manter Then
                ws.Delete()
            End If
        Next
    End Sub

    Private Sub RenomearAbasGeradas(wb As Object,
                                nomesTemporarios As List(Of String),
                                nomesDesejados As List(Of String))

        For i As Integer = 0 To nomesTemporarios.Count - 1
            Dim ws As Object = ObterPlanilha(wb, nomesTemporarios(i))

            If ws IsNot Nothing Then
                ws.Name = ObterNomeAbaDisponivel(wb, nomesDesejados(i))
            End If
        Next
    End Sub

    Private Function ObterNomeAbaDisponivel(wb As Object, nomeBase As String) As String
        Dim baseLimpa As String = LimitarNomeAba(nomeBase)

        If baseLimpa = "" Then baseLimpa = "Ficha"

        For tentativa As Integer = 0 To 50
            Dim candidato As String

            If tentativa = 0 Then
                candidato = baseLimpa
            Else
                Dim sufixo As String = " " & (tentativa + 1).ToString()
                Dim limite As Integer = 31 - sufixo.Length

                candidato = baseLimpa

                If candidato.Length > limite Then
                    candidato = candidato.Substring(0, limite)
                End If

                candidato &= sufixo
            End If

            If ObterPlanilha(wb, candidato) Is Nothing Then
                Return candidato
            End If
        Next

        Return LimitarNomeAba(baseLimpa & " " & DateTime.Now.ToString("ss"))
    End Function

    Private Sub PreencherFicha(ws As Object,
                           codigo As String,
                           nomeModelo As String,
                           tipoServico As String,
                           componente As String,
                           matriz As MatrizAuxiliar,
                           numeroEscala As String)

        Dim nomeSemCodigo As String =
            FormatarNomeModeloFicha(nomeModelo, codigo)

        Dim tipoFicha As String = "PILOTO"

        If tipoServico.Trim().ToLower() = "escala" Then
            tipoFicha = "COLEÇÃO"
        End If

        Dim componenteFicha As String = FormatarComponenteFicha(componente, tipoServico, matriz, numeroEscala)

        EscreverCelula(ws, "B4", codigo)
        EscreverCelula(ws, "B6", nomeSemCodigo)
        EscreverCelula(ws, "H5", tipoFicha)

        EscreverCelula(ws, "C7", componenteFicha)

        EscreverCelula(ws, "L2", matriz.Maquina)
        EscreverCelula(ws, "B9", matriz.Material)
        EscreverCelula(ws, "D9", matriz.Eixos)
        EscreverCelula(ws, "B11", matriz.Projetista)

        EscreverValorDepoisDoTitulo(
    ws,
    "TEXT",
    "E10",
    "Text:",
    TexturasEmLinha(matriz.TextoTexturas),
    False,
    True
)

        If Not String.IsNullOrWhiteSpace(matriz.ObservacoesAcabamento) Then
            EscreverValorDepoisDoTitulo(
                ws,
                "OBSERVAÇÕES",
                "J17",
                "OBSERVAÇÕES:",
                FormatarTextoFicha(matriz.ObservacoesAcabamento),
                True,
                False
            )
        End If

        If Not String.IsNullOrWhiteSpace(matriz.Acabamento) Then
            EscreverValorDepoisDoTitulo(
                ws,
                "ACABAMENTO",
                "J24",
                "ACABAMENTO:",
                FormatarTextoFicha(matriz.Acabamento),
                False,
                False
            )
        End If

    End Sub

    Private Sub EscreverValorDepoisDoTitulo(ws As Object,
                                        textoProcurado As String,
                                        celulaFallback As String,
                                        tituloPadrao As String,
                                        valor As String,
                                        Optional valorVermelho As Boolean = False,
                                        Optional valorNegrito As Boolean = False)

        If String.IsNullOrWhiteSpace(valor) Then Return

        Dim celula As Object = ProcurarCelula(ws, textoProcurado)

        If celula Is Nothing Then
            Try
                celula = ws.Range(celulaFallback)
            Catch
                Return
            End Try
        End If

        Try
            Dim destino As Object = celula

            If destino.MergeCells Then
                destino = destino.MergeArea.Cells(1, 1)
            End If

            Dim tituloFinal As String = tituloPadrao.Trim()

            Try
                Dim textoAtual As String = ""

                If destino.Value IsNot Nothing Then
                    textoAtual = destino.Value.ToString().Trim()
                End If

                If textoAtual <> "" Then
                    Dim posDoisPontos As Integer = textoAtual.IndexOf(":"c)

                    If posDoisPontos >= 0 Then
                        tituloFinal = textoAtual.Substring(0, posDoisPontos + 1).Trim()
                    Else
                        tituloFinal = textoAtual.Trim() & ":"
                    End If
                End If
            Catch
            End Try

            If Not tituloFinal.EndsWith(":") Then
                tituloFinal &= ":"
            End If

            Dim valorFinal As String = valor.Trim().ToUpperInvariant()
            Dim textoFinal As String = tituloFinal & " " & valorFinal

            destino.Value = textoFinal

            Dim inicioValor As Integer = tituloFinal.Length + 2
            Dim tamanhoValor As Integer = valorFinal.Length

            If tamanhoValor > 0 Then
                If valorVermelho Then
                    destino.Characters(inicioValor, tamanhoValor).Font.Color = 255
                End If

                If valorNegrito Then
                    destino.Characters(inicioValor, tamanhoValor).Font.Bold = True
                End If
            End If

        Catch
        End Try
    End Sub

    Private Function FormatarNomeModeloFicha(nomeModelo As String, codigo As String) As String
        If String.IsNullOrWhiteSpace(nomeModelo) Then Return ""

        Dim texto As String = nomeModelo.Trim()

        If Not String.IsNullOrWhiteSpace(codigo) AndAlso
       texto.StartsWith(codigo.Trim(), StringComparison.OrdinalIgnoreCase) Then

            texto = texto.Substring(codigo.Trim().Length)

            While texto.StartsWith("_") OrElse texto.StartsWith("-") OrElse texto.StartsWith(" ")
                texto = texto.Substring(1)
            End While
        Else
            Dim posUnderline As Integer = texto.IndexOf("_"c)

            If posUnderline > 0 Then
                Dim primeiraParte As String = texto.Substring(0, posUnderline)

                If SomenteDigitos(primeiraParte) Then
                    texto = texto.Substring(posUnderline + 1)
                End If
            End If
        End If

        Return FormatarTextoFicha(texto)
    End Function

    Private Function SomenteDigitos(texto As String) As Boolean
        If String.IsNullOrWhiteSpace(texto) Then Return False

        For Each ch As Char In texto
            If Not Char.IsDigit(ch) Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Sub EscreverTextoComTituloNaCaixa(ws As Object,
                                              textoProcurado As String,
                                              celulaFallback As String,
                                              titulo As String,
                                              valor As String,
                                              Optional vermelho As Boolean = False)

        If String.IsNullOrWhiteSpace(valor) Then Return

        Dim celula As Object = ProcurarCelula(ws, textoProcurado)

        If celula Is Nothing Then
            Try
                celula = ws.Range(celulaFallback)
            Catch
                Return
            End Try
        End If

        Try
            Dim destino As Object = celula

            If destino.MergeCells Then
                destino = destino.MergeArea.Cells(1, 1)
            End If

            Dim textoFinal As String = titulo & valor.Trim().ToUpperInvariant()

            destino.Value = textoFinal

            If vermelho Then
                destino.Font.Color = 255
            End If

        Catch
        End Try
    End Sub

    Private Sub EscreverTexturaNaCaixaText(ws As Object, valor As String)
        If String.IsNullOrWhiteSpace(valor) Then Return

        Dim celula As Object = ProcurarCelula(ws, "TEXT")

        If celula Is Nothing Then
            Try
                celula = ws.Range("E10")
            Catch
                Return
            End Try
        End If

        Try
            Dim destino As Object = celula

            If destino.MergeCells Then
                destino = destino.MergeArea.Cells(1, 1)
            End If

            Dim textoFinal As String = "Text:" & valor.Trim().ToUpperInvariant

            destino.Value = textoFinal
            destino.Font.Bold = False

            Dim inicioNegrito As Integer = 7
            Dim tamanhoNegrito As Integer = valor.Trim().Length

            destino.Characters(inicioNegrito, tamanhoNegrito).Font.Bold = True
        Catch
        End Try
    End Sub

    Private Function FormatarTextoFicha(textoOriginal As String) As String
        If String.IsNullOrWhiteSpace(textoOriginal) Then Return ""

        Dim texto As String = textoOriginal.Trim()

        texto = texto.Replace("_", " ")
        texto = texto.Replace("-", " ")

        While texto.Contains("  ")
            texto = texto.Replace("  ", " ")
        End While

        Return texto.Trim().ToUpperInvariant()
    End Function

    Private Function FormatarComponenteFicha(componente As String,
                                         tipoServico As String,
                                         matriz As MatrizAuxiliar,
                                         numeroEscala As String) As String

        If String.Equals(If(tipoServico, "").Trim(), "escala", StringComparison.OrdinalIgnoreCase) Then
            Dim tipoMatrizFicha As String = FormatarTextoFicha(ObterTipoMatrizParaFicha(matriz))
            Dim numeroFormatado As String = FormatarNumeroEscalaFicha(numeroEscala)

            If numeroFormatado <> "" Then
                Return tipoMatrizFicha & " " & numeroFormatado
            End If

            Return tipoMatrizFicha
        End If

        Return FormatarTextoFicha(componente)
    End Function

    Private Function FormatarNumeroEscalaFicha(numeroOriginal As String) As String
        If String.IsNullOrWhiteSpace(numeroOriginal) Then Return ""

        Dim numero As String = numeroOriginal.Trim()
        Dim partes As String() = numero.Split("-"c)
        Dim partesFormatadas As New List(Of String)()

        For Each parte As String In partes
            Dim limpa As String = parte.Trim()

            If limpa.Length = 4 AndAlso limpa.All(Function(ch) Char.IsDigit(ch)) Then
                partesFormatadas.Add(limpa.Substring(0, 2) & "/" & limpa.Substring(2, 2))
            Else
                partesFormatadas.Add(limpa)
            End If
        Next

        Return String.Join("-", partesFormatadas.ToArray())
    End Function

    Private Function ObterPrimeiroNumeroEscala(matriz As MatrizAuxiliar) As String
        If matriz Is Nothing OrElse String.IsNullOrWhiteSpace(matriz.TextoNumeros) Then
            Return ""
        End If

        Try
            Dim numeros As List(Of String) = ModuloFuncoes.QuebrarNumerosPorVirgula(matriz.TextoNumeros)

            If numeros.Count > 0 Then
                Return numeros(0)
            End If
        Catch
        End Try

        Return ""
    End Function

    Private Sub EscreverCelula(ws As Object, endereco As String, valor As String, Optional vermelho As Boolean = False)
        If String.IsNullOrWhiteSpace(valor) Then Return

        Try
            Dim destino As Object = ws.Range(endereco)

            Try
                If destino.MergeCells Then
                    destino = destino.MergeArea.Cells(1, 1)
                End If
            Catch
            End Try

            destino.Value = valor

            If vermelho Then
                destino.Font.Color = 255
            End If
        Catch
        End Try
    End Sub

    Private Function ObterTipoMatrizParaFicha(matriz As MatrizAuxiliar) As String
        If matriz Is Nothing Then Return ""

        Dim familia As String = ModuloFuncoes.LimparNomePasta(matriz.FamiliaMatriz)

        If familia = "encaixe" Then
            Return "Encaixe"
        End If

        If Not String.IsNullOrWhiteSpace(matriz.TipoMatriz) Then
            Return matriz.TipoMatriz.Trim()
        End If

        Return matriz.FamiliaMatriz.Trim()
    End Function
    Private Function ObterAbaBase(tipoMatriz As String) As String
        Dim tipo As String = ModuloFuncoes.LimparNomePasta(tipoMatriz)

        Select Case tipo
            Case "tampa"
                Return "TAMPA"

            Case "lateral", "lfe", "lfd", "lme", "lmd"
                Return "LATERAIS"

            Case "encaixe"
                Return "FERRAMENTAS"

            Case Else
                Return "GRAVAÇÃO"
        End Select
    End Function

    Private Function ObterPlanilha(wb As Object, nomeAba As String) As Object
        For i As Integer = 1 To wb.Worksheets.Count
            Dim ws As Object = wb.Worksheets(i)

            If String.Equals(ws.Name.ToString().Trim(), nomeAba, StringComparison.OrdinalIgnoreCase) Then
                Return ws
            End If
        Next

        Return Nothing
    End Function

    Private Sub ManterSomenteAba(wb As Object, abaParaManter As Object)
        For i As Integer = wb.Worksheets.Count To 1 Step -1
            Dim ws As Object = wb.Worksheets(i)

            If Not Object.ReferenceEquals(ws, abaParaManter) Then
                ws.Delete()
            End If
        Next
    End Sub

    Private Function EscreverAoLado(ws As Object,
                                textoProcurado As String,
                                valor As String,
                                Optional vermelho As Boolean = False) As Boolean
        If String.IsNullOrWhiteSpace(valor) Then Return False

        Dim celula As Object = ProcurarCelula(ws, textoProcurado)

        If celula Is Nothing Then
            Return False
        End If

        Try
            Dim area As Object = celula.MergeArea
            Dim linha As Integer = area.Row
            Dim colunaDestino As Integer = area.Column + area.Columns.Count

            Dim destino As Object = ws.Cells(linha, colunaDestino)

            Try
                If destino.MergeCells Then
                    destino = destino.MergeArea.Cells(1, 1)
                End If
            Catch
            End Try

            destino.Value = valor

            If vermelho Then
                destino.Font.Color = 255
            End If

            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub EscreverAoLadoOuCelula(ws As Object,
                                       textoProcurado As String,
                                       celulaFallBack As String,
                                       valor As String,
                                       Optional vermelho As Boolean = False)
        If String.IsNullOrWhiteSpace(valor) Then Return

        If EscreverAoLado(ws, textoProcurado, valor, vermelho) Then
            Return
        End If

        EscreverCelula(ws, celulaFallBack, valor, vermelho)
    End Sub

    Private Function ProcurarCelula(ws As Object, textoProcurado As String) As Object
        Dim usado As Object = ws.UsedRange

        For linha As Integer = 1 To usado.Rows.Count
            For coluna As Integer = 1 To usado.Columns.Count
                Dim celula As Object = usado.Cells(linha, coluna)
                Dim textoCelula As String = ""

                Try
                    If celula.Value IsNot Nothing Then
                        textoCelula = celula.Value.ToString().Trim()
                    End If
                Catch
                    textoCelula = ""
                End Try

                If String.Equals(NormalizarTexto(textoCelula), NormalizarTexto(textoProcurado), StringComparison.OrdinalIgnoreCase) Then
                    Return celula
                End If
            Next
        Next

        Return Nothing
    End Function

    Private Function NormalizarTexto(texto As String) As String
        If texto Is Nothing Then Return ""

        Dim resultado As String =
            ModuloFuncoes.RemoverAcentos(texto).Trim().ToUpperInvariant()

        resultado = resultado.Replace(":", "")
        resultado = resultado.Replace(".", "")
        resultado = resultado.Replace(" ", "")
        Return resultado
    End Function

    Private Function TexturasEmLinha(textoTexturas As String) As String
        If String.IsNullOrWhiteSpace(textoTexturas) Then Return ""

        Dim texto As String = textoTexturas.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim linhas As String() = texto.Split(New String() {vbLf}, StringSplitOptions.RemoveEmptyEntries)

        Dim nomes As New List(Of String)()

        For Each linha As String In linhas
            Dim nome As String = ExtrairCodigoTexturaFicha(linha)

            If nome <> "" Then
                nomes.Add(nome)
            End If
        Next

        Return String.Join(" ", nomes.ToArray())
    End Function

    Private Function ExtrairCodigoTexturaFicha(textoOriginal As String) As String
        If String.IsNullOrWhiteSpace(textoOriginal) Then Return ""

        Dim texto As String = textoOriginal.Trim()

        texto = texto.Replace(" ", "_")
        texto = texto.Replace("-", "_")

        Dim textoMin As String = texto.ToLowerInvariant()

        Dim prefixos As String() = {"gls_", "gl_", "gs_"}

        For Each prefixo As String In prefixos
            If textoMin.StartsWith(prefixo) Then
                Dim fim As Integer = prefixo.Length

                While fim < texto.Length AndAlso
                        Char.IsDigit(texto(fim))
                    fim += 1
                End While

                If fim > prefixo.Length Then
                    Return texto.Substring(0, fim).ToUpperInvariant()
                End If
            End If
        Next

        Return FormatarTextoFicha(texto)
    End Function

    Private Function RemoverCodigoDoNome(nomeModelo As String, codigo As String) As String
        If String.IsNullOrWhiteSpace(nomeModelo) Then Return ""

        Dim nome As String = nomeModelo.Trim()

        If Not String.IsNullOrWhiteSpace(codigo) AndAlso nome.StartsWith(codigo) Then
            nome = nome.Substring(codigo.Length).Trim()
        End If

        nome = nome.Trim("-"c, "_"c, " "c)

        Return nome
    End Function

    Private Function LimparNomeArquivo(nome As String) As String
        If String.IsNullOrWhiteSpace(nome) Then Return "sem_nome"

        Dim invalido As Char() = Path.GetInvalidFileNameChars()
        Dim resultado As String = nome.Trim()

        For Each ch As Char In invalido
            resultado = resultado.Replace(ch, "_"c)
        Next

        resultado = resultado.Replace(" ", "_")

        Return resultado
    End Function

    Private Function LimitarNomeAba(nome As String) As String
        Dim resultado As String = If(nome, "").Trim()

        If resultado = "" Then resultado = "Ficha"

        resultado = resultado.Replace(":", "")
        resultado = resultado.Replace("\", "")
        resultado = resultado.Replace("/", "")
        resultado = resultado.Replace("?", "")
        resultado = resultado.Replace("*", "")
        resultado = resultado.Replace("[", "")
        resultado = resultado.Replace("]", "")

        If resultado.Length > 31 Then
            resultado = resultado.Substring(0, 31)
        End If

        Return resultado
    End Function

    Private Sub LiberarObjeto(obj As Object)
        Try
            If obj IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Catch
        End Try
    End Sub

End Module