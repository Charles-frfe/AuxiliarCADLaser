Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Module IntegracaoRhino

    Private Class ArquivoRhinoPlanejado
        Public Caminho As String = ""
        Public Matrizes As New List(Of MatrizAuxiliar)()
    End Class

    Private Class PastaComponenteCandidata
        Public Caminho As String = ""
        Public Nome As String = ""
        Public Pontuacao As Integer = 0
        Public Profundidade As Integer = 0
    End Class

    Public Sub CriarArquivo(pastaModelo As String,
                            codigo As String,
                            tipoServico As String,
                            componente As String,
                            matrizes As List(Of MatrizAuxiliar))

        If String.IsNullOrWhiteSpace(pastaModelo) OrElse Not Directory.Exists(pastaModelo) Then
            Throw New Exception("Selecione uma pasta de modelo existente antes de criar o Rhino.")
        End If

        Dim codigoCincoDigitos As String = ObterCodigoCincoDigitos(codigo)
        Dim componentePasta As String = ModuloFuncoes.LimparNomePasta(componente)

        If codigoCincoDigitos = "" Then
            Throw New Exception("O código do modelo precisa conter os 5 dígitos usados no nome do arquivo Rhino.")
        End If

        If componentePasta = "" Then
            Throw New Exception("Informe o componente.")
        End If

        If matrizes Is Nothing OrElse matrizes.Count = 0 Then
            Throw New Exception("Adicione pelo menos uma matriz antes de criar o Rhino.")
        End If

        Dim pastaComponente As String = LocalizarOuCriarPastaComponente(
            pastaModelo, componente, componentePasta)

        Dim executavelRhino As String = LocalizarExecutavelRhino()
        If executavelRhino = "" Then
            Throw New OperationCanceledException("O executável do Rhino não foi selecionado.")
        End If

        Dim arquivos As List(Of ArquivoRhinoPlanejado) = PlanejarArquivos(
            pastaComponente,
            codigoCincoDigitos,
            tipoServico,
            matrizes
        )

        ConfirmarSubstituicaoEExecutarBackups(arquivos)

        Dim pastaScripts As String = Path.Combine(Path.GetTempPath(), "AuxiliarCADLaser")
        Directory.CreateDirectory(pastaScripts)

        Dim identificador As String = DateTime.Now.ToString("yyyyMMdd_HHmmssfff")
        Dim caminhoScript As String = Path.Combine(pastaScripts, "criar_rhino_" & identificador & ".py")
        Dim caminhoStatus As String = Path.Combine(pastaScripts, "criar_rhino_" & identificador & ".log")

        File.WriteAllText(
            caminhoScript,
            MontarScript(arquivos, caminhoStatus),
            New UTF8Encoding(False)
        )

        Dim argumentos As String = "/nosplash /notemplate /runscript=""-_RunPythonScript (" &
                                   caminhoScript & ")"""

        Dim inicio As New ProcessStartInfo()
        inicio.FileName = executavelRhino
        inicio.Arguments = argumentos
        inicio.UseShellExecute = True
        Dim processoCriacao As Process = Process.Start(inicio)

        AguardarCriacaoEAbrirArquivos(
            executavelRhino,
            arquivos,
            caminhoStatus,
            processoCriacao
        )
    End Sub

    Private Function LocalizarOuCriarPastaComponente(pastaModelo As String,
                                                      componenteOriginal As String,
                                                      componenteLimpo As String) As String
        Dim candidatas As New List(Of PastaComponenteCandidata)()

        Try
            ProcurarPastasComponente(
                pastaModelo,
                pastaModelo,
                componenteLimpo,
                candidatas,
                0,
                8
            )
        Catch ex As Exception
            Throw New Exception("Não foi possível pesquisar as pastas do componente." &
                                Environment.NewLine & ex.Message, ex)
        End Try

        If candidatas.Count > 0 Then
            candidatas.Sort(
                Function(a As PastaComponenteCandidata, b As PastaComponenteCandidata)
                    Dim porPontos As Integer = b.Pontuacao.CompareTo(a.Pontuacao)
                    If porPontos <> 0 Then Return porPontos

                    Dim porProfundidade As Integer = a.Profundidade.CompareTo(b.Profundidade)
                    If porProfundidade <> 0 Then Return porProfundidade

                    Return String.Compare(a.Caminho, b.Caminho, StringComparison.OrdinalIgnoreCase)
                End Function)
        End If

        Dim exatas As List(Of PastaComponenteCandidata) =
            candidatas.FindAll(Function(item As PastaComponenteCandidata) item.Pontuacao = 1000)

        If exatas.Count = 1 Then Return exatas(0).Caminho

        If candidatas.Count > 0 Then
            Dim opcoes As New List(Of OpcaoEscolhaAuxiliar)()

            opcoes.Add(New OpcaoEscolhaAuxiliar() With {
                .Texto = "Selecionar manualmente a pasta onde salvar os Rhinos",
                .Valor = "__SELECIONAR_MANUAL__"
            })

            opcoes.Add(New OpcaoEscolhaAuxiliar() With {
                .Texto = "Criar nova pasta no modelo: " & Path.Combine(pastaModelo, componenteLimpo),
                .Valor = "__CRIAR_NOVA__"
            })

            For Each candidata As PastaComponenteCandidata In candidatas
                opcoes.Add(New OpcaoEscolhaAuxiliar() With {
                    .Texto = candidata.Nome &
                             "  —  " &
                             candidata.Caminho &
                             "  [pontos: " & candidata.Pontuacao.ToString() & "]",
                    .Valor = candidata.Caminho
                })
            Next

            Dim mensagem As String

            If exatas.Count > 1 Then
                mensagem = "Foram encontradas mais de uma pasta EXATA para “" & componenteOriginal.Trim() & "”. Escolha onde salvar o Rhino:"
            Else
                mensagem = "Foram encontradas pastas parecidas com “" & componenteOriginal.Trim() & "”, inclusive em subpastas. Escolha onde salvar o Rhino:"
            End If

            Using tela As New FormEscolherOpcao(
                "Pasta do componente",
                mensagem,
                opcoes)

                If tela.ShowDialog() = DialogResult.OK Then
                    If tela.ValorSelecionado = "__SELECIONAR_MANUAL__" Then
                        Return SelecionarPastaComponenteManual(pastaModelo)
                    End If

                    If tela.ValorSelecionado = "__CRIAR_NOVA__" Then
                        Dim novaPastaEscolhida As String = Path.Combine(pastaModelo, componenteLimpo)
                        Directory.CreateDirectory(novaPastaEscolhida)
                        Return novaPastaEscolhida
                    End If

                    Return tela.ValorSelecionado
                End If
            End Using

            Throw New OperationCanceledException("Criação do Rhino cancelada: nenhuma pasta de componente foi escolhida.")
        End If

        Dim resposta As DialogResult = MessageBox.Show(
            "Nenhuma pasta parecida com o componente “" & componenteOriginal.Trim() & "” foi encontrada, mesmo pesquisando em subpastas." &
            Environment.NewLine & Environment.NewLine &
            "Clique SIM para selecionar manualmente a pasta onde salvar os Rhinos." &
            Environment.NewLine &
            "Clique NÃO para criar automaticamente a pasta abaixo:" &
            Environment.NewLine &
            Path.Combine(pastaModelo, componenteLimpo),
            "Pasta do componente",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question
        )

        If resposta = DialogResult.Yes Then
            Return SelecionarPastaComponenteManual(pastaModelo)
        End If

        If resposta = DialogResult.No Then
            Dim novaPasta As String = Path.Combine(pastaModelo, componenteLimpo)
            Directory.CreateDirectory(novaPasta)
            Return novaPasta
        End If

        Throw New OperationCanceledException("Criação do Rhino cancelada: nenhuma pasta de componente foi escolhida.")
    End Function

    Private Sub ProcurarPastasComponente(pastaAtual As String,
                                         pastaModelo As String,
                                         componenteLimpo As String,
                                         candidatas As List(Of PastaComponenteCandidata),
                                         profundidadeAtual As Integer,
                                         profundidadeMaxima As Integer)

        If profundidadeAtual >= profundidadeMaxima Then
            Return
        End If

        Dim subpastas As String()

        Try
            subpastas = Directory.GetDirectories(pastaAtual)
        Catch
            Return
        End Try

        For Each caminho As String In subpastas
            Dim nome As String = New DirectoryInfo(caminho).Name
            Dim nomeLimpo As String = ModuloFuncoes.LimparNomePasta(nome)
            Dim pontos As Integer = PontuarPastaComponente(nomeLimpo, componenteLimpo)

            If pontos > 0 Then
                candidatas.Add(New PastaComponenteCandidata() With {
                    .Caminho = caminho,
                    .Nome = nome,
                    .Pontuacao = pontos,
                    .Profundidade = profundidadeAtual + 1
                })
            End If

            ProcurarPastasComponente(
                caminho,
                pastaModelo,
                componenteLimpo,
                candidatas,
                profundidadeAtual + 1,
                profundidadeMaxima
            )
        Next
    End Sub

    Private Function SelecionarPastaComponenteManual(pastaInicial As String) As String
        Using dialogo As New FolderBrowserDialog()
            dialogo.Description = "Escolha a pasta onde os arquivos Rhino da escala serão criados."
            dialogo.ShowNewFolderButton = True

            If Directory.Exists(pastaInicial) Then
                dialogo.SelectedPath = pastaInicial
            End If

            If dialogo.ShowDialog() <> DialogResult.OK Then
                Throw New OperationCanceledException("Criação do Rhino cancelada: nenhuma pasta foi selecionada.")
            End If

            If String.IsNullOrWhiteSpace(dialogo.SelectedPath) OrElse Not Directory.Exists(dialogo.SelectedPath) Then
                Throw New OperationCanceledException("Criação do Rhino cancelada: a pasta selecionada não é válida.")
            End If

            Return dialogo.SelectedPath
        End Using
    End Function

    Private Function PontuarPastaComponente(nomePasta As String, componente As String) As Integer
        If nomePasta = "" OrElse componente = "" Then Return 0
        If String.Equals(nomePasta, componente, StringComparison.OrdinalIgnoreCase) Then Return 1000
        If nomePasta.StartsWith(componente & "_", StringComparison.OrdinalIgnoreCase) OrElse
           nomePasta.StartsWith(componente & "-", StringComparison.OrdinalIgnoreCase) Then Return 900
        If nomePasta.EndsWith("_" & componente, StringComparison.OrdinalIgnoreCase) OrElse
           nomePasta.EndsWith("-" & componente, StringComparison.OrdinalIgnoreCase) Then Return 880

        Dim tokensComponente() As String = componente.Split(New Char() {"_"c, "-"c}, StringSplitOptions.RemoveEmptyEntries)
        Dim tokensPasta As New HashSet(Of String)(
            nomePasta.Split(New Char() {"_"c, "-"c}, StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase)
        If tokensComponente.Length > 0 AndAlso
           Array.TrueForAll(tokensComponente, Function(token As String) tokensPasta.Contains(token)) Then
            Return 800 + tokensComponente.Length
        End If

        Return 0
    End Function

    Private Sub AguardarCriacaoEAbrirArquivos(executavelRhino As String,
                                              arquivos As List(Of ArquivoRhinoPlanejado),
                                              caminhoStatus As String,
                                              processoCriacao As Process)

        Dim horarioLimite As DateTime = DateTime.Now.AddMinutes(3)
        Dim resultado As String = ""

        Do While DateTime.Now < horarioLimite
            Application.DoEvents()
            Global.System.Threading.Thread.Sleep(300)

            Try
                If File.Exists(caminhoStatus) Then
                    resultado = File.ReadAllText(caminhoStatus, Encoding.UTF8).Trim()
                    If resultado <> "" Then
                        Exit Do
                    End If
                End If
            Catch
                ' O Rhino pode estar escrevendo o arquivo de status neste momento.
            End Try
        Loop

        If resultado = "" Then
            Throw New Exception("O Rhino não concluiu a criação dos arquivos no tempo esperado.")
        End If

        If Not String.Equals(resultado, "OK", StringComparison.OrdinalIgnoreCase) Then
            Throw New Exception("Não foi possível criar os arquivos Rhino." &
                                Environment.NewLine & Environment.NewLine & resultado)
        End If

        If processoCriacao IsNot Nothing Then
            Try
                Dim limiteFechamento As DateTime = DateTime.Now.AddSeconds(15)
                Do While Not processoCriacao.HasExited AndAlso DateTime.Now < limiteFechamento
                    Application.DoEvents()
                    Global.System.Threading.Thread.Sleep(250)
                Loop
            Catch
                ' A criação terminou; continua mesmo se o processo não puder ser consultado.
            End Try
        End If

        Global.System.Threading.Thread.Sleep(700)

        Dim arquivosCriados As New List(Of ArquivoRhinoPlanejado)()

        For Each arquivo As ArquivoRhinoPlanejado In arquivos
            If arquivo Is Nothing OrElse Not File.Exists(arquivo.Caminho) Then
                Continue For
            End If

            arquivosCriados.Add(arquivo)
        Next

        If arquivosCriados.Count = 0 Then
            Throw New Exception("A criação terminou, mas nenhum arquivo .3dm foi encontrado para abrir.")
        End If

        Dim primeiroArquivo As ArquivoRhinoPlanejado = arquivosCriados(0)

        Dim inicioArquivo As New ProcessStartInfo()
        inicioArquivo.FileName = executavelRhino
        inicioArquivo.Arguments = "/nosplash """ & primeiroArquivo.Caminho & """"
        inicioArquivo.UseShellExecute = True
        Process.Start(inicioArquivo)

        MessageBox.Show(
            arquivosCriados.Count.ToString() & " arquivo(s) Rhino criado(s) com sucesso." &
            Environment.NewLine & Environment.NewLine &
            "Apenas o primeiro arquivo foi aberto:" &
            Environment.NewLine &
            primeiroArquivo.Caminho,
            "Criar Rhino",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    End Sub

    Private Function ObterCodigoCincoDigitos(codigo As String) As String
        Dim encontrado As Match = Regex.Match(If(codigo, ""), "\d{5}")
        Return If(encontrado.Success, encontrado.Value, "")
    End Function

    Private Function PlanejarArquivos(pastaComponente As String,
                                      codigo As String,
                                      tipoServico As String,
                                      matrizes As List(Of MatrizAuxiliar)) As List(Of ArquivoRhinoPlanejado)

        Dim arquivos As New List(Of ArquivoRhinoPlanejado)()
        Dim ehEscala As Boolean = String.Equals(
            If(tipoServico, "").Trim(),
            "escala",
            StringComparison.OrdinalIgnoreCase
        )

        If Not ehEscala Then
            Dim arquivo As New ArquivoRhinoPlanejado()
            arquivo.Caminho = Path.Combine(pastaComponente, codigo & ".3dm")
            arquivo.Matrizes.AddRange(matrizes)
            arquivos.Add(arquivo)
            Return arquivos
        End If

        Dim matrizesPorNumero As New Dictionary(Of String, List(Of MatrizAuxiliar))(
            StringComparer.OrdinalIgnoreCase
        )
        Dim ordemNumeros As New List(Of String)()

        For Each matriz As MatrizAuxiliar In matrizes
            If matriz Is Nothing Then
                Continue For
            End If

            Dim numeros As List(Of String) = ModuloFuncoes.QuebrarNumerosPorVirgula(matriz.TextoNumeros)

            If numeros.Count = 0 Then
                Throw New Exception("Informe os números da escala para a matriz: " & matriz.TipoMatriz)
            End If

            For Each numero As String In numeros
                If Not matrizesPorNumero.ContainsKey(numero) Then
                    matrizesPorNumero.Add(numero, New List(Of MatrizAuxiliar)())
                    ordemNumeros.Add(numero)
                End If

                matrizesPorNumero(numero).Add(matriz)
            Next
        Next

        Dim pastasNumerosAusentes As New List(Of String)()
        For Each numero As String In ordemNumeros
            Dim pastaNumero As String = Path.Combine(pastaComponente, numero)
            If Not Directory.Exists(pastaNumero) Then pastasNumerosAusentes.Add(pastaNumero)
        Next

        If pastasNumerosAusentes.Count > 0 Then
            Dim resposta As DialogResult = MessageBox.Show(
                "As seguintes pastas de números não foram encontradas dentro do componente:" &
                Environment.NewLine & Environment.NewLine &
                String.Join(Environment.NewLine, pastasNumerosAusentes) &
                Environment.NewLine & Environment.NewLine &
                "Deseja criá-las para salvar os arquivos Rhino?",
                "Pastas dos números da escala",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            )

            If resposta <> DialogResult.Yes Then
                Throw New OperationCanceledException(
                    "Criação dos arquivos Rhino cancelada: as pastas dos números não foram criadas.")
            End If

            For Each pastaAusente As String In pastasNumerosAusentes
                Directory.CreateDirectory(pastaAusente)
            Next
        End If

        For Each numero As String In ordemNumeros
            Dim pastaNumero As String = Path.Combine(pastaComponente, numero)

            Dim arquivo As New ArquivoRhinoPlanejado()
            arquivo.Caminho = Path.Combine(pastaNumero, codigo & "_" & numero & ".3dm")
            arquivo.Matrizes.AddRange(matrizesPorNumero(numero))
            arquivos.Add(arquivo)
        Next

        Return arquivos
    End Function

    Private Sub ConfirmarSubstituicaoEExecutarBackups(arquivos As List(Of ArquivoRhinoPlanejado))
        Dim existentes As New List(Of ArquivoRhinoPlanejado)()

        For Each arquivo As ArquivoRhinoPlanejado In arquivos
            If File.Exists(arquivo.Caminho) Then
                existentes.Add(arquivo)
            End If
        Next

        If existentes.Count = 0 Then
            Return
        End If

        Dim resposta As DialogResult = MessageBox.Show(
            existentes.Count.ToString() & " arquivo(s) Rhino já existe(m)." &
            Environment.NewLine & Environment.NewLine &
            "Deseja criar cópias de segurança e substituí-los?",
            "Criar Rhino",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        )

        If resposta <> DialogResult.Yes Then
            Throw New OperationCanceledException("Criação dos arquivos Rhino cancelada.")
        End If

        For Each arquivo As ArquivoRhinoPlanejado In existentes
            Dim pasta As String = Path.GetDirectoryName(arquivo.Caminho)
            Dim nomeSemExtensao As String = Path.GetFileNameWithoutExtension(arquivo.Caminho)
            Dim backup As String = Path.Combine(
                pasta,
                nomeSemExtensao & ".backup_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".3dm"
            )
            File.Copy(arquivo.Caminho, backup, False)
        Next
    End Sub

    Private Function LocalizarExecutavelRhino() As String
        Dim candidatos As New List(Of String)()
        Dim config As ConfiguracaoAuxiliar = ModuloFuncoes.CarregarConfiguracao()

        If Not String.IsNullOrWhiteSpace(config.CaminhoRhino) AndAlso File.Exists(config.CaminhoRhino) Then
            Return config.CaminhoRhino
        End If

        Dim arquivosProgramas As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        Dim arquivosProgramasX86 As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
        Dim arquivosProgramas64 As String = Environment.GetEnvironmentVariable("ProgramW6432")

        ' Um executável AnyCPU pode iniciar em 32 bits. Nesse caso, ProgramFiles
        ' aponta para Program Files (x86), mas ProgramW6432 continua apontando
        ' para a pasta de programas 64 bits onde Rhino 7 e Rhino 8 são instalados.
        If String.IsNullOrWhiteSpace(arquivosProgramas64) Then
            arquivosProgramas64 = arquivosProgramas
        End If

        ' No computador de desenvolvimento, aceita Rhino 8. Na empresa, encontra Rhino 7.
        candidatos.Add(Path.Combine(arquivosProgramas64, "Rhino 8", "System", "Rhino.exe"))
        candidatos.Add(Path.Combine(arquivosProgramas64, "Rhino 7", "System", "Rhino.exe"))
        candidatos.Add(Path.Combine(arquivosProgramas, "Rhino 8", "System", "Rhino.exe"))
        candidatos.Add(Path.Combine(arquivosProgramas, "Rhino 7", "System", "Rhino.exe"))
        candidatos.Add(Path.Combine(arquivosProgramasX86, "Rhino 8", "System", "Rhino.exe"))
        candidatos.Add(Path.Combine(arquivosProgramasX86, "Rhino 7", "System", "Rhino.exe"))

        For Each candidato As String In candidatos
            If File.Exists(candidato) Then
                Return candidato
            End If
        Next

        Return SelecionarExecutavelRhino(config)
    End Function

    Private Function SelecionarExecutavelRhino(config As ConfiguracaoAuxiliar) As String
        Using dialogo As New OpenFileDialog()
            dialogo.Title = "Rhino não encontrado — selecione o arquivo Rhino.exe"
            dialogo.Filter = "Executável do Rhino (Rhino.exe)|Rhino.exe|Executáveis (*.exe)|*.exe"
            dialogo.CheckFileExists = True
            dialogo.Multiselect = False
            dialogo.FileName = "Rhino.exe"

            If dialogo.ShowDialog() <> DialogResult.OK Then Return ""

            If Not String.Equals(Path.GetFileName(dialogo.FileName), "Rhino.exe", StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show(
                    "Selecione o arquivo Rhino.exe.",
                    "Criar Rhino",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                Return ""
            End If

            config.CaminhoRhino = dialogo.FileName
            ModuloFuncoes.SalvarConfiguracao(config)
            MessageBox.Show(
                "Executável do Rhino registrado para as próximas utilizações.",
                "Criar Rhino",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            Return dialogo.FileName
        End Using
    End Function

    Private Function MontarScript(arquivos As List(Of ArquivoRhinoPlanejado),
                                  caminhoStatus As String) As String

        Dim sb As New StringBuilder()
        sb.AppendLine("# -*- coding: utf-8 -*-")
        sb.AppendLine("import traceback")
        sb.AppendLine("import Rhino")
        sb.AppendLine("import rhinoscriptsyntax as rs")
        sb.AppendLine("from System import Guid")
        sb.AppendLine("from System.Drawing import Color")
        sb.AppendLine("from System.IO import File")
        sb.AppendLine("from System.Text import Encoding")
        sb.AppendLine()
        sb.AppendLine("arquivo_status = u'" & EscaparPython(caminhoStatus) & "'")
        sb.AppendLine()
        sb.AppendLine("def adicionar(documento, nome, pai=Guid.Empty, cor=Color.Black, expandida=False):")
        sb.AppendLine("    camada = Rhino.DocObjects.Layer()")
        sb.AppendLine("    camada.Name = nome")
        sb.AppendLine("    camada.Color = cor")
        sb.AppendLine("    camada.ParentLayerId = pai")
        sb.AppendLine("    camada.IsVisible = True")
        sb.AppendLine("    camada.IsExpanded = expandida")
        sb.AppendLine("    indice = documento.Layers.Add(camada)")
        sb.AppendLine("    if indice < 0:")
        sb.AppendLine("        raise Exception(u'Não foi possível criar o layer: ' + nome)")
        sb.AppendLine("    return documento.Layers[indice].Id")
        sb.AppendLine()
        sb.AppendLine("try:")

        For i As Integer = 0 To arquivos.Count - 1
            AdicionarArquivoAoScript(sb, arquivos(i), i)
        Next

        sb.AppendLine("    File.WriteAllText(arquivo_status, u'OK', Encoding.UTF8)")
        sb.AppendLine("except Exception:")
        sb.AppendLine("    erro = traceback.format_exc()")
        sb.AppendLine("    File.WriteAllText(arquivo_status, erro, Encoding.UTF8)")
        sb.AppendLine("finally:")
        sb.AppendLine("    Rhino.RhinoApp.Exit(False)")
        Return sb.ToString()
    End Function

    Private Sub AdicionarArquivoAoScript(sb As StringBuilder,
                                     arquivo As ArquivoRhinoPlanejado,
                                     indiceArquivo As Integer)

        Dim prefixo As String = "a" & indiceArquivo.ToString()

        sb.AppendLine("    documento_" & prefixo & " = Rhino.RhinoDoc.CreateHeadless(None)")
        sb.AppendLine("    indice_padrao_" & prefixo & " = documento_" & prefixo & ".Layers.CurrentLayerIndex")
        sb.AppendLine("    camada_padrao_" & prefixo & " = documento_" & prefixo & ".Layers[indice_padrao_" & prefixo & "]")
        sb.AppendLine("    camada_padrao_" & prefixo & ".Name = u'Default'")
        sb.AppendLine("    camada_padrao_" & prefixo & ".IsExpanded = True")
        sb.AppendLine("    camada_padrao_" & prefixo & ".CommitChanges()")
        sb.AppendLine("    id_padrao_" & prefixo & " = camada_padrao_" & prefixo & ".Id")

        Dim variaveisTipo As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim texturasPorTipo As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
        Dim contadorTipo As Integer = 0
        Dim contadorTextura As Integer = 0

        For Each matriz As MatrizAuxiliar In arquivo.Matrizes
            If matriz Is Nothing OrElse String.IsNullOrWhiteSpace(matriz.TipoMatriz) Then
                Continue For
            End If

            Dim tipo As String = NomeLayerTipo(matriz.TipoMatriz)
            Dim variavelTipo As String = "tipo_" & prefixo & "_" & contadorTipo.ToString()

            If variaveisTipo.ContainsKey(tipo) Then
                variavelTipo = variaveisTipo(tipo)
            Else
                variaveisTipo.Add(tipo, variavelTipo)
                texturasPorTipo.Add(tipo, New List(Of String)())

                sb.AppendLine("    " & variavelTipo & " = adicionar(documento_" & prefixo & ", u'" & EscaparPython(tipo) & "', id_padrao_" & prefixo & ", expandida=True)")

                Dim variavelMatriz As String = "matriz_" & prefixo & "_" & contadorTipo.ToString()
                sb.AppendLine("    " & variavelMatriz & " = adicionar(documento_" & prefixo & ", u'Matriz', " & variavelTipo & ")")

                contadorTipo += 1
            End If

            Dim texturas As List(Of TexturaAuxiliar) = ModuloFuncoes.LerTexturas(matriz.TextoTexturas)

            For Each textura As TexturaAuxiliar In texturas
                Dim nomeTextura As String = textura.NomePasta

                If texturasPorTipo(tipo).Contains(nomeTextura) Then
                    Continue For
                End If

                texturasPorTipo(tipo).Add(nomeTextura)

                Dim variavelTextura As String = "textura_" & prefixo & "_" & contadorTextura.ToString()

                sb.AppendLine("    " & variavelTextura & " = adicionar(documento_" & prefixo & ", u'" & EscaparPython(nomeTextura) & "', " & variavelTipo & ", expandida=True)")
                sb.AppendLine("    adicionar(documento_" & prefixo & ", u'Map', " & variavelTextura & ", Color.FromArgb(0, 255, 0))")
                sb.AppendLine("    adicionar(documento_" & prefixo & ", u'Faces', " & variavelTextura & ")")

                contadorTextura += 1
            Next
        Next

        sb.AppendLine("    opcoes_" & prefixo & " = Rhino.FileIO.FileWriteOptions()")
        sb.AppendLine("    opcoes_" & prefixo & ".FileVersion = 7")
        sb.AppendLine("    opcoes_" & prefixo & ".SuppressDialogBoxes = True")
        sb.AppendLine("    if not documento_" & prefixo & ".WriteFile(u'" & EscaparPython(arquivo.Caminho) & "', opcoes_" & prefixo & "):")
        sb.AppendLine("        raise Exception(u'Falha ao gravar: " & EscaparPython(arquivo.Caminho) & "')")
        sb.AppendLine("    documento_" & prefixo & ".Dispose()")
    End Sub

    Private Function NomeLayerTipo(tipo As String) As String
        Select Case ModuloFuncoes.LimparNomePasta(tipo)
            Case "gravacao"
                Return "Gravação"
            Case "tampa"
                Return "Tampa"
            Case "laterais"
                Return "Laterais"
            Case "ferramentas"
                Return "Ferramentas"
            Case "lfe", "lfd", "lme", "lmd"
                Return tipo.Trim().ToUpperInvariant()
            Case Else
                Return tipo.Trim()
        End Select
    End Function

    Private Function EscaparPython(texto As String) As String
        If texto Is Nothing Then
            Return ""
        End If

        Return texto.Replace("\", "\\").Replace("'", "\'")
    End Function

End Module
