Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class ProjetoEncontrado
    Public Property Caminho As String = ""
    Public Property NomePasta As String = ""
    Public Property DataAlteracao As DateTime = DateTime.MinValue
End Class

Public Module LocalizadorProjetos

    Private Const LimiteResultados As Integer = 100

    Public Function Localizar(codigo As String, pastaRaiz As String) As List(Of ProjetoEncontrado)
        Dim resultado As New List(Of ProjetoEncontrado)()
        codigo = If(codigo, "").Trim()

        If Not Regex.IsMatch(codigo, "^\d{5}$") OrElse
           String.IsNullOrWhiteSpace(pastaRaiz) OrElse
           Not Directory.Exists(pastaRaiz) Then
            Return resultado
        End If

        Dim vistos As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Procurar(pastaRaiz, codigo, resultado, vistos)
        resultado.Sort(
            Function(a As ProjetoEncontrado, b As ProjetoEncontrado)
                Return b.DataAlteracao.CompareTo(a.DataAlteracao)
            End Function
        )
        Return resultado
    End Function

    Private Sub Procurar(pastaAtual As String,
                         codigo As String,
                         encontrados As List(Of ProjetoEncontrado),
                         vistos As HashSet(Of String))
        If encontrados.Count >= LimiteResultados Then Return

        Dim caminhoCompleto As String
        Try
            caminhoCompleto = Path.GetFullPath(pastaAtual)
        Catch
            Return
        End Try

        If vistos.Contains(caminhoCompleto) Then Return
        vistos.Add(caminhoCompleto)

        Try
            Dim info As New DirectoryInfo(caminhoCompleto)
            Dim corresponde As Boolean = Regex.IsMatch(
                info.Name,
                "^" & Regex.Escape(codigo) & "(?:$|[ _-])",
                RegexOptions.IgnoreCase
            )

            If corresponde Then
                encontrados.Add(New ProjetoEncontrado() With {
                    .Caminho = info.FullName,
                    .NomePasta = info.Name,
                    .DataAlteracao = info.LastWriteTime
                })
                ' A pasta encontrada é o projeto. Não pesquisa suas subpastas.
                Return
            End If

            If (info.Attributes And FileAttributes.ReparsePoint) = FileAttributes.ReparsePoint Then Return

            For Each subpasta As DirectoryInfo In info.EnumerateDirectories()
                Procurar(subpasta.FullName, codigo, encontrados, vistos)
                If encontrados.Count >= LimiteResultados Then Exit For
            Next
        Catch ex As UnauthorizedAccessException
        Catch ex As IOException
        Catch ex As System.Security.SecurityException
        End Try
    End Sub

End Module
