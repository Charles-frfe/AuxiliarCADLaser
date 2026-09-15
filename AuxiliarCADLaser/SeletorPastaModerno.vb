Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Module SeletorPastaModerno

    Private Const FOS_PICKFOLDERS As UInteger = &H20UI
    Private Const FOS_FORCEFILESYSTEM As UInteger = &H40UI
    Private Const FOS_PATHMUSTEXIST As UInteger = &H800UI
    Private Const FOS_NOCHANGEDIR As UInteger = &H8UI
    Private Const SIGDN_FILESYSPATH As UInteger = &H80058000UI
    Private Const HRESULT_CANCELADO As Integer = -2147023673

    Public Function Selecionar(titulo As String, pastaInicial As String) As String
        Dim dialogo As IFileDialog = Nothing
        Dim pastaInicialShell As IShellItem = Nothing
        Dim resultadoShell As IShellItem = Nothing

        Try
            Dim tipoDialogo As Type = Type.GetTypeFromCLSID(
                New Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")
            )

            dialogo = CType(Activator.CreateInstance(tipoDialogo), IFileDialog)

            Dim opcoes As UInteger = 0UI
            VerificarHResult(dialogo.GetOptions(opcoes))
            opcoes = opcoes Or FOS_PICKFOLDERS Or FOS_FORCEFILESYSTEM Or
                     FOS_PATHMUSTEXIST Or FOS_NOCHANGEDIR
            VerificarHResult(dialogo.SetOptions(opcoes))

            VerificarHResult(dialogo.SetTitle(titulo))
            VerificarHResult(dialogo.SetOkButtonLabel("Selecionar pasta"))

            If Not String.IsNullOrWhiteSpace(pastaInicial) AndAlso Directory.Exists(pastaInicial) Then
                Dim iidShellItem As New Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")
                Dim hrPasta As Integer = SHCreateItemFromParsingName(
                    pastaInicial,
                    IntPtr.Zero,
                    iidShellItem,
                    pastaInicialShell
                )

                If hrPasta = 0 AndAlso pastaInicialShell IsNot Nothing Then
                    ' Se a rede estiver temporariamente indisponível, o seletor
                    ' ainda deve abrir em vez de voltar para uma janela antiga.
                    dialogo.SetFolder(pastaInicialShell)
                End If
            End If

            Dim janelaPai As IntPtr = IntPtr.Zero

            If Form.ActiveForm IsNot Nothing Then
                janelaPai = Form.ActiveForm.Handle
            End If

            Dim hrExibir As Integer = dialogo.Show(janelaPai)

            If hrExibir = HRESULT_CANCELADO Then
                Return ""
            End If

            VerificarHResult(hrExibir)
            VerificarHResult(dialogo.GetResult(resultadoShell))

            Dim ponteiroCaminho As IntPtr = IntPtr.Zero

            Try
                VerificarHResult(resultadoShell.GetDisplayName(SIGDN_FILESYSPATH, ponteiroCaminho))

                Dim caminho As String = Marshal.PtrToStringUni(ponteiroCaminho)

                If Not String.IsNullOrWhiteSpace(caminho) AndAlso Directory.Exists(caminho) Then
                    Return caminho
                End If
            Finally
                If ponteiroCaminho <> IntPtr.Zero Then
                    Marshal.FreeCoTaskMem(ponteiroCaminho)
                End If
            End Try

            Return ""
        Finally
            LiberarCom(resultadoShell)
            LiberarCom(pastaInicialShell)
            LiberarCom(dialogo)
        End Try
    End Function

    Private Sub VerificarHResult(hResult As Integer)
        If hResult < 0 Then
            Marshal.ThrowExceptionForHR(hResult)
        End If
    End Sub

    Private Sub LiberarCom(objeto As Object)
        If objeto Is Nothing Then
            Return
        End If

        Try
            If Marshal.IsComObject(objeto) Then
                Marshal.FinalReleaseComObject(objeto)
            End If
        Catch
        End Try
    End Sub

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=True)>
    Private Function SHCreateItemFromParsingName(
        <MarshalAs(UnmanagedType.LPWStr)> caminho As String,
        contexto As IntPtr,
        ByRef iid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef item As IShellItem
    ) As Integer
    End Function

    <ComImport(), Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IFileDialog
        <PreserveSig()> Function Show(janelaPai As IntPtr) As Integer
        <PreserveSig()> Function SetFileTypes(quantidade As UInteger, tipos As IntPtr) As Integer
        <PreserveSig()> Function SetFileTypeIndex(indice As UInteger) As Integer
        <PreserveSig()> Function GetFileTypeIndex(ByRef indice As UInteger) As Integer
        <PreserveSig()> Function Advise(eventos As IntPtr, ByRef cookie As UInteger) As Integer
        <PreserveSig()> Function Unadvise(cookie As UInteger) As Integer
        <PreserveSig()> Function SetOptions(opcoes As UInteger) As Integer
        <PreserveSig()> Function GetOptions(ByRef opcoes As UInteger) As Integer
        <PreserveSig()> Function SetDefaultFolder(pasta As IShellItem) As Integer
        <PreserveSig()> Function SetFolder(pasta As IShellItem) As Integer
        <PreserveSig()> Function GetFolder(ByRef pasta As IShellItem) As Integer
        <PreserveSig()> Function GetCurrentSelection(ByRef item As IShellItem) As Integer
        <PreserveSig()> Function SetFileName(<MarshalAs(UnmanagedType.LPWStr)> nome As String) As Integer
        <PreserveSig()> Function GetFileName(ByRef nome As IntPtr) As Integer
        <PreserveSig()> Function SetTitle(<MarshalAs(UnmanagedType.LPWStr)> titulo As String) As Integer
        <PreserveSig()> Function SetOkButtonLabel(<MarshalAs(UnmanagedType.LPWStr)> texto As String) As Integer
        <PreserveSig()> Function SetFileNameLabel(<MarshalAs(UnmanagedType.LPWStr)> texto As String) As Integer
        <PreserveSig()> Function GetResult(ByRef item As IShellItem) As Integer
        <PreserveSig()> Function AddPlace(item As IShellItem, local As Integer) As Integer
        <PreserveSig()> Function SetDefaultExtension(<MarshalAs(UnmanagedType.LPWStr)> extensao As String) As Integer
        <PreserveSig()> Function Close(hResult As Integer) As Integer
        <PreserveSig()> Function SetClientGuid(ByRef guidCliente As Guid) As Integer
        <PreserveSig()> Function ClearClientData() As Integer
        <PreserveSig()> Function SetFilter(filtro As IntPtr) As Integer
    End Interface

    <ComImport(), Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IShellItem
        <PreserveSig()> Function BindToHandler(
            contexto As IntPtr,
            ByRef identificador As Guid,
            ByRef iid As Guid,
            ByRef objeto As IntPtr
        ) As Integer

        <PreserveSig()> Function GetParent(ByRef pastaPai As IShellItem) As Integer
        <PreserveSig()> Function GetDisplayName(tipoNome As UInteger, ByRef nome As IntPtr) As Integer
        <PreserveSig()> Function GetAttributes(mascara As UInteger, ByRef atributos As UInteger) As Integer
        <PreserveSig()> Function Compare(outroItem As IShellItem, dica As UInteger, ByRef ordem As Integer) As Integer
    End Interface

End Module
