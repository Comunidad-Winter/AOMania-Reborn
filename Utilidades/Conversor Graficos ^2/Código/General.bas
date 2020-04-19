Attribute VB_Name = "General"
'Option Explicit
  
Public Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long, ByVal wIDNewItem As Long, _
ByVal lpString As Any) As Long
Public Declare Function DrawMenuBar Lib "user32" _
(ByVal hWnd As Long) As Long
'
Public Const MF_BYCOMMAND = &H0&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
'
Public Const SC_CLOSE = &HF060&

'Constante para pasar que indica que _
 se abre el archivo en modo lectura
Public Const OF_READ = &H0&
  
' Api lOpen para abrir un archivo
Public Declare Function lOpen Lib "kernel32" Alias "_lopen" ( _
                ByVal lpPathName As String, _
                ByVal iReadWrite As Long) As Long
  
' Api lclose para cerrar el archivo
Public Declare Function lclose Lib "kernel32" Alias "_lclose" ( _
                ByVal hFile As Long) As Long
  
' Api GetFileSize para averiguar el tamaño
Public Declare Function GetFileSize Lib "kernel32" ( _
                ByVal hFile As Long, _
                lpFileSizeHigh As Long) As Long
  
Dim lpFSHigh As Long
  
' recupera el tamaño de todos los archivos _
  que cuelgan del directorio ( no incluye subdirectorios )
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function Tamaño(Numero As Long)
  
Dim Handle As Long
Dim Len_Archivo As Long
Dim Archivo As String
Dim Path As String

Path = frmMain.txtDirInicial

    ' verifica la barra separadora de path
    If Mid(Path, Len(Path), 1) <> "\" Then
       Path = Path & "\"
    End If
      
    'Buscamos todos los archivos del directorio
    Archivo = Dir(Path & "\" & Numero & ".bmp")
  
    While Archivo <> ""
        'Mientras no sean directorios
        If Archivo <> "." Or Archivo <> ".." Then
  
            'En handle almacenamos el n° de identificador de archivo. _
            Si es -1 es por que dio error _
            Para abrirlo utilizamos lopen, con la ruta y el tipo de acceso
  
            Handle = lOpen(Path & Archivo, OF_READ)
  
            ' Vamos Almacenando el tamaño
            Len_Archivo = Len_Archivo + GetFileSize(Handle, lpFSHigh)
  
        End If
    ' Buscamos el siguiente archivo
    Archivo = Dir
    Wend
    
    Tamaño = Round(Len_Archivo / 1024)
    ' Cerramos el archivo abierto
    lclose Handle
  
End Function

Public Function ObtenerDimension(ByVal mDimension As Integer)
    If mDimension <= 32 Then
        ObtenerDimension = 32
    ElseIf mDimension <= 64 And mDimension > 32 Then
        ObtenerDimension = 64
    ElseIf mDimension <= 128 And mDimension > 64 Then
        ObtenerDimension = 128
    ElseIf mDimension <= 256 And mDimension > 128 Then
        ObtenerDimension = 256
    ElseIf mDimension <= 512 And mDimension > 256 Then
        ObtenerDimension = 512
    ElseIf mDimension <= 1024 And mDimension > 512 Then
        ObtenerDimension = 1024
    ElseIf mDimension <= 2048 And mDimension > 1024 Then
        ObtenerDimension = 2048
    End If
    
    ObtenerDimension = ObtenerDimension + 4
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Function EsPotenciaDos(ByVal N As Integer) As Boolean
Dim i As Byte
For i = 1 To 10
    If 2 ^ i = N Then EsPotenciaDos = True: Exit Function
Next i
End Function
