Attribute VB_Name = "Mod_Sistema"
Option Explicit

' Estructura SHFILEOPSTRUCT o para usar con el Api
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
  
'Declaración Api SHFileOperation
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                                (lpFileOp As SHFILEOPSTRUCT) As Long
  
'Constantes
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40
  
  
' Subrutina que copia el archivo
Public Sub Copiar_Archivo(ByVal Origen As String, ByVal Destino As String)
  
Dim t_Op As SHFILEOPSTRUCT
  
    With t_Op
        .hWnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
  
    ' Se ejecuta la función Api pasandole la estructura
    SHFileOperation t_Op
      
     Kill App.Path & "\Registrador\Config.ini"
     
     Call Ordena
      
End Sub

Sub RevisaRegistro()
 
Dim NumRegistro As Long

NumRegistro = Val(GetVar(App.Path & "\Registrador\Config.ini", "Config", "Registros"))
NumReg = NumRegistro
 
 If NumRegistro > 0 Then
 Call Copiar_Archivo(App.Path & "\Registrador\Config.Ini", App.Path & "\Registrador\Envia.ini")
 Else
 End
 End If
 
End Sub

Sub Ordena()
 Call DatosRegistro
End Sub

Sub DatosRegistro()
    If ValReg = 0 Then
         ValReg = 1
    End If
    
    If ValReg > NumReg Then
        ValReg = 0
        Kill App.Path & "\Registrador\Envia.ini"
        Call RevisaRegistro
     Exit Sub
    End If
    
        Dim i As Integer
    
    For i = 1 To 5
    If i = 1 Then
    Nombre = GetVar(App.Path & "\Registrador\Envia.ini", "User" & ValReg, "Nombre")
    ElseIf i = 2 Then
    Password = GetVar(App.Path & "\Registrador\Envia.ini", "User" & ValReg, "Password")
    ElseIf i = 3 Then
    Email = GetVar(App.Path & "\Registrador\Envia.ini", "User" & ValReg, "Email")
    ElseIf i = 4 Then
    Clase = GetVar(App.Path & "\Registrador\Envia.ini", "User" & ValReg, "Clase")
    ElseIf i = 5 Then
    Raza = GetVar(App.Path & "\Registrador\Envia.ini", "User" & ValReg, "Raza")
    End If
    Next i
    
   Form1.Timer1.Enabled = True
   
End Sub

Sub EnviaRegistro()
   Call Form1.EnviaMail
End Sub
