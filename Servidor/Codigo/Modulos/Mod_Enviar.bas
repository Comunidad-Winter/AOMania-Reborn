Attribute VB_Name = "Mod_Enviar"
Option Explicit

Sub EnviaRegistro(Name As String, Password As String, Email As String, Clase As String, Raza As String)
    Dim NumRegistros As String
   
    NumRegistros = val(GetVar(App.Path & "\Registrador\Config.ini", "Config", "Registros"))
   
    NumRegistros = NumRegistros + 1
   
    Call WriteVar(App.Path & "\Registrador\Config.ini", "Config", "Registros", NumRegistros)
   
    Call WriteVar(App.Path & "\Registrador\Config.ini", "User" & NumRegistros, "Nombre", Name)
    Call WriteVar(App.Path & "\Registrador\Config.ini", "User" & NumRegistros, "Password", Password)
    Call WriteVar(App.Path & "\Registrador\Config.ini", "User" & NumRegistros, "Email", Email)
    Call WriteVar(App.Path & "\Registrador\Config.ini", "User" & NumRegistros, "Clase", Clase)
    Call WriteVar(App.Path & "\Registrador\Config.ini", "User" & NumRegistros, "Raza", Raza)
   
    Call Shell(App.Path & "\Registrador.exe", vbNormalFocus)
   
End Sub
