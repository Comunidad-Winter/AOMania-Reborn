Attribute VB_Name = "AoDefenderExternos"
Private Declare Function FindWindow _
    Lib "user32" _
    Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
        Public AoDefDetectName As String

Private NameCheats As String

Public Function AoDefDetect() As Boolean
If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.1.1"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("ART-MONEY")) Then
    AoDefDetect = True
    NameCheats = "ART-MONEY"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.0"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CROWN MAKRO")) Then
     AoDefDetect = True
     NameCheats = "CROWN MAKRO"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("A TRABAJAR...")) Then
     AoDefDetect = True
     NameCheats = "A TRABAJAR..."
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Project1")) Then
     AoDefDetect = True
     NameCheats = "Project1"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("ews")) Then
    AoDefDetect = True
    NameCheats = "ews"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    AoDefDetect = True
    NameCheats = "Pts"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.2"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.6")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.6"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.7")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.7"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.8")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.8"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.9")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.9"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 6.0")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 6.0"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    AoDefDetect = True
    NameCheats = "SOLOCOVO?"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    AoDefDetect = True
    NameCheats = "-=[ANUBYS RADAR]=-"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    AoDefDetect = True
    NameCheats = "CRAZY SPEEDER 1.05"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    AoDefDetect = True
    NameCheats = "SET !XSPEED.NET"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    AoDefDetect = True
    NameCheats = "SPEEDERXP V1.80 - UNREGISTERED"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
     AoDefDetect = True
     NameCheats = "CHEAT ENGINE 5.3"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.4")) Then
    AoDefDetect = True
    NameCheats = "CHEAT ENGINE 5.4"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VI@HOTMAIL.COM>")) Then
    AoDefDetect = True
    NameCheats = "MACROCRACK <GONZA_VI@HOTMAIL.COM>"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VJ@HOTMAIL.COM>")) Then
   AoDefDetect = True
   NameCheats = "MACROCRACK <GONZA_VJ@HOTMAIL.COM>"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VI@HOTMAIL.COM>")) Then
     AoDefDetect = True
     NameCheats = "MACRO CRACK <GONZA_VI@HOTMAIL.COM>"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    AoDefDetect = True
    NameCheats = "MACRO CRACK <GONZA_VJ@HOTMAIL.COM>"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHITS")) Then
   AoDefDetect = True
   NameCheats = "CHITS"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
   AoDefDetect = True
   NameCheats = "CHEAT ENGINE 5.1"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    AoDefDetect = True
    NameCheats = "A SPEEDER"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    AoDefDetect = True
    NameCheats = "MEMO :P"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
     AoDefDetect = True
     NameCheats = "ORK4M VERSION 1.5"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("ORKAM")) Then
   AoDefDetect = True
   NameCheats = "ORKAM"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACRO")) Then
  AoDefDetect = True
  NameCheats = "MACRO"
   Exit Function
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    AoDefDetect = True
    NameCheats = "BY FEDEX"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
   AoDefDetect = True
   NameCheats = "!XSPEED.NET +4.59"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    AoDefDetect = True
    NameCheats = "CAMBIA TITULOS DE CHEATS BY FEDEX"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    AoDefDetect = True
    NameCheats = "NEWENG OCULTO"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
     AoDefDetect = True
     NameCheats = "SERBIO ENGINE"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
     AoDefDetect = True
     NameCheats = "REYMIX ENGINE 5.3 PUBLIC"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    AoDefDetect = True
    NameCheats = "REY ENGINE 5.2"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
     AoDefDetect = True
     NameCheats = "AUTOCLICK - BY NIO_SHOOTER"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    AoDefDetect = True
    NameCheats = "TONNER MINER! :D [REG][SKLOV] 2.0"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    AoDefDetect = True
    NameCheats = "Buffy The vamp Slayer"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    AoDefDetect = True
    NameCheats = "Blorb Slayer 1.12.552 (BETA)"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    AoDefDetect = True
    NameCheats = "PumaEngine3.0"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
     AoDefDetect = True
     NameCheats = "Vicious Engine 5.0"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
   AoDefDetect = True
   NameCheats = "AkumaEngine33"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    AoDefDetect = True
    NameCheats = "Spuc3ngine"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
     AoDefDetect = True
     NameCheats = "Ultra Engine"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
   AoDefDetect = True
   NameCheats = "Engine"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
     AoDefDetect = True
     NameCheats = "Cheat Engine V5.4"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
     AoDefDetect = True
     NameCheats = "Cheat Engine V4.4"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    AoDefDetect = True
    NameCheats = "Cheat Engine V4.4 German Add-On"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    AoDefDetect = True
    NameCheats = "Cheat Engine V4.3"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
     AoDefDetect = True
     NameCheats = "Cheat Engine V4.2"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
     AoDefDetect = True
     NameCheats = "Cheat Engine V4.1.1"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
     AoDefDetect = True
     NameCheats = "Cheat Engine V3.3"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    AoDefDetect = True
    NameCheats = "Cheat Engine V3.2"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    AoDefDetect = True
    NameCheats = "Cheat Engine V3.1"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    AoDefDetect = True
    NameCheats = "Cheat Engine"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
     AoDefDetect = True
     NameCheats = "danza engine 5.2.150"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
   AoDefDetect = True
   NameCheats = "zenx engine"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
     AoDefDetect = True
     NameCheats = "MACROMAKER"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    AoDefDetect = True
    NameCheats = "MACREOMAKER - EDIT MACRO"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
     AoDefDetect = True
     NameCheats = "By Fedex"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    AoDefDetect = True
    NameCheats = "Macro Mage 1.0"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 Pete Powa")) Then
     AoDefDetect = True
     NameCheats = "Auto* v0.4 (c) 2001 Pete Powa"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
     AoDefDetect = True
     NameCheats = "Kizsada"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    AoDefDetect = True
    NameCheats = "Makro K33"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
     AoDefDetect = True
     NameCheats = "Super Saiyan"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
     AoDefDetect = True
     NameCheats = "Makro-Piringulete"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    AoDefDetect = True
    NameCheats = "Makro-Piringulete 2003"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
   AoDefDetect = True
   NameCheats = "TUKY2005"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("Countach")) Then
    AoDefDetect = True
    NameCheats = "Countach"
     Exit Function
    ElseIf FindWindow(vbNullString, UCase$("MacroRecorder")) Then
    AoDefDetect = True
    NameCheats = "MacroRecorder"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Ultimatemacros")) Then
     AoDefDetect = True
     NameCheats = "Ultimatemacros"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("MacroLauncher")) Then
    AoDefDetect = True
    NameCheats = "MacroLauncher"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.5")) Then
    AoDefDetect = True
    NameCheats = "Cheat Engine 5.5"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Auto Remo- TheFrank^")) Then
    AoDefDetect = True
    NameCheats = "Auto Remo- TheFrank^"
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO")) Then
     AoDefDetect = True
     NameCheats = "WPE PRO"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - " & AoDefDetectName & ".exe")) Then
     AoDefDetect = True
     NameCheats = "WPE PRO - " & AoDefDetectName & ".exe"
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - [WPEPRO2]")) Then
     AoDefDetect = True
     NameCheats = "WPE PRO - [WPEPRO2]"
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO [WPEPRO2]")) Then
     AoDefDetect = True
     NameCheats = "WPE PRO [WPEPRO2]"
   Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - " & AoDefDetectName & ".exe" & " - [WPEPRO2]")) Then
  AoDefDetect = True
  NameCheats = "WPE PRO - " & AoDefDetectName & ".exe" & " - [WPEPRO2]"
  Exit Function
ElseIf FindWindow(vbNullString, UCase$("rPE - rEdoX Packet Editor")) Then
  AoDefDetect = True
  NameCheats = "rPE - rEdoX Packet Editor"
  Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 7.1")) Then
  AoDefDetect = True
  NameCheats = "Cheat Engine 7.1"
  Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheats " & Chr(34) & "AO" & Chr(34))) Then
  AoDefDetect = True
  NameCheats = "Cheats " & Chr(34) & "AO" & Chr(34)
  Exit Function
ElseIf FindWindow(vbNullString, UCase$("CCleaner ")) Then
  AoDefDetect = True
  NameCheats = "CCleaner"
  Exit Function
ElseIf FindWindow(vbNullString, UCase$("Mouse Recorder Pro - Untitled 3")) Then
  AoDefDetect = True
  NameCheats = "Mouse Recorder Pro"
  Exit Function
End If

AoDefDetect = False
End Function
Public Sub AoDefCheat()
    Call SendData("ANTICH" & NameCheats)
    MsgBox "Se ha detectado algo inusual en el cliente. Se va a cerrar por seguridad.", vbCritical, "AoMania"
End Sub

