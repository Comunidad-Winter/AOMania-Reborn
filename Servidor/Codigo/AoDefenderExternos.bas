Attribute VB_Name = "AoDefenderExternos"
Private Declare Function FindWindow _
    Lib "user32" _
    Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
        Public AoDefDetectName As String
Public Function AoDefDetect() As Boolean
If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("ART-MONEY")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CROWN MAKRO")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("A TRABAJAR...")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Project1")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("ews")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.6")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.7")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.8")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.9")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 6.0")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.4")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VI@HOTMAIL.COM>")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VJ@HOTMAIL.COM>")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VI@HOTMAIL.COM>")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHITS")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("ORKAM")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACRO")) Then
  AoDefDetect = True
   Exit Function
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 Pete Powa")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
   AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("Countach")) Then
    AoDefDetect = True
     Exit Function
    ElseIf FindWindow(vbNullString, UCase$("MacroRecorder")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Ultimatemacros")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("MacroLauncher")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.5")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("Auto Remo- TheFrank^")) Then
    AoDefDetect = True
     Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - " & AoDefDetectName & ".exe")) Then
     AoDefDetect = True
      Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - [WPEPRO2]")) Then
     AoDefDetect = True
    Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO [WPEPRO2]")) Then
     AoDefDetect = True
   Exit Function
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - " & AoDefDetectName & ".exe" & " - [WPEPRO2]")) Then
  AoDefDetect = True
  Exit Function
ElseIf FindWindow(vbNullString, UCase$("rPE - rEdoX Packet Editor")) Then
  AoDefDetect = True
  Exit Function
End If
AoDefDetect = False
End Function
Public Sub AoDefCheat()
    MsgBox "AoDefender ha detectado cheats tu cliente será cerrado.", vbCritical, "AoDefender"
End Sub


