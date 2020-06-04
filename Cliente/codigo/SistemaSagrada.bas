Attribute VB_Name = "SistemaSagrada"
Option Explicit

'Intervalo por sagrada

Public Const IntervaloSagrada = 3600

'Npc Yeti Sagrada Oscura

Public Const NpcYetiOscura = 623
Public StatusYetiOscura As Boolean
Public Const MapaYetiOscura = 151
Public RepiteInvoYetiOscura As Boolean
Public MataYetiOscura       As Boolean

'Npc Yeti Sagrada Normal

Public Const NpcYeti = 630
Public StatusYeti As Boolean
Public Const MapaYeti = 151
Public RepiteInvoYeti As Boolean
Public MataYeti       As Boolean

'Npc Cleopatra

Public Const NpcCleopatra = 635
Public StatusCleopatra As Boolean
Public Const MapaCleopatra = 153
Public RepiteInvoCleopatra As Boolean
Public MataCleopatra       As Boolean

'Npc Rey Scorpion

Public Const NpcReyScorpion = 590
Public StatusReyScorpion As Boolean
Public Const MapaReyScorpion = 153
Public RepiteInvoReyScorpion As Boolean
Public MataReyScorpion       As Boolean

'Npc Dark Seth

Public Const NpcDarkSeth = 616
Public StatusDarkSeth As Boolean
Public Const MapaDarkSeth = 158
Public RepiteInvoDarkSeth As Boolean
Public MataDarkSeth       As Boolean

'Npc Tiburon Blanco Sagrado

Public Const NpcTiburonBlanco = 640
Public StatusTiburonBlanco As Boolean
Public Const MapaTiburonBlanco = 146
Public RepiteInvoTiburonBlanco As Boolean
Public MataTiburonBlanco       As Boolean

'Npc Elficas

Public Const NpcElfica = 636
Public StatusElfica As Boolean
Public Const MapaElfica = 55
Public RepiteInvoElfica As Boolean
Public MataElfica       As Boolean

'Npc Gran Dragon Rojo

Public Const NpcGranDragonRojo = 622
Public StatusGranDragonRojo As Boolean
Public Const MapaGranDragonRojo = 82
Public RepiteInvoGranDragonRojo As Boolean
Public MataGranDragonRojo       As Boolean

Sub LoadSagradas()

    StatusYetiOscura = False
    RepiteInvoYetiOscura = False
    MataYetiOscura = False
          
    StatusYeti = False
    RepiteInvoYeti = False
    MataYeti = False
     
    StatusCleopatra = False
    RepiteInvoCleopatra = False
    MataCleopatra = False
     
    StatusReyScorpion = False
    RepiteInvoReyScorpion = False
    MataReyScorpion = False
     
    StatusDarkSeth = False
    RepiteInvoDarkSeth = False
    MataDarkSeth = False
     
    StatusElfica = False
    RepiteInvoElfica = False
    MataElfica = False
     
    StatusGranDragonRojo = False
    RepiteInvoGranDragonRojo = False
    MataGranDragonRojo = False
     
    StatusTiburonBlanco = False
    RepiteInvoTiburonBlanco = False
    MataTiburonBlanco = False
     
End Sub

Private Function RevMapa(Mapa As Integer, X As Integer, Y As Integer)

    If MapData(Mapa, X, Y).Blocked = 1 Then
        RevMapa = True
        Exit Function

    End If

    RevMapa = False

End Function

Sub SpawnSagrada(Sagrada As String)
       
    Dim npc             As String
    Dim PositionSagrada As WorldPos
       
    Dim CordX           As Integer
    Dim CordY           As Integer
       
    CordX = RandomNumber(13, 87)
    CordY = RandomNumber(13, 87)
              
    Select Case Sagrada
          
        Case "YetiOscura"
              
            If RevMapa(MapaYetiOscura, CordX, CordY) Then
                RepiteInvoYetiOscura = True
            Else
                RepiteInvoYetiOscura = False
                StatusYetiOscura = True
                  
                npc = NpcYetiOscura
                  
                PositionSagrada.Map = MapaYetiOscura
                PositionSagrada.X = CordX
                PositionSagrada.Y = CordY
                  
                Call SpawnNpc(npc, PositionSagrada, True, False)
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Yeti Sagrado Oscuro en Aomania." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")

            End If
                    
            Exit Sub
          
        Case "Yeti"
              
            If RevMapa(MapaYeti, CordX, CordY) Then
                RepiteInvoYeti = True
            Else
                RepiteInvoYeti = False
                StatusYeti = True
                  
                npc = NpcYeti
                  
                PositionSagrada.Map = MapaYeti
                PositionSagrada.X = CordX
                PositionSagrada.Y = CordY
                  
                Call SpawnNpc(npc, PositionSagrada, True, False)
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Yeti Sagrado en Aomania." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")

            End If
                    
            Exit Sub
          
        Case "Cleopatra"
              
            If RevMapa(MapaCleopatra, CordX, CordY) Then
                RepiteInvoCleopatra = True
            Else
                RepiteInvoCleopatra = False
                StatusCleopatra = True
                  
                npc = NpcCleopatra
                  
                PositionSagrada.Map = MapaCleopatra
                PositionSagrada.X = CordX
                PositionSagrada.Y = CordY
                  
                Call SpawnNpc(npc, PositionSagrada, True, False)
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Cleopatra Sagrada en Aomania." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")

            End If
                    
            Exit Sub
          
        Case "ReyScorpion"
              
            If RevMapa(MapaReyScorpion, CordX, CordY) Then
                RepiteInvoReyScorpion = True
            Else
                RepiteInvoReyScorpion = False
                StatusReyScorpion = True
                  
                npc = NpcReyScorpion
                  
                PositionSagrada.Map = MapaReyScorpion
                PositionSagrada.X = CordX
                PositionSagrada.Y = CordY
                  
                Call SpawnNpc(npc, PositionSagrada, True, False)
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Rey Scorpion Sagrado en Aomania." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")

            End If
                    
            Exit Sub
          
        Case "DarkSeth"
              
            If RevMapa(MapaDarkSeth, CordX, CordY) Then
                RepiteInvoDarkSeth = True
            Else
                RepiteInvoDarkSeth = False
                StatusDarkSeth = True
                  
                npc = NpcDarkSeth
                  
                PositionSagrada.Map = MapaDarkSeth
                PositionSagrada.X = CordX
                PositionSagrada.Y = CordY
                  
                Call SpawnNpc(npc, PositionSagrada, True, False)
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Dark Seth Sagrado en Aomania." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")

            End If
                    
            Exit Sub
          
        Case "Elfica"
              
            If RevMapa(MapaElfica, CordX, CordY) Then
                RepiteInvoElfica = True
            Else
                RepiteInvoElfica = False
                StatusElfica = True
                  
                npc = NpcElfica
                  
                PositionSagrada.Map = MapaElfica
                PositionSagrada.X = CordX
                PositionSagrada.Y = CordY
                  
                Call SpawnNpc(npc, PositionSagrada, True, False)
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio hada Elfica Sagrada en Aomania." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")

            End If
                    
            Exit Sub
          
        Case "GranDragonRojo"
              
            If RevMapa(MapaGranDragonRojo, CordX, CordY) Then
                RepiteInvoGranDragonRojo = True
            Else
                RepiteInvoGranDragonRojo = False
                StatusGranDragonRojo = True
                  
                npc = NpcGranDragonRojo
                  
                PositionSagrada.Map = MapaGranDragonRojo
                PositionSagrada.X = CordX
                PositionSagrada.Y = CordY
                  
                Call SpawnNpc(npc, PositionSagrada, True, False)
                Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Gran Dragon Rojo Sagrado en Aomania." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toall, 0, 0, "TW3")

            End If
                    
            Exit Sub
          
        Case "TiburonBlanco"
              
            If RevMapa(MapaTiburonBlanco, CordX, CordY) Then
                RepiteInvoTiburonBlanco = True
            Else

                If HayAgua(MapaTiburonBlanco, CordX, CordY) Then
                    RepiteInvoTiburonBlanco = False
                    StatusTiburonBlanco = True
                  
                    npc = NpcTiburonBlanco
                  
                    PositionSagrada.Map = MapaTiburonBlanco
                    PositionSagrada.X = CordX
                    PositionSagrada.Y = CordY
                  
                    Call SpawnNpc(npc, PositionSagrada, True, False)
                    Call SendData(SendTarget.toall, 0, 0, "||Reaparecio Tiburon Blanco Sagrado en Aomania." & FONTTYPE_GUILD)
                    Call SendData(SendTarget.toall, 0, 0, "TW3")
                Else
                    RepiteInvoTiburonBlanco = True

                End If

            End If
                    
            Exit Sub
          
    End Select
       
End Sub
