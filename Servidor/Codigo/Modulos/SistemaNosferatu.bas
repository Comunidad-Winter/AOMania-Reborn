Attribute VB_Name = "SistemaNosferatu"
'|| Modulo: Sistema de Nosferatu
'|| Programado por: Bassinger
'|| 2019
'|| AoMania

Option Explicit

Public Const NpcNosfe     As Integer = 662
Public IntervaloNosfe     As Long
Public IntervaloMsjNosfe  As Long
Public StatusNosfe        As Boolean
Public RepiteInvoNosfe    As Boolean
Public MapaNosfe          As Integer
Public CordNosfeX         As Integer
Public CordNosfeY         As Integer
Public NickMataNosfe      As String
Public MataNosfe          As Boolean
Public Const ExpMataNosfe As Long = 500000
Public AvisoNosfe         As Boolean
Public TimeAvisoNosfe     As Long

Sub LoadNosfe()
    IntervaloNosfe = GetVar(App.Path & "\Dat\Ini\Nosferatu.ini", "Config", "IntervaloNosfe")
    IntervaloMsjNosfe = GetVar(App.Path & "\Dat\Ini\Nosferatu.ini", "Config", "IntervaloMsjNosfe")
    
    StatusNosfe = False
    RepiteInvoNosfe = False
    MataNosfe = False
    
    If IntervaloNosfe >= 1200 Then
        TimeAvisoNosfe = IntervaloNosfe - 600
        AvisoNosfe = True
    Else
        AvisoNosfe = False

    End If
    
End Sub

Sub InvocaNosfe()
         
    Dim NumMaps As Long
    Dim pos     As String
   
    DoEvents
   
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
   
    MapaNosfe = RandomNumber(1, NumMaps)
    CordNosfeX = RandomNumber(13, 87)
    CordNosfeY = RandomNumber(13, 87)
   
    If MapInfo(MapaNosfe).Zona = "CAMPO" Then
   
        If HayAgua(MapaNosfe, CordNosfeX, CordNosfeY) Then
           
            RepiteInvoNosfe = True

        Else
         
            If RestringeMapaNosfe(MapaNosfe) Then
          
                RepiteInvoNosfe = True
          
            Else

                If MapData(MapaNosfe, CordNosfeX, CordNosfeY).Blocked = "0" Then
                    Call SendData(SendTarget.toall, 0, 0, "||Nosferatu ha aparecido en el mapa " & MapaNosfe & FONTTYPE_GUILD)
                    Call SendData(SendTarget.toall, 0, 0, "TW107")
         
                    pos = MapaNosfe & CordNosfeX & CordNosfeY
                    StatusNosfe = True
                    RepiteInvoNosfe = False
       
                    Dim npc           As String
                    Dim PositionNosfe As WorldPos
       
                    npc = NpcNosfe
       
                    PositionNosfe.Map = MapaNosfe
                    PositionNosfe.X = CordNosfeX
                    PositionNosfe.Y = CordNosfeY
       
                    Call SpawnNpc(npc, PositionNosfe, True, False)
       
                Else
       
                    RepiteInvoNosfe = True
       
                End If
       
            End If
      
        End If
    
    Else
    
        RepiteInvoNosfe = True
      
    End If
   
End Sub

Function RestringeMapaNosfe(ByVal Mapa As Long)
      
    Select Case Mapa
        
        Case "0"
            RestringeMapaNosfe = True
            Exit Function
        
        Case "14"
            RestringeMapaNosfe = True
            Exit Function
        
        Case "22"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "47"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "48"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "57"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "66"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "67"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "68"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "69"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "70"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "71"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "72"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "73"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "74"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "75"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "76"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "77"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "78"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "79"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "80"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "83"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "85"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "87"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "88"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "89"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "90"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "91"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "92"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "93"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "103"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "105"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "106"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "107"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "108"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "109"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "110"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "111"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "112"
            RestringeMapaNosfe = True
            Exit Function
            
        Case "114"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "115"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "116"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "117"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "118"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "119"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "128"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "129"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "130"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "131"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "133"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "134"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "135"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "136"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "137"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "138"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "144"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "145"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "146"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "147"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "150"
            RestringeMapaNosfe = True
            Exit Function
       
        Case "154"
            RestringeMapaNosfe = True
            Exit Function
       
        Case "157"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "159"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "160"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "161"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "162"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "163"
            RestringeMapaNosfe = True
            Exit Function
       
        Case "165"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "166"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "167"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "169"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "170"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "171"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "173"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "175"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "176"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "177"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "178"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "179"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "180"
            RestringeMapaNosfe = True
            Exit Function
         
        Case "190"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "191"
            RestringeMapaNosfe = True
            Exit Function
      
        Case "192"
            RestringeMapaNosfe = True
            Exit Function
         
    End Select
     
    RestringeMapaNosfe = False

End Function
