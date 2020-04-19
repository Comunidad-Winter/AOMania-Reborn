Attribute VB_Name = "SistemaCriatura"
Option Explicit

Public NpcCriatura      As Integer
Public LoteriaCriatura  As Integer
Public DiasCriaturas    As Long
Public SistemaCriatura  As Boolean
Public ExpCriatura      As Boolean
Public OroCriatura      As Boolean
Public NombreCriatura   As String
Private ProximaCriatura As Long
Public RandomCase       As Integer
Public DiaEspecial      As Boolean
Public DiaEspecialExp   As Boolean
Public DiaEspecialOro   As Boolean

Sub DarExpNPC(npc As Integer, Exp As Integer)
     
    Dim LoopC As Integer
     
    For LoopC = 1 To LastNPC
     
        If Npclist(LoopC).Numero = npc Then
                
            Npclist(LoopC).GiveEXP = Npclist(LoopC).GiveEXP * Exp
              
        End If
     
    Next LoopC

End Sub

Sub DarOroNPC(npc As Integer, Oro As Integer)
     
    Dim LoopC As Integer
     
    For LoopC = 1 To LastNPC
     
        If Npclist(LoopC).Numero = npc Then
                
            Npclist(LoopC).GiveGLD = Npclist(LoopC).GiveGLD * Oro
              
        End If
     
    Next LoopC

End Sub

Sub QuitarExpNPC(npc As Integer, Exp As Integer)
     
    Dim LoopC As Integer
     
    For LoopC = 1 To LastNPC
     
        If Npclist(LoopC).Numero = npc Then
                
            Npclist(LoopC).GiveEXP = Npclist(LoopC).GiveEXP / Exp
              
        End If
     
    Next LoopC

End Sub

Sub QuitarOroNPC(npc As Integer, Oro As Integer)
     
    Dim LoopC As Integer
     
    For LoopC = 1 To LastNPC
     
        If Npclist(LoopC).Numero = npc Then
                
            Npclist(LoopC).GiveGLD = Npclist(LoopC).GiveGLD / Oro
              
        End If
     
    Next LoopC

End Sub

Sub Load_Criatura()
    DiasCriaturas = GetVar(DatPath & "\ini\sistemacriatura.ini", "Config", "Dias")

    If DiasCriaturas = 0 Then
        SistemaCriatura = False
    Else
        SistemaCriatura = True
        ProximaCriatura = 1

    End If

End Sub

Sub Save_Criatura()
    Call WriteVar(DatPath & "\ini\sistemacriatura.ini", "Config", "Dias", DiasCriaturas)

End Sub

Sub Timer_SistemaCriatura()

    If OnHor = 1 Then
        If SistemaCriatura = False And ProximaCriatura = 0 And DiasCriaturas = 0 Then
            RandomCase = RandomNumber(1, 15)
            Call CriaturasNormales(RandomCase)
            DiasCriaturas = DiasCriaturas + 1
            SistemaCriatura = True
            ProximaCriatura = 13
            Call Save_Criatura

        End If
                
        If SistemaCriatura = True And ProximaCriatura = 1 Then
            If DiasCriaturas = "20" Then
                Call DiaEspeciales
                ProximaCriatura = 13
            Else

                If DiaEspecialExp = True Then
                    Call QuitarDiaEspecial("Exp", "2")

                End If

                If DiaEspecialOro = True Then
                    Call QuitarDiaEspecial("Oro", "2")

                End If
             
                RandomCase = RandomNumber(1, 15)
                Call CriaturasNormales(RandomCase)
                DiasCriaturas = DiasCriaturas + 1
                SistemaCriatura = True
                Call Save_Criatura
                ProximaCriatura = 13

            End If

        End If
        
    End If
    
    If OnHor = 13 Then
        If SistemaCriatura = True And ProximaCriatura = 13 Then
            If DiasCriaturas = "20" Then
                Call DiaEspeciales
                ProximaCriatura = 1
            Else

                If DiaEspecialExp = True Then
                    Call QuitarDiaEspecial("Exp", "2")

                End If

                If DiaEspecialOro = True Then
                    Call QuitarDiaEspecial("Oro", "2")

                End If

                RandomCase = RandomNumber(1, 15)
                Call CriaturasNormales(RandomCase)
                DiasCriaturas = DiasCriaturas + 1
                SistemaCriatura = True
                Call Save_Criatura
                ProximaCriatura = 1

            End If

        End If

    End If

End Sub

Sub CriaturasNormales(EligeDia As Integer)
     
    If ExpCriatura = True Then
        DoEvents
        Call QuitarExpNPC(NpcCriatura, LoteriaCriatura)
        DoEvents
        ExpCriatura = False

    End If
          
    If OroCriatura = True Then
        DoEvents
        Call QuitarOroNPC(NpcCriatura, LoteriaCriatura)
        DoEvents
        OroCriatura = False

    End If
         
    Select Case EligeDia
         
        Case "1"
            NpcCriatura = 688
            LoteriaCriatura = 2
            NombreCriatura = "Ent"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "2"
            NpcCriatura = 613
            LoteriaCriatura = 3
            NombreCriatura = "Planta Carnivora"
            Call DarOroNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su oro ha sido aumentada x" & LoteriaCriatura & "." & _
                    FONTTYPE_TALK)
            OroCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "3"
            NpcCriatura = 691
            LoteriaCriatura = 2
            NombreCriatura = "Alma Infernal"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "4"
            NpcCriatura = 587
            LoteriaCriatura = 2
            NombreCriatura = "Hada"
            Call DarOroNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su oro ha sido aumentada x" & LoteriaCriatura & "." & _
                    FONTTYPE_TALK)
            OroCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "5"
            NpcCriatura = 553
            LoteriaCriatura = 3
            NombreCriatura = "Medusa"
            Call DarOroNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su oro ha sido aumentada x" & LoteriaCriatura & "." & _
                    FONTTYPE_TALK)
            OroCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "6"
            NpcCriatura = 553
            LoteriaCriatura = 2
            NombreCriatura = "Medusa"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "7"
            NpcCriatura = 655
            LoteriaCriatura = 2
            NombreCriatura = "Viuda Amarilla"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "8"
            NpcCriatura = 651
            LoteriaCriatura = 3
            NombreCriatura = "Viuda Verde"
            Call DarOroNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su oro ha sido aumentada x" & LoteriaCriatura & "." & _
                    FONTTYPE_TALK)
            OroCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "9"
            NpcCriatura = 652
            LoteriaCriatura = 3
            NombreCriatura = "Viuda Azul"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "10"
            NpcCriatura = 634
            LoteriaCriatura = 2
            NombreCriatura = "Hombre de las nieves"
            Call DarOroNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su oro ha sido aumentada x" & LoteriaCriatura & "." & _
                    FONTTYPE_TALK)
            OroCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "11"
            NpcCriatura = 538
            LoteriaCriatura = 2
            NombreCriatura = "Oso Pardo"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "12"
            NpcCriatura = 633
            LoteriaCriatura = 3
            NombreCriatura = "Sirena"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "13"
            NpcCriatura = 574
            LoteriaCriatura = 2
            NombreCriatura = "Momia"
            Call DarOroNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su oro ha sido aumentada x" & LoteriaCriatura & "." & _
                    FONTTYPE_TALK)
            OroCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "14"
            NpcCriatura = 633
            LoteriaCriatura = 2
            NombreCriatura = "Momia"
            Call DarExpNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su experencia ha sido aumentada x" & LoteriaCriatura & _
                    "." & FONTTYPE_TALK)
            ExpCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "15"
            NpcCriatura = 567
            LoteriaCriatura = 3
            NombreCriatura = "Golem de Oro"
            Call DarOroNPC(NpcCriatura, LoteriaCriatura)
            Call SendData(SendTarget.toall, 0, 0, "||Hoy es día de " & NombreCriatura & ", su oro ha sido aumentada x" & LoteriaCriatura & "." & _
                    FONTTYPE_TALK)
            OroCriatura = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
    End Select

End Sub

Private Sub DiaEspeciales()
    Dim EligeEspecial As Integer
    DiasCriaturas = 0
    Call Save_Criatura
    EligeEspecial = RandomNumber(1, 2)

    Select Case EligeEspecial
          
        Case "1"
              
            Call SendData(SendTarget.toall, 0, 0, "||¡Estáis de suerte! Día especial, la experencia ha sido aumentada por x2" & FONTTYPE_TALK)
            Call DarDiaEspecial("Exp", "2")
            LoteriaCriatura = 2
            DiaEspecialExp = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
            
        Case "2"
            Call SendData(SendTarget.toall, 0, 0, "||¡Estáis de suerte! Día especial, el oro ha sido aumentada por x2" & FONTTYPE_TALK)
            Call DarDiaEspecial("Oro", "2")
            LoteriaCriatura = 2
            DiaEspecialOro = True
            Call SendData(SendTarget.toall, 0, 0, "TW55")
          
    End Select

End Sub

Private Sub DarDiaEspecial(Tipo As String, Cantidad As Integer)
    Dim LoopC As Integer
    
    Select Case Tipo
    
        Case "Exp"
            
            For LoopC = 1 To NumNPCs
                 
                Npclist(LoopC).GiveEXP = Npclist(LoopC).GiveEXP * Cantidad
                    
            Next LoopC
             
        Case "Oro"

            For LoopC = 1 To NumNPCs
                   
                Npclist(LoopC).GiveGLD = Npclist(LoopC).GiveGLD * Cantidad
              
            Next LoopC
    
    End Select
        
End Sub

Private Sub QuitarDiaEspecial(Tipo As String, Cantidad As Integer)
    Dim LoopC As Integer
    
    Select Case Tipo
    
        Case "Exp"
            
            For LoopC = 1 To NumNPCs
                 
                Npclist(LoopC).GiveEXP = Npclist(LoopC).GiveEXP / Cantidad
                    
            Next LoopC

            DiaEspecialExp = False
             
        Case "Oro"

            For LoopC = 1 To NumNPCs
                   
                Npclist(LoopC).GiveGLD = Npclist(LoopC).GiveGLD / Cantidad
              
            Next LoopC

            DiaEspecialOro = False
    
    End Select
        
End Sub
