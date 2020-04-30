Attribute VB_Name = "Mod_TCP2"
Option Explicit

Sub HandleData2(ByVal Rdata As String)
   
    Dim Rs        As Integer

    Dim LoopC     As Integer

    Dim charindex As Long

    Dim X         As Integer
   
    Select Case UCase$(Left$(Rdata, 2))

        Case "XN"             '>>>>>> Coge información de quest NPC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(1, Rdata, 44)
            CharList(charindex).NpcType = ReadField(2, Rdata, 44)
            Exit Sub
       
        Case "XU"          '>>>>>>> Coge datos de quest usuario y abre frmquest
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Quest.NumQuests = ReadField(1, Rdata, 44)
         
            For LoopC = 1 To NumQuests
                Quest.InfoUser.UserQuest(LoopC) = ReadField(LoopC, Rdata, 44)
            Next LoopC
            
            frmQuest.Show , frmMain
            Exit Sub
        
        Case "XP"       '>>>>>>> Actualiza el proceso de la quest
           Rdata = Right$(Rdata, Len(Rdata) - 2)
           
           charindex = ReadField(1, Rdata, 44)
           ProcesoQuest = Val(ReadField(2, Rdata, 44))
           
           Exit Sub
           
        Case "XI"    '>>>>>>> Actualiza icono npc misiones
           Rdata = Right$(Rdata, Len(Rdata) - 2)
           
           charindex = ReadField(1, Rdata, 44)
           CharList(charindex).Icono = ReadField(2, Rdata, 44)
        Exit Sub

    End Select
   
    Select Case UCase$(Left$(Rdata, 3))
       
        Case "VPA"
           
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Rs = Val(ReadField(1, Rdata, 44))
           
            If Rs = 0 Then
                frmParty.Label1.Visible = True
            ElseIf Rs = 1 Then
              
                For LoopC = 1 To MaxVerParty
              
                    frmParty.Label2(LoopC).Caption = PartyData(LoopC).Name
                    frmParty.Label2(LoopC).Visible = True
                    frmParty.Label3(LoopC).Caption = PartyData(LoopC).MinHP & "/" & PartyData(LoopC).MaxHP
                    frmParty.Label3(LoopC).Visible = True
                    frmParty.Label4(LoopC).Visible = True
                    frmParty.Shape1(LoopC).Visible = True
                  
                    If PartyData(LoopC).MinHP > 0 Then
                        frmParty.Label4(LoopC).Width = (((PartyData(LoopC).MinHP / 100) / (PartyData(LoopC).MaxHP / 100)) * 101)
                    Else
                        frmParty.Label4(LoopC).Width = 0

                    End If
                 
                Next LoopC
              
                frmParty.cmdSalir.Visible = True
              
            End If
           
            frmParty.Show , frmMain
           
            Exit Sub
       
        Case "IVP"
           
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Rs = Val(ReadField(1, Rdata, 44))
           
            PartyData(Rs).Name = ReadField(2, Rdata, 44)
            PartyData(Rs).MinHP = Val(ReadField(3, Rdata, 44))
            PartyData(Rs).MaxHP = Val(ReadField(4, Rdata, 44))
           
            MaxVerParty = Rs
           
            Exit Sub
       
        Case "VPT"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
             
            charindex = Val(ReadField(1, Rdata, 44))
            
            CharList(charindex).Stats.MinHP = Val(ReadField(2, Rdata, 44))
            CharList(charindex).Stats.MaxHP = Val(ReadField(3, Rdata, 44))
            CharList(charindex).PartyIndex = Val(ReadField(4, Rdata, 44))
            Exit Sub
             
    End Select
   
    Select Case UCase$(Left$(Rdata, 4))
   
        Case "HUCT"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            TimeChange = Rdata
            Call DayNameChange(Rdata)
            Exit Sub
   
        Case "VLDB"
            frmValidarBanco.Show , frmMain
            Exit Sub
             
        Case "BANP"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            
            frmBancoInfo.LblBanco = ReadField(1, Rdata, 44)
            frmBancoInfo.LblOro = ReadField(2, Rdata, 44)
            frmBancoInfo.LblObj = ReadField(3, Rdata, 44)
            
            frmBancoInfo.Show , frmMain
            Exit Sub
       
        Case "BAND"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
           
            frmBancoDepositar.LblOro = Rdata
           
            frmBancoDepositar.Show , frmMain
            Exit Sub
        
        Case "BANF"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
           
            frmBancoFinal.LblBanco = ReadField(1, Rdata, 44)
            frmBancoFinal.LblOro = ReadField(2, Rdata, 44)
           
            frmBancoFinal.Show , frmMain
            Exit Sub
           
        Case "BANR"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
           
            frmBancoRetirar.LblBanco = Rdata
           
            frmBancoRetirar.Show , frmMain
            Exit Sub
        
        Case "HECA"
            frmOlvidarHechizo.Show , frmMain
            Exit Sub
        
        Case "LSTH"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
           
            frmOlvidarHechizo.List1.AddItem Rdata
            Exit Sub
        
        Case "ABRC"
          
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Rs = ReadField(1, Rdata, 64)
            Rdata = ReadField(2, Rdata, 64)
          
            ReDim Heads(1 To Rs)
            frmCabezas.List1.Clear
          
            For LoopC = 1 To Rs
                frmCabezas.List1.AddItem "Cabeza" & LoopC
                Heads(LoopC) = ReadField(LoopC, Rdata, 44)
            Next LoopC
          
            frmCabezas.Show , frmMain
            Exit Sub
        
        Case "LSTS"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
          
            ObjSastre(NumSastre) = ReadField(1, Rdata, 64)
            frmSastre.List1.AddItem ReadField(2, Rdata, 64)
            NumSastre = NumSastre + 1
          
            Exit Sub
          
        Case "ABRS"
            frmSastre.Show , frmMain
            Exit Sub
        
        Case "OBJH"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
          
            ObjHechizeria(NumHechizeria) = ReadField(1, Rdata, 64)
            frmHechiceria.List1.AddItem ReadField(2, Rdata, 64)
            NumHechizeria = NumHechizeria + 1
            Exit Sub
        
        Case "ABRH"
            frmHechiceria.Show , frmMain
            Exit Sub
           
        Case "OBHM"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
           
            ObjHerreroMagico(NumHerrero) = ReadField(1, Rdata, 64)
            frmHerreroMagico.List1.AddItem ReadField(2, Rdata, 64)
            NumHerrero = NumHerrero + 1
            Exit Sub
      
        Case "ABHM"
            frmHerreroMagico.Show , frmMain
            Exit Sub
          
    End Select

End Sub
