Attribute VB_Name = "Mod_TCP2"
Option Explicit

Sub HandleData2(ByVal rData As String)
   
    Dim Rs        As Integer

    Dim LoopC     As Integer

    Dim charindex As Long

    Dim X         As Integer
   
    Select Case UCase$(Left$(rData, 2))

        Case "XN"             '>>>>>> Coge información de quest NPC
            rData = Right$(rData, Len(rData) - 2)
            charindex = ReadField(1, rData, 44)
            CharList(charindex).NpcType = ReadField(2, rData, 44)
            Exit Sub
       
        Case "XU"          '>>>>>>> Coge datos de quest usuario y abre frmquest
            rData = Right$(rData, Len(rData) - 2)
            Quest.NumQuests = ReadField(1, rData, 44)
         
            For LoopC = 1 To NumQuests
                Quest.InfoUser.UserQuest(LoopC) = ReadField(LoopC, rData, 44)
            Next LoopC
            
            frmQuest.Show , frmMain
            Exit Sub

    End Select
   
    Select Case UCase$(Left$(rData, 3))
       
        Case "VPA"
           
            rData = Right$(rData, Len(rData) - 3)
            Rs = Val(ReadField(1, rData, 44))
           
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
              
                frmParty.CmdSalir.Visible = True
              
            End If
           
            frmParty.Show , frmMain
           
            Exit Sub
       
        Case "IVP"
           
            rData = Right$(rData, Len(rData) - 3)
            Rs = Val(ReadField(1, rData, 44))
           
            PartyData(Rs).Name = ReadField(2, rData, 44)
            PartyData(Rs).MinHP = Val(ReadField(3, rData, 44))
            PartyData(Rs).MaxHP = Val(ReadField(4, rData, 44))
           
            MaxVerParty = Rs
           
            Exit Sub
       
        Case "VPT"
            rData = Right$(rData, Len(rData) - 3)
             
            charindex = Val(ReadField(1, rData, 44))
            
            CharList(charindex).Stats.MinHP = Val(ReadField(2, rData, 44))
            CharList(charindex).Stats.MaxHP = Val(ReadField(3, rData, 44))
            CharList(charindex).PartyIndex = Val(ReadField(4, rData, 44))
            Exit Sub
             
    End Select
   
    Select Case UCase$(Left$(rData, 4))
   
        Case "HUCT"
            rData = Right$(rData, Len(rData) - 4)
            TimeChange = rData
            Call DayNameChange(rData)
            Exit Sub
   
        Case "VLDB"
            frmValidarBanco.Show , frmMain
            Exit Sub
             
        Case "BANP"
            rData = Right$(rData, Len(rData) - 4)
            
            frmBancoInfo.LblBanco = ReadField(1, rData, 44)
            frmBancoInfo.LblOro = ReadField(2, rData, 44)
            frmBancoInfo.LblObj = ReadField(3, rData, 44)
            
            frmBancoInfo.Show , frmMain
            Exit Sub
       
        Case "BAND"
            rData = Right$(rData, Len(rData) - 4)
           
            frmBancoDepositar.LblOro = rData
           
            frmBancoDepositar.Show , frmMain
            Exit Sub
        
        Case "BANF"
            rData = Right$(rData, Len(rData) - 4)
           
            frmBancoFinal.LblBanco = ReadField(1, rData, 44)
            frmBancoFinal.LblOro = ReadField(2, rData, 44)
           
            frmBancoFinal.Show , frmMain
            Exit Sub
           
        Case "BANR"
            rData = Right$(rData, Len(rData) - 4)
           
            frmBancoRetirar.LblBanco = rData
           
            frmBancoRetirar.Show , frmMain
            Exit Sub
        
        Case "HECA"
            frmOlvidarHechizo.Show , frmMain
            Exit Sub
        
        Case "LSTH"
            rData = Right$(rData, Len(rData) - 4)
           
            frmOlvidarHechizo.List1.AddItem rData
            Exit Sub
        
        Case "ABRC"
          
            rData = Right$(rData, Len(rData) - 4)
            Rs = ReadField(1, rData, 64)
            rData = ReadField(2, rData, 64)
          
            ReDim Heads(1 To Rs)
            frmCabezas.List1.Clear
          
            For LoopC = 1 To Rs
                frmCabezas.List1.AddItem "Cabeza" & LoopC
                Heads(LoopC) = ReadField(LoopC, rData, 44)
            Next LoopC
          
            frmCabezas.Show , frmMain
            Exit Sub
        
        Case "LSTS"
            rData = Right$(rData, Len(rData) - 4)
          
            ObjSastre(NumSastre) = ReadField(1, rData, 64)
            frmSastre.List1.AddItem ReadField(2, rData, 64)
            NumSastre = NumSastre + 1
          
            Exit Sub
          
        Case "ABRS"
            frmSastre.Show , frmMain
            Exit Sub
        
        Case "OBJH"
            rData = Right$(rData, Len(rData) - 4)
          
            ObjHechizeria(NumHechizeria) = ReadField(1, rData, 64)
            frmHechiceria.List1.AddItem ReadField(2, rData, 64)
            NumHechizeria = NumHechizeria + 1
            Exit Sub
        
        Case "ABRH"
            frmHechiceria.Show , frmMain
            Exit Sub
           
        Case "OBHM"
            rData = Right$(rData, Len(rData) - 4)
           
            ObjHerreroMagico(NumHerrero) = ReadField(1, rData, 64)
            frmHerreroMagico.List1.AddItem ReadField(2, rData, 64)
            NumHerrero = NumHerrero + 1
            Exit Sub
      
        Case "ABHM"
            frmHerreroMagico.Show , frmMain
            Exit Sub
          
    End Select

End Sub
