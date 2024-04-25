Attribute VB_Name = "EspadaSagrada"
Option Explicit

Public Const ObjEspadaNormal As Integer = 1064
Public Const ObjEspadaAse    As Integer = 1707
Public Const ObjArcoNormal   As Integer = 1001
Public Const ObjVaraNormal   As Integer = 1114
Private Const Hacha          As Integer = 3

Public Function EspadaSagrada(ObjIndex As Integer)
  
    Select Case ObjIndex
      
        Case ObjEspadaNormal
            EspadaSagrada = True
            Exit Function
      
        Case ObjEspadaAse
            EspadaSagrada = True
            Exit Function
       
        Case ObjArcoNormal
            EspadaSagrada = True
            Exit Function
        
        Case ObjVaraNormal
            EspadaSagrada = True
            Exit Function
                 
    End Select
   
    EspadaSagrada = False

End Function

Sub ChangeSagradaHit(UserIndex As Integer)
         
    With UserList(UserIndex)
        
        If .Sagrada.Enabled = 0 Then
            .Sagrada.Enabled = 1
        End If
        
        If .Invent.WeaponEqpObjIndex = ObjEspadaNormal Then
         
            Select Case UCase$(.Clase)
         
                Case "GUERRERO"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 10
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 14
                        .Sagrada.MinHit = 16
                    ElseIf .Stats.ELV <= 29 Then
                        .Sagrada.MaxHit = 22
                        .Sagrada.MinHit = 20
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 26
                        .Sagrada.MinHit = 25
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 29
                        .Sagrada.MinHit = 28
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 31
                        .Sagrada.MinHit = 33
                    End If
              
                Case "PALADIN"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 10
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 14
                        .Sagrada.MinHit = 16
                    ElseIf .Stats.ELV <= 29 Then
                        .Sagrada.MaxHit = 22
                        .Sagrada.MinHit = 20
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 26
                        .Sagrada.MinHit = 25
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 29
                        .Sagrada.MinHit = 28
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 31
                        .Sagrada.MinHit = 33
                    End If
              
                Case "LADRON"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 10
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 14
                        .Sagrada.MinHit = 16
                    ElseIf .Stats.ELV <= 29 Then
                        .Sagrada.MaxHit = 22
                        .Sagrada.MinHit = 20
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 26
                        .Sagrada.MinHit = 25
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 29
                        .Sagrada.MinHit = 28
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 31
                        .Sagrada.MinHit = 33
                    End If
              
                Case "CLERIGO"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 10
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 14
                        .Sagrada.MinHit = 16
                    ElseIf .Stats.ELV <= 29 Then
                        .Sagrada.MaxHit = 22
                        .Sagrada.MinHit = 20
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 26
                        .Sagrada.MinHit = 25
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 29
                        .Sagrada.MinHit = 28
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 31
                        .Sagrada.MinHit = 33
                    End If
              
                Case "BARDO"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 10
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 14
                        .Sagrada.MinHit = 16
                    ElseIf .Stats.ELV <= 29 Then
                        .Sagrada.MaxHit = 22
                        .Sagrada.MinHit = 20
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 26
                        .Sagrada.MinHit = 25
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 29
                        .Sagrada.MinHit = 28
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 31
                        .Sagrada.MinHit = 33
                    End If
              
                Case "DRUIDA"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 10
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 14
                        .Sagrada.MinHit = 16
                    ElseIf .Stats.ELV <= 29 Then
                        .Sagrada.MaxHit = 22
                        .Sagrada.MinHit = 20
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 26
                        .Sagrada.MinHit = 25
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 29
                        .Sagrada.MinHit = 28
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 31
                        .Sagrada.MinHit = 33
                    End If
                      
            End Select
          
        End If
          
        If .Invent.WeaponEqpObjIndex = ObjEspadaAse Then
               
            Select Case UCase(.Clase)
                 
                Case "ASESINO"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 10
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 14
                        .Sagrada.MinHit = 16
                    ElseIf .Stats.ELV <= 29 Then
                        .Sagrada.MaxHit = 22
                        .Sagrada.MinHit = 20
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 26
                        .Sagrada.MinHit = 25
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 29
                        .Sagrada.MinHit = 28
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 31
                        .Sagrada.MinHit = 33
                       End If
                 
            End Select
             
        End If
          
        If .Invent.WeaponEqpObjIndex = ObjArcoNormal Then
               
            Select Case UCase$(.Clase)

                Case "ARQUERO"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 5
                        .Sagrada.MinHit = 3
                    ElseIf .Stats.ELV <= 19 Then
                        .Sagrada.MaxHit = 9
                        .Sagrada.MinHit = 7
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 12
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 15
                        .Sagrada.MinHit = 13
                    ElseIf .Stats.ELV <= 39 Then
                        .Sagrada.MaxHit = 17
                        .Sagrada.MinHit = 15
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 20
                        .Sagrada.MinHit = 18
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 23
                        .Sagrada.MinHit = 21

                    End If
                  
                Case "CAZADOR"

                    If .Stats.ELV <= 14 Then
                        .Sagrada.MaxHit = 5
                        .Sagrada.MinHit = 3
                    ElseIf .Stats.ELV <= 19 Then
                        .Sagrada.MaxHit = 9
                        .Sagrada.MinHit = 7
                    ElseIf .Stats.ELV <= 24 Then
                        .Sagrada.MaxHit = 12
                        .Sagrada.MinHit = 10
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 15
                        .Sagrada.MinHit = 13
                    ElseIf .Stats.ELV <= 39 Then
                        .Sagrada.MaxHit = 17
                        .Sagrada.MinHit = 15
                    ElseIf .Stats.ELV <= 44 Then
                        .Sagrada.MaxHit = 20
                        .Sagrada.MinHit = 18
                    ElseIf .Stats.ELV <= 55 Then
                        .Sagrada.MaxHit = 23
                        .Sagrada.MinHit = 21

                    End If
               
            End Select

        End If
          
        If .Invent.WeaponEqpObjIndex = ObjVaraNormal Then
              
            Select Case UCase$(.Clase)

                Case "MAGO"

                    If .Stats.ELV <= 14 Then
                    ElseIf .Stats.ELV <= 24 Then
                    ElseIf .Stats.ELV <= 34 Then
                        .Sagrada.MaxHit = 30
                        .Sagrada.MinHit = 30
                    ElseIf .Stats.ELV <= 44 Then
                    ElseIf .Stats.ELV <= 55 Then

                    End If
                
                Case "BRUJO"

                    If .Stats.ELV <= 14 Then
                    ElseIf .Stats.ELV <= 24 Then
                    ElseIf .Stats.ELV <= 34 Then
                    ElseIf .Stats.ELV <= 44 Then
                    ElseIf .Stats.ELV <= 55 Then

                    End If
                  
            End Select
          
        End If
          
    End With
     
End Sub

Sub DeleteSagradaHit(UserIndex As Integer)

    With UserList(UserIndex)
       
        If .Sagrada.Enabled = 1 Then
            .Sagrada.MinHit = 0
            .Sagrada.MaxHit = 0
            .Sagrada.Enabled = 0

        End If
       
    End With

End Sub

Sub ConnectSagrada(UserIndex As Integer)
   Call ChangeSagradaHit(UserIndex)
End Sub
