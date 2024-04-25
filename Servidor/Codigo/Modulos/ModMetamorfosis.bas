Attribute VB_Name = "ModMetamorfosis"
Option Explicit

Public Sub DoLicantropo(ByVal UserIndex As Integer)
        
    With UserList(UserIndex)
             
        If UCase$(UserList(UserIndex).Raza) <> "LICANTROPO" Then Exit Sub
              
        If .flags.MetamorfosisLicantropo = 0 Then
           
           If .flags.Navegando = 1 Then Exit Sub
                 
            .flags.OldBody = .char.Body
            .flags.OldHead = .char.Head
            .char.Body = 173
            .char.Head = 0
            .char.WeaponAnim = NingunArma
            .char.Alas = NingunAlas
            .char.CascoAnim = NingunCasco
            .char.ShieldAnim = NingunEscudo
            
            .flags.MetamorfosisLicantropo = 1
            
        Else
            
            .char.Body = .flags.OldBody
            .char.Head = .flags.OldHead

            If .Invent.WeaponEqpObjIndex > 0 Then .char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
            If .Invent.CascoEqpObjIndex > 0 Then .char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
            If .Invent.EscudoEqpObjIndex > 0 Then .char.ShieldAnim = ObjData(.Invent.CascoEqpObjIndex).ShieldAnim
            If .Invent.AlaEqpObjIndex > 0 Then .char.Alas = ObjData(.Invent.AlaEqpObjIndex).Ropaje
                 
            .flags.MetamorfosisLicantropo = 0
                 
        End If
              
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim, UserList(UserIndex).char.Alas)
              
    End With
        
End Sub
