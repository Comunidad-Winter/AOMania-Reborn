Attribute VB_Name = "Trabajo"
Option Explicit

Private Const IntervaloOculto As Integer = 5500 ' el tiempo de aca se divide x 40ms = 12 Sec +-

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    With UserList(UserIndex)
        .Counters.Ocultando = .Counters.Ocultando - 1
        Call SendData(SendTarget.toIndex, UserIndex, 0, "INVI" & .Counters.Ocultando)

        If .Counters.Ocultando <= 0 Then
        
            .Counters.Ocultando = 0
            .flags.Oculto = 0
            Call SendData(SendTarget.toIndex, UserIndex, 0, "INVI0")

            If .flags.Invisible = 0 Then
                'no hace falta encriptar este (se jode el gil que bypassea esto)
                Call SendData(SendTarget.ToMap, 0, .pos.Map, "NOVER" & .char.CharIndex & ",0," & .PartyIndex)
     
                Call SendData(SendTarget.toIndex, UserIndex, 0, "Z11")

            End If
        
        End If

    End With

    Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")

End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    Dim Skill     As Byte
    Dim res       As Integer
    Dim Suerte    As Double
    Dim SegOculto As Double
    Dim Intervalo As Integer
    
    With UserList(UserIndex)
        Intervalo = "61,2"
        Skill = .Stats.UserSkills(eSkill.Ocultarse)

        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        If Skill = 0 Then
            SegOculto = 0
        Else
            SegOculto = Intervalo * Skill

        End If
        
        res = RandomNumber(1, Suerte)

        If res <= 5 Then
            .flags.Oculto = 1
            
            'Suerte = (-0.000001 * (100 - Skill) ^ 3)
            'Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            'Suerte = Suerte + (-0.0088 * (100 - Skill))
            'Suerte = Suerte + (0.9571)
            'Suerte = Suerte * IntervaloOculto
            
            .Counters.Ocultando = SegOculto 'CInt(Suerte)

            Call SendData(SendTarget.ToMap, 0, .pos.Map, "NOVER" & .char.CharIndex & ",1," & .PartyIndex)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Te has escondido entre las sombras!" & FONTTYPE_INFO)
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No has logrado esconderte!" & FONTTYPE_INFO)
                .flags.UltimoMensaje = 4

            End If

            '[/CDT]
        End If

    End With

    Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

    Dim ModNave As Long
    Dim ObjIndex As Integer
    
    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If ObjData(ObjIndex).Real = 1 Or ObjData(ObjIndex).Caos = 1 Or ObjData(ObjIndex).Nemes = 1 Or ObjData(ObjIndex).Templ = 1 Then
        If Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu faccion no puede usar este objeto." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If Not UseRangeFragata(UserIndex, ObjIndex) Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Tu rango aún no te permite usar ese item." & FONTTYPE_INFO)
        Exit Sub
    End If

    With UserList(UserIndex)
        ModNave = ModNavegacion(.Clase)

        If HayAgua(.pos.Map, .pos.X - 1, .pos.Y) And HayAgua(.pos.Map, .pos.X + 1, .pos.Y) And HayAgua(.pos.Map, .pos.X, .pos.Y - 1) And HayAgua( _
            .pos.Map, .pos.X, .pos.Y + 1) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes dejar de navegar en el agua!!" & FONTTYPE_INFO)
            Exit Sub

        End If

        If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
            'Call SendData(SendTarget.toindex, UserIndex, 0, "||No tenes suficientes conocimientos para usar este barco." & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & _
                " puntos en navegacion." & FONTTYPE_INFO)
            Exit Sub

        End If

        .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
        .Invent.BarcoSlot = Slot

        If .flags.Navegando = 0 Then
    
            .char.Head = 0
    
            If .flags.Muerto = 0 Then
                .char.Body = Barco.Ropaje
            Else
                .char.Body = iFragataFantasmal

            End If
    
            .char.ShieldAnim = NingunEscudo
            .char.WeaponAnim = NingunArma
            .char.CascoAnim = NingunCasco
    
            '[MaTeO 9]
            .char.Alas = NingunAlas
            '[/MaTeO 9]
    
            .flags.Navegando = 1
    
        Else
    
            .flags.Navegando = 0
    
            If .flags.Muerto = 0 Then
                .char.Head = .OrigChar.Head
        
                If .Invent.ArmourEqpObjIndex > 0 Then
                    .char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                Else
                    Call DarCuerpoDesnudo(UserIndex)

                End If
        
                If .Invent.EscudoEqpObjIndex > 0 Then .char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

                If .Invent.WeaponEqpObjIndex > 0 Then .char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim

                If .Invent.CascoEqpObjIndex > 0 Then .char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
            Else
                .char.Body = iCuerpoMuerto
                .char.Head = iCabezaMuerto
                .char.ShieldAnim = NingunEscudo
                .char.WeaponAnim = NingunArma
                .char.CascoAnim = NingunCasco
                '[MaTeO 9]
                .char.Alas = NingunAlas

                '[/MaTeO 9]
            End If

        End If

        '[MaTeO 9]
        Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .char.Body, .char.Head, .char.Heading, .char.WeaponAnim, .char.ShieldAnim, _
            .char.CascoAnim, .char.Alas)
                
        '[/MaTeO 9]
        Call SendData(SendTarget.toIndex, UserIndex, 0, "NAVEG")

    End With

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)

    'Call LogTarea("Sub FundirMineral")

    If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
        If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList( _
            UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList( _
            UserIndex).Clase) Then
            Call DoLingotes(UserIndex)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tenes conocimientos de mineria suficientes para trabajar este mineral." & _
                FONTTYPE_INFO)

        End If

    End If

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
                      
    'Call LogTarea("Sub TieneObjetos")

    Dim i     As Integer
    Dim Total As Long

    For i = 1 To MAX_INVENTORY_SLOTS

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount

        End If

    Next i

    If Cant <= Total Then
        TieneObjetos = True
        Exit Function

    End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
    'Call LogTarea("Sub QuitarObjetos")

    Dim i As Integer

    For i = 1 To MAX_INVENTORY_SLOTS

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        
            Call Desequipar(UserIndex, i)
        
            UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - Cant

            If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
                Cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
                UserList(UserIndex).Invent.Object(i).Amount = 0
                UserList(UserIndex).Invent.Object(i).ObjIndex = 0
            Else
                Cant = 0

            End If
        
            Call UpdateUserInv(False, UserIndex, i)
        
            If (Cant = 0) Then
                QuitarObjetos = True
                Exit Function

            End If

        End If

    Next i

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)

    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)

End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)

End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
        If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tenes suficientes madera." & FONTTYPE_INFO)
            CarpinteroTieneMateriales = False
            Exit Function

        End If

    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean

    If ObjData(ItemIndex).LingH > 0 Then
        If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tenes suficientes lingotes de hierro." & FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Exit Function

        End If

    End If

    If ObjData(ItemIndex).LingP > 0 Then
        If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tenes suficientes lingotes de plata." & FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Exit Function

        End If

    End If

    If ObjData(ItemIndex).LingO > 0 Then
        If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tenes suficientes lingotes de oro." & FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Exit Function

        End If

    End If

    HerreroTieneMateriales = True

End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean

    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData( _
        ItemIndex).SkHerreria

End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean

    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)

        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(ArmadurasHerrero)

        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    PuedeConstruirHerreria = False

End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    'Call LogTarea("Sub HerreroConstruirItem")
    If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
        Call HerreroQuitarMateriales(UserIndex, ItemIndex)

        ' AGREGAR FX
        If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has construido el arma!." & FONTTYPE_INFO)
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has construido el escudo!." & FONTTYPE_INFO)
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has construido el casco!." & FONTTYPE_INFO)
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has construido la armadura!." & FONTTYPE_INFO)

        End If

        Dim MiObj As Obj
        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

        End If

        Call SubirSkill(UserIndex, Herreria)
        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & MARTILLOHERRERO)
    
    End If

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean

    Dim i As Long

    For i = 1 To UBound(ObjCarpintero)

        If ObjCarpintero(i) = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function

        End If

    Next i

    PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If CarpinteroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData( _
        ItemIndex).SkCarpinteria And PuedeConstruirCarpintero(ItemIndex) And UserList(UserIndex).Invent.HerramientaEqpObjIndex = _
        SERRUCHO_CARPINTERO Then

        Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has construido el objeto!" & FONTTYPE_INFO)
    
        Dim MiObj As Obj
        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

        End If
    
        Call SubirSkill(UserIndex, Carpinteria)
        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & LABUROCARPINTERO)

    End If

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 13

        Case iMinerales.PlataCruda
            MineralesParaLingote = 25

        Case iMinerales.OroCrudo
            MineralesParaLingote = 50

        Case Else
            MineralesParaLingote = 10000

    End Select

End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)

    '    Call LogTarea("Sub DoLingotes")
    Dim Slot As Integer
    Dim obji As Integer

    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(UserIndex).Invent.Object(Slot).Amount < MineralesParaLingote(obji) Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No tienes suficientes minerales para hacer un lingote." & FONTTYPE_INFO)
        Exit Sub

    End If
    
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - MineralesParaLingote(obji)

    If UserList(UserIndex).Invent.Object(Slot).Amount < 1 Then
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0

    End If

    Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has obtenido un lingote!!!" & FONTTYPE_INFO)
    Dim nPos  As WorldPos
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

    End If

    Call UpdateUserInv(False, UserIndex, Slot)
    Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Has obtenido un lingote!" & FONTTYPE_INFO)

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End Sub

Function ModNavegacion(ByVal Clase As String) As Integer

    Select Case UCase$(Clase)

        Case "PIRATA"
            ModNavegacion = 1

        Case "TRABAJADOR"
            ModNavegacion = 1.2

        Case Else
            ModNavegacion = 2.3

    End Select

End Function

Function ModFundicion(ByVal Clase As String) As Integer

    Select Case UCase$(Clase)

        Case "TRABAJADOR"
            ModFundicion = 1

        Case Else
            ModFundicion = 3

    End Select

End Function

Function ModCarpinteria(ByVal Clase As String) As Integer

    Select Case UCase$(Clase)

        Case "TRABAJADOR"
            ModCarpinteria = 1

        Case Else
            ModCarpinteria = 3

    End Select

End Function

Function ModHerreriA(ByVal Clase As String) As Integer

    Select Case UCase$(Clase)

        Case "TRABAJADOR"
            ModHerreriA = 1

        Case Else
            ModHerreriA = 4

    End Select

End Function

Function ModDomar(ByVal Clase As String) As Integer

    Select Case UCase$(Clase)

        Case "DRUIDA"
            ModDomar = 6

        Case "CAZADOR"
            ModDomar = 6

        Case "CLERIGO"
            ModDomar = 7

        Case Else
            ModDomar = 10

    End Select

End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long

    With UserList(UserIndex).Stats
        CalcularPoderDomador = .UserAtributos(eAtributos.Carisma) * (.UserSkills(eSkill.Domar) / ModDomar(UserList(UserIndex).Clase)) + _
            RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) + RandomNumber(1, _
            .UserAtributos(eAtributos.Carisma) / 3)

    End With

End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function

        End If

    Next j

End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    'Call LogTarea("Sub DoDomar")

    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
    
        If Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La criatura ya te ha aceptado como su amo." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La criatura ya tiene amo." & FONTTYPE_INFO)
            Exit Sub

        End If
    
        If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(UserIndex) Then
            Dim Index As Integer
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            Index = FreeMascotaIndex(UserIndex)
            UserList(UserIndex).MascotasIndex(Index) = NpcIndex
            UserList(UserIndex).MascotasType(Index) = Npclist(NpcIndex).Numero
        
            Npclist(NpcIndex).MaestroUser = UserIndex
        
            Call FollowAmo(NpcIndex)
        
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||La criatura te ha aceptado como su amo." & FONTTYPE_INFO)
            Call SubirSkill(UserIndex, Domar)
        Else

            If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||No has logrado domar la criatura." & FONTTYPE_INFO)
                UserList(UserIndex).flags.UltimoMensaje = 5

            End If

        End If

    Else
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No podes controlar mas criaturas." & FONTTYPE_INFO)

    End If

End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .flags.AdminInvisible = 0 Then
        
            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .char.Body = .CharMimetizado.Body
                .char.Head = .CharMimetizado.Head
                .char.CascoAnim = .CharMimetizado.CascoAnim
                .char.ShieldAnim = .CharMimetizado.ShieldAnim
                .char.WeaponAnim = .CharMimetizado.WeaponAnim
                .char.Alas = .CharMimetizado.Alas
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0

            End If
        
            .flags.AdminInvisible = 1
            .flags.Invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .char.Body
            .flags.OldHead = .char.Head
            
            .char.Body = 0
            .char.Head = 0
            .char.ShieldAnim = NingunEscudo
            .char.WeaponAnim = NingunArma
            .char.CascoAnim = NingunCasco
            .char.Alas = NingunAlas
            
        Else
   
            .flags.AdminInvisible = 0
            .flags.Invisible = 0
            .flags.Oculto = 0
            .Counters.Ocultando = 0
            .char.Body = .flags.OldBody
            .char.Head = .flags.OldHead
            
            If .Invent.EscudoEqpObjIndex > 0 Then .char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

            If .Invent.WeaponEqpObjIndex > 0 Then .char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim

            If .Invent.CascoEqpObjIndex > 0 Then .char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
            If .Invent.AlaEqpObjIndex > 0 Then .char.Alas = ObjData(.Invent.AlaEqpObjIndex).Ropaje
        
        End If
    
        'vuelve a ser visible por la fuerza
        .showName = Not .showName
        'Call ChangeUserChar(SendTarget.ToPCArea, UserIndex, .pos.Map, UserIndex, .char.Body, .char.Head, _
         .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
                
        'Sucio, pero funciona, y siendo un comando administrativo de uso poco frecuente no molesta demasiado...
        Call EraseUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex)
        Call MakeUserChar(SendTarget.ToMap, 0, .pos.Map, UserIndex, .pos.Map, .pos.X, .pos.Y)
                    
        Call SendData(SendTarget.ToMap, 0, .pos.Map, "NOVER" & .char.CharIndex & ",0," & .PartyIndex)

    End With

End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    Dim Suerte    As Byte
    Dim exito     As Byte
    Dim raise     As Byte
    Dim Obj       As Obj
    Dim posMadera As WorldPos

    If Not LegalPos(Map, X, Y) Then Exit Sub

    With posMadera
        .Map = Map
        .X = X
        .Y = Y

    End With

    If Distancia(posMadera, UserList(UserIndex).pos) > 2 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Estás demasiado lejos para prender la fogata." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||No puedes hacer fogatas estando muerto." & FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Necesitas por lo menos tres troncos para hacer una fogata." & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
        Suerte = 3
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
        Suerte = 2
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
        Suerte = 1

    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        Obj.ObjIndex = FOGATA_APAG
        Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount \ 3
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    
        Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
    
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(UserIndex).flags.TargetObj = FOGATA_APAG
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||No has podido hacer la fogata." & FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 10

        End If

        '[/CDT]
    End If

    Call SubirSkill(UserIndex, Supervivencia)

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    Dim Suerte As Integer
    Dim res    As Integer

    If UserList(UserIndex).pos.Map = 1 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No puedes pescar en ciudades!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).pos.Map = 36 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No puedes pescar en ciudades!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).pos.Map = 34 Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Aqui no puedes pescar!!!" & FONTTYPE_INFO)
        Exit Sub

    End If

    If UCase$(UserList(UserIndex).Clase) = "TRABAJADOR" Then
        Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    Else
        Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= -1 Then
        Suerte = 35
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 11 Then
        Suerte = 30
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 21 Then
        Suerte = 28
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 31 Then
        Suerte = 24
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 41 Then
        Suerte = 22
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 51 Then
        Suerte = 20
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 61 Then
        Suerte = 18
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 71 Then
        Suerte = 15
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 81 Then
        Suerte = 13
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 91 Then
        Suerte = 7

    End If

    res = RandomNumber(1, Suerte)

    If res < 6 Then
        Dim nPos  As WorldPos
        Dim MiObj As Obj
    
        MiObj.Amount = 3
        MiObj.ObjIndex = Pescado
    
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

        End If
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Has pescado un lindo pez!" & FONTTYPE_INFO)
    
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 6

        End If

        '[/CDT]
    End If

    Call SubirSkill(UserIndex, Pesca)

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    Exit Sub

errhandler:
    Call LogError("Error en DoPescar")

End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    Dim iSkill     As Integer
    Dim Suerte     As Integer
    Dim res        As Integer
    Dim EsPescador As Boolean
                  
    If UCase(UserList(UserIndex).Clase) = "TRABAJADOR" Then
        Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
        EsPescador = True
    Else
        Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
        EsPescador = False

    End If

    iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)

    ' m = (60-11)/(1-10)
    ' y = mx - m*10 + 11

    Select Case iSkill

        Case 0
            Suerte = 0

        Case 1 To 10
            Suerte = 60

        Case 11 To 20
            Suerte = 54

        Case 21 To 30
            Suerte = 49

        Case 31 To 40
            Suerte = 43

        Case 41 To 50
            Suerte = 38

        Case 51 To 60
            Suerte = 32

        Case 61 To 70
            Suerte = 27

        Case 71 To 80
            Suerte = 21

        Case 81 To 90
            Suerte = 16

        Case 91 To 100
            Suerte = 11

        Case Else
            Suerte = 0

    End Select

    If Suerte > 0 Then
        res = RandomNumber(1, Suerte)
    
        If res < 6 Then
            Dim nPos                  As WorldPos
            Dim MiObj                 As Obj
            Dim PecesPosibles(1 To 4) As Integer
        
            PecesPosibles(1) = PESCADO1
            PecesPosibles(2) = PESCADO2
            PecesPosibles(3) = PESCADO3
            PecesPosibles(4) = PESCADO4
        
            If EsPescador = True Then
                MiObj.Amount = RandomNumber(1, 5)
            Else
                MiObj.Amount = 5

            End If

            MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

            End If
        
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Has pescado algunos peces!" & FONTTYPE_INFO)
        
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)

        End If
    
        Call SubirSkill(UserIndex, Pesca)

    End If

    Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")

End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

    If Not MapInfo(UserList(VictimaIndex).pos.Map).Pk Then Exit Sub

    If UserList(LadrOnIndex).flags.Seguro Then
        Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||Debes quitar el seguro para robar" & FONTTYPE_FIGHT)
        Exit Sub

    End If

    If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

    If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
        Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||No puedes robar a otros miembros de las fuerzas del caos" & FONTTYPE_FIGHT)
        Exit Sub

    End If

    If UserList(VictimaIndex).flags.Privilegios = PlayerType.User Then
        Dim Suerte As Integer
        Dim res    As Integer
    
        If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
            Suerte = 35
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
            Suerte = 30
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
            Suerte = 28
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
            Suerte = 24
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
            Suerte = 22
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
            Suerte = 20
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
            Suerte = 18
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
            Suerte = 15
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
            Suerte = 10
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 100 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
            Suerte = 5

        End If

        res = RandomNumber(1, Suerte)
    
        If res < 3 Then 'Exito robo
       
            If (RandomNumber(1, 50) < 25) And (UCase$(UserList(LadrOnIndex).Clase) = "LADRON") Then
                If TieneObjetosRobables(VictimaIndex) Then
                    Call RobarObjeto(LadrOnIndex, VictimaIndex)
                Else
                    Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene objetos." & FONTTYPE_INFO)

                End If

            Else 'Roba oro

                If UserList(VictimaIndex).Stats.GLD > 0 Then
                    Dim n As Integer
                
                    If UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                        n = RandomNumber(100, 1000)
                    Else
                        n = RandomNumber(1, 100)

                    End If

                    If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                    UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                
                    UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n

                    If UserList(LadrOnIndex).Stats.GLD > MaxOro Then UserList(LadrOnIndex).Stats.GLD = MaxOro
                
                    Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name & _
                        FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)

                End If

            End If

        Else
            Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||¡No has logrado robar nada!" & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!" & FONTTYPE_INFO)
            Call SendData(SendTarget.toIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " es un criminal!" & FONTTYPE_INFO)

        End If

        If Not Criminal(LadrOnIndex) Then
            Call VolverCriminal(LadrOnIndex)
        End If
    
        'If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then
        '    Call ExpulsarFaccionReal(LadrOnIndex)
        'End If

        'If UserList(LadrOnIndex).Faccion.Templario = 1 Then
        '    Call ExpulsarFaccionTemplario(LadrOnIndex)
        'End If

        UserList(LadrOnIndex).Reputacion.LadronesRep = UserList(LadrOnIndex).Reputacion.LadronesRep + vlLadron

        If UserList(LadrOnIndex).Reputacion.LadronesRep > MAXREP Then UserList(LadrOnIndex).Reputacion.LadronesRep = MAXREP
        Call SubirSkill(LadrOnIndex, Robar)

    End If

End Sub

Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean

    ' Agregué los barcos
    ' Esta funcion determina qué objetos son robables.

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And ObjData(OI).Real = 0 _
        And ObjData(OI).Caos = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

    'Call LogTarea("Sub RobarObjeto")
    Dim flag As Boolean
    Dim i    As Integer
    flag = False

    If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
        i = 1

        Do While Not flag And i <= MAX_INVENTORY_SLOTS

            'Hay objeto en este slot?
            If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
                If ObjEsRobable(VictimaIndex, i) Then
                    If RandomNumber(1, 10) < 4 Then flag = True

                End If

            End If

            If Not flag Then i = i + 1
        Loop
    Else
        i = 20

        Do While Not flag And i > 0

            'Hay objeto en este slot?
            If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
                If ObjEsRobable(VictimaIndex, i) Then
                    If RandomNumber(1, 10) < 4 Then flag = True

                End If

            End If

            If Not flag Then i = i - 1
        Loop

    End If

    If flag Then
        Dim MiObj As Obj
        Dim num   As Byte
        'Cantidad al azar
        num = RandomNumber(1, 5)
                
        If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
            num = UserList(VictimaIndex).Invent.Object(i).Amount

        End If
                
        MiObj.Amount = num
        MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
        UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
        If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
            Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

        End If
            
        Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
        If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(LadrOnIndex).pos, MiObj)

        End If
    
        Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toIndex, LadrOnIndex, 0, "||No has logrado robar un objetos." & FONTTYPE_INFO)

    End If

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)

    Dim Suerte As Integer
    Dim res    As Integer
    Dim Skill  As Byte
    
    With UserList(UserIndex)
    
        Skill = .Stats.UserSkills(eSkill.Apuñalar)
 
        If Skill <= 10 And Skill >= -1 Then
            Suerte = 200
        ElseIf Skill <= 20 And Skill >= 11 Then
            Suerte = 190
        ElseIf Skill <= 30 And Skill >= 21 Then
            Suerte = 180
        ElseIf Skill <= 40 And Skill >= 31 Then
            Suerte = 170
        ElseIf Skill <= 50 And Skill >= 41 Then
            Suerte = 160
        ElseIf Skill <= 60 And Skill >= 51 Then
            Suerte = 150
        ElseIf Skill <= 70 And Skill >= 61 Then
            Suerte = 140
        ElseIf Skill <= 80 And Skill >= 71 Then
            Suerte = 130
        ElseIf Skill <= 90 And Skill >= 81 Then
            Suerte = 120
        ElseIf Skill < 100 And Skill >= 91 Then
            Suerte = 110
        ElseIf Skill = 100 Then
            Suerte = 100

        End If

        If UCase$(.Clase) = "ASESINO" Then
            res = RandomNumber(1, Suerte)

            If res < 25 Then res = 0
        Else
            res = RandomNumber(1, Suerte * 1.2)

        End If

        If res < 15 Then
            Dim DañoApuñalar As Integer, DañoTotal As Integer
            Dim Heading As eHeading, tHeading As eHeading
            Heading = .char.Heading
            
            DañoApuñalar = ((.Stats.ELV * 2.4) + (Daño - 30))
            'FORMULA APU : ((.NivelUser * 2.4) + (Daño - 30)) x %PosicionGolpeo
        
            If VictimUserIndex <> 0 Then
                tHeading = UserList(VictimUserIndex).char.Heading
                
                DañoApuñalar = CInt(DañoApuñalar * BonoApuñalar(Heading, tHeading))
                DañoTotal = DañoApuñalar + Daño
                
                UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - DañoTotal
                
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & DañoApuñalar & _
                    FONTTYPE_APU)
                
                Call SendData(SendTarget.toIndex, VictimUserIndex, 0, "||Te ha apuñalado " & .Name & " por " & DañoApuñalar & FONTTYPE_APU)
                        
                Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "TW15")
                        
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, UserList(VictimUserIndex).pos.Map, "CFX" & UserList( _
                    VictimUserIndex).char.CharIndex & "," & 17 & "," & 1)
                        
                Call SendData(ToPCArea, UserIndex, .pos.Map, "||" & vbCyan & "°Apu! + " & DañoApuñalar & "!°" & CStr(.char.CharIndex))
            Else
                tHeading = Npclist(VictimNpcIndex).char.Heading
                
                DañoApuñalar = CInt(DañoApuñalar * BonoApuñalar(Heading, tHeading))
                DañoTotal = DañoApuñalar + Daño
            
                Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - DañoTotal
                Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has apuñalado la criatura por " & DañoApuñalar & FONTTYPE_APU)
                        
                Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "TW13")
                
                ' Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, Npclist(VictimNpcIndex).pos.Map, "CFX" & Npclist(VictimNpcIndex).char.CharIndex _
                  & "," & 17 & "," & 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, .pos.Map, "||" & vbCyan & "°Apu! + " & DañoApuñalar & "!°" & CStr(.char.CharIndex))
                        
                '[Alejo]
                Call CalcularDarExp(UserIndex, VictimNpcIndex, DañoTotal)

            End If

            Call SubirSkill(UserIndex, Apuñalar)
        Else
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No has logrado apuñalar a tu enemigo!" & FONTTYPE_FIGHT)

        End If

    End With

End Sub

Private Function BonoApuñalar(ByVal Heading As eHeading, ByVal tHeading As eHeading) As Single
    'GOLPE POR ESPALDA: %PosicionGolpe = 100%
    'GOLPE POR LATERAL: %PosicionGolpe = 50%
    'GOLPE POR DELANTE: %PosicionGolpe = 0%

    '  Mirando para el mismo lado, siempre va a ser apuñalada por la espalda.
    If Heading = tHeading Then
        BonoApuñalar = 2 ' Si no entendi mal al poner 1000%, quisieron poner el doble no ?
        Exit Function

    End If
    
    ' Yo mirando al Sur o al norte.
    ' El otro mirando al Este o Oeste, golpe de lateral.
    If Heading = eHeading.SOUTH Or Heading = eHeading.NORTH Then
        If tHeading = eHeading.EAST Or tHeading = eHeading.WEST Then
            BonoApuñalar = 0.5
            Exit Function

        End If

    End If
    
    ' y bueno por descarte, las siguiente accion significa que estan de frente.
    ' Face To Face
    BonoApuñalar = 1

End Function

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0

End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    Dim Suerte As Integer
    Dim res    As Integer

    If UCase$(UserList(UserIndex).Clase) = "TRABAJADOR" Then
        Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
    Else
        Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= -1 Then
        Suerte = 35
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 11 Then
        Suerte = 30
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 21 Then
        Suerte = 28
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 31 Then
        Suerte = 24
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 41 Then
        Suerte = 22
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 51 Then
        Suerte = 20
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 61 Then
        Suerte = 18
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 71 Then
        Suerte = 15
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 81 Then
        Suerte = 13
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 91 Then
        Suerte = 7

    End If

    res = RandomNumber(1, Suerte)

    If res < 6 Then
        Dim nPos  As WorldPos
        Dim MiObj As Obj
    
        If UCase$(UserList(UserIndex).Clase) = "TRABAJADOR" Then
            MiObj.Amount = RandomNumber(1, 5)
        Else
            MiObj.Amount = 1

        End If
    
        MiObj.ObjIndex = Leña
    
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
        
        End If
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Has conseguido algo de leña!" & FONTTYPE_INFO)
    
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No has obtenido leña!" & FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 8

        End If

        '[/CDT]
    End If

    Call SubirSkill(UserIndex, Talar)

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).trigger = 6 Then Exit Sub

    If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then
        UserList(UserIndex).Reputacion.BurguesRep = 0
        UserList(UserIndex).Reputacion.NobleRep = 0
        UserList(UserIndex).Reputacion.PlebeRep = 0
        UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + vlASALTO

        If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then UserList(UserIndex).Reputacion.BandidoRep = MAXREP

        'If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
        '    Call ExpulsarFaccionReal(UserIndex)
        'End If
    
        'If UserList(UserIndex).Faccion.Templario = 1 Then
        '    Call ExpulsarFaccionTemplario(UserIndex)
        'End If
    
    End If
    
    OnlineCriminal = OnlineCriminal + 1
    OnlineCiudadano = OnlineCiudadano - 1
    
#If MYSQL = 1 Then
    DoEvents
    Call Add_DataBase(UserIndex, "Ranking")
#End If

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)

    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).trigger = 6 Then Exit Sub

    UserList(UserIndex).Reputacion.LadronesRep = 0
    UserList(UserIndex).Reputacion.BandidoRep = 0
    UserList(UserIndex).Reputacion.AsesinoRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlASALTO

    If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then UserList(UserIndex).Reputacion.PlebeRep = MAXREP

    'If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    '    Call ExpulsarFaccionCaos(UserIndex)
    'End If
    
    'If UserList(UserIndex).Faccion.Nemesis = 1 Then
    '    Call ExpulsarFaccionNemesis(UserIndex)
    'End If
     
    OnlineCiudadano = OnlineCiudadano + 1
    OnlineCriminal = OnlineCriminal - 1
    
#If MYSQL = 1 Then
    DoEvents
    Call Add_DataBase(UserIndex, "Ranking")
#End If

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    Dim Suerte As Integer
    Dim res    As Integer
    Dim metal  As Integer

    If UCase$(UserList(UserIndex).Clase) = "TRABAJADOR" Then
        Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
    Else
        Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= -1 Then
        Suerte = 35
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 11 Then
        Suerte = 30
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 21 Then
        Suerte = 28
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 31 Then
        Suerte = 24
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 41 Then
        Suerte = 22
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 51 Then
        Suerte = 20
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 61 Then
        Suerte = 18
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 71 Then
        Suerte = 15
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 81 Then
        Suerte = 10
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 91 Then
        Suerte = 7

    End If

    res = RandomNumber(1, Suerte)

    If res <= 5 Then
        Dim MiObj As Obj
        Dim nPos  As WorldPos
    
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
        MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    
        If UCase$(UserList(UserIndex).Clase) = "TRABAJADOR" Then
            MiObj.Amount = RandomNumber(1, 6)
        Else
            MiObj.Amount = 1

        End If
    
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Has extraido algunos minerales!" & FONTTYPE_INFO)
    
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡No has conseguido nada!" & FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 9

        End If

        '[/CDT]
    End If

    Call SubirSkill(UserIndex, Mineria)

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

    UserList(UserIndex).Counters.IdleCount = 0

    Dim Suerte  As Integer
    Dim res     As Integer
    Dim Cant    As Integer

    'Barrin 3/10/03
    'Esperamos a que se termine de concentrar
    Dim TActual As Long
    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
        Exit Sub

    End If

    If UserList(UserIndex).Counters.bPuedeMeditar = False Then
        UserList(UserIndex).Counters.bPuedeMeditar = True

    End If

    If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
        Call SendData(SendTarget.toIndex, UserIndex, 0, "Z16")
        Call SendData(SendTarget.toIndex, UserIndex, 0, "MEDOK")
        UserList(UserIndex).flags.Meditando = False
        UserList(UserIndex).char.FX = 0
        UserList(UserIndex).char.loops = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 0 & "," & 0)
        Exit Sub

    End If
    
    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
        If UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
            Suerte = 35
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
            Suerte = 30
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
            Suerte = 28
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
            Suerte = 24
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
            Suerte = 22
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
            Suerte = 20
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
            Suerte = 18
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
            Suerte = 15
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
            Suerte = 10
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
            Suerte = 5

        End If
   
    Else
   
        If UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 10 Then
            Suerte = 35
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 110 Then
            Suerte = 30
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 210 Then
            Suerte = 28
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 310 Then
            Suerte = 24
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 410 Then
            Suerte = 22
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 510 Then
            Suerte = 20
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 610 Then
            Suerte = 18
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 710 Then
            Suerte = 15
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 810 Then
            Suerte = 10
        ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 910 Then
            Suerte = 5

        End If
    
    End If
   
    res = RandomNumber(1, Suerte)

    If res = 1 Then
        Cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Cant

        If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList( _
            UserIndex).Stats.MaxMAN
    
        If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
            Call SendData(SendTarget.toIndex, UserIndex, 0, "||¡Has recuperado " & Cant & " puntos de mana!" & FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 22

        End If
    
        Call SendData(SendTarget.toIndex, UserIndex, 0, "ASM" & UserList(UserIndex).Stats.MinMAN)
        Call SubirSkill(UserIndex, Meditar)

    End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

    Dim Suerte As Integer
    Dim res    As Integer

    If UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= -1 Then
        Suerte = 35
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 11 Then
        Suerte = 30
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 21 Then
        Suerte = 28
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 31 Then
        Suerte = 24
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 41 Then
        Suerte = 22
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 51 Then
        Suerte = 20
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 61 Then
        Suerte = 18
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 71 Then
        Suerte = 15
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 81 Then
        Suerte = 10
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 91 Then
        Suerte = 5

    End If

    res = RandomNumber(1, Suerte)

    If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call SendData(SendTarget.toIndex, UserIndex, 0, "||Has logrado desarmar a tu oponente!" & FONTTYPE_FIGHT)

        If UserList(VictimIndex).Stats.ELV < 20 Then Call SendData(SendTarget.toIndex, VictimIndex, 0, "||Tu oponente te ha desarmado!" & _
            FONTTYPE_FIGHT)

    End If

End Sub


