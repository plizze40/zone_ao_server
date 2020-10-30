Attribute VB_Name = "Trabajo"
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Option Explicit

Public Sub DoOcultarse(UserIndex As Integer)
    On Error GoTo errhandler
    Dim Suerte As Integer

    Suerte = 50 - 0.35 * UserList(UserIndex).Stats.UserSkills(Ocultarse)

    If TiempoTranscurrido(UserList(UserIndex).Counters.LastOculto) < 0.5 Then Exit Sub
    UserList(UserIndex).Counters.LastOculto = Timer

    If UserList(UserIndex).Clase = CAZADOR Or UserList(UserIndex).Clase = ASESINO Or UserList(UserIndex).Clase = LADRON Then Suerte = Suerte - 5

    If CInt(RandomNumber(1, Suerte)) <= 5 Then
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.Invisible = 1
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, ("V3" & UserList(UserIndex).Char.CharIndex & ",1"))
        Call SendData(ToIndex, UserIndex, 0, "V7")
        Call SubirSkill(UserIndex, Ocultarse, 15)
    Else: Call SendData(ToIndex, UserIndex, 0, "EN")
    End If

    Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub
Public Sub DoNavega(UserIndex As Integer, Slot As Integer)
    Dim Barco As ObjData, Skill As Byte

    Barco = ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex)

    If UserList(UserIndex).Clase <> PIRATA And UserList(UserIndex).Clase <> PESCADOR Then
        Skill = Barco.MinSkill * 2
    ElseIf UserList(UserIndex).Invent.Object(Slot).OBJIndex = 474 Then
        Skill = 40
    Else: Skill = Barco.MinSkill
    End If

    If UserList(UserIndex).Stats.UserSkills(Navegacion) < Skill Then
        If Skill <= 100 Then
            Call SendData(ToIndex, UserIndex, 0, "!7" & Skill)
        Else: Call SendData(ToIndex, UserIndex, 0, "||Esta embarcación solo puede ser usada por piratas." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If

    If UserList(UserIndex).Stats.ELV < 18 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes ser nivel 18 o superior para poder navegar." & FONTTYPE_INFO)
        Exit Sub
    End If

    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
    UserList(UserIndex).Invent.BarcoSlot = Slot

    If UserList(UserIndex).flags.Navegando = 0 Then
        UserList(UserIndex).Char.Head = 0

        If UserList(UserIndex).flags.Muerto = 0 Then
            UserList(UserIndex).Char.Body = Barco.Ropaje
        Else
            UserList(UserIndex).Char.Body = iFragataFantasmal
        End If

        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
        UserList(UserIndex).flags.Navegando = 1
    Else
        UserList(UserIndex).flags.Navegando = 0

        If UserList(UserIndex).flags.Muerto = 0 Then
            UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head

            If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            Else: Call DarCuerpoDesnudo(UserIndex)
            End If

            If UserList(UserIndex).Invent.EscudoEqpObjIndex Then _
               UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
            If UserList(UserIndex).Invent.WeaponEqpObjIndex Then _
               UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
            If UserList(UserIndex).Invent.CascoEqpObjIndex Then _
               UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
        Else
            UserList(UserIndex).Char.Body = iCuerpoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        End If
    End If

    Call ChangeUserCharB(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(ToIndex, UserIndex, 0, "NAVEG")

End Sub
Public Sub FundirMineral(UserIndex As Integer)

    If UserList(UserIndex).flags.TargetObjInvIndex Then
        If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(Mineria) / ModFundicion(UserList(UserIndex).Clase) Then
            Call DoLingotes(UserIndex)
        Else: Call SendData(ToIndex, UserIndex, 0, "!8")
        End If
    End If

End Sub
Function TieneObjetos(ItemIndex As Integer, Cant As Integer, UserIndex As Integer) As Boolean
    Dim i As Byte
    Dim Total As Long

    For i = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(i).OBJIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount
        End If
    Next

    If Cant <= Total Then
        TieneObjetos = True
        Exit Function
    End If

End Function
Function QuitarObjetos(ItemIndex As Integer, Cant As Integer, UserIndex As Integer) As Boolean
    Dim i As Byte

    For i = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(i).OBJIndex = ItemIndex Then

            Call Desequipar(UserIndex, i)

            UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - Cant
            If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
                Cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
                UserList(UserIndex).Invent.Object(i).Amount = 0
                UserList(UserIndex).Invent.Object(i).OBJIndex = 0
            Else
                Cant = 0
            End If

            Call UpdateUserInv(False, UserIndex, i)

            If (Cant = 0) Then
                QuitarObjetos = True
                Exit Function
            End If
        End If
    Next

End Function
Sub HerreroQuitarMateriales(UserIndex As Integer, ItemIndex As Integer, cantT As Integer)
    Dim Descuento As Single

    Descuento = 1

    If UserList(UserIndex).Clase = HERRERO Then
        If UserList(UserIndex).Recompensas(1) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO Then
            If CInt(RandomNumber(1, 4)) <= 1 Then Descuento = 0.5
        ElseIf UserList(UserIndex).Recompensas(1) = 2 And (ObjData(ItemIndex).SubTipo = OBJTYPE_CASCO Or ObjData(ItemIndex).SubTipo = OBJTYPE_ESCUDO) Then
            Descuento = 0.5
        End If
        Descuento = Descuento * (1 - 0.25 * Buleano(UserList(UserIndex).Recompensas(3) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO))
    End If

    If ObjData(ItemIndex).LingH Then Call QuitarObjetos(LingoteHierro, Descuento * Int(ObjData(ItemIndex).LingH * ModMateriales(UserList(UserIndex).Clase) * cantT), UserIndex)
    If ObjData(ItemIndex).LingP Then Call QuitarObjetos(LingotePlata, Descuento * Int(ObjData(ItemIndex).LingP * ModMateriales(UserList(UserIndex).Clase) * cantT), UserIndex)
    If ObjData(ItemIndex).LingO Then Call QuitarObjetos(LingoteOro, Descuento * Int(ObjData(ItemIndex).LingO * ModMateriales(UserList(UserIndex).Clase) * cantT), UserIndex)

End Sub
Sub CarpinteroQuitarMateriales(UserIndex As Integer, ItemIndex As Integer, cantT As Integer)
    Dim Descuento As Single

    cantT = Maximo(1, cantT)

    If UserList(UserIndex).Clase = CARPINTERO And UserList(UserIndex).Recompensas(2) = 2 And ObjData(ItemIndex).ObjType = OBJTYPE_BARCOS Then
        Descuento = 0.8
    Else
        Descuento = 1
    End If

    If ObjData(ItemIndex).Madera Then
        Call QuitarObjetos(Leña, CInt(Descuento * ObjData(ItemIndex).Madera * ModMadera(UserList(UserIndex).Clase) * cantT), UserIndex)
    End If

    If ObjData(ItemIndex).MaderaElfica Then
        Call QuitarObjetos(LeñaElfica, CInt(Descuento * ObjData(ItemIndex).MaderaElfica * ModMadera(UserList(UserIndex).Clase) * cantT), UserIndex)
    End If

End Sub
Function CarpinteroTieneMateriales(UserIndex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
    Dim Descuento As Single

    cantT = Maximo(1, cantT)

    If UserList(UserIndex).Clase = CARPINTERO And UserList(UserIndex).Recompensas(2) = 2 And ObjData(ItemIndex).ObjType = OBJTYPE_BARCOS Then
        Descuento = 0.8
    Else
        Descuento = 1
    End If

    If ObjData(ItemIndex).Madera Then
        If Not TieneObjetos(Leña, CInt(Descuento * ObjData(ItemIndex).Madera * ModMadera(UserList(UserIndex).Clase) * cantT), UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "!9")
            CarpinteroTieneMateriales = False
            Exit Function
        End If
    End If

    If ObjData(ItemIndex).MaderaElfica Then
        If Not TieneObjetos(LeñaElfica, CInt(Descuento * ObjData(ItemIndex).MaderaElfica * ModMadera(UserList(UserIndex).Clase) * cantT), UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "!9")
            CarpinteroTieneMateriales = False
            Exit Function
        End If
    End If

    CarpinteroTieneMateriales = True

End Function
Function Piel(UserIndex As Integer, Tipo As Byte, Obj As Integer) As Integer

    Select Case Tipo
    Case 1
        Piel = ObjData(Obj).PielLobo
        If UserList(UserIndex).Clase = SASTRE And UserList(UserIndex).Stats.ELV >= 18 Then Piel = Piel * 0.8
    Case 2
        Piel = ObjData(Obj).PielOsoPardo
        If UserList(UserIndex).Clase = SASTRE And UserList(UserIndex).Stats.ELV >= 18 Then Piel = Piel * 0.8
    Case 3
        Piel = ObjData(Obj).PielOsoPolar
        If UserList(UserIndex).Clase = SASTRE And UserList(UserIndex).Stats.ELV >= 18 Then Piel = Piel * 0.8
    End Select

End Function
Function SastreTieneMateriales(UserIndex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
    Dim PielL As Integer, PielO As Integer, PielP As Integer
    cantT = Maximo(1, cantT)

    PielL = ObjData(ItemIndex).PielLobo
    PielO = ObjData(ItemIndex).PielOsoPardo
    PielP = ObjData(ItemIndex).PielOsoPolar

    If UserList(UserIndex).Clase = SASTRE And UserList(UserIndex).Stats.ELV >= 18 Then
        PielL = 0.8 * PielL
        PielO = 0.8 * PielO
        PielP = 0.8 * PielP
    End If

    If PielL Then
        If Not TieneObjetos(PLobo, CInt(PielL * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "0A")
            SastreTieneMateriales = False
            Exit Function
        End If
    End If

    If PielO Then
        If Not TieneObjetos(POsoPardo, CInt(PielO * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "0A")
            SastreTieneMateriales = False
            Exit Function
        End If
    End If

    If PielP Then
        If Not TieneObjetos(POsoPolar, CInt(PielP * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "0A")
            SastreTieneMateriales = False
            Exit Function
        End If
    End If

    SastreTieneMateriales = True

End Function
Sub SastreQuitarMateriales(UserIndex As Integer, ItemIndex As Integer, cantT As Integer)
    Dim PielL As Integer, PielO As Integer, PielP As Integer

    PielL = ObjData(ItemIndex).PielLobo
    PielO = ObjData(ItemIndex).PielOsoPardo
    PielP = ObjData(ItemIndex).PielOsoPolar

    If UserList(UserIndex).Clase = SASTRE And UserList(UserIndex).Stats.ELV >= 18 Then
        PielL = 0.8 * PielL
        PielO = 0.8 * PielO
        PielP = 0.8 * PielP
    End If

    If PielL Then Call QuitarObjetos(PLobo, CInt(PielL * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex)
    If PielO Then Call QuitarObjetos(POsoPardo, CInt(PielO * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex)
    If PielP Then Call QuitarObjetos(POsoPolar, CInt(PielP * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex)

End Sub
Public Sub SastreConstruirItem(UserIndex As Integer, ItemIndex As Integer, cantT As Integer)

    If SastreTieneMateriales(UserIndex, ItemIndex, cantT) And _
       UserList(UserIndex).Stats.UserSkills(Sastreria) / ModRopas(UserList(UserIndex).Clase) >= _
       ObjData(ItemIndex).SkSastreria And _
       PuedeConstruirSastre(ItemIndex, UserIndex) And _
       UserList(UserIndex).Invent.HerramientaEqpObjIndex = HILAR_SASTRE Then

        Call SastreQuitarMateriales(UserIndex, ItemIndex, cantT)
        Call SendData(ToIndex, UserIndex, 0, "0C")

        Dim MiObj As Obj
        MiObj.Amount = Maximo(1, cantT)
        MiObj.OBJIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)

        Call CheckUserLevel(UserIndex)

        Call SubirSkill(UserIndex, Sastreria, 5)

    Else
        Call SendData(ToIndex, UserIndex, 0, "0D")

    End If

End Sub

Public Function PuedeConstruirSastre(ItemIndex As Integer, UserIndex As Integer) As Boolean
    Dim i As Long
    Dim N As Integer

    N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

    For i = 1 To UBound(ObjSastre)
        If ObjSastre(i) = ItemIndex Then
            PuedeConstruirSastre = True
            Exit Function
        End If
    Next

    PuedeConstruirSastre = False

End Function

Function HerreroTieneMateriales(UserIndex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
    Dim Descuento As Single

    Descuento = 1

    If UserList(UserIndex).Clase = HERRERO Then
        If UserList(UserIndex).Recompensas(1) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO Then
            If CInt(RandomNumber(1, 4)) <= 1 Then Descuento = 0.5
        ElseIf UserList(UserIndex).Recompensas(1) = 2 And (ObjData(ItemIndex).SubTipo = OBJTYPE_CASCO Or ObjData(ItemIndex).SubTipo = OBJTYPE_ESCUDO) Then
            Descuento = 0.5
        End If
        Descuento = Descuento * (1 - 0.25 * Buleano(UserList(UserIndex).Recompensas(3) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO))
    End If

    If ObjData(ItemIndex).LingH Then
        If Not TieneObjetos(LingoteHierro, Descuento * Int(ObjData(ItemIndex).LingH * ModMateriales(UserList(UserIndex).Clase) * cantT), UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "0E")
            HerreroTieneMateriales = False
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).LingP Then
        If Not TieneObjetos(LingotePlata, Descuento * Int(ObjData(ItemIndex).LingP * ModMateriales(UserList(UserIndex).Clase) * cantT), UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "0F")
            HerreroTieneMateriales = False
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).LingO Then
        If Not TieneObjetos(LingoteOro, Descuento * Int(ObjData(ItemIndex).LingO * ModMateriales(UserList(UserIndex).Clase) * cantT), UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "0G")
            HerreroTieneMateriales = False
            Exit Function
        End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(UserIndex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex, cantT) And UserList(UserIndex).Stats.UserSkills(Herreria) >= ObjData(ItemIndex).SkHerreria * ModHerreriA(UserList(UserIndex).Clase)
End Function
Public Function PuedeConstruirHerreria(ItemIndex As Integer) As Boolean
    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)
        If ArmasHerrero(i).Index = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function
        End If
    Next

    For i = 1 To UBound(ArmadurasHerrero)
        If ArmadurasHerrero(i).Index = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function
        End If
    Next

    For i = 1 To UBound(CascosHerrero)
        If CascosHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function
        End If
    Next

    For i = 1 To UBound(EscudosHerrero)
        If EscudosHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function
        End If
    Next

    PuedeConstruirHerreria = False

End Function
Public Sub HerreroConstruirItem(UserIndex As Integer, ItemIndex As Integer, cantT As Integer)

    If cantT > 10 Then
        Call SendData(ToIndex, UserIndex, 0, "0H")
        Exit Sub
    End If

    If PuedeConstruir(UserIndex, ItemIndex, cantT) And PuedeConstruirHerreria(ItemIndex) Then
        Call HerreroQuitarMateriales(UserIndex, ItemIndex, cantT)

        Select Case ObjData(ItemIndex).ObjType
        Case OBJTYPE_WEAPON
            Call SendData(ToIndex, UserIndex, 0, "0I")
        Case OBJTYPE_ESCUDO
            Call SendData(ToIndex, UserIndex, 0, "0L")
        Case OBJTYPE_CASCO
            Call SendData(ToIndex, UserIndex, 0, "0K")
        Case OBJTYPE_ARMOUR
            Call SendData(ToIndex, UserIndex, 0, "0J")
        End Select
        cantT = cantT * (1 + Buleano(CInt(RandomNumber(1, 10)) <= 1 And UserList(UserIndex).Clase = HERRERO And UserList(UserIndex).Recompensas(3) = 2))
        Dim MiObj As Obj
        MiObj.Amount = cantT
        MiObj.OBJIndex = ItemIndex
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)

        Call CheckUserLevel(UserIndex)
        Call SubirSkill(UserIndex, Herreria, 5)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & MARTILLOHERRERO)
    Else

    End If

End Sub
Public Function PuedeConstruirCarpintero(ItemIndex As Integer) As Boolean
    Dim i As Long

    For i = 1 To UBound(ObjCarpintero)
        If ObjCarpintero(i).Index = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function
        End If
    Next
    PuedeConstruirCarpintero = False

End Function
Public Sub CarpinteroConstruirItem(UserIndex As Integer, ItemIndex As Integer, cantT As Integer)

    If CarpinteroTieneMateriales(UserIndex, ItemIndex, cantT) And _
       UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
       ObjData(ItemIndex).SkCarpinteria And _
       PuedeConstruirCarpintero(ItemIndex) And _
       UserList(UserIndex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then

        Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, cantT)
        Call SendData(ToIndex, UserIndex, 0, "0M")

        Dim MiObj As Obj
        MiObj.Amount = cantT
        MiObj.OBJIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)


        Call CheckUserLevel(UserIndex)

        Call SubirSkill(UserIndex, Carpinteria, 5)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & LABUROCARPINTERO)
    End If

End Sub
Public Sub DoLingotes(UserIndex As Integer)
    Dim Minimo As Integer

    Select Case ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    Case LingoteHierro
        Minimo = 6
    Case LingotePlata
        Minimo = 18
    Case LingoteOro
        Minimo = 34
    End Select

    If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvslot).Amount < Minimo Then
        Call SendData(ToIndex, UserIndex, 0, "M3")
        Exit Sub
    End If

    Dim nPos As WorldPos
    Dim MiObj As Obj

    MiObj.Amount = UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvslot).Amount / Minimo
    MiObj.OBJIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex

    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)

    UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvslot).Amount = 0
    UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvslot).OBJIndex = 0

    Call UpdateUserInv(False, UserIndex, UserList(UserIndex).flags.TargetObjInvslot)
    Call SendData(ToIndex, UserIndex, 0, "M1")

End Sub
Function ModFundicion(Clase As Byte) As Single

    Select Case (Clase)
    Case MINERO, HERRERO
        ModFundicion = 1
    Case TRABAJADOR, EXPERTO_MINERALES
        ModFundicion = 2.5
    Case Else
        ModFundicion = 3
    End Select

End Function
Function ModHerreriA(Clase As Byte) As Single

    Select Case (Clase)
    Case HERRERO
        ModHerreriA = 1
    Case Else
        ModHerreriA = 3
    End Select

End Function
Function ModCarpinteria(Clase As Byte) As Single

    Select Case (Clase)
    Case CARPINTERO
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
    End Select

End Function
Function ModMateriales(Clase As Byte) As Single

    Select Case (Clase)
    Case HERRERO
        ModMateriales = 1
    Case Else
        ModMateriales = 3
    End Select

End Function
Function ModMadera(Clase As Byte) As Double

    Select Case (Clase)
    Case CARPINTERO
        ModMadera = 1
    Case Else
        ModMadera = 3
    End Select

End Function
Function ModSastre(Clase As Byte) As Double

    Select Case (Clase)
    Case SASTRE
        ModSastre = 1
    Case Else
        ModSastre = 3
    End Select

End Function
Function ModRopas(Clase As Byte) As Double

    Select Case (Clase)
    Case SASTRE
        ModRopas = 1
    Case Else
        ModRopas = 3
    End Select

End Function
Function FreeMascotaIndex(UserIndex As Integer) As Integer
    Dim j As Integer

    For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next

End Function
Sub DoDomar(UserIndex As Integer, NpcIndex As Integer)


    If UserList(UserIndex).NroMascotas < 3 Then

        If Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "0N")
            Exit Sub
        End If

        If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "0Ñ")
            Exit Sub
        End If

        If Npclist(NpcIndex).flags.Domable <= UserList(UserIndex).Stats.UserSkills(Domar) Then
            Dim Index As Integer
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
            Index = FreeMascotaIndex(UserIndex)
            UserList(UserIndex).MascotasIndex(Index) = NpcIndex
            UserList(UserIndex).MascotasType(Index) = Npclist(NpcIndex).Numero

            Npclist(NpcIndex).MaestroUser = UserIndex

            Call QuitarNPCDeLista(Npclist(NpcIndex).Numero, Npclist(NpcIndex).POS.Map)

            Call FollowAmo(NpcIndex)

            Call SendData(ToIndex, UserIndex, 0, "0O")
            Call SubirSkill(UserIndex, Domar)

        Else

            Call SendData(ToIndex, UserIndex, 0, "||Necesitas " & Npclist(NpcIndex).flags.Domable & " puntos para domar a esta criatura. " & FONTTYPE_INFO)

        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "0Q")
    End If

End Sub
Sub DoAdminInvisible(UserIndex As Integer)

    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, ("V3" & UserList(UserIndex).Char.CharIndex & ",1"))
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, "QDL" & UserList(UserIndex).Char.CharIndex)
    Else
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.Invisible = 0
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, ("V3" & UserList(UserIndex).Char.CharIndex & ",0"))
    End If

End Sub
Sub TratarDeHacerFogata(Map As Integer, X As Integer, Y As Integer, UserIndex As Integer)
    Dim Suerte As Byte
    Dim exito As Byte
    Dim raise As Byte
    Dim Obj As Obj, nPos As WorldPos

    If Not LegalPos(Map, X, Y) Then Exit Sub
    nPos.Map = Map
    nPos.X = X
    nPos.Y = Y

    If Not MapInfo(Map).Pk Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes hacer fogatas en zonas seguras." & FONTTYPE_WARNING)
        Exit Sub
    End If

    If Distancia(nPos, UserList(UserIndex).POS) > 4 Then
        Call SendData(ToIndex, UserIndex, 0, "DL")
        Exit Sub
    End If

    If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
        Call SendData(ToIndex, UserIndex, 0, "0R")
        Exit Sub
    End If

    If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
        Suerte = 3
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
        Suerte = 2
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 Then
        Suerte = 1
    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        Obj.OBJIndex = FOGATA_APAG
        Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount / 3

        If Obj.Amount > 1 Then
            Call SendData(ToIndex, UserIndex, 0, "0S" & Obj.Amount)
        Else
            Call SendData(ToIndex, UserIndex, 0, "0T")
        End If

        Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)

        Dim Fogatita As New cGarbage
        Fogatita.Map = Map
        Fogatita.X = X
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)

    Else
        Call SendData(ToIndex, UserIndex, 0, "0U")
    End If

    Call SubirSkill(UserIndex, Supervivencia)


End Sub
Public Sub DoTalar(UserIndex As Integer, Elfico As Boolean)
    On Error GoTo errhandler
    Dim MiObj As Obj
    Dim Factor As Integer
    Dim Esfuerzo As Integer

    If UserList(UserIndex).Clase = TALADOR Then
        Esfuerzo = EsfuerzoTalarLeñador
    Else
        Esfuerzo = EsfuerzoTalarGeneral
    End If

    If UserList(UserIndex).Stats.MinSta >= Esfuerzo Then
        Call QuitarSta(UserIndex, Esfuerzo)
        Call SendUserSTA(UserIndex)
    Else
        Call SendData(ToIndex, UserIndex, 0, "9E")
        Exit Sub
    End If

    If Elfico Then
        MiObj.OBJIndex = LeñaElfica
        Factor = 6
    Else
        MiObj.OBJIndex = Leña
        Factor = 5
    End If



    If UserList(UserIndex).Clase = TALADOR Then
        MiObj.Amount = Fix(4 + ((0.29 + 0.07 * Buleano(UserList(UserIndex).Recompensas(1) = 1)) * UserList(UserIndex).Stats.UserSkills(Talar)))
    Else: MiObj.Amount = 1
    End If

    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)

    Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).POS.Map, "TW" & SOUND_TALAR)
    Call SubirSkill(UserIndex, Talar, 5)

    Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub
Public Sub DoPescar(UserIndex As Integer)
    On Error GoTo errhandler
    Dim Esfuerzo As Integer
    Dim MiObj As Obj

    If UserList(UserIndex).Clase = PESCADOR Then
        Esfuerzo = EsfuerzoPescarPescador
    Else
        Esfuerzo = EsfuerzoPescarGeneral
    End If

    If UserList(UserIndex).Stats.MinSta >= Esfuerzo Then
        Call QuitarSta(UserIndex, Esfuerzo)
        Call SendUserSTA(UserIndex)
    Else
        Call SendData(ToIndex, UserIndex, 0, "9E")
        Exit Sub
    End If

    MiObj.OBJIndex = Pescado


    If UserList(UserIndex).Clase = PESCADOR Then
        If UserList(UserIndex).Recompensas(1) = 2 And UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.HerramientaEqpObjIndex = RED_PESCA And CInt(RandomNumber(1, 10)) <= 1 Then MiObj.OBJIndex = PescadoCaro + CInt(RandomNumber(1, 3))
        MiObj.Amount = Fix(4 + ((0.29 + 0.07 * Buleano(UserList(UserIndex).Recompensas(1) = 1)) * UserList(UserIndex).Stats.UserSkills(Pesca)))
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> RED_PESCA Then MiObj.Amount = MiObj.Amount / 2
    Else: MiObj.Amount = 1
    End If

    Call SubirSkill(UserIndex, Pesca, 5)
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SOUND_PESCAR)

    Exit Sub

errhandler:
    Call LogError("Error en DoPescar")

End Sub
Public Function Buleano(A As Boolean) As Byte

    Buleano = -A

End Function
Public Sub DoRobar(LadronIndex As Integer, VictimaIndex As Integer)
    Dim Res As Integer
    Dim N As Long

    If Not PuedeRobar(LadronIndex, VictimaIndex) Then Exit Sub

    UserList(LadronIndex).Counters.LastRobo = Timer

    Res = RandomNumber(1, 100)

    If Res > UserList(LadronIndex).Stats.UserSkills(Robar) \ 10 + 25 * Buleano(UserList(LadronIndex).Clase = LADRON) + 5 * Buleano(UserList(LadronIndex).Clase = LADRON And UserList(LadronIndex).Recompensas(1) = 2) Then
        Call SendData(ToIndex, LadronIndex, 0, "X0")
        Call SendData(ToIndex, VictimaIndex, 0, "Y0" & UserList(LadronIndex).Name)
    ElseIf UserList(LadronIndex).Clase = LADRON And TieneObjetosRobables(VictimaIndex) And Res <= 10 * Buleano(UserList(LadronIndex).Recompensas(2) = 2) + 10 * Buleano(UserList(LadronIndex).Recompensas(3) = 2) Then
        Call RobarObjeto(LadronIndex, VictimaIndex)
    ElseIf UserList(VictimaIndex).Stats.GLD = 0 Then
        Call SendData(ToIndex, LadronIndex, 0, "W0")
        Call SendData(ToIndex, VictimaIndex, 0, "Y0" & UserList(LadronIndex).Name)
    Else
        N = Minimo((1 + 0.1 * Buleano(UserList(LadronIndex).Recompensas(1) = 1 And UserList(LadronIndex).Clase = LADRON)) * (RandomNumber(1, (UserList(LadronIndex).Stats.UserSkills(Robar) * (UserList(VictimaIndex).Stats.ELV / 10) * UserList(LadronIndex).Stats.ELV)) / (10 + 10 * Buleano(Not UserList(LadronIndex).Clase = LADRON))), UserList(VictimaIndex).Stats.GLD)

        UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
        Call AddtoVar(UserList(LadronIndex).Stats.GLD, N, MAXORO)

        Call SendData(ToIndex, LadronIndex, 0, "U0" & UserList(VictimaIndex).Name & "," & N)
        Call SendData(ToIndex, VictimaIndex, 0, "V0" & UserList(LadronIndex).Name & "," & N)

        Call SendUserORO(LadronIndex)
        Call SendUserORO(VictimaIndex)
    End If

    Call SubirSkill(LadronIndex, Robar)

End Sub
Public Function ObjEsRobable(VictimaIndex As Integer, Slot As Byte) As Boolean
    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).OBJIndex
    If OI = 0 Then Exit Function

    ObjEsRobable = ObjData(OI).ObjType <> OBJTYPE_LLAVES And _
                   ObjData(OI).ObjType <> OBJTYPE_BARCOS And _
                   Not ObjData(OI).Real And _
                   Not ObjData(OI).Caos And _
                   Not ObjData(OI).NoSeCae

End Function
Public Sub RobarObjeto(LadronIndex As Integer, VictimaIndex As Integer)
    Dim IndexRobo As Byte
    Dim MiObj As Obj
    Dim Num As Byte

    Do
        IndexRobo = RandomNumber(1, MAX_INVENTORY_SLOTS)
        If ObjEsRobable(VictimaIndex, IndexRobo) Then Exit Do
    Loop

    MiObj.OBJIndex = UserList(VictimaIndex).Invent.Object(IndexRobo).OBJIndex

    Num = Minimo(RandomNumber(1, 4 + 96 * Buleano(ObjData(MiObj.OBJIndex).ObjType = OBJTYPE_POCIONES)), UserList(VictimaIndex).Invent.Object(IndexRobo).Amount)

    If UserList(VictimaIndex).Invent.Object(IndexRobo).Equipped = 1 Then Call Desequipar(VictimaIndex, IndexRobo)

    MiObj.Amount = Num

    UserList(VictimaIndex).Invent.Object(IndexRobo).Amount = UserList(VictimaIndex).Invent.Object(IndexRobo).Amount - Num
    If UserList(VictimaIndex).Invent.Object(IndexRobo).Amount <= 0 Then Call QuitarUserInvItem(VictimaIndex, CByte(IndexRobo), 1)

    If Not MeterItemEnInventario(LadronIndex, MiObj) Then Call TirarItemAlPiso(UserList(LadronIndex).POS, MiObj)

    Call SendData(ToIndex, LadronIndex, 0, "||Has robado " & ObjData(MiObj.OBJIndex).Name & " (" & MiObj.Amount & ")." & FONTTYPE_INFO)
    Call UpdateUserInv(False, VictimaIndex, CByte(IndexRobo))

End Sub
Public Sub DoApuñalar(UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)
    Dim Prob As Integer

    Prob = 20 - 1.2 * UserList(UserIndex).Stats.UserSkills(Apuñalar) \ 10

    Select Case UserList(UserIndex).Clase
    Case ASESINO
        Prob = Prob - 3 - Buleano(UserList(UserIndex).Recompensas(3) = 2)
    Case BARDO
        Prob = Prob - 2 - Buleano(UserList(UserIndex).Recompensas(3) = 1)
    End Select

    If RandomNumber(1, Prob) <= 1 Then
        If VictimUserIndex Then
            If UserList(UserIndex).Clase = ASESINO And UserList(UserIndex).Recompensas(3) = 1 Then
                Daño = Daño * 1.7
            Else: Daño = Daño * 1.5
            End If
            If Not UserList(VictimUserIndex).flags.Quest And UserList(VictimUserIndex).flags.Privilegios = 0 Then
                UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Daño
                Call SendUserHP(VictimUserIndex)
            End If
            Call SendData(ToIndex, UserIndex, 0, "5K" & UserList(VictimUserIndex).Name & "," & Daño)
            Call SendData(ToIndex, VictimUserIndex, 0, "5L" & UserList(UserIndex).Name & "," & Daño)
        ElseIf VictimNpcIndex Then
            Select Case UserList(UserIndex).Clase
            Case ASESINO
                Daño = Daño * 2
            Case Else
                Daño = Daño * 1.5
            End Select
            Call SendData(ToIndex, UserIndex, 0, "5M" & Daño)
            Call ExperienciaPorGolpe(UserIndex, VictimNpcIndex, Daño)
            Call VerNPCMuere(VictimNpcIndex, Daño, UserIndex)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "5N")
    End If

End Sub
Public Sub QuitarSta(UserIndex As Integer, Cantidad As Integer)

    If UserList(UserIndex).flags.Quest Or UserList(UserIndex).flags.Privilegios > 2 Then Exit Sub
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0

End Sub
Public Sub DoMineria(UserIndex As Integer, Mineral As Integer)
    On Error GoTo errhandler
    Dim MiObj As Obj
    Dim Esfuerzo As Integer

    If UserList(UserIndex).Clase = MINERO Then
        Esfuerzo = EsfuerzoExcavarMinero
    Else: Esfuerzo = EsfuerzoExcavarGeneral
    End If

    If UserList(UserIndex).Stats.MinSta >= Esfuerzo Then
        Call QuitarSta(UserIndex, Esfuerzo)
        Call SendUserSTA(UserIndex)
    Else
        Call SendData(ToIndex, UserIndex, 0, "9E")
        Exit Sub
    End If

    MiObj.OBJIndex = Mineral



    If UserList(UserIndex).Clase = MINERO Then
        MiObj.Amount = Fix(4 + ((0.29 + 0.07 * Buleano(UserList(UserIndex).Recompensas(1) = 1 And UserList(UserIndex).Invent.HerramientaEqpObjIndex = PICO_EXPERTO)) * UserList(UserIndex).Stats.UserSkills(Mineria)))
    Else: MiObj.Amount = 1
    End If

    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
    Call SubirSkill(UserIndex, Mineria, 5)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SOUND_MINERO)

    Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub
Public Sub DoMeditar(UserIndex As Integer)

    UserList(UserIndex).Counters.IdleCount = Timer

    Dim Suerte As Integer
    Dim Res As Integer
    Dim Cant As Integer

    If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
        Call SendData(ToIndex, UserIndex, 0, "D9")
        Call SendData(ToIndex, UserIndex, 0, "MEDOK")
        UserList(UserIndex).flags.Meditando = False
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
        Exit Sub
    End If

    If UserList(UserIndex).Stats.UserSkills(Meditar) <= 10 _
       And UserList(UserIndex).Stats.UserSkills(Meditar) >= -1 Then
        Suerte = 35
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 20 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 11 Then
        Suerte = 30
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 30 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 21 Then
        Suerte = 28
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 40 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 31 Then
        Suerte = 24
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 50 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 41 Then
        Suerte = 22
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 60 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 51 Then
        Suerte = 20
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 70 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 61 Then
        Suerte = 18
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 80 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 71 Then
        Suerte = 15
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 90 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 81 Then
        Suerte = 10
    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 99 _
           And UserList(UserIndex).Stats.UserSkills(Meditar) >= 91 Then
        Suerte = 8

    ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) = 100 Then

        Suerte = 5
    End If
    Res = RandomNumber(1, Suerte)

    If Res = 1 Then
        If UserList(UserIndex).Stats.MaxMAN > 0 Then Cant = Maximo(1, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3))
        Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Cant, UserList(UserIndex).Stats.MaxMAN)
        Call SendData(ToIndex, UserIndex, 0, "MN" & Cant)
        Call SendUserMANA(UserIndex)
        Call SubirSkill(UserIndex, Meditar)
    End If

End Sub
Public Sub InicioTrabajo(UserIndex As Integer, Trabajo As Long, TrabajoPos As WorldPos)


    If Distancia(TrabajoPos, UserList(UserIndex).POS) > 2 Then
        Call SendData(ToIndex, UserIndex, 0, "DL")
        Exit Sub
    End If


    Select Case Trabajo



    Case Pesca


        If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> OBJTYPE_CAÑA And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> RED_PESCA Then
            Call SendData(ToIndex, UserIndex, 0, "%6")
            Exit Sub
        End If

        If MapData(UserList(UserIndex).POS.Map, TrabajoPos.X, TrabajoPos.Y).Agua = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "6N")
            Exit Sub
        End If

    Case Talar

        If Trabajo = Talar Then
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                Call SendData(ToIndex, UserIndex, 0, "%7")
                Exit Sub
            End If
        End If



        If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).trigger = 4 Then
            Call SendData(ToIndex, UserIndex, 0, "0W")
            Exit Sub
        End If

        If Not ObjData(MapData(TrabajoPos.Map, TrabajoPos.X, TrabajoPos.Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_ARBOLES Then
            Call SendData(ToIndex, UserIndex, 0, "2S")
            Exit Sub
        End If

    Case Mineria

        If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PICO_EXPERTO Then
            Call SendData(ToIndex, UserIndex, 0, "%9")
            Exit Sub
        End If

        If Not ObjData(MapData(TrabajoPos.Map, TrabajoPos.X, TrabajoPos.Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_YACIMIENTO Then
            Call SendData(ToIndex, UserIndex, 0, "7N")
            Exit Sub
        End If

    End Select


    UserList(UserIndex).flags.Trabajando = Trabajo

    UserList(UserIndex).TrabajoPos.X = TrabajoPos.X
    UserList(UserIndex).TrabajoPos.Y = TrabajoPos.Y
    Call SendData(ToIndex, UserIndex, 0, "%0")
    Call SendData(ToIndex, UserIndex, 0, "MT")


End Sub
