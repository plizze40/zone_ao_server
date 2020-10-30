Attribute VB_Name = "InvUsuario"
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
Public Sub AcomodarItems(UserIndex As Integer, Item1 As Byte, Item2 As Byte)
    Dim tObj As UserOBJ
    Dim tObj2 As UserOBJ

    tObj = UserList(UserIndex).Invent.Object(Item1)
    tObj2 = UserList(UserIndex).Invent.Object(Item2)

    UserList(UserIndex).Invent.Object(Item1) = tObj2
    UserList(UserIndex).Invent.Object(Item2) = tObj

    If tObj.Equipped = 1 Then
        Select Case ObjData(tObj.OBJIndex).ObjType
        Case OBJTYPE_WEAPON
            UserList(UserIndex).Invent.WeaponEqpSlot = Item2
        Case OBJTYPE_HERRAMIENTAS
            UserList(UserIndex).Invent.HerramientaEqpslot = Item2
        Case OBJTYPE_BARCOS
            UserList(UserIndex).Invent.BarcoSlot = Item2
        Case OBJTYPE_ARMOUR
            Select Case ObjData(tObj.OBJIndex).SubTipo
            Case OBJTYPE_CASCO
                UserList(UserIndex).Invent.CascoEqpSlot = Item2
            Case OBJTYPE_ARMADURA
                UserList(UserIndex).Invent.ArmourEqpSlot = Item2
            Case OBJTYPE_ESCUDO
                UserList(UserIndex).Invent.EscudoEqpSlot = Item2
            End Select
        Case OBJTYPE_FLECHAS
            UserList(UserIndex).Invent.MunicionEqpSlot = Item2
        End Select
    End If

    If tObj2.Equipped = 1 Then
        Select Case ObjData(tObj2.OBJIndex).ObjType
        Case OBJTYPE_WEAPON
            UserList(UserIndex).Invent.WeaponEqpSlot = Item1
        Case OBJTYPE_HERRAMIENTAS
            UserList(UserIndex).Invent.HerramientaEqpslot = Item1
        Case OBJTYPE_BARCOS
            UserList(UserIndex).Invent.BarcoSlot = Item1
        Case OBJTYPE_ARMOUR
            Select Case ObjData(tObj2.OBJIndex).SubTipo
            Case OBJTYPE_CASCO
                UserList(UserIndex).Invent.CascoEqpSlot = Item1
            Case OBJTYPE_ARMADURA
                UserList(UserIndex).Invent.ArmourEqpSlot = Item1
            Case OBJTYPE_ESCUDO
                UserList(UserIndex).Invent.EscudoEqpSlot = Item1
            End Select
        Case OBJTYPE_FLECHAS
            UserList(UserIndex).Invent.MunicionEqpSlot = Item1
        End Select
    End If

    Call UpdateUserInv(False, UserIndex, Item1)
    Call UpdateUserInv(False, UserIndex, Item2)

End Sub

Public Sub CalcularSta(UserIndex As Integer)

    Select Case UserList(UserIndex).Clase
    Case CIUDADANO, TRABAJADOR, EXPERTO_MINERALES
        UserList(UserIndex).Stats.MaxSta = 15 * UserList(UserIndex).Stats.ELV
    Case MINERO
        UserList(UserIndex).Stats.MaxSta = (15 + AdicionalSTMinero) * UserList(UserIndex).Stats.ELV
    Case HERRERO
        UserList(UserIndex).Stats.MaxSta = 15 * UserList(UserIndex).Stats.ELV
    Case TALADOR
        UserList(UserIndex).Stats.MaxSta = (15 + AdicionalSTLeñador) * UserList(UserIndex).Stats.ELV
    Case CARPINTERO
        UserList(UserIndex).Stats.MaxSta = 15 * UserList(UserIndex).Stats.ELV
    Case PESCADOR
        UserList(UserIndex).Stats.MaxSta = (15 + AdicionalSTPescador) * UserList(UserIndex).Stats.ELV
    Case Is <= 37
        UserList(UserIndex).Stats.MaxSta = 15 * UserList(UserIndex).Stats.ELV
    Case MAGO, NIGROMANTE
        UserList(UserIndex).Stats.MaxSta = (15 - AdicionalSTLadron / 2) * UserList(UserIndex).Stats.ELV
    Case Else
        UserList(UserIndex).Stats.MaxSta = 15 * UserList(UserIndex).Stats.ELV
    End Select

    UserList(UserIndex).Stats.MaxSta = 60 + UserList(UserIndex).Stats.MaxSta
    UserList(UserIndex).Stats.MinSta = Minimo(UserList(UserIndex).Stats.MinSta, UserList(UserIndex).Stats.MaxSta)

End Sub
Public Sub VerObjetosEquipados(UserIndex As Integer)

    With UserList(UserIndex).Invent
        If .CascoEqpSlot Then
            .Object(.CascoEqpSlot).Equipped = 1
            .CascoEqpObjIndex = .Object(.CascoEqpSlot).OBJIndex
            UserList(UserIndex).Char.CascoAnim = ObjData(.CascoEqpObjIndex).CascoAnim
        Else
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        End If

        If .BarcoSlot Then .BarcoObjIndex = .Object(.BarcoSlot).OBJIndex

        If .ArmourEqpSlot Then
            .Object(.ArmourEqpSlot).Equipped = 1
            .ArmourEqpObjIndex = .Object(.ArmourEqpSlot).OBJIndex
            UserList(UserIndex).Char.Body = ObjData(.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If

        If .WeaponEqpSlot Then
            .Object(.WeaponEqpSlot).Equipped = 1
            .WeaponEqpObjIndex = .Object(.WeaponEqpSlot).OBJIndex
            UserList(UserIndex).Char.WeaponAnim = ObjData(.WeaponEqpObjIndex).WeaponAnim
        End If

        If .EscudoEqpSlot Then
            .Object(.EscudoEqpSlot).Equipped = 1
            .EscudoEqpObjIndex = .Object(.EscudoEqpSlot).OBJIndex
            UserList(UserIndex).Char.ShieldAnim = ObjData(.EscudoEqpObjIndex).ShieldAnim
        Else
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        End If

        If .MunicionEqpSlot Then
            .Object(.MunicionEqpSlot).Equipped = 1
            .MunicionEqpObjIndex = .Object(.MunicionEqpSlot).OBJIndex
        End If

        If .HerramientaEqpslot Then
            .Object(.HerramientaEqpslot).Equipped = 1
            .HerramientaEqpObjIndex = .Object(.HerramientaEqpslot).OBJIndex
        End If
    End With

End Sub
Public Function TieneObjetosRobables(UserIndex As Integer) As Boolean
    On Error Resume Next
    Dim i As Byte

    For i = 1 To MAX_INVENTORY_SLOTS
        If ObjEsRobable(UserIndex, i) Then
            TieneObjetosRobables = True
            Exit For
        End If
    Next

End Function
Function ClaseBase(Clase As Byte) As Boolean

    ClaseBase = (Clase = CIUDADANO Or Clase = TRABAJADOR Or Clase = EXPERTO_MINERALES Or _
                 Clase = EXPERTO_MADERA Or Clase = LUCHADOR Or Clase = CON_MANA Or _
                 Clase = HECHICERO Or Clase = ORDEN_SAGRADA Or Clase = NATURALISTA Or _
                 Clase = SIGILOSO Or Clase = SIN_MANA Or Clase = BANDIDO Or _
                 Clase = CABALLERO)

End Function
Function ClaseMana(Clase As Byte) As Boolean

    ClaseMana = (Clase >= CON_MANA And Clase < SIN_MANA)

End Function
Function ClaseNoMana(Clase As Byte) As Boolean

    ClaseNoMana = (Clase >= SIN_MANA)

End Function
Function ClaseTrabajadora(Clase As Byte) As Boolean

    ClaseTrabajadora = (Clase > CIUDADANO And Clase < LUCHADOR)

End Function
Function ClasePuedeHechizo(UserIndex As Integer, ByVal OBJIndex As Integer) As Boolean
    On Error GoTo manejador
    Dim flag As Boolean

    If UserList(UserIndex).flags.Privilegios Then
        ClasePuedeHechizo = True
        Exit Function
    End If

    If ObjData(OBJIndex).ClaseProhibida(1) > 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If ObjData(OBJIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                ClasePuedeHechizo = True
                Exit Function
            End If
        Next
    Else: ClasePuedeHechizo = True
    End If

    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarHechizo")
End Function
Function ClasePuedeUsarItem(UserIndex As Integer, ByVal OBJIndex As Integer) As Boolean
    On Error GoTo manejador
    Dim flag As Boolean

    If UserList(UserIndex).flags.Privilegios Then
        ClasePuedeUsarItem = True
        Exit Function
    End If

    If Len(ObjData(OBJIndex).ClaseProhibida(1)) > 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES

            If ObjData(OBJIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                ClasePuedeUsarItem = False
                Exit Function
            ElseIf ObjData(OBJIndex).ClaseProhibida(i) = 0 Then
                Exit For
            End If
        Next
    End If

    ClasePuedeUsarItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function
Function RazaPuedeUsarItem(UserIndex As Integer, ByVal OBJIndex As Integer) As Boolean
    On Error GoTo manejador
    Dim flag As Boolean

    If UserList(UserIndex).flags.Privilegios Then
        RazaPuedeUsarItem = True
        Exit Function
    End If

    If Len(ObjData(OBJIndex).RazaProhibida(1)) > 0 Then
        Dim i As Integer
        For i = 1 To NUMRAZAS
            If (ObjData(OBJIndex).RazaProhibida(i)) = (UserList(UserIndex).Raza) Then
                RazaPuedeUsarItem = False
                Exit Function
            End If
        Next
        RazaPuedeUsarItem = True
    Else
        RazaPuedeUsarItem = True
    End If

    Exit Function

manejador:
    LogError ("Error en RazaPuedeUsarItem")
End Function
Sub QuitarNewbieObj(UserIndex As Integer)
    Dim j As Byte

    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).OBJIndex Then
            If ObjData(UserList(UserIndex).Invent.Object(j).OBJIndex).Newbie = 1 Then _
               Call QuitarVariosItem(UserIndex, j, MAX_INVENTORY_OBJS)
            Call UpdateUserInv(False, UserIndex, j)
        End If
    Next

End Sub
Sub LimpiarInventario(UserIndex As Integer)
    Dim j As Byte

    For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).OBJIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
    Next

    UserList(UserIndex).Invent.NroItems = 0

    UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
    UserList(UserIndex).Invent.ArmourEqpSlot = 0

    UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
    UserList(UserIndex).Invent.WeaponEqpSlot = 0

    UserList(UserIndex).Invent.CascoEqpObjIndex = 0
    UserList(UserIndex).Invent.CascoEqpSlot = 0

    UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
    UserList(UserIndex).Invent.EscudoEqpSlot = 0

    UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
    UserList(UserIndex).Invent.HerramientaEqpslot = 0

    UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
    UserList(UserIndex).Invent.MunicionEqpSlot = 0

    UserList(UserIndex).Invent.BarcoObjIndex = 0
    UserList(UserIndex).Invent.BarcoSlot = 0

End Sub
Sub TirarOro(ByVal Cantidad As Long, UserIndex As Integer)
    On Error GoTo errhandler

    If UserList(UserIndex).POS.Map = 8 Or _
       UserList(UserIndex).POS.Map = 17 Or _
       UserList(UserIndex).POS.Map = 88 Or _
       UserList(UserIndex).POS.Map = 7 Or _
       UserList(UserIndex).POS.Map = 212 Or _
       UserList(UserIndex).POS.Map = 15 Or _
       UserList(UserIndex).POS.Map = 16 Or _
       UserList(UserIndex).POS.Map = 12 Or _
       UserList(UserIndex).POS.Map = 10 Or _
       UserList(UserIndex).POS.Map = 11 Or _
       UserList(UserIndex).POS.Map = 14 Then Exit Sub

    Dim nPos As WorldPos
    If Cantidad > 100000 Then Exit Sub

    If Cantidad <= 0 Or Cantidad > UserList(UserIndex).Stats.GLD Then Exit Sub

    Dim MiObj As Obj

    MiObj.OBJIndex = iORO

    If UserList(UserIndex).flags.Privilegios Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & Cantidad & " Objeto:" & ObjData(MiObj.OBJIndex).Name, False)

    Do While Cantidad > 0
        MiObj.Amount = Minimo(Cantidad, MAX_INVENTORY_OBJS)

        nPos = TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
        If nPos.Map = 0 Then Exit Sub

        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
        Cantidad = Cantidad - MiObj.Amount
    Loop

    Exit Sub

errhandler:

End Sub
Sub QuitarUserInvItem(UserIndex As Integer, ByVal Slot As Byte, Cantidad As Integer)
    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)

    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad

    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
    End If

End Sub
Sub QuitarUnItem(UserIndex As Integer, ByVal Slot As Byte)
    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 And UserList(UserIndex).Invent.Object(Slot).Amount = 1 Then Call Desequipar(UserIndex, Slot)

    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - 1

    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        Call SendData(ToIndex, UserIndex, 0, "3I" & Slot)
    Else
        Call SendData(ToIndex, UserIndex, 0, "2I" & Slot)
    End If

End Sub
Sub QuitarBebida(UserIndex As Integer, ByVal Slot As Byte)

    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)


    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - 1


    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        Call SendData(ToIndex, UserIndex, 0, "6I" & Slot & "," & UserList(UserIndex).Stats.MinAGU)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")

    Else
        Call SendData(ToIndex, UserIndex, 0, "6J" & Slot & "," & UserList(UserIndex).Stats.MinAGU)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")

    End If

End Sub
Sub QuitarComida(UserIndex As Integer, ByVal Slot As Byte)

    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)


    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - 1


    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        Call SendData(ToIndex, UserIndex, 0, "7K" & Slot & "," & UserList(UserIndex).Stats.MinHam)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "7")

    Else
        Call SendData(ToIndex, UserIndex, 0, "6K" & Slot & "," & UserList(UserIndex).Stats.MinHam)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "7")

    End If

End Sub

Sub QuitarPocion(UserIndex As Integer, ByVal Slot As Byte)

    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)


    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - 1

    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        Call SendData(ToIndex, UserIndex, 0, "4J" & Slot)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")

    Else
        Call SendData(ToIndex, UserIndex, 0, "3J" & Slot)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")
    End If

End Sub

Sub QuitarPocionMana(UserIndex As Integer, ByVal Slot As Byte)

    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)


    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - 1


    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        Call SendData(ToIndex, UserIndex, 0, "8I" & Slot & "," & UserList(UserIndex).Stats.MinMAN)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")

    Else
        Call SendData(ToIndex, UserIndex, 0, "7I" & Slot & "," & UserList(UserIndex).Stats.MinMAN)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")

    End If

End Sub
Sub QuitarPocionVida(UserIndex As Integer, ByVal Slot As Byte)

    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)


    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - 1

    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        Call SendData(ToIndex, UserIndex, 0, "2J" & Slot & "," & UserList(UserIndex).Stats.MinHP)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")

    Else
        Call SendData(ToIndex, UserIndex, 0, "9I" & Slot & "," & UserList(UserIndex).Stats.MinHP)
        Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).POS.Map, "TW" & "46")

    End If

End Sub
Sub QuitarVariosItem(UserIndex As Integer, ByVal Slot As Byte, Cantidad As Integer)

    Dim MiObj As Obj

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 And UserList(UserIndex).Invent.Object(Slot).Amount <= Cantidad Then Call Desequipar(UserIndex, Slot)


    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad


    If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        Call SendData(ToIndex, UserIndex, 0, "3I" & Slot)
    Else
        Call SendData(ToIndex, UserIndex, 0, "4I" & Slot & "," & Cantidad)
    End If

End Sub
Sub UpdateUserInv(ByVal UpdateAll As Boolean, UserIndex As Integer, Slot As Byte, Optional JustAmount As Boolean)
    Dim i As Byte

    If UpdateAll Then
        For i = 1 To MAX_INVENTORY_SLOTS
            Call SendUserItem(UserIndex, i, JustAmount)
        Next
    Else
        Call SendUserItem(UserIndex, Slot, JustAmount)
    End If

End Sub
Sub DropObj(UserIndex As Integer, Slot As Byte, ByVal Num As Integer, Map As Integer, X As Integer, Y As Integer)
    Dim Obj As Obj

    If Num Then
        If Num > UserList(UserIndex).Invent.Object(Slot).Amount Then Num = UserList(UserIndex).Invent.Object(Slot).Amount


        If MapData(UserList(UserIndex).POS.Map, X, Y).OBJInfo.OBJIndex = 0 Then
            If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 And UserList(UserIndex).Invent.Object(Slot).Amount <= Num Then Call Desequipar(UserIndex, Slot)
            Obj.OBJIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
            If UserList(UserIndex).flags.Privilegios < 2 Then
                If ObjData(Obj.OBJIndex).NoComerciable = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "2W")
                    Exit Sub
                End If

                If ObjData(Obj.OBJIndex).Newbie = 1 And EsNewbie(UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "3W")
                    Exit Sub
                End If
            End If

            Obj.Amount = Num

            Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
            Call QuitarVariosItem(UserIndex, Slot, Num)

            If UserList(UserIndex).flags.Privilegios Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & Num & " Objeto:" & ObjData(Obj.OBJIndex).Name, False)
        Else
            Call SendData(ToIndex, UserIndex, 0, "4W")
        End If

    End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal Num As Integer, Map As Integer, X As Integer, Y As Integer)

    MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - Num

    If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
        MapData(Map, X, Y).OBJInfo.OBJIndex = 0
        MapData(Map, X, Y).OBJInfo.Amount = 0
        Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
    End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, X As Integer, Y As Integer)


    MapData(Map, X, Y).OBJInfo = Obj
    Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.OBJIndex).GrhIndex & "," & X & "," & Y)

End Sub

Function MeterItemEnInventario(UserIndex As Integer, MiObj As Obj) As Boolean
    On Error GoTo errhandler



    Dim X As Integer
    Dim Y As Integer
    Dim Slot As Byte


    Slot = 1
    Do Until UserList(UserIndex).Invent.Object(Slot).OBJIndex = MiObj.OBJIndex And _
       UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
        Slot = Slot + 1
        If Slot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop


    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
            Slot = Slot + 1
            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "5W")
                MeterItemEnInventario = False
                Exit Function
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
    End If


    If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then

        UserList(UserIndex).Invent.Object(Slot).OBJIndex = MiObj.OBJIndex
        UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
    Else
        UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
    End If

    MeterItemEnInventario = True

    Call UpdateUserInv(False, UserIndex, Slot)


    Exit Function
errhandler:

End Function


Sub GetObj(UserIndex As Integer)

    Dim Obj As ObjData
    Dim MiObj As Obj


    If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).OBJInfo.OBJIndex Then

        If ObjData(MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).OBJInfo.OBJIndex).Agarrable <> 1 Then
            Dim X As Integer
            Dim Y As Integer
            Dim Slot As Byte

            X = UserList(UserIndex).POS.X
            Y = UserList(UserIndex).POS.Y
            Obj = ObjData(MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).OBJInfo.OBJIndex)
            MiObj.Amount = MapData(UserList(UserIndex).POS.Map, X, Y).OBJInfo.Amount
            MiObj.OBJIndex = MapData(UserList(UserIndex).POS.Map, X, Y).OBJInfo.OBJIndex

            If ObjData(MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_GUITA Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MapData(UserList(UserIndex).POS.Map, X, Y).OBJInfo.Amount
                Call SendUserORO(UserIndex)
                Call EraseObj(ToMap, 0, UserList(UserIndex).POS.Map, MapData(UserList(UserIndex).POS.Map, X, Y).OBJInfo.Amount, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
                If UserList(UserIndex).flags.Privilegios Then Call LogGM(UserList(UserIndex).Name, "Agarro oro:" & MiObj.Amount, False)

                Exit Sub
            End If


            If Not MeterItemEnInventario(UserIndex, MiObj) Then

            Else

                Call EraseObj(ToMap, 0, UserList(UserIndex).POS.Map, MapData(UserList(UserIndex).POS.Map, X, Y).OBJInfo.Amount, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
                If UserList(UserIndex).flags.Privilegios Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.OBJIndex).Name, False)
            End If

        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "8K")
    End If

End Sub
Sub Desequipar(UserIndex As Integer, ByVal Slot As Byte)



    Dim Obj As ObjData
    If Slot = 0 Then Exit Sub
    If UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0 Then Exit Sub

    Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex)

    Select Case Obj.ObjType


    Case OBJTYPE_WEAPON

        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0

        Call ChangeUserArma(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, NingunArma)

    Case OBJTYPE_FLECHAS

        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0

    Case OBJTYPE_HERRAMIENTAS

        If UserList(UserIndex).flags.Trabajando Then
            If UserList(UserIndex).flags.CodigoTrabajo Then
                Exit Sub
            Else
                Call SacarModoTrabajo(UserIndex)
            End If
        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpslot = 0

    Case OBJTYPE_ARMOUR
        If UserList(UserIndex).flags.Montado = 1 Then Exit Sub

        Select Case Obj.SubTipo

        Case OBJTYPE_ARMADURA
            UserList(UserIndex).Invent.Object(Slot).Equipped = 0
            UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
            UserList(UserIndex).Invent.ArmourEqpSlot = 0
            If UserList(UserIndex).flags.Transformado = 0 Then
                Call DarCuerpoDesnudo(UserIndex)
                Call ChangeUserBody(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, UserList(UserIndex).Char.Body)
            End If

        Case OBJTYPE_CASCO
            UserList(UserIndex).Invent.Object(Slot).Equipped = 0
            UserList(UserIndex).Invent.CascoEqpObjIndex = 0
            UserList(UserIndex).Invent.CascoEqpSlot = 0
            If UserList(UserIndex).flags.Transformado = 0 Then
                Call ChangeUserCasco(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, NingunCasco)
            End If
        Case OBJTYPE_ESCUDO
            UserList(UserIndex).Invent.Object(Slot).Equipped = 0
            UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
            UserList(UserIndex).Invent.EscudoEqpSlot = 0
            If UserList(UserIndex).flags.Transformado = 0 Then
                Call ChangeUserEscudo(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, NingunEscudo)
            End If
        End Select

    End Select

    Call DesequiparItem(UserIndex, Slot)

End Sub
Function SexoPuedeUsarItem(UserIndex As Integer, ByVal OBJIndex As Integer) As Boolean
    On Error GoTo errhandler

    If UserList(UserIndex).flags.Privilegios Then
        SexoPuedeUsarItem = True
        Exit Function
    End If

    If ObjData(OBJIndex).MUJER = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero = MUJER
    ElseIf ObjData(OBJIndex).HOMBRE = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero = HOMBRE
    Else
        SexoPuedeUsarItem = True
    End If

    Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function
Function FaccionClasePuedeUsarItem(UserIndex As Integer, ByVal OBJIndex As Integer) As Boolean
    Dim i As Integer

    If UserList(UserIndex).flags.Privilegios Then
        FaccionClasePuedeUsarItem = True
        Exit Function
    End If

    For i = 1 To Minimo(UserList(UserIndex).Faccion.Jerarquia, 3)
        If Armaduras(UserList(UserIndex).Faccion.Bando, i, TipoClase(UserIndex), TipoRaza(UserIndex)) = OBJIndex Then
            FaccionClasePuedeUsarItem = True
            Exit Function
        End If
    Next

End Function
Function FaccionPuedeUsarItem(UserIndex As Integer, ByVal OBJIndex As Integer) As Boolean

    If UserList(UserIndex).flags.Privilegios Then
        FaccionPuedeUsarItem = True
        Exit Function
    End If

    If ObjData(OBJIndex).Real >= 1 Then
        FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.Bando = Real And UserList(UserIndex).Faccion.Jerarquia >= ObjData(OBJIndex).Jerarquia)
    ElseIf ObjData(OBJIndex).Caos >= 1 Then
        FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.Bando = Caos And UserList(UserIndex).Faccion.Jerarquia >= ObjData(OBJIndex).Jerarquia)
    Else: FaccionPuedeUsarItem = True
    End If

End Function
Function PuedeUsarObjeto(UserIndex As Integer, ByVal OBJIndex As Integer) As Byte

    Select Case ObjData(OBJIndex).ObjType
    Case OBJTYPE_WEAPON

        If Not (OBJIndex = 367 And UserList(UserIndex).Clase = ASESINO) Then
            If Not RazaPuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 5
                Exit Function
            End If

            If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 2
                Exit Function
            End If
        End If

        If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
            PuedeUsarObjeto = 4
            Exit Function
        End If

    Case OBJTYPE_HERRAMIENTAS

        If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
            PuedeUsarObjeto = 2
            Exit Function
        End If

    Case OBJTYPE_ARMOUR

        Select Case ObjData(OBJIndex).SubTipo

        Case OBJTYPE_ARMADURA

            If Not RazaPuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 5
                Exit Function
            End If

            If Not SexoPuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 1
                Exit Function
            End If

            If ObjData(OBJIndex).Real = 0 And ObjData(OBJIndex).Caos = 0 Then
                If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                    PuedeUsarObjeto = 2
                    Exit Function
                End If
            Else
                If Not FaccionPuedeUsarItem(UserIndex, OBJIndex) Then
                    PuedeUsarObjeto = 3
                    Exit Function
                End If
                If Not FaccionClasePuedeUsarItem(UserIndex, OBJIndex) Then
                    PuedeUsarObjeto = 2
                    Exit Function
                End If
            End If

            If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 4
                Exit Function
            End If

        Case OBJTYPE_CASCO

            If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 2
                Exit Function
            End If

            If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 4
                Exit Function
            End If

        Case OBJTYPE_ESCUDO

            If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 2
                Exit Function
            End If

            If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 4
                Exit Function
            End If

        Case OBJTYPE_PERGAMINOS
            If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                PuedeUsarObjeto = 2
                Exit Function
            End If

        End Select
    End Select

    PuedeUsarObjeto = 0

End Function
Function SkillPuedeUsarItem(UserIndex As Integer, ByVal OBJIndex As Integer) As Boolean

    If UserList(UserIndex).flags.Privilegios Then
        SkillPuedeUsarItem = True
        Exit Function
    End If

    If ObjData(OBJIndex).SkillCombate > UserList(UserIndex).Stats.UserSkills(Armas) Then Exit Function
    If ObjData(OBJIndex).SkillApuñalar > UserList(UserIndex).Stats.UserSkills(Apuñalar) Then Exit Function
    If ObjData(OBJIndex).SkillProyectiles > UserList(UserIndex).Stats.UserSkills(Proyectiles) Then Exit Function
    If ObjData(OBJIndex).SkResistencia > UserList(UserIndex).Stats.UserSkills(Resis) Then Exit Function
    If ObjData(OBJIndex).SkDefensa > UserList(UserIndex).Stats.UserSkills(Defensa) Then Exit Function
    If ObjData(OBJIndex).SkillTacticas > UserList(UserIndex).Stats.UserSkills(Tacticas) Then Exit Function

    SkillPuedeUsarItem = True

End Function
Sub EquiparInvItem(UserIndex As Integer, Slot As Byte)
    On Error GoTo errhandler


    Dim Obj As ObjData
    Dim OBJIndex As Integer

    OBJIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
    Obj = ObjData(OBJIndex)

    If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "6W")
        Exit Sub
    End If

    Select Case Obj.ObjType
    Case OBJTYPE_WEAPON

        If Not (OBJIndex = 367 And UserList(UserIndex).Clase = ASESINO) Then
            If Not RazaPuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "8W")
                Exit Sub
            End If

            If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "2X")
                Exit Sub
            End If
        End If

        If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "7W")
            Exit Sub
        End If

        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            Call Desequipar(UserIndex, Slot)
            Exit Sub
        End If


        If UserList(UserIndex).Invent.WeaponEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
        UserList(UserIndex).Invent.WeaponEqpSlot = Slot


        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil > 0 And UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
            Call ChangeUserEscudo(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, 0)
        End If

        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SOUND_SACARARMA)

        Call ChangeUserArma(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, Obj.WeaponAnim)
        Call EquiparItem(UserIndex, Slot)

    Case OBJTYPE_HERRAMIENTAS
        If Not RazaPuedeUsarItem(UserIndex, OBJIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "8W")
            Exit Sub
        End If


        If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "2X")
            Exit Sub
        End If


        If UserList(UserIndex).Invent.Object(Slot).Equipped Then

            Call Desequipar(UserIndex, Slot)
            Exit Sub
        End If


        If UserList(UserIndex).Invent.HerramientaEqpObjIndex Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpslot)
        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = OBJIndex
        UserList(UserIndex).Invent.HerramientaEqpslot = Slot
        Call EquiparItem(UserIndex, Slot)

    Case OBJTYPE_FLECHAS


        If UserList(UserIndex).Invent.Object(Slot).Equipped Then

            Call Desequipar(UserIndex, Slot)
            Exit Sub
        End If


        If UserList(UserIndex).Invent.MunicionEqpObjIndex Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
        End If

        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
        UserList(UserIndex).Invent.MunicionEqpSlot = Slot
        Call EquiparItem(UserIndex, Slot)

    Case OBJTYPE_ARMOUR

        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub

        Select Case Obj.SubTipo

        Case OBJTYPE_ARMADURA

            If Not RazaPuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "8W")
                Exit Sub
            End If

            If Not SexoPuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "8W")
                Exit Sub
            End If

            If ObjData(OBJIndex).Real = 0 And ObjData(OBJIndex).Caos = 0 Then
                If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "2X")
                    Exit Sub
                End If
            Else
                If Not FaccionPuedeUsarItem(UserIndex, OBJIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "%?")
                    Exit Sub
                End If
                If Not FaccionClasePuedeUsarItem(UserIndex, OBJIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||Tu clase o raza no puede usar ese objeto." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If

            If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "7W")
                Exit Sub
            End If


            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Exit Sub
            End If


            If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
            End If


            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
            UserList(UserIndex).Invent.ArmourEqpSlot = Slot

            UserList(UserIndex).flags.Desnudo = 0

            If UserList(UserIndex).flags.Transformado = 0 Then Call ChangeUserBody(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, Obj.Ropaje)
            Call EquiparItem(UserIndex, Slot)

        Case OBJTYPE_CASCO

            If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "2X")
                Exit Sub
            End If

            If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "7W")
                Exit Sub
            End If

            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Exit Sub
            End If


            If UserList(UserIndex).Invent.CascoEqpObjIndex Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
            End If



            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
            UserList(UserIndex).Invent.CascoEqpSlot = Slot

            Call ChangeUserCasco(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, Obj.CascoAnim)
            Call EquiparItem(UserIndex, Slot)

        Case OBJTYPE_ESCUDO

            If Not ClasePuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "2X")
                Exit Sub
            End If

            If Not SkillPuedeUsarItem(UserIndex, OBJIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "7W")
                Exit Sub
            End If

            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Exit Sub
            End If


            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Exit Sub
            End If


            If UserList(UserIndex).Invent.EscudoEqpObjIndex Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)



            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
            UserList(UserIndex).Invent.EscudoEqpSlot = Slot

            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                Call ChangeUserArma(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, 0)
            End If

            Call ChangeUserEscudo(ToMap, 0, UserList(UserIndex).POS.Map, UserIndex, Obj.ShieldAnim)
            Call EquiparItem(UserIndex, Slot)

        End Select
    End Select


    Exit Sub
errhandler:
    Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(UserIndex As Integer, ItemIndex As Integer) As Boolean
    On Error GoTo errhandler


    If UserList(UserIndex).Raza = HUMANO Or _
       UserList(UserIndex).Raza = ELFO Or _
       UserList(UserIndex).Raza = ELFO_OSCURO Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
    End If


    Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function
Sub SacarModoTrabajo(UserIndex As Integer)

    UserList(UserIndex).flags.Trabajando = 0
    UserList(UserIndex).TrabajoPos.X = 0
    UserList(UserIndex).TrabajoPos.Y = 0
    UserList(UserIndex).flags.CodigoTrabajo = 0

    Call SendData(ToIndex, UserIndex, 0, "%I")
    Call SendData(ToIndex, UserIndex, 0, "MT")

End Sub
Sub UseInvItem(UserIndex As Integer, Slot As Byte, ByVal Click As Byte)
    Dim Obj As ObjData
    Dim OBJIndex As Integer
    Dim TargObj As ObjData
    Dim MiObj As Obj

    Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex)

    If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "6W")
        Exit Sub
    End If

    OBJIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
    UserList(UserIndex).flags.TargetObjInvIndex = OBJIndex
    UserList(UserIndex).flags.TargetObjInvslot = Slot

    Select Case Obj.ObjType

    Case OBJTYPE_USEONCE
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        Call AddtoVar(UserList(UserIndex).Stats.MinHam, Obj.MinHam, UserList(UserIndex).Stats.MaxHam)
        UserList(UserIndex).flags.Hambre = 0

        Call QuitarComida(UserIndex, Slot)

    Case OBJTYPE_GUITA
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        Call SendUserORO(UserIndex)

    Case OBJTYPE_WEAPON
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If ObjData(OBJIndex).proyectil = 1 Then
            If TiempoTranscurrido(UserList(UserIndex).Counters.LastFlecha) < IntervaloUserFlechas Then Exit Sub
            If TiempoTranscurrido(UserList(UserIndex).Counters.LastHechizo) < IntervaloUserPuedeHechiGolpe Then Exit Sub
            Call SendData(ToIndex, UserIndex, 0, "T01" & Proyectiles)
        Else
            If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
            If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_LEÑA And UserList(UserIndex).Invent.Object(Slot).OBJIndex = DAGA Then Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
        End If

    Case OBJTYPE_POCIONES
        If TiempoTranscurrido(UserList(UserIndex).Counters.LastGolpe) < (IntervaloUserPuedeAtacar / 2) Then
            Call SendData(ToIndex, UserIndex, 0, "6X")
            Exit Sub
        End If

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If



        Select Case Obj.TipoPocion

        Case 1
            UserList(UserIndex).flags.DuracionEfecto = Timer
            UserList(UserIndex).flags.TomoPocion = True

            Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(Obj.MinModificador, Obj.MaxModificador), Minimo(UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
            Call UpdateFuerzaYAg(UserIndex)

            Call QuitarPocion(UserIndex, Slot)


        Case 2
            UserList(UserIndex).flags.DuracionEfecto = Timer
            UserList(UserIndex).flags.TomoPocion = True

            Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(fuerza), RandomNumber(Obj.MinModificador, Obj.MaxModificador), Minimo(UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
            Call UpdateFuerzaYAg(UserIndex)

            Call QuitarPocion(UserIndex, Slot)

        Case 3

            AddtoVar UserList(UserIndex).Stats.MinHP, RandomNumber(Obj.MinModificador, Obj.MaxModificador), UserList(UserIndex).Stats.MaxHP


            Call QuitarPocionVida(UserIndex, Slot)






        Case 4

            Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, Obj.MaxModificador), UserList(UserIndex).Stats.MaxMAN)


            Call QuitarPocionMana(UserIndex, Slot)
        Case 5
            If UserList(UserIndex).flags.Envenenado = 1 Then
                UserList(UserIndex).flags.Envenenado = 0
                Call SendData(ToIndex, UserIndex, 0, "8X")
            End If

            Call QuitarPocion(UserIndex, Slot)

        End Select

    Case OBJTYPE_BEBIDA
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0


        Call QuitarBebida(UserIndex, Slot)

    Case OBJTYPE_LLAVES
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)

        If TargObj.ObjType = OBJTYPE_PUERTAS Then

            If TargObj.Cerrada = 1 Then

                If TargObj.Llave Then
                    If TargObj.Clave = Obj.Clave Then

                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.OBJIndex _
                              = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.OBJIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.OBJIndex
                        Call SendData(ToIndex, UserIndex, 0, "9X")
                        Exit Sub
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "2Y")
                        Exit Sub
                    End If
                Else
                    If TargObj.Clave = Obj.Clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.OBJIndex _
                              = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.OBJIndex).IndexCerradaLlave
                        Call SendData(ToIndex, UserIndex, 0, "3Y")
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.OBJIndex
                        Exit Sub
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "2Y")
                        Exit Sub
                    End If
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "4Y")
                Exit Sub
            End If

        End If

    Case OBJTYPE_BOTELLAVACIA
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Agua = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "9F")
            Exit Sub
        End If
        MiObj.Amount = 1
        MiObj.OBJIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).IndexAbierta
        Call QuitarUnItem(UserIndex, Slot)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
        End If

    Case OBJTYPE_BOTELLALLENA
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call EnviarHyS(UserIndex)
        MiObj.Amount = 1
        MiObj.OBJIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).IndexCerrada
        Call QuitarUnItem(UserIndex, Slot)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
        End If

    Case OBJTYPE_HERRAMIENTAS

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If UserList(UserIndex).Stats.MinSta = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "9E")
            Exit Sub
        End If

        If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "%J")
            Exit Sub
        End If

        If UserList(UserIndex).flags.Trabajando Then
            Call SendData(ToIndex, UserIndex, 0, "%K")
            Exit Sub
        End If

        Select Case OBJIndex
        Case OBJTYPE_CAÑA, RED_PESCA
            Call SendData(ToIndex, UserIndex, 0, "T01" & Pesca)
        Case HACHA_LEÑADOR
            Call SendData(ToIndex, UserIndex, 0, "T01" & Talar)
        Case PIQUETE_MINERO, PICO_EXPERTO
            Call SendData(ToIndex, UserIndex, 0, "T01" & Mineria)
        Case MARTILLO_HERRERO
            Call SendData(ToIndex, UserIndex, 0, "T01" & Herreria)
        Case SERRUCHO_CARPINTERO
            Call EnviarObjConstruibles(UserIndex)
            Call SendData(ToIndex, UserIndex, 0, "SFC")
        Case HILAR_SASTRE
            Call EnviarRopasConstruibles(UserIndex)
            Call SendData(ToIndex, UserIndex, 0, "SFS")

        End Select

    Case OBJTYPE_WARP

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU" & FONTTYPE_INFO)
            Exit Sub
        End If
        If Not UserList(UserIndex).flags.TargetNpcTipo = 6 Then
            Call SendData(ToIndex, UserIndex, 0, "5Y")
            Exit Sub
        Else
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).POS, UserList(UserIndex).POS) > 4 Then
                Call SendData(ToIndex, UserIndex, 0, "6Y")
                Exit Sub
            Else
                If val(Obj.WI) = val(UserList(UserIndex).POS.Map) Then
                    Call WarpUserChar(UserIndex, Obj.WMapa, Obj.WX, Obj.WY, True)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & SND_WARP)
                    Call UpdateUserInv(False, UserIndex, Slot)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Ese pasaje no te lo he vendido yo, lárgate!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                    Exit Sub
                End If
            End If
        End If

    Case OBJTYPE_PERGAMINOS
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If

        If Not ClasePuedeHechizo(UserIndex, UserList(UserIndex).Invent.Object(Slot).OBJIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede aprender este hechizo." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).flags.Hambre = 0 And _
           UserList(UserIndex).flags.Sed = 0 Then
            Call AgregarHechizo(UserIndex, Slot)
            Call UpdateUserInv(False, UserIndex, Slot)
        Else
            Call SendData(ToIndex, UserIndex, 0, "7F")
        End If

    Case OBJTYPE_MINERALES
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "T01" & FundirMetal)

    Case OBJTYPE_INSTRUMENTOS
        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "MU")
            Exit Sub
        End If
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "TW" & Obj.Snd1)

    Case OBJTYPE_BARCOS
        If UserList(UserIndex).flags.Montado = 1 Then Exit Sub

        If ((LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X - 1, UserList(UserIndex).POS.Y, True) Or _
             LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y - 1, True) Or _
             LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X + 1, UserList(UserIndex).POS.Y, True) Or _
             LegalPos(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y + 1, True)) And _
             UserList(UserIndex).flags.Navegando = 0) _
             Or UserList(UserIndex).flags.Navegando = 1 Then
            Call DoNavega(UserIndex, CInt(Slot))
        Else
            Call SendData(ToIndex, UserIndex, 0, "2G")
        End If

    End Select

End Sub
Sub EnviarArmasConstruibles(UserIndex As Integer)
    Dim i As Integer, cad As String
    Dim Descuento As Single

    If UserList(UserIndex).Clase = HERRERO And UserList(UserIndex).Recompensas(3) = 1 Then
        Descuento = 0.75
    Else: Descuento = 1
    End If

    For i = 1 To UBound(ArmasHerrero)
        If ObjData(ArmasHerrero(i).Index).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).Clase) Then
            If ArmasHerrero(i).Recompensa = 0 Or UserList(UserIndex).Recompensas(2) = 1 Then
                cad = cad & ObjData(ArmasHerrero(i).Index).Name & " (" & ObjData(ArmasHerrero(i).Index).MinHit & "/" & ObjData(ArmasHerrero(i).Index).MaxHit & ")" & " - (" & Int(val(ObjData(ArmasHerrero(i).Index).LingH * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(ArmasHerrero(i).Index).LingP * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(ArmasHerrero(i).Index).LingO * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & ")" _
                    & "," & ArmasHerrero(i).Index & ","
            End If
        End If
    Next

    Call SendData(ToIndex, UserIndex, 0, "LAH" & cad)

End Sub
Sub EnviarObjConstruibles(UserIndex As Integer)
    Dim i As Integer, cad As String, Coste As Integer

    For i = 1 To UBound(ObjCarpintero)
        If ObjData(ObjCarpintero(i).Index).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(Carpinteria) / ModCarpinteria(UserList(UserIndex).Clase) Then
            If ObjCarpintero(i).Recompensa = 0 Or (UserList(UserIndex).Clase = CARPINTERO And UserList(UserIndex).Recompensas(1) = ObjCarpintero(i).Recompensa) Then
                Coste = ObjData(ObjCarpintero(i).Index).Madera
                If UserList(UserIndex).Clase = CARPINTERO And UserList(UserIndex).Recompensas(2) = 2 And ObjData(ObjCarpintero(i).Index).ObjType = OBJTYPE_BARCOS Then Coste = Coste * 0.8
                cad = cad & ObjData(ObjCarpintero(i).Index).Name & " (" & CLng(Coste * ModMadera(UserList(UserIndex).Clase)) & ") - (" & CLng(val(ObjData(ObjCarpintero(i).Index).MaderaElfica) * ModMadera(UserList(UserIndex).Clase)) & ")" & "," & ObjCarpintero(i).Index & ","
            End If
        End If
    Next

    Call SendData(ToIndex, UserIndex, 0, "OBR" & cad)

End Sub
Sub EnviarRopasConstruibles(UserIndex As Integer)
    Dim PielP As Integer, PielL As Integer, PielO As Integer
    Dim N As Integer

    Dim i As Integer, cad As String
    N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

    For i = 1 To UBound(ObjSastre)
        If ObjData(ObjSastre(i)).SkSastreria <= UserList(UserIndex).Stats.UserSkills(Sastreria) / ModRopas(UserList(UserIndex).Clase) Then
            PielP = ObjData(ObjSastre(i)).PielOsoPolar
            PielL = ObjData(ObjSastre(i)).PielLobo
            PielO = ObjData(ObjSastre(i)).PielOsoPardo
            If UserList(UserIndex).Clase = SASTRE And UserList(UserIndex).Stats.ELV >= 18 Then
                PielL = PielL * 0.8
                PielO = PielO * 0.8
                PielP = PielP * 0.8
            End If
            cad = cad & ObjData(ObjSastre(i)).Name & " (" & ObjData(ObjSastre(i)).MinDef & "/" & ObjData(ObjSastre(i)).MaxDef & ")" & " - (" & CLng(PielL * ModSastre(UserList(UserIndex).Clase)) & "/" & CLng(PielO * ModSastre(UserList(UserIndex).Clase)) & "/" & CLng(PielP * ModSastre(UserList(UserIndex).Clase)) & ")" & "," & ObjSastre(i) & ","
        End If
    Next

    Call SendData(ToIndex, UserIndex, 0, "SAR" & cad)

End Sub
Sub EnviarArmadurasConstruibles(UserIndex As Integer)
    Dim i As Integer, cad As String
    Dim Descuento As Single

    If UserList(UserIndex).Clase = HERRERO And UserList(UserIndex).Recompensas(3) = 1 Then
        Descuento = 0.75
    Else: Descuento = 1
    End If

    For i = 1 To UBound(ArmadurasHerrero)
        If ObjData(ArmadurasHerrero(i).Index).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).Clase) Then
            If ArmadurasHerrero(i).Recompensa = 0 Or UserList(UserIndex).Recompensas(2) = 2 Then
                cad = cad & ObjData(ArmadurasHerrero(i).Index).Name & " (" & ObjData(ArmadurasHerrero(i).Index).MinDef & "/" & ObjData(ArmadurasHerrero(i).Index).MaxDef & ")" & " - (" & Int(val(ObjData(ArmadurasHerrero(i).Index).LingH * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(ArmadurasHerrero(i).Index).LingP * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(ArmadurasHerrero(i).Index).LingO * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & ")" _
                    & "," & ArmadurasHerrero(i).Index & ","
            End If
        End If
    Next

    Call SendData(ToIndex, UserIndex, 0, "LAR" & cad)


End Sub
Sub EnviarCascosConstruibles(UserIndex As Integer)
    Dim i As Integer, cad As String
    Dim Descuento As Single

    If UserList(UserIndex).Clase = HERRERO And UserList(UserIndex).Recompensas(1) = 2 Then
        Descuento = 0.5
    Else: Descuento = 1
    End If

    For i = 1 To UBound(CascosHerrero)
        If ObjData(CascosHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).Clase) Then
            cad = cad & ObjData(CascosHerrero(i)).Name & " (" & ObjData(CascosHerrero(i)).MinDef & "/" & ObjData(CascosHerrero(i)).MaxDef & ")" & " - (" & Int(val(ObjData(CascosHerrero(i)).LingH * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(CascosHerrero(i)).LingP * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(CascosHerrero(i)).LingO * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & ")" _
                & "," & CascosHerrero(i) & ","
        End If
    Next

    Call SendData(ToIndex, UserIndex, 0, "CAS" & cad)

End Sub
Sub EnviarEscudosConstruibles(UserIndex As Integer)
    Dim i As Integer, cad As String
    Dim Descuento As Single

    If UserList(UserIndex).Clase = HERRERO And UserList(UserIndex).Recompensas(1) = 2 Then
        Descuento = 0.5
    Else: Descuento = 1
    End If

    For i = 1 To UBound(EscudosHerrero)
        If ObjData(EscudosHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).Clase) Then
            cad = cad & ObjData(EscudosHerrero(i)).Name & " (" & ObjData(EscudosHerrero(i)).MinDef & "/" & ObjData(EscudosHerrero(i)).MaxDef & ") - (" & Int(val(ObjData(EscudosHerrero(i)).LingH * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(EscudosHerrero(i)).LingP * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & "/" & Int(val(ObjData(EscudosHerrero(i)).LingO * Descuento) * ModMateriales(UserList(UserIndex).Clase)) & ")" _
                & "," & EscudosHerrero(i) & ","
        End If
    Next

    Call SendData(ToIndex, UserIndex, 0, "ESC" & cad)



End Sub
Sub TirarTodo(UserIndex As Integer)
    On Error Resume Next

    If UserList(UserIndex).POS.Map = 8 Or _
       UserList(UserIndex).POS.Map = 17 Or _
       UserList(UserIndex).POS.Map = 88 Or _
       UserList(UserIndex).POS.Map = 1 Or _
       UserList(UserIndex).POS.Map = 212 Or _
       UserList(UserIndex).POS.Map = 7 Or _
       UserList(UserIndex).POS.Map = 15 Or _
       UserList(UserIndex).POS.Map = 16 Or _
       UserList(UserIndex).POS.Map = 12 Or _
       UserList(UserIndex).POS.Map = 10 Or _
       UserList(UserIndex).POS.Map = 11 Or _
       UserList(UserIndex).POS.Map = 14 Then Exit Sub

    Call TirarTodosLosItems(UserIndex)
    Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)

End Sub
Public Function ItemSeCae(Index As Integer) As Boolean

    ItemSeCae = (ObjData(Index).Real = 0 And _
                 ObjData(Index).Caos = 0 And _
                 ObjData(Index).ObjType <> OBJTYPE_LLAVES And _
                 ObjData(Index).ObjType <> OBJTYPE_BARCOS And _
                 Not ObjData(Index).NoSeCae)

End Function
Sub TirarTodosLosItems(UserIndex As Integer)

    If UserList(UserIndex).POS.Map = 8 Or _
       UserList(UserIndex).POS.Map = 17 Or _
       UserList(UserIndex).POS.Map = 7 Or _
       UserList(UserIndex).POS.Map = 1 Or _
       UserList(UserIndex).POS.Map = 212 Or _
       UserList(UserIndex).POS.Map = 88 Or _
       UserList(UserIndex).POS.Map = 15 Or _
       UserList(UserIndex).POS.Map = 16 Or _
       UserList(UserIndex).POS.Map = 12 Or _
       UserList(UserIndex).POS.Map = 10 Or _
       UserList(UserIndex).POS.Map = 11 Or _
       UserList(UserIndex).POS.Map = 14 Then Exit Sub

    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    Dim PosibilidadesZafa As Integer
    Dim ZafaMinerales As Boolean

    If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).trigger = 6 Then Exit Sub

    If UserList(UserIndex).Clase = PIRATA And UserList(UserIndex).Recompensas(2) = 1 And CInt(RandomNumber(1, 10)) <= 1 Then Exit Sub

    If UserList(UserIndex).Clase = MINERO Then
        If UserList(UserIndex).Recompensas(1) = 2 Then PosibilidadesZafa = 2
        If UserList(UserIndex).Recompensas(3) = 2 Then PosibilidadesZafa = PosibilidadesZafa + 3
        ZafaMinerales = CInt(RandomNumber(1, 10)) <= PosibilidadesZafa
    End If

    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).OBJIndex
        If ItemIndex Then
            If ItemSeCae(ItemIndex) And Not (ObjData(ItemIndex).ObjType = OBJTYPE_MINERALES And ZafaMinerales) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Call Tilelibre(UserList(UserIndex).POS, NuevaPos)
                If NuevaPos.X <> 0 And NuevaPos.Y Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.OBJIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
            End If
        End If

    Next

End Sub
Function ItemNewbie(ItemIndex As Integer) As Boolean

    ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function
Sub TirarTodosLosItemsNoNewbies(UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer

    If UserList(UserIndex).POS.Map = 8 Or _
       UserList(UserIndex).POS.Map = 17 Or _
       UserList(UserIndex).POS.Map = 7 Or _
       UserList(UserIndex).POS.Map = 1 Or _
       UserList(UserIndex).POS.Map = 15 Or _
       UserList(UserIndex).POS.Map = 212 Or _
       UserList(UserIndex).POS.Map = 16 Or _
       UserList(UserIndex).POS.Map = 12 Or _
       UserList(UserIndex).POS.Map = 10 Or _
       UserList(UserIndex).POS.Map = 11 Or _
       UserList(UserIndex).POS.Map = 14 Then Exit Sub

    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).OBJIndex
        If ItemIndex Then
            If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).POS, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.OBJIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
            End If

        End If
    Next

End Sub
