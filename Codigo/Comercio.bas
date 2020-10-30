Attribute VB_Name = "Comercio"
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
Sub UserCompraObj(UserIndex As Integer, ByVal OBJIndex As Integer, NpcIndex As Integer, Cantidad As Integer)
    Dim Infla As Integer
    Dim Desc As Single
    Dim unidad As Long, monto As Long
    Dim Slot As Byte
    Dim ObjI As Integer
    Dim Encontre As Boolean

    ObjI = Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(OBJIndex).OBJIndex

    Slot = 1
    Do Until UserList(UserIndex).Invent.Object(Slot).OBJIndex = ObjI And _
       UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS

        Slot = Slot + 1
        If Slot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop

    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).OBJIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "5P")
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
    End If

    If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        UserList(UserIndex).Invent.Object(Slot).OBJIndex = ObjI
        UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad
        Infla = (Npclist(NpcIndex).Inflacion * ObjData(ObjI).Valor) \ 100

        Desc = Descuento(UserIndex)

        unidad = Int(((ObjData(Npclist(NpcIndex).Invent.Object(OBJIndex).OBJIndex).Valor + Infla) / Desc))
        If unidad = 0 Then unidad = 1
        monto = unidad * Cantidad
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - (monto * 0)

        Call SubirSkill(UserIndex, Comerciar)

        If ObjData(ObjI).ObjType = OBJTYPE_LLAVES Then Call LogVentaCasa(UserList(UserIndex).Name & " compro " & ObjData(ObjI).Name)
        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNpc, CByte(OBJIndex), Cantidad, UserIndex)

        Call UpdateUserInv(False, UserIndex, Slot)
    Else
        Call SendData(ToIndex, UserIndex, 0, "5P")
    End If

End Sub
Sub UpdateNPCInv(UpdateAll As Boolean, UserIndex As Integer, NpcIndex As Integer, Slot As Byte)
    Dim i As Byte
    Dim MiObj As UserOBJ

    If UpdateAll Then
        For i = 1 To MAX_NPCINVENTORY_SLOTS
            Call SendNPCItem(UserIndex, NpcIndex, i, UpdateAll)
        Next
    Else
        Call SendNPCItem(UserIndex, NpcIndex, i, UpdateAll)
    End If

End Sub
Sub SendNPCItem(UserIndex As Integer, NpcIndex As Integer, Slot As Byte, ByVal AllInfo As Boolean)
    Dim MiObj As UserOBJ
    Dim Infla As Long
    Dim Desc As Single
    Dim val As Long

    MiObj = Npclist(NpcIndex).Invent.Object(Slot)

    Desc = Descuento(UserIndex)

    If Desc >= 0 And Desc <= 1 Then Desc = 1




    If MiObj.OBJIndex Then
        If AllInfo Then
            Infla = (Npclist(NpcIndex).Inflacion * ObjData(MiObj.OBJIndex).Valor) / 100
            val = Maximo(1, Int((ObjData(MiObj.OBJIndex).Valor + Infla) / Desc))
            Call SendData(ToIndex, UserIndex, 0, "OTII" & Slot _
                                               & "," & ObjData(MiObj.OBJIndex).Name _
                                               & "," & MiObj.Amount _
                                               & "," & val _
                                               & "," & ObjData(MiObj.OBJIndex).GrhIndex _
                                               & "," & MiObj.OBJIndex _
                                               & "," & ObjData(MiObj.OBJIndex).ObjType _
                                               & "," & ObjData(MiObj.OBJIndex).MaxHit _
                                               & "," & ObjData(MiObj.OBJIndex).MinHit _
                                               & "," & ObjData(MiObj.OBJIndex).MaxDef _
                                               & "," & ObjData(MiObj.OBJIndex).MinDef _
                                               & "," & ObjData(MiObj.OBJIndex).TipoPocion _
                                               & "," & ObjData(MiObj.OBJIndex).MaxModificador _
                                               & "," & ObjData(MiObj.OBJIndex).MinModificador _
                                               & "," & PuedeUsarObjeto(UserIndex, MiObj.OBJIndex))
        Else
            Call SendData(ToIndex, UserIndex, 0, "OTIC" & Slot & "," & MiObj.Amount)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "OTIV" & Slot)
    End If

End Sub
Sub IniciarComercioNPC(UserIndex As Integer)
    On Error GoTo errhandler

    Call UpdateNPCInv(True, UserIndex, UserList(UserIndex).flags.TargetNpc, 0)
    Call SendData(ToIndex, UserIndex, 0, "INITCOM")
    UserList(UserIndex).flags.Comerciando = True

errhandler:

End Sub
Sub NPCVentaItem(UserIndex As Integer, ByVal i As Integer, Cantidad As Integer, NpcIndex As Integer)
    On Error GoTo errhandler
    Dim Infla As Long
    Dim val As Long
    Dim Desc As Single

    If Cantidad < 1 Then Exit Sub


    Infla = (Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).OBJIndex).Valor) / 100
    Desc = Descuento(UserIndex)

    val = Fix((ObjData(Npclist(NpcIndex).Invent.Object(i).OBJIndex).Valor + Infla) / Desc)
    If val = 0 Then val = 1

    If UserList(UserIndex).Stats.GLD >= (val * Cantidad) Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount > 0 Or Npclist(UserList(UserIndex).flags.TargetNpc).InvReSpawn = 0 Then
            If Cantidad > Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount And Npclist(UserList(UserIndex).flags.TargetNpc).InvReSpawn = 1 Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount
            Call UserCompraObj(UserIndex, CInt(i), UserList(UserIndex).flags.TargetNpc, Cantidad)
            Call SendUserORO(UserIndex)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "2Q")
        Exit Sub
    End If

errhandler:

End Sub
Sub NPCCompraItem(UserIndex As Integer, ByVal Item As Byte, Cantidad As Integer)
    On Error GoTo errhandler

    If ObjData(UserList(UserIndex).Invent.Object(Item).OBJIndex).Newbie = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "6P")
        Exit Sub
    End If

    If ObjData(UserList(UserIndex).Invent.Object(Item).OBJIndex).NoSeCae = 1 Or ObjData(UserList(UserIndex).Invent.Object(Item).OBJIndex).Real > 0 Or ObjData(UserList(UserIndex).Invent.Object(Item).OBJIndex).Caos > 0 Or ObjData(UserList(UserIndex).Invent.Object(Item).OBJIndex).Newbie = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes vender este item." & FONTTYPE_WARNING)
        Exit Sub
    End If

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
        If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        UserList(UserIndex).Invent.Object(Item).Amount = UserList(UserIndex).Invent.Object(Item).Amount - Cantidad
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + (ObjData(UserList(UserIndex).Invent.Object(Item).OBJIndex).Valor / 3 * Cantidad)
        If UserList(UserIndex).Invent.Object(Item).Amount <= 0 Then
            UserList(UserIndex).Invent.Object(Item).Amount = 0
            UserList(UserIndex).Invent.Object(Item).OBJIndex = 0
            UserList(UserIndex).Invent.Object(Item).Equipped = 0
        End If
        Call SubirSkill(UserIndex, Comerciar)
        Call UpdateUserInv(False, UserIndex, Item)
    End If

    Call SendUserORO(UserIndex)
    Exit Sub
errhandler:

End Sub
Public Function Descuento(UserIndex As Integer) As Single

    Descuento = CSng(Minimo(10 + (Fix((UserList(UserIndex).Stats.UserSkills(Comerciar) + UserList(UserIndex).Stats.UserAtributos(Carisma) - 10) / 10)), 20)) / 10

End Function
