Attribute VB_Name = "modTienda"
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


Public DineroTotalVentas As Double
Public NumeroVentas As Long

Option Explicit
Sub TiendaVentaItem(UserIndex As Integer, ByVal i As Integer, Cantidad As Integer, NpcIndex As Integer)
    On Error GoTo errhandler
    Dim Vendedor As Integer

    If Cantidad < 1 Or Npclist(NpcIndex).NPCtype <> NPCTYPE_TIENDA Then Exit Sub

    Vendedor = Npclist(NpcIndex).flags.TiendaUser

    If UserList(UserIndex).Stats.GLD >= (UserList(Vendedor).Tienda.Object(i).Precio * Cantidad) Then
        If UserList(Vendedor).Tienda.Object(i).Amount Then
            If Cantidad > UserList(Vendedor).Tienda.Object(i).Amount Then Cantidad = UserList(Vendedor).Tienda.Object(i).Amount
            Call TiendaCompraItem(UserIndex, CInt(i), UserList(UserIndex).flags.TargetNpc, Cantidad)
            Call SendUserORO(UserIndex)
        Else
            Call SendData(ToIndex, UserIndex, 0, "OTIV" & i)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "2Q")
        Exit Sub
    End If

errhandler:

End Sub
Sub TiendaCompraItem(UserIndex As Integer, Slot As Byte, NpcIndex As Integer, Cantidad As Integer)
    Dim Vendedor As Integer
    Dim ObjI As Integer
    Dim Encontre As Boolean
    Dim MiObj As Obj

    Vendedor = Npclist(NpcIndex).flags.TiendaUser

    If (UserList(Vendedor).Tienda.Object(Slot).Amount <= 0) Then Exit Sub

    ObjI = UserList(Vendedor).Tienda.Object(Slot).OBJIndex

    MiObj.OBJIndex = ObjI
    MiObj.Amount = Cantidad

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call SendData(ToIndex, UserIndex, 0, "5P")
        Exit Sub
    End If

    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad

    Call VendedorVenta(Vendedor, CByte(Slot), Cantidad, UserIndex)

End Sub
Sub VendedorVenta(Vendedor As Integer, Slot As Byte, Cantidad As Integer, Comprador As Integer)

    Call SendData(ToIndex, Vendedor, 0, "/R" & UserList(Comprador).Name & "," & ObjData(UserList(Vendedor).Tienda.Object(Slot).OBJIndex).Name & "," & Cantidad & "," & UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad)
    UserList(Vendedor).Stats.Banco = UserList(Vendedor).Stats.Banco + UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad
    UserList(Vendedor).Tienda.Gold = UserList(Vendedor).Tienda.Gold + UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad
    DineroTotalVentas = DineroTotalVentas + UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad
    NumeroVentas = NumeroVentas + 1

    UserList(Vendedor).Tienda.Object(Slot).Amount = UserList(Vendedor).Tienda.Object(Slot).Amount - Cantidad

    If UserList(Vendedor).Tienda.Object(Slot).Amount <= 0 Then
        UserList(Vendedor).Tienda.Object(Slot).Amount = 0
        UserList(Vendedor).Tienda.Object(Slot).OBJIndex = 0
        UserList(Vendedor).Tienda.Object(Slot).Precio = 0
        UserList(Vendedor).Tienda.NroItems = UserList(Vendedor).Tienda.NroItems - 1
        If UserList(Vendedor).Tienda.NroItems <= 0 Then
            Npclist(UserList(Vendedor).Tienda.NpcTienda).flags.TiendaUser = 0
            UserList(Vendedor).Tienda.NpcTienda = 0
            Call SendData(ToIndex, Vendedor, 0, "/S")
            Call SendData(ToIndex, Comprador, 0, "FINCOMOK")
            Exit Sub
        End If
    End If

    Call UpdateTiendaC(False, Comprador, UserList(Vendedor).Tienda.NpcTienda, Slot)

    Exit Sub
errhandler:

End Sub
Sub IniciarComercioTienda(UserIndex As Integer, NpcIndex As Integer)

    Call UpdateTiendaC(True, UserIndex, NpcIndex, 0)
    Call SendData(ToIndex, UserIndex, 0, "INITCOM")
    UserList(UserIndex).flags.Comerciando = True

End Sub
Public Sub IniciarAlquiler(UserIndex As Integer)

    If Not (ClaseTrabajadora(UserList(UserIndex).Clase) And Not EsNewbie(UserIndex)) And Not (UserList(UserIndex).Stats.ELV >= 25 And UserList(UserIndex).Stats.UserSkills(Comerciar) >= 65) Then
        Call SendData(ToIndex, UserIndex, 0, "/V" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    If UserList(UserIndex).Tienda.NpcTienda > 0 And UserList(UserIndex).Tienda.NpcTienda <> UserList(UserIndex).flags.TargetNpc Then
        Call SendData(ToIndex, UserIndex, 0, "/W" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    Call UpdateTiendaV(True, UserIndex, 0)
    Call SendData(ToIndex, UserIndex, 0, "INITIENDA")
    UserList(UserIndex).flags.Comerciando = True

End Sub
Sub UpdateTiendaV(ByVal UpdateAll As Boolean, UserIndex As Integer, Slot As Byte, Optional ByVal TodaInfo As Boolean)
    Dim i As Byte

    If UpdateAll Then
        For i = 1 To MAX_TIENDA_SLOTS
            Call SendTiendaItemV(UserIndex, i, UpdateAll)
        Next
    Else
        Call SendTiendaItemV(UserIndex, Slot, TodaInfo)
    End If

End Sub
Sub SendTiendaItemV(UserIndex As Integer, Slot As Byte, TodaInfo As Boolean)
    Dim MiObj As TiendaObj

    MiObj = UserList(UserIndex).Tienda.Object(Slot)

    If MiObj.OBJIndex Then
        If TodaInfo Then
            Call SendData(ToIndex, UserIndex, 0, "OTII" & Slot _
                                               & "," & ObjData(MiObj.OBJIndex).Name _
                                               & "," & MiObj.Amount _
                                               & "," & MiObj.Precio _
                                               & "," & ObjData(MiObj.OBJIndex).GrhIndex _
                                               & "," & MiObj.OBJIndex _
                                               & "," & ObjData(MiObj.OBJIndex).ObjType _
                                               & "," & ObjData(MiObj.OBJIndex).MaxHit _
                                               & "," & ObjData(MiObj.OBJIndex).MinHit _
                                               & "," & ObjData(MiObj.OBJIndex).MaxDef _
                                               & "," & ObjData(MiObj.OBJIndex).MinDef _
                                               & "," & ObjData(MiObj.OBJIndex).TipoPocion _
                                               & "," & ObjData(MiObj.OBJIndex).MaxModificador _
                                               & "," & ObjData(MiObj.OBJIndex).MinModificador)
        Else
            Call SendData(ToIndex, UserIndex, 0, "OTIC " & Slot & "," & MiObj.Amount)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "OTIV" & Slot)
    End If

End Sub
Sub UpdateTiendaC(ByVal UpdateAll As Boolean, UserIndex As Integer, NpcIndex As Integer, Slot As Byte)
    Dim i As Byte

    If UpdateAll Then
        For i = 1 To MAX_TIENDA_SLOTS
            Call SendTiendaItemC(UserIndex, NpcIndex, i, UpdateAll)
        Next
    Else
        Call SendTiendaItemC(UserIndex, NpcIndex, Slot, UpdateAll)
    End If

End Sub
Sub SendTiendaItemC(UserIndex As Integer, NpcIndex As Integer, Slot As Byte, ByVal TodaInfo As Boolean)
    Dim MiObj As TiendaObj

    MiObj = UserList(Npclist(NpcIndex).flags.TiendaUser).Tienda.Object(Slot)

    If MiObj.OBJIndex Then
        If TodaInfo Then
            Call SendData(ToIndex, UserIndex, 0, "OTII" & Slot _
                                               & "," & ObjData(MiObj.OBJIndex).Name _
                                               & "," & MiObj.Amount _
                                               & "," & MiObj.Precio _
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
Sub UserSacaVenta(UserIndex As Integer, Slot As Byte, Cantidad As Integer)
    On Error GoTo errhandler

    If UserList(UserIndex).Tienda.Object(Slot).Amount Then
        If Cantidad > UserList(UserIndex).Tienda.Object(Slot).Amount Then Cantidad = UserList(UserIndex).Tienda.Object(Slot).Amount
        Call UserSacaObjVenta(UserIndex, CInt(Slot), Cantidad)
    End If

    Exit Sub
errhandler:

End Sub
Sub UserPoneVenta(UserIndex As Integer, Slot As Byte, Cantidad As Integer, Precio As Long)
    On Error GoTo errhandler

    If ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).Newbie Then
        Call SendData(ToIndex, UserIndex, 0, "/H")
        Exit Sub
    End If

    If ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).NoSeCae Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes poner este objeto a la venta." & FONTTYPE_INFO)
        Exit Sub
    End If

    If ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).Caos > 0 Or ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).Real Then
        Call SendData(ToIndex, UserIndex, 0, "/I")
        Exit Sub
    End If

    If Precio = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "/M")
        Exit Sub
    End If

    If UserList(UserIndex).Tienda.NpcTienda = 0 Then
        UserList(UserIndex).Tienda.NpcTienda = UserList(UserIndex).flags.TargetNpc
        Npclist(UserList(UserIndex).flags.TargetNpc).flags.TiendaUser = UserIndex
    End If

    If UserList(UserIndex).Invent.Object(Slot).Amount > 0 And UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
        If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).Amount
        Call UserDaObjVenta(UserIndex, CInt(Slot), Cantidad, Precio)
    End If

    Exit Sub
errhandler:

End Sub
Sub UserSacaObjVenta(UserIndex As Integer, ByVal Itemslot As Byte, Cantidad As Integer)
    Dim MiObj As Obj

    If Cantidad < 1 Then Exit Sub

    MiObj.OBJIndex = UserList(UserIndex).Tienda.Object(Itemslot).OBJIndex
    MiObj.Amount = Cantidad

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call SendData(ToIndex, UserIndex, 0, "/J")
        Exit Sub
    End If

    UserList(UserIndex).Tienda.Object(Itemslot).Amount = UserList(UserIndex).Tienda.Object(Itemslot).Amount - Cantidad

    If UserList(UserIndex).Tienda.Object(Itemslot).Amount <= 0 Then
        UserList(UserIndex).Tienda.Object(Itemslot).Amount = 0
        UserList(UserIndex).Tienda.Object(Itemslot).OBJIndex = 0
        UserList(UserIndex).Tienda.Object(Itemslot).Precio = 0
        UserList(UserIndex).Tienda.NroItems = UserList(UserIndex).Tienda.NroItems - 1
        If UserList(UserIndex).Tienda.NroItems <= 0 Then
            Npclist(UserList(UserIndex).Tienda.NpcTienda).flags.TiendaUser = 0
            UserList(UserIndex).Tienda.NpcTienda = 0
        End If
    End If

    Call UpdateTiendaV(False, UserIndex, Itemslot)

End Sub
Sub UserDaObjVenta(UserIndex As Integer, ByVal Itemslot As Byte, Cantidad As Integer, ByVal Precio As Long)
    Dim Slot As Byte
    Dim ObjI As Integer
    Dim SlotHayado As Boolean

    If Cantidad < 1 Then Exit Sub

    ObjI = UserList(UserIndex).Invent.Object(Itemslot).OBJIndex

    For Slot = 1 To MAX_TIENDA_SLOTS
        If UserList(UserIndex).Tienda.Object(Slot).OBJIndex = ObjI Then
            SlotHayado = True
            Exit For
        End If
    Next

    If Not SlotHayado Then
        For Slot = 1 To MAX_TIENDA_SLOTS
            If UserList(UserIndex).Tienda.Object(Slot).OBJIndex = 0 Then
                If UserList(UserIndex).Tienda.NroItems + UserList(UserIndex).BancoInvent.NroItems + 1 > MAX_BANCOINVENTORY_SLOTS Then
                    Call SendData(ToIndex, UserIndex, 0, "/K")
                    Exit Sub
                End If
                UserList(UserIndex).Tienda.NroItems = UserList(UserIndex).Tienda.NroItems + 1
                SlotHayado = True
                Exit For
            End If
        Next
    End If

    If Not SlotHayado Then
        Call SendData(ToIndex, UserIndex, 0, "/G")
        Exit Sub
    End If

    If UserList(UserIndex).Tienda.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        UserList(UserIndex).Tienda.Object(Slot).OBJIndex = ObjI
        UserList(UserIndex).Tienda.Object(Slot).Amount = UserList(UserIndex).Tienda.Object(Slot).Amount + Cantidad
        UserList(UserIndex).Tienda.Object(Slot).Precio = Precio
        Call QuitarUserInvItem(UserIndex, CByte(Itemslot), Cantidad)
    Else
        Call SendData(ToIndex, UserIndex, 0, "/G")
    End If

    Call UpdateUserInv(False, UserIndex, CByte(Itemslot))
    Call UpdateTiendaV(False, UserIndex, Slot, True)

End Sub
Sub DevolverItemsVenta(UserIndex As Integer)
    Dim i As Byte


    For i = 1 To MAX_TIENDA_SLOTS
        If UserList(UserIndex).Tienda.Object(i).OBJIndex Then Call TiendaABoveda(UserIndex, i)
    Next

End Sub
Sub TiendaABoveda(UserIndex As Integer, Itemslot As Byte)
    Dim Slot As Byte
    Dim ObjI As Integer
    Dim SlotHayado As Boolean

    ObjI = UserList(UserIndex).Tienda.Object(Itemslot).OBJIndex

    For Slot = 1 To MAX_BANCOINVENTORY_SLOTS
        If UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex = ObjI Then
            SlotHayado = True
            Exit For
        End If
    Next

    If Not SlotHayado Then
        For Slot = 1 To MAX_BANCOINVENTORY_SLOTS
            If UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex = 0 Then
                UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
                SlotHayado = True
                Exit For
            End If
        Next
    End If

    If Not SlotHayado Then Exit Sub

    If UserList(UserIndex).BancoInvent.Object(Slot).Amount + UserList(UserIndex).Tienda.Object(Itemslot).Amount <= MAX_INVENTORY_OBJS Then
        UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex = ObjI
        UserList(UserIndex).BancoInvent.Object(Slot).Amount = UserList(UserIndex).BancoInvent.Object(Slot).Amount + UserList(UserIndex).Tienda.Object(Itemslot).Amount
        UserList(UserIndex).Tienda.Object(Itemslot).Amount = 0
        UserList(UserIndex).Tienda.Object(Itemslot).OBJIndex = 0
        UserList(UserIndex).Tienda.Object(Itemslot).Precio = 0
        UserList(UserIndex).Tienda.NroItems = UserList(UserIndex).Tienda.NroItems - 1
    End If

End Sub
