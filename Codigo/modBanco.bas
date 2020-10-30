Attribute VB_Name = "modBanco"
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
Sub IniciarDeposito(UserIndex As Integer)
    On Error GoTo errhandler

    Call UpdateBancoInv(True, UserIndex, 0)
    Call SendData(ToIndex, UserIndex, 0, "INITBANCO")
    UserList(UserIndex).flags.Comerciando = True

errhandler:

End Sub
Sub UpdateBancoInv(UpdateAll As Boolean, UserIndex As Integer, Slot As Byte, Optional ByVal TodaInfo As Boolean)
    Dim i As Byte

    If UpdateAll Then
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            Call EnviarBancoItem(UserIndex, i, UpdateAll)
        Next
    Else
        Call EnviarBancoItem(UserIndex, Slot, TodaInfo)
    End If

End Sub
Sub EnviarBancoItem(UserIndex As Integer, Slot As Byte, ByVal AllInfo As Boolean)
    Dim MiObj As UserOBJ

    MiObj = UserList(UserIndex).BancoInvent.Object(Slot)

    If MiObj.OBJIndex Then
        If AllInfo Then
            Call SendData(ToIndex, UserIndex, 0, "OTII" & Slot _
                                               & "," & ObjData(MiObj.OBJIndex).Name _
                                               & "," & MiObj.Amount _
                                               & "," & 0 _
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
            Call SendData(ToIndex, UserIndex, 0, "OTIC" & Slot & "," & MiObj.Amount)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "OTIV" & Slot)
    End If

End Sub
Sub UserRetiraItem(UserIndex As Integer, ByVal i As Byte, Cantidad As Integer)
    On Error GoTo errhandler

    If Cantidad < 1 Then Exit Sub

    If UserList(UserIndex).BancoInvent.Object(i).Amount Then
        If Cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then Cantidad = UserList(UserIndex).BancoInvent.Object(i).Amount
        Call UserReciveObj(UserIndex, CInt(i), Cantidad)
        Call UpdateBancoInv(False, UserIndex, i)
    End If

errhandler:

End Sub
Sub UserReciveObj(UserIndex As Integer, ByVal OBJIndex As Integer, Cantidad As Integer)
    Dim Slot As Byte
    Dim ObjI As Integer


    If UserList(UserIndex).BancoInvent.Object(OBJIndex).Amount <= 0 Then Exit Sub

    ObjI = UserList(UserIndex).BancoInvent.Object(OBJIndex).OBJIndex



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
                Call SendData(ToIndex, UserIndex, 0, "5W")
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
    End If




    If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then


        UserList(UserIndex).Invent.Object(Slot).OBJIndex = ObjI
        UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad


        Call UpdateUserInv(False, UserIndex, Slot)
        Call QuitarBancoInvItem(UserIndex, CByte(OBJIndex), Cantidad)

    Else
        Call SendData(ToIndex, UserIndex, 0, "5W")
    End If


End Sub

Sub QuitarBancoInvItem(UserIndex As Integer, Slot As Byte, Cantidad As Integer)
    Dim OBJIndex As Integer
    OBJIndex = UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex

    UserList(UserIndex).BancoInvent.Object(Slot).Amount = UserList(UserIndex).BancoInvent.Object(Slot).Amount - Cantidad

    If UserList(UserIndex).BancoInvent.Object(Slot).Amount <= 0 Then
        UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
        UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex = 0
        UserList(UserIndex).BancoInvent.Object(Slot).Amount = 0
    End If

End Sub
Sub UserDepositaItem(UserIndex As Integer, ByVal Item As Integer, Cantidad As Integer)
    On Error GoTo errhandler

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
        If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        Call UserDejaObj(UserIndex, CInt(Item), Cantidad)
    End If

errhandler:

End Sub
Sub UserDejaObj(UserIndex As Integer, ByVal OBJIndex As Integer, Cantidad As Integer)
    Dim Slot As Byte
    Dim ObjI As Integer

    If Cantidad < 1 Then Exit Sub

    ObjI = UserList(UserIndex).Invent.Object(OBJIndex).OBJIndex

    Slot = 1
    Do Until UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex = ObjI And _
       UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        Slot = Slot + 1

        If Slot > MAX_BANCOINVENTORY_SLOTS Then Exit Do
    Loop

    If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex = 0
            Slot = Slot + 1
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "9Y")
                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
    End If

    If UserList(UserIndex).Tienda.NroItems + UserList(UserIndex).BancoInvent.NroItems > MAX_BANCOINVENTORY_SLOTS Then
        Call SendData(ToIndex, UserIndex, 0, "/L")
        Exit Sub
    End If

    If Slot <= MAX_BANCOINVENTORY_SLOTS Then
        If UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex = ObjI
            UserList(UserIndex).BancoInvent.Object(Slot).Amount = UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad
            Call QuitarUserInvItem(UserIndex, CByte(OBJIndex), Cantidad)
            Call UpdateBancoInv(False, UserIndex, Slot, True)
        Else
            Call SendData(ToIndex, UserIndex, 0, "9Y")
        End If
        Call UpdateUserInv(False, UserIndex, CByte(OBJIndex))
    Else
        Call QuitarUserInvItem(UserIndex, CByte(OBJIndex), Cantidad)
    End If

End Sub


