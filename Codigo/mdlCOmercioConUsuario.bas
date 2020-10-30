Attribute VB_Name = "mdlCOmercioConUsuario"
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

Public Type tCOmercioUsuario
    DestUsu As Integer
    DestNick As String
    Objeto As Integer
    Cant As Long

    Acepto As Boolean
End Type
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
    On Error GoTo errhandler

    If UserList(Origen).ComUsu.DestUsu = Destino And _
       UserList(Destino).ComUsu.DestUsu = Origen Then

        Call UpdateUserInv(True, Origen, 0)

        Call SendData(ToIndex, Origen, 0, "INITCOMUSU")
        UserList(Origen).flags.Comerciando = True


        Call UpdateUserInv(True, Destino, 0)

        Call SendData(ToIndex, Destino, 0, "INITCOMUSU")
        UserList(Destino).flags.Comerciando = True
    Else

        Call SendData(ToIndex, Destino, 0, "||" & UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR." & FONTTYPE_TALK)
        UserList(Destino).flags.TargetUser = Origen

    End If

    Exit Sub
errhandler:

End Sub
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
    Dim ObjInd As Integer
    Dim ObjCant As Long

    ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
    If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
        ObjInd = iORO
    Else
        ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).OBJIndex
    End If

    If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

    If ObjInd > 0 And ObjCant Then
        Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & ObjInd & "," & ObjData(ObjInd).Name & "," & ObjCant & "," & 0 & "," & ObjData(ObjInd).GrhIndex & "," _
                                        & ObjData(ObjInd).ObjType & "," _
                                        & ObjData(ObjInd).MaxHit & "," _
                                        & ObjData(ObjInd).MinHit & "," _
                                        & ObjData(ObjInd).MaxDef & "," _
                                        & ObjData(ObjInd).Valor \ 3)
    End If

End Sub
Public Sub FinComerciarUsu(UserIndex As Integer)

    If UserIndex = 0 Then Exit Sub

    With UserList(UserIndex)
        If .ComUsu.DestUsu Then Call SendData(ToIndex, UserIndex, 0, "FINCOMUSUOK")
        .ComUsu.Acepto = False
        .ComUsu.Cant = 0
        .ComUsu.DestUsu = 0
        .ComUsu.Objeto = 0
        .ComUsu.DestNick = ""
        .flags.Comerciando = False
    End With

End Sub
Public Sub AceptarComercioUsu(UserIndex As Integer)
    Dim Obj1 As Obj, Obj2 As Obj
    Dim OtroUserIndex As Integer
    Dim TerminarAhora As Boolean

    OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu

    If OtroUserIndex <= 0 Then
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    End If

    TerminarAhora = (UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex) Or _
                    (UserList(OtroUserIndex).Name <> UserList(UserIndex).ComUsu.DestNick) Or _
                    (UserList(UserIndex).Name <> UserList(OtroUserIndex).ComUsu.DestNick) Or _
                    (Not UserList(OtroUserIndex).flags.UserLogged Or Not UserList(UserIndex).flags.UserLogged)

    If TerminarAhora Then
        Call FinComerciarUsu(UserIndex)
        Call FinComerciarUsu(OtroUserIndex)
        Exit Sub
    End If

    UserList(UserIndex).ComUsu.Acepto = True

    If Not UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto Then
        Call SendData(ToIndex, UserIndex, 0, "||El otro usuario aun no ha aceptado tu oferta." & FONTTYPE_TALK)
        Exit Sub
    End If

    If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
        Obj1.OBJIndex = iORO
        If UserList(UserIndex).ComUsu.Cant > UserList(UserIndex).Stats.GLD Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
            TerminarAhora = True
        End If
    Else
        Obj1.Amount = UserList(UserIndex).ComUsu.Cant
        Obj1.OBJIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).OBJIndex
        If Obj1.Amount > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Amount Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
            TerminarAhora = True
        End If
    End If
    If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
        Obj2.OBJIndex = iORO
        If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
            Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
            TerminarAhora = True
        End If
    Else
        Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
        Obj2.OBJIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).OBJIndex
        If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
            Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
            TerminarAhora = True
        End If
    End If

    If TerminarAhora Then
        Call FinComerciarUsu(UserIndex)
        Call FinComerciarUsu(OtroUserIndex)
        Exit Sub
    End If


    If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then

        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
        Call SendUserORO(OtroUserIndex)

        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
        Call SendUserORO(UserIndex)
    Else

        If Not MeterItemEnInventario(UserIndex, Obj2) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Obj2)
        Call QuitarObjetos(Obj2.OBJIndex, Obj2.Amount, OtroUserIndex)
    End If


    If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then

        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Cant
        Call SendUserORO(UserIndex)

        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Cant
        Call SendUserORO(OtroUserIndex)
    Else

        If Not MeterItemEnInventario(OtroUserIndex, Obj1) Then Call TirarItemAlPiso(UserList(OtroUserIndex).POS, Obj1)
        Call QuitarObjetos(Obj1.OBJIndex, Obj1.Amount, UserIndex)
    End If



    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserInv(True, OtroUserIndex, 0)

    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)

End Sub



