Attribute VB_Name = "modParty"
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
Sub RepartirExp(UserIndex As Integer, Experiencia As Long, Died As Boolean)
    On Error GoTo Error
    Dim i As Integer
    Dim Miembro As Integer, Cant As Long, Exp As Long

    For i = 1 To Party(UserList(UserIndex).PartyIndex).NroMiembros
        Miembro = Party(UserList(UserIndex).PartyIndex).MiembrosIndex(i)
        If Miembro > 0 Then
            With UserList(Miembro)
                If .flags.Muerto = 0 And .POS.Map = UserList(UserIndex).POS.Map Then
                    Cant = Cant + 1 + Buleano(Not ClaseTrabajadora(.Clase))
                End If
            End With
        End If
    Next

    For i = 1 To Party(UserList(UserIndex).PartyIndex).NroMiembros
        Miembro = Party(UserList(UserIndex).PartyIndex).MiembrosIndex(i)
        If Miembro > 0 Then
            With UserList(Miembro)
                If .flags.Muerto = 0 And .POS.Map = UserList(UserIndex).POS.Map Then
                    Exp = ((1 + Buleano(Not ClaseTrabajadora(.Clase))) * Experiencia) / Cant
                    .Stats.Exp = .Stats.Exp + Exp
                    If UserIndex = Miembro Then
                        If Died Then
                            Call SendData(ToIndex, UserIndex, 0, "EL" & Exp)
                        Else: Call SendData(ToIndex, UserIndex, 0, "EX" & Exp)
                        End If
                    End If
                    Call SendUserEXP(Miembro)
                    Call CheckUserLevel(Miembro)
                End If
            End With
        End If
    Next

    Exit Sub
Error:
    Call Logear("Errores", "Error en RepartirExp " & UserList(UserIndex).Name & Err.Description)
End Sub

Sub EnviarIntegrantesParty(UserIndex As Integer)
    Dim i As Integer
    Dim SD As String

    If UserList(UserIndex).flags.Party = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "PARTYN")
        Exit Sub
    End If

    For i = 1 To MAXPARTYUSERS
        If Party(UserList(UserIndex).PartyIndex).MiembrosIndex(i) > 0 Then
            SD = SD & UserList(Party(UserList(UserIndex).PartyIndex).MiembrosIndex(i)).Name & " (Lvl " & UserList(Party(UserList(UserIndex).PartyIndex).MiembrosIndex(i)).Stats.ELV & "),"
        End If
    Next

    If Party(UserList(UserIndex).PartyIndex).MiembrosIndex(1) = UserIndex Then
        SD = "PARTYL" & SD
    Else
        SD = "PARTYI" & SD
    End If

    Call SendData(ToIndex, UserIndex, 0, SD)

End Sub
Public Function NoPuedeEntrarParty(UserIndex As Integer, Invitado As Integer) As Byte

    If EsNewbie(UserIndex) Then
        NoPuedeEntrarParty = 1
        Exit Function
    End If

    If EsNewbie(Invitado) Then
        NoPuedeEntrarParty = 15
        Exit Function
    End If

    If UserList(UserIndex).Stats.UserAtributos(Carisma) * UserList(UserIndex).Stats.UserSkills(Liderazgo) < 100 Then
        NoPuedeEntrarParty = 2
        Exit Function
    End If

    If UserIndex = Invitado Then
        NoPuedeEntrarParty = 3
        Exit Function
    End If

    If UserList(Invitado).Faccion.Bando = Enemigo(UserList(UserIndex).Faccion.Bando) Then
        NoPuedeEntrarParty = 4
        Exit Function
    End If

    If UserList(Invitado).flags.Muerto Then
        NoPuedeEntrarParty = 5
        Exit Function
    End If

    If UserList(Invitado).flags.Party Then
        If UserList(Invitado).PartyIndex = UserList(UserIndex).PartyIndex Then
            NoPuedeEntrarParty = 6
        Else: NoPuedeEntrarParty = 13
        End If
        Exit Function
    End If

    If UserList(Invitado).flags.Ofreciente > 0 And UserList(Invitado).flags.Ofreciente <> UserIndex Then
        NoPuedeEntrarParty = 14
        Exit Function
    End If

    If Distancia(UserList(Invitado).POS, UserList(UserIndex).POS) > 6 Then
        NoPuedeEntrarParty = 8
        Exit Function
    End If

    If UserList(UserIndex).flags.Party Then
        If UserList(Invitado).Faccion.Bando = Enemigo(Party(UserList(UserIndex).PartyIndex).Bando) Then
            NoPuedeEntrarParty = 12
        ElseIf MinLevel(UserList(UserIndex).PartyIndex) + 5 < UserList(Invitado).Stats.ELV Or MaxLevel(UserList(UserIndex).PartyIndex) - 5 > UserList(Invitado).Stats.ELV Then
            NoPuedeEntrarParty = 10
        ElseIf Party(UserList(UserIndex).PartyIndex).NroMiembros = MAXPARTYUSERS Then
            NoPuedeEntrarParty = 9
        ElseIf Party(UserList(UserIndex).PartyIndex).MiembrosIndex(1) <> UserIndex Then
            NoPuedeEntrarParty = 11
        End If
    ElseIf Abs(UserList(UserIndex).Stats.ELV - UserList(Invitado).Stats.ELV) > 5 Then
        NoPuedeEntrarParty = 10
    End If

End Function
Public Sub DoInvitar(UserIndex As Integer, Invitado As Integer)

    If UserIndex = 0 Or Invitado = 0 Then Exit Sub

    Select Case NoPuedeEntrarParty(UserIndex, Invitado)
    Case 0
        UserList(Invitado).flags.Ofreciente = UserIndex
        Call SendData(ToIndex, UserIndex, 0, "||Le ofreciste a " & UserList(Invitado).Name & " entrar en el party." & FONTTYPE_PARTY)
        Call SendData(ToIndex, Invitado, 0, "INVPAR" & UserList(UserIndex).Name)
    Case 1
        Call SendData(ToIndex, UserIndex, 0, "||Los newbies no pueden formar un party." & FONTTYPE_PARTY)
    Case 2
        Call SendData(ToIndex, UserIndex, 0, "||Por tu carisma necesitas " & Redondea(100 / UserList(UserIndex).Stats.UserAtributos(Carisma)) & " puntos de liderazgo para hacer un party." & FONTTYPE_PARTY)
    Case 3
        Call SendData(ToIndex, UserIndex, 0, "||No podés invitarte a party a ti mismo." & FONTTYPE_PARTY)
    Case 4
        Call SendData(ToIndex, UserIndex, 0, "||¡No podés entrar en party con miembros de " & ListaBandos(UserList(Invitado).Faccion.Bando) & "!" & FONTTYPE_PARTY)
    Case 5
        Call SendData(ToIndex, UserIndex, 0, "||No podés invitar al party a un muerto." & FONTTYPE_PARTY)
    Case 6
        Call SendData(ToIndex, UserIndex, 0, "||El usuario ya está en tu party." & FONTTYPE_PARTY)
    Case 7
        Call SendData(ToIndex, UserIndex, 0, "||Ya le ofreciste party al usuario." & FONTTYPE_PARTY)
    Case 8
        Call SendData(ToIndex, UserIndex, 0, "DL")
    Case 9
        Call SendData(ToIndex, UserIndex, 0, "||Tu party ya está lleno." & FONTTYPE_PARTY)
    Case 10
        Call SendData(ToIndex, UserIndex, 0, "||La diferencia de niveles es demasiado grande." & FONTTYPE_PARTY)
    Case 11
        Call SendData(ToIndex, UserIndex, 0, "||Solo el lider del party (" & UserList(Party(UserList(UserIndex).PartyIndex).MiembrosIndex(1)).Name & ") puede invitar gente a unirse al party." & FONTTYPE_PARTY)
    Case 12
        Call SendData(ToIndex, UserIndex, 0, "||Tu party ya está asignado a " & ListaBandos(Party(UserList(UserIndex).PartyIndex).Bando) & "." & FONTTYPE_PARTY)
    Case 13
        Call SendData(ToIndex, UserIndex, 0, "||El usuario ya está en otro party." & FONTTYPE_PARTY)
    Case 14
        Call SendData(ToIndex, UserIndex, 0, "||Alguien ya le ofreció party al usuario." & FONTTYPE_PARTY)
    Case 15
        Call SendData(ToIndex, UserIndex, 0, "||No puedes invitar a un newbie a un party." & FONTTYPE_PARTY)
    End Select

End Sub
Function MaxLevel(PartyIndex As Integer) As Byte
    Dim i As Integer

    For i = 1 To Party(PartyIndex).NroMiembros
        If Party(PartyIndex).MiembrosIndex(i) > 0 Then
            If UserList(Party(PartyIndex).MiembrosIndex(i)).Stats.ELV > MaxLevel Then
                MaxLevel = UserList(Party(PartyIndex).MiembrosIndex(i)).Stats.ELV
            End If
        End If
    Next

End Function
Function MinLevel(PartyIndex As Integer) As Byte
    Dim i As Integer

    MinLevel = STAT_MAXELV

    For i = 1 To Party(PartyIndex).NroMiembros
        If Party(PartyIndex).MiembrosIndex(i) > 0 Then
            If UserList(Party(PartyIndex).MiembrosIndex(i)).Stats.ELV < MinLevel Then
                MinLevel = UserList(Party(PartyIndex).MiembrosIndex(i)).Stats.ELV
            End If
        End If
    Next

End Function
Sub CrearParty(UserIndex As Integer)
    Dim PartyIndex As Integer
    Dim Lider As Integer
    Dim i As Integer

    If UserIndex = 0 Then Exit Sub
    Lider = UserList(UserIndex).flags.Ofreciente
    If Lider = 0 Then Exit Sub

    For i = 1 To UBound(Party)
        If Party(i).NroMiembros = 0 Then
            PartyIndex = i
            Party(PartyIndex).NroMiembros = 2
            Party(PartyIndex).MiembrosIndex(1) = Lider
            Party(PartyIndex).MiembrosIndex(2) = UserIndex
            UserList(Lider).PartyIndex = PartyIndex
            UserList(UserIndex).PartyIndex = PartyIndex
            UserList(Lider).flags.Party = 1
            UserList(UserIndex).flags.Party = 1
            UserList(UserIndex).flags.Ofreciente = 0
            UserList(Lider).flags.Ofreciente = 0
            Call SendData(ToIndex, UserIndex, 0, "||Entraste en party con " & UserList(Lider).Name & "." & FONTTYPE_PARTY)
            Call SendData(ToIndex, Lider, 0, "||" & UserList(UserIndex).Name & " aceptó entrar en party con vos." & FONTTYPE_PARTY)
            Party(PartyIndex).Bando = UserList(Lider).Faccion.Bando
            If UserList(UserIndex).Faccion.Bando Then Party(PartyIndex).Bando = UserList(UserIndex).Faccion.Bando
            Exit Sub
        End If
    Next

End Sub
Sub EntrarAlParty(UserIndex As Integer, PartyIndex As Integer)
    Dim i As Integer

    If UserIndex = 0 Or PartyIndex = 0 Then Exit Sub

    For i = 1 To MAXPARTYUSERS
        If Party(PartyIndex).MiembrosIndex(i) = 0 Then
            Party(PartyIndex).MiembrosIndex(i) = UserIndex
            Party(PartyIndex).NroMiembros = Party(PartyIndex).NroMiembros + 1
            UserList(UserIndex).flags.Party = 1
            UserList(UserIndex).PartyIndex = PartyIndex
            UserList(UserIndex).flags.Ofreciente = 0
            Call SendData(ToIndex, UserIndex, 0, "||Entraste al party." & FONTTYPE_PARTY)
            Call SendData(ToParty, UserIndex, 0, "||" & UserList(UserIndex).Name & " entró al party." & FONTTYPE_PARTY)
            If Party(PartyIndex).Bando = Neutral Then Party(PartyIndex).Bando = UserList(UserIndex).Faccion.Bando
            Exit Sub
        End If
    Next

End Sub
Sub EcharDelParty(UserIndex As Integer)
    Dim PartyIndex As Integer
    Dim i As Integer
    Dim NroMiembro As Integer

    If UserIndex = 0 Then Exit Sub

    PartyIndex = UserList(UserIndex).PartyIndex
    Party(PartyIndex).NroMiembros = Party(PartyIndex).NroMiembros - 1

    For i = 1 To MAXPARTYUSERS
        If Party(PartyIndex).MiembrosIndex(i) = UserIndex Then
            NroMiembro = i
            Exit For
        End If
    Next

    For i = NroMiembro + 1 To MAXPARTYUSERS
        Party(PartyIndex).MiembrosIndex(i - 1) = Party(PartyIndex).MiembrosIndex(i)
    Next

    Party(PartyIndex).MiembrosIndex(MAXPARTYUSERS) = 0

    Call SendData(ToParty, UserIndex, 0, "||" & UserList(UserIndex).Name & " fue echado del party." & FONTTYPE_PARTY)
    Call SendData(ToIndex, UserIndex, 0, "||Fuiste echado del party." & FONTTYPE_PARTY)
    Call SendData(ToIndex, UserIndex, 0, "PARTYN")

    UserList(UserIndex).flags.Party = 0
    UserList(UserIndex).PartyIndex = 0

End Sub
Sub SacarDelParty(UserIndex As Integer)
    Dim PartyIndex As Integer
    Dim i As Integer
    Dim NroMiembro As Integer

    If UserIndex = 0 Then Exit Sub

    PartyIndex = UserList(UserIndex).PartyIndex
    Party(PartyIndex).NroMiembros = Party(PartyIndex).NroMiembros - 1

    For i = 1 To MAXPARTYUSERS
        If Party(PartyIndex).MiembrosIndex(i) = UserIndex Then
            NroMiembro = i
            Exit For
        End If
    Next

    If NroMiembro = 1 Then
        Call RomperParty(UserIndex)
        Exit Sub
    End If

    For i = NroMiembro + 1 To MAXPARTYUSERS
        Party(PartyIndex).MiembrosIndex(i - 1) = Party(PartyIndex).MiembrosIndex(i)
    Next

    Party(PartyIndex).MiembrosIndex(MAXPARTYUSERS) = 0

    Call SendData(ToParty, UserIndex, 0, "||" & UserList(UserIndex).Name & " salió del party." & FONTTYPE_PARTY)

    Call SendData(ToIndex, UserIndex, 0, "||Saliste del party." & FONTTYPE_PARTY)


    Call SendData(ToIndex, UserIndex, 0, "PARTYN")
    UserList(UserIndex).flags.Party = 0
    UserList(UserIndex).PartyIndex = 0

    Exit Sub

Error:
    Call Logear("Errores", "Error en SacarDelParty-" & Err.Description)
End Sub
Sub RomperParty(UserIndex As Integer)
    On Error GoTo Error
    Dim i As Integer
    Dim PartyIndex As Integer

    If UserIndex = 0 Then Exit Sub

    PartyIndex = UserList(UserIndex).PartyIndex

    If PartyIndex = 0 Then Exit Sub

    Call SendData(ToParty, UserIndex, 0, "||El party fue disuelto." & FONTTYPE_PARTY)
    Call SendData(ToParty, UserIndex, 0, "PARTYN")

    For i = 1 To MAXPARTYUSERS
        If Party(PartyIndex).MiembrosIndex(i) > 0 Then
            UserList(Party(PartyIndex).MiembrosIndex(i)).flags.Party = 0
            UserList(Party(PartyIndex).MiembrosIndex(i)).PartyIndex = 0
            Party(PartyIndex).MiembrosIndex(i) = 0
            Party(PartyIndex).Bando = Neutral
        End If
    Next

    Party(PartyIndex).NroMiembros = 0
    Exit Sub
Error:
    Call Logear("Errores", "Error en RomperParty-" & Err.Description)
End Sub
