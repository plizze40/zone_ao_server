Attribute VB_Name = "modSubclases"
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

Sub EnviarSubclase(UserIndex As Integer)

    If PuedeSubirClase(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "RECOM" & UserList(UserIndex).Clase)

End Sub
Sub RecibirRecompensa(UserIndex As Integer, Eleccion As Byte)
    Dim Recompensa As Byte
    Dim i As Integer

    Recompensa = PuedeRecompensa(UserIndex)

    If Recompensa = 0 Then Exit Sub

    UserList(UserIndex).Recompensas(Recompensa) = Eleccion

    If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeHP Then
        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeHP, STAT_MAXHP)
        Call SendUserMAXHP(UserIndex)
    End If

    If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeMP Then
        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeMP, 2000 + 200 * Buleano(UserList(UserIndex).Clase = MAGO) * 200 + 300 * Buleano(UserList(UserIndex).Clase = MAGO And UserList(UserIndex).Recompensas(2) = 2))
        Call SendUserMAXMANA(UserIndex)
    End If

    For i = 1 To 2
        If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i).OBJIndex Then
            If Not MeterItemEnInventario(UserIndex, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i)) Then Call TirarItemAlPiso(UserList(UserIndex).POS, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i))
        End If
    Next

    If PuedeRecompensa(UserIndex) = 0 Then Call SendData(ToIndex, UserIndex, 0, "SURE0")

End Sub
Sub RecibirSubclase(Clase As Byte, UserIndex As Integer)

    If Not PuedeSubirClase(UserIndex) Then Exit Sub

    Select Case UserList(UserIndex).Clase
    Case CIUDADANO
        If Clase = 1 Then
            UserList(UserIndex).Clase = TRABAJADOR
        Else: UserList(UserIndex).Clase = LUCHADOR
        End If

    Case TRABAJADOR
        Select Case Clase
        Case 1
            UserList(UserIndex).Clase = EXPERTO_MINERALES
        Case 2
            UserList(UserIndex).Clase = EXPERTO_MADERA
        Case 3
            UserList(UserIndex).Clase = PESCADOR
        Case 4
            UserList(UserIndex).Clase = SASTRE
        End Select

    Case EXPERTO_MINERALES
        If Clase = 1 Then
            UserList(UserIndex).Clase = MINERO
        Else: UserList(UserIndex).Clase = HERRERO
        End If

    Case EXPERTO_MADERA
        If Clase = 1 Then
            UserList(UserIndex).Clase = TALADOR
        Else: UserList(UserIndex).Clase = CARPINTERO
        End If

    Case LUCHADOR
        Call Aprenderhechizo(UserIndex, 2)
        Call Aprenderhechizo(UserIndex, 8)
        Call Aprenderhechizo(UserIndex, 5)
        Call Aprenderhechizo(UserIndex, 15)
        Call Aprenderhechizo(UserIndex, 23)
        Call Aprenderhechizo(UserIndex, 24)
        Call Aprenderhechizo(UserIndex, 25)
        Call Aprenderhechizo(UserIndex, 10)
        If Clase = 1 Then
            UserList(UserIndex).Clase = CON_MANA
            Call Aprenderhechizo(UserIndex, 2)
            UserList(UserIndex).Stats.MaxMAN = 100
            Call SendUserMAXMANA(UserIndex)
            If Not PuedeSubirClase(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUCL0")
            Exit Sub
        Else: UserList(UserIndex).Clase = SIN_MANA
        End If

    Case CON_MANA
        Select Case Clase
        Case 1
            UserList(UserIndex).Clase = HECHICERO
        Case 2
            UserList(UserIndex).Clase = ORDEN_SAGRADA
        Case 3
            UserList(UserIndex).Clase = NATURALISTA
        Case 4
            UserList(UserIndex).Clase = SIGILOSO
        End Select

    Case HECHICERO
        If Clase = 1 Then
            UserList(UserIndex).Clase = MAGO
        Else: UserList(UserIndex).Clase = NIGROMANTE
        End If

    Case ORDEN_SAGRADA
        If Clase = 1 Then
            UserList(UserIndex).Clase = PALADIN
        Else
            UserList(UserIndex).Clase = CLERIGO
        End If

    Case NATURALISTA
        If Clase = 1 Then
            UserList(UserIndex).Clase = BARDO
        Else: UserList(UserIndex).Clase = DRUIDA
        End If

    Case SIGILOSO
        If Clase = 1 Then
            UserList(UserIndex).Clase = ASESINO
        Else: UserList(UserIndex).Clase = CAZADOR
        End If

    Case SIN_MANA
        If Clase = 1 Then
            UserList(UserIndex).Clase = BANDIDO
        Else: UserList(UserIndex).Clase = CABALLERO
        End If

    Case BANDIDO
        If Clase = 1 Then
            UserList(UserIndex).Clase = PIRATA
        Else: UserList(UserIndex).Clase = LADRON
        End If

    Case CABALLERO
        If Clase = 1 Then
            UserList(UserIndex).Clase = GUERRERO
        Else: UserList(UserIndex).Clase = ARQUERO
        End If
    End Select

    Call CalcularValores(UserIndex)
    Call SendData(ToIndex, UserIndex, 0, "/0" & ListaClases(UserList(UserIndex).Clase))

    If Not PuedeSubirClase(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUCL0")

End Sub
