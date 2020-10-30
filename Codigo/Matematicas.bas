Attribute VB_Name = "Matematicas"
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
Public Function Porcentaje(Total As Variant, Porc As Variant) As Long

    Porcentaje = Total * (Porc / 100)

End Function
Sub RestVar(Var As Variant, Take As Variant, MIN As Variant)

    Var = Maximo(Var - Take, MIN)

End Sub
Sub AddtoVar(Var As Variant, Addon As Variant, MAX As Variant)

    Var = Minimo(Var + Addon, MAX)

End Sub
Function Distancia(wp1 As WorldPos, wp2 As WorldPos)

    Distancia = (Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100))

End Function
Function TipoClase(UserIndex As Integer) As Byte

    Select Case UserList(UserIndex).Clase
    Case PALADIN, ASESINO, CAZADOR
        TipoClase = 2
    Case CLERIGO, BARDO, LADRON
        TipoClase = 3
    Case MAGO, NIGROMANTE, DRUIDA
        TipoClase = 4
    Case Else
        TipoClase = 1
    End Select

End Function
Public Function TipoRaza(UserIndex As Integer) As Byte

    If UserList(UserIndex).Raza = ENANO Or UserList(UserIndex).Raza = GNOMO Then
        TipoRaza = 2
    Else: TipoRaza = 1
    End If

End Function
Public Function RazaBaja(UserIndex As Integer) As Boolean

    RazaBaja = (UserList(UserIndex).Raza = ENANO Or UserList(UserIndex).Raza = GNOMO)

End Function
Function Distance(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As Double

    Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function
