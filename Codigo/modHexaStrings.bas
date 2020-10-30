Attribute VB_Name = "modHexaStrings"
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
Public Function hexMd52Asc(ByVal md5 As String) As String
    Dim i As Integer, L As String

    md5 = UCase$(md5)
    If Len(md5) Mod 2 = 1 Then md5 = "0" & md5

    For i = 1 To Len(md5) \ 2
        L = Mid$(md5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(L))
    Next

End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    Dim i As Integer, L As String
    For i = 1 To Len(hex)
        L = Mid$(hex, i, 1)
        Select Case L
        Case "A": L = 10
        Case "B": L = 11
        Case "C": L = 12
        Case "D": L = 13
        Case "E": L = 14
        Case "F": L = 15
        End Select

        hexHex2Dec = (L * 16 ^ ((Len(hex) - i))) + hexHex2Dec
    Next
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim i As Integer, L As String
    For i = 1 To Len(Text)
        L = Mid$(Text, i, 1)
        txtOffset = txtOffset & Chr$((Asc(L) + off) Mod 256)
    Next
End Function
