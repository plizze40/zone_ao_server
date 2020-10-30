VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const FRENTE = 1

Private Cola As Collection

Public Sub Reset()
    On Error Resume Next

    Dim i As Integer
    For i = 1 To Me.Longitud
        Cola.Remove FRENTE
    Next

End Sub

Public Property Get Longitud() As Integer
    Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
    IndexValido = i >= 1 And i <= Me.Longitud
End Function

Private Sub Class_Initialize()
    Set Cola = New Collection
End Sub
Public Function VerElemento(Index As Integer) As Variant
    On Error Resume Next

    If IndexValido(Index) Then
        VerElemento = Cola(Index)
    Else: VerElemento = 0
    End If

End Function
Public Sub Push(info As String, Nombre As String)
    On Error Resume Next

    Call Cola.Add(UCase$(Nombre))

End Sub
Public Function Pop() As Variant
    On Error Resume Next

    If Cola.Count Then
        Pop = Cola(FRENTE)
        Call Cola.Remove(FRENTE)
    End If

End Function
Public Function PopByVal() As Variant
    On Error Resume Next

    If Cola.Count Then PopByVal = Cola.Item(1)

End Function
Public Function Existe(ByVal Nombre As String) As Boolean
    On Error Resume Next
    Dim f As Boolean
    Dim V As Variant
    Dim i As Integer

    Dim NombreEnMayusculas As String
    NombreEnMayusculas = UCase$(Nombre)

    i = 1
    f = False

    For i = 1 To Me.Longitud
        V = Me.VerElemento(i)
        If V = NombreEnMayusculas Then
            Existe = True
            Exit Function
        End If
    Next

End Function
Public Sub Quitar(Nombre As String)
    On Error Resume Next
    Dim V As Variant
    Dim i As Integer
    Dim NombreEnMayusculas As String

    NombreEnMayusculas = UCase$(Nombre)

    For i = 1 To Me.Longitud
        V = Me.VerElemento(i)
        If V = NombreEnMayusculas Then
            Call Cola.Remove(i)
            Exit Sub
        End If
    Next

End Sub
Public Sub QuitarIndex(Index As Integer)
    On Error Resume Next

    If IndexValido(Index) Then Call Cola.Remove(Index)

End Sub
Private Sub Class_Terminate()

    Set Cola = Nothing

End Sub
