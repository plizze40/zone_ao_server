Attribute VB_Name = "PathFinding"
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

Private Const ROWS = 100
Private Const COLUMS = 100
Private Const MAXINT = 1000
Private Const Walkable = 0

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Dim TilePosX As Integer, TilePosY As Integer
Attribute TilePosY.VB_VarUserMemId = 1073741825

Dim MyVert As tVertice
Attribute MyVert.VB_VarUserMemId = 1073741827
Dim MyFin As tVertice
Attribute MyFin.VB_VarUserMemId = 1073741828

Dim Iter As Integer
Attribute Iter.VB_VarUserMemId = 1073741829

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
    Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS
End Function

Private Function IsWalkable(Map As Integer, ByVal row As Integer, ByVal Col As Integer, NpcIndex As Integer) As Boolean
    IsWalkable = MapData(Map, row, Col).Blocked = 0 And MapData(Map, row, Col).NpcIndex = 0

    If MapData(Map, row, Col).UserIndex Then
        If MapData(Map, row, Col).UserIndex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False
    End If

End Function

Private Sub ProcessAdjacents(MapIndex As Integer, T() As tIntermidiateWork, vfila As Integer, vcolu As Integer, NpcIndex As Integer)
    Dim V As tVertice
    Dim j As Integer

    j = vfila - 1
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then

            If T(j, vcolu).DistV = MAXINT Then

                T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                T(j, vcolu).PrevV.X = vcolu
                T(j, vcolu).PrevV.Y = vfila

                V.X = vcolu
                V.Y = j
                Call Push(V)
            End If
        End If
    End If
    j = vfila + 1

    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then

            If T(j, vcolu).DistV = MAXINT Then

                T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                T(j, vcolu).PrevV.X = vcolu
                T(j, vcolu).PrevV.Y = vfila

                V.X = vcolu
                V.Y = j
                Call Push(V)
            End If
        End If
    End If

    If Limites(vfila, vcolu - 1) Then
        If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then

            If T(vfila, vcolu - 1).DistV = MAXINT Then

                T(vfila, vcolu - 1).DistV = T(vfila, vcolu).DistV + 1
                T(vfila, vcolu - 1).PrevV.X = vcolu
                T(vfila, vcolu - 1).PrevV.Y = vfila

                V.X = vcolu - 1
                V.Y = vfila
                Call Push(V)
            End If
        End If
    End If

    If Limites(vfila, vcolu + 1) Then
        If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then

            If T(vfila, vcolu + 1).DistV = MAXINT Then

                T(vfila, vcolu + 1).DistV = T(vfila, vcolu).DistV + 1
                T(vfila, vcolu + 1).PrevV.X = vcolu
                T(vfila, vcolu + 1).PrevV.Y = vfila

                V.X = vcolu + 1
                V.Y = vfila
                Call Push(V)
            End If
        End If
    End If


End Sub


Public Sub SeekPath(NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)







    Dim cur_npc_pos As tVertice
    Dim tar_npc_pos As tVertice
    Dim V As tVertice
    Dim NpcMap As Integer
    Dim steps As Integer

    NpcMap = Npclist(NpcIndex).POS.Map

    steps = 0

    cur_npc_pos.X = Npclist(NpcIndex).POS.Y
    cur_npc_pos.Y = Npclist(NpcIndex).POS.X

    tar_npc_pos.X = Npclist(NpcIndex).PFINFO.Target.X
    tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y

    Call InitializeTable(TmpArray, cur_npc_pos)
    Call InitQueue


    Call Push(cur_npc_pos)

    Do While (Not IsEmpty)
        If steps > MaxSteps Then Exit Do
        V = Pop
        If V.X = tar_npc_pos.X And V.Y = tar_npc_pos.Y Then Exit Do
        Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
    Loop

    Call MakePath(NpcIndex)

End Sub

Private Sub MakePath(NpcIndex As Integer)




    Dim Pasos As Integer
    Dim miV As tVertice
    Dim i As Integer

    Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.X).DistV
    Npclist(NpcIndex).PFINFO.PathLenght = Pasos


    If Pasos = MAXINT Then

        Npclist(NpcIndex).PFINFO.NoPath = True
        Npclist(NpcIndex).PFINFO.PathLenght = 0
        Exit Sub
    End If

    ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice

    miV.X = Npclist(NpcIndex).PFINFO.Target.X
    miV.Y = Npclist(NpcIndex).PFINFO.Target.Y

    For i = Pasos To 1 Step -1
        Npclist(NpcIndex).PFINFO.Path(i) = miV
        miV = TmpArray(miV.Y, miV.X).PrevV
    Next

    Npclist(NpcIndex).PFINFO.CurPos = 1
    Npclist(NpcIndex).PFINFO.NoPath = False

End Sub

Private Sub InitializeTable(T() As tIntermidiateWork, S As tVertice, Optional ByVal MaxSteps As Integer = 30)




    Dim j As Integer, k As Integer
    Const anymap = 1
    For j = S.Y - MaxSteps To S.Y + MaxSteps
        For k = S.X - MaxSteps To S.X + MaxSteps
            If InMapBounds(j, k) Then
                T(j, k).Known = False
                T(j, k).DistV = MAXINT
                T(j, k).PrevV.X = 0
                T(j, k).PrevV.Y = 0
            End If
        Next
    Next

    T(S.Y, S.X).Known = False
    T(S.Y, S.X).DistV = 0

End Sub

