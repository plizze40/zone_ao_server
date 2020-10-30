Attribute VB_Name = "modArena"
Option Explicit



'   '''   '''   '''
'   '0'   '1'   '2'
'   '''   '''   '''
'
'   '''   '''   '''
'   '3'   '4'   '5'
'   '''   '''   '''
'
'   '''   '''   '''
'   '6'   '7'   '8'
'   '''   '''   '''


Private Const START_X = 10
Private Const START_Y = 10

Private Const OFFSET_X = 10
Private Const OFFSET_Y = 10

Private Const ARENA_COUNT = 9
Private Const MAP_ID = 5    ' <-CAMBIAR

Private Const MAX_MEMBER_TEAM = 3
Private Const MIN_MEMBER_TEAM = 1

Private Type Team
    Members(MAX_MEMBER_TEAM - 1) As Long
    MembersCount As Long
End Type

Private Type Ring
    MapID As Long
    sX As Long
    sY As Long
    eX As Long
    eY As Long
    Available As Boolean

    LeftTeam As Team
    RightTeam As Team
End Type

Public Type RingInfo
    MapID As Long
    sX As Long
    sY As Long
    eX As Long
    eY As Long
End Type

Private Rings(ARENA_COUNT - 1) As Ring

Public Sub InitializeRings()
    Dim i As Long
    Dim offsetX As Long
    Dim offsetY As Long

    offsetX = START_X
    offsetY = START_Y

    For i = 0 To ARENA_COUNT - 1
        With Rings(i)
            .sX = offsetX
            .sY = offsetY
            .eX = .sX + OFFSET_X
            .eY = .sY + OFFSET_Y
            .MapID = MAP_ID

            offsetX = .eX + OFFSET_X
            If (i + 1) Mod 3 = 0 Then
                offsetX = START_X
                offsetY = .eY + OFFSET_Y
            End If
        End With
    Next i
End Sub

Public Function FindFreeRing() As Long
    Dim i As Long
    For i = 0 To ARENA_COUNT - 1
        With Rings(i)
            If .Available Then
                FindFreeRing = i
                Exit Function
            End If
        End With
    Next i

    FindFreeRing = -1
End Function

Public Sub FreeRing(ByVal index As Long)
    With Rings(index)
        Call ClearRing(index)
    End With
End Sub

Private Sub ClearRing(ByVal index As Long)
    With Rings(index)
        .Available = True
        .LeftTeam.MembersCount = 0
        .RightTeam.MembersCount = 0
    End With
End Sub

Public Function GetRingInfo(ByVal index As Long) As RingInfo
    With Rings(index)
        GetRingInfo.eX = .eX
        GetRingInfo.eY = .eY
        GetRingInfo.sX = .sX
        GetRingInfo.sY = .sY
        GetRingInfo.MapID = .MapID
    End With
End Function
