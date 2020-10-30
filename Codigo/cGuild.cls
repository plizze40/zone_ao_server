VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGuild"
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

Public ElectionPeriod As Long

Public Founder As String
Public FundationDate As String
Public Description As String
Private Cod(0 To 7) As String
Public Codex As String
Public Leader As String
Public SubLeader As String
Public Bando As Byte
Public Gold As Double
Public GuildName As String
Public URL As String
Public GuildExperience As Double
Public DaysSinceLastElection As Long
Public GuildNews As String

Public AlliedGuilds As Collection
Public EnemyGuilds As Collection
Public Members As Collection
Public Solicitudes As Collection
Public PeacePropositions As Collection
Public Votes As Collection
Public Elections As Boolean

Private Type tVotes
    Name As String
    Cant As Long
End Type
Public Function Initialize(ByVal GuildInfo As String, ByVal FounderName As String) As Boolean
    Initialize = ParseGuildInfo(GuildInfo, FounderName)
End Function
Private Function ParseGuildInfo(GuildInfo As String, ByVal FounderName As String) As Boolean
    On Error GoTo errhandler

    Founder = FounderName
    Leader = Founder
    FundationDate = Now
    Description = Trim$(ReadField(1, GuildInfo, Asc("¬")))
    GuildName = Trim$(ReadField(2, GuildInfo, Asc("¬")))
    Bando = UserList(NameIndex(FounderName)).Faccion.Bando

    If Not TestName(GuildName) Then
        ParseGuildInfo = False
        Exit Function
    End If

    URL = ReadField(3, GuildInfo, Asc("¬"))
    Codex = ReadField(4, GuildInfo, Asc("¬"))
    GuildNews = "Clan iniciado."
    ParseGuildInfo = True

    Exit Function

errhandler:

End Function
Public Function CodexLenght() As Integer
    Dim T As Integer, NumC As Integer

    For T = 1 To Len(Codex)
        If Mid$(Codex, T, 1) = "|" Then NumC = NumC + 1
    Next

    CodexLenght = NumC + 1

End Function
Private Function TestName(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(Mid$(cad, i, 1))

        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            TestName = False
            Exit Function
        End If

    Next

    TestName = True

End Function
Public Sub UpdateCodexAndDesc(ByVal rdata As String)

    Description = ReadField(1, rdata, Asc("¬"))
    Codex = ReadField(2, rdata, Asc("¬"))

End Sub
Private Sub Class_Initialize()

    Set AlliedGuilds = New Collection
    Set EnemyGuilds = New Collection
    Set Members = New Collection
    Set Solicitudes = New Collection
    Set PeacePropositions = New Collection
    Set Votes = New Collection

    ElectionPeriod = 45

End Sub
Public Function IsEnemy(ByVal GuildName As String) As Boolean
    Dim k As Integer

    GuildName = UCase$(GuildName)

    For k = 1 To EnemyGuilds.Count
        If UCase$(EnemyGuilds(k)) = GuildName Then
            IsEnemy = True
            Exit Function
        End If
    Next

End Function
Public Function IsAllie(ByVal GuildName As String) As Boolean
    Dim k As Integer

    GuildName = UCase$(GuildName)

    For k = 1 To AlliedGuilds.Count
        If UCase$(AlliedGuilds(k)) = GuildName Then
            IsAllie = True
            Exit Function
        End If
    Next

End Function
Public Function SolicitudesIncludes(ByVal UserName As String) As Boolean
    Dim k As Integer

    UserName = UCase$(UserName)

    For k = 1 To Solicitudes.Count
        If UCase$(Solicitudes(k).UserName) = UserName Then
            SolicitudesIncludes = True
            Exit Function
        End If
    Next


End Function
Public Sub RemoveSolicitud(ByVal UserName As String)
    Dim k As Integer

    UserName = UCase$(UserName)

    For k = 1 To Solicitudes.Count
        If UCase$(Solicitudes(k).UserName) = UserName Then
            Call Solicitudes.Remove(k)
            Exit Sub
        End If
    Next

End Sub

Public Sub RemoveProposition(ByVal UserName As String)

    Dim k As Integer


    UserName = UCase$(UserName)

    For k = 1 To PeacePropositions.Count
        If UCase$(PeacePropositions(k).UserName) = UserName Then
            Call PeacePropositions.Remove(k)
            Exit Sub
        End If
    Next



End Sub


Public Sub RemoveAllie(ByVal Name As String)
    Dim k As Integer
    Name = UCase$(Name)
    For k = 1 To AlliedGuilds.Count
        If UCase$(AlliedGuilds(k)) = Name Then
            Call AlliedGuilds.Remove(k)
            Exit Sub
        End If
    Next
End Sub
Public Sub RemoveEnemy(ByVal Name As String)
    Dim k As Integer

    Name = UCase$(Name)

    For k = 1 To EnemyGuilds.Count
        If UCase$(EnemyGuilds(k)) = Name Then
            Call EnemyGuilds.Remove(k)
            Exit Sub
        End If
    Next

End Sub
Public Sub RemoveMember(ByVal UserName As String)
    Dim k As Integer

    UserName = UCase$(UserName)

    For k = 1 To Members.Count
        If UCase$(Members(k)) = UserName Then
            Call Members.Remove(k)
            Exit Sub
        End If
    Next

End Sub
Public Function GetPeaceRequest(ByVal UserName As String) As cSolicitud
    Dim k As Integer

    UserName = UCase$(UserName)

    For k = 1 To PeacePropositions.Count
        If UCase$(PeacePropositions(k).UserName) = UserName Then
            Set GetPeaceRequest = PeacePropositions(k)
            Exit Function
        End If
    Next

End Function
Public Sub TestSolicitudBound()

    If Me.Solicitudes.Count > 25 Then
        Dim i As Integer
        For i = 1 To Me.Solicitudes.Count
            Solicitudes.Remove 1
        Next
    End If

End Sub
Public Function GetSolicitud(ByVal UserName As String) As cSolicitud
    Dim k As Integer

    UserName = UCase$(UserName)

    For k = 1 To Solicitudes.Count
        If UCase$(Solicitudes(k).UserName) = UCase$(UserName) Then
            Set GetSolicitud = Solicitudes(k)
            Exit Function
        End If
    Next


End Function

Public Function IncludesPeaceOffer(ByVal Name As String) As Boolean

    Dim k As Integer

    Name = UCase$(Name)

    For k = 1 To PeacePropositions.Count
        If UCase$(PeacePropositions(k).UserName) = Name Then
            IncludesPeaceOffer = True
            Exit Function
        End If
    Next



End Function

Public Sub ResetVotes()

    Dim k As Integer
    For k = 1 To Votes.Count
        Call Votes.Remove(1)
    Next

End Sub


Public Function IsMember(ByVal Name As String) As Boolean

    Dim k As Integer

    Name = UCase$(Name)

    For k = 1 To Members.Count
        If UCase$(Members(k)) = Name Then
            IsMember = True
            Exit Function
        End If
    Next

End Function


Public Function NuevoLider() As String

    Dim j As Integer, c As Integer

    Dim totals() As tVotes

    If Members.Count = 0 Then
        NuevoLider = "CLANCLAUSURADO"
        Exit Function
    End If

    ReDim totals(1 To Members.Count)


    For j = 1 To Members.Count

        totals(j).Name = Members(j)
        totals(j).Cant = 0

    Next


    For j = 1 To Votes.Count

        For c = 1 To Members.Count
            If UCase$(totals(c).Name) = UCase$(Votes(j)) Then
                totals(c).Cant = totals(c).Cant + 1
                Exit For
            End If
        Next

    Next

    Dim MAX As Integer
    MAX = -1

    For j = 1 To Members.Count
        If totals(j).Cant > MAX Then
            MAX = totals(j).Cant
            NuevoLider = totals(j).Name
        End If
    Next

End Function
