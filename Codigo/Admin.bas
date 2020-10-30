Attribute VB_Name = "Admin"
Option Explicit

Public Type tMotd
    Texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public NPCs As Long

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long

Public tInicioServer As Single
Public EstadisticasWeb As New clsEstadisticasIPC

Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizadoUsuario As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloMover As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Single
Public IntervaloUserFlechas As Single
Public IntervaloUserSH As Single
Public IntervaloUserPuedeGolpeHechi As Single
Public IntervaloUserPuedeHechiGolpe As Single
Public IntervaloUserPuedePocion As Single
Public IntervaloUserPuedePocionC As Single
Public IntervaloUserPuedeCastear As Single
Public IntervaloFlechasCazadores As Single
Public IntervaloUserPuedeUsar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long
Public MinutosWs As Long
Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte
Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
    On Error Resume Next

    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    AuxInteger = SD(N)
    AuxInteger2 = SDM(N)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)

End Function
Sub ReSpawnOrigPosNpcs()
    On Error Resume Next
    Dim i As Integer
    Dim MiNPC As Npc

    For i = 1 To LastNPC
        If Npclist(i).flags.NPCActive Then
            If InMapBounds(Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
            End If

            If Npclist(i).Contadores.TiempoExistencia Then
                Call MuereNpc(i, 0)
            End If
        End If
    Next

End Sub
Sub WorldSave()
    On Error Resume Next

    Dim LoopX As Integer
    Dim Porc As Long

    Call ReSpawnOrigPosNpcs

    Dim j As Integer, k As Integer

    For j = 1 To NumMaps
        If MapInfo(j).BackUp Then k = k + 1
    Next

    FrmStat.ProgressBar1.MIN = 0
    FrmStat.ProgressBar1.MAX = k
    FrmStat.ProgressBar1.value = 0

    For LoopX = 1 To NumMaps


        If MapInfo(LoopX).BackUp Then
            Call SaveMapData(LoopX)
            FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
        End If

    Next

    FrmStat.Visible = False

    If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
    If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

    For LoopX = 1 To LastNPC
        If Npclist(LoopX).InvReSpawn Then Call BackUPnPc(LoopX)
    Next

    Call SendData(ToAll, 0, 0, "3P")


End Sub
Public Sub Encarcelar(UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")

    UserList(UserIndex).Counters.TiempoPena = 60 * Minutos
    UserList(UserIndex).flags.Encarcelado = 1
    UserList(UserIndex).Counters.Pena = Timer

    Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)

    If Len(GmName) = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "2O" & Minutos)
    Else
        Call SendData(ToIndex, UserIndex, 0, "3O" & GmName & "," & Minutos)
    End If

End Sub
Public Sub BanTemporal(ByVal Nombre As String, ByVal Dias As Integer, Causa As String, Baneador As String)
    Dim tBan As tBaneo

    Set tBan = New tBaneo
    tBan.Name = UCase$(Nombre)
    tBan.FechaLiberacion = (Now + Dias)
    tBan.Causa = Causa
    tBan.Baneador = UCase$(Baneador)

    Call Baneos.Add(tBan)
    Call SaveBan(Baneos.Count)
    Call SendData(ToAdmins, 0, 0, "||" & Nombre & " fue baneado por " & Causa & " durante los próximos " & Dias & " días." & FONTTYPE_FENIX)

End Sub
