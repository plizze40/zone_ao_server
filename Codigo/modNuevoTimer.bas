Attribute VB_Name = "modNuevoTimer"
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

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Dim hGameTimer As Long
Dim hNpcCanAttack As Long
Dim hNpcAITimer As Long
Dim hAutoTimer As Long
Public Sub NpcAITimer(Enabled As Boolean)
    On Error GoTo Error

    If Enabled Then
        If hNpcAITimer Then KillTimer 0, hNpcAITimer
        hNpcAITimer = SetTimer(0, 0, 420, AddressOf NpcAITimerProc)
    Else
        If hNpcAITimer = 0 Then Exit Sub
        KillTimer 0, hNpcAITimer
        hNpcAITimer = 0
    End If

    Exit Sub
Error:
    Call LogError("Error en NpcAiTimer: " & Err.Description)
End Sub
Sub NpcAITimerProc(ByVal hwnd As Long, ByVal hEvent As Long, ByVal TimerID As Long, ByVal lpTimerFunc As Long)
    On Error GoTo ErrorHandler
    Dim NpcIndex As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim UseAI As Integer

    If Not haciendoBK Then
        For NpcIndex = 1 To LastNPC
            If Npclist(NpcIndex).flags.NPCActive Then
                If Npclist(NpcIndex).flags.Paralizado = 0 Then
                    If Npclist(NpcIndex).POS.Map Then
                        If MapInfo(Npclist(NpcIndex).POS.Map).NumUsers And Npclist(NpcIndex).Movement <> ESTATICO Then Call NPCMovementAI(NpcIndex)
                    End If
                ElseIf Npclist(NpcIndex).flags.Paralizado = 2 Then Call NPCAtacaAlFrente(NpcIndex)
                End If
            End If
        Next
    End If

    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).POS.Map)
    Call MuereNpc(NpcIndex, 0)

End Sub
Public Sub NpcCanAttack(Enabled As Boolean)
    On Error GoTo Error

    If Enabled Then
        If hNpcCanAttack Then KillTimer 0, hNpcCanAttack
        hNpcCanAttack = SetTimer(0, 0, 3000, AddressOf NpcCanAttackProc)
    Else
        If hNpcCanAttack = 0 Then Exit Sub
        KillTimer 0, hNpcCanAttack
        hNpcCanAttack = 0
    End If

    Exit Sub
Error:
    Call LogError("Error en NpcCanAttack: " & Err.Description)

End Sub
Sub NpcCanAttackProc(ByVal hwnd As Long, ByVal hEvent As Long, ByVal TimerID As Long, ByVal lpTimerFunc As Long)
    Dim Npc As Integer
    On Error GoTo Error

    For Npc = 1 To LastNPC
        If Npclist(Npc).flags.NPCActive And Npclist(Npc).POS.Map Then
            If Npclist(Npc).Numero <> 89 Then Npclist(Npc).CanAttack = 1
            If Npclist(Npc).flags.Paralizado Then Call EfectoParalisisNpc(Npc)
        End If
    Next Npc

    Exit Sub
Error:
    Call LogError("Error en NpcCanAttackProc: " & Err.Description)
End Sub
Public Sub AutoTimer(Enabled As Boolean)
    On Error GoTo Error

    If Enabled Then
        If hAutoTimer Then KillTimer 0, hAutoTimer
        hAutoTimer = SetTimer(0, 0, 60000, AddressOf AutoTimerProc)
    Else
        If hAutoTimer = 0 Then Exit Sub
        KillTimer 0, hAutoTimer
        hAutoTimer = 0
    End If
    Exit Sub
Error:
    Call LogError("Error en AutoTimer:" & Err.Description)

End Sub
Public Sub EfectoParalisisNpc(NpcIndex As Integer)

    If Npclist(NpcIndex).Contadores.Paralisis Then
        Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 2
    Else
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).flags.QuienParalizo = 0
    End If

End Sub
Sub RegistrarDataDia()
    On Error GoTo errhandler
    Dim FileAntiguo As String

    FileAntiguo = App.Path & "\LOGS\Data " & Format(Now - 7, "dd-mm") & ".log"
    If FileExist(FileAntiguo, vbNormal) Then Call Kill(FileAntiguo)

    Dim nfile As Integer
    nfile = FreeFile

    Open App.Path & "\LOGS\Data\Data " & Format(Now, "dd-mm") & ".log" For Append Shared As #nfile
    Print #nfile, "### " & Format(Now, "dd/mm") & " ###"
    Print #nfile, "Enviada: " & Data(Dia, Actual, Enviada, Mensages) & "/" & Data(Dia, Actual, Enviada, Letras) & "(" & Round(Data(Dia, Actual, Enviada, Letras) / 1048576, 2) & " mb)"
    Print #nfile, "Recibida: " & Data(Dia, Actual, Recibida, Mensages) & "/" & Data(Dia, Actual, Recibida, Letras) & "(" & Round(Data(Dia, Actual, Recibida, Letras) / 1048576, 2) & " mb)"
    Print #nfile, "Users Online: " & Round(Onlines(Dia) / 1440, 2)
    Print #nfile, ""
    Close #nfile

    Exit Sub
errhandler:

End Sub
Sub RegistrarData()
    On Error GoTo errhandler
    Dim FileAntiguo As String

    FileAntiguo = App.Path & "\LOGS\Data " & Format(Now - 7, "dd-mm") & ".log"
    If FileExist(FileAntiguo, vbNormal) Then Call Kill(FileAntiguo)

    Dim nfile As Integer
    nfile = FreeFile

    Open App.Path & "\LOGS\Data\Data " & Format(Now, "dd-mm") & ".log" For Append Shared As #nfile
    Print #nfile, "### " & Format(Now - 1 / 24, "hh:mm") & "-" & Format(Now, "hh:mm") & " ###"
    Print #nfile, "Enviada: " & Data(Hora, Actual, Enviada, Mensages) & "/" & Data(Hora, Actual, Enviada, Letras) & "(" & Round(Data(Hora, Actual, Enviada, Letras) / 1048576, 2) & " mb)"
    Print #nfile, "Recibida: " & Data(Hora, Actual, Recibida, Mensages) & "/" & Data(Hora, Actual, Recibida, Letras) & "(" & Round(Data(Hora, Actual, Recibida, Letras) / 1048576, 2) & " mb)"
    Print #nfile, "Users Online: " & Round(Onlines(Actual) / 60, 2)
    Print #nfile, ""
    Close #nfile

    Exit Sub
errhandler:

End Sub
Sub PasarDataDia()

    Call RegistrarDataDia


    Data(Dia, Actual, Recibida, Mensages) = 0
    Data(Dia, Actual, Recibida, Letras) = 0

    Data(Dia, Actual, Enviada, Mensages) = 0
    Data(Dia, Actual, Enviada, Letras) = 0

    Onlines(Dia) = 0

End Sub
Sub PasarDataHora()

    Call RegistrarData


    Data(Dia, Actual, Recibida, Mensages) = Data(Dia, Actual, Recibida, Mensages) + Data(Hora, Actual, Recibida, Mensages)
    Data(Dia, Actual, Recibida, Letras) = Data(Dia, Actual, Recibida, Letras) + Data(Hora, Actual, Recibida, Letras)

    Data(Dia, Actual, Enviada, Mensages) = Data(Dia, Actual, Enviada, Mensages) + Data(Hora, Actual, Enviada, Mensages)
    Data(Dia, Actual, Enviada, Letras) = Data(Dia, Actual, Enviada, Letras) + Data(Hora, Actual, Enviada, Letras)

    Onlines(Last) = Onlines(Actual)
    Onlines(Dia) = Onlines(Dia) + Onlines(Actual)


    Data(Hora, Last, Recibida, Mensages) = Data(Hora, Actual, Recibida, Mensages)
    Data(Hora, Last, Recibida, Letras) = Data(Hora, Actual, Recibida, Letras)

    Data(Hora, Last, Enviada, Mensages) = Data(Hora, Actual, Enviada, Mensages)
    Data(Hora, Last, Enviada, Letras) = Data(Hora, Actual, Enviada, Letras)


    Data(Hora, Actual, Recibida, Mensages) = 0
    Data(Hora, Actual, Recibida, Letras) = 0

    Data(Hora, Actual, Enviada, Mensages) = 0
    Data(Hora, Actual, Enviada, Letras) = 0

    Onlines(Actual) = 0

End Sub
Sub PasarDataMinuto()


    Data(Hora, Actual, Recibida, Mensages) = Data(Hora, Actual, Recibida, Mensages) + Data(Minuto, Actual, Recibida, Mensages)
    Data(Hora, Actual, Recibida, Letras) = Data(Hora, Actual, Recibida, Letras) + Data(Minuto, Actual, Recibida, Letras)

    Data(Hora, Actual, Enviada, Mensages) = Data(Hora, Actual, Enviada, Mensages) + Data(Minuto, Actual, Enviada, Mensages)
    Data(Hora, Actual, Enviada, Letras) = Data(Hora, Actual, Enviada, Letras) + Data(Minuto, Actual, Enviada, Letras)

    Onlines(Actual) = Onlines(Actual) + NumUsers


    Data(Minuto, Last, Recibida, Mensages) = Data(Minuto, Actual, Recibida, Mensages)
    Data(Minuto, Last, Recibida, Letras) = Data(Minuto, Actual, Recibida, Letras)

    Data(Minuto, Last, Enviada, Mensages) = Data(Minuto, Actual, Enviada, Mensages)
    Data(Minuto, Last, Enviada, Letras) = Data(Minuto, Actual, Enviada, Letras)


    Data(Minuto, Actual, Recibida, Mensages) = 0
    Data(Minuto, Actual, Recibida, Letras) = 0

    Data(Minuto, Actual, Enviada, Mensages) = 0
    Data(Minuto, Actual, Enviada, Letras) = 0

End Sub
Sub AutoTimerProc(ByVal hwnd As Long, ByVal hEvent As Long, ByVal TimerID As Long, ByVal lpTimerFunc As Long)
    Static Minutos As Long
    Static MinsSocketReset As Long
    Static MinsPjesSave As Long
    Static MinutosCodigoTrabajar As Long
    Dim i As Integer
    Static MinWSLargo As Long
    On Error GoTo errhandler

    Call ComprobarCerrar

    For i = 1 To Baneos.Count
        If Baneos(i).FechaLiberacion <= Now Then
            Call SendData(ToAdmins, 0, 0, "||Se ha concluido la sentencia de ban de " & Baneos(i).Name & "." & FONTTYPE_FIGHT)
            Call ChangeBan(Baneos(i).Name, 0)
            Call Baneos.Remove(i)
            Call SaveBans
        End If
    Next

    If Len(MensajeRepeticion) > 0 Then
        If TiempoRepeticion > 0 Then
            TiempoRepeticion = TiempoRepeticion - 1
            If TiempoRepeticion Mod IntervaloRepeticion = 0 Then Call SendData(ToAll, 0, 0, "||" & MensajeRepeticion & FONTTYPE_TALK & ENDC)
            If TiempoRepeticion = 0 Then
                Call SendData(ToAdmins, 0, 0, "||Se ha terminado la repetición del mensaje: " & MensajeRepeticion & "." & FONTTYPE_FENIX)
                IntervaloRepeticion = 0
                MensajeRepeticion = ""
            End If
        Else
            TiempoRepeticion = TiempoRepeticion + 1
            If TiempoRepeticion Mod IntervaloRepeticion = 0 Then Call SendData(ToAll, 0, 0, "||" & MensajeRepeticion & FONTTYPE_TALK & ENDC)
            If TiempoRepeticion = 0 Then TiempoRepeticion = -IntervaloRepeticion
        End If
    End If

    Minutos = Minutos + 1
    MinWSLargo = MinWSLargo + 1

    Call MostrarNumUsers

    If MinWSLargo >= 240 Then
        Call DoBackUp(True)
        MinWSLargo = 0
        Minutos = 0
    End If

    If Minutos >= 90 Then
        Call DoBackUp
        Call aClon.VaciarColeccion
        Minutos = 0
    End If

    If Time >= #6:28:00 AM# And Time <= #6:29:01 AM# And Worldsaves Then
        Call SendData(ToAll, 0, 0, "||Un nuevo día ha comenzado..." & FONTTYPE_FENIX)
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        Call DayElapsed
    End If


    Dim N As Integer
    N = FreeFile(1)
    Open App.Path & "\LOGS\numusers.log" For Output Shared As N
    Print #N, NumUsers
    Close #N

    Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave")

End Sub
Public Sub ComprobarCerrar()

    If val(GetVar(App.Path & "\Executor.ini", "EXECUTOR", "Cerrar")) = 1 Then
        Call LogMain(" Server apagado por el Executor.")
        Call WriteVar(App.Path & "\Executor.ini", "EXECUTOR", "Cerrar", 0)
        Call DoBackUp(True)
        End
    End If

End Sub
