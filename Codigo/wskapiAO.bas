Attribute VB_Name = "wskapiAO"
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

#If UsarQueSocket = 1 Then
    #Const WSAPI_CREAR_LABEL = True

    Private Const SD_RECEIVE As Long = &H0
    Private Const SD_SEND As Long = &H1
    Private Const SD_BOTH As Long = &H2

    Private Const MAX_TIEMPOIDLE_COLALLENA = 1
    Private Const MAX_COLASALIDA_COUNT = 800

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

    Private Const WS_CHILD = &H40000000
    Public Const GWL_WNDPROC = (-4)

Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr As New Collection

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long
Public SockListen As Long
#End If
Public Sub IniciaWsApi()
    #If UsarQueSocket = 1 Then

        hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)
        OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
        ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

        Dim Desc As String
        Call StartWinsock(Desc)

    #End If
End Sub
Public Function StartWinsock(sDescription As String) As Boolean

    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        Else: WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp

End Function
Public Sub LimpiaWsApi(ByVal hwnd As Long)
    #If UsarQueSocket = 1 Then

        If WSAStartedUp Then
            Call EndWinsock
        End If

    #End If
End Sub
Public Function BuscaSlotSock(S As Long) As Long
    #If UsarQueSocket = 1 Then
        Dim i As Integer

        For i = 1 To MaxUsers
            If UserList(i).ConnID = S Then
                BuscaSlotSock = i
                Exit Function
            End If
        Next i

        BuscaSlotSock = -1

    #End If
End Function
Public Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    #If UsarQueSocket = 1 Then

        Dim Ret As Long
        Dim Tmp As String

        WndProc = 0

        Select Case msg
        Case 1025
            Dim S As Long, E As Long
            Dim N As Integer
            S = wParam
            E = WSAGetSelectEvent(lParam)

            Select Case E
            Case FD_ACCEPT
                Call EventoSockAccept(S)
            Case FD_READ

                N = BuscaSlotSock(S)
                If N < 0 Then
                    Call ApiCloseSocket(S)
                    Exit Function
                End If


                Tmp = Space(4096)

                Ret = recv(S, Tmp, Len(Tmp), 0)

                If Ret < 0 Then
                    Call CloseSocket(N)
                    Exit Function
                End If

                Tmp = Left(Tmp, Ret)

                Call EventoSockRead(N, Tmp)

            Case FD_CLOSE
                N = BuscaSlotSock(S)

                If N < 0 Then
                    Call ApiCloseSocket(S)
                Else: Call EventoSockClose(N)
                End If

            End Select
        Case Else
            WndProc = CallWindowProc(OldWProc, hwnd, msg, wParam, lParam)
        End Select

    #End If
End Function
Public Sub WsApiEnviar(Slot As Integer, str As String)
    #If UsarQueSocket = 1 Then
        Dim Ret As String

        If UserList(Slot).ConnID > -1 Then
            Ret = send(ByVal UserList(Slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)
            If Ret < 0 Then Exit Sub
        End If
    #End If

End Sub
Public Sub LogCustom(ByVal str As String)
    #If (UsarQueSocket = 1) Then

        Dim nfile As Integer
        nfile = FreeFile
        Open App.Path & "\logs\custom.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & str
        Close #nfile

        Exit Sub

errhandler:

    #End If
End Sub
Public Sub EventoSockAccept(SockID As Long)
    #If UsarQueSocket = 1 Then
        On Error GoTo ErrorHandler
        Dim NewIndex As Integer
        Dim Ret As Long
        Dim Tam As Long, sa As sockaddr
        Dim NuevoSock As Long
        Dim i As Long

        NewIndex = NextOpenUser

        If NewIndex <= MaxUsers Then

            Tam = sockaddr_size

            Ret = accept(SockID, sa, Tam)
            If Ret = INVALID_SOCKET Then
                Call LogCriticEvent("Error en Accept() API")
                Exit Sub
            End If
            NuevoSock = Ret

            UserList(NewIndex).ip = GetAscIP(sa.sin_addr)

            For i = 1 To BanIps.Count
                If BanIps.Item(i) = UserList(NewIndex).ip Then
                    Call ApiCloseSocket(NuevoSock)
                    Exit Sub
                End If
            Next i

            If aDos.MaxConexiones(UserList(NewIndex).ip) Then
                UserList(NewIndex).ConnID = -1
                Call aDos.RestarConexion(UserList(NewIndex).ip)
                Call ApiCloseSocket(NuevoSock)
            End If

            UserList(NewIndex).ConnID = NuevoSock
            Set UserList(NewIndex).CommandsBuffer = New CColaArray

        Else
            Call LogCriticEvent("No acepte conexion porque no tenia slots")
        End If
        Exit Sub
ErrorHandler:
        Call LogError("Error en EventoSockAccept. " & Err.Description)
        Call CloseSocket(NewIndex)


    #End If
End Sub
Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos As String)
    #If UsarQueSocket = 1 Then
        On Error GoTo ErrorHandler
        Dim T() As String
        Dim LoopC As Long

        UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos

        If InStr(1, UserList(Slot).RDBuffer, Chr(2)) > 0 Then UserList(Slot).RDBuffer = "CLIENTEVIEJO" & ENDC

        T = Split(UserList(Slot).RDBuffer, ENDC)
        If UBound(T) > 0 Then
            UserList(Slot).RDBuffer = T(UBound(T))

            For LoopC = 0 To UBound(T) - 1
                If ClientsCommandsQueue = 1 Then
                    If Len(T(LoopC)) > 0 Then If Not UserList(Slot).CommandsBuffer.Push(T(LoopC)) Then Call Cerrar_Usuario(Slot)
                Else
                    If UserList(Slot).ConnID <> -1 Then
                        Call HandleData(Slot, T(LoopC))
                    Else: Exit Sub
                    End If
                End If
            Next LoopC
        End If

        Exit Sub

ErrorHandler:
        Call LogError("Error en Socket read. " & Err.Description)
        Call CloseSocket(Slot)

    #End If
End Sub
#If Win16 Then
Public Function kSendData(ByVal S%, vMessage As Variant) As Integer
#ElseIf Win32 Then
Public Function kSendData(ByVal S&, vMessage As Variant) As Long
#End If
Dim TheMsg() As Byte, sTemp$
TheMsg = ""
Select Case VarType(vMessage)
Case 8209
    sTemp = vMessage
    TheMsg = sTemp
Case 8
    #If Win32 Then
        sTemp = StrConv(vMessage, vbFromUnicode)
    #Else
        sTemp = vMessage
    #End If
Case Else
    sTemp = CStr(vMessage)
    #If Win32 Then
        sTemp = StrConv(vMessage, vbFromUnicode)
    #Else
        sTemp = vMessage
    #End If
End Select
TheMsg = sTemp
If UBound(TheMsg) > -1 Then
    kSendData = send(S, TheMsg(0), UBound(TheMsg) + 1, 0)
End If
End Function
Public Sub EventoSockClose(Slot As Integer)

    #If UsarQueSocket = 1 Then
        If UserList(Slot).flags.UserLogged Then
            Call Cerrar_Usuario(Slot)
        Else: Call CloseSocket(Slot)
        End If
    #End If

End Sub


Public Sub WSApiReiniciarSockets()
    #If UsarQueSocket = 1 Then
        Dim i As Long

        If SockListen >= 0 Then Call ApiCloseSocket(SockListen)

        For i = 1 To MaxUsers
            If UserList(i).ConnID <> -1 And UserList(i).ConnIDvalida Then
                Call CloseSocket(i)
            End If
        Next i


        ReDim UserList(1 To MaxUsers)
        For i = 1 To MaxUsers
            UserList(i).ConnID = -1
            UserList(i).ConnIDvalida = False
        Next i

        LastUser = 1
        NumUsers = 0

        Call LimpiaWsApi(frmMain.hwnd)
        Call Sleep(100)
        Call IniciaWsApi
        SockListen = ListenForConnect(Puerto, hWndMsg, "")

    #End If
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)
    #If UsarQueSocket = 1 Then
        Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
        Call ShutDown(Socket, SD_BOTH)
    #End If
End Sub
