VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servidor FlamiusAO  ~ Argentum Online ~"
   ClientHeight    =   3540
   ClientLeft      =   1950
   ClientTop       =   1695
   ClientWidth     =   6615
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000007&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Tlimpiar 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   720
   End
   Begin VB.Timer retos2vs2 
      Interval        =   60000
      Left            =   3360
      Top             =   720
   End
   Begin VB.Timer TimerMeditar 
      Interval        =   400
      Left            =   2880
      Top             =   720
   End
   Begin VB.Data ADODB 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox i 
      Height          =   3180
      ItemData        =   "frmMain.frx":1042
      Left            =   5160
      List            =   "frmMain.frx":1049
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Mensaje BroadCast >>"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.Timer timeravisar 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   1800
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   15000
         Left            =   2280
         Top             =   120
      End
      Begin VB.Timer torneos 
         Interval        =   15000
         Left            =   2280
         Top             =   600
      End
      Begin VB.Timer TimerTrabaja 
         Interval        =   10000
         Left            =   4200
         Top             =   120
      End
      Begin VB.Timer CmdExec 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3240
         Tag             =   "S"
         Top             =   120
      End
      Begin VB.Timer UserTimer 
         Interval        =   1000
         Left            =   2760
         Top             =   120
      End
      Begin VB.Timer TimerFatuo 
         Interval        =   2500
         Left            =   3720
         Top             =   120
      End
      Begin VB.Timer tRevisarCabs 
         Left            =   10000
         Top             =   480
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   4200
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         Protocol        =   4
         RemoteHost      =   "fenixao.localstrike.com.ar"
         URL             =   "http://fenixao.localstrike.com.ar/descargas/Clave.txt"
         Document        =   "/descargas/Clave.txt"
         RequestTimeout  =   30
      End
      Begin VB.Label CantUsuarios 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblCantUsers 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Usuarios Online:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2400
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensaje BroadCast:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
      Begin VB.TextBox BroadMsg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar Mensaje BroadCast"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   4695
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   6480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   0
      X2              =   5160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "&FlamiusAO"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "SysTray Servidor"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar Servidor"
      End
      Begin VB.Menu mnuSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Cerrar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function
Private Sub CmdExec_Timer()
    On Error Resume Next

    #If UsarQueSocket = 1 Then
        Dim i As Integer

        For i = 1 To MaxUsers
            If UserList(i).ConnID <> -1 Then
                If Not UserList(i).CommandsBuffer.IsEmpty Then Call HandleData(i, UserList(i).CommandsBuffer.Pop)
            End If
        Next i

    #End If

End Sub
Private Sub cmdMore_Click()

    If cmdMore.caption = "Mensaje BroadCast >>" Then
        Me.Height = 4395
        cmdMore.caption = "<< Ocultar"
    Else
        Me.Height = 2070
        cmdMore.caption = "Mensaje BroadCast >>"
    End If

End Sub

Private Sub Command1_Click()
    Call SendData(ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub
Public Sub InitMain(f As Byte)

    If f Then
        Call mnuSystray_Click
    Else: frmMain.Show
    End If

End Sub
Private Sub Form_Load()

    Call mnuSystray_Click
    Codifico = RandomNumber(1, 99)

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX

        Case WM_LBUTTONDBLCLK
            WindowState = vbNormal
            Visible = True
            Dim hProcess As Long
            GetWindowThreadProcessId hwnd, hProcess
            AppActivate hProcess
        Case WM_RBUTTONUP
            hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
            PopupMenu mnuPopUp, , , , mnuMostrar
            If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
    End If

End Sub
Private Sub QuitarIconoSystray()
    On Error Resume Next


    Dim i As Integer
    Dim nid As NOTIFYICONDATA

    nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

    i = Shell_NotifyIconA(NIM_DELETE, nid)

End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    Call QuitarIconoSystray
    #If UsarQueSocket = 1 Then
        Call LimpiaWsApi(frmMain.hwnd)
    #Else
        Socket1.Cleanup
    #End If

    Call DescargaNpcsDat

    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers
        If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
    Next


    Call LogMain(" Server cerrado")
    End

End Sub

Private Sub mnuCerrar_Click()

    Call SaveGuildsNew

    If MsgBox("Si cierra el servidor puede provocar la perdida de datos." & vbCrLf & vbCrLf & "¿Desea hacerlo de todas maneras?", vbYesNo + vbExclamation, "Advertencia") = vbYes Then Call ApagarSistema

End Sub
Private Sub mnusalir_Click()

    Call mnuCerrar_Click

End Sub
Public Sub mnuMostrar_Click()
    On Error Resume Next

    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0

End Sub
Private Sub mnuServidor_Click()

    frmServidor.Visible = True

End Sub
Private Sub mnuSystray_Click()
    Dim i As Integer
    Dim S As String
    Dim nid As NOTIFYICONDATA

    S = "Servidor HarmonyAO"
    nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    i = Shell_NotifyIconA(NIM_ADD, nid)

    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = False

End Sub
Private Sub Socket1_Blocking(Status As Integer, Cancel As Integer)
    Cancel = True
End Sub
Private Sub Socket2_Connect(Index As Integer)

    Set UserList(Index).CommandsBuffer = New CColaArray

End Sub
Private Sub Socket2_Disconnect(Index As Integer)

    If UserList(Index).flags.UserLogged And _
       UserList(Index).Counters.Saliendo = False Then
        Call Cerrar_Usuario(Index)
    Else: Call CloseSocket(Index)
    End If

End Sub
Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)


    #If UsarQueSocket = 0 Then
        On Error GoTo ErrorHandler
        Dim LoopC As Integer
        Dim RD As String
        Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
        Dim CR As Integer
        Dim tChar As String
        Dim sChar As Integer
        Dim eChar As Integer
        Dim aux$
        Dim OrigCad As String
        Dim LenRD As Long

        Call Socket2(Index).Read(RD, DataLength)

        OrigCad = RD
        LenRD = Len(RD)

        If LenRD = 0 Then
            UserList(Index).AntiCuelgue = UserList(Index).AntiCuelgue + 1
            If UserList(Index).AntiCuelgue >= 150 Then
                UserList(Index).AntiCuelgue = 0
                Call LogError("!!!! Detectado bucle infinito de eventos socket2_read. cerrando indice " & Index)
                Socket2(Index).Disconnect
                Call CloseSocket(Index)
                Exit Sub
            End If
        Else
            UserList(Index).AntiCuelgue = 0
        End If

        If Len(UserList(Index).RDBuffer) > 0 Then
            RD = UserList(Index).RDBuffer & RD
            UserList(Index).RDBuffer = ""
        End If

        sChar = 1
        For LoopC = 1 To LenRD

            tChar = Mid$(RD, LoopC, 1)

            If tChar = ENDC Then
                CR = CR + 1
                eChar = LoopC - sChar
                rBuffer(CR) = Mid$(RD, sChar, eChar)
                sChar = LoopC + 1
            End If

        Next LoopC

        If Len(RD) - (sChar - 1) <> 0 Then UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))

        For LoopC = 1 To CR
            If ClientsCommandsQueue = 1 Then
                If Len(rBuffer(LoopC)) > 0 Then If Not UserList(Index).CommandsBuffer.Push(rBuffer(LoopC)) Then Call Cerrar_Usuario(Index)
            Else
                If UserList(Index).ConnID <> -1 Then
                    Call HandleData(Index, rBuffer(LoopC))
                Else
                    Exit Sub
                End If
            End If
        Next LoopC

        Exit Sub

ErrorHandler:
        Call LogError("Error en Socket read. " & Err.Description & " Numero paquetes:" & UserList(Index).NumeroPaquetesPorMiliSec & " . Rdata:" & OrigCad)
        Call CloseSocket(Index)
    #End If
End Sub


Private Sub retos2vs2_Timer()

    If OPCDuelos.OCUP Then
        OPCDuelos.Tiempo = OPCDuelos.Tiempo - 1
        If OPCDuelos.Tiempo <= 0 Then
            UserList(OPCDuelos.J1).Reto.Received_Request = False
            UserList(OPCDuelos.J1).Reto.Send_Request = False
            UserList(OPCDuelos.J1).Reto.Retando_2 = False

            UserList(OPCDuelos.J2).Reto.Received_Request = False
            UserList(OPCDuelos.J2).Reto.Send_Request = False
            UserList(OPCDuelos.J2).Reto.Retando_2 = False

            UserList(OPCDuelos.J3).Reto.Received_Request = False
            UserList(OPCDuelos.J3).Reto.Send_Request = False
            UserList(OPCDuelos.J3).Reto.Retando_2 = False

            UserList(OPCDuelos.J4).Reto.Received_Request = False
            UserList(OPCDuelos.J4).Reto.Send_Request = False
            UserList(OPCDuelos.J4).Reto.Retando_2 = False

            Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y)    'los mando a ulla
            Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y + 1)    'los mando a ulla

            frmMain.retos2vs2.Enabled = False    '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
            OPCDuelos.OCUP = False
            OPCDuelos.Tiempo = 0
        End If
    End If

End Sub

Private Sub Timer1_Timer()
    On Error GoTo errordm:
    hyunpt = hyunpt + 1
    Select Case hyunpt
    Case 53
        Call SendData(ToAll, 0, 0, "||DeathMatch> En 2 minutos y 30 segundos se realizará un deathmatch automatico." & FONTTYPE_GUILD)
    Case 58
        Call SendData(ToAll, 0, 0, "||DeathMatch> En 1 minuto y 15 segundos se realizará un deathmatch automatico." & FONTTYPE_GUILD)
    Case 62
        Call SendData(ToAll, 0, 0, "||DeathMatch> En 15 segundos se realizará un deathmatch automatico." & FONTTYPE_GUILD)
    Case 63
        Call death_comienza(RandomNumber(8, 14))
    Case 65
        If deathesp = True Then
            Call Deathauto_Cancela
            hyunpt = 2
        Else
            hyunpt = 2
        End If
    End Select
errordm:
End Sub

Private Sub timeravisar_Timer()
    Call SendData(ToAll, 0, 0, "||Usuarios, les recuerdo que hay nuevas actualizaciones!, asegurense de estar al dia, descargandose estas. Gracias." & FONTTYPE_GUILD)
End Sub

Private Sub TimerFatuo_Timer()
    On Error GoTo Error
    Dim i As Integer

    For i = 1 To LastNPC
        If Npclist(i).flags.NPCActive And Npclist(i).Numero = 89 Then Npclist(i).CanAttack = 1
    Next

    Exit Sub

Error:
    Call LogError("Error en TimerFatuo: " & Err.Description)
End Sub
Private Sub TimerMeditar_Timer()
    Dim i As Integer

    For i = 1 To LastUser
        If UserList(i).flags.Meditando Then Call TimerMedita(i)
    Next

End Sub
Sub TimerMedita(UserIndex As Integer)
    Dim Cant As Single

    If TiempoTranscurrido(UserList(UserIndex).Counters.tInicioMeditar) >= TIEMPO_INICIOMEDITAR Then
        Cant = UserList(UserIndex).Counters.ManaAcumulado + UserList(UserIndex).Stats.MaxMAN * (1 + UserList(UserIndex).Stats.UserSkills(Meditar) * 0.05) / 100
        If Cant <= 0.75 Then
            UserList(UserIndex).Counters.ManaAcumulado = Cant
            Exit Sub
        Else
            Cant = Round(Cant)
            UserList(UserIndex).Counters.ManaAcumulado = 0
        End If
        Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Cant, UserList(UserIndex).Stats.MaxMAN)
        Call SendData(ToIndex, UserIndex, 0, "MN" & Cant)
        Call SubirSkill(UserIndex, Meditar)
        If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
            Call SendData(ToIndex, UserIndex, 0, "D9")
            Call SendData(ToIndex, UserIndex, 0, "MEDOK")
            UserList(UserIndex).flags.Meditando = False
            UserList(UserIndex).Char.FX = 0
            UserList(UserIndex).Char.loops = 0
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
        End If
    End If

    Call SendUserMANA(UserIndex)

End Sub
Private Sub TimerTrabaja_Timer()
    Dim i As Integer
    On Error GoTo Error

    For i = 1 To LastUser
        If UserList(i).flags.Trabajando Then
            UserList(i).Counters.IdleCount = Timer

            Select Case UserList(i).flags.Trabajando
            Case Pesca
                Call DoPescar(i)

            Case Talar
                Call DoTalar(i, ObjData(MapData(UserList(i).POS.Map, UserList(i).TrabajoPos.X, UserList(i).TrabajoPos.Y).OBJInfo.OBJIndex).ArbolElfico = 1)

            Case Mineria
                Call DoMineria(i, ObjData(MapData(UserList(i).POS.Map, UserList(i).TrabajoPos.X, UserList(i).TrabajoPos.Y).OBJInfo.OBJIndex).MineralIndex)
            End Select
        End If
    Next
    Exit Sub
Error:
    Call LogError("Error en TimerTrabaja: " & Err.Description)

End Sub

Public Sub Tlimpiar_Timer()
    MinutosTLimpiar = MinutosTLimpiar + 1

    Call LimpiarItemsMundo
    MinutosTLimpiar = 1

End Sub

Private Sub torneos_Timer()
    AutoTorneo = AutoTorneo + 1
    Select Case AutoTorneo
    Case 84
        Call SendData(ToAll, 0, 0, "||Torneo> En 2 minutos y 30 segundos se realizará un torneo automatico." & FONTTYPE_GUILD)
    Case 89
        Call SendData(ToAll, 0, 0, "||Torneo> En 1 minuto y 15 segundos se realizará un torneo automatico." & FONTTYPE_GUILD)
    Case 93
        Call SendData(ToAll, 0, 0, "||Torneo> En 15 segundos se realizará un torneo automatico." & FONTTYPE_GUILD)
    Case 94
        Call torneos_auto(RandomNumber(3, 4))    ' con esto se hace un random si el torneo sera de 4 a 16 participantes.
    Case 96
        If Torneo_Esperando = True Then
            Call Torneoauto_Cancela
            AutoTorneo = 2
        Else
            AutoTorneo = 2
        End If
    End Select
End Sub

Private Sub UserTimer_Timer()
    On Error GoTo Error
    Static Andaban As Boolean, Contador As Single
    Dim Andan As Boolean, UI As Integer, i As Integer, XXN As Integer

    'matute
    If encuestas.activa = 1 Then
        encuestas.Tiempo = encuestas.Tiempo + 1
        If encuestas.Tiempo = 15 Then
            Call SendData(ToAll, 0, 0, "||Faltan 15 segundos para finalizar la encuesta." & FONTTYPE_TALK)
        ElseIf encuestas.Tiempo = 30 Then
            Call SendData(ToAll, 0, 0, "||RESULTADOS DE LA ENCUESTA:" & FONTTYPE_FENIX)
            Call SendData(ToAll, 0, 0, "||VOTOS POSITIVOS: " & encuestas.votosSI & " | VOTOS NEGATIVOS: " & encuestas.votosNP & FONTTYPE_TALK)
            If encuestas.votosNP < encuestas.votosSI Then
                Call SendData(ToAll, 0, 0, "||Opción ganadora: SI" & FONTTYPE_FENIX)
            ElseIf encuestas.votosSI < encuestas.votosNP Then
                Call SendData(ToAll, 0, 0, "||Opción ganadora: NO" & FONTTYPE_FENIX)
            ElseIf encuestas.votosNP = encuestas.votosSI Then
                Call SendData(ToAll, 0, 0, "||Opción ganadora: NINGUNA - EMPATE" & FONTTYPE_FENIX)
            End If
            encuestas.activa = 0
            encuestas.Tiempo = 0
            encuestas.votosNP = 0
            encuestas.votosSI = 0
            For XXN = 1 To LastUser
                If UserList(XXN).flags.votoencuesta = 1 Then UserList(XXN).flags.votoencuesta = 0
            Next XXN
        End If
        Exit Sub
    End If
    '//////'matute

    If CuentaRegresiva Then
        CuentaRegresiva = CuentaRegresiva - 1

        If CuentaRegresiva = 0 Then
            Call SendData(ToMap, 0, GMCuenta, "||YA!!!" & FONTTYPE_FIGHT)
            Me.Enabled = False
        Else
            Call SendData(ToMap, 0, GMCuenta, "||" & CuentaRegresiva & "..." & FONTTYPE_INFO)
        End If
    End If

    For i = 1 To LastUser
        If UserList(i).ConnID <> -1 Then DayStats.Segundos = DayStats.Segundos + 1
    Next

    If TiempoTranscurrido(Contador) >= 10 Then
        Contador = Timer
        Andan = EstadisticasWeb.EstadisticasAndando()
        If Not Andaban And Andan Then Call InicializaEstadisticas
        Andaban = Andan
    End If

    For UI = 1 To LastUser
        If UserList(UI).flags.UserLogged And UserList(UI).ConnID <> -1 Then
            Call TimerPiquete(UI)
            If UserList(UI).flags.Protegido > 1 Then Call TimerProtEntro(UI)
            If UserList(UI).flags.Encarcelado Then Call TimerCarcel(UI)
            If UserList(UI).flags.Muerto = 0 Then
                If UserList(UI).flags.Paralizado Then Call TimerParalisis(UI)
                If UserList(UI).flags.BonusFlecha Then Call TimerFlecha(UI)
                If UserList(UI).flags.Ceguera = 1 Then Call TimerCeguera(UI)
                If UserList(UI).flags.Envenenado = 1 Then Call TimerVeneno(UI)
                If UserList(UI).flags.Envenenado = 2 Then Call TimerVenenoDoble(UI)
                If UserList(UI).flags.Estupidez = 1 Then Call TimerEstupidez(UI)
                If UserList(UI).flags.AdminInvisible = 0 And UserList(UI).flags.Invisible = 1 And UserList(UI).flags.Oculto = 0 Then Call TimerInvisibilidad(UI)
                If UserList(UI).flags.Desnudo = 1 Then Call TimerFrio(UI)
                If UserList(UI).flags.TomoPocion Then Call TimerPocion(UI)
                If UserList(UI).flags.Transformado Then Call TimerTransformado(UI)
                If UserList(UI).NroMascotas Then Call TimerInvocacion(UI)
                If UserList(UI).flags.Oculto Then Call TimerOculto(UI)
                If UserList(UI).flags.Sacrificando Then Call TimerSacrificando(UI)

                Call TimerHyS(UI)
                Call TimerSanar(UI)
                Call TimerStamina(UI)
            End If
            If EnviarEstats Then
                Call SendUserStatsBox(UI)
                EnviarEstats = False
            End If
            Call TimerIdleCount(UI)
            If UserList(UI).Counters.Saliendo Then Call TimerSalir(UI)
        End If
    Next

    Exit Sub

Error:
    Call LogError("Error en UserTimer:" & Err.Description & " " & UI)

End Sub
Public Sub TimerOculto(UserIndex As Integer)
    Dim ClaseBuena As Boolean

    ClaseBuena = UserList(UserIndex).Clase = GUERRERO Or UserList(UserIndex).Clase = ARQUERO Or UserList(UserIndex).Clase = CAZADOR

    If RandomNumber(1, 10 + UserList(UserIndex).Stats.UserSkills(Ocultarse) / 4 + 15 * Buleano(ClaseBuena) + 25 * Buleano(ClaseBuena And Not UserList(UserIndex).Clase = GUERRERO And UserList(UserIndex).Invent.ArmourEqpObjIndex = 360)) <= 5 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.Invisible = 0
        Call SendData(ToMap, 0, UserList(UserIndex).POS.Map, ("V3" & UserList(UserIndex).Char.CharIndex & ",0"))
        Call SendData(ToIndex, UserIndex, 0, "V5")
    End If

End Sub
Public Sub TimerStamina(UserIndex As Integer)

    If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta And UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 And UserList(UserIndex).flags.Desnudo = 0 Then
        If (Not UserList(UserIndex).flags.Descansar And TiempoTranscurrido(UserList(UserIndex).Counters.STACounter) >= StaminaIntervaloSinDescansar) Or _
           (UserList(UserIndex).flags.Descansar And TiempoTranscurrido(UserList(UserIndex).Counters.STACounter) >= StaminaIntervaloDescansar) Then
            UserList(UserIndex).Counters.STACounter = Timer
            UserList(UserIndex).Stats.MinSta = Minimo(UserList(UserIndex).Stats.MinSta + CInt(RandomNumber(5, Porcentaje(UserList(UserIndex).Stats.MaxSta, 15))), UserList(UserIndex).Stats.MaxSta)
            If TiempoTranscurrido(UserList(UserIndex).Counters.CartelStamina) >= 10 Then
                UserList(UserIndex).Counters.CartelStamina = Timer
                Call SendData(ToIndex, UserIndex, 0, "MV")
            End If
            EnviarEstats = True
        End If
    End If

End Sub
Sub TimerTransformado(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Transformado) >= IntervaloInvisible Then
        Call DoTransformar(UserIndex)
    End If

End Sub
Sub TimerInvisibilidad(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Invisibilidad) >= IntervaloInvisible Then
        Call SendData(ToIndex, UserIndex, 0, "V6")
        Call QuitarInvisible(UserIndex)
    End If

End Sub
Sub TimerFlecha(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.BonusFlecha) >= 45 Then
        UserList(UserIndex).Counters.BonusFlecha = 0
        UserList(UserIndex).flags.BonusFlecha = False
        Call SendData(ToIndex, UserIndex, 0, "||Se acabó el efecto del Arco Encantado." & FONTTYPE_INFO)
    End If

End Sub
Sub TimerPiquete(UserIndex As Integer)

    If MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).trigger = 5 Then
        UserList(UserIndex).Counters.PiqueteC = UserList(UserIndex).Counters.PiqueteC + 1
        If UserList(UserIndex).Counters.PiqueteC Mod 5 = 0 Then Call SendData(ToIndex, UserIndex, 0, "9N")
        If UserList(UserIndex).Counters.PiqueteC >= 25 Then
            UserList(UserIndex).Counters.PiqueteC = 0
            Call Encarcelar(UserIndex, 3)
        End If
    Else: UserList(UserIndex).Counters.PiqueteC = 0
    End If

End Sub
Public Sub TimerProtEntro(UserIndex As Integer)
    On Error GoTo Error

    UserList(UserIndex).Counters.Protegido = UserList(UserIndex).Counters.Protegido - 1
    If UserList(UserIndex).Counters.Protegido <= 0 Then UserList(UserIndex).flags.Protegido = 0

    Exit Sub

Error:
    Call LogError("Error en TimerProtEntro" & " " & Err.Description)
End Sub
Sub TimerParalisis(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Paralisis) >= IntervaloParalizadoUsuario Then
        UserList(UserIndex).Counters.Paralisis = 0
        UserList(UserIndex).flags.Paralizado = 0
        Call SendData(ToIndex, UserIndex, 0, "P8")
    End If

End Sub
Sub TimerCeguera(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Ceguera) >= IntervaloParalizadoUsuario / 2 Then
        UserList(UserIndex).Counters.Ceguera = 0
        UserList(UserIndex).flags.Ceguera = 0
        Call SendData(ToIndex, UserIndex, 0, "NSEGUE")
    End If

End Sub
Sub TimerEstupidez(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Estupidez) >= IntervaloParalizadoUsuario Then
        UserList(UserIndex).Counters.Estupidez = 0
        UserList(UserIndex).flags.Estupidez = 0
        Call SendData(ToIndex, UserIndex, 0, "NESTUP")
    End If

End Sub
Sub TimerCarcel(UserIndex As Integer)

    Dim j As Byte

    If TiempoTranscurrido(UserList(UserIndex).Counters.Pena) >= UserList(UserIndex).Counters.TiempoPena Then
        UserList(UserIndex).Counters.TiempoPena = 0
        UserList(UserIndex).flags.Encarcelado = 0
        UserList(UserIndex).Counters.Pena = 0
        If UserList(UserIndex).POS.Map = Prision.Map Then
            Call WarpUserChar(UserIndex, Libertad.Map, Libertad.X, Libertad.Y, True)
            Call SendData(ToIndex, UserIndex, 0, "4P")
        End If
    End If

End Sub
Sub TimerVenenoDoble(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Veneno) >= 2 Then
        If TiempoTranscurrido(UserList(UserIndex).flags.EstasEnvenenado) >= 8 Then
            UserList(UserIndex).flags.Envenenado = 0
            UserList(UserIndex).flags.EstasEnvenenado = 0
            UserList(UserIndex).Counters.Veneno = 0
        Else
            Call SendData(ToIndex, UserIndex, 0, "1M")
            UserList(UserIndex).Counters.Veneno = Timer
            If Not UserList(UserIndex).flags.Quest Then
                UserList(UserIndex).Stats.MinHP = Maximo(0, UserList(UserIndex).Stats.MinHP - 25)
                If UserList(UserIndex).Stats.MinHP = 0 Then
                    Call UserDie(UserIndex)
                Else: EnviarEstats = True
                End If
            End If
        End If
    End If

End Sub
Sub UserSacrificado(UserIndex As Integer)
    Dim MiObj As Obj

    MiObj.OBJIndex = Gema
    MiObj.Amount = UserList(UserIndex).Stats.ELV ^ 2

    Call MakeObj(ToMap, UserIndex, UserList(UserIndex).POS.Map, MiObj, UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y)
    Call UserDie(UserIndex)

    UserList(UserList(UserIndex).flags.Sacrificador).flags.Sacrificado = 0
    Call SendData(ToIndex, UserList(UserIndex).flags.Sacrificador, 0, "||Sacrificaste a " & UserList(UserIndex).Name & " por " & MiObj.Amount & " partes de la piedra filosofal." & FONTTYPE_INFO)
    UserList(UserIndex).flags.Ban = 1
    Call CloseSocket(UserIndex)

End Sub
Sub TimerSacrificando(UserIndex As Integer)

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - 10
    UserList(UserList(UserIndex).flags.Sacrificador).Stats.MinMAN = Minimo(0, UserList(UserList(UserIndex).flags.Sacrificador).Stats.MinMAN - 50)
    Call SendUserMANA(UserList(UserIndex).flags.Sacrificador)

    If UserList(UserList(UserIndex).flags.Sacrificador).Stats.MinMAN = 0 Then Call CancelarSacrificio(UserIndex)
    If UserList(UserIndex).Stats.MinHP <= 0 Then Call UserSacrificado(UserIndex)

    EnviarEstats = True

End Sub
Sub TimerVeneno(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Veneno) >= IntervaloVeneno Then
        If TiempoTranscurrido(UserList(UserIndex).flags.EstasEnvenenado) >= IntervaloVeneno * 10 Then
            UserList(UserIndex).flags.Envenenado = 0
            UserList(UserIndex).flags.EstasEnvenenado = 0
            UserList(UserIndex).Counters.Veneno = 0
        Else
            Call SendData(ToIndex, UserIndex, 0, "1M")
            UserList(UserIndex).Counters.Veneno = Timer
            If Not UserList(UserIndex).flags.Quest Then
                UserList(UserIndex).Stats.MinHP = Maximo(0, UserList(UserIndex).Stats.MinHP - RandomNumber(1, 5))
                If UserList(UserIndex).Stats.MinHP = 0 Then
                    Call UserDie(UserIndex)
                Else: EnviarEstats = True
                End If
            End If
        End If
    End If

End Sub
Public Sub TimerFrio(UserIndex As Integer)

    If UserList(UserIndex).flags.Privilegios > 1 Then Exit Sub

    If TiempoTranscurrido(UserList(UserIndex).Counters.Frio) >= IntervaloFrio Then
        UserList(UserIndex).Counters.Frio = Timer
        If MapInfo(UserList(UserIndex).POS.Map).Terreno = Nieve Then
            If TiempoTranscurrido(UserList(UserIndex).Counters.CartelFrio) >= 5 Then
                UserList(UserIndex).Counters.CartelFrio = Timer
                Call SendData(ToIndex, UserIndex, 0, "1K")
            End If
            If Not UserList(UserIndex).flags.Quest Then
                UserList(UserIndex).Stats.MinHP = Maximo(0, UserList(UserIndex).Stats.MinHP - Porcentaje(UserList(UserIndex).Stats.MaxHP, 5))
                EnviarEstats = True
                If UserList(UserIndex).Stats.MinHP = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "1L")
                    Call UserDie(UserIndex)
                End If
            End If
        End If
        Call QuitarSta(UserIndex, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        If TiempoTranscurrido(UserList(UserIndex).Counters.CartelFrio) >= 10 Then
            UserList(UserIndex).Counters.CartelFrio = Timer
            Call SendData(ToIndex, UserIndex, 0, "FR")
        End If
        EnviarEstats = True
    End If

End Sub
Sub TimerPocion(UserIndex As Integer)
    If TiempoTranscurrido(UserList(UserIndex).flags.DuracionEfecto) >= 80 Then
        Call Parpa(UserIndex)
        If TiempoTranscurrido(UserList(UserIndex).flags.DuracionEfecto) >= 70 Then
            UserList(UserIndex).flags.DuracionEfecto = 0
            UserList(UserIndex).flags.TomoPocion = False
            UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
            UserList(UserIndex).Stats.UserAtributos(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza)
            Call UpdateFuerzaYAg(UserIndex)
        End If
    End If
End Sub
Public Sub TimerHyS(UserIndex As Integer)
    Dim EnviaInfo As Boolean

    If UserList(UserIndex).flags.Privilegios > 1 Or (UserList(UserIndex).Clase = TALADOR And UserList(UserIndex).Recompensas(1) = 2) Or UserList(UserIndex).flags.Quest Then Exit Sub

    If TiempoTranscurrido(UserList(UserIndex).Counters.AGUACounter) >= IntervaloSed Then
        If UserList(UserIndex).flags.Sed = 0 Then
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
            If UserList(UserIndex).Stats.MinAGU <= 0 Then
                UserList(UserIndex).Stats.MinAGU = 0
                UserList(UserIndex).flags.Sed = 1
            End If
            EnviaInfo = True
        End If
        UserList(UserIndex).Counters.AGUACounter = Timer
    End If

    If TiempoTranscurrido(UserList(UserIndex).Counters.COMCounter) >= IntervaloHambre Then
        If UserList(UserIndex).flags.Hambre = 0 Then
            UserList(UserIndex).Counters.COMCounter = Timer
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10
            If UserList(UserIndex).Stats.MinHam <= 0 Then
                UserList(UserIndex).Stats.MinHam = 0
                UserList(UserIndex).flags.Hambre = 1
            End If
            EnviaInfo = True
        End If
        UserList(UserIndex).Counters.COMCounter = Timer
    End If

    If EnviaInfo Then Call EnviarHambreYsed(UserIndex)

End Sub
Sub TimerSanar(UserIndex As Integer)

    If (UserList(UserIndex).flags.Descansar And TiempoTranscurrido(UserList(UserIndex).Counters.HPCounter) >= SanaIntervaloDescansar) Or _
       (Not UserList(UserIndex).flags.Descansar And TiempoTranscurrido(UserList(UserIndex).Counters.HPCounter) >= SanaIntervaloSinDescansar) Then
        If (Not Lloviendo Or Not Intemperie(UserIndex)) And UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP And UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
            If UserList(UserIndex).flags.Descansar Then
                UserList(UserIndex).Stats.MinHP = Minimo(UserList(UserIndex).Stats.MaxHP, UserList(UserIndex).Stats.MinHP + Porcentaje(UserList(UserIndex).Stats.MaxHP, 20))
                If UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MinHP And UserList(UserIndex).Stats.MaxSta = UserList(UserIndex).Stats.MinSta Then
                    Call SendData(ToIndex, UserIndex, 0, "DOK")
                    Call SendData(ToIndex, UserIndex, 0, "DN")
                    UserList(UserIndex).flags.Descansar = False
                End If
            Else
                UserList(UserIndex).Stats.MinHP = Minimo(UserList(UserIndex).Stats.MaxHP, UserList(UserIndex).Stats.MinHP + Porcentaje(UserList(UserIndex).Stats.MaxHP, 5))
            End If
            Call SendData(ToIndex, UserIndex, 0, "1N")
            EnviarEstats = True
        End If
        UserList(UserIndex).Counters.HPCounter = Timer
    End If

End Sub
Sub TimerInvocacion(UserIndex As Integer)
    Dim i As Integer
    Dim NpcIndex As Integer

    If UserList(UserIndex).flags.Privilegios > 0 Or UserList(UserIndex).flags.Quest Then Exit Sub

    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(i) Then
            NpcIndex = UserList(UserIndex).MascotasIndex(i)
            If Npclist(NpcIndex).Contadores.TiempoExistencia > 0 And TiempoTranscurrido(Npclist(NpcIndex).Contadores.TiempoExistencia) >= IntervaloInvocacion + 10 * Buleano(Npclist(NpcIndex).Numero = 92) Then Call MuereNpc(NpcIndex, 0)
        End If
    Next

End Sub
Public Sub TimerIdleCount(UserIndex As Integer)

    If UserList(UserIndex).flags.Privilegios = 0 And UserList(UserIndex).flags.Trabajando = 0 And TiempoTranscurrido(UserList(UserIndex).Counters.IdleCount) >= IntervaloParaConexion And Not UserList(UserIndex).Counters.Saliendo Then
        Call SendData(ToIndex, UserIndex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
        Call SendData(ToIndex, UserIndex, 0, "FINOK")
        Call CloseSocket(UserIndex)
    End If

End Sub
Sub TimerSalir(UserIndex As Integer)

    If TiempoTranscurrido(UserList(UserIndex).Counters.Salir) >= IntervaloCerrarConexion Then
        Call SendData(ToIndex, UserIndex, 0, "FINOK")
        Call CloseSocket(UserIndex)
    End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
