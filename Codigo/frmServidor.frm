VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servidor"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Reiniciar 
      Caption         =   "Reiniciar servidor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Actualizar npcs.dat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Reload Server.ini"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   4095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Update MOTD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   4095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Unban All IPs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cargar BackUp del mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Guardar todos los personajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4440
      Width           =   4095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Unban All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   4095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Debug listening socket"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Debug Npcs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   4095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Stats de los slots"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Trafico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Reload Lista Nombres Prohibidos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Actualizar hechizos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Configurar intervalos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ReSpawn Guardias en posiciones originales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar objetos.dat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hacer un Backup del mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   5160
      Width           =   945
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reiniciar sockets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      Height          =   3855
      Left            =   120
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   120
      Top             =   4080
      Width           =   4335
   End
End
Attribute VB_Name = "frmServidor"
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
Private Sub Command1_Click()
    Call LoadOBJData
End Sub

Private Sub Command10_Click()
    frmTrafic.Show
End Sub

Private Sub Command11_Click()
    frmConID.Show
End Sub

Private Sub Command12_Click()
    frmDebugNpc.Show
End Sub

Private Sub Command14_Click()
    Call LoadMotd
End Sub

Private Sub Command15_Click()
    On Error Resume Next

    Dim Fn As String
    Dim cad As String
    Dim N As Integer, k As Integer

    Fn = App.Path & "\logs\GenteBanned.log"

    If FileExist(Fn, vbNormal) Then
        N = FreeFile
        Open Fn For Input Shared As #N
        Do While Not EOF(N)
            k = k + 1
            Input #N, cad
            Call ChangeBan(cad, 0)

        Loop
        Close #N
        MsgBox "Se han habilitado " & k & " personajes."
        Kill Fn
    End If




End Sub

Private Sub Command16_Click()
    Call LoadSini
End Sub

Private Sub Command17_Click()
    Call DescargaNpcsDat
    Call CargaNpcsDat

End Sub
Private Sub Command18_Click()

    Me.MousePointer = 11
    Call GuardarUsuarios
    Me.MousePointer = 0
    MsgBox "Grabado de personajes OK!"

End Sub

Private Sub Command19_Click()
    Dim i As Long, N As Long

    N = BanIps.Count
    For i = 1 To BanIps.Count
        BanIps.Remove 1
    Next

    MsgBox "Se han habilitado " & N & " ipes"

End Sub

Private Sub Command2_Click()
    frmServidor.Visible = False
End Sub

Private Sub Command4_Click()

    On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Show
    Call DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
    Exit Sub

eh:
    Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Command5_Click()


    On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.caption = "Reiniciando."

    FrmStat.Show

    If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"



    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    LastUser = 0
    NumUsers = 0
    NumNoGMs = 0

    ReDim Npclist(1 To MAXNPCS) As Npc
    ReDim CharList(1 To MAXCHARS) As Integer

    Call LoadSini
    Call CargarBackUp
    Call LoadOBJData



    If frmMain.Visible Then frmMain.txStatus.caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
    Call ReSpawnOrigPosNpcs
End Sub

Private Sub Command7_Click()
    FrmInterv.Show
End Sub

Private Sub Command8_Click()
    Call CargarHechizos

End Sub

Private Sub Command9_Click()
    Call CargarForbidenWords
End Sub

Private Sub Form_Deactivate()
    frmServidor.Visible = False
End Sub

Private Sub Reiniciar_Click()
    ShellExecute frmMain.hwnd, "open", App.Path & "/Server.exe", "", "", 1
    Call ApagarSistema
End Sub
