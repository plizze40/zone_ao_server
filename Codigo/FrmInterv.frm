VERSION 5.00
Begin VB.Form FrmInterv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervalos ~ Servidor FlamiusAO ~"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "FrmInterv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Intervalos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame11 
      Caption         =   "NPCs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   2640
      TabIndex        =   51
      Top             =   2280
      Width           =   1665
      Begin VB.Frame Frame4 
         Caption         =   "A.I.:"
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
         Height          =   1650
         Left            =   150
         TabIndex        =   52
         Top             =   270
         Width           =   1365
         Begin VB.TextBox txtAI 
            Height          =   285
            Left            =   150
            TabIndex        =   15
            Text            =   "0"
            Top             =   1200
            Width           =   1050
         End
         Begin VB.TextBox txtNPCPuedeAtacar 
            Height          =   285
            Left            =   135
            TabIndex        =   14
            Text            =   "0"
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "AI"
            Height          =   195
            Left            =   165
            TabIndex        =   54
            Top             =   960
            Width           =   150
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Puede atacar"
            Height          =   195
            Left            =   150
            TabIndex        =   53
            Top             =   375
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Clima / Ambiente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   4440
      TabIndex        =   45
      Top             =   2280
      Width           =   2505
      Begin VB.Frame Frame7 
         Caption         =   "Frio y Fx Ambientales:"
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
         Height          =   1650
         Left            =   165
         TabIndex        =   46
         Top             =   300
         Width           =   2220
         Begin VB.TextBox txtCmdExec 
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Text            =   "0"
            Top             =   1200
            Width           =   915
         End
         Begin VB.TextBox txtIntervaloPerdidaStaminaLluvia 
            Height          =   300
            Left            =   1080
            TabIndex        =   17
            Text            =   "0"
            Top             =   630
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloWAVFX 
            Height          =   300
            Left            =   120
            TabIndex        =   16
            Text            =   "0"
            Top             =   630
            Width           =   810
         End
         Begin VB.TextBox txtIntervaloFrio 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Text            =   "0"
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "TimerExec"
            Height          =   195
            Left            =   1080
            TabIndex        =   50
            Top             =   960
            Width           =   750
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Sta. Lluvia"
            Height          =   195
            Left            =   1080
            TabIndex        =   49
            Top             =   375
            Width           =   750
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "FxS"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   375
            Width           =   270
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Frio"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   6830
      Begin VB.Frame Frame9 
         Caption         =   "Otros:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   120
         TabIndex        =   37
         Top             =   330
         Width           =   1290
         Begin VB.TextBox txtIntervaloParaConexion 
            Height          =   300
            Left            =   120
            TabIndex        =   0
            Text            =   "0"
            Top             =   495
            Width           =   930
         End
         Begin VB.TextBox txtTrabajo 
            Height          =   300
            Left            =   120
            TabIndex        =   1
            Text            =   "0"
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "IntervaloCon"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Trabajo"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   900
            Width           =   540
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Combate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   1440
         TabIndex        =   34
         Top             =   330
         Width           =   1290
         Begin VB.TextBox txtPuedeAtacar 
            Height          =   300
            Left            =   135
            TabIndex        =   3
            Text            =   "0"
            Top             =   1140
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloLanzaHechizo 
            Height          =   300
            Left            =   150
            TabIndex        =   2
            Text            =   "0"
            Top             =   525
            Width           =   930
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Puede Atacar"
            Height          =   195
            Left            =   135
            TabIndex        =   36
            Top             =   900
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Lanza Spell"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   285
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ham. / Sed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   5400
         TabIndex        =   31
         Top             =   330
         Width           =   1290
         Begin VB.TextBox txtIntervaloHambre 
            Height          =   285
            Left            =   150
            TabIndex        =   8
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloSed 
            Height          =   285
            Left            =   150
            TabIndex        =   9
            Text            =   "0"
            Top             =   1140
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hambre"
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sed"
            Height          =   195
            Left            =   165
            TabIndex        =   32
            Top             =   900
            Width           =   285
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sanar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   4080
         TabIndex        =   28
         Top             =   330
         Width           =   1290
         Begin VB.TextBox txtSanaIntervaloDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   6
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtSanaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   7
            Text            =   "0"
            Top             =   1140
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   255
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   29
            Top             =   900
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stamina:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   2760
         TabIndex        =   25
         Top             =   330
         Width           =   1290
         Begin VB.TextBox txtStaminaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   5
            Text            =   "0"
            Top             =   1140
            Width           =   1050
         End
         Begin VB.TextBox txtStaminaIntervaloDescansar 
            Height          =   285
            Left            =   165
            TabIndex        =   4
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   27
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   255
            Width           =   990
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Magia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   2415
      Begin VB.Frame Frame10 
         Caption         =   "Duracion de Spells:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1650
         Left            =   135
         TabIndex        =   40
         Top             =   270
         Width           =   2120
         Begin VB.TextBox txtInvocacion 
            Height          =   300
            Left            =   1080
            TabIndex        =   13
            Text            =   "0"
            Top             =   1170
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloInvisible 
            Height          =   300
            Left            =   1080
            TabIndex        =   11
            Text            =   "0"
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloParalizado 
            Height          =   300
            Left            =   120
            TabIndex        =   12
            Text            =   "0"
            Top             =   1170
            Width           =   795
         End
         Begin VB.TextBox txtIntervaloVeneno 
            Height          =   300
            Left            =   120
            TabIndex        =   10
            Text            =   "0"
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Invocacion"
            Height          =   195
            Left            =   1080
            TabIndex        =   44
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Invisible"
            Height          =   195
            Left            =   1080
            TabIndex        =   43
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Paralizado"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Veneno"
            Height          =   180
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton ok 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "FrmInterv"
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
Public Sub AplicarIntervalos()


    SanaIntervaloSinDescansar = val(txtSanaIntervaloSinDescansar.Text)
    StaminaIntervaloSinDescansar = val(txtStaminaIntervaloSinDescansar.Text)
    SanaIntervaloDescansar = val(txtSanaIntervaloDescansar.Text)
    StaminaIntervaloDescansar = val(txtStaminaIntervaloDescansar.Text)
    IntervaloSed = val(txtIntervaloSed.Text)
    IntervaloHambre = val(txtIntervaloHambre.Text)
    IntervaloVeneno = val(txtIntervaloVeneno.Text)
    IntervaloParalizado = val(txtIntervaloParalizado.Text)
    IntervaloInvisible = val(txtIntervaloInvisible.Text)
    IntervaloFrio = val(txtIntervaloFrio.Text)
    IntervaloWavFx = val(txtIntervaloWAVFX.Text)
    IntervaloInvocacion = val(txtInvocacion.Text)
    IntervaloParaConexion = val(txtIntervaloParaConexion.Text)



    IntervaloUserPuedeCastear = val(txtIntervaloLanzaHechizo.Text)
    IntervaloUserPuedeAtacar = val(txtPuedeAtacar.Text)


End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Call AplicarIntervalos

End Sub

Private Sub Command2_Click()

    On Error GoTo Err


    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar", str(SanaIntervaloSinDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar", str(StaminaIntervaloSinDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar", str(SanaIntervaloDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar", str(StaminaIntervaloDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed", str(IntervaloSed))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre", str(IntervaloHambre))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno", str(IntervaloVeneno))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado", str(IntervaloParalizado))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible", str(IntervaloInvisible))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio", str(IntervaloFrio))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX", str(IntervaloWavFx))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion", str(IntervaloInvocacion))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion", str(IntervaloParaConexion))



    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", str(IntervaloUserPuedeCastear))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", str(IntervaloUserPuedeAtacar))

    MsgBox "Los intervalos se han guardado sin problemas.", vbInformation, "Servidor HarmonyAO"

    Exit Sub
Err:
    MsgBox "Error al intentar grabar los intervalos"
End Sub

Private Sub ok_Click()
    Me.Visible = False
End Sub

