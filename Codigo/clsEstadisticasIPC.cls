VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadisticasIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Option Explicit

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_USER = &H400

Private sMensaje As Long, hVentana As Long, hVentanaMia As Long
Attribute hVentana.VB_VarUserMemId = 1073938432
Attribute hVentanaMia.VB_VarUserMemId = 1073938432

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDnext = 2


Public Enum EstaNotificaciones
    CANTIDAD_ONLINE = 1
    RECORD_USUARIOS = 2
    UPTIME_SERVER = 3
    CANTIDAD_MAPAS = 4
    EVENTO_NUEVO_CLAN = 5

    HANDLE_WND_SERVER = 100
End Enum



Private Function BuscaVentana(Wnd As Long, str As String) As Long
    Dim W As Long
    Dim L As Long
    Dim T As String


    W = GetWindow(Wnd, GW_HWNDFIRST)

    While W <> 0
        L = GetWindowTextLength(W)

        If L Then
            T = Space$(L + 1)
            L = GetWindowText(W, T, L + 1)

            If Left$(T, Len(str)) = str Then
                BuscaVentana = W
                Exit Function
            End If
        End If

        W = GetWindow(W, GW_HWNDnext)
    Wend

    BuscaVentana = 0

End Function

Public Function Informar(ByVal QueCosa As EstaNotificaciones, ByVal Parametro As Long) As Long
    Call BuscaWndEstadisticas
    Informar = SendMessageLong(hVentana, sMensaje, QueCosa, Parametro)

End Function

Public Function EstadisticasAndando() As Boolean
    Dim Ret As Long

    Call BuscaWndEstadisticas

    EstadisticasAndando = (hVentana <> 0)

End Function

Public Sub Inicializa(ByVal hwnd As Long)
    hVentanaMia = hwnd
    sMensaje = RegisterWindowMessage("EstadisticasAO")

End Sub

Private Sub BuscaWndEstadisticas()
    hVentana = BuscaVentana(hVentanaMia, "Servidor de estadisticas AO")

End Sub
