Attribute VB_Name = "Recpass"
Option Explicit

Dim oMail As clsCDOmail

Public Function EnviarCorreo(ByVal UserNick As String, ByVal UserMail As String) As Boolean

    Set oMail = New clsCDOmail

    With oMail
        .Servidor = "smtp.gmail.com"
        .Puerto = 465
        .UseAuntentificacion = True
        .SSL = True
        .Usuario = "USUARIO GMAIL (SIN @gmail.com)"
        .PassWord = "CONTRASEÑA DE GMAIL"
        .Asunto = "Datos del personaje " & UserNick
        .De = "FlamiusAO Staff"
        .Para = UserMail
        .Mensaje = "Estimado usuario, le informamos que el personaje " & UserNick & " tiene nueva contraseña. Ésta nueva password fué solicitada desde el juego para su recuperación. La nueva contraseña es " & ObtenerPassword(UserNick) & ".Rogamos memoizarla o cambiarla. Atte. HarmonyAO Staff."
        If .Enviar_Backup Then
            EnviarCorreo = True
        Else
            EnviarCorreo = False
        End If
    End With

    Set oMail = Nothing

End Function
