Attribute VB_Name = "mdlRetos"
Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = TIndex
    UserList(TIndex).flags.EstaDueleando = True
    Call WarpUserChar(TIndex, 16, 37, 44)
    UserList(TIndex).flags.Oponente = UserIndex
    Call WarpUserChar(UserIndex, 16, 58, 56)
    Call SendData(ToAll, 0, 0, "||Ring 1> " & UserList(TIndex).Name & " y " & UserList(UserIndex).Name & " van a competir en un reto." & FONTTYPE_RETOS)
End Sub
Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    Call WarpUserChar(UserIndex, 17, 50, 50)
    Call WarpUserChar(TIndex, 17, 51, 51)
    UserList(TIndex).flags.EsperandoDuelo = False
    UserList(TIndex).flags.Oponente = 0
    UserList(TIndex).flags.EstaDueleando = False
End Sub
Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Ring 1> " & UserList(Ganador).Name & " venció a " & UserList(Perdedor).Name & " en un reto." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
End Sub
Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Ring 1> El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).Name & "." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
End Sub
