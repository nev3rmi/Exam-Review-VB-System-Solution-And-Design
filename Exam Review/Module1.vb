Module Module1
    Public a As Integer = 100
    Public Function b() As Integer
        MsgBox(a)
        Return a
    End Function

    Friend Function c() As Integer
        Return a
    End Function
End Module
