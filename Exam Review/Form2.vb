Option Strict Off
Option Explicit On

Public Class Form2
    Dim a As Integer = 0
    Dim b As Integer = 1
    Dim c As Integer = 2

    Private Sub testHash()
        Dim hshA As New Hashtable
        Dim lisA As New List(Of Hashtable)
        hshA.Add("FName", "FA1")
        hshA.Add("LName", "FA2")
        hshA.Add("CName", "FA3")
        lisA.Add(hshA)
        hshA = New Hashtable
        hshA.Add("FName", "FB1")
        hshA.Add("LName", "FB2")
        hshA.Add("CName", "FB3")
        lisA.Add(hshA)
        Dim FinalAnswer As String = CType(lisA.Item(1)("FName"), String)
        Debug.WriteLine(FinalAnswer)
        Debug.WriteLine("---------------")
        For Each hshB In lisA
            For Each Val As DictionaryEntry In hshB
                Debug.WriteLine(Val.Key)
                Debug.WriteLine(Val.Value)
            Next
        Next
    End Sub


    Private Function d(a As Integer) As Integer
        a = 100
        Return a
    End Function

    Private Function e(ByVal b As Integer) As Integer
        b = 200
        Return b
    End Function

    Private Function f(ByRef c As Integer) As Integer
        c = 300
        Return c
    End Function

    Private Sub Button1_Click(sender As Object, v As EventArgs) Handles Button1.Click
        testHash()
        Debug.WriteLine(d(a)) ' 100
        Debug.WriteLine(e(b)) ' 200
        Debug.WriteLine(f(c)) ' 300
        Debug.WriteLine(a) ' 0
        Debug.WriteLine(b) ' 1
        Debug.WriteLine(c) ' 300
    End Sub
End Class