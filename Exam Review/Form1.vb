Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RichTextBox2.Text = ""
        Dim GetValue As String = RichTextBox1.Text
        Dim Results As List(Of Hashtable) = controlDB(ConnectString, GetValue, True)
        Dim HeaderIsUse As Boolean = False
        Dim CollectThisString As String = String.Empty
        For Each Result In Results
            CollectThisString = String.Empty
            For Each Element As DictionaryEntry In Result
                If HeaderIsUse = False Then
                    RichTextBox2.AppendText(CType(Element.Key, String) + " | ")
                End If
                CollectThisString = CollectThisString + CType(Element.Value, String) + " | "
            Next
            If HeaderIsUse = False Then
                HeaderIsUse = True
                RichTextBox2.AppendText(br + "-----------------------------------" + br)
            End If
            RichTextBox2.AppendText(CollectThisString)
            RichTextBox2.AppendText(br)
        Next
    End Sub
End Class
