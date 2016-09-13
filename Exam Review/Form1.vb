Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Clean Result Windows
        RichTextBox2.Text = ""
        ' Set Variable
        Dim GetValue As String = RichTextBox1.Text
        Dim RegexTest As New List(Of Regex)
        RegexTest.Add(New Regex("^[SELECT]+"))
        RegexTest.Add(New Regex("^[INSERT]+"))
        RegexTest.Add(New Regex("^[UPDATE]+"))
        RegexTest.Add(New Regex("^[DELETE]+"))
        ' If it is a select

        If RegexTest.Item(0).IsMatch(GetValue) = True Then
            Dim Results As List(Of Hashtable) = controlDB(connectString, GetValue, True)
            Dim CountResults As Integer = Results.Count()
            Dim HeaderIsUse As Boolean = False
            Dim CollectThisString As String = String.Empty
            If CountResults > 0 Then
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
            Else
                RichTextBox2.AppendText("There is no record in this query!")
            End If
        ElseIf RegexTest.Item(1).IsMatch(GetValue) = True Or RegexTest.Item(2).IsMatch(GetValue) = True Or RegexTest.Item(3).IsMatch(GetValue) = True Then
                Dim Results As Boolean = controlDB(connectString, GetValue)
                If Results = True Then
                    RichTextBox2.Text = Results.ToString
                Else
                    RichTextBox2.Text = Results.ToString
                End If
            Else
                RichTextBox2.Text = "Error!"
        End If

    End Sub
End Class
