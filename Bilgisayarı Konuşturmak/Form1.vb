Public Class Form1
    Dim Jarvis = CreateObject("Sapi.spvoice")

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Jarvis.Speak(TextBox1.Text)
    End Sub
End Class
