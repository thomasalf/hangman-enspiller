Public Class Form1

    Private spillmatrise(9) As String 'Ordet som brukes i denne runden
    Private ordliste(10) As String 'Alle ordene som kan brukes i spillet





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Laster inn ord i ordlistematrisen
        ordliste(0) = "abbedissen"
        ordliste(1) = "abdominalt"
        ordliste(2) = "horekunder"
        ordliste(3) = "hormonbruk"
        ordliste(4) = "horisonten"
        ordliste(5) = "test"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Sjekker lengden på ordlisten og laster verdien inn i variabel
        Dim ordlistelengde As Integer
        ordlistelengde = ordliste.Length
        MsgBox("ordlistelengde: " & ordlistelengde)

        'Velger tilfeldig tall mellom 1 og ordlistelengde og laster verdien inn i variabel
        Dim tilfeldig As Integer = CInt(Int((ordlistelengde * Rnd()) + 1))

        'Henter inn ord

    End Sub
End Class
