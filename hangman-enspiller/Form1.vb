Public Class Form1

    Private spillmatrise() As Char 'Ordet som brukes i denne runden
    Private ordliste(6) As String 'Alle ordene som kan brukes i spillet
    Private bokstavsky(28) As String 'Bokstavene i alfabetet

    Private spillordlengde As Integer 'Lengden på ordet som brukes i denne runden
    Private bokstaverFunnet As Integer 'Antall bokstaver som er funnet av i spillordet denne runden





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Laster inn ord i ordlistematrisen
        ordliste(0) = "a"
        ordliste(1) = "ab"
        ordliste(2) = "abc"
        ordliste(3) = "abcd"
        ordliste(4) = "abcde"
        ordliste(5) = "abcdea"
        ordliste(6) = "abcdeab"

        'laster inn bokstavene i alfabetet
        bokstavsky(0) = "a"
        bokstavsky(1) = "b"
        bokstavsky(2) = "c"
        bokstavsky(3) = "d"
        bokstavsky(4) = "e"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Viser bokstaver i bokstavsky
        For i = 11 To 39 'skal vise label 11 til og med label 39
            Me.Controls("label" & i).Visible = True
        Next

        'Sjekker lengden på ordlisten og laster verdien inn i variabel
        Dim ordlistelengde As Integer
        ordlistelengde = ordliste.Length
        MsgBox("ordlistelengde: " & ordlistelengde)

        'Velger tilfeldig tall mellom 0 og ordlistelengde og laster verdien inn i variabel
        Randomize()
        Dim tilfeldig As Integer = CInt(Int((ordlistelengde * Rnd())))
        MsgBox("tilfeldig tall: " & tilfeldig)

        'Henter inn ord basert på det tilfeldige tallet
        Dim spillord As String = ordliste(tilfeldig - 1)

        spillordlengde = spillord.Length
        MsgBox("Spillord: " & spillord)

        spillmatrise = spillord.ToCharArray
        MsgBox("spillmatrise: " & spillmatrise(0))

        'viser underscores basert på spillordlengde
        For i = 1 To spillordlengde
            Me.Controls("label" & i).Visible = True
        Next


    End Sub

    
    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        'Fjerner bokstaven A fra bokstavsky
        Label11.Visible = False
        'sjekker om noen av bokstaven A finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = "a" Then
                Me.Controls("label" & i).Text = "A"
            End If
        Next
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        'sjekker om noen av bokstaven B finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = "b" Then
                Me.Controls("label" & i).Text = "B"
            End If
        Next
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        'sjekker om noen av bokstaven C finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = "c" Then
                Me.Controls("label" & i).Text = "C"
            End If
        Next
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        'sjekker om noen av bokstaven D finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = "d" Then
                Me.Controls("label" & i).Text = "D"
            End If
        Next
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        'sjekker om noen av bokstaven D finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = "e" Then
                Me.Controls("label" & i).Text = "E"
            End If
        Next
    End Sub



End Class
