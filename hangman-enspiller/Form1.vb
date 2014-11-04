Public Class Form1

    Private spillmatrise() As Char 'Ordet som brukes i denne runden
    Private ordliste(6) As String 'Alle ordene som kan brukes i spillet
    Private bokstavsky(28) As String 'Bokstavene i alfabetet

    Private spillordlengde As Integer 'Lengden på ordet som brukes i denne runden
    'Private bokstaverFunnet As Integer 'Antall bokstaver som er funnet av i spillordet denne runden
    Private liv As Integer 'Antallet liv/streker på tegningen spilleren har til rådighet hver runde
    Private bokstaverIgjen 'Antallet bokstaver som gjenstår å finne denne runden

    'Funksjon som kjøres hver gang en bokstav i bokstavskyen klikkes på
    Private Function bokstavsjekk(ByVal valgtBokstav As String, valgtLabel As Integer)
        'Fjerner bokstaven fra bokstavsky
        Me.Controls("label" & valgtLabel).Visible = False
        'definerer variabelen antallForekomster
        Dim antallForekomster As Integer = 0
        'sjekker om noen av bokstaven finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = valgtBokstav Then
                Me.Controls("label" & i).Text = valgtBokstav 'xxx endre til Caps lock her
                antallForekomster = antallForekomster + 1
            End If
        Next
        'oppdaterer antall bokstaver som gjenstår før spillet er vunnet
        bokstaverIgjen = bokstaverIgjen - antallForekomster

        'sjekker om spilleren har vunnet
        If bokstaverIgjen = 0 Then
            'fjerner bokstavskyen
            For i = 11 To 39 'skal skjule label 11 til og med label 39
                Me.Controls("label" & i).Visible = False
            Next
            MsgBox("Du har vunnet!") 'viser vinnerskjermbilde

        End If

        'sjekker om spilleren skal miste liv
        If antallForekomster = 0 Then
            liv = liv - 1
        End If

        'sjekker om spilleren er død/hengt
        If liv = 0 Then
            MsgBox("Du har tapt og har blitt hengt..")
        End If


    End Function



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
        Dim tilfeldig As Integer = CInt(Int((ordlistelengde * Rnd()))) 'xxx her er det noe feil som gjør at "tilfeldig" kan bli lik null.
        MsgBox("tilfeldig tall: " & tilfeldig)

        'Henter inn ord basert på det tilfeldige tallet
        Dim spillord As String = ordliste(tilfeldig - 1)

        spillordlengde = spillord.Length
        MsgBox("Spillord: " & spillord)

        spillmatrise = spillord.ToCharArray
        MsgBox("spillmatrise: " & spillmatrise(0))

        'legger inn underscores/"nullstiller" i alle labels
        For i = 1 To 10
            Me.Controls("label" & i).Text = "_"
        Next

        'viser underscores basert på spillordlengde
        For i = 1 To spillordlengde
            Me.Controls("label" & i).Visible = True
        Next

        'setter variabelen "liv" til 10 (som tilsvarer maksimalt antall streker på tegningen)
        liv = 1 'har satt denne til et annet tall enn 10 for å kunne teste
        MsgBox("liv = " & liv)

        'setter variabelen bokstaverIgjen lik lengden på spillordet
        bokstaverIgjen = spillordlengde
        MsgBox("Antall bokstaver igjen å finne/bokstaverIgjen " & bokstaverIgjen)


    End Sub


    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        bokstavsjekk("a", 11)
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        bokstavsjekk("b", 12)
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        bokstavsjekk("c", 13)
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        bokstavsjekk("d", 14)
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        bokstavsjekk("e", 15)
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click

    End Sub
End Class
