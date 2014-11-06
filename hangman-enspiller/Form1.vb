Public Class Form1

    Private spillmatrise() As Char 'Ordet som brukes i denne runden
    Private ordliste(10) As String 'Alle ordene som kan brukes i spillet
    Private bokstavsky(28) As String 'Bokstavene i alfabetet

    Private spillordlengde As Integer 'Lengden på ordet som brukes i denne runden
    'Private bokstaverFunnet As Integer 'Antall bokstaver som er funnet av i spillordet denne runden
    Private liv As Integer 'Antallet liv/streker på tegningen spilleren har til rådighet hver runde
    Private bokstaverIgjen As Integer 'Antallet bokstaver som gjenstår å finne denne runden

    'Funksjon som kjøres hver gang en bokstav i bokstavskyen klikkes på
    Private Function bokstavsjekk(ByVal valgtBokstav As String, valgtLabel As Integer)
        'Fjerner bokstaven fra bokstavsky
        Me.Controls("label" & valgtLabel).Visible = False
        'definerer variabelen antallForekomster
        Dim antallForekomster As Integer = 0
        'sjekker om noen av bokstaven finnes i spillmatrisen 
        For i = 1 To (spillordlengde)
            If spillmatrise(i - 1) = valgtBokstav Then
                Me.Controls("label" & i).Text = valgtBokstav
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
        Select Case liv
            Case 8
                PictureBox1.Visible = True
            Case 7
                PictureBox1.Visible = False
                PictureBox2.Visible = True
            Case 6
                PictureBox2.Visible = False
                PictureBox3.Visible = True
            Case 5
                PictureBox3.Visible = False
                PictureBox4.Visible = True
            Case 4
                PictureBox4.Visible = False
                PictureBox5.Visible = True
            Case 3
                PictureBox5.Visible = False
                PictureBox6.Visible = True
            Case 2
                PictureBox6.Visible = False
                PictureBox7.Visible = True
            Case 1
                PictureBox7.Visible = False
                PictureBox8.Visible = True
            Case 0
                PictureBox8.Visible = False
                PictureBox9.Visible = True
                MsgBox("Du har tapt og har blitt hengt..")
        End Select

        'gammel "sjekk om spilleren er død"-kode
        'If liv = 0 Then
        'PictureBox1.Visible = False
        'PictureBox9.Visible = True
        'MsgBox("Du har tapt og har blitt hengt..")
        'End If


    End Function


    'Laster inn ord i matrise og bokstaver i bokstavsky. Viser tomt bilde.
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Laster inn ord i ordlistematrisen
        ordliste(0) = "A"
        ordliste(1) = "AB"
        ordliste(2) = "ABC"
        ordliste(3) = "ABCD"
        ordliste(4) = "ABCDE"
        ordliste(5) = "ABCDEF"
        ordliste(6) = "ABCDEFA"
        ordliste(7) = "ABCCDGAD"
        ordliste(8) = "ABHAIJKLM"
        ordliste(9) = "AABBHHGGAA"
        ordliste(10) = "BBCCDDKKLLM"

        'laster inn bokstavene i alfabetet
        bokstavsky(0) = "A"
        bokstavsky(1) = "B"
        bokstavsky(2) = "C"
        bokstavsky(3) = "D"
        bokstavsky(4) = "E"

        'viser tomt bilde / bakgrunnsbilde
        For i = 2 To 9
            Me.Controls("Picturebox" & i).Visible = False
        Next
        PictureBox1.Visible = True
    End Sub

    'Starter spillet
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
        Dim tilfeldig As Integer = CInt(Int((ordlistelengde * Rnd()) + 1)) 'xxx Tror dette er ok nå. Her var det noe feil som gjorder at "tilfeldig" kunne bli lik null.
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

        'setter variabelen "liv" til 8 (som tilsvarer maksimalt antall streker på tegningen)
        liv = 8 'xxx har muligens satt denne til et annet tall enn 8 for å kunne teste
        MsgBox("liv = " & liv)

        'setter variabelen bokstaverIgjen lik lengden på spillordet
        bokstaverIgjen = spillordlengde
        MsgBox("Antall bokstaver igjen å finne/bokstaverIgjen " & bokstaverIgjen)


    End Sub


    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        bokstavsjekk("A", 11)
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        bokstavsjekk("B", 12)
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        bokstavsjekk("C", 13)
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        bokstavsjekk("D", 14)
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        bokstavsjekk("E", 15)
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        bokstavsjekk("F", 16)
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        bokstavsjekk("G", 17)
    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        bokstavsjekk("H", 18)
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        bokstavsjekk("I", 19)
    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        bokstavsjekk("J", 20)
    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        bokstavsjekk("K", 21)
    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click
        bokstavsjekk("L", 22)
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        bokstavsjekk("M", 23)
    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        bokstavsjekk("N", 24)
    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
        bokstavsjekk("O", 25)
    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click
        bokstavsjekk("P", 26)
    End Sub

    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        bokstavsjekk("Q", 27)
    End Sub

    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click
        bokstavsjekk("R", 28)
    End Sub

    Private Sub Label29_Click(sender As Object, e As EventArgs) Handles Label29.Click
        bokstavsjekk("S", 29)
    End Sub

    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
        bokstavsjekk("T", 30)
    End Sub

    Private Sub Label31_Click(sender As Object, e As EventArgs) Handles Label31.Click
        bokstavsjekk("U", 31)
    End Sub

    Private Sub Label32_Click(sender As Object, e As EventArgs) Handles Label32.Click
        bokstavsjekk("V", 32)
    End Sub

    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click
        bokstavsjekk("W", 33)
    End Sub

    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        bokstavsjekk("X", 34)
    End Sub

    Private Sub Label35_Click(sender As Object, e As EventArgs) Handles Label35.Click
        bokstavsjekk("Y", 35)
    End Sub

    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        bokstavsjekk("Z", 36)
    End Sub

    Private Sub Label37_Click(sender As Object, e As EventArgs) Handles Label37.Click
        bokstavsjekk("Æ", 37)
    End Sub

    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click
        bokstavsjekk("Ø", 38)
    End Sub

    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click
        bokstavsjekk("Å", 39)
    End Sub


    Private Sub btn2player_Click(sender As Object, e As EventArgs) Handles btn2player.Click
        'viser tomt bilde / bakgrunnsbilde
        For i = 2 To 9
            Me.Controls("Picturebox" & i).Visible = False
        Next
        PictureBox1.Visible = True


        'Starter spillet

        'Viser bokstaver i bokstavsky
        For i = 11 To 39 'skal vise label 11 til og med label 39
            Me.Controls("label" & i).Visible = True
        Next

        'Lar motspiller taste inn spillordet i en inputbox, som konverteres til store bokstaver
        Dim spillord As String


        'sjekker om ordet er godkjent (1-10 bokstaver, A-Å, ingen spesialtegn eller tall)
        Dim godkjent As Boolean
        Dim antallGodkjenteBokstaver
        godkjent = False


        Do
            antallGodkjenteBokstaver = 0
            Do 'denne løkka sjekker lengden på ordet
                spillord = InputBox("Skriv inn ord. Maks 10 bokstaver. A-Å").ToUpper
            Loop Until 0 < spillord.Length And spillord.Length < 11

            spillmatrise = spillord.ToCharArray 'laster inn ord i matrise

            For i = 0 To (spillord.Length - 1) 'løkke: kontrollerer at hver bokstav er i området A-Å  XXX fungerer ikke med Æ og Ø..
                godkjent = spillmatrise(i) Like "[A-Å]"
                If spillmatrise(i) = "Æ" Or spillmatrise(i) = "Ø" Then 'Nødløsning for å få godkjent bokstavene Æ og Ø
                    godkjent = True
                End If
                ListBox1.Items.Add("bokstav nr " & (i + 1) & " : " & spillmatrise(i) & " " & godkjent) 'XXX til debug-bruk
                If godkjent = True Then
                    antallGodkjenteBokstaver = antallGodkjenteBokstaver + 1
                End If
            Next

        Loop Until antallGodkjenteBokstaver = spillord.Length







        spillordlengde = spillord.Length
        MsgBox("Spillord: " & spillord)

        'spillmatrise = spillord.ToCharArray
        MsgBox("spillmatrise: " & spillmatrise(0))

        'legger inn underscores/"nullstiller" i alle labels
        For i = 1 To 10
            Me.Controls("label" & i).Text = "_"
        Next

        'viser underscores basert på spillordlengde
        For i = 1 To spillordlengde
            Me.Controls("label" & i).Visible = True
        Next

        'setter variabelen "liv" til 8 (som tilsvarer maksimalt antall streker på tegningen)
        liv = 8 'xxx har muligens satt denne til et annet tall enn 8 for å kunne teste
        MsgBox("liv = " & liv)

        'setter variabelen bokstaverIgjen lik lengden på spillordet
        bokstaverIgjen = spillordlengde
        MsgBox("Antall bokstaver igjen å finne/bokstaverIgjen " & bokstaverIgjen)

    End Sub
End Class
