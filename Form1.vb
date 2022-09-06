'**************************
'* Sodavandsautomat       *
'**************************
'* Iván Viemoes Cuevas    *
'* Programmering C        *
'* 09-02-2020             *
'**************************

Public Class Form1

    'Variable til brug for drag-n-drop & til køb og valg af sodavand
    Dim Dragging As Boolean = False
    Dim PictureBoxWhileDragging As PictureBox = Nothing
    Dim StartMousePos As Point
    Dim sum As Integer = 0
    Dim PictureBoxClick As Integer = 0
    Dim SodavandBetalt As PictureBox = Nothing
    Dim SodaOutput As Boolean = False
    Dim Sodavand As String


    'Venstre museknap aktiveres på den PictureBox, som skal trækkes
    'Bemærk: så længe museknappen er nede, er det objekt, som blev klikket på,
    'altså PictureBoxSomKanTrækkes1 eller 2, det aktive objekt, 
    'dvs. PictureBoxSomKanTrækkes modtager f.eks. .MouseMove og .MouseUp -events
    Private Sub PictureBoxSomKanTrækkes_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox_1krone.MouseDown, PictureBox_2krone.MouseDown, PictureBox_5krone.MouseDown, PictureBox_10krone.MouseDown, PictureBox_20krone.MouseDown

        'Sæt flag: Drag-n-drop i gang
        Dragging = True

        'Opret ny PictureBox til at vise billede MENS der bliver trukket
        'en kopi af det objekt, der blev klikket på
        PictureBoxWhileDragging = New PictureBox
        PictureBoxWhileDragging.Size = sender.Size
        PictureBoxWhileDragging.Image = sender.Image
        PictureBoxWhileDragging.BackColor = sender.backcolor
        PictureBoxWhileDragging.SizeMode = sender.SizeMode
        PictureBoxWhileDragging.Tag = sender.tag

        'Husk positionen af musemarkøren på det tidspunkt hvor museknappen blev aktiveret
        StartMousePos = sender.Location - e.Location

        Me.Controls.Add(PictureBoxWhileDragging) 'PictureBoxWhileDragging skal "tilhøre" Form1
    End Sub

    'Musemarkøren flyttes mens PictureBoxSomKanTrækkesx er det aktive objekt
    Private Sub PictureBoxSomKanTrækkes_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox_1krone.MouseMove, PictureBox_1krone.MouseDown, PictureBox_2krone.MouseMove, PictureBox_2krone.MouseDown, PictureBox_5krone.MouseMove, PictureBox_5krone.MouseDown, PictureBox_10krone.MouseMove, PictureBox_10krone.MouseDown, PictureBox_20krone.MouseMove, PictureBox_20krone.MouseDown

        If Dragging Then
            'En PictureBox er ved at blive trukket. Flyt billedet med musemarkøren
            PictureBoxWhileDragging.Location = StartMousePos + e.Location
            PictureBoxWhileDragging.BringToFront()
        End If
    End Sub

    'Museknap slippes: drop
    Private Sub PictureBoxSomKanTrækkes_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox_1krone.MouseUp, PictureBox_2krone.MouseUp, PictureBox_5krone.MouseUp, PictureBox_10krone.MouseUp, PictureBox_20krone.MouseUp

        If Dragging Then
            'En PictureBox er ved at blive trukket
            Dragging = False

            'Se om vi er over feltet hvor vi må droppe
            If (PictureBoxWhileDragging.Bounds.IntersectsWith(PictureBox_CoinSlot.Bounds)) Then
                'Drop over PictureBoxFelt(x, y) er tilladt. Sæt billede her.
                sum = PictureBoxWhileDragging.Tag + sum
                TextBox_IndsattePenge.Text = sum & " kr."
            End If
        End If

        'Fjern midlertidigt billede
        PictureBoxWhileDragging.Image = Nothing
        PictureBoxWhileDragging.Visible = False
        PictureBoxWhileDragging = Nothing
    End Sub

    Private Sub Sodavand_Click(sender As Object, e As MouseEventArgs) Handles CocaCola.Click, Fanta.Click, FaxeKondi.Click, Sport.Click

        'Hvis der er en sodavand i lugen, og du prøver at købe en ny,
        'fortæller den dig, du skal tage en ny sodavand først
        If SodaOutput = True Then
            MsgBox("Du skal tage din sodavand før du køber en til.")
            Exit Sub
        End If

        'Den giver denne variabel samme værdi som det valgte billede
        PictureBoxClick = sender.tag

        'Hvis sum er mindre end det sodavanden koster fortæl køberen ikke har indsat nok penge,
        'og afslut processen
        If sum < PictureBoxClick Then
            MsgBox("Du har ikke indsat nok penge")
            Exit Sub
        End If

        'Sæt sum = sum - det sodavanden kostede
        sum = sum - PictureBoxClick

        'Vis den valgte sodavand i lugen
        PictureBox_SodaOutput.Image = sender.Image

        'Der er en sodavand i lugen
        SodaOutput = True

        'Sætter sodavand lig med navnet på det valgte billede
        Sodavand = sender.name

        'Fortæller hvilken sodavand du har købt, for hvor meget, og hvor mange penge du får tilbage
        MsgBox("Du har købt en " & Sodavand & " for " & PictureBoxClick & " kr." & vbNewLine &
               "Du får " & sum & " kr. tilbage")

        'Sætter summen lig med 0, da køberen har fået sine penge tilbage, og viser det i textboxen
        sum = 0
        TextBox_IndsattePenge.Text = sum & " kr."

        'kalder subben der deaktivere mønterne
        DeaktiverMønter()

    End Sub

    Private Sub SodavandOutput_Click(sender As Object, e As MouseEventArgs) Handles PictureBox_SodaOutput.Click
        'Når der klikkes på sodavnden i lugen, vises sodavnden ikke længere
        PictureBox_SodaOutput.Image = Nothing

        'Lugen er tom igen
        SodaOutput = False

        'Kalder subben der aktivere mønterne
        AktiverMønter()

    End Sub

    'Laver en sub, der deaktivere mønterne
    Sub DeaktiverMønter()

        PictureBox_1krone.Enabled = False
        PictureBox_2krone.Enabled = False
        PictureBox_5krone.Enabled = False
        PictureBox_10krone.Enabled = False
        PictureBox_20krone.Enabled = False

    End Sub

    'Laver en sub, der aktivere mønterne
    Sub AktiverMønter()

        PictureBox_1krone.Enabled = True
        PictureBox_2krone.Enabled = True
        PictureBox_5krone.Enabled = True
        PictureBox_10krone.Enabled = True
        PictureBox_20krone.Enabled = True

    End Sub
End Class
