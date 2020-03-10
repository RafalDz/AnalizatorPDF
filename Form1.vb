Imports System.IO
Imports System.Text

Public Class Form1

    Dim pages As Single
    Dim plik As String
    Dim ilosc As Single
    Dim danaStrony, iloscStron As String
    Dim zapisanyRozmiarCiag As String
    Dim kolejneznaki, kolejneznakiA, kolejneznakiB As Single
    Dim bokA, bokB As Single
    Dim rekord As Byte
    Dim policz, zlicz As Single
    Dim odczyt, znalazlemlStrony As Boolean
    Dim spr, autor, czytajod As Byte
    Dim wymiarA4, wymiarA3, wymiarA3m, wymiarA2, wymiarA1, wymiarA0 As Single
    Dim iloscSTR As String



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        wymiarA4 = 210
        wymiarA3 = 297
        wymiarA3m = 297
        wymiarA2 = 420
        wymiarA1 = 594
        wymiarA0 = 841
        czysc()
    End Sub




    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()
        czysc()


        odczyt = False
        'openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = "pdf files (*.pdf)|*.pdf|txt files (*.txt)|*.txt|All files (*.*)|*.*"
        openFileDialog1.FilterIndex = 1
        openFileDialog1.RestoreDirectory = True

        Label3.Text = "Analizuje plik,  proszę o cierpliwość... "
        Label3.BackColor = Color.Yellow


        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then

                    ' wyciagniecie nazwy pliku
                    Dim nazwa As String
                    Dim Separator As Char = CChar("\")
                    Dim Rezultat As String() = New String() {}
                    Dim wielkosc As Double

                    plik = openFileDialog1.FileName
                    Rezultat = plik.Split(Separator)
                    nazwa = Rezultat(UBound(Rezultat))
                    wielkosc = Int(myStream.Length)

                    Label1.Text = nazwa 'Wczytano plik:

                    Select Case wielkosc  'Wielkość pliku
                        Case Is > 1099511627776
                            Label2.Text = Format$((myStream.Length / 1099511627776), "0.0") & " TB"
                        Case Is > 134217728
                            Label2.Text = Format$((myStream.Length / 134217728), "0.0") & " GB"
                        Case Is > 1048576
                            Label2.Text = Format$((myStream.Length / 1048576), "0.0") & " MB"
                        Case Is > 1024
                            Label2.Text = Format$((myStream.Length / 1024), "0.0") & " KB"
                        Case Else
                            Label2.Text = Format$((myStream.Length), "0.0") & " B"
                    End Select

                    'Open the stream and read it back.
                    myStream = File.OpenRead(plik)
                    Dim b(HScrollBar1.Value) As Byte ' liczba w nawiasie wilkosc bufora w bajtach domyslna 1024 teraz pobierana z suwaka 
                    Dim temp As UTF8Encoding = New UTF8Encoding(True)

                    Do While myStream.Read(b, 0, b.Length) > 0
                        ' podaje wersje odczytanego pdf-a
                        If spr = 0 Then
                            Label13.Text = Mid(temp.GetString(b), 6, 3)
                            spr = 1
                        End If


                        ''   /Count  lub "/N    szuka ilosci stron w ciagu
                        Dim znalezionyLiczbaStron1 As Integer = temp.GetString(b).IndexOf("/Count ")
                        Dim znalezionyLiczbaStron2 As Integer = temp.GetString(b).IndexOf("/N ")

                        If znalezionyLiczbaStron2 <> -1 And znalazlemlStrony = False Then

                            znalazlemlStrony = True
                            iloscStron = temp.GetString(b).Chars((znalezionyLiczbaStron2))

                            Do While iloscStron <> "/"

                                iloscStron = temp.GetString(b).Chars((znalezionyLiczbaStron2 + 4) + kolejneznakiB)
                                kolejneznakiB += 1
                                '   Label7.Text = Label7.Text & vbCrLf & iloscStron
                            Loop
                            iloscSTR = Mid(temp.GetString(b), znalezionyLiczbaStron2 + 4, kolejneznakiB)

                        ElseIf znalezionyLiczbaStron1 <> -1 Then
                            ' If znalezionyLiczbaStron1 <> -1 And znalazlemlStrony = False Then
                            znalazlemlStrony = True
                            iloscStron = temp.GetString(b).Chars((znalezionyLiczbaStron1))

                            Do While iloscStron <> " "

                                iloscStron = temp.GetString(b).Chars((znalezionyLiczbaStron1 + 8) + kolejneznakiA)
                                kolejneznakiA += 1
                            Loop
                            iloscSTR = Mid(temp.GetString(b), znalezionyLiczbaStron1 + 8, kolejneznakiA)

                        End If

                        Label10.Text = iloscSTR  ' wyswietla ilosc stron na podstawie danej PAGE w pdf-ie



                        ''   /MediaBox[0 0 595.276 841.89]   lub /MediaBox [0 0 595.276 841.89]  szukana wymiarów stron w ciagu
                        Dim szukaj As String = "/MediaBox"
                        Dim znaleziony As Integer = temp.GetString(b).IndexOf(szukaj)
                        If znaleziony <> -1 Then
                            odczyt = True
                            ilosc = ilosc + 1

                            danaStrony = temp.GetString(b).Chars((znaleziony))


                            Do While danaStrony <> "]"
                                danaStrony = temp.GetString(b).Chars((znaleziony) + kolejneznaki)
                                kolejneznaki += 1
                            Loop

                            '  sprawdzenie  w czy jest spacja przed nawiasem 
                            If (Mid(temp.GetString(b), znaleziony + 10, 1)) = " " Then czytajod = 12
                            If (Mid(temp.GetString(b), znaleziony + 10, 1)) = "[" Then czytajod = 11

                            zapisanyRozmiarCiag = Mid(temp.GetString(b), znaleziony + czytajod, kolejneznaki - czytajod)

                            'wlaczenie dodatkowych pol gdy sa odkryte
                            If autor >= 4 Then
                                If CheckBox1.Checked = False Then TextBox1.Text = TextBox1.Text & "(" & ilosc & ") ciag = " & zapisanyRozmiarCiag & vbCrLf
                                If CheckBox1.Checked = True Then TextBox1.Text = TextBox1.Text & temp.GetString(b) & vbCrLf
                            End If

                            ' Trim - usuwa spacje z poczatku i konca ciagu
                            zapisanyRozmiarCiag = Trim(zapisanyRozmiarCiag)

                            bokA = podajWymiarA(zapisanyRozmiarCiag)
                            bokB = podajWymiarB(zapisanyRozmiarCiag)

                            If autor >= 4 Then TextBox2.Text = TextBox2.Text & "(" & ilosc & ") " & bokA & "x" & bokB & vbCrLf

                            If bokA <= bokB Then
                                Select Case bokA
                                    Case Is <= wymiarA4
                                        wstawA4(bokA, bokB)
                                    Case wymiarA4 + 0.1 To wymiarA3

                                        If bokB <= 420 Then wstawA3(bokA, bokB)
                                        If bokB > 420 Then wstawA3m(bokA, bokB)

                                    Case wymiarA3 + 0.1 To wymiarA2
                                        wstawA2(bokA, bokB)
                                    Case wymiarA2 + 0.1 To wymiarA1
                                        wstawA1(bokA, bokB)
                                    Case wymiarA1 + 0.1 To wymiarA0
                                        wstawA0(bokA, bokB)
                                    Case Is >= wymiarA0 + 0.1
                                        wstawA00(bokA, bokB)
                                End Select
                            ElseIf bokA > bokB Then
                                Select Case bokB
                                    Case Is <= wymiarA4
                                        wstawA4(bokA, bokB)
                                    Case wymiarA4 + 0.1 To wymiarA3

                                        If bokA <= 420 Then wstawA3(bokA, bokB)
                                        If bokA > 420 Then wstawA3m(bokA, bokB)
                                       '' wstawA3(bokA, bokB)
                                    Case wymiarA3 + 0.1 To wymiarA2
                                        wstawA2(bokA, bokB)
                                    Case wymiarA2 + 0.1 To wymiarA1
                                        wstawA1(bokA, bokB)
                                    Case wymiarA1 + 0.1 To wymiarA0
                                        wstawA0(bokA, bokB)
                                    Case Is >= wymiarA0 + 0.1
                                        wstawA00(bokA, bokB)
                                End Select
                            End If

                        End If
                        Label8.Text = ilosc ' wyswietla ilosc stron na podstawie rekordow
                        kolejneznaki = 0
                        ' kolejneznakiA = 0
                        ' kolejneznakiB = 0

                    Loop
                    myStream.Close()
                End If
            Catch Ex As Exception
                MessageBox.Show("Nie można odczytać pliku z dysku. błąd nr: " & Ex.Message)
                Label3.Text = "Nie można odczytać pliku z dysku. błąd nr: " & Ex.Message
                Label3.BackColor = Color.Red
            Finally
                Select Case odczyt
                    Case True
                        Label3.Text = "Obliczam..."
                        Label3.BackColor = Color.Green
                        oblicz()
                        obliczKopie()
                    Case False
                        Label3.Text = "Nie udało sie pobrać danych    :-("
                        Label3.BackColor = Color.Red
                End Select

                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        czysc()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' TextBox1.Text = TextBox1.Text & temp.GetString(b) & vbCrLf
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        oblicz()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        HScrollBarA4.Value = 0
        HScrollBarA3.Value = 0
        HScrollBarA3m.Value = 0
        HScrollBarA2.Value = 0
        HScrollBarA1.Value = 0
        HScrollBarA0.Value = 0
        wymiarA4 = 210
        wymiarA3 = 297
        wymiarA3m = 297
        wymiarA2 = 420
        wymiarA1 = 594
        wymiarA0 = 841
        czysc()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        HScrollBarA4.Value = 1
        HScrollBarA3.Value = 1
        HScrollBarA3m.Value = 1
        HScrollBarA2.Value = 1
        HScrollBarA1.Value = 1
        HScrollBarA0.Value = 1
        wymiarA4 = 210 + 1
        wymiarA3 = 297 + 1
        wymiarA3m = 297 + 1
        wymiarA2 = 420 + 1
        wymiarA1 = 594 + 1
        wymiarA0 = 841 + 1
        czysc()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        HScrollBarA4.Value = 3
        HScrollBarA3.Value = 3
        HScrollBarA3m.Value = 3
        HScrollBarA2.Value = 3
        HScrollBarA1.Value = 3
        HScrollBarA0.Value = 3
        wymiarA4 = 210 + 3
        wymiarA3 = 297 + 3
        wymiarA3m = 297 + 3
        wymiarA2 = 420 + 3
        wymiarA1 = 594 + 3
        wymiarA0 = 841 + 3
        czysc()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        HScrollBarA4.Value = 5
        HScrollBarA3.Value = 5
        HScrollBarA3m.Value = 5
        HScrollBarA2.Value = 5
        HScrollBarA1.Value = 5
        HScrollBarA0.Value = 5
        wymiarA4 = 210 + 5
        wymiarA3 = 297 + 5
        wymiarA3m = 297 + 5
        wymiarA2 = 420 + 5
        wymiarA1 = 594 + 5
        wymiarA0 = 841 + 5
        czysc()
    End Sub

    Private Sub LablAutor_Click(sender As Object, e As EventArgs) Handles LablAutor.Click
        autor += 1
        If autor >= 4 Then
            LablAutor.ForeColor = Color.Teal
            FormBorderStyle = FormBorderStyle.Sizable
        End If
    End Sub

    Private Sub HScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar1.Scroll

        Select Case HScrollBar1.Value  'Wielkość pliku
            Case Is > 1048576
                Label16.Text = "bufor:" & Format$((HScrollBar1.Value / 1048576), "0.0") & "MB"
            Case Else
                Label16.Text = "bufor:" & Format((HScrollBar1.Value / 1024), "0.0") & "KB"
        End Select
    End Sub

    Private Sub HScrollBarA4_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBarA4.Scroll
        wymiarA4 = 210 + HScrollBarA4.Value
        czysc()
    End Sub

    Private Sub HScrollBarA3_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBarA3.Scroll
        wymiarA3 = 297 + HScrollBarA3.Value
        czysc()
    End Sub

    Private Sub HScrollBarA3m_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBarA3m.Scroll
        wymiarA3m = 297 + HScrollBarA3m.Value
        czysc()
    End Sub
    Private Sub HScrollBarA2_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBarA2.Scroll
        wymiarA2 = 420 + HScrollBarA2.Value
        czysc()
    End Sub

    Private Sub HScrollBarA1_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBarA1.Scroll
        wymiarA1 = 594 + HScrollBarA1.Value
        czysc()
    End Sub

    Private Sub HScrollBarA0_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBarA0.Scroll
        wymiarA0 = 841 + HScrollBarA0.Value
        czysc()
    End Sub

    Private Function podajWymiarA(zapisanyRozmiarCiag As String) As Single

        Dim l1, l2 As Single
        Dim dana, dana_1, dana_2, dana_3 As Byte
        Dim dlCiagu As Byte

        dana = 0
        dlCiagu = Len(zapisanyRozmiarCiag)

        For petla = 1 To dlCiagu
            If Mid(zapisanyRozmiarCiag, petla, 1) = "." Then Mid$(zapisanyRozmiarCiag, petla, 1) = ","
        Next

        For petla = 1 To dlCiagu
            If Mid(zapisanyRozmiarCiag, petla, 1) = " " And dana = 0 Then
                dana = 1
                dana_1 = petla
            End If

            If Mid(zapisanyRozmiarCiag, petla, 1) = " " And dana = 1 And petla <> dana_1 Then
                dana = 2
                dana_2 = petla
            End If

            If Mid(zapisanyRozmiarCiag, petla, 1) = " " And dana = 2 And petla <> dana_2 Then
                dana = 3
                dana_3 = petla
            End If
        Next

        l1 = CSng((Mid(zapisanyRozmiarCiag, 1, (dana_1 - 1))))
        l2 = CSng((Mid(zapisanyRozmiarCiag, (dana_2 + 1), ((dana_3 - dana_2) - 1))))

        If l1 < 0 Then l1 = l1 * (-1)
        If l2 < 0 Then l2 = l2 * (-1)

        ' w mm 
        Return CUInt((l1 + l2) * 0.3528)
    End Function

    Private Function podajWymiarB(zapisanyRozmiarCiag As String) As Single

        Dim l1, l2 As Single
        Dim dana, dana_1, dana_2, dana_3 As Byte
        Dim dlCiagu As Byte

        dana = 0
        dlCiagu = Len(zapisanyRozmiarCiag)

        For petla = 1 To dlCiagu
            If Mid(zapisanyRozmiarCiag, petla, 1) = "." Then Mid$(zapisanyRozmiarCiag, petla, 1) = ","
        Next

        For petla = 1 To dlCiagu
            If Mid(zapisanyRozmiarCiag, petla, 1) = " " And dana = 0 Then
                dana = 1
                dana_1 = petla
            End If

            If Mid(zapisanyRozmiarCiag, petla, 1) = " " And dana = 1 And petla <> dana_1 Then
                dana = 2
                dana_2 = petla
            End If

            If Mid(zapisanyRozmiarCiag, petla, 1) = " " And dana = 2 And petla <> dana_2 Then
                dana = 3
                dana_3 = petla
            End If
        Next

        l1 = CSng((Mid(zapisanyRozmiarCiag, (dana_1 + 1), ((dana_2 - dana_1) - 1))))
        l2 = CSng((Mid(zapisanyRozmiarCiag, (dana_3 + 1), (dlCiagu - dana_3))))

        If l1 < 0 Then l1 = l1 * (-1)
        If l2 < 0 Then l2 = l2 * (-1)

        ' w mm 
        Return CUInt((l1 + l2) * 0.3528)
    End Function

    Sub wstawA4(bokA, bokB)

        rekord = DataGridView_A4.RowCount - 1
        DataGridView_A4.Rows.Add()

        rekord += 1

        If bokA <= bokB Then
            DataGridView_A4.Rows(rekord - 1).Cells(0).Value = bokA
            DataGridView_A4.Rows(rekord - 1).Cells(1).Value = bokB
        ElseIf bokA > bokB Then
            DataGridView_A4.Rows(rekord - 1).Cells(0).Value = bokB
            DataGridView_A4.Rows(rekord - 1).Cells(1).Value = bokA
        End If
    End Sub

    Sub wstawA3(bokA, bokB)

        rekord = DataGridView_A3.RowCount - 1
        DataGridView_A3.Rows.Add()

        rekord += 1

        If bokA <= bokB Then
            DataGridView_A3.Rows(rekord - 1).Cells(0).Value = bokA
            DataGridView_A3.Rows(rekord - 1).Cells(1).Value = bokB
        ElseIf bokA > bokB Then
            DataGridView_A3.Rows(rekord - 1).Cells(0).Value = bokB
            DataGridView_A3.Rows(rekord - 1).Cells(1).Value = bokA
        End If
    End Sub

    Sub wstawA3m(bokA, bokB)

        rekord = DataGridView_A3m.RowCount - 1
        DataGridView_A3m.Rows.Add()

        rekord += 1

        If bokA <= bokB Then
            DataGridView_A3m.Rows(rekord - 1).Cells(0).Value = bokA
            DataGridView_A3m.Rows(rekord - 1).Cells(1).Value = bokB
        ElseIf bokA > bokB Then
            DataGridView_A3m.Rows(rekord - 1).Cells(0).Value = bokB
            DataGridView_A3m.Rows(rekord - 1).Cells(1).Value = bokA
        End If
    End Sub

    Sub wstawA2(bokA, bokB)

        rekord = DataGridView_A2.RowCount - 1
        DataGridView_A2.Rows.Add()

        rekord += 1

        If bokA <= bokB Then
            DataGridView_A2.Rows(rekord - 1).Cells(0).Value = bokA
            DataGridView_A2.Rows(rekord - 1).Cells(1).Value = bokB
        ElseIf bokA > bokB Then
            DataGridView_A2.Rows(rekord - 1).Cells(0).Value = bokB
            DataGridView_A2.Rows(rekord - 1).Cells(1).Value = bokA
        End If
    End Sub

    Sub wstawA1(bokA, bokB)

        rekord = DataGridView_A1.RowCount - 1
        DataGridView_A1.Rows.Add()

        rekord += 1

        If bokA <= bokB Then
            DataGridView_A1.Rows(rekord - 1).Cells(0).Value = bokA
            DataGridView_A1.Rows(rekord - 1).Cells(1).Value = bokB
        ElseIf bokA > bokB Then
            DataGridView_A1.Rows(rekord - 1).Cells(0).Value = bokB
            DataGridView_A1.Rows(rekord - 1).Cells(1).Value = bokA
        End If
    End Sub

    Sub wstawA0(bokA, bokB)

        rekord = DataGridView_A0.RowCount - 1
        DataGridView_A0.Rows.Add()

        rekord += 1

        If bokA <= bokB Then
            DataGridView_A0.Rows(rekord - 1).Cells(0).Value = bokA
            DataGridView_A0.Rows(rekord - 1).Cells(1).Value = bokB
        ElseIf bokA > bokB Then
            DataGridView_A0.Rows(rekord - 1).Cells(0).Value = bokB
            DataGridView_A0.Rows(rekord - 1).Cells(1).Value = bokA
        End If
    End Sub

    Sub wstawA00(bokA, bokB)

        rekord = DataGridView_A00.RowCount - 1
        DataGridView_A00.Rows.Add()

        rekord += 1

        If bokA <= bokB Then
            DataGridView_A00.Rows(rekord - 1).Cells(0).Value = bokA
            DataGridView_A00.Rows(rekord - 1).Cells(1).Value = bokB
        ElseIf bokA > bokB Then
            DataGridView_A00.Rows(rekord - 1).Cells(0).Value = bokB
            DataGridView_A00.Rows(rekord - 1).Cells(1).Value = bokA
        End If
    End Sub

    Sub oblicz()

        'oblicza A4
        policz = 0
        zlicz = 0
        If (DataGridView_A4.RowCount) >= 1 And DataGridView_A4.Rows(0).Cells(0).Value <> 0 Then
            For licz_A4 = 0 To (DataGridView_A4.RowCount - 2)
                policz = DataGridView_A4.Rows(licz_A4).Cells(1).Value
                policz = policz / 297
                If policz < 1 Then policz = 1
                DataGridView_A4.Rows(licz_A4).Cells(2).Value = policz
                zlicz = zlicz + DataGridView_A4.Rows(licz_A4).Cells(2).Value
            Next
            Select Case (zlicz - Fix(zlicz))
                Case 0
                    Lbl_A4.Text = zlicz
                Case Else
                    Lbl_A4.Text = FormatNumber(zlicz, 2)
            End Select
        End If

        'oblicza A3
        policz = 0
        zlicz = 0
        If (DataGridView_A3.RowCount) >= 1 And DataGridView_A3.Rows(0).Cells(0).Value <> 0 Then
            For licz_A3 = 0 To (DataGridView_A3.RowCount - 2)
                policz = DataGridView_A3.Rows(licz_A3).Cells(1).Value
                policz = policz / 420
                If policz < 1 Then policz = 1
                DataGridView_A3.Rows(licz_A3).Cells(2).Value = policz
                zlicz = zlicz + DataGridView_A3.Rows(licz_A3).Cells(2).Value
            Next
            Select Case (zlicz - Fix(zlicz))
                Case 0
                    Lbl_A3.Text = zlicz
                Case Else
                    Lbl_A3.Text = FormatNumber(zlicz, 2)
            End Select
        End If

        'oblicza A3m
        policz = 0
        zlicz = 0
        If (DataGridView_A3m.RowCount) >= 1 And DataGridView_A3m.Rows(0).Cells(0).Value <> 0 Then
            For licz_A3m = 0 To (DataGridView_A3m.RowCount - 2)
                policz = DataGridView_A3m.Rows(licz_A3m).Cells(1).Value
                policz = policz / 420
                If policz < 1 Then policz = 1
                DataGridView_A3m.Rows(licz_A3m).Cells(2).Value = policz
                zlicz = zlicz + DataGridView_A3m.Rows(licz_A3m).Cells(2).Value
            Next
            Select Case (zlicz - Fix(zlicz))
                Case 0
                    Lbl_A3m.Text = zlicz
                Case Else
                    Lbl_A3m.Text = FormatNumber(zlicz, 2)
            End Select
        End If

        'oblicza A2
        policz = 0
        zlicz = 0
        If (DataGridView_A2.RowCount) >= 1 And DataGridView_A2.Rows(0).Cells(0).Value <> 0 Then
            For licz_A2 = 0 To (DataGridView_A2.RowCount - 2)
                policz = DataGridView_A2.Rows(licz_A2).Cells(1).Value
                policz = policz / 594
                If policz < 1 Then policz = 1
                DataGridView_A2.Rows(licz_A2).Cells(2).Value = policz
                zlicz = zlicz + DataGridView_A2.Rows(licz_A2).Cells(2).Value
            Next
            Select Case (zlicz - Fix(zlicz))
                Case 0
                    Lbl_A2.Text = zlicz
                Case Else
                    Lbl_A2.Text = FormatNumber(zlicz, 2)
            End Select
        End If

        'oblicza A1
        policz = 0
        zlicz = 0
        If (DataGridView_A1.RowCount) >= 1 And DataGridView_A1.Rows(0).Cells(0).Value <> 0 Then
            For licz_A1 = 0 To (DataGridView_A1.RowCount - 2)
                policz = DataGridView_A1.Rows(licz_A1).Cells(1).Value
                policz = policz / 841
                If policz < 1 Then policz = 1
                DataGridView_A1.Rows(licz_A1).Cells(2).Value = policz
                zlicz = zlicz + DataGridView_A1.Rows(licz_A1).Cells(2).Value
            Next
            Select Case (zlicz - Fix(zlicz))
                Case 0
                    Lbl_A1.Text = zlicz
                Case Else
                    Lbl_A1.Text = FormatNumber(zlicz, 2)
            End Select
        End If

        'oblicza A0
        policz = 0
        zlicz = 0
        If (DataGridView_A0.RowCount) >= 1 And DataGridView_A0.Rows(0).Cells(0).Value <> 0 Then
            For licz_A0 = 0 To (DataGridView_A0.RowCount - 2)
                policz = DataGridView_A0.Rows(licz_A0).Cells(1).Value
                policz = policz / 1189
                If policz < 1 Then policz = 1
                DataGridView_A0.Rows(licz_A0).Cells(2).Value = policz
                zlicz = zlicz + DataGridView_A0.Rows(licz_A0).Cells(2).Value
            Next
            Select Case (zlicz - Fix(zlicz))
                Case 0
                    Lbl_A0.Text = zlicz
                Case Else
                    Lbl_A0.Text = FormatNumber(zlicz, 2)
            End Select
        End If

        'oblicza A00
        policz = 0
        zlicz = 0
        If (DataGridView_A00.RowCount) >= 1 And DataGridView_A00.Rows(0).Cells(0).Value <> 0 Then
            For licz_A00 = 0 To (DataGridView_A00.RowCount - 2)
                policz = DataGridView_A00.Rows(licz_A00).Cells(1).Value
                policz = policz / 1189
                If policz < 1 Then policz = 1
                DataGridView_A00.Rows(licz_A00).Cells(2).Value = policz
                zlicz = zlicz + DataGridView_A00.Rows(licz_A00).Cells(2).Value
            Next
            Select Case (zlicz - Fix(zlicz))
                Case 0
                    Lbl_A00.Text = zlicz
                Case Else
                    Lbl_A00.Text = FormatNumber(zlicz, 2)
            End Select
        End If

        'likwidacja zaznaczenia w tabelach
        DataGridView_A4.ClearSelection()
        DataGridView_A3.ClearSelection()
        DataGridView_A3m.ClearSelection()
        DataGridView_A2.ClearSelection()
        DataGridView_A1.ClearSelection()
        DataGridView_A0.ClearSelection()
        DataGridView_A00.ClearSelection()

        Label3.Text = "Gotowe    :-)"
        Label3.BackColor = Color.Lime
    End Sub

    Private Sub HScrr_komplety_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrr_komplety.Scroll
        Lbl_komplety.Text = "x " & HScrr_komplety.Value
        obliczKopie()
    End Sub

    Sub obliczKopie()

        If Lbl_A4.Text <> "" Then Lbl_kompA4.Text = HScrr_komplety.Value * CSng(Lbl_A4.Text)
        If Lbl_A3.Text <> "" Then Lbl_kompA3.Text = HScrr_komplety.Value * CSng(Lbl_A3.Text)
        If Lbl_A3m.Text <> "" Then Lbl_kompA3m.Text = HScrr_komplety.Value * CSng(Lbl_A3m.Text)
        If Lbl_A2.Text <> "" Then Lbl_kompA2.Text = HScrr_komplety.Value * CSng(Lbl_A2.Text)
        If Lbl_A1.Text <> "" Then Lbl_kompA1.Text = HScrr_komplety.Value * CSng(Lbl_A1.Text)
        If Lbl_A0.Text <> "" Then Lbl_kompA0.Text = HScrr_komplety.Value * CSng(Lbl_A0.Text)
        If Lbl_A00.Text <> "" Then Lbl_kompA00.Text = HScrr_komplety.Value * CSng(Lbl_A00.Text)

    End Sub

    Sub czysc()
        TextBox1.Text = ""
        TextBox2.Text = ""
        Label1.Text = ""
        Label2.Text = ""
        Label3.Text = ""
        Label5.Text = ""
        Label8.Text = ""
        Label10.Text = ""
        Label13.Text = ""
        ilosc = 0
        kolejneznaki = 0
        kolejneznakiA = 0
        kolejneznakiB = 0
        rekord = 0
        spr = 0
        znalazlemlStrony = False

        usunA4()
        usunA3()
        usunA3m()
        usunA2()
        usunA1()
        usunA0()
        usunA00()

        Lbl_A4.Text = ""
        Lbl_A3.Text = ""
        Lbl_A3m.Text = ""
        Lbl_A2.Text = ""
        Lbl_A1.Text = ""
        Lbl_A0.Text = ""
        Lbl_A00.Text = ""

        Lbl_kompA4.Text = ""
        Lbl_kompA3.Text = ""
        Lbl_kompA3m.Text = ""
        Lbl_kompA2.Text = ""
        Lbl_kompA1.Text = ""
        Lbl_kompA0.Text = ""
        Lbl_kompA00.Text = ""

        Lab_W_A4.Text = wymiarA4
        Lab_W_A3.Text = wymiarA3
        Lab_W_A3m.Text = wymiarA3m
        Lab_W_A2.Text = wymiarA2
        Lab_W_A1.Text = wymiarA1
        Lab_W_A0.Text = wymiarA0
        Lbl_komplety.Text = "x " & HScrr_komplety.Value

        Select Case wymiarA4
            Case 210
                Lab_W_A4.BackColor = Color.AliceBlue
            Case Else
                Lab_W_A4.BackColor = Color.Pink
        End Select

        Select Case wymiarA3
            Case 297
                Lab_W_A3.BackColor = Color.AliceBlue
            Case Else
                Lab_W_A3.BackColor = Color.Pink
        End Select

        Select Case wymiarA3m
            Case 297
                Lab_W_A3m.BackColor = Color.AliceBlue
            Case Else
                Lab_W_A3m.BackColor = Color.Pink
        End Select

        Select Case wymiarA2
            Case 420
                Lab_W_A2.BackColor = Color.AliceBlue
            Case Else
                Lab_W_A2.BackColor = Color.Pink
        End Select

        Select Case wymiarA1
            Case 594
                Lab_W_A1.BackColor = Color.AliceBlue
            Case Else
                Lab_W_A1.BackColor = Color.Pink
        End Select

        Select Case wymiarA0
            Case 841
                Lab_W_A0.BackColor = Color.AliceBlue
            Case Else
                Lab_W_A0.BackColor = Color.Pink
        End Select
    End Sub

    Sub usunA4()

        If (DataGridView_A4.RowCount) <= 1 Then Exit Sub

        'likwiduje wszystkie wiersze w DataGridView_A4
        For usun = 1 To (DataGridView_A4.RowCount - 1)
            DataGridView_A4.Rows.RemoveAt(0)
        Next
    End Sub

    Sub usunA3()

        If (DataGridView_A3.RowCount) <= 1 Then Exit Sub

        'likwiduje wszystkie wiersze w DataGridView_A3
        For usun = 1 To (DataGridView_A3.RowCount - 1)
            DataGridView_A3.Rows.RemoveAt(0)
        Next
    End Sub

    Sub usunA3m()

        If (DataGridView_A3m.RowCount) <= 1 Then Exit Sub

        'likwiduje wszystkie wiersze w DataGridView_A3
        For usun = 1 To (DataGridView_A3m.RowCount - 1)
            DataGridView_A3m.Rows.RemoveAt(0)
        Next
    End Sub

    Sub usunA2()

        If (DataGridView_A2.RowCount) <= 1 Then Exit Sub

        'likwiduje wszystkie wiersze w DataGridView_A2
        For usun = 1 To (DataGridView_A2.RowCount - 1)
            DataGridView_A2.Rows.RemoveAt(0)
        Next
    End Sub

    Sub usunA1()

        If (DataGridView_A1.RowCount) <= 1 Then Exit Sub

        'likwiduje wszystkie wiersze w DataGridView_A1
        For usun = 1 To (DataGridView_A1.RowCount - 1)
            DataGridView_A1.Rows.RemoveAt(0)
        Next
    End Sub

    Sub usunA0()

        If (DataGridView_A0.RowCount) <= 1 Then Exit Sub

        'likwiduje wszystkie wiersze w DataGridView_A0
        For usun = 1 To (DataGridView_A0.RowCount - 1)
            DataGridView_A0.Rows.RemoveAt(0)
        Next
    End Sub

    Sub usunA00()

        If (DataGridView_A00.RowCount) <= 1 Then Exit Sub

        'likwiduje wszystkie wiersze w DataGridView_A00
        For usun = 1 To (DataGridView_A00.RowCount - 1)
            DataGridView_A00.Rows.RemoveAt(0)
        Next
    End Sub





End Class

