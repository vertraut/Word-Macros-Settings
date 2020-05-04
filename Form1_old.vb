Public Class Form1
    Dim fonts_name As New System.Drawing.Text.InstalledFontCollection
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox2.DataSource = fonts_name.Families
        ComboBox2.DisplayMember = "Name"
        start_program()


    End Sub

    Private Sub start_program()
        Settings_read(0, ComboBox2) ' установка названия шрифта
        Settings_read(1, ComboBox1) ' установка размера шрифта
        Settings_read(8, TextBox2) 'установка цвета шрифта
        Settings_read(9, TextBox3) 'установка цвета фона ячейки

        Settings_read(2, RadioButton5) 'установка выравнивания
        Settings_read(3, ComboBox4) ' установка отступов верх/низ
        Settings_read(4, ComboBox3) ' установка отступов лево/право
        Settings_read(5, CheckBox3)
        Settings_read(6, CheckBox2)
        Settings_read(7, CheckBox1)
        color_set(TextBox4, TextBox3, 0) ' красим образец цвета фона (маленький квадратик)
        color_set(TextBox6, TextBox2, 0) ' красим образец цвета текста (маленький квадратик)
        test_show()
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FileOpen(1, "D:\Macros Settings\settings.txt", OpenMode.Output)
        Dim s As String
        Dim text_align As String = 0
        Dim bold_status, italic_status, underline_status, format_status As String 'переменные для статуса жирного, курсива и подчеркнутого

        If RadioButton1.Checked Then
            text_align = 11
        ElseIf RadioButton2.Checked Then
            text_align = 12
        ElseIf RadioButton3.Checked Then
            text_align = 13
        ElseIf RadioButton4.Checked Then
            text_align = 21
        ElseIf RadioButton5.Checked Then
            text_align = 22
        ElseIf RadioButton6.Checked Then
            text_align = 23
        ElseIf RadioButton7.Checked Then
            text_align = 31
        ElseIf RadioButton8.Checked Then
            text_align = 32
        ElseIf RadioButton9.Checked Then
            text_align = 33
        End If

        If CheckBox1.Checked Then
            underline_status = "1"
        Else underline_status = "0"
        End If
        If CheckBox2.Checked Then
            italic_status = "1"
        Else italic_status = "0"
        End If
        If CheckBox3.Checked Then
            bold_status = "1"
        Else bold_status = "0"
        End If

        If CheckBox5.Checked Then
            format_status = "1"
        Else
            format_status = "0"
        End If

        s = ComboBox2.Text + "#" + ComboBox1.Text + "#" + text_align + "#" + ComboBox4.Text + "#" + ComboBox3.Text + "#" + bold_status + "#" + italic_status + "#" + underline_status + "#" + TextBox2.Text + "#" + TextBox3.Text + "#" + format_status
        Print(1, s)
        FileClose(1)
    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click 'Вставка RGB цвета для текста из буфера
        TextBox2.Text = My.Computer.Clipboard.GetText()
        color_set(TextBox6, TextBox2, 0)
        test_show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click 'Вставка RGB цвета для заливки из буфера
        TextBox3.Text = My.Computer.Clipboard.GetText()
        color_set(TextBox4, TextBox3, 0)
        test_show()
    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click 'При нажатии на кнопку "тест"
        test_show()
    End Sub

    Private Sub test_show()
        Dim text_test As String = "Пример"

        ' ниже конструкция elseif для позиционирования текста, по вертикали текст позиционируется при помощи пустых строк, top -без пустой строк, middle- одна пустая строка, bottom- две пустые строки
        If RadioButton1.Checked Then
            TextBox5.TextAlign = HorizontalAlignment.Left
            TextBox5.Text = text_test
        ElseIf RadioButton2.Checked Then
            TextBox5.Text = text_test
            TextBox5.TextAlign = HorizontalAlignment.Center
        ElseIf RadioButton3.Checked Then
            TextBox5.Text = text_test
            TextBox5.TextAlign = HorizontalAlignment.Right
        ElseIf RadioButton4.Checked Then
            TextBox5.Text = Environment.NewLine + text_test
            TextBox5.TextAlign = HorizontalAlignment.Left
        ElseIf RadioButton5.Checked Then
            TextBox5.Text = Environment.NewLine + text_test
            TextBox5.TextAlign = HorizontalAlignment.Center
        ElseIf RadioButton6.Checked Then
            TextBox5.Text = Environment.NewLine + text_test
            TextBox5.TextAlign = HorizontalAlignment.Right
        ElseIf RadioButton7.Checked Then
            TextBox5.Text = Environment.NewLine + Environment.NewLine + text_test
            TextBox5.TextAlign = HorizontalAlignment.Left
        ElseIf RadioButton8.Checked Then
            TextBox5.Text = Environment.NewLine + Environment.NewLine + text_test
            TextBox5.TextAlign = HorizontalAlignment.Center
        ElseIf RadioButton9.Checked Then
            TextBox5.Text = Environment.NewLine + Environment.NewLine + text_test
            TextBox5.TextAlign = HorizontalAlignment.Right
        End If

        'тут устанавливается цвет шрифта и фона, а также шрифт и его размер
        color_set(TextBox5, TextBox2, 1)
        color_set(TextBox5, TextBox3, 0)

        Dim TempFont = TextBox5.Font
        Dim FontStyle1 = FontStyle.Regular
        If CheckBox2.Checked Then FontStyle1 = FontStyle1 Or FontStyle.Italic
        If CheckBox1.Checked Then FontStyle1 = FontStyle1 Or FontStyle.Underline
        If CheckBox3.Checked Then FontStyle1 = FontStyle1 Or FontStyle.Bold
        TextBox5.Font = New System.Drawing.Font(ComboBox2.Text, 10, FontStyle1)
    End Sub

    Private Sub TextBox2_Validated(sender As Object, e As EventArgs) Handles TextBox2.Validated
        color_set(TextBox6, TextBox2, 0)
    End Sub

    Private Sub TextBox3_Validated(sender As Object, e As EventArgs) Handles TextBox3.Validated
        color_set(TextBox4, TextBox3, 0)
    End Sub



    Private Sub color_set(boxObj As Object, colorObj As Object, goal As Integer) 'Функция применения цвета, boxObj - что красим, colorObj - каким цветом, goal - текст или фон
        If goal = 0 Then
            Try
                boxObj.BackColor = Color.FromArgb(CInt(colorObj.Text.Split(",")(0).ToString), CInt(colorObj.Text.Split(",")(1).ToString), CInt(colorObj.Text.Split(",")(2).ToString))
            Catch ex As Exception

            End Try

        Else
            Try
                boxObj.ForeColor = Color.FromArgb(CInt(colorObj.Text.Split(",")(0).ToString), CInt(colorObj.Text.Split(",")(1).ToString), CInt(colorObj.Text.Split(",")(2).ToString))
            Catch ex As Exception

            End Try

        End If

    End Sub

    Private Sub Settings_read(num_setting As Integer, box As Object)
        Dim s, split_setting As String
        Dim i As Integer
        i = FreeFile()
        Try
            FileOpen(i, "D:\Macros Settings\settings.txt", OpenMode.Input, OpenAccess.Default)
            While Not EOF(i)
                s = LineInput(i)
            End While
            FileClose(i)
            split_setting = s.Split("#")(num_setting).ToString()


            If num_setting = 2 Then
                If split_setting = 11 Then
                    RadioButton1.Checked = True
                ElseIf split_setting = "12" Then
                    RadioButton2.Checked = True
                ElseIf split_setting = "13" Then
                    RadioButton3.Checked = True
                ElseIf split_setting = "21" Then
                    RadioButton4.Checked = True
                ElseIf split_setting = "22" Then
                    RadioButton5.Checked = True
                ElseIf split_setting = "23" Then
                    RadioButton6.Checked = True
                ElseIf split_setting = "31" Then
                    RadioButton7.Checked = True
                ElseIf split_setting = "32" Then
                    RadioButton8.Checked = True
                ElseIf split_setting = "33" Then
                    RadioButton9.Checked = True
                End If


            ElseIf num_setting = 5 Or num_setting = 6 Or num_setting = 7 Then 'жирный, курсив или подчеркнутый
                If split_setting = "1" Then
                    box.Checked = True
                End If


            Else
                box.text = split_setting
            End If

        Catch ex As Exception ' действия, если не удалось открыть файл настроек
            GroupBox4.Visible = True
            TextBox1.Text = "Пожалуйста создайте папку" + Environment.NewLine + "D:\Macros Settings\" + Environment.NewLine + "Файл с настройками будет создан автоматически!"
        End Try


    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GroupBox4.Visible = False
        start_program()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Exit()
    End Sub

    Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click

        If (ColorDialog1.ShowDialog() = DialogResult.OK) Then
            Dim R, G, B As String
            R = ColorDialog1.Color.R
            G = ColorDialog1.Color.G
            B = ColorDialog1.Color.B

            TextBox2.Text = R + "," + G + "," + B
            color_set(TextBox6, TextBox2, 0)
            test_show()
        End If

    End Sub

    Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles TextBox4.Click

        If (ColorDialog1.ShowDialog() = DialogResult.OK) Then
            Dim R, G, B As String
            R = ColorDialog1.Color.R
            G = ColorDialog1.Color.G
            B = ColorDialog1.Color.B

            TextBox3.Text = R + "," + G + "," + B
            color_set(TextBox4, TextBox3, 0)
            test_show()
        End If

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked Then
            Panel1.Visible = False
        Else Panel1.Visible = True

        End If
    End Sub
End Class





