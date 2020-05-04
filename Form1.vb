Public Class Form1



    Dim fonts_name As New System.Drawing.Text.InstalledFontCollection

    Public mem_color_text(), mem_color_fill(), mem_color_head_text(), mem_color_head_fill(), mem_color_text_pipetka(), mem_color_bg_pipetka(), mem_color_text_textbox(), mem_color_fill_textbox() As Button 'массивы памяти цветов для текста и заливки + для шапки
    Public radiobtn_align(,), head_radiobtn_align(,), textbox_radiobtn_align(,) As RadioButton
    Public otladka As String
    Public timer_save As Integer = 0 'таймер для отображения что настройки сохранены
    Public call_pixcolor As Button ' служит для запоминания кнопки, которая была нажата для выбора цвета с экрана
    Public color_pixcolor As String ' служит для запоминания цвета выбранного пипеткой
    Public set_setting_table_complete As Boolean = False
    Public set_setting_head_complete As Boolean = False
    Public set_setting_textbox_complete As Boolean = False
    Public Window_view_memory As Integer ' служит для запоминания состояния вкладки "Таблица"
    Public windows_width As Integer ' ширина рабочего окна

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox2.DataSource = fonts_name.Families
        ComboBox2.DisplayMember = "Name"
        ComboBox8.DataSource = fonts_name.Families
        ComboBox8.DisplayMember = "Name"
        ComboBox11.DataSource = fonts_name.Families
        ComboBox11.DisplayMember = "Name"
        ComboBox15.DataSource = fonts_name.Families
        ComboBox15.DisplayMember = "Name"
        windows_width = 602 ' задаем ширину окна по умолчанию


        start_program()

    End Sub

    Public Sub start_program()

        mem_color_text = New Button(7) {Me.color_text_memory_1, Me.color_text_memory_2, Me.color_text_memory_3, Me.color_text_memory_4, Me.color_text_memory_5, Me.color_text_memory_6, Me.color_text_memory_7, Me.color_text_memory_8}
        mem_color_fill = New Button(7) {Me.color_fill_memory_1, Me.color_fill_memory_2, Me.color_fill_memory_3, Me.color_fill_memory_4, Me.color_fill_memory_5, Me.color_fill_memory_6, Me.color_fill_memory_7, Me.color_fill_memory_8}
        mem_color_head_text = New Button(7) {Me.color_head_text_memory_1, Me.color_head_text_memory_2, Me.color_head_text_memory_3, Me.color_head_text_memory_4, Me.color_head_text_memory_5, Me.color_head_text_memory_6, Me.color_head_text_memory_7, Me.color_head_text_memory_8}
        mem_color_head_fill = New Button(7) {Me.color_head_fill_memory_1, Me.color_head_fill_memory_2, Me.color_head_fill_memory_3, Me.color_head_fill_memory_4, Me.color_head_fill_memory_5, Me.color_head_fill_memory_6, Me.color_head_fill_memory_7, Me.color_head_fill_memory_8}
        mem_color_text_pipetka = New Button(7) {Me.color_text_pipetka_memory_1, Me.color_text_pipetka_memory_2, Me.color_text_pipetka_memory_3, Me.color_text_pipetka_memory_4, Me.color_text_pipetka_memory_5, Me.color_text_pipetka_memory_6, Me.color_text_pipetka_memory_7, Me.color_text_pipetka_memory_8}
        mem_color_bg_pipetka = New Button(7) {Me.color_bg_pipetka_memory_1, Me.color_bg_pipetka_memory_2, Me.color_bg_pipetka_memory_3, Me.color_bg_pipetka_memory_4, Me.color_bg_pipetka_memory_5, Me.color_bg_pipetka_memory_6, Me.color_bg_pipetka_memory_7, Me.color_bg_pipetka_memory_8}
        mem_color_text_textbox = New Button(7) {Me.tb_t_1, Me.tb_t_2, Me.tb_t_3, Me.tb_t_4, Me.tb_t_5, Me.tb_t_6, Me.tb_t_7, Me.tb_t_8}
        mem_color_fill_textbox = New Button(7) {Me.tb_f_1, Me.tb_f_2, Me.tb_f_3, Me.tb_f_4, Me.tb_f_5, Me.tb_f_6, Me.tb_f_7, Me.tb_f_8}

        radiobtn_align = New RadioButton(2, 2) {{Me.RadioButton1, Me.RadioButton2, Me.RadioButton3}, {Me.RadioButton4, Me.RadioButton5, Me.RadioButton6}, {Me.RadioButton7, Me.RadioButton8, Me.RadioButton9}}
        head_radiobtn_align = New RadioButton(2, 2) {{Me.Head_RadioButton1, Me.Head_RadioButton2, Me.Head_RadioButton3}, {Me.Head_RadioButton4, Me.Head_RadioButton5, Me.Head_RadioButton6}, {Me.Head_RadioButton7, Me.Head_RadioButton8, Me.Head_RadioButton9}}
        textbox_radiobtn_align = New RadioButton(2, 2) {{Me.Textbox_RadioButton1, Me.Textbox_RadioButton2, Me.Textbox_RadioButton3}, {Me.Textbox_RadioButton4, Me.Textbox_RadioButton5, Me.Textbox_RadioButton6}, {Me.Textbox_RadioButton7, Me.Textbox_RadioButton8, Me.Textbox_RadioButton9}}


        GroupBox1.TabStop = False
        Window_view(1)
        Settings_read(0, ComboBox2, 1) ' установка названия шрифта
        Settings_read(1, ComboBox1, 1) ' установка размера шрифта
        Settings_read(2, RadioButton5, 1) 'установка выравнивания
        Settings_read(3, ComboBox4, 1) ' установка отступов верх/низ
        Settings_read(4, ComboBox3, 1) ' установка отступов лево/право
        Settings_read(5, CheckBox3, 1) 'жирный
        Settings_read(6, CheckBox2, 1) 'курсив
        Settings_read(7, CheckBox1, 1) 'подчеркнутый
        Settings_read(8, TextBox2, 1) 'установка цвета шрифта
        Settings_read(9, TextBox3, 1) 'установка цвета фона ячейки
        Settings_read(10, CheckBox5, 1) 'использовать ж, к, подчеркнутый?
        Settings_read(11, CheckBox4, 1) 'проверка настроек шапки надо или нет
        Settings_read(12, CheckBox10, 1) 'проверка настроек ячеек надо или нет
        Settings_read(14, CheckBox15, 1) 'проверка нужно ли красить текст
        Settings_read(15, CheckBox16, 1) 'нужно ли красить фон
        Settings_read(13, ComboBox10, 1) 'как выравнивать саму таблицу (по ширине, по содержимому, зафискисровать, не выравнивать)
        color_set(Button10, TextBox3.Text, 0) ' красим образец цвета фона 
        color_set(Button9, TextBox2.Text, 0) ' красим образец цвета текста 
        color_memory_set(1) 'заносим цвета из памяти на панель для текста
        color_memory_set(2) 'заносим цвета из памяти на панель для фона


        set_setting_table_complete = True

        test_show(1)
    End Sub

    Private Sub TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected

        If TabControl1.SelectedIndex = 0 Then 'Если выбрана вкладка таблица
            Me.TopMost = False
            Window_view(Window_view_memory)
        End If

        If TabControl1.SelectedIndex = 1 Then 'Если выбрана пипетка
            Me.TopMost = True
            Label19.Visible = False
            Button3.Visible = False
            Me.Height = 249
            Me.Width = 300
            set_controlTab_size(210) 'устанавливает размер окна с вкладками


            color_set(Button62, TextBox9.Text, 0) ' красим образец цвета фона 
            color_set(Button30, TextBox8.Text, 0) ' красим образец цвета текста 
            color_memory_set(7) 'заносим цвета из памяти на панель для текста
            color_memory_set(8) 'заносим цвета из памяти на панель для фона
            Button62.BackColor = color_text_pipetka_memory_1.BackColor
            Button30.BackColor = color_bg_pipetka_memory_1.BackColor
            TextBox9.Text = CStr(Button62.BackColor.R) + "," + CStr(Button62.BackColor.G) + "," + CStr(Button62.BackColor.B)
            TextBox8.Text = CStr(Button30.BackColor.R) + "," + CStr(Button30.BackColor.G) + "," + CStr(Button30.BackColor.B)
        End If

        If TabControl1.SelectedIndex = 2 Then 'Если выбрана вкладка "текстовый блок"
            Label19.Visible = False
            Button3.Visible = False
            Me.Height = 249
            Me.Width = 530
            set_controlTab_size(210) 'устанавливает размер окна с вкладками
            Me.TopMost = False
            Button3.Visible = True
            Button3.Location = New Point(10, 170)


            Settings_textbox_read(0, ComboBox15) ' установка названия шрифта
            Settings_textbox_read(1, ComboBox16) ' установка размера шрифта
            Settings_textbox_read(2, Textbox_RadioButton5) 'установка выравнивания
            Settings_textbox_read(3, CheckBox21) 'жирный
            Settings_textbox_read(4, CheckBox20) 'курсив
            Settings_textbox_read(5, CheckBox19) 'подчеркнутый
            Settings_textbox_read(6, TextBox11) 'установка цвета шрифта
            Settings_textbox_read(7, TextBox10) 'установка цвета фона ячейки
            color_set(Button84, TextBox11.Text, 0) ' красим образец цвета фона 
            color_set(Button80, TextBox10.Text, 0) ' красим образец цвета текста 
            color_memory_set(9) 'заносим цвета из памяти на панель для текста
            color_memory_set(10) 'заносим цвета из памяти на панель для фона

            set_setting_textbox_complete = True

        End If

    End Sub


    Private Sub Format_head_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged 'если пользователь хочет задать формат шапки таблицы
        head_set()
        apply_changes(1)
    End Sub

    Private Sub Format_cells_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged 'формат выделенных ячеек
        cells_set()
        apply_changes(3)
    End Sub


    Private Sub head_set() 'функция, которая разворачивает окно программы для редактирование шапки, если это необходимо
        If CheckBox4.Checked = True Then 'показать настройки шапки
            CheckBox4.TabStop = False
            GroupBox5.TabStop = False
            Settings_read(0, ComboBox8, 2) ' установка названия шрифта
            Settings_read(1, ComboBox5, 2) ' установка размера шрифта
            Settings_read(8, TextBox5, 2) 'установка цвета шрифта
            Settings_read(9, TextBox4, 2) 'установка цвета фона ячейки
            Settings_read(2, Head_RadioButton5, 2) 'установка выравнивания
            Settings_read(3, ComboBox7, 2) ' установка отступов верх/низ
            Settings_read(4, ComboBox6, 2) ' установка отступов лево/право
            Settings_read(5, CheckBox9, 2) 'жирный
            Settings_read(6, CheckBox8, 2) 'курсив
            Settings_read(7, CheckBox7, 2) 'подчеркнутый
            Settings_read(11, ComboBox9, 2) 'количество строк шапки
            color_set(Button14, TextBox5.Text, 0) ' красим образец цвета текста
            color_set(Button12, TextBox4.Text, 0) ' красим образец цвета фона 
            color_memory_set(3) 'установка цветов цвета в память для текста
            color_memory_set(4) 'установка цветов в память для фона
            Settings_read(12, CheckBox17, 2) 'проверка нужно ли красить текст
            Settings_read(13, CheckBox18, 2) 'нужно ли красить фон
            set_setting_head_complete = True


            If CheckBox10.Checked = True Then 'проверяем стоит ли галочка "формат ячеек"
                Window_view(4)
            Else
                Window_view(2)
            End If

        Else 'скрыть настройки шапки

            If CheckBox10.Checked = True Then 'проверяем стоит ли галочка "формат ячеек"
                Window_view(3)
            Else
                Window_view(1)
            End If
            CheckBox4.TabStop = True

        End If

        test_show(2)
    End Sub




    Private Sub cells_set() 'функция, которая разворачивает окно программы для редактирование формата ячеек
        If CheckBox10.Checked = True Then
            CheckBox10.TabStop = False

            GroupBox9.TabStop = False

            If CheckBox4.Checked Then
                Window_view(4) 'отображаем все три окошка
            Else
                Window_view(3) 'отображаем 2 окошка основное+для ячеек
            End If
            '    Settings_read(0, ComboBox8, 3) ' установка названия шрифта
            '    Settings_read(1, ComboBox5, 3) ' установка размера шрифта
            '    Settings_read(8, TextBox5, 3) 'установка цвета шрифта
            '    Settings_read(9, TextBox4, 3) 'установка цвета фона ячейки

            '    Settings_read(2, Head_RadioButton5, 3) 'установка выравнивания
            '    Settings_read(3, ComboBox7, 3) ' установка отступов верх/низ
            '    Settings_read(4, ComboBox6, 3) ' установка отступов лево/право
            '    Settings_read(5, CheckBox9, 3) 'жирный
            '    Settings_read(6, CheckBox8, 3) 'курсив
            '    Settings_read(7, CheckBox7, 3) 'подчеркнутый
            '    Settings_read(11, ComboBox9, 3) 'количество строк шапки
            '    color_set(Button14, TextBox5.Text, 0) ' красим образец цвета текста
            '    color_set(Button12, TextBox4.Text, 0) ' красим образец цвета фона 
            '    color_memory_set(5) 'установка цветов цвета в память для текста
            '    color_memory_set(6) 'установка цветов в память для фона
            '    set_setting_head_complete = True
            '    test_show(3)

        Else
            '  
            '    CheckBox4.TabStop = True
            ' 

            If CheckBox4.Checked Then
                Window_view(2) 'если была снята галочка с "формат. ячеек", но установлена галочка "формат. шапки", то отображаем главн. панель +формат шапки
            Else
                Window_view(1) 'как и выше, но формат. шапки не используется, поэтому оставляем только одно окно
            End If

            CheckBox10.TabStop = True
            Panel5.Visible = False
            GroupBox9.TabStop = True
            Label31.Visible = False


        End If

    End Sub

    Private Sub Settings_save(goal As Integer)
        Dim dir As String 'путь к файлу который нужно сохранить
        Dim radiobtn_array(3, 3) As RadioButton 'массив в который запишется массив радиокнопок выравнивания текста
        Dim checkbox_bold_show, checkbox_italic_show, checbox_underline_show, save_format, add_head, set_color_text, set_color_fill As CheckBox
        Dim font, size, top_bottom, left_right, count_rows_head, autosize As ComboBox
        Dim s As String
        Dim text_align As String = 0
        Dim bold_status, italic_status, underline_status, format_status, head_status As String 'переменные для статуса жирного, курсива и подчеркнутого
        Dim set_color_text_status, set_color_fill_status As String

        Dim color_text, color_fill As TextBox

        If goal = 1 Then
            dir = "D:\Macros Settings\settings.txt"
            radiobtn_array = radiobtn_align
            checkbox_bold_show = CheckBox3
            checkbox_italic_show = CheckBox2
            checbox_underline_show = CheckBox1
            save_format = CheckBox5
            font = ComboBox2
            size = ComboBox1
            top_bottom = ComboBox4
            left_right = ComboBox3
            color_text = TextBox2
            color_fill = TextBox3
            add_head = CheckBox4
            set_color_text = CheckBox15
            set_color_fill = CheckBox16


            If add_head.Checked Then
                head_status = "1"
            Else head_status = "0"
            End If

        ElseIf goal = 2 Then
            dir = "D:\Macros Settings\head_settings.txt"
            radiobtn_array = head_radiobtn_align
            checkbox_bold_show = CheckBox9
            checkbox_italic_show = CheckBox8
            checbox_underline_show = CheckBox7
            save_format = CheckBox6
            font = ComboBox8
            size = ComboBox5
            top_bottom = ComboBox7
            left_right = ComboBox6
            color_text = TextBox5
            color_fill = TextBox4
            count_rows_head = ComboBox9
            set_color_text = CheckBox17
            set_color_fill = CheckBox18


        ElseIf goal = 4 Then
            dir = "D:\Macros Settings\textbox_settings.txt"
            radiobtn_array = textbox_radiobtn_align
            checkbox_bold_show = CheckBox21
            checkbox_italic_show = CheckBox20
            checbox_underline_show = CheckBox19
            font = ComboBox15
            size = ComboBox16
            color_text = TextBox11
            color_fill = TextBox10

        End If
        FileOpen(1, dir, OpenMode.Output)

        For i = 0 To 2
            For y = 0 To 2
                If radiobtn_array(i, y).Checked Then
                    text_align = CStr(i + 1) + CStr(y + 1)
                End If
            Next
        Next

        If checbox_underline_show.Checked Then
            underline_status = "1"
        Else underline_status = "0"
        End If
        If checkbox_italic_show.Checked Then
            italic_status = "1"
        Else italic_status = "0"
        End If
        If checkbox_bold_show.Checked Then
            bold_status = "1"
        Else bold_status = "0"
        End If

        If goal <> 4 Then
            If save_format.Checked Then
                format_status = "1"
            Else
                format_status = "0"
            End If
            If set_color_text.Checked Then
                set_color_text_status = "1"
            Else set_color_text_status = "0"
            End If

            If set_color_fill.Checked Then
                set_color_fill_status = "1"
            Else set_color_fill_status = "0"
            End If
        End If

        If goal = 1 Then
            s = font.Text + "#" + size.Text + "#" + text_align + "#" + top_bottom.Text + "#" + left_right.Text + "#" + bold_status + "#" + italic_status + "#" + underline_status + "#" + color_text.Text + "#" + color_fill.Text + "#" + format_status + "#" + head_status + "#" + "доп.формат" + "#" + CStr(ComboBox10.SelectedIndex) + "#" + set_color_text_status + "#" + set_color_fill_status

        ElseIf goal = 2 Then
            s = font.Text + "#" + size.Text + "#" + text_align + "#" + top_bottom.Text + "#" + left_right.Text + "#" + bold_status + "#" + italic_status + "#" + underline_status + "#" + color_text.Text + "#" + color_fill.Text + "#" + format_status + "#" + ComboBox9.Text + "#" + set_color_text_status + "#" + set_color_fill_status


        ElseIf goal = 4 Then
            s = font.Text + "#" + size.Text + "#" + text_align + "#" + bold_status + "#" + italic_status + "#" + underline_status + "#" + color_text.Text + "#" + color_fill.Text
        End If

        Print(1, s)
        FileClose(1)
        Timer2.Enabled = True

    End Sub


    Private Sub apply_changes(goal As Integer)

        If goal = 1 And set_setting_table_complete Then
            test_show(goal)
            Settings_save(goal)
        ElseIf goal = 2 And set_setting_head_complete Then
            test_show(goal)
            Settings_save(goal)

        ElseIf goal = 4 And set_setting_textbox_complete Then 'измнения в текстовом блоке
            ' test_show(goal) 'для текстового блока нет тестового отображения
            Settings_save(goal)

        End If

    End Sub

    Private Sub Button_color_Press_down(sender As Object, e As EventArgs) Handles Button6.MouseDown, Button4.MouseDown, Button34.MouseDown, Button33.MouseDown, Button82.MouseDown, Button81.MouseDown, Button103.MouseDown, Button102.MouseDown

        Dim BM As New Bitmap(My.Resources.pipetka)
        Dim Ico As Icon = Icon.FromHandle(BM.GetHicon)
        sender.cursor = New Cursor(Ico.Handle)
        sender.Image = Nothing

        If sender.name = Button6.Name Then
            call_pixcolor = Button6
            Panel3.Location = New Point(343, 28)
        ElseIf sender.name = Button4.Name Then
            call_pixcolor = Button4
            Panel3.Location = New Point(343, 28)
        ElseIf sender.name = Button34.Name Then
            call_pixcolor = Button34
            Panel3.Location = New Point(343, 245)
        ElseIf sender.name = Button33.Name Then
            call_pixcolor = Button33
            Panel3.Location = New Point(343, 245)
        ElseIf sender.name = Button82.Name Then
            call_pixcolor = Button82
            Panel3.Location = New Point(20, 1)
        ElseIf sender.name = Button81.Name Then
            call_pixcolor = Button81
            Panel3.Location = New Point(20, 1)
        ElseIf sender.name = Button103.Name Then
            call_pixcolor = Button103
            Panel3.Location = New Point(255, 1)
        ElseIf sender.name = Button102.Name Then
            call_pixcolor = Button102
            Panel3.Location = New Point(255, 1)
        End If
        Timer1.Enabled = True
    End Sub


    Private Sub Button_color_Press_up(sender As Object, e As EventArgs) Handles Button6.MouseUp, Button4.MouseUp, Button34.MouseUp, Button33.MouseUp, Button82.MouseUp, Button81.MouseUp, Button103.MouseUp, Button102.MouseUp

        set_color_getpixel()
        Timer1.Enabled = False
        Panel3.Visible = False
        sender.cursor = Cursors.Hand
        sender.Image = My.Resources.pipetka_btn

    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click 'При нажатии на кнопку "тест"
        If sender.name = Button7.Name Then
            test_show(1)
        Else
            test_show(2)
        End If
    End Sub

    Private Sub test_show(goal As Integer)
        Dim text_test As String = "Пример"
        Dim button_test_show As Button
        Dim textbox_text_show, textbox_fill_show As TextBox
        Dim checkbox_bold_show, checkbox_italic_show, checbox_underline_show, format_status As CheckBox
        Dim fonts_show, size_font_show As ComboBox
        Dim radiobutton_array(3, 3) As RadioButton
        If goal = 1 Then
            button_test_show = Button7
            textbox_text_show = TextBox2
            textbox_fill_show = TextBox3
            checkbox_bold_show = CheckBox3
            checkbox_italic_show = CheckBox2
            checbox_underline_show = CheckBox1
            fonts_show = ComboBox2
            size_font_show = ComboBox1
            radiobutton_array = radiobtn_align
            format_status = CheckBox5
        ElseIf goal = 2 Then
            button_test_show = Button32
            textbox_text_show = TextBox5
            textbox_fill_show = TextBox4
            checkbox_bold_show = CheckBox9
            checkbox_italic_show = CheckBox8
            checbox_underline_show = CheckBox7
            fonts_show = ComboBox8
            size_font_show = ComboBox5
            radiobutton_array = head_radiobtn_align
            format_status = CheckBox6
        End If
        button_test_show.Text = text_test
        ' ниже конструкция elseif для позиционирования текста
        If radiobutton_array(0, 0).Checked Then
            button_test_show.TextAlign = ContentAlignment.TopLeft
        ElseIf radiobutton_array(0, 1).Checked Then
            button_test_show.TextAlign = ContentAlignment.TopCenter
        ElseIf radiobutton_array(0, 2).Checked Then
            button_test_show.TextAlign = ContentAlignment.TopRight
        ElseIf radiobutton_array(1, 0).Checked Then
            button_test_show.TextAlign = ContentAlignment.MiddleLeft
        ElseIf radiobutton_array(1, 1).Checked Then
            button_test_show.TextAlign = ContentAlignment.MiddleCenter
        ElseIf radiobutton_array(1, 2).Checked Then
            button_test_show.TextAlign = ContentAlignment.MiddleRight
        ElseIf radiobutton_array(2, 0).Checked Then
            button_test_show.TextAlign = ContentAlignment.BottomLeft
        ElseIf radiobutton_array(2, 1).Checked Then
            button_test_show.TextAlign = ContentAlignment.BottomCenter
        ElseIf radiobutton_array(2, 2).Checked Then
            button_test_show.TextAlign = ContentAlignment.BottomRight
        End If

        'тут устанавливается цвет шрифта и фона, а также шрифт и его размер
        color_set(button_test_show, textbox_text_show.Text, 1)
        color_set(button_test_show, textbox_fill_show.Text, 0)

        Dim TempFont = button_test_show.Font
        Dim FontStyle1 = FontStyle.Regular
        If format_status.Checked = False Then
            If checkbox_italic_show.Checked Then FontStyle1 = FontStyle1 Or FontStyle.Italic
            If checbox_underline_show.Checked Then FontStyle1 = FontStyle1 Or FontStyle.Underline
            If checkbox_bold_show.Checked Then FontStyle1 = FontStyle1 Or FontStyle.Bold
        End If
        button_test_show.Font = New System.Drawing.Font(fonts_show.Text, CInt(size_font_show.Text), FontStyle1)
    End Sub


    Private Sub color_set(boxObj As Object, colorObj As String, goal As Integer) 'Функция применения цвета, boxObj - что красим, colorObj - каким цветом, goal - текст или фон
        If goal = 0 Then
            Try
                boxObj.BackColor = Color.FromArgb(CInt(colorObj.Split(",")(0).ToString), CInt(colorObj.Split(",")(1).ToString), CInt(colorObj.Split(",")(2).ToString))
            Catch ex As Exception

            End Try

        Else
            Try
                boxObj.ForeColor = Color.FromArgb(CInt(colorObj.Split(",")(0).ToString), CInt(colorObj.Split(",")(1).ToString), CInt(colorObj.Split(",")(2).ToString))
            Catch ex As Exception

            End Try

        End If

    End Sub

    Private Sub Settings_read(num_setting As Integer, box As Object, goal As Integer) 'открытие файла с настройками и чтение настроек 

        Dim dir As String
        Dim align_radiobox(3, 3) As RadioButton

        If goal = 1 Then 'путь к файлу настроек для всей таблицы
            align_radiobox = radiobtn_align
            dir = "D:\Macros Settings\settings.txt"
        ElseIf goal = 2 Then
            'путь к файлу настроек для шапки таблицы
            align_radiobox = head_radiobtn_align
            dir = "D:\Macros Settings\head_settings.txt"
        End If
        Dim s, split_setting As String

        s = file_open(dir)

        split_setting = s.Split("#")(num_setting).ToString()

        If num_setting = 2 Then 'проверка какой радиобокс выравнивания установлен
            'обращаемся к элементу массива радиобаттанов с помощью двух цифр из файла настроек, первая цифра номер строки, вторая номер столбца
            'Microsoft.VisualBasic.Val преобразует char элемент строки в цифру
            align_radiobox(Microsoft.VisualBasic.Val(split_setting.Chars(0)) - 1, Microsoft.VisualBasic.Val(split_setting.Chars(1)) - 1).Checked = True

        ElseIf num_setting = 5 Or num_setting = 6 Or num_setting = 7 Then 'жирный, курсив или подчеркнутый
            If split_setting = "1" Then
                box.Checked = True
            End If

        ElseIf num_setting = 13 And goal = 1 Then 'какое выравнивание таблицы стоит?
            box.SelectedIndex = split_setting



        ElseIf num_setting = 11 And goal = 1 Then 'стоит ли галочка на "делать шапку"?
            If split_setting = "1" Then
                box.Checked = True
            End If


        ElseIf num_setting = 14 And goal = 1 Then 'стоит ли галочка на "цвет текста" для всей таблицы?
            If split_setting = "1" Then
                box.Checked = True
            End If

        ElseIf num_setting = 15 And goal = 1 Then 'стоит ли галочка на "цвет фона" для всей таблицы?
            If split_setting = "1" Then
                box.Checked = True
            End If

        ElseIf num_setting = 12 And goal = 2 Then 'стоит ли галочка на "цвет текста" для всей таблицы?
            If split_setting = "1" Then
                box.Checked = True
            End If

        ElseIf num_setting = 13 And goal = 2 Then 'стоит ли галочка на "цвет фона" для всей таблицы?
            If split_setting = "1" Then
                box.Checked = True
            End If

        ElseIf num_setting = 10 Then
            If split_setting = "1" Then
                box.Checked = True
            Else
                box.Checked = False
            End If



        Else
            box.Text = split_setting
        End If

    End Sub 'конец чтения настроек из файла



    Private Sub Settings_textbox_read(num_setting As Integer, box As Object) 'открытие файла с настройками и чтение настроек 

        Dim dir As String
        Dim align_radiobox(3, 3) As RadioButton
        align_radiobox = textbox_radiobtn_align

        Dim s, split_setting As String

        s = file_open("D:\Macros Settings\textbox_settings.txt")

        split_setting = s.Split("#")(num_setting).ToString()

        If num_setting = 2 Then 'проверка какой радиобокс выравнивания установлен
            'обращаемся к элементу массива радиобаттанов с помощью двух цифр из файла настроек, первая цифра номер строки, вторая номер столбца
            'Microsoft.VisualBasic.Val преобразует char элемент строки в цифру
            align_radiobox(Microsoft.VisualBasic.Val(split_setting.Chars(0)) - 1, Microsoft.VisualBasic.Val(split_setting.Chars(1)) - 1).Checked = True

        ElseIf num_setting = 3 Or num_setting = 4 Or num_setting = 5 Then 'жирный, курсив или подчеркнутый
            If split_setting = "1" Then
                box.Checked = True
            End If

        Else
            box.Text = split_setting
        End If

    End Sub 'конец чтения настроек из файла




    Function file_open(file_dir As String) As String
        Dim i As Integer
        Dim s As String
        otladka += "- Запрос на чтение файла: " + file_dir + " -"
        Try


            i = FreeFile()

            FileOpen(i, file_dir, OpenMode.Input, OpenAccess.Default)
            While Not EOF(i)
                s = LineInput(i)
            End While
            FileClose(i)

            file_open = s
            otladka += "- файл " + file_dir + " открыт успешною. в функцию вернулась строка '" + s + "' -"

        Catch ex As Exception ' действия, если не удалось открыть файл настроек
            My.Computer.FileSystem.CreateDirectory("D:\Macros Settings\style")
            s = "Arial#8#11#0.05#0.10#0#0#0#0,0,0#255,255,255#0#1#1#1#1#1"
            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\settings.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)
            s = "Arial#8#22#0.10#0.10#1#0#0#0,0,0#255,255,255#0#1"
            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\head_settings.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)

            s = "255,255,255#255,255,255#255,255,255#255,255,255#255,255,255#255,255,255#255,255,255#255,255,255"

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\color_text_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\color_fill_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\color_head_text_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\color_head_fill_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\pipetka_text_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\pipetka_bg_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\textbox_text_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\textbox_fill_memory.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, s)
            FileClose(i)

            i = FreeFile()
            FileOpen(i, "D:\Macros Settings\style\style_name.txt", OpenMode.Output, OpenAccess.Default)
            Print(i, "Стиль 1" + Environment.NewLine + "Стиль 2" + Environment.NewLine + "Стиль 3" + Environment.NewLine + "Стиль 4" + Environment.NewLine + "Стиль 5" + Environment.NewLine + "Стиль 6" + Environment.NewLine + "Стиль 7" + Environment.NewLine + "Стиль 8")
            FileClose(i)

            s = file_open(file_dir)
            file_open = s
        End Try


    End Function
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, Button10.Click, Button14.Click, Button12.Click, Button62.Click, Button30.Click, Button84.Click, Button80.Click 'окно выбора цвета для текста
        Dim source_color_code As TextBox
        Dim goal As Integer

        If sender.Name = Button9.Name Then 'выбор цвета для текста таблицы
            source_color_code = TextBox2
            goal = 1
        ElseIf sender.Name = Button10.Name Then 'выбор цвета для фона таблицы
            source_color_code = TextBox3
            goal = 2
        ElseIf sender.name = Button14.Name Then 'выбор цвета для текста шапки
            source_color_code = TextBox5
            goal = 3
        ElseIf sender.name = Button12.Name Then 'выбор цвета для фона шапки
            source_color_code = TextBox4
            goal = 4
        ElseIf sender.name = Button62.Name Then 'Пипетка для текста
            source_color_code = TextBox9
            goal = 7
        ElseIf sender.name = Button30.Name Then 'Пипетка для фона
            source_color_code = TextBox8
            goal = 8
        ElseIf sender.name = Button84.Name Then 'Пипетка для текста
            source_color_code = TextBox11
            goal = 9
        ElseIf sender.name = Button80.Name Then 'Пипетка для фона
            source_color_code = TextBox10
            goal = 10

        End If

        If (ColorDialog1.ShowDialog() = DialogResult.OK) Then
            Dim R, G, B As String
            R = ColorDialog1.Color.R
            G = ColorDialog1.Color.G
            B = ColorDialog1.Color.B

            source_color_code.Text = R + "," + G + "," + B
            color_set(sender, source_color_code.Text, 0)
            upd_and_save_mem(goal)
            If goal < 5 Then
                apply_changes(1)
                apply_changes(2)
                'apply_changes(3) пока нет формирования отдельных ячеек это надо закоментить
            End If

        End If

    End Sub



    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged, CheckBox6.CheckedChanged 'применять ли жирный, курсив и подчеркнутый для макроса?
        Dim num_panel As Panel
        Dim goal As Integer
        If sender.Name = CheckBox5.Name Then
            num_panel = Panel1
            goal = 1
        ElseIf sender.Name = CheckBox6.Name Then
            num_panel = Panel2
            goal = 2
        End If

        If sender.Checked Then
            num_panel.Visible = False
        Else num_panel.Visible = True
        End If
        apply_changes(goal)
    End Sub


    Private Sub CheckBox15_16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged, CheckBox16.CheckedChanged 'применять ли заливку ячейки и цвет текста для таблички
        apply_changes(1)
    End Sub

    Private Sub CheckBox17_18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged, CheckBox18.CheckedChanged 'применять ли заливку ячейки и цвет текста для шапки
        apply_changes(2)
    End Sub





    Private Sub color_memory_set(goal As Integer) '1 - для шрифта, 2- для фона 3 - для шрифта шапки 4-для фона шапки
        Dim colors_array(7) As String
        Dim colors_mem_button(7) As Button
        'запускаем функцию считывания цветов для шрифта или фона
        colors_array = color_memory_read(goal)

        'Для нужного массива объектов устанавливаем цвета в цикле
        If goal = 1 Then 'цвет текста таблицы
            colors_mem_button = mem_color_text 'копируем массив цветов текста таблицы
        ElseIf goal = 2 Then 'цвет фона таблицы
            colors_mem_button = mem_color_fill 'копируем массив цветов фона таблицы
        ElseIf goal = 3 Then 'цвет текста шапки
            colors_mem_button = mem_color_head_text 'копируем массив цветов текста таблицы
        ElseIf goal = 4 Then 'цвет фона шапки
            colors_mem_button = mem_color_head_fill 'копируем массив цветов фона таблицы
        ElseIf goal = 7 Then 'пипетка цвет текста
            colors_mem_button = mem_color_text_pipetka
            'копируем массив цветов текста таблицы
        ElseIf goal = 8 Then 'пипетка цвет заливки
            colors_mem_button = mem_color_bg_pipetka 'копируем массив цветов фона таблицы
        ElseIf goal = 9 Then 'пипетка цвет текста
            colors_mem_button = mem_color_text_textbox
            'копируем массив цветов текста таблицы
        ElseIf goal = 10 Then 'пипетка цвет заливки
            colors_mem_button = mem_color_fill_textbox 'копируем массив цветов фона таблицы
        End If

        For indexA = 0 To 7
            color_set(colors_mem_button(indexA), colors_array(indexA), 0)
        Next indexA

    End Sub
    Function color_memory_read(color_for As Integer) As String() 'считываем сохраненные цвета из файлов (1 - для текста, 2- для фона)
        'если 1, то читаем для текста, если 2, то читаем для фона
        Dim colors As String
        Dim colors_array(7) As String

        If color_for = 1 Then
            colors = file_open("D:\Macros Settings\color_text_memory.txt")
        ElseIf color_for = 2 Then
            colors = file_open("D:\Macros Settings\color_fill_memory.txt")
        ElseIf color_for = 3 Then
            colors = file_open("D:\Macros Settings\color_head_text_memory.txt")
        ElseIf color_for = 4 Then
            colors = file_open("D:\Macros Settings\color_head_fill_memory.txt")


        ElseIf color_for = 7 Then
            colors = file_open("D:\Macros Settings\pipetka_text_memory.txt")

        ElseIf color_for = 8 Then
            colors = file_open("D:\Macros Settings\pipetka_bg_memory.txt")

        ElseIf color_for = 9 Then
            colors = file_open("D:\Macros Settings\textbox_text_memory.txt")

        ElseIf color_for = 10 Then
            colors = file_open("D:\Macros Settings\textbox_fill_memory.txt")
        End If

        For indexA = 0 To 7
            colors_array(indexA) = colors.Split("#")(indexA).ToString()
        Next indexA

        color_memory_read = colors_array

    End Function


    Private Sub choose_from_memory(goal As Integer, color As Color) 'действия при нажатии на цвета в памяти
        Dim R, G, B As String
        R = color.R
        G = color.G
        B = color.B

        If goal = 1 Then 'цвет текста всей таблицы из памяти
            TextBox2.Text = R + "," + G + "," + B
            Button9.BackColor = color
            upd_and_save_mem(goal)
            apply_changes(1)
        ElseIf goal = 2 Then 'цвет фона всей таблицы из памяти
            TextBox3.Text = R + "," + G + "," + B
            Button10.BackColor = color
            upd_and_save_mem(goal)
            apply_changes(1)
        ElseIf goal = 3 Then 'цвет текста шапки таблицы из памяти
            TextBox5.Text = R + "," + G + "," + B
            Button14.BackColor = color
            upd_and_save_mem(goal)
            apply_changes(2)
        ElseIf goal = 4 Then 'цвет фона шапки таблицы из памяти
            TextBox4.Text = R + "," + G + "," + B
            Button12.BackColor = color
            upd_and_save_mem(goal)
            apply_changes(2)
        ElseIf goal = 7 Then 'цвет текста пипетка
            TextBox9.Text = R + "," + G + "," + B
            Button62.BackColor = color
            upd_and_save_mem(goal)
        ElseIf goal = 8 Then 'цвет фона пипетка
            TextBox8.Text = R + "," + G + "," + B
            Button30.BackColor = color
            upd_and_save_mem(goal)
        ElseIf goal = 9 Then 'цвет текста текстовый блок
            TextBox11.Text = R + "," + G + "," + B
            Button84.BackColor = color
            upd_and_save_mem(goal)
            apply_changes(4)
        ElseIf goal = 10 Then 'цвет фона текстовый блок
            TextBox10.Text = R + "," + G + "," + B
            Button80.BackColor = color
            upd_and_save_mem(goal)
            apply_changes(4)
        End If


    End Sub

    Private Sub upd_and_save_mem(goal As Integer) 'получаем цель - текст или фон

        'имеем массив цветов в памяти
        'когда происходит изменение цвета, мы должны понять, если ли такой цвет в памяти, если да, то просто переместить его на первое место
        'если нет, то перезаписть массив со сдвигом на 1, а на первое место записать новый цвет
        'после чего сохранить массив в файл
        Dim color, dir As String
        Dim color_array(7) As Button
        Dim obj As TextBox
        Dim flag As Integer = 0 'найден цвет или нет? 0- не найден, 1 - найден


        If goal = 1 Then
            dir = "D:\Macros Settings\color_text_memory.txt"
            color_array = mem_color_text
            obj = TextBox2
        ElseIf goal = 2 Then
            dir = "D:\Macros Settings\color_fill_memory.txt"
            color_array = mem_color_fill
            obj = TextBox3
        ElseIf goal = 3 Then
            dir = "D:\Macros Settings\color_head_text_memory.txt"
            color_array = mem_color_head_text
            obj = TextBox5
        ElseIf goal = 4 Then
            dir = "D:\Macros Settings\color_head_fill_memory.txt"
            color_array = mem_color_head_fill
            obj = TextBox4


        ElseIf goal = 7 Then
            dir = "D:\Macros Settings\pipetka_text_memory.txt"
            color_array = mem_color_text_pipetka
            obj = TextBox9
        ElseIf goal = 8 Then
            dir = "D:\Macros Settings\pipetka_bg_memory.txt"
            color_array = mem_color_bg_pipetka
            obj = TextBox8
        ElseIf goal = 9 Then
            dir = "D:\Macros Settings\textbox_text_memory.txt"
            color_array = mem_color_text_textbox
            obj = TextBox11
        ElseIf goal = 10 Then
            dir = "D:\Macros Settings\textbox_fill_memory.txt"
            color_array = mem_color_fill_textbox
            obj = TextBox10


        End If

        For i = 0 To 7
            color = CStr(color_array(i).BackColor.R) + "," + CStr(color_array(i).BackColor.G) + "," + CStr(color_array(i).BackColor.B)

            If obj.Text = color And flag = 0 Then 'если цвет совпал с цветом из памяти первый раз (flag = 0)
                flag = 1
                Dim buff_color As Color
                buff_color = color_array(i).BackColor
                For j = i To 1 Step -1
                    color_array(j).BackColor = color_array(j - 1).BackColor
                Next j
                color_array(0).BackColor = buff_color
            End If
        Next i

        If flag = 0 Then
            For i = 7 To 1 Step -1
                color_array(i).BackColor = color_array(i - 1).BackColor
            Next
            color_set(color_array(0), obj.Text, 0)
        End If

        'сохранение:
        Dim file_num As Integer = FreeFile()
        FileOpen(file_num, dir, OpenMode.Output)
        Dim s As String

        For i = 0 To 7

            s += CStr(color_array(i).BackColor.R) + "," + CStr(color_array(i).BackColor.G) + "," + CStr(color_array(i).BackColor.B)
            If i < 7 Then
                s += "#"
            End If
        Next i
        Print(file_num, s)
        FileClose(file_num)
    End Sub

    Function control_imput_color_rgb(goal As Integer) As Boolean ' Функция проверки правильности ввода в текстовое поле кода цвета
        Dim code As String = ""
        Dim text_code As TextBox
        Dim R, G, B As Integer
        If goal = 1 Then
            code = TextBox2.Text
            text_code = TextBox2
        ElseIf goal = 2 Then
            code = TextBox3.Text
            text_code = TextBox3
        ElseIf goal = 3 Then
            code = TextBox5.Text
            text_code = TextBox5
        ElseIf goal = 4 Then
            code = TextBox4.Text
            text_code = TextBox4
        ElseIf goal = 7 Then
            code = TextBox9.Text
            text_code = TextBox9
        ElseIf goal = 8 Then
            code = TextBox8.Text
            text_code = TextBox8
        ElseIf goal = 9 Then
            code = TextBox11.Text
            text_code = TextBox11
        ElseIf goal = 10 Then
            code = TextBox10.Text
            text_code = TextBox10
        End If
        Try

            R = CInt(code.Split(",")(0).ToString())
            G = CInt(code.Split(",")(1).ToString())
            B = CInt(code.Split(",")(2).ToString())
            If R >= 0 And R <= 255 And G >= 0 And G <= 255 And B >= 0 And B <= 255 Then
                control_imput_color_rgb = True
            Else control_imput_color_rgb = False
                text_code.Text = "Ошибка!"
            End If
        Catch ex As Exception
            text_code.Text = "Ошибка!"
        End Try
    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, Button15.Click, Button8.Click, Button13.Click, Button63.Click, Button31.Click, Button85.Click, Button83.Click

        Dim goal_buttom As Button
        Dim goal_textbox As TextBox
        Dim goal_mem, goal_apply_changes As Integer

        If sender.name = Button5.Name Then
            goal_textbox = TextBox2
            goal_buttom = Button9
            goal_mem = 1
            goal_apply_changes = 1
        ElseIf sender.name = Button8.Name Then
            goal_buttom = Button10
            goal_textbox = TextBox3
            goal_mem = 2
            goal_apply_changes = 1
        ElseIf sender.name = Button15.Name Then
            goal_buttom = Button14
            goal_textbox = TextBox5
            goal_mem = 3
            goal_apply_changes = 2
        ElseIf sender.name = Button13.Name Then
            goal_buttom = Button12
            goal_textbox = TextBox4
            goal_mem = 4
            goal_apply_changes = 2
        ElseIf sender.name = Button63.Name Then
            goal_buttom = Button62
            goal_textbox = TextBox9
            goal_mem = 7
        ElseIf sender.name = Button31.Name Then
            goal_buttom = Button30
            goal_textbox = TextBox8
            goal_mem = 8
        ElseIf sender.name = Button85.Name Then
            goal_buttom = Button84
            goal_textbox = TextBox11
            goal_mem = 9
        ElseIf sender.name = Button83.Name Then
            goal_buttom = Button80
            goal_textbox = TextBox10
            goal_mem = 10
        End If

        goal_textbox.Text = Replace(goal_textbox.Text, ".", ",")

        If control_imput_color_rgb(goal_mem) Then
            color_set(goal_buttom, goal_textbox.Text, 0)
            upd_and_save_mem(goal_mem)
            apply_changes(goal_apply_changes)
        End If

    End Sub


    'нажатие на цвет текста для всей таблицы из памяти
    Private Sub Color_text_memory_1_Click(sender As Object, e As EventArgs) Handles color_text_memory_1.Click, color_text_memory_2.Click, color_text_memory_3.Click, color_text_memory_4.Click, color_text_memory_5.Click, color_text_memory_6.Click, color_text_memory_7.Click, color_text_memory_8.Click
        choose_from_memory(1, sender.BackColor)
    End Sub

    'нажатие на цвет заливки для всей таблицы из памяти
    Private Sub Color_fill_memory_1_Click(sender As Object, e As EventArgs) Handles color_fill_memory_1.Click, color_fill_memory_2.Click, color_fill_memory_3.Click, color_fill_memory_4.Click, color_fill_memory_5.Click, color_fill_memory_6.Click, color_fill_memory_7.Click, color_fill_memory_8.Click
        choose_from_memory(2, sender.BackColor)
    End Sub

    'нажатие на цвет текста для шапки из памяти
    Private Sub color_head_text_memory_1_Click(sender As Object, e As EventArgs) Handles color_head_text_memory_1.Click, color_head_text_memory_2.Click, color_head_text_memory_3.Click, color_head_text_memory_4.Click, color_head_text_memory_5.Click, color_head_text_memory_6.Click, color_head_text_memory_7.Click, color_head_text_memory_8.Click
        choose_from_memory(3, sender.BackColor)
    End Sub

    'нажатие на цвет заливки для шапки из памяти
    Private Sub color_head_fill_memory_1_Click(sender As Object, e As EventArgs) Handles color_head_fill_memory_1.Click, color_head_fill_memory_2.Click, color_head_fill_memory_3.Click, color_head_fill_memory_4.Click, color_head_fill_memory_5.Click, color_head_fill_memory_6.Click, color_head_fill_memory_7.Click, color_head_fill_memory_8.Click
        choose_from_memory(4, sender.BackColor)
    End Sub

    Private Sub Color_text_pipetka_memory_1_Click(sender As Object, e As EventArgs) Handles color_text_pipetka_memory_1.Click, color_text_pipetka_memory_2.Click, color_text_pipetka_memory_3.Click, color_text_pipetka_memory_4.Click, color_text_pipetka_memory_5.Click, color_text_pipetka_memory_6.Click, color_text_pipetka_memory_7.Click, color_text_pipetka_memory_8.Click
        choose_from_memory(7, sender.BackColor)
    End Sub

    Private Sub Color_bg_pipetka_memory_1_Click(sender As Object, e As EventArgs) Handles color_bg_pipetka_memory_1.Click, color_bg_pipetka_memory_2.Click, color_bg_pipetka_memory_3.Click, color_bg_pipetka_memory_4.Click, color_bg_pipetka_memory_5.Click, color_bg_pipetka_memory_6.Click, color_bg_pipetka_memory_7.Click, color_bg_pipetka_memory_8.Click
        choose_from_memory(8, sender.BackColor)
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim corsorpoint As Point = MousePosition
        Dim R, B, G1, pixel
        Dim screen_detected As Boolean = True
        Dim all_screen() As Screen = Screen.AllScreens
        Dim count_screen As Integer = all_screen.Length
        Dim num_screen = 0
        Dim start_position, finish_position As Integer 'Начальная точка отчета по ширине это 0 первого монитора.
        start_position = 0
        finish_position = all_screen(0).Bounds.Width 'Конечная точка отчета по ширине, правый нижний угол текущего монитора

        If count_screen > 0 Then 'Если мониторов больше одного
            For i = 0 To count_screen - 1
                If corsorpoint.X >= start_position And corsorpoint.X <= finish_position - 1 Then
                Else
                    start_position += all_screen(i).Bounds.Width
                    finish_position += all_screen(i + 1).Bounds.Width
                    num_screen += 1
                End If
            Next i
        End If

        Dim myBmp As New Bitmap(finish_position, all_screen(num_screen).Bounds.Height)
        Dim g As Graphics = Graphics.FromImage(myBmp)

        g.CopyFromScreen(Point.Empty, Point.Empty, myBmp.Size)


        Dim crop_screen_bitmap As New Bitmap(20, 20)
        Dim crop_screen_graphics As Graphics = Graphics.FromImage(crop_screen_bitmap)
        crop_screen_graphics.CopyFromScreen(New Point(MousePosition.X - 10, MousePosition.Y - 10), New Point(0.0), crop_screen_bitmap.Size)
        Panel3.Visible = True

        PictureBox1.BackgroundImage = crop_screen_bitmap

        g.Dispose()
        Try
            If start_position = 0 Then
                pixel = myBmp.GetPixel(corsorpoint.X, corsorpoint.Y)
            ElseIf start_position > 0 Then
                pixel = myBmp.GetPixel(corsorpoint.X - 1, corsorpoint.Y)
            End If
            R = pixel.R
            B = pixel.B
            G1 = pixel.G
            myBmp.Dispose()
            color_pixcolor = CStr(R) + "," + CStr(G1) + "," + CStr(B)
            TextBox1.Text = color_pixcolor
            Button1.BackColor = pixel

            If call_pixcolor.Name = Button6.Name Then
                Button9.BackColor = pixel
                TextBox2.Text = color_pixcolor
            ElseIf call_pixcolor.Name = Button4.Name Then
                Button10.BackColor = pixel
                TextBox3.Text = color_pixcolor
            ElseIf call_pixcolor.Name = Button34.Name Then
                Button14.BackColor = pixel
                TextBox5.Text = color_pixcolor
            ElseIf call_pixcolor.Name = Button33.Name Then
                Button12.BackColor = pixel
                TextBox4.Text = color_pixcolor
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub set_color_getpixel()
        Dim color_textbox As TextBox
        Dim button_color As Button
        Dim num_applychanges, num_changes As Integer

        If call_pixcolor.Name = Button6.Name Then
            color_textbox = TextBox2
            button_color = Button9
            num_applychanges = 1
            num_changes = 1

        ElseIf call_pixcolor.Name = Button4.Name Then
            color_textbox = TextBox3
            button_color = Button10
            num_applychanges = 1
            num_changes = 2

        ElseIf call_pixcolor.Name = Button34.Name Then
            color_textbox = TextBox5
            button_color = Button14
            num_applychanges = 2
            num_changes = 3
        ElseIf call_pixcolor.Name = Button33.Name Then
            color_textbox = TextBox4
            button_color = Button12
            num_applychanges = 2
            num_changes = 4


        ElseIf call_pixcolor.Name = Button82.Name Then
            color_textbox = TextBox9
            button_color = Button62
            ' num_applychanges = 5
            num_changes = 7
        ElseIf call_pixcolor.Name = Button81.Name Then
            color_textbox = TextBox8
            button_color = Button30
            'num_applychanges = 5
            num_changes = 8
        ElseIf call_pixcolor.Name = Button103.Name Then
            color_textbox = TextBox11
            button_color = Button84
            num_applychanges = 4
            num_changes = 9
        ElseIf call_pixcolor.Name = Button102.Name Then
            color_textbox = TextBox10
            button_color = Button80
            num_applychanges = 4
            num_changes = 10
        End If

        color_textbox.Text = color_pixcolor
        If control_imput_color_rgb(num_changes) Then
            color_set(button_color, color_textbox.Text, 0)
            apply_changes(num_applychanges)
            upd_and_save_mem(num_changes)
        End If

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        If timer_save < 2 Then
            Button3.Text = "Сохранено"
            color_set(Button3, "202,255,157", 0)
            timer_save += 1
        Else
            Button3.Text = "Сохранить"
            color_set(Button3, "255,228,196", 0)
            timer_save = 0
            Timer2.Enabled = False
        End If
    End Sub

    Private Sub ComboBox_Table_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown, ComboBox2.KeyDown, ComboBox3.KeyDown, ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            apply_changes(1)
        End If
    End Sub

    Private Sub ComboBox_Table_Leave(sender As Object, e As EventArgs) Handles ComboBox1.Leave, ComboBox2.Leave, ComboBox3.Leave, ComboBox4.Leave, ComboBox2.SelectedIndexChanged, ComboBox1.SelectedIndexChanged, ComboBox3.SelectedIndexChanged, ComboBox4.SelectedIndexChanged
        apply_changes(1)
    End Sub

    Private Sub ComboBox_TextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox15.KeyDown, ComboBox16.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            apply_changes(4)
        End If
    End Sub

    Private Sub ComboBox_TextBox_Leave(sender As Object, e As EventArgs) Handles ComboBox15.Leave, ComboBox16.Leave, ComboBox15.SelectedIndexChanged, ComboBox16.SelectedIndexChanged
        apply_changes(4)
    End Sub


    Private Sub ComboBox_Head_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox5.KeyDown, ComboBox6.KeyDown, ComboBox7.KeyDown, ComboBox8.KeyDown, ComboBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            apply_changes(2)
        End If
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, ComboBox7.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox9.Leave, ComboBox8.Leave, ComboBox6.Leave, ComboBox7.Leave, ComboBox5.Leave
        apply_changes(2)
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged, CheckBox8.CheckedChanged, CheckBox7.CheckedChanged
        apply_changes(2)
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged, CheckBox2.CheckedChanged, CheckBox1.CheckedChanged
        apply_changes(1)
    End Sub

    Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged, RadioButton6.CheckedChanged, RadioButton7.CheckedChanged, RadioButton8.CheckedChanged, RadioButton9.CheckedChanged
        sender.Tabstop = False
        apply_changes(1)
    End Sub

    Private Sub RadioButton_textbox_CheckedChanged(sender As Object, e As EventArgs) Handles Textbox_RadioButton1.CheckedChanged, Textbox_RadioButton2.CheckedChanged, Textbox_RadioButton3.CheckedChanged, Textbox_RadioButton4.CheckedChanged, Textbox_RadioButton5.CheckedChanged, Textbox_RadioButton6.CheckedChanged, Textbox_RadioButton7.CheckedChanged, Textbox_RadioButton8.CheckedChanged, Textbox_RadioButton9.CheckedChanged
        sender.Tabstop = False
        apply_changes(4)
    End Sub



    Private Sub Plus_Click(sender As Object, e As EventArgs) Handles Button27.Click, Button2.Click, Button17.Click, Button23.Click, Button25.Click, Button21.Click, Button19.Click, Button105.Click
        If sender.Name = Button27.Name Then
            ComboBox1.Text = CStr(CInt(ComboBox1.Text) + 1)
            apply_changes(1)
        End If
        If sender.Name = Button2.Name Then
            ComboBox4.Text = CStr(Format(CDbl(ComboBox4.Text.Replace(".", ",")) + 0.05, "#,##0.00")).Replace(",", ".")
            apply_changes(1)
        End If
        If sender.Name = Button17.Name Then
            ComboBox3.Text = CStr(Format(CDbl(ComboBox3.Text.Replace(".", ",")) + 0.05, "#,##0.00")).Replace(",", ".")
            apply_changes(1)
        End If
        If sender.Name = Button23.Name Then
            ComboBox9.Text = CStr(CInt(ComboBox9.Text) + 1)
            apply_changes(2)
        End If
        If sender.Name = Button25.Name Then
            ComboBox5.Text = CStr(CInt(ComboBox5.Text) + 1)
            apply_changes(2)
        End If
        If sender.Name = Button21.Name Then
            ComboBox7.Text = CStr(Format(CDbl(ComboBox7.Text.Replace(".", ",")) + 0.05, "#,##0.00")).Replace(",", ".")
            apply_changes(2)
        End If
        If sender.Name = Button19.Name Then
            ComboBox6.Text = CStr(Format(CDbl(ComboBox6.Text.Replace(".", ",")) + 0.05, "#,##0.00")).Replace(",", ".")
            apply_changes(2)
        End If

        If sender.Name = Button105.Name Then 'текстовый блок
            ComboBox16.Text = CStr(CInt(ComboBox16.Text) + 1)
            apply_changes(4)
        End If


    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        Settings_save(1)
    End Sub

    Private Sub Minus_Click(sender As Object, e As EventArgs) Handles Button26.Click, Button11.Click, Button16.Click, Button22.Click, Button24.Click, Button20.Click, Button18.Click, Button104.Click
        Dim i As Double
        Dim y As Integer

        If sender.Name = Button26.Name Then
            y = CInt(ComboBox1.Text) - 1
            If y < 1 Then
                y = 1
            End If
            ComboBox1.Text = CStr(y)
            apply_changes(1)
        End If
        If sender.Name = Button11.Name Then
            i = CDbl(ComboBox4.Text.Replace(".", ",")) - 0.05
            If i < 0 Then
                i = 0
            End If
            ComboBox4.Text = CStr(Format(i, "#,##0.00")).Replace(",", ".")
            apply_changes(1)
        End If
        If sender.Name = Button16.Name Then
            i = CDbl(ComboBox3.Text.Replace(".", ",")) - 0.05
            If i < 0 Then
                i = 0
            End If
            ComboBox3.Text = CStr(Format(i, "#,##0.00")).Replace(",", ".")
            apply_changes(1)
        End If
        If sender.Name = Button22.Name Then
            y = CInt(ComboBox9.Text) - 1
            If y < 1 Then
                y = 1
            End If
            ComboBox9.Text = CStr(y)
            apply_changes(2)
        End If
        If sender.Name = Button24.Name Then
            y = CInt(ComboBox5.Text) - 1
            If y < 1 Then
                y = 1
            End If
            ComboBox5.Text = CStr(y)
            apply_changes(2)
        End If
        If sender.Name = Button20.Name Then
            i = CDbl(ComboBox7.Text.Replace(".", ",")) - 0.05
            If i < 0 Then
                i = 0
            End If
            ComboBox7.Text = CStr(Format(i, "#,##0.00")).Replace(",", ".")
            apply_changes(2)
        End If
        If sender.Name = Button18.Name Then
            i = CDbl(ComboBox6.Text.Replace(".", ",")) - 0.05
            If i < 0 Then
                i = 0
            End If
            ComboBox6.Text = CStr(Format(i, "#,##0.00")).Replace(",", ".")
            apply_changes(2)
        End If

        If sender.Name = Button104.Name Then 'текстовый блок
            y = CInt(ComboBox16.Text) - 1
            If y < 1 Then
                y = 1
            End If
            ComboBox16.Text = CStr(y)
            apply_changes(4)
        End If
    End Sub

    Private Sub Head_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles Head_RadioButton1.CheckedChanged, Head_RadioButton2.CheckedChanged, Head_RadioButton3.CheckedChanged, Head_RadioButton4.CheckedChanged, Head_RadioButton5.CheckedChanged, Head_RadioButton6.CheckedChanged, Head_RadioButton7.CheckedChanged, Head_RadioButton8.CheckedChanged, Head_RadioButton9.CheckedChanged
        sender.Tabstop = False
        apply_changes(2)
    End Sub

    Private Sub Tb_t_Click(sender As Object, e As EventArgs) Handles tb_t_1.Click, tb_t_2.Click, tb_t_3.Click, tb_t_4.Click, tb_t_5.Click, tb_t_6.Click, tb_t_7.Click, tb_t_8.Click
        choose_from_memory(9, sender.BackColor)
    End Sub

    Private Sub Tb_f_Click(sender As Object, e As EventArgs) Handles tb_f_1.Click, tb_f_2.Click, tb_f_3.Click, tb_f_4.Click, tb_f_5.Click, tb_f_6.Click, tb_f_7.Click, tb_f_8.Click
        choose_from_memory(10, sender.BackColor)
    End Sub

    Private Sub CheckBox_textbox_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged, CheckBox20.CheckedChanged, CheckBox19.CheckedChanged ' жкч для тексбокса
        apply_changes(4)
    End Sub

    Private Sub Button66_Click(sender As Object, e As EventArgs) Handles Button66.Click 'кнопка память стилей
        'читаем название стилей из файла

        Dim name_styles(8) As String
        Dim num As Integer
        num = 0
        ' Open file.
        FileOpen(1, "D:\Macros Settings\style\style_name.txt", OpenMode.Input, OpenAccess.Default)
        ' Читаем файл до конца
        While Not EOF(1)
            ' Read line into variable.
            name_styles(num) = LineInput(1)
            num = num + 1
        End While

        ' Устанавливаем имена стилей
        Name_style1.Text = name_styles(0)
        Name_style2.Text = name_styles(1)
        Name_style3.Text = name_styles(2)
        Name_style4.Text = name_styles(3)
        Name_style5.Text = name_styles(4)
        Name_style6.Text = name_styles(5)
        Name_style7.Text = name_styles(6)
        Name_style8.Text = name_styles(7)
        FileClose(1)

        'изменяем ширину окна
        If windows_width < 650 Then
            windows_width = 755 ' ширина окна со списком стилей
        Else
            windows_width = 602 ' ширина окна без списка стилей
        End If
        Me.Width = windows_width 'устанавливаем ширину окна
    End Sub

    Private Sub save_style(sender As Object, e As EventArgs) Handles Style1.Click, Style2.Click, Style3.Click, Style4.Click, Style5.Click, Style6.Click, Style7.Click, Style8.Click
        'Запись настроек в файл стиля
        debug.Text = sender.name + " Сохранен!"
    End Sub

    Private Sub save_name_style(sender As Object, e As EventArgs) Handles Name_style1.TextChanged, Name_style2.TextChanged, Name_style3.TextChanged, Name_style4.TextChanged, Name_style5.TextChanged, Name_style6.TextChanged, Name_style7.TextChanged, Name_style8.TextChanged
        'сохраниение имени стиля при его измнении
        debug.Text = sender.name + " " + sender.text
    End Sub



    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown, TextBox5.KeyDown, TextBox3.KeyDown, TextBox4.KeyDown

        Dim goal_textbox As TextBox
        Dim goal_button As Button

        If sender.Text = TextBox2.Text Then
            goal_textbox = TextBox2
            goal_button = Button9
        End If

        If sender.Text = TextBox5.Text Then
            goal_textbox = TextBox5
            goal_button = Button14
        End If

        If sender.Text = TextBox3.Text Then
            goal_textbox = TextBox3
            goal_button = Button10
        End If

        If sender.Text = TextBox4.Text Then
            goal_textbox = TextBox4
            goal_button = Button12
        End If

        If e.KeyCode = Keys.Enter Then

            If control_imput_color_rgb(1) Then
                color_set(goal_button, goal_textbox.Text, 0)
                upd_and_save_mem(1)
                apply_changes(1)
            End If

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click 'сохранение настроек
        Settings_save(1)
        Settings_save(2)
        Settings_save(4)
    End Sub



    Private Sub Window_view(view_id As Int16) 'Функция показывающая или скрывающая доп. опции для таблиц
        Window_view_memory = view_id
        If view_id = 1 Then 'Показывать только настройки для таблицы
            set_window_size(279, windows_width) ' устанавливает размер всего окна программы
            set_controlTab_size(240) 'устанавливает размер окна с вкладками
            set_saveBtn_position(10, 195) 'устанавливает положение кнопки "сохранить"
            Label19.Visible = False
            Label31.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
        ElseIf view_id = 2 Then 'Показывать настройки для таблицы и шапки таблицы
            set_window_size(507, windows_width) ' устанавливает размер всего окна программы
            set_controlTab_size(468) 'устанавливает размер окна с вкладками
            set_label_head_position(10, 205) 'устанавливает надпись "формат. шапки"
            set_saveBtn_position(10, 420) 'устанавливает положение кнопки "сохранить" 
            Label19.Visible = True
            Label31.Visible = False
            Panel4.Visible = True
            Panel5.Visible = False
        ElseIf view_id = 3 Then 'Показывать настройки для таблицы и отдельных ячеек
            set_window_size(507, windows_width) ' устанавливает размер всего окна программы
            set_controlTab_size(468) 'устанавливает размер окна с вкладками
            set_label_cells_position(10, 205) 'устанавливает надпись "формат. шапки"
            set_saveBtn_position(10, 390) 'устанавливает положение кнопки "сохранить"
            Label19.Visible = False
            Label31.Visible = True
            Panel4.Visible = False
            Panel5.Visible = True
            Panel5.Location = New Point(3, 212)
        ElseIf view_id = 4 Then 'Показывать настройки для таблицы, шапки таблицы и окно для отдельных ячеек

            set_window_size(715, windows_width) ' устанавливает размер всего окна программы
            set_controlTab_size(675) 'устанавливает размер окна с вкладками
            set_label_head_position(10, 205) 'устанавливает надпись "формат. шапки"
            set_label_cells_position(10, 421) 'устанавливает надпись "формат. ячеек"
            set_saveBtn_position(10, 620) 'устанавливает положение кнопки "сохранить"
            Label19.Visible = True
            Label31.Visible = True
            Panel4.Visible = True
            Panel5.Visible = True
            Panel5.Location = New Point(3, 427)
        End If

    End Sub



    Private Sub set_window_size(heigt, width)
        Me.Height = heigt
        Me.Width = width
    End Sub

    Private Sub ComboBox17_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub set_controlTab_size(heigt)
        TabControl1.Height = heigt
    End Sub

    Private Sub set_label_head_position(x, y)
        Label19.Location = New Point(x, y)
    End Sub

    Private Sub set_label_cells_position(x, y)
        Label31.Location = New Point(x, y)
    End Sub

    Private Sub set_saveBtn_position(x, y)
        Button3.Location = New Point(x, y)
    End Sub

End Class




