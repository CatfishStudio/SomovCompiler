Module CORE
    '//КОМЕНТАРИИ В МОДУЛЕ////////////////////////////////////////////////////////
    Private MultilineComent As Boolean = False
    Private ModuleCodeNoneComent As New RichTextBox 'Модуль без коментариев

    '//ОПЕРАТОРЫ КОМПИЛЯТОРА//////////////////////////////////////////////////////
    Private LISTOperators As New ListBox 'Перечень всех заризервированных операторов



    '//ТАБЛИЦА ИДЕНТИФИКАТОРОВ////////////////////////////////////////////////////
    Private TABLE_ID(0) As String        'все элементы модуля.
    'ПЕРЕМЕННЫХ ------------------------------
    '-----------Имя-Тип-Значение--------------
    Public Perem(1, 1, 0) As String     'переменные
    Public countPerem As Integer        'количество переменных
    Public IndexStatusPerem As Integer  'индекс переменной "Состояние"
    '-----------Имя-Тип-Тело--------
    Public Func(1, 1, 0) As String      'функции
    Public countFunc As Integer         'количество Функций
    '-----------Имя-| |-Тело-------------
    Public Proc(1, 1, 0) As String      'процедуры
    Public countProc As Integer         'количество процедур

    'ЗАПОЛНЕНИЕ ОБЩЕЙ ТАБЛИЦЫ ИДЕНТИФИКАТОРОВ ЭЛЕМЕНТАМИ КОДА
    Private Sub AddTableID(ByVal ModuleCode As RichTextBox)
        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String

        For i As Integer = 0 To MCLength
            'Сохраняем номер строки модуля в таблицу
            ReDim Preserve TABLE_ID(TABLE_ID.Length)
            Dim actionLine As Integer = i + 1
            TABLE_ID(TABLE_ID.Length - 1) = "Строка: " + actionLine.ToString


            STROKA = ModuleCode.Lines.GetValue(i)
            'группировка в ТОКЕНЫ
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If (Simvol = " " And MultilineComent = False) Then
                    '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                    If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                        ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                        TABLE_ID(TABLE_ID.Length - 1) = Operator
                    End If
                    Operator = ""
                Else
                    If (Simvol = ";" Or Simvol = ";" And MultilineComent = False) Then
                        '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                        If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                            ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                            TABLE_ID(TABLE_ID.Length - 1) = Operator
                            ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                            TABLE_ID(TABLE_ID.Length - 1) = ";"
                        Else
                            ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                            TABLE_ID(TABLE_ID.Length - 1) = ";"
                        End If
                        Operator = ""
                    Else
                        If (Simvol = "," And MultilineComent = False) Then
                            '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                            If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                TABLE_ID(TABLE_ID.Length - 1) = Operator
                                ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                TABLE_ID(TABLE_ID.Length - 1) = ","
                            Else
                                ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                TABLE_ID(TABLE_ID.Length - 1) = ","
                            End If
                            Operator = ""
                        Else
                            If (Simvol = "(" And MultilineComent = False) Then
                                '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                    TABLE_ID(TABLE_ID.Length - 1) = Operator
                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                    TABLE_ID(TABLE_ID.Length - 1) = "("
                                Else
                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                    TABLE_ID(TABLE_ID.Length - 1) = "("
                                End If
                                Operator = ""
                            Else
                                If (Simvol = ")" And MultilineComent = False) Then
                                    '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                    If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                        ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                        TABLE_ID(TABLE_ID.Length - 1) = Operator
                                        ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                        TABLE_ID(TABLE_ID.Length - 1) = ")"
                                    Else
                                        ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                        TABLE_ID(TABLE_ID.Length - 1) = ")"
                                    End If
                                    Operator = ""
                                Else
                                    If (Simvol = "{" And MultilineComent = False) Then
                                        '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                        If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                            ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                            TABLE_ID(TABLE_ID.Length - 1) = Operator
                                            ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                            TABLE_ID(TABLE_ID.Length - 1) = "{"
                                        Else
                                            ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                            TABLE_ID(TABLE_ID.Length - 1) = "{"
                                        End If
                                        Operator = ""
                                    Else
                                        If (Simvol = "}" And MultilineComent = False) Then
                                            '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                            If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                                ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                                TABLE_ID(TABLE_ID.Length - 1) = Operator
                                                ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                                TABLE_ID(TABLE_ID.Length - 1) = "}"
                                            Else
                                                ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                                TABLE_ID(TABLE_ID.Length - 1) = "}"
                                            End If
                                            Operator = ""
                                        Else
                                            If (iOperator = STROKA.Length - 1) Then
                                                '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                                Operator = Operator + Simvol
                                                If (Operator <> "" And Operator <> " " And Operator <> "/*" And MultilineComent = False) Then
                                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) 'Увеличиваем массив
                                                    TABLE_ID(TABLE_ID.Length - 1) = Operator
                                                End If
                                                If (Operator = "/*") Then MultilineComent = True
                                                If (Operator = "*/") Then MultilineComent = False
                                                Operator = ""
                                            Else
                                                'не учитывая табуляцию (177580) группируем в ТОКЕН
                                                If (Simvol.GetHashCode <> 177580 And Simvol.GetHashCode <> -842352729) Then
                                                    Operator = Operator + Simvol
                                                    If (Operator = "//") Then Exit For
                                                    If (Operator = "/*") Then MultilineComent = True
                                                    If (Operator = "*/") Then MultilineComent = False
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            'очистка...
            Operator = ""
            Simvol = ""
        Next

        'ВЫГРУЗКА МОДУЛЯ БЕЗ КОМЕНТАРИЕВ
        Dim Element As String
        Dim StrokaCode As String

        For iTest As Integer = 1 To TABLE_ID.Length - 1
            Element = TABLE_ID(iTest).ToString

            If (Element.Length > 6) Then StrokaCode = Element.Chars(0) + Element.Chars(1) + Element.Chars(2) + Element.Chars(3) + Element.Chars(4) + Element.Chars(5) + Element.Chars(6)
            If (StrokaCode = "Строка:") Then
                ModuleCodeNoneComent.Text = ModuleCodeNoneComent.Text + System.Environment.NewLine
            Else
                If (Element = "Перем" Or Element = "перем" Or Element = "Процедура" Or Element = "процедура" Or Element = "Функция" Or Element = "функция" Or Element = "{" Or Element = "}" Or Element = "(" Or Element = ")" Or Element = ";" Or Element = ",") Then
                    ModuleCodeNoneComent.Text = ModuleCodeNoneComent.Text + Element
                Else
                    ModuleCodeNoneComent.Text = ModuleCodeNoneComent.Text + " " + Element
                End If

            End If
            StrokaCode = ""
            Element = ""
        Next
    End Sub

    'ЗАГРУЗКА ТАБЛИЦЫ ИДЕНТИФИКАТОРОВ ПЕРЕМЕННЫМИ, ФУНКЦИЯМИ, ПРОЦЕДУРАМИ.
    Private Sub add_Perem_Func_Proc(ByVal ModuleCode As RichTextBox)
        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim InputPARAM As Boolean = False 'параметры процедуры или функции

        For i As Integer = 0 To MCLength
            STROKA = ModuleCode.Lines.GetValue(i)
            'Определяем ТИП
            TYPE(STROKA)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'группировка в ТОКЕНЫ
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If Simvol = " " Or Simvol = ";" Or Simvol = ";" Or Simvol = "," Or Simvol = "(" Or Simvol = ")" Or Simvol = "{" Or Simvol = "}" Then 'Прекращение группирование ТОКЕНА
                    '//ПЕРЕМЕННЫЕ//////////////////////////////////////
                    If (TYPE_peremen = True) Then
                        If Operator = "Перем" Or Operator = "перем" Or Operator = "Тип" Or Operator = "тип" Or Operator = "Строка" Or Operator = "строка" Or Operator = "Число" Or Operator = "число" Or Operator = "" Or Operator = " " Then
                            'нет действий.
                        Else
                            '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                            If (SyntaxPerem(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! Строка: " + LineError.ToString + " (переменная [" + Operator + "] уже существует, её повторное объявление недопустимо.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                            countPerem = countPerem + 1 'учет количества переменных
                            ReDim Preserve Perem(1, 1, countPerem) 'Увеличиваем массив
                            '**ИМЯ ПЕРЕМЕННОЙ**************************
                            Perem(0, 0, countPerem) = Operator
                            '**ТИП ЧИСЛО*******************************
                            If (TYPE_chislo = True) Then
                                Perem(0, 1, countPerem) = "Число"
                                'ЗНАЧЕНИЕ ПО УМОЛЧАНИЮ
                                Perem(1, 1, countPerem) = "0"
                            End If
                            '**ТИП СТРОКА******************************
                            If (TYPE_stroka = True) Then
                                Perem(0, 1, countPerem) = "Строка"
                                'ЗНАЧЕНИЕ ПО УМОЛЧАНИЮ
                                Perem(1, 1, countPerem) = ""
                            End If
                        End If
                    End If

                    '//ПРОЦЕДУРЫ//////////////////////////////////////
                    If (TYPE_procedure = True) Then
                        If Operator = "Процедура" Or Operator = "процедура" Or Operator = "(" Or Operator = ")" Or Operator = "{" Or Operator = "}" Or Operator = "" Or Operator = " " Then
                            'нет действий.
                        Else
                            '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                            If (SyntaxProc(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! Строка: " + LineError.ToString + " (Процедура с именем [" + Operator + "] уже объявлена, повторное объявление недопустимо.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                            countProc = countProc + 1
                            ReDim Preserve Proc(1, 1, countProc) 'Увеличиваем массив
                            '**ИМЯ ПРОЦЕДУРЫ***************************
                            Proc(0, 0, countProc) = Operator
                            Proc(0, 1, countProc) = "null"
                            '**ТЕЛО ПРОЦЕДУРЫ *************************
                            Proc(1, 1, countProc) = BodyProc(ModuleCode, i)
                        End If
                    End If
                    If (TYPE_procedure_Param = True) Then
                        If Operator = "Процедура" Or Operator = "процедура" Or Operator = "" Or Operator = " " Then
                            'нет действий.
                        Else
                            'Параметры процедуры записываем в таблицу переменных
                            If (Operator = "Перем" Or Operator = "перем") Then
                                InputPARAM = True
                            End If
                            '**********************************************
                            If (InputPARAM = True) Then 'ПАРАМЕТРЫ ПРОЦЕДУРЫ
                                If Operator = "Перем" Or Operator = "перем" Or Operator = "Тип" Or Operator = "тип" Or Operator = "Строка" Or Operator = "строка" Or Operator = "Число" Or Operator = "число" Or Operator = "" Or Operator = " " Then
                                    'нет действий.
                                Else
                                    '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                                    If (SyntaxPerem(Operator) = True) Then
                                        Dim LineError As Integer = i
                                        SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! Строка: " + LineError.ToString + " (переменная [" + Operator + "] уже существует, её повторное объявление не возможно.)")
                                        SomovCompileModule.WError.Show()
                                        Exit For
                                    End If
                                    '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                    countPerem = countPerem + 1 'учет количества переменных
                                    ReDim Preserve Perem(1, 1, countPerem) 'Увеличиваем массив
                                    '**ИМЯ ПЕРЕМЕННОЙ**************************
                                    Perem(0, 0, countPerem) = Operator
                                    '**ТИП ЧИСЛО*******************************
                                    If (TYPE_chislo = True) Then
                                        Perem(0, 1, countPerem) = "Число"
                                        'ЗНАЧЕНИЕ ПО УМОЛЧАНИЮ
                                        Perem(1, 1, countPerem) = "0"
                                    End If
                                    '**ТИП СТРОКА******************************
                                    If (TYPE_stroka = True) Then
                                        Perem(0, 1, countPerem) = "Строка"
                                        'ЗНАЧЕНИЕ ПО УМОЛЧАНИЮ
                                        Perem(1, 1, countPerem) = ""
                                    End If
                                End If
                            Else    'ИМЯ ПРОЦЕДУРЫ И ЕЁ ТЕЛО
                                '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                                If (SyntaxProc(Operator) = True) Then
                                    Dim LineError As Integer = i
                                    SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! Строка: " + LineError.ToString + " (Процедура с именем [" + Operator + "] уже объявлена, повторное объявление недопустимо.)")
                                    SomovCompileModule.WError.Show()
                                    Exit For
                                End If
                                '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                countProc = countProc + 1
                                ReDim Preserve Proc(1, 1, countProc) 'Увеличиваем массив
                                '**ИМЯ ПРОЦЕДУРЫ***************************
                                Proc(0, 0, countProc) = Operator
                                Proc(0, 1, countProc) = "null"
                                '**ТЕЛО ПРОЦЕДУРЫ *************************
                                Proc(1, 1, countProc) = BodyProc(ModuleCode, i)
                            End If
                        End If
                    End If

                    '//ФУНКЦИИ//////////////////////////////////////
                    If (TYPE_function = True) Then
                        If Operator = "Функция" Or Operator = "функция" Or Operator = "(" Or Operator = ")" Or Operator = "{" Or Operator = "}" Or Operator = "" Or Operator = " " Then
                            'нет действий.
                        Else
                            '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                            If (SyntaxFunc(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! Строка: " + LineError.ToString + " (Функция с именем [" + Operator + "] уже объявлена, повторное объявление недопустимо.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                            countFunc = countFunc + 1
                            ReDim Preserve Func(1, 1, countFunc) 'Увеличиваем массив
                            '**ИМЯ ФУНКЦИИ ****************************
                            Func(0, 0, countFunc) = Operator
                            '**ТИП ФУНКЦИИ ****************************
                            Func(0, 1, countFunc) = TypeFunc(ModuleCode, i)
                            '**ТЕЛО ФУНКЦИИ ***************************
                            Func(1, 1, countFunc) = BodyFunc(ModuleCode, i)
                        End If
                    End If
                    If (TYPE_function_Param = True) Then
                        If Operator = "Функция" Or Operator = "функция" Or Operator = "" Or Operator = " " Then
                            'нет действий.
                        Else
                            'Параметры функции записываем в таблицу переменных
                            If (Operator = "Перем" Or Operator = "перем") Then
                                InputPARAM = True
                            End If
                            '**********************************************
                            If (InputPARAM = True) Then 'ПАРАМЕТРЫ ФУНКЦИИ
                                If Operator = "Перем" Or Operator = "перем" Or Operator = "Тип" Or Operator = "тип" Or Operator = "Строка" Or Operator = "строка" Or Operator = "Число" Or Operator = "число" Or Operator = "" Or Operator = " " Then
                                    'нет действий.
                                Else
                                    '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                                    If (SyntaxPerem(Operator) = True) Then
                                        Dim LineError As Integer = i
                                        SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! Строка: " + LineError.ToString + " (переменная [" + Operator + "] уже существует, её повторное объявление не возможно.)")
                                        SomovCompileModule.WError.Show()
                                        Exit For
                                    End If
                                    '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                    countPerem = countPerem + 1 'учет количества переменных
                                    ReDim Preserve Perem(1, 1, countPerem) 'Увеличиваем массив
                                    '**ИМЯ ПЕРЕМЕННОЙ**************************
                                    Perem(0, 0, countPerem) = Operator
                                    '**ТИП ЧИСЛО*******************************
                                    If (TYPE_chislo = True) Then
                                        Perem(0, 1, countPerem) = "Число"
                                        'ЗНАЧЕНИЕ ПО УМОЛЧАНИЮ
                                        Perem(1, 1, countPerem) = "0"
                                    End If
                                    '**ТИП СТРОКА******************************
                                    If (TYPE_stroka = True) Then
                                        Perem(0, 1, countPerem) = "Строка"
                                        'ЗНАЧЕНИЕ ПО УМОЛЧАНИЮ
                                        Perem(1, 1, countPerem) = ""
                                    End If
                                    If (Operator = "Состояние") Then IndexStatusPerem = countPerem
                                End If
                            Else    'ИМЯ ФУНКЦИИ И ЕЁ ТЕЛО
                            '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                            If (SyntaxFunc(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! Строка: " + LineError.ToString + " (Функция с именем [" + Operator + "] уже объявлена, повторное объявление недопустимо.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                            countFunc = countFunc + 1
                            ReDim Preserve Func(1, 1, countFunc) 'Увеличиваем массив
                            '**ИМЯ ПРОЦЕДУРЫ***************************
                            Func(0, 0, countFunc) = Operator
                            '**ТИП ФУНКЦИИ ****************************
                            Func(0, 1, countFunc) = TypeFunc(ModuleCode, i)
                            '**ТЕЛО ПРОЦЕДУРЫ *************************
                            Func(1, 1, countFunc) = BodyFunc(ModuleCode, i)
                            End If
                        End If
                    End If
                    '...

                    Operator = ""
                Else
                    'не учитывая табуляцию (177580) группируем в ТОКЕН
                    If (Simvol.GetHashCode <> 177580 And Simvol.GetHashCode <> -842352729) Then Operator = Operator + Simvol
                End If
            Next
            'очистка...
            ClearTYPE()
            Operator = ""
            Simvol = ""
            InputPARAM = False
        Next
    End Sub

    'ТЕЛО ПРОЦЕДУРЫ
    Private Function BodyProc(ByVal ModuleCode As RichTextBox, ByVal StartProc As Integer) As String
        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim Open_Close As Integer 'подсчет открытых и закрытых фигурных скобок { }
        Dim RTB As RichTextBox = New RichTextBox

        StartProc = StartProc + 1
        For i As Integer = StartProc To MCLength
            STROKA = ModuleCode.Lines.GetValue(i)

            'обработка строки программного кода
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If Simvol = "{" Then '
                    Open_Close = Open_Close + 1 'Открытая скобка {
                End If
                If Simvol = "}" Then '
                    Open_Close = Open_Close - 1  'Закрытая скобка }
                End If
                Simvol = ""
            Next

            If (i <> StartProc) Then 'Завершение получения тела процедуры
                If (Open_Close = 0) Then
                    Return RTB.Text
                    Exit Function
                Else
                    If (RTB.Text = "") Then
                        RTB.Text = STROKA
                    Else
                        RTB.Text = RTB.Text + System.Environment.NewLine + STROKA
                    End If
                End If
            End If
            Simvol = ""
        Next
    End Function

    'ТИП И ТЕЛО ФУНКЦИИ
    Private Function TypeFunc(ByVal ModuleCode As RichTextBox, ByVal StartProc As Integer) As String
        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim ThisTYPE As Boolean = False

        STROKA = ModuleCode.Lines.GetValue(StartProc)

        'обработка строки программного кода
        For iOperator As Integer = 0 To STROKA.Length - 1
            Simvol = STROKA.Chars(iOperator)
            If (Simvol = " " Or Simvol = "(" Or Simvol = ")") Then
                If (Simvol = ")") Then
                    ThisTYPE = True
                End If

                Operator = ""
            Else
                Operator = Operator + Simvol
                If (iOperator = STROKA.Length - 1 And ThisTYPE = True) Then
                    Return Operator
                    Exit Function
                End If
            End If
            Simvol = ""
        Next
    End Function
    Private Function BodyFunc(ByVal ModuleCode As RichTextBox, ByVal StartProc As Integer) As String
        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim Open_Close As Integer 'подсчет открытых и закрытых фигурных скобок { }
        Dim RTB As RichTextBox = New RichTextBox

        StartProc = StartProc + 1
        For i As Integer = StartProc To MCLength
            STROKA = ModuleCode.Lines.GetValue(i)

            'обработка строки программного кода
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If Simvol = "{" Then '
                    Open_Close = Open_Close + 1 'Открытая скобка {
                End If
                If Simvol = "}" Then '
                    Open_Close = Open_Close - 1  'Закрытая скобка }
                End If
                Simvol = ""
            Next

            If (i <> StartProc) Then 'Завершение получения тела процедуры
                If (Open_Close = 0) Then
                    Return RTB.Text
                    Exit Function
                Else
                    If (RTB.Text = "") Then
                        RTB.Text = STROKA
                    Else
                        RTB.Text = RTB.Text + System.Environment.NewLine + STROKA
                    End If
                End If
            End If
            Simvol = ""
        Next
    End Function



    '//ЛЕКСИЧЕСКИЙ АНАЛИЗ////////////////////////////////////////////////////////
    'Определяем какую иформацию несет данный оператор.
    Private TYPE_peremen As Boolean
    Private TYPE_chislo As Boolean 'тип число
    Private TYPE_stroka As Boolean 'тип строка
    Private TYPE_object As Boolean 'тип объект
    Private TYPE_prisvoit As Boolean 'присвоить
    Private TYPE_message As Boolean 'сообщить
    Private TYPE_zapros As Boolean 'запрос
    Private TYPE_question As Boolean 'вопрос
    Private TYPE_procedure As Boolean 'процедура
    Private TYPE_procedure_Param As Boolean 'процедура с параметрами
    Private TYPE_procedure_RUN As Boolean 'выполнить процедуру
    Private TYPE_function As Boolean 'функция
    Private TYPE_function_Param As Boolean 'функция с параметрами
    Private TYPE_matematika As Boolean

    'ОПРЕДЕЛЕНИЕ ТИПА ПОЛУЧЕННОГО ОПЕРАТОРА
    Private Sub TYPE(ByVal Stroka As String) 'Тип
        ClearTYPE() 'очистка предыдущих определений типов
        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        For iOperator As Integer = 0 To Stroka.Length - 1
            Simvol = Stroka.Chars(iOperator)
            If Simvol = " " Or Simvol = ";" Or Simvol = ";" Or Simvol = "(" Or Simvol = ")" Then 'Если символ точка с запятой значит оператор прочитан

                'Определяем значение оператора.
                If Operator = "Перем" Or Operator = "перем" Then TYPE_peremen = True
                If Operator = "Число" Or Operator = "число" Then TYPE_chislo = True
                If Operator = "Строка" Or Operator = "строка" Then TYPE_stroka = True
                If Operator = "Форма" Or Operator = "форма" Then TYPE_object = True
                If Operator = "=" Then TYPE_prisvoit = True
                If Operator = "+" Or Operator = "-" Or Operator = "*" Or Operator = "/" Then
                    TYPE_matematika = True  'математические вычисления
                    TYPE_peremen = False
                End If
                'Встроенные функции
                If Operator = "Сообщить" Or Operator = "сообщить" Then TYPE_message = True
                If Operator = "Вопрос" Or Operator = "вопрос" Then TYPE_question = True
                If Operator = "Запрос" Or Operator = "запрос" Then TYPE_zapros = True
                'Определяем тип ПРОЦЕДУРА ********************************
                If (Operator = "Процедура" Or Operator = "процедура") Then
                    TYPE_procedure = True   'Процедура без параметров
                End If
                If (TYPE_procedure = True And TYPE_peremen = True) Then
                    TYPE_procedure_Param = True 'Процедура с параметами
                    TYPE_procedure = False
                    TYPE_peremen = False
                End If
                For i As Integer = 1 To CORE.countProc
                    If (Proc(0, 0, i).ToString = Operator) Then
                        TYPE_procedure_RUN = True
                        Exit Sub
                    End If
                Next

                'Определяем тип ФУНКЦИЯ **********************************
                If (Operator = "Функция" Or Operator = "функция") Then
                    TYPE_function = True    'Функция без параметров
                End If
                If (TYPE_function = True And TYPE_peremen = True) Then
                    TYPE_function_Param = True  'Функция с параметрами
                    TYPE_function = False
                    TYPE_peremen = False
                End If

                '-----------------------------
                Operator = ""
                Simvol = ""
            Else
                'не учитывая табуляцию (WinXP 177580, WinSeven -842352729) группируем в ТОКЕН
                If (Simvol.GetHashCode <> 177580 And Simvol.GetHashCode <> -842352729 And Simvol <> " ") Then Operator = Operator + Simvol
            End If
        Next
    End Sub



    '//СИНТАКСИЧЕСКИЙ И СЕМАНТИЧЕСКИЙ АНАЛИЗ///////////////////////////////////
    'Проверка имён Переменных
    Private Function SyntaxPerem(ByVal Test As String) As Boolean
        'Проверка повторения имён переменных*****
        For i As Integer = 1 To CORE.countPerem
            If (Perem(0, 0, i).ToString = Test) Then
                Return True
                Exit Function
            End If
        Next
        '****************************************
    End Function

    'Проверка имен Процедур
    Private Function SyntaxProc(ByVal Test As String) As Boolean
        'Проверка повторения имён переменных*****
        For i As Integer = 1 To CORE.countProc
            If (Proc(0, 0, i).ToString = Test) Then
                Return True
                Exit Function
            End If
        Next
        '****************************************
    End Function

    'Проверка имен Функций
    Private Function SyntaxFunc(ByVal Test As String) As Boolean
        'Проверка повторения имён переменных*****
        For i As Integer = 1 To CORE.countFunc
            If (Func(0, 0, i).ToString = Test) Then
                Return True
                Exit Function
            End If
        Next
        '****************************************
    End Function

    'Полная проверка всего модуля
    Private Function SyntaxAll(ByVal Test As RichTextBox) As Boolean

        If (SomovCompileModule.WErrorShow = True) Then
            Return True
            Exit Function
        End If

        'ВЫГРУЗКА МОДУЛЯ БЕЗ КОМЕНТАРИЕВ
        Dim Element As String
        Dim StrokaCode As String
        Dim ActionLine As String
        Dim StringOPEN As Boolean = False   'строка открыта/закрыта
        Dim ListPeremFor As New ListBox         'дополнительный лист объявленных переменных в цикле

        LISTOperators.Items.Add("")
        LISTOperators.Items.Add(" ")
        LISTOperators.Items.Add("(")
        LISTOperators.Items.Add(")")
        LISTOperators.Items.Add("{")
        LISTOperators.Items.Add("}")
        LISTOperators.Items.Add("'")
        LISTOperators.Items.Add("""")
        LISTOperators.Items.Add(",")
        LISTOperators.Items.Add(".")
        LISTOperators.Items.Add("0")
        LISTOperators.Items.Add("1")
        LISTOperators.Items.Add("2")
        LISTOperators.Items.Add("3")
        LISTOperators.Items.Add("4")
        LISTOperators.Items.Add("5")
        LISTOperators.Items.Add("6")
        LISTOperators.Items.Add("7")
        LISTOperators.Items.Add("8")
        LISTOperators.Items.Add("9")

        'Проверка на наличие "Главной" функции
        Dim StartFunc As Boolean = False
        For i As Integer = 1 To CORE.countFunc
            If (Func(0, 0, i).ToString = "Главная" Or Func(0, 0, i).ToString = "главная") Then
                StartFunc = True
                Exit For
            End If
        Next
        If (StartFunc = False) Then
            SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! в модуле не объявлена ГЛАВНАЯ функция.(Функция Главная(Перем Состояние Тип Строка) Число)")
            SomovCompileModule.WError.Show()
            Return True
            Exit Function
        End If
        '************************************


        For iTest As Integer = 1 To TABLE_ID.Length - 1
            Element = TABLE_ID(iTest).ToString

            If (Element.Length > 6) Then StrokaCode = Element.Chars(0) + Element.Chars(1) + Element.Chars(2) + Element.Chars(3) + Element.Chars(4) + Element.Chars(5) + Element.Chars(6)
            If (StrokaCode = "Строка:") Then
                ActionLine = Element
                'проверяем символ окончания строки ;
                If (iTest > 1) Then
                    If (TABLE_ID(iTest - 1).ToString = ";" Or TABLE_ID(iTest - 1).ToString = ")" Or TABLE_ID(iTest - 1).ToString = "Иначе" Or TABLE_ID(iTest - 1).ToString = "иначе" Or TABLE_ID(iTest - 1).ToString = "Строка" Or TABLE_ID(iTest - 1).ToString = "строка" Or TABLE_ID(iTest - 1).ToString = "Число" Or TABLE_ID(iTest - 1).ToString = "число" Or TABLE_ID(iTest - 1).ToString = "{" Or TABLE_ID(iTest - 1).ToString = "}") Then
                        'СИМВОЛ КОНЦА СТРОКИ НАЙДЕН - ВСЁ ВЕРНО!!!
                    Else
                        'Вывод сообщения об ошибке ---------------
                        Dim FrontElementTest As String = TABLE_ID(iTest - 1).ToString
                        If (FrontElementTest.Length > 6) Then FrontElementTest = FrontElementTest.Chars(0) + FrontElementTest.Chars(1) + FrontElementTest.Chars(2) + FrontElementTest.Chars(3) + FrontElementTest.Chars(4) + FrontElementTest.Chars(5) + FrontElementTest.Chars(6)
                        If (FrontElementTest <> "Строка:") Then
                            SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! " + ActionLine + " (не найден символ конца строки ; точка с запятой)")
                            SomovCompileModule.WError.Show()
                        End If
                    End If
                End If
            Else
                'ПРОВЕРКА ВСЕГО МОДУЛЯ НА ВОЗМОЖНЫЕ ОШИБКИ
                '1.Проверка элемента в списке заризервированных операторов
                Dim TextN1 As Boolean = False
                For iTestOperator As Integer = 0 To LISTOperators.Items.Count - 1
                    If (Element.Chars(0) = """" Or Element.Chars(0) = "'" Or Element.Chars(Element.Length - 1) = """" Or Element.Chars(Element.Length - 1) = "'") Then
                        If (Element.Chars(0) = """" And Element.Chars(Element.Length - 1) = """") Then
                            TextN1 = True 'оператор определён как строка всё хорошо!
                            StringOPEN = False 'означает что строка закрыта
                            Exit For
                        End If
                        If (Element.Chars(0) = "'" And Element.Chars(Element.Length - 1) = "'") Then
                            TextN1 = True 'оператор определён как строка всё хорошо!
                            StringOPEN = False 'означает что строка закрыта
                            Exit For
                        End If
                        If (StringOPEN = False) Then
                            StringOPEN = True   'означает что строка открыта
                        Else
                            StringOPEN = False 'означает что строка закрыта
                        End If
                        TextN1 = True 'оператор определён как строка всё хорошо!
                        Exit For
                    End If
                    If (StringOPEN = True) Then
                        TextN1 = True 'оператор определён как строка всё хорошо!
                        Exit For
                    End If

                    If (Element.Chars(0) = "0" Or Element.Chars(0) = "1" Or Element.Chars(0) = "2" Or Element.Chars(0) = "3" Or Element.Chars(0) = "4" Or Element.Chars(0) = "5" Or Element.Chars(0) = "6" Or Element.Chars(0) = "7" Or Element.Chars(0) = "8" Or Element.Chars(0) = "9") Then
                        TextN1 = True 'оператор определён как число всё хорошо!
                        Exit For
                    End If
                    If (Element = LISTOperators.Items.Item(iTestOperator)) Then
                        TextN1 = True 'оператор определён как заризервированный всё хорошо!
                        Exit For
                    End If
                Next
                '2.Просмотр листа переменных цикла
                For iLPFor As Integer = 0 To ListPeremFor.Items.Count - 1
                    If (Element = ListPeremFor.Items.Item(iLPFor)) Then
                        TextN1 = True 'оператор определён как дополнительная переменная цикла всё хорошо!
                        Exit For
                    End If
                Next
                '3.Проверка элемента в списках объявленных переменных, процедур и функций
                If (TextN1 = False And StringOPEN = False) Then
                    'просмотр имён переменных
                    If (SyntaxPerem(Element) = False) Then
                        'просмотр имён процедур
                        If (SyntaxProc(Element) = False) Then
                            'просмотр имён функций
                            If (SyntaxFunc(Element) = False) Then
                                If (iTest - 2 > 0) Then
                                    If (TABLE_ID(iTest - 2).ToString = "Для " Or TABLE_ID(iTest - 2).ToString = "Для") Then
                                        ListPeremFor.Items.Add(Element)
                                    Else
                                        'Вывод сообщения об ошибке ---------------
                                        SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! " + ActionLine + " (" + Element + " - объект не определён)")
                                        SomovCompileModule.WError.Show()
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
                '4 проверяем правильность написания фигурных скобок { }
                If (Element = "{" Or Element = "}") Then
                    Dim front As String 'спереди
                    Dim back As String 'сзади
                    front = TABLE_ID(iTest - 1).ToString
                    If (iTest < TABLE_ID.Length - 1) Then
                        back = TABLE_ID(iTest + 1).ToString
                    Else
                        back = "Строка:"
                    End If

                    If (front.Length > 6) Then front = front.Chars(0) + front.Chars(1) + front.Chars(2) + front.Chars(3) + front.Chars(4) + front.Chars(5) + front.Chars(6)
                    If (back.Length > 6) Then back = back.Chars(0) + back.Chars(1) + back.Chars(2) + back.Chars(3) + back.Chars(4) + back.Chars(5) + back.Chars(6)

                    If (front = "Строка:" And back = "Строка:") Then
                        'СИМВОЛ НАПИСАН С НОВОЙ СТРОКИ - ВСЁ ВЕРНО!!!
                    Else
                        'Вывод сообщения об ошибке ---------------
                        SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! " + ActionLine + " (символ " + Element + " - должен быть написан с новой строки и отдельно от остальных операторов.)")
                        SomovCompileModule.WError.Show()
                    End If
                End If
            End If

            StrokaCode = ""
            Element = ""
        Next

        If (SomovCompileModule.WErrorShow = True) Then
            Return True
            Exit Function
        End If

    End Function



    '//ВЫПОЛНЕНИЕ////////////////////////////////////////////////////////////////
    Public Sub RUN(ByVal Code As RichTextBox, ByVal Operators As ListBox)
        'ОТЛАДКА МОДУЛЯ (Выполнение программного кода)
        CORE.ClearModule()              'Предварительная очистка тиблицы идентификаторов
        '****************************************
        'Подготовка окна сообщения об ошибках
        If (WErrorShow = False) Then
            SomovCompileModule.LoadWError()
            WError.ListBox1.Items.Clear()
        Else
            WError.ListBox1.Items.Clear()
        End If
        '****************************************
        CORE.LISTOperators = Operators
        CORE.AddTableID(Code)
        CORE.add_Perem_Func_Proc(ModuleCodeNoneComent)   'загружаем таблицу идентификаторов
        '****************************************
        If (CORE.SyntaxAll(ModuleCodeNoneComent) = True) Then Exit Sub 'полная проверка модуля
        '****************************************
        RUN_MAIN()   'Непосредственное выполнение кода начиная с "Главной" функции
    End Sub
    'Начало выполнения программного кода
    Private Sub RUN_MAIN()
        Dim MainFunction As New RichTextBox
        'НАЧАЛО ВЫПОЛНЕНИЯ С "ГЛАВНОЙ" ФУНКЦИИ
        For i As Integer = 1 To CORE.countFunc
            If (Func(0, 0, i).ToString = "Главная" Or Func(0, 0, i).ToString = "Главная") Then
                MainFunction.Text = Func(1, 1, i)
                Exit For
            End If
        Next
        '************************************
        Perem(1, 1, IndexStatusPerem) = "Обработка главной функции начата!"
        ExecuteCode(MainFunction) 'Отправляем на выполнение "Главную" функцию
    End Sub
    'Выполнение программного кода
    Private Sub ExecuteCode(ByVal ModuleCode As RichTextBox)
        Perem(1, 1, IndexStatusPerem) = "Выполнение программного кода!"

        'ОБРАБОТКА Условий и Циклов
        Dim LevelOpenClose As Integer = 0       'Определение уровня в режиме реального времени
        Dim IfAndCycle(1, 1, 1, 100) As String  'Структура: --ЦИКЛ--ЕСЛИ--ИНАЧЕ--УРОВЕНЬ--
        Dim CycleCode As New RichTextBox
        '--------------------------

        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        Dim iSTROKA As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String

        For StrokaRUN As Integer = 0 To iSTROKA
            STROKA = ModuleCode.Lines.GetValue(StrokaRUN)

            '//ТИП СТРОКИ//////////////////////////////////////
            Dim Test As String = STROKA
            If (Test.Length > 6) Then 'для "ЕСЛИ" и "ДЛЯ("
                Test = Test.Remove(0, 1)
                Test = Test.Remove(4, Test.Length - 4)
            Else
                If (Test.Length > 5) Then 'для "ИНАЧЕ"
                    Test = Test.Remove(0, 1)
                    Test = Test.Remove(5, Test.Length - 5)
                Else
                    If (Test.Length > 4) Then 'для "ДЛЯ"
                        Test = Test.Remove(0, 1)
                        Test = Test.Remove(3, Test.Length - 3)
                    Else
                        If (Test.Length = 1) Then 'для "{" и "}"
                            If (Test = "{") Then    'ОТКРЫТИЕ УРОВНЯ
                                LevelOpenClose = LevelOpenClose + 1
                                IfAndCycle(1, 1, 1, LevelOpenClose) = LevelOpenClose.ToString   'УРОВЕНЬ
                            End If
                            If (Test = "}") Then    'ЗАКРЫТИЕ УРОВНЯ
                                LevelOpenClose = LevelOpenClose - 1
                                IfAndCycle(1, 1, 1, LevelOpenClose) = LevelOpenClose.ToString   'УРОВЕНЬ
                            End If
                        End If
                    End If
                End If
            End If

            '//УСЛОВИЯ////////////////////////////////////////
            If (Test = "Если" Or Test = "если") Then
                Perem(1, 1, IndexStatusPerem) = "Выполнение условия!"

                If (VerificationIF(STROKA) = True) Then
                    IfAndCycle(0, 0, 0, LevelOpenClose + 1) = "NULL"    'НЕ ЦИКЛ
                    IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "ДА"      'УСЛОВИЕ = ИСТИНО
                    IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "НЕТ"     'ИНАЧЕ = ЛОЖНО
                    IfAndCycle(1, 1, 1, LevelOpenClose + 1) = ""        'УРОВЕНЬ
                Else
                    IfAndCycle(0, 0, 0, LevelOpenClose + 1) = "NULL"    'НЕ ЦИКЛ
                    IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "НЕТ"     'УСЛОВИЕ = ЛОЖНО
                    IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "ДА"      'ИНАЧЕ = ИСТИНО
                    IfAndCycle(1, 1, 1, LevelOpenClose + 1) = ""        'УРОВЕНЬ
                End If
            End If
            If (Test = "Иначе" Or Test = "иначе") Then  'МЕНЯЕМ ЛОГИКУ
                If (IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "ДА") Then 'ЕСЛИ УСЛОВИЕ = ИСТИНО
                    IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "НЕТ" 'УСЛОВИЕ = ЛОЖНО
                    IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "ДА"  'ИНАЧЕ = ИСТИНО
                Else
                    If (IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "НЕТ") Then 'ЕСЛИ УСЛОВИЕ = ЛОЖНО
                        IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "ДА" 'УСЛОВИЕ = ИСТИНО
                        IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "НЕТ"  'ИНАЧЕ = ЛОЖНО
                    End If
                End If

            End If

            '//ЦИКЛЫ//////////////////////////////////////////
            If (Test = "Для(" Or Test = "для(") Then
                Perem(1, 1, IndexStatusPerem) = "Выполнение цикла!"

                'ПРОВЕРКА ПРЕДЫДУЩИХ УРОВНЕЙ УСЛОВИЙ
                Dim All_TRUE As Boolean = True 'ВСЕ ПРЕДЫДУЩИЕ ДА/НЕТ
                If (LevelOpenClose > 0) Then 'ПРОВЕРКА ПРЕДЫДУЩИХ УСЛОВИЙ
                    For iTestALL As Integer = 1 To LevelOpenClose
                        If (IfAndCycle(1, 0, 0, iTestALL) = "НЕТ") Then
                            All_TRUE = False
                            Exit For
                        End If
                    Next
                End If

                Dim CodeLine As String = STROKA
                Dim FOR_OpenClose As Integer = 1
                Dim JampLine As Integer = 0
                CycleCode.Text = ""

                For iFOR As Integer = StrokaRUN + 2 To iSTROKA
                    STROKA = ModuleCode.Lines.GetValue(iFOR)
                    If (STROKA = "{") Then FOR_OpenClose = FOR_OpenClose + 1
                    If (STROKA = "}") Then FOR_OpenClose = FOR_OpenClose - 1
                    If (FOR_OpenClose = 0) Then
                        JampLine = iFOR
                        Exit For
                    Else
                        CycleCode.Text = CycleCode.Text + System.Environment.NewLine + STROKA
                    End If
                Next
                StrokaRUN = JampLine
                If (All_TRUE = True) Then ExecuteFOR(CodeLine, CycleCode)
                STROKA = ""
            End If


            '//ВЫПОЛНЕНИЕ////////////////////////////////////
            If (LevelOpenClose = 0) Then    'НЕТ УСЛОВИЙ И ЦИКЛОВ

                'Определяем ТИПЫ операторов в данной строке
                TYPE(STROKA)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                '//ПРИСВОИТЬ//ФУНКЦИЯ/////////////////////////////
                If (TYPE_prisvoit = True) Then
                    Perem(1, 1, IndexStatusPerem) = "Присвоить значение!"
                    SetVariable(STROKA)
                End If

                '//СООБЩИТЬ//ФУНКЦИЯ//////////////////////////////
                If (TYPE_message = True) Then
                    ShowMessage(STROKA)
                End If

                '//ПРОЦЕДУРА//////////////////////////////////////
                If (TYPE_procedure_RUN = True) Then
                    Perem(1, 1, IndexStatusPerem) = "Обработка процедуры!"
                    ExecutePROC(STROKA)
                End If

            Else    'УЧИТЫВАЮТСЯ УСЛОВИЯ
                If (IfAndCycle(1, 0, 0, LevelOpenClose) = "ДА") Then

                    Dim All_TRUE As Boolean = True 'ВСЕ ПРЕДЫДУЩИЕ ДА/НЕТ
                    If (LevelOpenClose > 0) Then 'ПРОВЕРКА ПРЕДЫДУЩИХ УСЛОВИЙ
                        For iTestALL As Integer = 1 To LevelOpenClose
                            If (IfAndCycle(1, 0, 0, iTestALL) = "НЕТ") Then
                                All_TRUE = False
                                Exit For
                            End If
                        Next
                    End If


                    If (All_TRUE = True) Then
                        'Определяем ТИПЫ операторов в данной строке
                        TYPE(STROKA)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                        '//ПРИСВОИТЬ//ФУНКЦИЯ/////////////////////////////
                        If (TYPE_prisvoit = True) Then
                            Perem(1, 1, IndexStatusPerem) = "Присвоить значение!"
                            SetVariable(STROKA)
                        End If

                        '//СООБЩИТЬ//ФУНКЦИЯ//////////////////////////////
                        If (TYPE_message = True) Then
                            ShowMessage(STROKA)
                        End If

                        '//ПРОЦЕДУРА//////////////////////////////////////
                        If (TYPE_procedure_RUN = True) Then
                            Perem(1, 1, IndexStatusPerem) = "Обработка процедуры!"
                            ExecutePROC(STROKA)
                        End If
                    End If
                End If

            End If
        Next
    End Sub
    '--Присвоить значение переменной, выполнение вычислений
    Private Sub SetVariable(ByVal CodeLine As String)
        Perem(1, 1, IndexStatusPerem) = "Присвоить значение!"

        Dim Simvol As String 'символ из строки
        Dim Operator As String 'получаем конечный оператор
        Dim ActionPerem As String

        For iOperator As Integer = 0 To CodeLine.Length - 1
            Simvol = CodeLine.Chars(iOperator)
            If Simvol = " " Or Simvol = ";" Or Simvol = ";" Then 'Если символ пробел или точка с запятой значит оператор прочитан

                If Operator = "=" Then
                    'Отправляем имя переменной и значение на обработку
                    Dim ValueLine As String
                    ValueLine = CodeLine.Remove(0, iOperator)
                    If (AddValue(ActionPerem, ValueLine) = False) Then
                        Exit Sub
                    Else
                        'Произошла ошибка
                        SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! в строке [" + CodeLine + "] (не удалось выполнить присвоение)")
                        SomovCompileModule.WError.Show()
                        Exit Sub
                    End If
                Else
                    If (Operator <> " " And Operator <> "") Then ActionPerem = Operator 'имя переменной которой присвоено значение или выражение
                End If

                '-----------------------------
                Operator = ""
                Simvol = ""
            Else
                'не учитывая табуляцию (177580) группируем в ТОКЕН
                If (Simvol.GetHashCode <> 177580 And Simvol.GetHashCode <> -842352729 And Simvol <> " ") Then Operator = Operator + Simvol
            End If
        Next
    End Sub
    Private Function AddValue(ByVal NamePerem As String, ByVal CodeLine As String) As Boolean
        Perem(1, 1, IndexStatusPerem) = "Присвоить значение!"
        Try
            Dim TestTYPE As String
            Dim Token As String
            Dim Zapros As String
            Dim Result As String

            If (CodeLine.Chars(CodeLine.Length - 1) <> ";") Then
                'Произошла ошибка (предупредение)
                SomovCompileModule.WError.ListBox1.Items.Add("Предупреждение! в строке [" + CodeLine + "] (отсутствует символа конца строки ; точка с запятой)")
                SomovCompileModule.WError.Show()
            End If

            'ВЫЯВЛЕНИЕ ЗАПРОСА И ПРИСВОЕНОЙ СТРОКИ ***********************************
            For iTest As Integer = 0 To CodeLine.Length - 1
                'ЕСЛИ ЗАПРОС К ПОЛЬЗОВАТЕЛЮ ------------------------------------------
                If (Token = "Запрос(" Or Token = "запрос(") Then 'определил как запрос
                    TestTYPE = "Запрос" '!!! ТИП ДЕЙСТВИЯ НАД ЗНАЧЕНИЕМ
                    Dim kv As Integer = iTest + 1
                    If (CodeLine.Chars(kv) = """" Or CodeLine.Chars(kv) = "'") Then  'Параметр передан ввиде строки
                        Zapros = CodeLine.Remove(0, kv + 1)
                        Zapros = Zapros.Remove(Zapros.Length - 3, 3)
                        Result = InputBox(Zapros, "Запрос:")
                        Exit For
                    Else    'параметр передан в виде переменной
                        Zapros = CodeLine.Remove(0, kv)
                        Zapros = Zapros.Remove(Zapros.Length - 2, 2)
                        For iPerem As Integer = 1 To CORE.countPerem
                            If (Perem(0, 0, iPerem).ToString = Zapros) Then
                                Result = InputBox(Perem(1, 1, iPerem), "Запрос:")
                                Exit For
                            End If
                        Next
                    End If
                    Token = ""
                    Exit For
                Else
                    If (CodeLine.Chars(iTest) <> " " And CodeLine.Chars(iTest) <> "" And CodeLine.Chars(iTest) <> ";") Then Token = Token + CodeLine.Chars(iTest)
                End If
                '---------------------------------------------------------------------
                'ПРОСТОЕ ПРИСВАИВАНИЕ СТРОКИ -----------------------------------------
                If (CodeLine.Chars(1) = """" Or CodeLine.Chars(1) = "'") Then
                    TestTYPE = "Строка" '!!! ТИП ДЕЙСТВИЯ НАД ЗНАЧЕНИЕМ
                    'ПЕРЕМЕННОЙ БУДЕТ ПРИСВОЕНА СТРОКА
                    If (CodeLine.Chars(CodeLine.Length - 2) = """" Or CodeLine.Chars(CodeLine.Length - 2) = "'") Then
                        Result = CodeLine.Remove(0, 2)
                        Result = Result.Remove(Result.Length - 2, 2)
                        Exit For
                    End If
                End If
            Next
            'ЕСЛИ ЗНАЧЕНИЕ ЭТО ПЕРЕМЕННАЯ
            If (Token <> "") Then   'проверка значения на выражение переменной
                For i As Integer = 1 To CORE.countPerem
                    If (Perem(0, 0, i).ToString = Token) Then
                        Result = Perem(1, 1, i) 'значение из переменной
                        TestTYPE = Perem(0, 1, i) '!!! ТИП ДЕЙСТВИЯ НАД ЗНАЧЕНИЕМ
                        Token = ""
                        Exit For
                    End If
                Next
            End If
            '*************************************************************************

            'ВЫЯВЛЕНИЕ ЧИСЛА, ФУНКЦИИ ИЛИ МАТЕМАТИЧЕСКОГО ВЫРАЖЕНИЯ ******************
            If (TestTYPE <> "Запрос" And TestTYPE <> "Строка") Then
                Result = MATEMATIKA(CodeLine)
                If (Result = "ОШИБКА!ОШИБКА!ОШИБКА!") Then    'Проверка на ошибки
                    Return True 'произошла ошибка
                    Exit Function
                End If
                TestTYPE = "Неопределенное" '!!! ТИП ДЕЙСТВИЯ НАД ЗНАЧЕНИЕМ
            End If

            'ЗАПИСЫВАЕМ ЗНАЧЕНИЕ В ПЕРЕМЕННУЮ ***************************************
            For i As Integer = 1 To CORE.countPerem
                If (Perem(0, 0, i).ToString = NamePerem) Then
                    'проверка соответствия типа значения типу переменной
                    If (Perem(0, 1, i) = TestTYPE Or TestTYPE = "Запрос" Or TestTYPE = "Неопределенное") Then
                        Perem(1, 1, i) = Result
                        Return False    'безошибочное выполнение
                        Exit Function
                    Else
                        Return True 'произошла ошибка
                        Exit Function
                    End If
                    Exit For
                End If
            Next
            '**********************************************************************
        Catch ex As Exception
            Return True 'произошла ошибка
            Exit Function
        End Try
    End Function
    Private Function MATEMATIKA(ByVal DataLine As String) As String
        Perem(1, 1, IndexStatusPerem) = "Математические вычисления!"

        'ЕСЛИ ЗНАЧЕНИЕ НАЧИНАЕТСЯ С ЧИСЛА
        Dim Token As String
        Dim FuncName, FuncParam As String
        Dim Elements(0) As String
        Dim DataL As String = DataLine.Remove(0, 1)
        DataL = DataL.Remove(DataL.Length - 1, 1)

        'ВЫРАЖЕНИЕ ОПРЕДЕЛЕНО КАК МАТЕМАТИЧЕСКОЕ
        If (TYPE_matematika = True) Then
            For i As Integer = 0 To DataL.Length - 1
                If (DataL.Chars(i) = " ") Then
                    ReDim Preserve Elements(Elements.Length)
                    Elements(Elements.Length - 1) = Token
                    Token = ""
                Else
                    If (DataL.Chars(i) <> " " And DataL.Chars(i) <> "") Then Token = Token + DataL.Chars(i)
                End If
                If (DataL.Length - 1 = i) Then  'Конец строки
                    ReDim Preserve Elements(Elements.Length)
                    Elements(Elements.Length - 1) = Token
                    Token = ""
                End If
            Next
        Else    'ВЫРАЖЕНИЕ НЕ МАТЕМАТИЧЕСКОЕ
            'ЕСЛИ ЗНАЧЕНИЕ НАЧИНАЕТСЯ С ЧИСЛА
            If (DataL.Chars(0) = "0" Or DataL.Chars(0) = "1" Or DataL.Chars(0) = "2" Or DataL.Chars(0) = "3" Or DataL.Chars(0) = "4" Or DataL.Chars(0) = "5" Or DataL.Chars(0) = "6" Or DataL.Chars(0) = "7" Or DataL.Chars(0) = "8" Or DataL.Chars(0) = "9" Or DataL.Chars(0) = "0") Then
                Return DataL
                Exit Function

            Else    'ЕСЛИ ЗНАЧЕНИЕ НАЧИНАЕТСЯ С СИМВОЛА

                'ПРОСМОТР ПЕРЕМЕННЫХ--------------------
                For i As Integer = 1 To CORE.countPerem
                    If (Perem(0, 0, i).ToString = DataL) Then
                        'проверка соответствия типа значения типу переменной
                        If (Perem(0, 1, i) = "Число") Then
                            Return Perem(1, 1, i)
                            Exit Function
                        Else
                            'произошла ошибка
                            Return "ОШИБКА!ОШИБКА!ОШИБКА!"
                            Exit Function
                        End If
                        Exit For
                    End If
                Next
                'ПРОСМОТР ФУНКЦИЙ------------------------------------------------------
                For i As Integer = 0 To DataL.Length - 1
                    If (DataL.Chars(i) = "(") Then
                        FuncName = DataL.Remove(i, DataL.Length - i)
                        FuncParam = DataL.Remove(0, i + 2)
                        If (FuncParam <> "") Then FuncParam = FuncParam.Remove(FuncParam.Length - 1, 1)
                        Exit For
                    End If
                Next
                For i As Integer = 1 To CORE.countFunc
                    If (Func(0, 0, i).ToString = FuncName) Then
                        Return ExecuteFUNC(FuncName, FuncParam, Func(0, 1, i).ToString, Func(1, 1, i).ToString)
                        Exit Function
                    End If
                Next
                '------------------------------------------------
            End If
        End If

        'ВЫПОЛНЯЕМ МАТЕМАТИЧЕСКИЕ РАСЧЁТЫ************************
        If (TYPE_matematika = True) Then
            'ОКОНЧАТЕЛЬНОЕ ВЫЧИСЛЕНИЕ

            Dim SUMMA As Double
            Dim Element1, Element2 As Double

            'ПРИОРИТЕТНОЕ ВЫЧИСЛЕНИЕ ****************************************************
            For iMat As Integer = 1 To Elements.Length - 1
                If (Elements(iMat) = "*" Or Elements(iMat) = "/") Then 'УМНОЖЕНИЕ ИЛИ ДЕЛЕНИЕ
                    '--ЧИСЛО СЛЕВА ОТ ЗНАКА -------------------------------------------------
                    'ЕСЛИ НАЧИНАЕТСЯ С ЧИСЛА - ЗНАЧИТ ЭТО ЧИСЛО А НЕ ПЕРЕМЕННАЯ
                    If (Elements(iMat - 1).Chars(0) = "0" Or Elements(iMat - 1).Chars(0) = "1" Or Elements(iMat - 1).Chars(0) = "2" Or Elements(iMat - 1).Chars(0) = "3" Or Elements(iMat - 1).Chars(0) = "4" Or Elements(iMat - 1).Chars(0) = "5" Or Elements(iMat - 1).Chars(0) = "6" Or Elements(iMat - 1).Chars(0) = "7" Or Elements(iMat - 1).Chars(0) = "8" Or Elements(iMat - 1).Chars(0) = "9" Or Elements(iMat - 1).Chars(0) = "0") Then
                        'ЕСЛИ ЧИСЛО ДРОБНОЕ
                        Dim ElementD As String
                        For iChar As Integer = 0 To Elements(iMat - 1).Length - 1
                            If (Elements(iMat - 1).Chars(iChar) = ".") Then
                                ElementD = ElementD + ","
                            Else
                                ElementD = ElementD + Elements(iMat - 1).Chars(iChar)
                            End If
                        Next
                        Element1 = ElementD
                        ElementD = ""
                    Else    'ЭТО СЛУЧАЙ С ПЕРЕМЕННОЙ
                        'ПРОСМОТР ПЕРЕМЕННЫХ--------------------
                        For i As Integer = 1 To CORE.countPerem
                            If (Perem(0, 0, i).ToString = Elements(iMat - 1)) Then
                                'проверка соответствия типа значения типу переменной
                                If (Perem(0, 1, i) = "Число") Then
                                    'ЕСЛИ ЧИСЛО ДРОБНОЕ
                                    Dim ElementD As String
                                    For iChar As Integer = 0 To Perem(1, 1, i).Length - 1
                                        If (Perem(1, 1, i).Chars(iChar) = ".") Then
                                            ElementD = ElementD + ","
                                        Else
                                            ElementD = ElementD + Perem(1, 1, i).Chars(iChar)
                                        End If
                                    Next
                                    Element1 = ElementD
                                    ElementD = ""
                                Else
                                    Element1 = 0
                                End If
                                Exit For
                            End If
                        Next
                    End If
                    '---------------------------------------------------------------------
                    '--ЧИСЛО СПРАВА ОТ ЗНАКА ---------------------------------------------
                    'ЕСЛИ НАЧИНАЕТСЯ С ЧИСЛА - ЗНАЧИТ ЭТО ЧИСЛО А НЕ ПЕРЕМЕННАЯ
                    If (Elements(iMat + 1).Chars(0) = "0" Or Elements(iMat + 1).Chars(0) = "1" Or Elements(iMat + 1).Chars(0) = "2" Or Elements(iMat + 1).Chars(0) = "3" Or Elements(iMat + 1).Chars(0) = "4" Or Elements(iMat + 1).Chars(0) = "5" Or Elements(iMat + 1).Chars(0) = "6" Or Elements(iMat + 1).Chars(0) = "7" Or Elements(iMat + 1).Chars(0) = "8" Or Elements(iMat + 1).Chars(0) = "9" Or Elements(iMat + 1).Chars(0) = "0") Then
                        'ЕСЛИ ЧИСЛО ДРОБНОЕ
                        Dim ElementD As String
                        For iChar As Integer = 0 To Elements(iMat + 1).Length - 1
                            If (Elements(iMat + 1).Chars(iChar) = ".") Then
                                ElementD = ElementD + ","
                            Else
                                ElementD = ElementD + Elements(iMat + 1).Chars(iChar)
                            End If
                        Next
                        Element2 = ElementD
                        ElementD = ""
                    Else    'ЭТО СЛУЧАЙ С ПЕРЕМЕННОЙ
                        'ПРОСМОТР ПЕРЕМЕННЫХ--------------------
                        For i As Integer = 1 To CORE.countPerem
                            If (Perem(0, 0, i).ToString = Elements(iMat + 1)) Then
                                'проверка соответствия типа значения типу переменной
                                If (Perem(0, 1, i) = "Число") Then
                                    'ЕСЛИ ЧИСЛО ДРОБНОЕ
                                    Dim ElementD As String
                                    For iChar As Integer = 0 To Perem(1, 1, i).Length - 1
                                        If (Perem(1, 1, i).Chars(iChar) = ".") Then
                                            ElementD = ElementD + ","
                                        Else
                                            ElementD = ElementD + Perem(1, 1, i).Chars(iChar)
                                        End If
                                    Next
                                    Element2 = ElementD
                                    ElementD = ""
                                Else
                                    Element2 = 0
                                End If
                                Exit For
                            End If
                        Next
                    End If
                    '------------------------------------------------------------------------
                    'ВЫПОЛНЯЕМ ВЫЧИСЛЕНИЕ
                    If (Elements(iMat) = "*") Then
                        SUMMA = Element1 * Element2
                        Elements(iMat - 1) = "NULL"
                        Elements(iMat) = "NULL"
                        Elements(iMat + 1) = SUMMA.ToString
                        SUMMA = 0
                        Element1 = 0
                        Element2 = 0
                    Else
                        If (Elements(iMat) = "/") Then
                            SUMMA = Element1 / Element2
                            Elements(iMat - 1) = "NULL"
                            Elements(iMat) = "NULL"
                            Elements(iMat + 1) = SUMMA.ToString
                            SUMMA = 0
                            Element1 = 0
                            Element2 = 0
                        End If
                    End If
                End If
            Next

            'ОКОНЧАТЕЛЬНОЕ ВЫЧИСЛЕНИЕ
            'Математические действия
            Dim TYPE_plus As Boolean = False
            Dim TYPE_minus As Boolean = False
            Dim TYPE_result As Boolean = False
            For iMat As Integer = 1 To Elements.Length - 1
                If (Elements(iMat) = "+") Then TYPE_plus = True
                If (Elements(iMat) = "-") Then TYPE_minus = True
                If (Elements(iMat) <> "NULL" And Elements(iMat) <> "+" And Elements(iMat) <> "-") Then
                    If (TYPE_plus = False And TYPE_minus = False) Then 'ЧИСЛО СЛЕВА ОТ ЗНАКА
                        'ЕСЛИ НАЧИНАЕТСЯ С ЧИСЛА - ЗНАЧИТ ЭТО ЧИСЛО А НЕ ПЕРЕМЕННАЯ
                        If (Elements(iMat).Chars(0) = "0" Or Elements(iMat).Chars(0) = "1" Or Elements(iMat).Chars(0) = "2" Or Elements(iMat).Chars(0) = "3" Or Elements(iMat).Chars(0) = "4" Or Elements(iMat).Chars(0) = "5" Or Elements(iMat).Chars(0) = "6" Or Elements(iMat).Chars(0) = "7" Or Elements(iMat).Chars(0) = "8" Or Elements(iMat).Chars(0) = "9" Or Elements(iMat).Chars(0) = "0") Then
                            'ЕСЛИ ЧИСЛО ДРОБНОЕ
                            Dim ElementD As String
                            For iChar As Integer = 0 To Elements(iMat).Length - 1
                                If (Elements(iMat).Chars(iChar) = ".") Then
                                    ElementD = ElementD + ","
                                Else
                                    ElementD = ElementD + Elements(iMat).Chars(iChar)
                                End If
                            Next
                            Element1 = ElementD
                            ElementD = ""
                        Else    'ЭТО СЛУЧАЙ С ПЕРЕМЕННОЙ
                            'ПРОСМОТР ПЕРЕМЕННЫХ--------------------
                            For i As Integer = 1 To CORE.countPerem
                                If (Perem(0, 0, i).ToString = Elements(iMat)) Then
                                    'проверка соответствия типа значения типу переменной
                                    If (Perem(0, 1, i) = "Число") Then
                                        'ЕСЛИ ЧИСЛО ДРОБНОЕ
                                        Dim ElementD As String
                                        For iChar As Integer = 0 To Perem(1, 1, i).Length - 1
                                            If (Perem(1, 1, i).Chars(iChar) = ".") Then
                                                ElementD = ElementD + ","
                                            Else
                                                ElementD = ElementD + Perem(1, 1, i).Chars(iChar)
                                            End If
                                        Next
                                        Element1 = ElementD
                                        ElementD = ""
                                    Else
                                        Element1 = 0
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    Else 'ЧИСЛО СПРАВА ОТ ЗНАКА
                        TYPE_result = True
                        'ЕСЛИ НАЧИНАЕТСЯ С ЧИСЛА - ЗНАЧИТ ЭТО ЧИСЛО А НЕ ПЕРЕМЕННАЯ
                        If (Elements(iMat).Chars(0) = "0" Or Elements(iMat).Chars(0) = "1" Or Elements(iMat).Chars(0) = "2" Or Elements(iMat).Chars(0) = "3" Or Elements(iMat).Chars(0) = "4" Or Elements(iMat).Chars(0) = "5" Or Elements(iMat).Chars(0) = "6" Or Elements(iMat).Chars(0) = "7" Or Elements(iMat).Chars(0) = "8" Or Elements(iMat).Chars(0) = "9" Or Elements(iMat).Chars(0) = "0") Then
                            'ЕСЛИ ЧИСЛО ДРОБНОЕ
                            Dim ElementD As String
                            For iChar As Integer = 0 To Elements(iMat).Length - 1
                                If (Elements(iMat).Chars(iChar) = ".") Then
                                    ElementD = ElementD + ","
                                Else
                                    ElementD = ElementD + Elements(iMat).Chars(iChar)
                                End If
                            Next
                            Element2 = ElementD
                            ElementD = ""
                        Else    'ЭТО СЛУЧАЙ С ПЕРЕМЕННОЙ
                            'ПРОСМОТР ПЕРЕМЕННЫХ--------------------
                            For i As Integer = 1 To CORE.countPerem
                                If (Perem(0, 0, i).ToString = Elements(iMat)) Then
                                    'проверка соответствия типа значения типу переменной
                                    If (Perem(0, 1, i) = "Число") Then
                                        'ЕСЛИ ЧИСЛО ДРОБНОЕ
                                        Dim ElementD As String
                                        For iChar As Integer = 0 To Perem(1, 1, i).Length - 1
                                            If (Perem(1, 1, i).Chars(iChar) = ".") Then
                                                ElementD = ElementD + ","
                                            Else
                                                ElementD = ElementD + Perem(1, 1, i).Chars(iChar)
                                            End If
                                        Next
                                        Element2 = ElementD
                                        ElementD = ""
                                    Else
                                        Element2 = 0
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    'ВЫПОЛНЯЕМ ВЫЧИСЛЕНИЕ------------------------
                    If (TYPE_plus = True And TYPE_result = True) Then
                        If (Element1 = 0) Then Element1 = SUMMA
                        SUMMA = Element1 + Element2
                        Element1 = 0
                        Element2 = 0
                        TYPE_plus = False
                        TYPE_minus = False
                        TYPE_result = False
                    End If
                    If (TYPE_minus = True And TYPE_result = True) Then
                        If (Element1 = 0) Then Element1 = SUMMA
                        SUMMA = Element1 - Element2
                        Element1 = 0
                        Element2 = 0
                        TYPE_plus = False
                        TYPE_minus = False
                        TYPE_result = False
                    End If
                    '--------------------------------------------
                End If
            Next

            If (SUMMA = 0) Then
                If (Element1 <> 0) Then SUMMA = Element1
                If (Element2 <> 0) Then SUMMA = Element2
            End If

            'РЕЗУЛЬТАТ ВЫЧИСЛЕНИЯ
            Return SUMMA.ToString
            Exit Function
        End If
    End Function
    '--Сообщить (выводит сообщение пользователю)
    Private Sub ShowMessage(ByVal CodeLine As String)
        Dim DataL As String
        Dim FuncName, FuncParam As String
        Dim Token As String

        If (CodeLine.Chars(CodeLine.Length - 1) <> ";") Then
            'Произошла ошибка
            SomovCompileModule.WError.ListBox1.Items.Add("Предупреждение! в строке [" + CodeLine + "] (отсутствует символа конца строки ; точка с запятой)")
            SomovCompileModule.WError.Show()
            Exit Sub
        End If

        'ОПРЕДЕЛЯЕМ ПЕРЕДАННОЕ ЗНАЧЕНИЕ
        For i As Integer = 0 To CodeLine.Length - 1
            If (CodeLine.Chars(i) = " ") Then
                If (Token = "Сообщить(" Or Token = "сообщить(") Then
                    DataL = CodeLine.Remove(0, i + 1)
                    DataL = DataL.Remove(DataL.Length - 2, 2)
                    Token = ""
                    Exit For
                Else
                    Token = ""
                End If
            Else
                'не учитывая табуляцию (WinXP 177580, WinSeven -842352729) группируем в ТОКЕН
                If (CodeLine.GetHashCode <> 177580 And CodeLine.GetHashCode <> -842352729 And CodeLine.Chars(i) <> " ") Then Token = Token + CodeLine.Chars(i)
            End If
        Next

        '--------------------------------------------------------------------------
        'ВЫПОЛНЕНИЕ СООБЩЕНИЯ
        If (DataL.Chars(0) = """" Or DataL.Chars(0) = "'") Then 'ЗНАЧЕНИЕ СТРОКА
            If (DataL.Chars(DataL.Length - 1) = """" Or DataL.Chars(DataL.Length - 1) = "'" And CodeLine.Chars(CodeLine.Length - 1) = ";") Then
                DataL = DataL.Remove(0, 1)
                DataL = DataL.Remove(DataL.Length - 1, 1)
                MsgBox(DataL.ToString, MsgBoxStyle.OKOnly, "Сообщить:")
            Else
                'Произошла ошибка
                SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! в строке [" + CodeLine + "] (нет закрывающей кавычки или отсутствует символа конца строки ; точка с запятой)")
                SomovCompileModule.WError.Show()
                Exit Sub
            End If
        Else    'ЗНАЧЕНИЕ ПЕРЕМЕННАЯ ИЛИ ФУНКЦИЯ
            'ПРОСМОТР ПЕРЕМЕННЫХ---------------------------------------------------
            For i As Integer = 1 To CORE.countPerem
                If (Perem(0, 0, i).ToString = DataL) Then
                    If (CodeLine.Chars(CodeLine.Length - 1) = ";") Then
                        MsgBox(Perem(1, 1, i).ToString, MsgBoxStyle.OKOnly, "Сообщить:")
                        Exit Sub
                    Else
                        'Произошла ошибка
                        SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! в строке [" + CodeLine + "] (отсутствует символа конца строки ; точка с запятой)")
                        SomovCompileModule.WError.Show()
                    End If
                End If
            Next

            'ПРОСМОТР ФУНКЦИЙ------------------------------------------------------
            For i As Integer = 0 To DataL.Length - 1
                If (DataL.Chars(i) = "(") Then
                    FuncName = DataL.Remove(i, DataL.Length - i)
                    FuncParam = DataL.Remove(0, i + 2)
                    If (FuncParam <> "") Then FuncParam = FuncParam.Remove(FuncParam.Length - 1, 1)
                    Exit For
                End If
            Next
            For i As Integer = 1 To CORE.countFunc
                If (Func(0, 0, i).ToString = FuncName) Then
                    If (CodeLine.Chars(CodeLine.Length - 1) = ";") Then
                        MsgBox(ExecuteFUNC(FuncName, FuncParam, Func(0, 1, i).ToString, Func(1, 1, i).ToString), MsgBoxStyle.OKOnly, "Сообщить:")
                        Exit Sub
                    Else
                        'Произошла ошибка
                        SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! в строке [" + CodeLine + "] (отсутствует символа конца строки ; точка с запятой)")
                        SomovCompileModule.WError.Show()
                    End If

                End If
            Next
            '--------------------------------------------------
        End If
        Perem(1, 1, IndexStatusPerem) = "Вывод сообщений!"
    End Sub
    '--Выполнение ФУНКЦИИ
    Private Function ExecuteFUNC(ByVal NameFunction As String, ByVal ParamFunction As String, ByVal TypeFunction As String, ByVal CodeFunction As String) As String
        Perem(1, 1, IndexStatusPerem) = "Выполнение функции!"

        Dim Token, stroka As String
        Dim FuncFind As Boolean = False 'Отметка при выявлении заданой функции
        Dim Param(0) As String 'Массив имен переданных параметров
        Dim ParamIndex As Integer = 1
        Dim ParamBeginEnd As Boolean = False 'Определяет начало описания параметров и конец)
        Dim FuncReturn As Boolean = False   'оператор ВОЗВРАТ был или не был определён
        Dim ModuleCode As New RichTextBox
        ModuleCode.Text = CodeFunction

        'ВЫПОЛНЕНИЕ ПРИСВОЕНИЯ ЗНАЧЕНИЙ ПЕРЕДАННЫМ ПАРАМЕТРАМ
        If (ParamFunction <> "") Then
            'ИМЕНА ПЕРЕМЕННЫХ ПЕРЕДАННЫХ В КАЧЕСТВЕ ПАРАМЕТРОВ
            For i As Integer = 0 To ParamFunction.Length - 1
                If (ParamFunction.Chars(i) = " " Or ParamFunction.Chars(i) = "," Or i = ParamFunction.Length - 1) Then
                    If (i = ParamFunction.Length - 1 And ParamFunction.GetHashCode <> 177580 And ParamFunction.GetHashCode <> -842352729 And ParamFunction.Chars(i) <> " ") Then Token = Token + ParamFunction.Chars(i)
                    If (Token <> "") Then
                        'В СЛУЧАЕ ЕСЛИ ПАРАМЕТР ЭТО ЗНАЧЕНИЕ: ЧИСЛО
                        If (Token.Chars(0) = "0" Or Token.Chars(0) = "1" Or Token.Chars(0) = "2" Or Token.Chars(0) = "3" Or Token.Chars(0) = "4" Or Token.Chars(0) = "5" Or Token.Chars(0) = "6" Or Token.Chars(0) = "7" Or Token.Chars(0) = "8" Or Token.Chars(0) = "9" Or Token.Chars(0) = "0") Then
                            ReDim Preserve Param(Param.Length)
                            Param(Param.Length - 1) = Token
                        Else    'В СЛУЧАЕ ЕСЛИ ПАРАМЕТР ЭТО ЗНАЧЕНИЕ: СТРОКА
                            If (Token.Chars(0) = """" Or Token.Chars(0) = "'") Then
                                stroka = ParamFunction.Remove(0, 1)
                                stroka = stroka.Remove(stroka.Length - 1, 1)
                                ReDim Preserve Param(Param.Length)
                                Param(Param.Length - 1) = stroka
                                stroka = ""
                                Exit For

                            Else    'В СЛУЧАЕ ЕСЛИ ПАРАМЕТР ЭТО ПЕРЕМЕННАЯ
                                For iPerem As Integer = 1 To CORE.countPerem
                                    If (Perem(0, 0, iPerem).ToString = Token) Then
                                        ReDim Preserve Param(Param.Length)
                                        Param(Param.Length - 1) = Perem(1, 1, iPerem)
                                    End If
                                Next
                            End If
                        End If

                    End If
                    Token = ""
                Else
                    'не учитывая табуляцию (WinXP 177580, WinSeven -842352729) группируем в ТОКЕН
                    If (ParamFunction.GetHashCode <> 177580 And ParamFunction.GetHashCode <> -842352729 And ParamFunction.Chars(i) <> " ") Then Token = Token + ParamFunction.Chars(i)
                End If
            Next
            'ПРОСМОТР ТАБЛИЦЫ ИДЕНТИФИКАТОРОВ
            For i As Integer = 1 To TABLE_ID.Length - 1
                If (FuncFind = True And ParamBeginEnd = True And TABLE_ID(i) <> "Перем" And TABLE_ID(i) <> "перем" And TABLE_ID(i) <> "," And TABLE_ID(i) <> " " And TABLE_ID(i) <> "Тип" And TABLE_ID(i) <> "тип" And TABLE_ID(i) <> "Строка" And TABLE_ID(i) <> "строка" And TABLE_ID(i) <> "Число" And TABLE_ID(i) <> "число") Then
                    'ПРОСМОТР ПЕРЕМЕННЫХ И ИХ ЗНАЧЕНИЙ
                    For iPerem As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, iPerem).ToString = TABLE_ID(i)) Then
                            Perem(1, 1, iPerem) = Param(ParamIndex)
                            ParamIndex = ParamIndex + 1
                            Exit For
                        End If
                    Next
                End If
                If (TABLE_ID(i) = NameFunction) Then    'НАЙДЕНА ФУНКЦИЯ С УКАЗАНЫМ ИМЕНЕМ
                    FuncFind = True 'Функция определена
                End If
                If (TABLE_ID(i) = "(") Then
                    ParamBeginEnd = True 'ОТКРЫТО ЧТЕНИЕ ПАРАМЕТРОВ ФУНКЦИИ
                End If
                If (TABLE_ID(i) = ")") Then
                    ParamBeginEnd = False 'ЗАКРЫТО ЧТЕНИЕ ПАРАМЕТРОВ ФУНКЦИИ
                    'ОБРАБОТКА КОДА ФУНКЦИИ
                    ExecuteCode(ModuleCode)
                    'ВЫПОЛНЯЕМ ВОЗВРАТ ИЗ ФУНКЦИИ
                    'ПРОСМОТР ТАБЛИЦЫ ИДЕНТИФИКАТОРОВ
                    For iReturn As Integer = 1 To TABLE_ID.Length - 1
                        If (TABLE_ID(iReturn) = NameFunction) Then    'НАЙДЕНА ФУНКЦИЯ С УКАЗАНЫМ ИМЕНЕМ
                            FuncReturn = True
                        End If
                        If (TABLE_ID(iReturn) = "Возврат" Or TABLE_ID(i) = "возврат") Then
                            If (FuncReturn = True) Then
                                'ПРОСМОТР ПЕРЕМЕННЫХ И ИХ ЗНАЧЕНИЙ
                                For iPerem As Integer = 1 To CORE.countPerem
                                    If (Perem(0, 0, iPerem).ToString = TABLE_ID(iReturn + 1)) Then
                                        Return Perem(1, 1, iPerem).ToString
                                        Exit Function
                                    End If
                                Next
                                Return ""
                                Exit Function
                            End If
                        End If
                    Next
                End If
            Next
        Else    'ОБРАБОТКА ФУНКЦИИ БЕЗ ПАРАМЕТРОВ
            ExecuteCode(ModuleCode)
            'ВЫПОЛНЯЕМ ВОЗВРАТ ИЗ ФУНКЦИИ
            'ПРОСМОТР ТАБЛИЦЫ ИДЕНТИФИКАТОРОВ
            For i As Integer = 1 To TABLE_ID.Length - 1
                If (TABLE_ID(i) = NameFunction) Then    'НАЙДЕНА ФУНКЦИЯ С УКАЗАНЫМ ИМЕНЕМ
                    FuncReturn = True
                End If
                If (TABLE_ID(i) = "Возврат" Or TABLE_ID(i) = "возврат") Then
                    If (FuncReturn = True) Then
                        'ПРОСМОТР ПЕРЕМЕННЫХ И ИХ ЗНАЧЕНИЙ
                        For iPerem As Integer = 1 To CORE.countPerem
                            If (Perem(0, 0, iPerem).ToString = TABLE_ID(i + 1)) Then
                                Return Perem(1, 1, iPerem).ToString
                                Exit Function
                            End If
                        Next
                        Return ""
                        Exit Function
                    End If
                End If
            Next
        End If
    End Function
    '--Выполнить ПРОЦЕДУРЫ
    Private Sub ExecutePROC(ByVal NameParamProcedure As String)
        Perem(1, 1, IndexStatusPerem) = "Выполнение процедуры!"

        Dim Token, stroka As String
        Dim ProcFind As Boolean = False 'Отметка при выявлении заданой процедуры
        Dim Param(0) As String 'Массив имен переданных параметров
        Dim ParamIndex As Integer = 1
        Dim ParamBeginEnd As Boolean = False 'Определяет начало описания параметров и конец)

        Dim ProcName, ProcParam As String
        Dim ModuleCode As New RichTextBox

        'Убираем пробел в начале
        If (NameParamProcedure.Chars(0) = " ") Then NameParamProcedure = NameParamProcedure.Remove(0, 1)

        'Определяем имя процедуры
        For i As Integer = 0 To NameParamProcedure.Length - 1
            If (NameParamProcedure.Chars(i) = "(") Then
                ProcName = NameParamProcedure.Remove(i, NameParamProcedure.Length - i)
                ProcParam = NameParamProcedure.Remove(0, i + 2)
                If (ProcParam <> "") Then ProcParam = ProcParam.Remove(ProcParam.Length - 1, 1)
                Exit For
            End If
        Next
        'Определяем параметры переданные в процедуру
        For i As Integer = 1 To CORE.countProc
            If (Proc(0, 0, i).ToString = ProcName) Then
                ModuleCode.Text = Proc(1, 1, i).ToString
                Exit For
            End If
        Next

        'Если процедура с параметрами выполняем их обработку
        If (ProcParam <> "") Then
            'ИМЕНА ПЕРЕМЕННЫХ ПЕРЕДАННЫХ В КАЧЕСТВЕ ПАРАМЕТРОВ
            For i As Integer = 0 To ProcParam.Length - 1
                If (ProcParam.Chars(i) = " " Or ProcParam.Chars(i) = "," Or i = ProcParam.Length - 1) Then
                    If (i = ProcParam.Length - 1 And ProcParam.GetHashCode <> 177580 And ProcParam.GetHashCode <> -842352729 And ProcParam.Chars(i) <> " ") Then Token = Token + ProcParam.Chars(i)
                    If (Token <> "") Then
                        'В СЛУЧАЕ ЕСЛИ ПАРАМЕТР ЭТО ЗНАЧЕНИЕ: ЧИСЛО
                        If (Token.Chars(0) = "0" Or Token.Chars(0) = "1" Or Token.Chars(0) = "2" Or Token.Chars(0) = "3" Or Token.Chars(0) = "4" Or Token.Chars(0) = "5" Or Token.Chars(0) = "6" Or Token.Chars(0) = "7" Or Token.Chars(0) = "8" Or Token.Chars(0) = "9" Or Token.Chars(0) = "0") Then
                            ReDim Preserve Param(Param.Length)
                            Param(Param.Length - 1) = Token
                        Else    'В СЛУЧАЕ ЕСЛИ ПАРАМЕТР ЭТО ЗНАЧЕНИЕ: СТРОКА
                            If (Token.Chars(0) = """" Or Token.Chars(0) = "'") Then
                                stroka = ProcParam.Remove(0, 1)
                                stroka = stroka.Remove(stroka.Length - 1, 1)
                                If (stroka.Chars(stroka.Length - 1) = """" Or stroka.Chars(stroka.Length - 1) = "'") Then stroka = stroka.Remove(stroka.Length - 1, 1)
                                ReDim Preserve Param(Param.Length)
                                Param(Param.Length - 1) = stroka
                                stroka = ""
                                Exit For
                            Else    'В СЛУЧАЕ ЕСЛИ ПАРАМЕТР ЭТО ПЕРЕМЕННАЯ
                                If (Token.Chars(Token.Length - 1) = ")") Then Token = Token.Remove(Token.Length - 1, 1)
                                For iPerem As Integer = 1 To CORE.countPerem
                                    If (Perem(0, 0, iPerem).ToString = Token) Then
                                        ReDim Preserve Param(Param.Length)
                                        Param(Param.Length - 1) = Perem(1, 1, iPerem)
                                    End If
                                Next
                            End If
                        End If

                    End If
                    Token = ""
                Else
                    'не учитывая табуляцию (WinXP 177580, WinSeven -842352729) группируем в ТОКЕН
                    If (ProcParam.GetHashCode <> 177580 And ProcParam.GetHashCode <> -842352729 And ProcParam.Chars(i) <> " ") Then Token = Token + ProcParam.Chars(i)
                End If
            Next
            'ПРОСМОТР ТАБЛИЦЫ ИДЕНТИФИКАТОРОВ
            For i As Integer = 1 To TABLE_ID.Length - 1
                If (ProcFind = True And ParamBeginEnd = True And TABLE_ID(i) <> "Перем" And TABLE_ID(i) <> "перем" And TABLE_ID(i) <> "," And TABLE_ID(i) <> " " And TABLE_ID(i) <> "Тип" And TABLE_ID(i) <> "тип" And TABLE_ID(i) <> "Строка" And TABLE_ID(i) <> "строка" And TABLE_ID(i) <> "Число" And TABLE_ID(i) <> "число") Then
                    'ПРОСМОТР ПЕРЕМЕННЫХ И ИХ ЗНАЧЕНИЙ
                    For iPerem As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, iPerem).ToString = TABLE_ID(i)) Then
                            Perem(1, 1, iPerem) = Param(ParamIndex)
                            ParamIndex = ParamIndex + 1
                            Exit For
                        End If
                    Next
                End If
                If (TABLE_ID(i) = ProcName) Then    'НАЙДЕНА ПРОЦЕДУРА С УКАЗАНЫМ ИМЕНЕМ
                    ProcFind = True 'Функция определена
                End If
                If (TABLE_ID(i) = "(") Then
                    ParamBeginEnd = True 'ОТКРЫТО ЧТЕНИЕ ПАРАМЕТРОВ ПРОЦЕДУРЫ
                End If
                If (TABLE_ID(i) = ")") Then
                    ParamBeginEnd = False 'ЗАКРЫТО ЧТЕНИЕ ПАРАМЕТРОВ ПРОЦЕДУРЫ
                    Exit For
                End If
            Next
        End If

        'ОБРАБОТКА КОДА ПРОЦЕДУРЫ
        ExecuteCode(ModuleCode)
    End Sub
    '--Условие (проверка выполнения заданного условия)
    Private Function VerificationIF(ByVal CodeLine As String) As Boolean
        Perem(1, 1, IndexStatusPerem) = "Выполнение условий!"
        Try
            Dim STROKA As String = CodeLine
            Dim Token As String
            Dim KV As Boolean = False   'Определение начала и конца строки
            Dim TYPE_IF As String = ""  'тип условия
            Dim LEFT_OPERATOR, RIGHT_OPERATOR As String 'левый и правый оператор
            Dim LEFT_OPERATOR_D, RIGHT_OPERATOR_D As Double 'левый и правый оператор (ЧИСЛА)

            STROKA = STROKA.Remove(0, 5)

            For i As Integer = 0 To STROKA.Length - 1
                If (STROKA.Chars(i) = """" And KV = False) Then
                    KV = True 'Открывается кавычка
                Else
                    If (STROKA.Chars(i) = """" And KV = True) Then
                        KV = False  'Закрывается кавычка
                    End If
                End If

                If (STROKA.Chars(i) = "(" Or STROKA.Chars(i) = " " Or STROKA.Chars(i) = ")") Then
                    If (Token = "==") Then
                        TYPE_IF = "=="
                        Token = ""
                    End If
                    If (Token = ">") Then
                        TYPE_IF = ">"
                        Token = ""
                    End If
                    If (Token = "<") Then
                        TYPE_IF = "<"
                        Token = ""
                    End If
                    If (Token = ">=") Then
                        TYPE_IF = ">="
                        Token = ""
                    End If
                    If (Token = "<=") Then
                        TYPE_IF = "<="
                        Token = ""
                    End If
                    If (Token = "!=") Then
                        TYPE_IF = "!="
                        Token = ""
                    End If

                    If (KV = True And STROKA.Chars(i) = " ") Then Token = Token + STROKA.Chars(i)
                    'ОПЕРАНДЫ УСЛОВИЯ -----------------------------
                    If (TYPE_IF = "") Then 'ОПРЕДЕЛЯЕМ ЛЕВЫЙ ОПЕРАНД
                        If (KV = False) Then LEFT_OPERATOR = Token
                    Else 'ОПРЕДЕЛЯЕМ ПРАВЫЙ ОПЕРАНД
                        If (KV = False) Then RIGHT_OPERATOR = Token
                    End If
                    '----------------------------------------------

                    If (KV = False) Then Token = ""
                Else
                    'не учитывая табуляцию (WinXP 177580, WinSeven -842352729) группируем в ТОКЕН
                    If (STROKA.GetHashCode <> 177580 And STROKA.GetHashCode <> -842352729 And STROKA.Chars(i) <> " ") Then Token = Token + STROKA.Chars(i)
                End If
            Next

            'ОПРЕДЕЛЯЕМ ТИП И ЗНАЧЕНИЕ ПОЛУЧЕННЫХ ОПЕРАНДОВ *************
            'ЛЕВЫЙ ОПЕРАНД ..............................................
            If (LEFT_OPERATOR.Chars(0) = "0" Or LEFT_OPERATOR.Chars(0) = "1" Or LEFT_OPERATOR.Chars(0) = "2" Or LEFT_OPERATOR.Chars(0) = "3" Or LEFT_OPERATOR.Chars(0) = "4" Or LEFT_OPERATOR.Chars(0) = "5" Or LEFT_OPERATOR.Chars(0) = "6" Or LEFT_OPERATOR.Chars(0) = "7" Or LEFT_OPERATOR.Chars(0) = "8" Or LEFT_OPERATOR.Chars(0) = "9" Or LEFT_OPERATOR.Chars(0) = "0") Then
                'ЗНАЧИТ ЭТО ЧИСЛО В ЧИСТОМ ВИДЕ
                'ЕСЛИ ЧИСЛО ДРОБНОЕ
                Dim ElementD As String = ""
                For iChar As Integer = 0 To LEFT_OPERATOR.Length - 1
                    If (LEFT_OPERATOR.Chars(iChar) = ".") Then
                        ElementD = ElementD + ","
                    Else
                        ElementD = ElementD + LEFT_OPERATOR.Chars(iChar)
                    End If
                Next
                LEFT_OPERATOR = ElementD
            Else
                If (LEFT_OPERATOR.Chars(0) = """" And LEFT_OPERATOR.Chars(LEFT_OPERATOR.Length - 1) = """") Then
                    'ЗНАЧИТ ЭТО СТРОКА В ЧИСТОМ ВИДЕ
                    LEFT_OPERATOR = LEFT_OPERATOR.Remove(0, 1)
                    LEFT_OPERATOR = LEFT_OPERATOR.Remove(LEFT_OPERATOR.Length - 1, 1)
                Else
                    'ЗНАЧИТ ЭТО ИМЯ ПЕРЕМЕННОЙ
                    For i As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, i).ToString = LEFT_OPERATOR) Then
                            LEFT_OPERATOR = Perem(1, 1, i).ToString 'ЗНАЧЕНИЕ ПЕРЕМЕННОЙ
                            If (LEFT_OPERATOR.Chars(0) = "0" Or LEFT_OPERATOR.Chars(0) = "1" Or LEFT_OPERATOR.Chars(0) = "2" Or LEFT_OPERATOR.Chars(0) = "3" Or LEFT_OPERATOR.Chars(0) = "4" Or LEFT_OPERATOR.Chars(0) = "5" Or LEFT_OPERATOR.Chars(0) = "6" Or LEFT_OPERATOR.Chars(0) = "7" Or LEFT_OPERATOR.Chars(0) = "8" Or LEFT_OPERATOR.Chars(0) = "9" Or LEFT_OPERATOR.Chars(0) = "0") Then
                                'ЗНАЧИТ ЭТО ЧИСЛО В ЧИСТОМ ВИДЕ
                                'ЕСЛИ ЧИСЛО ДРОБНОЕ
                                Dim ElementD As String = ""
                                For iChar As Integer = 0 To LEFT_OPERATOR.Length - 1
                                    If (LEFT_OPERATOR.Chars(iChar) = ".") Then
                                        ElementD = ElementD + ","
                                    Else
                                        ElementD = ElementD + LEFT_OPERATOR.Chars(iChar)
                                    End If
                                Next
                                LEFT_OPERATOR = ElementD
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
            'ПРАВЫЙ ОПЕРАНД ..............................................
            If (RIGHT_OPERATOR.Chars(0) = "0" Or RIGHT_OPERATOR.Chars(0) = "1" Or RIGHT_OPERATOR.Chars(0) = "2" Or RIGHT_OPERATOR.Chars(0) = "3" Or RIGHT_OPERATOR.Chars(0) = "4" Or RIGHT_OPERATOR.Chars(0) = "5" Or RIGHT_OPERATOR.Chars(0) = "6" Or RIGHT_OPERATOR.Chars(0) = "7" Or RIGHT_OPERATOR.Chars(0) = "8" Or RIGHT_OPERATOR.Chars(0) = "9" Or RIGHT_OPERATOR.Chars(0) = "0") Then
                'ЗНАЧИТ ЭТО ЧИСЛО В ЧИСТОМ ВИДЕ
                'ЕСЛИ ЧИСЛО ДРОБНОЕ
                Dim ElementD As String = ""
                For iChar As Integer = 0 To RIGHT_OPERATOR.Length - 1
                    If (RIGHT_OPERATOR.Chars(iChar) = ".") Then
                        ElementD = ElementD + ","
                    Else
                        ElementD = ElementD + RIGHT_OPERATOR.Chars(iChar)
                    End If
                Next
                RIGHT_OPERATOR = ElementD
            Else
                If (RIGHT_OPERATOR.Chars(0) = """" And RIGHT_OPERATOR.Chars(RIGHT_OPERATOR.Length - 1) = """") Then
                    'ЗНАЧИТ ЭТО СТРОКА В ЧИСТОМ ВИДЕ
                    RIGHT_OPERATOR = RIGHT_OPERATOR.Remove(0, 1)
                    RIGHT_OPERATOR = RIGHT_OPERATOR.Remove(RIGHT_OPERATOR.Length - 1, 1)
                Else
                    'ЗНАЧИТ ЭТО ИМЯ ПЕРЕМЕННОЙ
                    For i As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, i).ToString = RIGHT_OPERATOR) Then
                            RIGHT_OPERATOR = Perem(1, 1, i).ToString 'ЗНАЧЕНИЕ ПЕРЕМЕННОЙ
                            If (RIGHT_OPERATOR.Chars(0) = "0" Or RIGHT_OPERATOR.Chars(0) = "1" Or RIGHT_OPERATOR.Chars(0) = "2" Or RIGHT_OPERATOR.Chars(0) = "3" Or RIGHT_OPERATOR.Chars(0) = "4" Or RIGHT_OPERATOR.Chars(0) = "5" Or RIGHT_OPERATOR.Chars(0) = "6" Or RIGHT_OPERATOR.Chars(0) = "7" Or RIGHT_OPERATOR.Chars(0) = "8" Or RIGHT_OPERATOR.Chars(0) = "9" Or RIGHT_OPERATOR.Chars(0) = "0") Then
                                'ЕСЛИ ЧИСЛО ДРОБНОЕ
                                Dim ElementD As String = ""
                                For iChar As Integer = 0 To RIGHT_OPERATOR.Length - 1
                                    If (RIGHT_OPERATOR.Chars(iChar) = ".") Then
                                        ElementD = ElementD + ","
                                    Else
                                        ElementD = ElementD + RIGHT_OPERATOR.Chars(iChar)
                                    End If
                                Next
                                RIGHT_OPERATOR = ElementD
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
            'ВЫПОЛНЕНИЕ ***********************************************
            If (TYPE_IF = "==") Then
                If (LEFT_OPERATOR = RIGHT_OPERATOR) Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            End If
            If (TYPE_IF = ">") Then
                LEFT_OPERATOR_D = LEFT_OPERATOR
                RIGHT_OPERATOR_D = RIGHT_OPERATOR
                If (LEFT_OPERATOR_D > RIGHT_OPERATOR_D) Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            End If
            If (TYPE_IF = "<") Then
                LEFT_OPERATOR_D = LEFT_OPERATOR
                RIGHT_OPERATOR_D = RIGHT_OPERATOR
                If (LEFT_OPERATOR_D < RIGHT_OPERATOR_D) Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            End If
            If (TYPE_IF = ">=") Then
                LEFT_OPERATOR_D = LEFT_OPERATOR
                RIGHT_OPERATOR_D = RIGHT_OPERATOR
                If (LEFT_OPERATOR_D >= RIGHT_OPERATOR_D) Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            End If
            If (TYPE_IF = "<=") Then
                LEFT_OPERATOR_D = LEFT_OPERATOR
                RIGHT_OPERATOR_D = RIGHT_OPERATOR
                If (LEFT_OPERATOR_D <= RIGHT_OPERATOR_D) Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            End If
            If (TYPE_IF = "!=") Then
                If (LEFT_OPERATOR <> RIGHT_OPERATOR) Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            End If

        Catch ex As Exception
            'Произошла ошибка
            SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! в строке [" + CodeLine + "] (неверно задано условие)")
            SomovCompileModule.WError.Show()
            Exit Function
        End Try
    End Function
    '--Обработка заданного условия ЦИКЛА "ДЛЯ" и ВЫПОЛНЕНИЕ ЕГО
    Private Sub ExecuteFOR(ByVal CodeLine As String, ByVal ModuleCode As RichTextBox)
        Perem(1, 1, IndexStatusPerem) = "Выполнение цикла!"
        Try
            Dim STROKA As String = CodeLine
            Dim Token As String
            Dim Ravno As Boolean = False
            Dim PO As Boolean = False
            Dim ActionPerem As String = ""      'Имя переменной
            Dim StartNumCycle As Integer = 0    'Стартовое значение цикла
            Dim KolCycle As Integer = 0         'Количество циклов

            TYPE_prisvoit = False
            STROKA = STROKA.Remove(0, 6)
            STROKA = STROKA.Remove(STROKA.Length - 1, 1)

            For i As Integer = 0 To STROKA.Length - 1
                If (STROKA.Chars(i) = " " Or i = STROKA.Length - 1) Then

                    If (PO = False) Then    'ОПЕРАТОР ПО

                        If (Ravno = False) Then 'ЗНАК РАВЕНСТВА
                            If (ActionPerem = "") Then ActionPerem = Token
                        Else
                            '**СИНТАКСИЧЕСКАЯ ПРОВЕРКА*****************
                            If (SyntaxPerem(ActionPerem) = True) Then
                                For iPerem As Integer = 1 To CORE.countPerem
                                    If (Perem(0, 0, iPerem).ToString = ActionPerem) Then
                                        'ЗНАЧЕНИЕ ПЕРЕМЕННОЙ
                                        Perem(1, 1, iPerem) = Token
                                        StartNumCycle = Token
                                        Ravno = False
                                        Exit For
                                    End If
                                Next
                            Else
                                '**ЗАПОЛНЕНИЕ ТАБЛИЦЫ *********************
                                countPerem = countPerem + 1 'учет количества переменных
                                ReDim Preserve Perem(1, 1, countPerem) 'Увеличиваем массив
                                '**ИМЯ ПЕРЕМЕННОЙ**************************
                                Perem(0, 0, countPerem) = ActionPerem
                                '**ТИП ЧИСЛО*******************************
                                Perem(0, 1, countPerem) = "Число"
                                'ЗНАЧЕНИЕ ПЕРЕМЕННОЙ
                                Perem(1, 1, countPerem) = Token
                                StartNumCycle = Token
                                Ravno = False
                            End If

                        End If
                    Else
                        'не учитывая табуляцию (WinXP 177580, WinSeven -842352729) группируем в ТОКЕН
                        If (STROKA.GetHashCode <> 177580 And STROKA.GetHashCode <> -842352729 And STROKA.Chars(i) <> " ") Then Token = Token + STROKA.Chars(i)
                        KolCycle = Token
                        Exit For
                    End If
                    If (Token = "=") Then Ravno = True
                    If (Token = "По" Or Token = "по") Then PO = True
                    Token = ""
                Else
                    'не учитывая табуляцию (WinXP 177580, WinSeven -842352729) группируем в ТОКЕН
                    If (STROKA.GetHashCode <> 177580 And STROKA.GetHashCode <> -842352729 And STROKA.Chars(i) <> " ") Then Token = Token + STROKA.Chars(i)
                End If
            Next

            'ВЫПОЛНЕНИЕ ЦИКЛА
            For iExecute As Integer = StartNumCycle To KolCycle
                For iPerem As Integer = 1 To CORE.countPerem
                    If (Perem(0, 0, iPerem).ToString = ActionPerem) Then
                        'ЗНАЧЕНИЕ ПЕРЕМЕННОЙ
                        Perem(1, 1, countPerem) = iExecute.ToString
                        Exit For
                    End If
                Next
                ExecuteCode(ModuleCode)
            Next

        Catch ex As Exception
            'Произошла ошибка
            SomovCompileModule.WError.ListBox1.Items.Add("Ошибка! в строке [" + CodeLine + "] (неверно задано условие)")
            SomovCompileModule.WError.Show()
            Exit Sub
        End Try
    End Sub


    '//ОЧИСТКА///////////////////////////////////////////////////////////////////
    Private Sub ClearTYPE() 'Очистка
        TYPE_peremen = False
        TYPE_chislo = False
        TYPE_stroka = False
        TYPE_object = False
        TYPE_prisvoit = False
        TYPE_message = False
        TYPE_zapros = False
        TYPE_question = False
        TYPE_procedure = False
        TYPE_procedure_Param = False
        TYPE_procedure_RUN = False
        TYPE_function = False
        TYPE_function_Param = False
        TYPE_matematika = False
    End Sub

    Public Sub ClearModule() 'Очистка модуля.
        ReDim CORE.TABLE_ID(0)
        ReDim CORE.Perem(1, 1, 0)
        ReDim CORE.Proc(1, 1, 0)
        ReDim CORE.Func(1, 1, 0)
        countPerem = 0
        countProc = 0
        countFunc = 0
        MultilineComent = False
        ClearTYPE()
        ModuleCodeNoneComent.Clear()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Module
