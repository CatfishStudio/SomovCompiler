Module CORE
    '//���������� � ������////////////////////////////////////////////////////////
    Private MultilineComent As Boolean = False
    Private ModuleCodeNoneComent As New RichTextBox '������ ��� �����������

    '//��������� �����������//////////////////////////////////////////////////////
    Private LISTOperators As New ListBox '�������� ���� ����������������� ����������



    '//������� ���������������////////////////////////////////////////////////////
    Private TABLE_ID(0) As String        '��� �������� ������.
    '���������� ------------------------------
    '-----------���-���-��������--------------
    Public Perem(1, 1, 0) As String     '����������
    Public countPerem As Integer        '���������� ����������
    Public IndexStatusPerem As Integer  '������ ���������� "���������"
    '-----------���-���-����--------
    Public Func(1, 1, 0) As String      '�������
    Public countFunc As Integer         '���������� �������
    '-----------���-| |-����-------------
    Public Proc(1, 1, 0) As String      '���������
    Public countProc As Integer         '���������� ��������

    '���������� ����� ������� ��������������� ���������� ����
    Private Sub AddTableID(ByVal ModuleCode As RichTextBox)
        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String

        For i As Integer = 0 To MCLength
            '��������� ����� ������ ������ � �������
            ReDim Preserve TABLE_ID(TABLE_ID.Length)
            Dim actionLine As Integer = i + 1
            TABLE_ID(TABLE_ID.Length - 1) = "������: " + actionLine.ToString


            STROKA = ModuleCode.Lines.GetValue(i)
            '����������� � ������
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If (Simvol = " " And MultilineComent = False) Then
                    '**���������� ������� *********************
                    If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                        ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                        TABLE_ID(TABLE_ID.Length - 1) = Operator
                    End If
                    Operator = ""
                Else
                    If (Simvol = ";" Or Simvol = ";" And MultilineComent = False) Then
                        '**���������� ������� *********************
                        If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                            ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                            TABLE_ID(TABLE_ID.Length - 1) = Operator
                            ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                            TABLE_ID(TABLE_ID.Length - 1) = ";"
                        Else
                            ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                            TABLE_ID(TABLE_ID.Length - 1) = ";"
                        End If
                        Operator = ""
                    Else
                        If (Simvol = "," And MultilineComent = False) Then
                            '**���������� ������� *********************
                            If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                TABLE_ID(TABLE_ID.Length - 1) = Operator
                                ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                TABLE_ID(TABLE_ID.Length - 1) = ","
                            Else
                                ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                TABLE_ID(TABLE_ID.Length - 1) = ","
                            End If
                            Operator = ""
                        Else
                            If (Simvol = "(" And MultilineComent = False) Then
                                '**���������� ������� *********************
                                If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                    TABLE_ID(TABLE_ID.Length - 1) = Operator
                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                    TABLE_ID(TABLE_ID.Length - 1) = "("
                                Else
                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                    TABLE_ID(TABLE_ID.Length - 1) = "("
                                End If
                                Operator = ""
                            Else
                                If (Simvol = ")" And MultilineComent = False) Then
                                    '**���������� ������� *********************
                                    If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                        ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                        TABLE_ID(TABLE_ID.Length - 1) = Operator
                                        ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                        TABLE_ID(TABLE_ID.Length - 1) = ")"
                                    Else
                                        ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                        TABLE_ID(TABLE_ID.Length - 1) = ")"
                                    End If
                                    Operator = ""
                                Else
                                    If (Simvol = "{" And MultilineComent = False) Then
                                        '**���������� ������� *********************
                                        If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                            ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                            TABLE_ID(TABLE_ID.Length - 1) = Operator
                                            ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                            TABLE_ID(TABLE_ID.Length - 1) = "{"
                                        Else
                                            ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                            TABLE_ID(TABLE_ID.Length - 1) = "{"
                                        End If
                                        Operator = ""
                                    Else
                                        If (Simvol = "}" And MultilineComent = False) Then
                                            '**���������� ������� *********************
                                            If (Operator <> "" And Operator <> " " And MultilineComent = False) Then
                                                ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                                TABLE_ID(TABLE_ID.Length - 1) = Operator
                                                ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                                TABLE_ID(TABLE_ID.Length - 1) = "}"
                                            Else
                                                ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                                TABLE_ID(TABLE_ID.Length - 1) = "}"
                                            End If
                                            Operator = ""
                                        Else
                                            If (iOperator = STROKA.Length - 1) Then
                                                '**���������� ������� *********************
                                                Operator = Operator + Simvol
                                                If (Operator <> "" And Operator <> " " And Operator <> "/*" And MultilineComent = False) Then
                                                    ReDim Preserve TABLE_ID(TABLE_ID.Length) '����������� ������
                                                    TABLE_ID(TABLE_ID.Length - 1) = Operator
                                                End If
                                                If (Operator = "/*") Then MultilineComent = True
                                                If (Operator = "*/") Then MultilineComent = False
                                                Operator = ""
                                            Else
                                                '�� �������� ��������� (177580) ���������� � �����
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
            '�������...
            Operator = ""
            Simvol = ""
        Next

        '�������� ������ ��� �����������
        Dim Element As String
        Dim StrokaCode As String

        For iTest As Integer = 1 To TABLE_ID.Length - 1
            Element = TABLE_ID(iTest).ToString

            If (Element.Length > 6) Then StrokaCode = Element.Chars(0) + Element.Chars(1) + Element.Chars(2) + Element.Chars(3) + Element.Chars(4) + Element.Chars(5) + Element.Chars(6)
            If (StrokaCode = "������:") Then
                ModuleCodeNoneComent.Text = ModuleCodeNoneComent.Text + System.Environment.NewLine
            Else
                If (Element = "�����" Or Element = "�����" Or Element = "���������" Or Element = "���������" Or Element = "�������" Or Element = "�������" Or Element = "{" Or Element = "}" Or Element = "(" Or Element = ")" Or Element = ";" Or Element = ",") Then
                    ModuleCodeNoneComent.Text = ModuleCodeNoneComent.Text + Element
                Else
                    ModuleCodeNoneComent.Text = ModuleCodeNoneComent.Text + " " + Element
                End If

            End If
            StrokaCode = ""
            Element = ""
        Next
    End Sub

    '�������� ������� ��������������� �����������, ���������, �����������.
    Private Sub add_Perem_Func_Proc(ByVal ModuleCode As RichTextBox)
        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim InputPARAM As Boolean = False '��������� ��������� ��� �������

        For i As Integer = 0 To MCLength
            STROKA = ModuleCode.Lines.GetValue(i)
            '���������� ���
            TYPE(STROKA)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            '����������� � ������
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If Simvol = " " Or Simvol = ";" Or Simvol = ";" Or Simvol = "," Or Simvol = "(" Or Simvol = ")" Or Simvol = "{" Or Simvol = "}" Then '����������� ������������� ������
                    '//����������//////////////////////////////////////
                    If (TYPE_peremen = True) Then
                        If Operator = "�����" Or Operator = "�����" Or Operator = "���" Or Operator = "���" Or Operator = "������" Or Operator = "������" Or Operator = "�����" Or Operator = "�����" Or Operator = "" Or Operator = " " Then
                            '��� ��������.
                        Else
                            '**�������������� ��������*****************
                            If (SyntaxPerem(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("������! ������: " + LineError.ToString + " (���������� [" + Operator + "] ��� ����������, � ��������� ���������� �����������.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**���������� ������� *********************
                            countPerem = countPerem + 1 '���� ���������� ����������
                            ReDim Preserve Perem(1, 1, countPerem) '����������� ������
                            '**��� ����������**************************
                            Perem(0, 0, countPerem) = Operator
                            '**��� �����*******************************
                            If (TYPE_chislo = True) Then
                                Perem(0, 1, countPerem) = "�����"
                                '�������� �� ���������
                                Perem(1, 1, countPerem) = "0"
                            End If
                            '**��� ������******************************
                            If (TYPE_stroka = True) Then
                                Perem(0, 1, countPerem) = "������"
                                '�������� �� ���������
                                Perem(1, 1, countPerem) = ""
                            End If
                        End If
                    End If

                    '//���������//////////////////////////////////////
                    If (TYPE_procedure = True) Then
                        If Operator = "���������" Or Operator = "���������" Or Operator = "(" Or Operator = ")" Or Operator = "{" Or Operator = "}" Or Operator = "" Or Operator = " " Then
                            '��� ��������.
                        Else
                            '**�������������� ��������*****************
                            If (SyntaxProc(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("������! ������: " + LineError.ToString + " (��������� � ������ [" + Operator + "] ��� ���������, ��������� ���������� �����������.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**���������� ������� *********************
                            countProc = countProc + 1
                            ReDim Preserve Proc(1, 1, countProc) '����������� ������
                            '**��� ���������***************************
                            Proc(0, 0, countProc) = Operator
                            Proc(0, 1, countProc) = "null"
                            '**���� ��������� *************************
                            Proc(1, 1, countProc) = BodyProc(ModuleCode, i)
                        End If
                    End If
                    If (TYPE_procedure_Param = True) Then
                        If Operator = "���������" Or Operator = "���������" Or Operator = "" Or Operator = " " Then
                            '��� ��������.
                        Else
                            '��������� ��������� ���������� � ������� ����������
                            If (Operator = "�����" Or Operator = "�����") Then
                                InputPARAM = True
                            End If
                            '**********************************************
                            If (InputPARAM = True) Then '��������� ���������
                                If Operator = "�����" Or Operator = "�����" Or Operator = "���" Or Operator = "���" Or Operator = "������" Or Operator = "������" Or Operator = "�����" Or Operator = "�����" Or Operator = "" Or Operator = " " Then
                                    '��� ��������.
                                Else
                                    '**�������������� ��������*****************
                                    If (SyntaxPerem(Operator) = True) Then
                                        Dim LineError As Integer = i
                                        SomovCompileModule.WError.ListBox1.Items.Add("������! ������: " + LineError.ToString + " (���������� [" + Operator + "] ��� ����������, � ��������� ���������� �� ��������.)")
                                        SomovCompileModule.WError.Show()
                                        Exit For
                                    End If
                                    '**���������� ������� *********************
                                    countPerem = countPerem + 1 '���� ���������� ����������
                                    ReDim Preserve Perem(1, 1, countPerem) '����������� ������
                                    '**��� ����������**************************
                                    Perem(0, 0, countPerem) = Operator
                                    '**��� �����*******************************
                                    If (TYPE_chislo = True) Then
                                        Perem(0, 1, countPerem) = "�����"
                                        '�������� �� ���������
                                        Perem(1, 1, countPerem) = "0"
                                    End If
                                    '**��� ������******************************
                                    If (TYPE_stroka = True) Then
                                        Perem(0, 1, countPerem) = "������"
                                        '�������� �� ���������
                                        Perem(1, 1, countPerem) = ""
                                    End If
                                End If
                            Else    '��� ��������� � Ũ ����
                                '**�������������� ��������*****************
                                If (SyntaxProc(Operator) = True) Then
                                    Dim LineError As Integer = i
                                    SomovCompileModule.WError.ListBox1.Items.Add("������! ������: " + LineError.ToString + " (��������� � ������ [" + Operator + "] ��� ���������, ��������� ���������� �����������.)")
                                    SomovCompileModule.WError.Show()
                                    Exit For
                                End If
                                '**���������� ������� *********************
                                countProc = countProc + 1
                                ReDim Preserve Proc(1, 1, countProc) '����������� ������
                                '**��� ���������***************************
                                Proc(0, 0, countProc) = Operator
                                Proc(0, 1, countProc) = "null"
                                '**���� ��������� *************************
                                Proc(1, 1, countProc) = BodyProc(ModuleCode, i)
                            End If
                        End If
                    End If

                    '//�������//////////////////////////////////////
                    If (TYPE_function = True) Then
                        If Operator = "�������" Or Operator = "�������" Or Operator = "(" Or Operator = ")" Or Operator = "{" Or Operator = "}" Or Operator = "" Or Operator = " " Then
                            '��� ��������.
                        Else
                            '**�������������� ��������*****************
                            If (SyntaxFunc(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("������! ������: " + LineError.ToString + " (������� � ������ [" + Operator + "] ��� ���������, ��������� ���������� �����������.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**���������� ������� *********************
                            countFunc = countFunc + 1
                            ReDim Preserve Func(1, 1, countFunc) '����������� ������
                            '**��� ������� ****************************
                            Func(0, 0, countFunc) = Operator
                            '**��� ������� ****************************
                            Func(0, 1, countFunc) = TypeFunc(ModuleCode, i)
                            '**���� ������� ***************************
                            Func(1, 1, countFunc) = BodyFunc(ModuleCode, i)
                        End If
                    End If
                    If (TYPE_function_Param = True) Then
                        If Operator = "�������" Or Operator = "�������" Or Operator = "" Or Operator = " " Then
                            '��� ��������.
                        Else
                            '��������� ������� ���������� � ������� ����������
                            If (Operator = "�����" Or Operator = "�����") Then
                                InputPARAM = True
                            End If
                            '**********************************************
                            If (InputPARAM = True) Then '��������� �������
                                If Operator = "�����" Or Operator = "�����" Or Operator = "���" Or Operator = "���" Or Operator = "������" Or Operator = "������" Or Operator = "�����" Or Operator = "�����" Or Operator = "" Or Operator = " " Then
                                    '��� ��������.
                                Else
                                    '**�������������� ��������*****************
                                    If (SyntaxPerem(Operator) = True) Then
                                        Dim LineError As Integer = i
                                        SomovCompileModule.WError.ListBox1.Items.Add("������! ������: " + LineError.ToString + " (���������� [" + Operator + "] ��� ����������, � ��������� ���������� �� ��������.)")
                                        SomovCompileModule.WError.Show()
                                        Exit For
                                    End If
                                    '**���������� ������� *********************
                                    countPerem = countPerem + 1 '���� ���������� ����������
                                    ReDim Preserve Perem(1, 1, countPerem) '����������� ������
                                    '**��� ����������**************************
                                    Perem(0, 0, countPerem) = Operator
                                    '**��� �����*******************************
                                    If (TYPE_chislo = True) Then
                                        Perem(0, 1, countPerem) = "�����"
                                        '�������� �� ���������
                                        Perem(1, 1, countPerem) = "0"
                                    End If
                                    '**��� ������******************************
                                    If (TYPE_stroka = True) Then
                                        Perem(0, 1, countPerem) = "������"
                                        '�������� �� ���������
                                        Perem(1, 1, countPerem) = ""
                                    End If
                                    If (Operator = "���������") Then IndexStatusPerem = countPerem
                                End If
                            Else    '��� ������� � Ũ ����
                            '**�������������� ��������*****************
                            If (SyntaxFunc(Operator) = True) Then
                                Dim LineError As Integer = i
                                SomovCompileModule.WError.ListBox1.Items.Add("������! ������: " + LineError.ToString + " (������� � ������ [" + Operator + "] ��� ���������, ��������� ���������� �����������.)")
                                SomovCompileModule.WError.Show()
                                Exit For
                            End If
                            '**���������� ������� *********************
                            countFunc = countFunc + 1
                            ReDim Preserve Func(1, 1, countFunc) '����������� ������
                            '**��� ���������***************************
                            Func(0, 0, countFunc) = Operator
                            '**��� ������� ****************************
                            Func(0, 1, countFunc) = TypeFunc(ModuleCode, i)
                            '**���� ��������� *************************
                            Func(1, 1, countFunc) = BodyFunc(ModuleCode, i)
                            End If
                        End If
                    End If
                    '...

                    Operator = ""
                Else
                    '�� �������� ��������� (177580) ���������� � �����
                    If (Simvol.GetHashCode <> 177580 And Simvol.GetHashCode <> -842352729) Then Operator = Operator + Simvol
                End If
            Next
            '�������...
            ClearTYPE()
            Operator = ""
            Simvol = ""
            InputPARAM = False
        Next
    End Sub

    '���� ���������
    Private Function BodyProc(ByVal ModuleCode As RichTextBox, ByVal StartProc As Integer) As String
        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim Open_Close As Integer '������� �������� � �������� �������� ������ { }
        Dim RTB As RichTextBox = New RichTextBox

        StartProc = StartProc + 1
        For i As Integer = StartProc To MCLength
            STROKA = ModuleCode.Lines.GetValue(i)

            '��������� ������ ������������ ����
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If Simvol = "{" Then '
                    Open_Close = Open_Close + 1 '�������� ������ {
                End If
                If Simvol = "}" Then '
                    Open_Close = Open_Close - 1  '�������� ������ }
                End If
                Simvol = ""
            Next

            If (i <> StartProc) Then '���������� ��������� ���� ���������
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

    '��� � ���� �������
    Private Function TypeFunc(ByVal ModuleCode As RichTextBox, ByVal StartProc As Integer) As String
        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim ThisTYPE As Boolean = False

        STROKA = ModuleCode.Lines.GetValue(StartProc)

        '��������� ������ ������������ ����
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
        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        Dim MCLength As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String
        Dim Open_Close As Integer '������� �������� � �������� �������� ������ { }
        Dim RTB As RichTextBox = New RichTextBox

        StartProc = StartProc + 1
        For i As Integer = StartProc To MCLength
            STROKA = ModuleCode.Lines.GetValue(i)

            '��������� ������ ������������ ����
            For iOperator As Integer = 0 To STROKA.Length - 1
                Simvol = STROKA.Chars(iOperator)
                If Simvol = "{" Then '
                    Open_Close = Open_Close + 1 '�������� ������ {
                End If
                If Simvol = "}" Then '
                    Open_Close = Open_Close - 1  '�������� ������ }
                End If
                Simvol = ""
            Next

            If (i <> StartProc) Then '���������� ��������� ���� ���������
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



    '//����������� ������////////////////////////////////////////////////////////
    '���������� ����� ��������� ����� ������ ��������.
    Private TYPE_peremen As Boolean
    Private TYPE_chislo As Boolean '��� �����
    Private TYPE_stroka As Boolean '��� ������
    Private TYPE_object As Boolean '��� ������
    Private TYPE_prisvoit As Boolean '���������
    Private TYPE_message As Boolean '��������
    Private TYPE_zapros As Boolean '������
    Private TYPE_question As Boolean '������
    Private TYPE_procedure As Boolean '���������
    Private TYPE_procedure_Param As Boolean '��������� � �����������
    Private TYPE_procedure_RUN As Boolean '��������� ���������
    Private TYPE_function As Boolean '�������
    Private TYPE_function_Param As Boolean '������� � �����������
    Private TYPE_matematika As Boolean

    '����������� ���� ����������� ���������
    Private Sub TYPE(ByVal Stroka As String) '���
        ClearTYPE() '������� ���������� ����������� �����
        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        For iOperator As Integer = 0 To Stroka.Length - 1
            Simvol = Stroka.Chars(iOperator)
            If Simvol = " " Or Simvol = ";" Or Simvol = ";" Or Simvol = "(" Or Simvol = ")" Then '���� ������ ����� � ������� ������ �������� ��������

                '���������� �������� ���������.
                If Operator = "�����" Or Operator = "�����" Then TYPE_peremen = True
                If Operator = "�����" Or Operator = "�����" Then TYPE_chislo = True
                If Operator = "������" Or Operator = "������" Then TYPE_stroka = True
                If Operator = "�����" Or Operator = "�����" Then TYPE_object = True
                If Operator = "=" Then TYPE_prisvoit = True
                If Operator = "+" Or Operator = "-" Or Operator = "*" Or Operator = "/" Then
                    TYPE_matematika = True  '�������������� ����������
                    TYPE_peremen = False
                End If
                '���������� �������
                If Operator = "��������" Or Operator = "��������" Then TYPE_message = True
                If Operator = "������" Or Operator = "������" Then TYPE_question = True
                If Operator = "������" Or Operator = "������" Then TYPE_zapros = True
                '���������� ��� ��������� ********************************
                If (Operator = "���������" Or Operator = "���������") Then
                    TYPE_procedure = True   '��������� ��� ����������
                End If
                If (TYPE_procedure = True And TYPE_peremen = True) Then
                    TYPE_procedure_Param = True '��������� � ����������
                    TYPE_procedure = False
                    TYPE_peremen = False
                End If
                For i As Integer = 1 To CORE.countProc
                    If (Proc(0, 0, i).ToString = Operator) Then
                        TYPE_procedure_RUN = True
                        Exit Sub
                    End If
                Next

                '���������� ��� ������� **********************************
                If (Operator = "�������" Or Operator = "�������") Then
                    TYPE_function = True    '������� ��� ����������
                End If
                If (TYPE_function = True And TYPE_peremen = True) Then
                    TYPE_function_Param = True  '������� � �����������
                    TYPE_function = False
                    TYPE_peremen = False
                End If

                '-----------------------------
                Operator = ""
                Simvol = ""
            Else
                '�� �������� ��������� (WinXP 177580, WinSeven -842352729) ���������� � �����
                If (Simvol.GetHashCode <> 177580 And Simvol.GetHashCode <> -842352729 And Simvol <> " ") Then Operator = Operator + Simvol
            End If
        Next
    End Sub



    '//�������������� � ������������� ������///////////////////////////////////
    '�������� ��� ����������
    Private Function SyntaxPerem(ByVal Test As String) As Boolean
        '�������� ���������� ��� ����������*****
        For i As Integer = 1 To CORE.countPerem
            If (Perem(0, 0, i).ToString = Test) Then
                Return True
                Exit Function
            End If
        Next
        '****************************************
    End Function

    '�������� ���� ��������
    Private Function SyntaxProc(ByVal Test As String) As Boolean
        '�������� ���������� ��� ����������*****
        For i As Integer = 1 To CORE.countProc
            If (Proc(0, 0, i).ToString = Test) Then
                Return True
                Exit Function
            End If
        Next
        '****************************************
    End Function

    '�������� ���� �������
    Private Function SyntaxFunc(ByVal Test As String) As Boolean
        '�������� ���������� ��� ����������*****
        For i As Integer = 1 To CORE.countFunc
            If (Func(0, 0, i).ToString = Test) Then
                Return True
                Exit Function
            End If
        Next
        '****************************************
    End Function

    '������ �������� ����� ������
    Private Function SyntaxAll(ByVal Test As RichTextBox) As Boolean

        If (SomovCompileModule.WErrorShow = True) Then
            Return True
            Exit Function
        End If

        '�������� ������ ��� �����������
        Dim Element As String
        Dim StrokaCode As String
        Dim ActionLine As String
        Dim StringOPEN As Boolean = False   '������ �������/�������
        Dim ListPeremFor As New ListBox         '�������������� ���� ����������� ���������� � �����

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

        '�������� �� ������� "�������" �������
        Dim StartFunc As Boolean = False
        For i As Integer = 1 To CORE.countFunc
            If (Func(0, 0, i).ToString = "�������" Or Func(0, 0, i).ToString = "�������") Then
                StartFunc = True
                Exit For
            End If
        Next
        If (StartFunc = False) Then
            SomovCompileModule.WError.ListBox1.Items.Add("������! � ������ �� ��������� ������� �������.(������� �������(����� ��������� ��� ������) �����)")
            SomovCompileModule.WError.Show()
            Return True
            Exit Function
        End If
        '************************************


        For iTest As Integer = 1 To TABLE_ID.Length - 1
            Element = TABLE_ID(iTest).ToString

            If (Element.Length > 6) Then StrokaCode = Element.Chars(0) + Element.Chars(1) + Element.Chars(2) + Element.Chars(3) + Element.Chars(4) + Element.Chars(5) + Element.Chars(6)
            If (StrokaCode = "������:") Then
                ActionLine = Element
                '��������� ������ ��������� ������ ;
                If (iTest > 1) Then
                    If (TABLE_ID(iTest - 1).ToString = ";" Or TABLE_ID(iTest - 1).ToString = ")" Or TABLE_ID(iTest - 1).ToString = "�����" Or TABLE_ID(iTest - 1).ToString = "�����" Or TABLE_ID(iTest - 1).ToString = "������" Or TABLE_ID(iTest - 1).ToString = "������" Or TABLE_ID(iTest - 1).ToString = "�����" Or TABLE_ID(iTest - 1).ToString = "�����" Or TABLE_ID(iTest - 1).ToString = "{" Or TABLE_ID(iTest - 1).ToString = "}") Then
                        '������ ����� ������ ������ - �Ѩ �����!!!
                    Else
                        '����� ��������� �� ������ ---------------
                        Dim FrontElementTest As String = TABLE_ID(iTest - 1).ToString
                        If (FrontElementTest.Length > 6) Then FrontElementTest = FrontElementTest.Chars(0) + FrontElementTest.Chars(1) + FrontElementTest.Chars(2) + FrontElementTest.Chars(3) + FrontElementTest.Chars(4) + FrontElementTest.Chars(5) + FrontElementTest.Chars(6)
                        If (FrontElementTest <> "������:") Then
                            SomovCompileModule.WError.ListBox1.Items.Add("������! " + ActionLine + " (�� ������ ������ ����� ������ ; ����� � �������)")
                            SomovCompileModule.WError.Show()
                        End If
                    End If
                End If
            Else
                '�������� ����� ������ �� ��������� ������
                '1.�������� �������� � ������ ����������������� ����������
                Dim TextN1 As Boolean = False
                For iTestOperator As Integer = 0 To LISTOperators.Items.Count - 1
                    If (Element.Chars(0) = """" Or Element.Chars(0) = "'" Or Element.Chars(Element.Length - 1) = """" Or Element.Chars(Element.Length - 1) = "'") Then
                        If (Element.Chars(0) = """" And Element.Chars(Element.Length - 1) = """") Then
                            TextN1 = True '�������� �������� ��� ������ �� ������!
                            StringOPEN = False '�������� ��� ������ �������
                            Exit For
                        End If
                        If (Element.Chars(0) = "'" And Element.Chars(Element.Length - 1) = "'") Then
                            TextN1 = True '�������� �������� ��� ������ �� ������!
                            StringOPEN = False '�������� ��� ������ �������
                            Exit For
                        End If
                        If (StringOPEN = False) Then
                            StringOPEN = True   '�������� ��� ������ �������
                        Else
                            StringOPEN = False '�������� ��� ������ �������
                        End If
                        TextN1 = True '�������� �������� ��� ������ �� ������!
                        Exit For
                    End If
                    If (StringOPEN = True) Then
                        TextN1 = True '�������� �������� ��� ������ �� ������!
                        Exit For
                    End If

                    If (Element.Chars(0) = "0" Or Element.Chars(0) = "1" Or Element.Chars(0) = "2" Or Element.Chars(0) = "3" Or Element.Chars(0) = "4" Or Element.Chars(0) = "5" Or Element.Chars(0) = "6" Or Element.Chars(0) = "7" Or Element.Chars(0) = "8" Or Element.Chars(0) = "9") Then
                        TextN1 = True '�������� �������� ��� ����� �� ������!
                        Exit For
                    End If
                    If (Element = LISTOperators.Items.Item(iTestOperator)) Then
                        TextN1 = True '�������� �������� ��� ����������������� �� ������!
                        Exit For
                    End If
                Next
                '2.�������� ����� ���������� �����
                For iLPFor As Integer = 0 To ListPeremFor.Items.Count - 1
                    If (Element = ListPeremFor.Items.Item(iLPFor)) Then
                        TextN1 = True '�������� �������� ��� �������������� ���������� ����� �� ������!
                        Exit For
                    End If
                Next
                '3.�������� �������� � ������� ����������� ����������, �������� � �������
                If (TextN1 = False And StringOPEN = False) Then
                    '�������� ��� ����������
                    If (SyntaxPerem(Element) = False) Then
                        '�������� ��� ��������
                        If (SyntaxProc(Element) = False) Then
                            '�������� ��� �������
                            If (SyntaxFunc(Element) = False) Then
                                If (iTest - 2 > 0) Then
                                    If (TABLE_ID(iTest - 2).ToString = "��� " Or TABLE_ID(iTest - 2).ToString = "���") Then
                                        ListPeremFor.Items.Add(Element)
                                    Else
                                        '����� ��������� �� ������ ---------------
                                        SomovCompileModule.WError.ListBox1.Items.Add("������! " + ActionLine + " (" + Element + " - ������ �� ��������)")
                                        SomovCompileModule.WError.Show()
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
                '4 ��������� ������������ ��������� �������� ������ { }
                If (Element = "{" Or Element = "}") Then
                    Dim front As String '�������
                    Dim back As String '�����
                    front = TABLE_ID(iTest - 1).ToString
                    If (iTest < TABLE_ID.Length - 1) Then
                        back = TABLE_ID(iTest + 1).ToString
                    Else
                        back = "������:"
                    End If

                    If (front.Length > 6) Then front = front.Chars(0) + front.Chars(1) + front.Chars(2) + front.Chars(3) + front.Chars(4) + front.Chars(5) + front.Chars(6)
                    If (back.Length > 6) Then back = back.Chars(0) + back.Chars(1) + back.Chars(2) + back.Chars(3) + back.Chars(4) + back.Chars(5) + back.Chars(6)

                    If (front = "������:" And back = "������:") Then
                        '������ ������� � ����� ������ - �Ѩ �����!!!
                    Else
                        '����� ��������� �� ������ ---------------
                        SomovCompileModule.WError.ListBox1.Items.Add("������! " + ActionLine + " (������ " + Element + " - ������ ���� ������� � ����� ������ � �������� �� ��������� ����������.)")
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



    '//����������////////////////////////////////////////////////////////////////
    Public Sub RUN(ByVal Code As RichTextBox, ByVal Operators As ListBox)
        '������� ������ (���������� ������������ ����)
        CORE.ClearModule()              '��������������� ������� ������� ���������������
        '****************************************
        '���������� ���� ��������� �� �������
        If (WErrorShow = False) Then
            SomovCompileModule.LoadWError()
            WError.ListBox1.Items.Clear()
        Else
            WError.ListBox1.Items.Clear()
        End If
        '****************************************
        CORE.LISTOperators = Operators
        CORE.AddTableID(Code)
        CORE.add_Perem_Func_Proc(ModuleCodeNoneComent)   '��������� ������� ���������������
        '****************************************
        If (CORE.SyntaxAll(ModuleCodeNoneComent) = True) Then Exit Sub '������ �������� ������
        '****************************************
        RUN_MAIN()   '���������������� ���������� ���� ������� � "�������" �������
    End Sub
    '������ ���������� ������������ ����
    Private Sub RUN_MAIN()
        Dim MainFunction As New RichTextBox
        '������ ���������� � "�������" �������
        For i As Integer = 1 To CORE.countFunc
            If (Func(0, 0, i).ToString = "�������" Or Func(0, 0, i).ToString = "�������") Then
                MainFunction.Text = Func(1, 1, i)
                Exit For
            End If
        Next
        '************************************
        Perem(1, 1, IndexStatusPerem) = "��������� ������� ������� ������!"
        ExecuteCode(MainFunction) '���������� �� ���������� "�������" �������
    End Sub
    '���������� ������������ ����
    Private Sub ExecuteCode(ByVal ModuleCode As RichTextBox)
        Perem(1, 1, IndexStatusPerem) = "���������� ������������ ����!"

        '��������� ������� � ������
        Dim LevelOpenClose As Integer = 0       '����������� ������ � ������ ��������� �������
        Dim IfAndCycle(1, 1, 1, 100) As String  '���������: --����--����--�����--�������--
        Dim CycleCode As New RichTextBox
        '--------------------------

        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        Dim iSTROKA As Integer = ModuleCode.Lines.Length - 1
        Dim STROKA As String

        For StrokaRUN As Integer = 0 To iSTROKA
            STROKA = ModuleCode.Lines.GetValue(StrokaRUN)

            '//��� ������//////////////////////////////////////
            Dim Test As String = STROKA
            If (Test.Length > 6) Then '��� "����" � "���("
                Test = Test.Remove(0, 1)
                Test = Test.Remove(4, Test.Length - 4)
            Else
                If (Test.Length > 5) Then '��� "�����"
                    Test = Test.Remove(0, 1)
                    Test = Test.Remove(5, Test.Length - 5)
                Else
                    If (Test.Length > 4) Then '��� "���"
                        Test = Test.Remove(0, 1)
                        Test = Test.Remove(3, Test.Length - 3)
                    Else
                        If (Test.Length = 1) Then '��� "{" � "}"
                            If (Test = "{") Then    '�������� ������
                                LevelOpenClose = LevelOpenClose + 1
                                IfAndCycle(1, 1, 1, LevelOpenClose) = LevelOpenClose.ToString   '�������
                            End If
                            If (Test = "}") Then    '�������� ������
                                LevelOpenClose = LevelOpenClose - 1
                                IfAndCycle(1, 1, 1, LevelOpenClose) = LevelOpenClose.ToString   '�������
                            End If
                        End If
                    End If
                End If
            End If

            '//�������////////////////////////////////////////
            If (Test = "����" Or Test = "����") Then
                Perem(1, 1, IndexStatusPerem) = "���������� �������!"

                If (VerificationIF(STROKA) = True) Then
                    IfAndCycle(0, 0, 0, LevelOpenClose + 1) = "NULL"    '�� ����
                    IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "��"      '������� = ������
                    IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "���"     '����� = �����
                    IfAndCycle(1, 1, 1, LevelOpenClose + 1) = ""        '�������
                Else
                    IfAndCycle(0, 0, 0, LevelOpenClose + 1) = "NULL"    '�� ����
                    IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "���"     '������� = �����
                    IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "��"      '����� = ������
                    IfAndCycle(1, 1, 1, LevelOpenClose + 1) = ""        '�������
                End If
            End If
            If (Test = "�����" Or Test = "�����") Then  '������ ������
                If (IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "��") Then '���� ������� = ������
                    IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "���" '������� = �����
                    IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "��"  '����� = ������
                Else
                    If (IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "���") Then '���� ������� = �����
                        IfAndCycle(1, 0, 0, LevelOpenClose + 1) = "��" '������� = ������
                        IfAndCycle(1, 1, 0, LevelOpenClose + 1) = "���"  '����� = �����
                    End If
                End If

            End If

            '//�����//////////////////////////////////////////
            If (Test = "���(" Or Test = "���(") Then
                Perem(1, 1, IndexStatusPerem) = "���������� �����!"

                '�������� ���������� ������� �������
                Dim All_TRUE As Boolean = True '��� ���������� ��/���
                If (LevelOpenClose > 0) Then '�������� ���������� �������
                    For iTestALL As Integer = 1 To LevelOpenClose
                        If (IfAndCycle(1, 0, 0, iTestALL) = "���") Then
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


            '//����������////////////////////////////////////
            If (LevelOpenClose = 0) Then    '��� ������� � ������

                '���������� ���� ���������� � ������ ������
                TYPE(STROKA)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                '//���������//�������/////////////////////////////
                If (TYPE_prisvoit = True) Then
                    Perem(1, 1, IndexStatusPerem) = "��������� ��������!"
                    SetVariable(STROKA)
                End If

                '//��������//�������//////////////////////////////
                If (TYPE_message = True) Then
                    ShowMessage(STROKA)
                End If

                '//���������//////////////////////////////////////
                If (TYPE_procedure_RUN = True) Then
                    Perem(1, 1, IndexStatusPerem) = "��������� ���������!"
                    ExecutePROC(STROKA)
                End If

            Else    '����������� �������
                If (IfAndCycle(1, 0, 0, LevelOpenClose) = "��") Then

                    Dim All_TRUE As Boolean = True '��� ���������� ��/���
                    If (LevelOpenClose > 0) Then '�������� ���������� �������
                        For iTestALL As Integer = 1 To LevelOpenClose
                            If (IfAndCycle(1, 0, 0, iTestALL) = "���") Then
                                All_TRUE = False
                                Exit For
                            End If
                        Next
                    End If


                    If (All_TRUE = True) Then
                        '���������� ���� ���������� � ������ ������
                        TYPE(STROKA)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                        '//���������//�������/////////////////////////////
                        If (TYPE_prisvoit = True) Then
                            Perem(1, 1, IndexStatusPerem) = "��������� ��������!"
                            SetVariable(STROKA)
                        End If

                        '//��������//�������//////////////////////////////
                        If (TYPE_message = True) Then
                            ShowMessage(STROKA)
                        End If

                        '//���������//////////////////////////////////////
                        If (TYPE_procedure_RUN = True) Then
                            Perem(1, 1, IndexStatusPerem) = "��������� ���������!"
                            ExecutePROC(STROKA)
                        End If
                    End If
                End If

            End If
        Next
    End Sub
    '--��������� �������� ����������, ���������� ����������
    Private Sub SetVariable(ByVal CodeLine As String)
        Perem(1, 1, IndexStatusPerem) = "��������� ��������!"

        Dim Simvol As String '������ �� ������
        Dim Operator As String '�������� �������� ��������
        Dim ActionPerem As String

        For iOperator As Integer = 0 To CodeLine.Length - 1
            Simvol = CodeLine.Chars(iOperator)
            If Simvol = " " Or Simvol = ";" Or Simvol = ";" Then '���� ������ ������ ��� ����� � ������� ������ �������� ��������

                If Operator = "=" Then
                    '���������� ��� ���������� � �������� �� ���������
                    Dim ValueLine As String
                    ValueLine = CodeLine.Remove(0, iOperator)
                    If (AddValue(ActionPerem, ValueLine) = False) Then
                        Exit Sub
                    Else
                        '��������� ������
                        SomovCompileModule.WError.ListBox1.Items.Add("������! � ������ [" + CodeLine + "] (�� ������� ��������� ����������)")
                        SomovCompileModule.WError.Show()
                        Exit Sub
                    End If
                Else
                    If (Operator <> " " And Operator <> "") Then ActionPerem = Operator '��� ���������� ������� ��������� �������� ��� ���������
                End If

                '-----------------------------
                Operator = ""
                Simvol = ""
            Else
                '�� �������� ��������� (177580) ���������� � �����
                If (Simvol.GetHashCode <> 177580 And Simvol.GetHashCode <> -842352729 And Simvol <> " ") Then Operator = Operator + Simvol
            End If
        Next
    End Sub
    Private Function AddValue(ByVal NamePerem As String, ByVal CodeLine As String) As Boolean
        Perem(1, 1, IndexStatusPerem) = "��������� ��������!"
        Try
            Dim TestTYPE As String
            Dim Token As String
            Dim Zapros As String
            Dim Result As String

            If (CodeLine.Chars(CodeLine.Length - 1) <> ";") Then
                '��������� ������ (�������������)
                SomovCompileModule.WError.ListBox1.Items.Add("��������������! � ������ [" + CodeLine + "] (����������� ������� ����� ������ ; ����� � �������)")
                SomovCompileModule.WError.Show()
            End If

            '��������� ������� � ���������� ������ ***********************************
            For iTest As Integer = 0 To CodeLine.Length - 1
                '���� ������ � ������������ ------------------------------------------
                If (Token = "������(" Or Token = "������(") Then '��������� ��� ������
                    TestTYPE = "������" '!!! ��� �������� ��� ���������
                    Dim kv As Integer = iTest + 1
                    If (CodeLine.Chars(kv) = """" Or CodeLine.Chars(kv) = "'") Then  '�������� ������� ����� ������
                        Zapros = CodeLine.Remove(0, kv + 1)
                        Zapros = Zapros.Remove(Zapros.Length - 3, 3)
                        Result = InputBox(Zapros, "������:")
                        Exit For
                    Else    '�������� ������� � ���� ����������
                        Zapros = CodeLine.Remove(0, kv)
                        Zapros = Zapros.Remove(Zapros.Length - 2, 2)
                        For iPerem As Integer = 1 To CORE.countPerem
                            If (Perem(0, 0, iPerem).ToString = Zapros) Then
                                Result = InputBox(Perem(1, 1, iPerem), "������:")
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
                '������� ������������ ������ -----------------------------------------
                If (CodeLine.Chars(1) = """" Or CodeLine.Chars(1) = "'") Then
                    TestTYPE = "������" '!!! ��� �������� ��� ���������
                    '���������� ����� ��������� ������
                    If (CodeLine.Chars(CodeLine.Length - 2) = """" Or CodeLine.Chars(CodeLine.Length - 2) = "'") Then
                        Result = CodeLine.Remove(0, 2)
                        Result = Result.Remove(Result.Length - 2, 2)
                        Exit For
                    End If
                End If
            Next
            '���� �������� ��� ����������
            If (Token <> "") Then   '�������� �������� �� ��������� ����������
                For i As Integer = 1 To CORE.countPerem
                    If (Perem(0, 0, i).ToString = Token) Then
                        Result = Perem(1, 1, i) '�������� �� ����������
                        TestTYPE = Perem(0, 1, i) '!!! ��� �������� ��� ���������
                        Token = ""
                        Exit For
                    End If
                Next
            End If
            '*************************************************************************

            '��������� �����, ������� ��� ��������������� ��������� ******************
            If (TestTYPE <> "������" And TestTYPE <> "������") Then
                Result = MATEMATIKA(CodeLine)
                If (Result = "������!������!������!") Then    '�������� �� ������
                    Return True '��������� ������
                    Exit Function
                End If
                TestTYPE = "��������������" '!!! ��� �������� ��� ���������
            End If

            '���������� �������� � ���������� ***************************************
            For i As Integer = 1 To CORE.countPerem
                If (Perem(0, 0, i).ToString = NamePerem) Then
                    '�������� ������������ ���� �������� ���� ����������
                    If (Perem(0, 1, i) = TestTYPE Or TestTYPE = "������" Or TestTYPE = "��������������") Then
                        Perem(1, 1, i) = Result
                        Return False    '������������ ����������
                        Exit Function
                    Else
                        Return True '��������� ������
                        Exit Function
                    End If
                    Exit For
                End If
            Next
            '**********************************************************************
        Catch ex As Exception
            Return True '��������� ������
            Exit Function
        End Try
    End Function
    Private Function MATEMATIKA(ByVal DataLine As String) As String
        Perem(1, 1, IndexStatusPerem) = "�������������� ����������!"

        '���� �������� ���������� � �����
        Dim Token As String
        Dim FuncName, FuncParam As String
        Dim Elements(0) As String
        Dim DataL As String = DataLine.Remove(0, 1)
        DataL = DataL.Remove(DataL.Length - 1, 1)

        '��������� ���������� ��� ��������������
        If (TYPE_matematika = True) Then
            For i As Integer = 0 To DataL.Length - 1
                If (DataL.Chars(i) = " ") Then
                    ReDim Preserve Elements(Elements.Length)
                    Elements(Elements.Length - 1) = Token
                    Token = ""
                Else
                    If (DataL.Chars(i) <> " " And DataL.Chars(i) <> "") Then Token = Token + DataL.Chars(i)
                End If
                If (DataL.Length - 1 = i) Then  '����� ������
                    ReDim Preserve Elements(Elements.Length)
                    Elements(Elements.Length - 1) = Token
                    Token = ""
                End If
            Next
        Else    '��������� �� ��������������
            '���� �������� ���������� � �����
            If (DataL.Chars(0) = "0" Or DataL.Chars(0) = "1" Or DataL.Chars(0) = "2" Or DataL.Chars(0) = "3" Or DataL.Chars(0) = "4" Or DataL.Chars(0) = "5" Or DataL.Chars(0) = "6" Or DataL.Chars(0) = "7" Or DataL.Chars(0) = "8" Or DataL.Chars(0) = "9" Or DataL.Chars(0) = "0") Then
                Return DataL
                Exit Function

            Else    '���� �������� ���������� � �������

                '�������� ����������--------------------
                For i As Integer = 1 To CORE.countPerem
                    If (Perem(0, 0, i).ToString = DataL) Then
                        '�������� ������������ ���� �������� ���� ����������
                        If (Perem(0, 1, i) = "�����") Then
                            Return Perem(1, 1, i)
                            Exit Function
                        Else
                            '��������� ������
                            Return "������!������!������!"
                            Exit Function
                        End If
                        Exit For
                    End If
                Next
                '�������� �������------------------------------------------------------
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

        '��������� �������������� ���ר��************************
        If (TYPE_matematika = True) Then
            '������������� ����������

            Dim SUMMA As Double
            Dim Element1, Element2 As Double

            '������������ ���������� ****************************************************
            For iMat As Integer = 1 To Elements.Length - 1
                If (Elements(iMat) = "*" Or Elements(iMat) = "/") Then '��������� ��� �������
                    '--����� ����� �� ����� -------------------------------------------------
                    '���� ���������� � ����� - ������ ��� ����� � �� ����������
                    If (Elements(iMat - 1).Chars(0) = "0" Or Elements(iMat - 1).Chars(0) = "1" Or Elements(iMat - 1).Chars(0) = "2" Or Elements(iMat - 1).Chars(0) = "3" Or Elements(iMat - 1).Chars(0) = "4" Or Elements(iMat - 1).Chars(0) = "5" Or Elements(iMat - 1).Chars(0) = "6" Or Elements(iMat - 1).Chars(0) = "7" Or Elements(iMat - 1).Chars(0) = "8" Or Elements(iMat - 1).Chars(0) = "9" Or Elements(iMat - 1).Chars(0) = "0") Then
                        '���� ����� �������
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
                    Else    '��� ������ � ����������
                        '�������� ����������--------------------
                        For i As Integer = 1 To CORE.countPerem
                            If (Perem(0, 0, i).ToString = Elements(iMat - 1)) Then
                                '�������� ������������ ���� �������� ���� ����������
                                If (Perem(0, 1, i) = "�����") Then
                                    '���� ����� �������
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
                    '--����� ������ �� ����� ---------------------------------------------
                    '���� ���������� � ����� - ������ ��� ����� � �� ����������
                    If (Elements(iMat + 1).Chars(0) = "0" Or Elements(iMat + 1).Chars(0) = "1" Or Elements(iMat + 1).Chars(0) = "2" Or Elements(iMat + 1).Chars(0) = "3" Or Elements(iMat + 1).Chars(0) = "4" Or Elements(iMat + 1).Chars(0) = "5" Or Elements(iMat + 1).Chars(0) = "6" Or Elements(iMat + 1).Chars(0) = "7" Or Elements(iMat + 1).Chars(0) = "8" Or Elements(iMat + 1).Chars(0) = "9" Or Elements(iMat + 1).Chars(0) = "0") Then
                        '���� ����� �������
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
                    Else    '��� ������ � ����������
                        '�������� ����������--------------------
                        For i As Integer = 1 To CORE.countPerem
                            If (Perem(0, 0, i).ToString = Elements(iMat + 1)) Then
                                '�������� ������������ ���� �������� ���� ����������
                                If (Perem(0, 1, i) = "�����") Then
                                    '���� ����� �������
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
                    '��������� ����������
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

            '������������� ����������
            '�������������� ��������
            Dim TYPE_plus As Boolean = False
            Dim TYPE_minus As Boolean = False
            Dim TYPE_result As Boolean = False
            For iMat As Integer = 1 To Elements.Length - 1
                If (Elements(iMat) = "+") Then TYPE_plus = True
                If (Elements(iMat) = "-") Then TYPE_minus = True
                If (Elements(iMat) <> "NULL" And Elements(iMat) <> "+" And Elements(iMat) <> "-") Then
                    If (TYPE_plus = False And TYPE_minus = False) Then '����� ����� �� �����
                        '���� ���������� � ����� - ������ ��� ����� � �� ����������
                        If (Elements(iMat).Chars(0) = "0" Or Elements(iMat).Chars(0) = "1" Or Elements(iMat).Chars(0) = "2" Or Elements(iMat).Chars(0) = "3" Or Elements(iMat).Chars(0) = "4" Or Elements(iMat).Chars(0) = "5" Or Elements(iMat).Chars(0) = "6" Or Elements(iMat).Chars(0) = "7" Or Elements(iMat).Chars(0) = "8" Or Elements(iMat).Chars(0) = "9" Or Elements(iMat).Chars(0) = "0") Then
                            '���� ����� �������
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
                        Else    '��� ������ � ����������
                            '�������� ����������--------------------
                            For i As Integer = 1 To CORE.countPerem
                                If (Perem(0, 0, i).ToString = Elements(iMat)) Then
                                    '�������� ������������ ���� �������� ���� ����������
                                    If (Perem(0, 1, i) = "�����") Then
                                        '���� ����� �������
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
                    Else '����� ������ �� �����
                        TYPE_result = True
                        '���� ���������� � ����� - ������ ��� ����� � �� ����������
                        If (Elements(iMat).Chars(0) = "0" Or Elements(iMat).Chars(0) = "1" Or Elements(iMat).Chars(0) = "2" Or Elements(iMat).Chars(0) = "3" Or Elements(iMat).Chars(0) = "4" Or Elements(iMat).Chars(0) = "5" Or Elements(iMat).Chars(0) = "6" Or Elements(iMat).Chars(0) = "7" Or Elements(iMat).Chars(0) = "8" Or Elements(iMat).Chars(0) = "9" Or Elements(iMat).Chars(0) = "0") Then
                            '���� ����� �������
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
                        Else    '��� ������ � ����������
                            '�������� ����������--------------------
                            For i As Integer = 1 To CORE.countPerem
                                If (Perem(0, 0, i).ToString = Elements(iMat)) Then
                                    '�������� ������������ ���� �������� ���� ����������
                                    If (Perem(0, 1, i) = "�����") Then
                                        '���� ����� �������
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
                    '��������� ����������------------------------
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

            '��������� ����������
            Return SUMMA.ToString
            Exit Function
        End If
    End Function
    '--�������� (������� ��������� ������������)
    Private Sub ShowMessage(ByVal CodeLine As String)
        Dim DataL As String
        Dim FuncName, FuncParam As String
        Dim Token As String

        If (CodeLine.Chars(CodeLine.Length - 1) <> ";") Then
            '��������� ������
            SomovCompileModule.WError.ListBox1.Items.Add("��������������! � ������ [" + CodeLine + "] (����������� ������� ����� ������ ; ����� � �������)")
            SomovCompileModule.WError.Show()
            Exit Sub
        End If

        '���������� ���������� ��������
        For i As Integer = 0 To CodeLine.Length - 1
            If (CodeLine.Chars(i) = " ") Then
                If (Token = "��������(" Or Token = "��������(") Then
                    DataL = CodeLine.Remove(0, i + 1)
                    DataL = DataL.Remove(DataL.Length - 2, 2)
                    Token = ""
                    Exit For
                Else
                    Token = ""
                End If
            Else
                '�� �������� ��������� (WinXP 177580, WinSeven -842352729) ���������� � �����
                If (CodeLine.GetHashCode <> 177580 And CodeLine.GetHashCode <> -842352729 And CodeLine.Chars(i) <> " ") Then Token = Token + CodeLine.Chars(i)
            End If
        Next

        '--------------------------------------------------------------------------
        '���������� ���������
        If (DataL.Chars(0) = """" Or DataL.Chars(0) = "'") Then '�������� ������
            If (DataL.Chars(DataL.Length - 1) = """" Or DataL.Chars(DataL.Length - 1) = "'" And CodeLine.Chars(CodeLine.Length - 1) = ";") Then
                DataL = DataL.Remove(0, 1)
                DataL = DataL.Remove(DataL.Length - 1, 1)
                MsgBox(DataL.ToString, MsgBoxStyle.OKOnly, "��������:")
            Else
                '��������� ������
                SomovCompileModule.WError.ListBox1.Items.Add("������! � ������ [" + CodeLine + "] (��� ����������� ������� ��� ����������� ������� ����� ������ ; ����� � �������)")
                SomovCompileModule.WError.Show()
                Exit Sub
            End If
        Else    '�������� ���������� ��� �������
            '�������� ����������---------------------------------------------------
            For i As Integer = 1 To CORE.countPerem
                If (Perem(0, 0, i).ToString = DataL) Then
                    If (CodeLine.Chars(CodeLine.Length - 1) = ";") Then
                        MsgBox(Perem(1, 1, i).ToString, MsgBoxStyle.OKOnly, "��������:")
                        Exit Sub
                    Else
                        '��������� ������
                        SomovCompileModule.WError.ListBox1.Items.Add("������! � ������ [" + CodeLine + "] (����������� ������� ����� ������ ; ����� � �������)")
                        SomovCompileModule.WError.Show()
                    End If
                End If
            Next

            '�������� �������------------------------------------------------------
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
                        MsgBox(ExecuteFUNC(FuncName, FuncParam, Func(0, 1, i).ToString, Func(1, 1, i).ToString), MsgBoxStyle.OKOnly, "��������:")
                        Exit Sub
                    Else
                        '��������� ������
                        SomovCompileModule.WError.ListBox1.Items.Add("������! � ������ [" + CodeLine + "] (����������� ������� ����� ������ ; ����� � �������)")
                        SomovCompileModule.WError.Show()
                    End If

                End If
            Next
            '--------------------------------------------------
        End If
        Perem(1, 1, IndexStatusPerem) = "����� ���������!"
    End Sub
    '--���������� �������
    Private Function ExecuteFUNC(ByVal NameFunction As String, ByVal ParamFunction As String, ByVal TypeFunction As String, ByVal CodeFunction As String) As String
        Perem(1, 1, IndexStatusPerem) = "���������� �������!"

        Dim Token, stroka As String
        Dim FuncFind As Boolean = False '������� ��� ��������� ������� �������
        Dim Param(0) As String '������ ���� ���������� ����������
        Dim ParamIndex As Integer = 1
        Dim ParamBeginEnd As Boolean = False '���������� ������ �������� ���������� � �����)
        Dim FuncReturn As Boolean = False   '�������� ������� ��� ��� �� ��� ��������
        Dim ModuleCode As New RichTextBox
        ModuleCode.Text = CodeFunction

        '���������� ���������� �������� ���������� ����������
        If (ParamFunction <> "") Then
            '����� ���������� ���������� � �������� ����������
            For i As Integer = 0 To ParamFunction.Length - 1
                If (ParamFunction.Chars(i) = " " Or ParamFunction.Chars(i) = "," Or i = ParamFunction.Length - 1) Then
                    If (i = ParamFunction.Length - 1 And ParamFunction.GetHashCode <> 177580 And ParamFunction.GetHashCode <> -842352729 And ParamFunction.Chars(i) <> " ") Then Token = Token + ParamFunction.Chars(i)
                    If (Token <> "") Then
                        '� ������ ���� �������� ��� ��������: �����
                        If (Token.Chars(0) = "0" Or Token.Chars(0) = "1" Or Token.Chars(0) = "2" Or Token.Chars(0) = "3" Or Token.Chars(0) = "4" Or Token.Chars(0) = "5" Or Token.Chars(0) = "6" Or Token.Chars(0) = "7" Or Token.Chars(0) = "8" Or Token.Chars(0) = "9" Or Token.Chars(0) = "0") Then
                            ReDim Preserve Param(Param.Length)
                            Param(Param.Length - 1) = Token
                        Else    '� ������ ���� �������� ��� ��������: ������
                            If (Token.Chars(0) = """" Or Token.Chars(0) = "'") Then
                                stroka = ParamFunction.Remove(0, 1)
                                stroka = stroka.Remove(stroka.Length - 1, 1)
                                ReDim Preserve Param(Param.Length)
                                Param(Param.Length - 1) = stroka
                                stroka = ""
                                Exit For

                            Else    '� ������ ���� �������� ��� ����������
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
                    '�� �������� ��������� (WinXP 177580, WinSeven -842352729) ���������� � �����
                    If (ParamFunction.GetHashCode <> 177580 And ParamFunction.GetHashCode <> -842352729 And ParamFunction.Chars(i) <> " ") Then Token = Token + ParamFunction.Chars(i)
                End If
            Next
            '�������� ������� ���������������
            For i As Integer = 1 To TABLE_ID.Length - 1
                If (FuncFind = True And ParamBeginEnd = True And TABLE_ID(i) <> "�����" And TABLE_ID(i) <> "�����" And TABLE_ID(i) <> "," And TABLE_ID(i) <> " " And TABLE_ID(i) <> "���" And TABLE_ID(i) <> "���" And TABLE_ID(i) <> "������" And TABLE_ID(i) <> "������" And TABLE_ID(i) <> "�����" And TABLE_ID(i) <> "�����") Then
                    '�������� ���������� � �� ��������
                    For iPerem As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, iPerem).ToString = TABLE_ID(i)) Then
                            Perem(1, 1, iPerem) = Param(ParamIndex)
                            ParamIndex = ParamIndex + 1
                            Exit For
                        End If
                    Next
                End If
                If (TABLE_ID(i) = NameFunction) Then    '������� ������� � �������� ������
                    FuncFind = True '������� ����������
                End If
                If (TABLE_ID(i) = "(") Then
                    ParamBeginEnd = True '������� ������ ���������� �������
                End If
                If (TABLE_ID(i) = ")") Then
                    ParamBeginEnd = False '������� ������ ���������� �������
                    '��������� ���� �������
                    ExecuteCode(ModuleCode)
                    '��������� ������� �� �������
                    '�������� ������� ���������������
                    For iReturn As Integer = 1 To TABLE_ID.Length - 1
                        If (TABLE_ID(iReturn) = NameFunction) Then    '������� ������� � �������� ������
                            FuncReturn = True
                        End If
                        If (TABLE_ID(iReturn) = "�������" Or TABLE_ID(i) = "�������") Then
                            If (FuncReturn = True) Then
                                '�������� ���������� � �� ��������
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
        Else    '��������� ������� ��� ����������
            ExecuteCode(ModuleCode)
            '��������� ������� �� �������
            '�������� ������� ���������������
            For i As Integer = 1 To TABLE_ID.Length - 1
                If (TABLE_ID(i) = NameFunction) Then    '������� ������� � �������� ������
                    FuncReturn = True
                End If
                If (TABLE_ID(i) = "�������" Or TABLE_ID(i) = "�������") Then
                    If (FuncReturn = True) Then
                        '�������� ���������� � �� ��������
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
    '--��������� ���������
    Private Sub ExecutePROC(ByVal NameParamProcedure As String)
        Perem(1, 1, IndexStatusPerem) = "���������� ���������!"

        Dim Token, stroka As String
        Dim ProcFind As Boolean = False '������� ��� ��������� ������� ���������
        Dim Param(0) As String '������ ���� ���������� ����������
        Dim ParamIndex As Integer = 1
        Dim ParamBeginEnd As Boolean = False '���������� ������ �������� ���������� � �����)

        Dim ProcName, ProcParam As String
        Dim ModuleCode As New RichTextBox

        '������� ������ � ������
        If (NameParamProcedure.Chars(0) = " ") Then NameParamProcedure = NameParamProcedure.Remove(0, 1)

        '���������� ��� ���������
        For i As Integer = 0 To NameParamProcedure.Length - 1
            If (NameParamProcedure.Chars(i) = "(") Then
                ProcName = NameParamProcedure.Remove(i, NameParamProcedure.Length - i)
                ProcParam = NameParamProcedure.Remove(0, i + 2)
                If (ProcParam <> "") Then ProcParam = ProcParam.Remove(ProcParam.Length - 1, 1)
                Exit For
            End If
        Next
        '���������� ��������� ���������� � ���������
        For i As Integer = 1 To CORE.countProc
            If (Proc(0, 0, i).ToString = ProcName) Then
                ModuleCode.Text = Proc(1, 1, i).ToString
                Exit For
            End If
        Next

        '���� ��������� � ����������� ��������� �� ���������
        If (ProcParam <> "") Then
            '����� ���������� ���������� � �������� ����������
            For i As Integer = 0 To ProcParam.Length - 1
                If (ProcParam.Chars(i) = " " Or ProcParam.Chars(i) = "," Or i = ProcParam.Length - 1) Then
                    If (i = ProcParam.Length - 1 And ProcParam.GetHashCode <> 177580 And ProcParam.GetHashCode <> -842352729 And ProcParam.Chars(i) <> " ") Then Token = Token + ProcParam.Chars(i)
                    If (Token <> "") Then
                        '� ������ ���� �������� ��� ��������: �����
                        If (Token.Chars(0) = "0" Or Token.Chars(0) = "1" Or Token.Chars(0) = "2" Or Token.Chars(0) = "3" Or Token.Chars(0) = "4" Or Token.Chars(0) = "5" Or Token.Chars(0) = "6" Or Token.Chars(0) = "7" Or Token.Chars(0) = "8" Or Token.Chars(0) = "9" Or Token.Chars(0) = "0") Then
                            ReDim Preserve Param(Param.Length)
                            Param(Param.Length - 1) = Token
                        Else    '� ������ ���� �������� ��� ��������: ������
                            If (Token.Chars(0) = """" Or Token.Chars(0) = "'") Then
                                stroka = ProcParam.Remove(0, 1)
                                stroka = stroka.Remove(stroka.Length - 1, 1)
                                If (stroka.Chars(stroka.Length - 1) = """" Or stroka.Chars(stroka.Length - 1) = "'") Then stroka = stroka.Remove(stroka.Length - 1, 1)
                                ReDim Preserve Param(Param.Length)
                                Param(Param.Length - 1) = stroka
                                stroka = ""
                                Exit For
                            Else    '� ������ ���� �������� ��� ����������
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
                    '�� �������� ��������� (WinXP 177580, WinSeven -842352729) ���������� � �����
                    If (ProcParam.GetHashCode <> 177580 And ProcParam.GetHashCode <> -842352729 And ProcParam.Chars(i) <> " ") Then Token = Token + ProcParam.Chars(i)
                End If
            Next
            '�������� ������� ���������������
            For i As Integer = 1 To TABLE_ID.Length - 1
                If (ProcFind = True And ParamBeginEnd = True And TABLE_ID(i) <> "�����" And TABLE_ID(i) <> "�����" And TABLE_ID(i) <> "," And TABLE_ID(i) <> " " And TABLE_ID(i) <> "���" And TABLE_ID(i) <> "���" And TABLE_ID(i) <> "������" And TABLE_ID(i) <> "������" And TABLE_ID(i) <> "�����" And TABLE_ID(i) <> "�����") Then
                    '�������� ���������� � �� ��������
                    For iPerem As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, iPerem).ToString = TABLE_ID(i)) Then
                            Perem(1, 1, iPerem) = Param(ParamIndex)
                            ParamIndex = ParamIndex + 1
                            Exit For
                        End If
                    Next
                End If
                If (TABLE_ID(i) = ProcName) Then    '������� ��������� � �������� ������
                    ProcFind = True '������� ����������
                End If
                If (TABLE_ID(i) = "(") Then
                    ParamBeginEnd = True '������� ������ ���������� ���������
                End If
                If (TABLE_ID(i) = ")") Then
                    ParamBeginEnd = False '������� ������ ���������� ���������
                    Exit For
                End If
            Next
        End If

        '��������� ���� ���������
        ExecuteCode(ModuleCode)
    End Sub
    '--������� (�������� ���������� ��������� �������)
    Private Function VerificationIF(ByVal CodeLine As String) As Boolean
        Perem(1, 1, IndexStatusPerem) = "���������� �������!"
        Try
            Dim STROKA As String = CodeLine
            Dim Token As String
            Dim KV As Boolean = False   '����������� ������ � ����� ������
            Dim TYPE_IF As String = ""  '��� �������
            Dim LEFT_OPERATOR, RIGHT_OPERATOR As String '����� � ������ ��������
            Dim LEFT_OPERATOR_D, RIGHT_OPERATOR_D As Double '����� � ������ �������� (�����)

            STROKA = STROKA.Remove(0, 5)

            For i As Integer = 0 To STROKA.Length - 1
                If (STROKA.Chars(i) = """" And KV = False) Then
                    KV = True '����������� �������
                Else
                    If (STROKA.Chars(i) = """" And KV = True) Then
                        KV = False  '����������� �������
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
                    '�������� ������� -----------------------------
                    If (TYPE_IF = "") Then '���������� ����� �������
                        If (KV = False) Then LEFT_OPERATOR = Token
                    Else '���������� ������ �������
                        If (KV = False) Then RIGHT_OPERATOR = Token
                    End If
                    '----------------------------------------------

                    If (KV = False) Then Token = ""
                Else
                    '�� �������� ��������� (WinXP 177580, WinSeven -842352729) ���������� � �����
                    If (STROKA.GetHashCode <> 177580 And STROKA.GetHashCode <> -842352729 And STROKA.Chars(i) <> " ") Then Token = Token + STROKA.Chars(i)
                End If
            Next

            '���������� ��� � �������� ���������� ��������� *************
            '����� ������� ..............................................
            If (LEFT_OPERATOR.Chars(0) = "0" Or LEFT_OPERATOR.Chars(0) = "1" Or LEFT_OPERATOR.Chars(0) = "2" Or LEFT_OPERATOR.Chars(0) = "3" Or LEFT_OPERATOR.Chars(0) = "4" Or LEFT_OPERATOR.Chars(0) = "5" Or LEFT_OPERATOR.Chars(0) = "6" Or LEFT_OPERATOR.Chars(0) = "7" Or LEFT_OPERATOR.Chars(0) = "8" Or LEFT_OPERATOR.Chars(0) = "9" Or LEFT_OPERATOR.Chars(0) = "0") Then
                '������ ��� ����� � ������ ����
                '���� ����� �������
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
                    '������ ��� ������ � ������ ����
                    LEFT_OPERATOR = LEFT_OPERATOR.Remove(0, 1)
                    LEFT_OPERATOR = LEFT_OPERATOR.Remove(LEFT_OPERATOR.Length - 1, 1)
                Else
                    '������ ��� ��� ����������
                    For i As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, i).ToString = LEFT_OPERATOR) Then
                            LEFT_OPERATOR = Perem(1, 1, i).ToString '�������� ����������
                            If (LEFT_OPERATOR.Chars(0) = "0" Or LEFT_OPERATOR.Chars(0) = "1" Or LEFT_OPERATOR.Chars(0) = "2" Or LEFT_OPERATOR.Chars(0) = "3" Or LEFT_OPERATOR.Chars(0) = "4" Or LEFT_OPERATOR.Chars(0) = "5" Or LEFT_OPERATOR.Chars(0) = "6" Or LEFT_OPERATOR.Chars(0) = "7" Or LEFT_OPERATOR.Chars(0) = "8" Or LEFT_OPERATOR.Chars(0) = "9" Or LEFT_OPERATOR.Chars(0) = "0") Then
                                '������ ��� ����� � ������ ����
                                '���� ����� �������
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
            '������ ������� ..............................................
            If (RIGHT_OPERATOR.Chars(0) = "0" Or RIGHT_OPERATOR.Chars(0) = "1" Or RIGHT_OPERATOR.Chars(0) = "2" Or RIGHT_OPERATOR.Chars(0) = "3" Or RIGHT_OPERATOR.Chars(0) = "4" Or RIGHT_OPERATOR.Chars(0) = "5" Or RIGHT_OPERATOR.Chars(0) = "6" Or RIGHT_OPERATOR.Chars(0) = "7" Or RIGHT_OPERATOR.Chars(0) = "8" Or RIGHT_OPERATOR.Chars(0) = "9" Or RIGHT_OPERATOR.Chars(0) = "0") Then
                '������ ��� ����� � ������ ����
                '���� ����� �������
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
                    '������ ��� ������ � ������ ����
                    RIGHT_OPERATOR = RIGHT_OPERATOR.Remove(0, 1)
                    RIGHT_OPERATOR = RIGHT_OPERATOR.Remove(RIGHT_OPERATOR.Length - 1, 1)
                Else
                    '������ ��� ��� ����������
                    For i As Integer = 1 To CORE.countPerem
                        If (Perem(0, 0, i).ToString = RIGHT_OPERATOR) Then
                            RIGHT_OPERATOR = Perem(1, 1, i).ToString '�������� ����������
                            If (RIGHT_OPERATOR.Chars(0) = "0" Or RIGHT_OPERATOR.Chars(0) = "1" Or RIGHT_OPERATOR.Chars(0) = "2" Or RIGHT_OPERATOR.Chars(0) = "3" Or RIGHT_OPERATOR.Chars(0) = "4" Or RIGHT_OPERATOR.Chars(0) = "5" Or RIGHT_OPERATOR.Chars(0) = "6" Or RIGHT_OPERATOR.Chars(0) = "7" Or RIGHT_OPERATOR.Chars(0) = "8" Or RIGHT_OPERATOR.Chars(0) = "9" Or RIGHT_OPERATOR.Chars(0) = "0") Then
                                '���� ����� �������
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
            '���������� ***********************************************
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
            '��������� ������
            SomovCompileModule.WError.ListBox1.Items.Add("������! � ������ [" + CodeLine + "] (������� ������ �������)")
            SomovCompileModule.WError.Show()
            Exit Function
        End Try
    End Function
    '--��������� ��������� ������� ����� "���" � ���������� ���
    Private Sub ExecuteFOR(ByVal CodeLine As String, ByVal ModuleCode As RichTextBox)
        Perem(1, 1, IndexStatusPerem) = "���������� �����!"
        Try
            Dim STROKA As String = CodeLine
            Dim Token As String
            Dim Ravno As Boolean = False
            Dim PO As Boolean = False
            Dim ActionPerem As String = ""      '��� ����������
            Dim StartNumCycle As Integer = 0    '��������� �������� �����
            Dim KolCycle As Integer = 0         '���������� ������

            TYPE_prisvoit = False
            STROKA = STROKA.Remove(0, 6)
            STROKA = STROKA.Remove(STROKA.Length - 1, 1)

            For i As Integer = 0 To STROKA.Length - 1
                If (STROKA.Chars(i) = " " Or i = STROKA.Length - 1) Then

                    If (PO = False) Then    '�������� ��

                        If (Ravno = False) Then '���� ���������
                            If (ActionPerem = "") Then ActionPerem = Token
                        Else
                            '**�������������� ��������*****************
                            If (SyntaxPerem(ActionPerem) = True) Then
                                For iPerem As Integer = 1 To CORE.countPerem
                                    If (Perem(0, 0, iPerem).ToString = ActionPerem) Then
                                        '�������� ����������
                                        Perem(1, 1, iPerem) = Token
                                        StartNumCycle = Token
                                        Ravno = False
                                        Exit For
                                    End If
                                Next
                            Else
                                '**���������� ������� *********************
                                countPerem = countPerem + 1 '���� ���������� ����������
                                ReDim Preserve Perem(1, 1, countPerem) '����������� ������
                                '**��� ����������**************************
                                Perem(0, 0, countPerem) = ActionPerem
                                '**��� �����*******************************
                                Perem(0, 1, countPerem) = "�����"
                                '�������� ����������
                                Perem(1, 1, countPerem) = Token
                                StartNumCycle = Token
                                Ravno = False
                            End If

                        End If
                    Else
                        '�� �������� ��������� (WinXP 177580, WinSeven -842352729) ���������� � �����
                        If (STROKA.GetHashCode <> 177580 And STROKA.GetHashCode <> -842352729 And STROKA.Chars(i) <> " ") Then Token = Token + STROKA.Chars(i)
                        KolCycle = Token
                        Exit For
                    End If
                    If (Token = "=") Then Ravno = True
                    If (Token = "��" Or Token = "��") Then PO = True
                    Token = ""
                Else
                    '�� �������� ��������� (WinXP 177580, WinSeven -842352729) ���������� � �����
                    If (STROKA.GetHashCode <> 177580 And STROKA.GetHashCode <> -842352729 And STROKA.Chars(i) <> " ") Then Token = Token + STROKA.Chars(i)
                End If
            Next

            '���������� �����
            For iExecute As Integer = StartNumCycle To KolCycle
                For iPerem As Integer = 1 To CORE.countPerem
                    If (Perem(0, 0, iPerem).ToString = ActionPerem) Then
                        '�������� ����������
                        Perem(1, 1, countPerem) = iExecute.ToString
                        Exit For
                    End If
                Next
                ExecuteCode(ModuleCode)
            Next

        Catch ex As Exception
            '��������� ������
            SomovCompileModule.WError.ListBox1.Items.Add("������! � ������ [" + CodeLine + "] (������� ������ �������)")
            SomovCompileModule.WError.Show()
            Exit Sub
        End Try
    End Sub


    '//�������///////////////////////////////////////////////////////////////////
    Private Sub ClearTYPE() '�������
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

    Public Sub ClearModule() '������� ������.
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
