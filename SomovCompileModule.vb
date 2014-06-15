Module SomovCompileModule
    Public PathBase As String
    Public NameBase As String
    Public StringConnection As String
    Public MyProgramDirectory As String

    Public FileOpenMyProgram As String

    '������� ���� �����������
    Public SCompile As Compiler
    Public Sub LoadSCompile()
        SCompile = New Compiler
    End Sub

    '���� � ���������
    Public AboutShow As Boolean = False
    Public SCAbout As About
    Public Sub LoadAbout()
        SCAbout = New About
    End Sub

    '���� ��������� �� �������
    Public WErrorShow As Boolean = False
    Public WError As WindowERROR
    Public Sub LoadWError()
        WError = New WindowERROR
    End Sub

    '���� ������
    Public WSearchShow As Boolean = False
    Public WSearch As Search
    Public Sub LoadWSearch()
        WSearch = New Search
    End Sub
End Module
