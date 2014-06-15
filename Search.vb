Public Class Search
    Inherits System.Windows.Forms.Form

#Region " Код, автоматически созданный конструктором форм Windows "

    Public Sub New()
        MyBase.New()

        'Этот вызов требуется конструктором форм Windows.
        InitializeComponent()

        'Добавьте код инициализации после вызова InitializeComponent()

    End Sub

    'Форма переопределяет метод Dispose для очистки списка компонентов.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Требуется конструктором форм Windows
    Private components As System.ComponentModel.IContainer

    'ПРИМЕЧАНИЕ: следующая процедура требуется для конструктора форм Windows.
    'Ее можно изменить в конструкторе форм Windows.  
    'Не изменяйте ее в редакторе исходного текста.
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Search))
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox1.Location = New System.Drawing.Point(8, 8)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(248, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(260, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(64, 28)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Поиск."
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Search
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(330, 38)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ComboBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Search"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Поиск:"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Search_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        WSearchShow = True
    End Sub

    Private Sub Search_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        WSearchShow = False
    End Sub

    Dim FindIndex As Integer = 0
    Dim FindText As String
    Dim FindLast As Integer = 0
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            'Сохранение слов поиска
            Dim AddFind As Boolean = True
            For i As Integer = 0 To ComboBox1.Items.Count - 1
                If (ComboBox1.Items.Item(i) = ComboBox1.Text) Then
                    AddFind = False
                    Exit For
                End If
            Next
            If (AddFind = True) Then ComboBox1.Items.Add(ComboBox1.Text)

            'Поиск в модуле
            If (FindText <> ComboBox1.Text) Then
                FindIndex = 0
                FindLast = 0
                FindText = ComboBox1.Text
            End If
            If (SomovCompileModule.SCompile.RichTextBox1.Find(ComboBox1.Text, FindIndex, SomovCompileModule.SCompile.RichTextBox1.TextLength - 1, RichTextBoxFinds.None)) Then
                SomovCompileModule.SCompile.RichTextBox1.Select()
                FindIndex = SomovCompileModule.SCompile.RichTextBox1.SelectionStart + SomovCompileModule.SCompile.RichTextBox1.SelectionLength
                If (FindLast = SomovCompileModule.SCompile.RichTextBox1.SelectionStart) Then
                    MsgBox("Поиск завершен.", MsgBoxStyle.OKOnly, "Сообщение:")
                    FindIndex = 0
                    FindLast = 0
                    FindText = ComboBox1.Text
                Else
                    FindLast = SomovCompileModule.SCompile.RichTextBox1.SelectionStart
                    SomovCompileModule.SCompile.Focus()
                End If
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub
End Class
