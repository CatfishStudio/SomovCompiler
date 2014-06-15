Imports System.IO
Imports System.Text
Public Class Compiler
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents DataSet11 As SomovCompiler.DataSet1
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem26 As System.Windows.Forms.MenuItem
    Friend WithEvents DataSet21 As SomovCompiler.DataSet2
    Friend WithEvents MenuItem27 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem28 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem31 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem32 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem34 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem35 As System.Windows.Forms.MenuItem
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents FontDialog1 As System.Windows.Forms.FontDialog
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem37 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem38 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem39 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem40 As System.Windows.Forms.MenuItem
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ContextMenu3 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem41 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem42 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem43 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem44 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem45 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem46 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem47 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem48 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem49 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem50 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Compiler))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.MenuItem18 = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.MenuItem38 = New System.Windows.Forms.MenuItem
        Me.MenuItem39 = New System.Windows.Forms.MenuItem
        Me.MenuItem40 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuItem26 = New System.Windows.Forms.MenuItem
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox
        Me.ContextMenu3 = New System.Windows.Forms.ContextMenu
        Me.MenuItem41 = New System.Windows.Forms.MenuItem
        Me.MenuItem42 = New System.Windows.Forms.MenuItem
        Me.MenuItem43 = New System.Windows.Forms.MenuItem
        Me.MenuItem44 = New System.Windows.Forms.MenuItem
        Me.MenuItem45 = New System.Windows.Forms.MenuItem
        Me.MenuItem46 = New System.Windows.Forms.MenuItem
        Me.MenuItem47 = New System.Windows.Forms.MenuItem
        Me.MenuItem48 = New System.Windows.Forms.MenuItem
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem27 = New System.Windows.Forms.MenuItem
        Me.MenuItem28 = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.MenuItem30 = New System.Windows.Forms.MenuItem
        Me.MenuItem31 = New System.Windows.Forms.MenuItem
        Me.MenuItem32 = New System.Windows.Forms.MenuItem
        Me.MenuItem33 = New System.Windows.Forms.MenuItem
        Me.MenuItem34 = New System.Windows.Forms.MenuItem
        Me.MenuItem49 = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.MenuItem37 = New System.Windows.Forms.MenuItem
        Me.MenuItem50 = New System.Windows.Forms.MenuItem
        Me.MenuItem35 = New System.Windows.Forms.MenuItem
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.DataSet11 = New SomovCompiler.DataSet1
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.DataSet21 = New SomovCompiler.DataSet2
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog
        Me.FontDialog1 = New System.Windows.Forms.FontDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel6.SuspendLayout()
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet21, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem15, Me.MenuItem38, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem12, Me.MenuItem13, Me.MenuItem7, Me.MenuItem10, Me.MenuItem8, Me.MenuItem9, Me.MenuItem11, Me.MenuItem14})
        Me.MenuItem1.Text = "&Файл"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 0
        Me.MenuItem12.Shortcut = System.Windows.Forms.Shortcut.F2
        Me.MenuItem12.Text = "Создать новый модуль"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 1
        Me.MenuItem13.Text = "-"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.Shortcut = System.Windows.Forms.Shortcut.F3
        Me.MenuItem7.Text = "Открыть файл."
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 3
        Me.MenuItem10.Text = "-"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 4
        Me.MenuItem8.Shortcut = System.Windows.Forms.Shortcut.F4
        Me.MenuItem8.Text = "Сохранить файл."
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 5
        Me.MenuItem9.Text = "Сохранить как..."
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 6
        Me.MenuItem11.Text = "-"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 7
        Me.MenuItem14.Shortcut = System.Windows.Forms.Shortcut.F12
        Me.MenuItem14.Text = "Выход."
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 1
        Me.MenuItem15.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem16, Me.MenuItem17, Me.MenuItem18, Me.MenuItem19, Me.MenuItem20, Me.MenuItem21, Me.MenuItem22, Me.MenuItem23, Me.MenuItem24})
        Me.MenuItem15.Text = "&Правка"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 0
        Me.MenuItem16.Shortcut = System.Windows.Forms.Shortcut.CtrlZ
        Me.MenuItem16.Text = "Отменить."
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 1
        Me.MenuItem17.Text = "-"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 2
        Me.MenuItem18.Shortcut = System.Windows.Forms.Shortcut.CtrlX
        Me.MenuItem18.Text = "Вырезать."
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 3
        Me.MenuItem19.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.MenuItem19.Text = "Копировать."
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 4
        Me.MenuItem20.Shortcut = System.Windows.Forms.Shortcut.CtrlV
        Me.MenuItem20.Text = "Вставить."
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 5
        Me.MenuItem21.Text = "Удалить."
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 6
        Me.MenuItem22.Text = "-"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 7
        Me.MenuItem23.Shortcut = System.Windows.Forms.Shortcut.CtrlA
        Me.MenuItem23.Text = "Выделить всё."
        '
        'MenuItem24
        '
        Me.MenuItem24.Index = 8
        Me.MenuItem24.Shortcut = System.Windows.Forms.Shortcut.CtrlF
        Me.MenuItem24.Text = "Найти."
        '
        'MenuItem38
        '
        Me.MenuItem38.Index = 2
        Me.MenuItem38.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem39, Me.MenuItem40})
        Me.MenuItem38.Text = "&Вид"
        '
        'MenuItem39
        '
        Me.MenuItem39.Index = 0
        Me.MenuItem39.Text = "Окно ошибок выполнения."
        '
        'MenuItem40
        '
        Me.MenuItem40.Index = 1
        Me.MenuItem40.Text = "Окно поиска."
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5, Me.MenuItem6})
        Me.MenuItem2.Text = "&Отладка"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Shortcut = System.Windows.Forms.Shortcut.F5
        Me.MenuItem5.Text = "Отладить модуль."
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.Shortcut = System.Windows.Forms.Shortcut.F6
        Me.MenuItem6.Text = "Построить проект."
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 4
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4})
        Me.MenuItem3.Text = "&Справка"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.Text = "О программе."
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 429)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(642, 16)
        Me.StatusBar1.TabIndex = 1
        Me.StatusBar1.Text = "StatusBar1"
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.StatusBarPanel1.Text = "Строка:"
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.Text = "Модуль не сохранён!"
        Me.StatusBarPanel2.Width = 1000
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(642, 72)
        Me.Panel1.TabIndex = 2
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Transparent
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Location = New System.Drawing.Point(216, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 64)
        Me.Button3.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.Button3, "Отладить модуль.")
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Transparent
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(112, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 64)
        Me.Button2.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Button2, "Сохранить файл.")
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Transparent
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(8, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(96, 64)
        Me.Button1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Button1, "Открыть файл.")
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BackColor = System.Drawing.Color.Transparent
        Me.Panel4.BackgroundImage = CType(resources.GetObject("Panel4.BackgroundImage"), System.Drawing.Image)
        Me.Panel4.Location = New System.Drawing.Point(490, 8)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(143, 56)
        Me.Panel4.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.TreeView1)
        Me.Panel2.Controls.Add(Me.Splitter2)
        Me.Panel2.Controls.Add(Me.Panel5)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel2.Location = New System.Drawing.Point(394, 72)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(248, 357)
        Me.Panel2.TabIndex = 3
        '
        'TreeView1
        '
        Me.TreeView1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TreeView1.ContextMenu = Me.ContextMenu2
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.ImageIndex = 2
        Me.TreeView1.ImageList = Me.ImageList1
        Me.TreeView1.Indent = 27
        Me.TreeView1.ItemHeight = 24
        Me.TreeView1.Location = New System.Drawing.Point(0, 0)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(246, 232)
        Me.TreeView1.TabIndex = 2
        '
        'ContextMenu2
        '
        Me.ContextMenu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem25, Me.MenuItem26})
        '
        'MenuItem25
        '
        Me.MenuItem25.Checked = True
        Me.MenuItem25.Index = 0
        Me.MenuItem25.Text = "По алфавиту (возрастание)"
        '
        'MenuItem26
        '
        Me.MenuItem26.Index = 1
        Me.MenuItem26.Text = "По алфавиту (убывание)"
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(24, 20)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Splitter2
        '
        Me.Splitter2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Splitter2.Location = New System.Drawing.Point(0, 232)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(246, 3)
        Me.Splitter2.TabIndex = 1
        Me.Splitter2.TabStop = False
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.RichTextBox2)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel5.Location = New System.Drawing.Point(0, 235)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(246, 120)
        Me.Panel5.TabIndex = 0
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox2.BackColor = System.Drawing.SystemColors.Control
        Me.RichTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox2.ContextMenu = Me.ContextMenu3
        Me.RichTextBox2.Location = New System.Drawing.Point(0, 0)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(242, 115)
        Me.RichTextBox2.TabIndex = 0
        Me.RichTextBox2.Text = ""
        Me.RichTextBox2.WordWrap = False
        '
        'ContextMenu3
        '
        Me.ContextMenu3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem41, Me.MenuItem42, Me.MenuItem43, Me.MenuItem44, Me.MenuItem45, Me.MenuItem46, Me.MenuItem47, Me.MenuItem48})
        '
        'MenuItem41
        '
        Me.MenuItem41.Index = 0
        Me.MenuItem41.Text = "Отменить."
        '
        'MenuItem42
        '
        Me.MenuItem42.Index = 1
        Me.MenuItem42.Text = "-"
        '
        'MenuItem43
        '
        Me.MenuItem43.Index = 2
        Me.MenuItem43.Text = "Вырезать."
        '
        'MenuItem44
        '
        Me.MenuItem44.Index = 3
        Me.MenuItem44.Text = "Копировать."
        '
        'MenuItem45
        '
        Me.MenuItem45.Index = 4
        Me.MenuItem45.Text = "Вставить."
        '
        'MenuItem46
        '
        Me.MenuItem46.Index = 5
        Me.MenuItem46.Text = "Удалить."
        '
        'MenuItem47
        '
        Me.MenuItem47.Index = 6
        Me.MenuItem47.Text = "-"
        '
        'MenuItem48
        '
        Me.MenuItem48.Index = 7
        Me.MenuItem48.Text = "Выделить всё."
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Splitter1.Location = New System.Drawing.Point(391, 72)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 357)
        Me.Splitter1.TabIndex = 4
        Me.Splitter1.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Panel6)
        Me.Panel3.Controls.Add(Me.RichTextBox1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 72)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(391, 357)
        Me.Panel3.TabIndex = 5
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.TextBox2)
        Me.Panel6.Controls.Add(Me.ListBox1)
        Me.Panel6.Location = New System.Drawing.Point(64, 152)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(224, 104)
        Me.Panel6.TabIndex = 22
        Me.Panel6.Visible = False
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox2.Location = New System.Drawing.Point(0, 0)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(224, 20)
        Me.TextBox2.TabIndex = 1
        Me.TextBox2.Text = ""
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListBox1.Location = New System.Drawing.Point(0, 21)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(224, 80)
        Me.ListBox1.TabIndex = 0
        '
        'RichTextBox1
        '
        Me.RichTextBox1.AcceptsTab = True
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.ContextMenu = Me.ContextMenu1
        Me.RichTextBox1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox1.Location = New System.Drawing.Point(0, 0)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(389, 355)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem27, Me.MenuItem28, Me.MenuItem29, Me.MenuItem30, Me.MenuItem31, Me.MenuItem32, Me.MenuItem33, Me.MenuItem34, Me.MenuItem49, Me.MenuItem36, Me.MenuItem37, Me.MenuItem50, Me.MenuItem35})
        '
        'MenuItem27
        '
        Me.MenuItem27.Index = 0
        Me.MenuItem27.Shortcut = System.Windows.Forms.Shortcut.CtrlZ
        Me.MenuItem27.Text = "Отменить."
        '
        'MenuItem28
        '
        Me.MenuItem28.Index = 1
        Me.MenuItem28.Text = "-"
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 2
        Me.MenuItem29.Shortcut = System.Windows.Forms.Shortcut.CtrlX
        Me.MenuItem29.Text = "Вырезать."
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 3
        Me.MenuItem30.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.MenuItem30.Text = "Копировать."
        '
        'MenuItem31
        '
        Me.MenuItem31.Index = 4
        Me.MenuItem31.Shortcut = System.Windows.Forms.Shortcut.CtrlV
        Me.MenuItem31.Text = "Вставить."
        '
        'MenuItem32
        '
        Me.MenuItem32.Index = 5
        Me.MenuItem32.Text = "Удалить."
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 6
        Me.MenuItem33.Text = "-"
        '
        'MenuItem34
        '
        Me.MenuItem34.Index = 7
        Me.MenuItem34.Shortcut = System.Windows.Forms.Shortcut.CtrlA
        Me.MenuItem34.Text = "Выделить всё."
        '
        'MenuItem49
        '
        Me.MenuItem49.Index = 8
        Me.MenuItem49.Text = "-"
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 9
        Me.MenuItem36.Text = "Шрифт текста."
        '
        'MenuItem37
        '
        Me.MenuItem37.Index = 10
        Me.MenuItem37.Text = "Цвет текста."
        '
        'MenuItem50
        '
        Me.MenuItem50.Index = 11
        Me.MenuItem50.Text = "-"
        '
        'MenuItem35
        '
        Me.MenuItem35.Index = 12
        Me.MenuItem35.Shortcut = System.Windows.Forms.Shortcut.CtrlF
        Me.MenuItem35.Text = "Найти."
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, Тип" & _
        "Объекта, ФайлВПапке, ФайлИдентификатор FROM manual"
        Me.OleDbSelectCommand1.Connection = Me.OleDbConnection1
        '
        'OleDbConnection1
        '
        Me.OleDbConnection1.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""C:\ВНУ\Ответы\Системное ПО\SomovCompiler\База данных\" & _
        "help.mdb"";Jet OLEDB:Engine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:S" & _
        "ystem database=;Jet OLEDB:SFP=False;persist security info=False;Extended Propert" & _
        "ies=;Mode=ReadWrite;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create System Dat" & _
        "abase=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Witho" & _
        "ut Replica Repair=False;User ID=Admin;Jet OLEDB:Global Bulk Transactions=1"
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO manual(ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла," & _
        " ТипОбъекта, ФайлВПапке, ФайлИдентификатор) VALUES (?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.OleDbConnection1
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, "ДатаПоследнегоСохранения"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ПапкаИдентификатор"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("СодержаниеФайла", System.Data.OleDb.OleDbType.VarWChar, 0, "СодержаниеФайла"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, "ТипОбъекта"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлВПапке"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлИдентификатор"))
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE manual SET ДатаПоследнегоСохранения = ?, ПапкаИдентификатор = ?, Содержани" & _
        "еФайла = ?, ТипОбъекта = ?, ФайлВПапке = ?, ФайлИдентификатор = ? WHERE (Строка " & _
        "= ?) AND (ДатаПоследнегоСохранения = ? OR ? IS NULL AND ДатаПоследнегоСохранения" & _
        " IS NULL) AND (ПапкаИдентификатор = ? OR ? IS NULL AND ПапкаИдентификатор IS NUL" & _
        "L) AND (ТипОбъекта = ? OR ? IS NULL AND ТипОбъекта IS NULL) AND (ФайлВПапке = ? " & _
        "OR ? IS NULL AND ФайлВПапке IS NULL) AND (ФайлИдентификатор = ? OR ? IS NULL AND" & _
        " ФайлИдентификатор IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.OleDbConnection1
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, "ДатаПоследнегоСохранения"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ПапкаИдентификатор"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("СодержаниеФайла", System.Data.OleDb.OleDbType.VarWChar, 0, "СодержаниеФайла"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, "ТипОбъекта"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлВПапке"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлИдентификатор"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM manual WHERE (Строка = ?) AND (ДатаПоследнегоСохранения = ? OR ? IS N" & _
        "ULL AND ДатаПоследнегоСохранения IS NULL) AND (ПапкаИдентификатор = ? OR ? IS NU" & _
        "LL AND ПапкаИдентификатор IS NULL) AND (ТипОбъекта = ? OR ? IS NULL AND ТипОбъек" & _
        "та IS NULL) AND (ФайлВПапке = ? OR ? IS NULL AND ФайлВПапке IS NULL) AND (ФайлИд" & _
        "ентификатор = ? OR ? IS NULL AND ФайлИдентификатор IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.OleDbConnection1
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDataAdapter1
        '
        Me.OleDbDataAdapter1.DeleteCommand = Me.OleDbDeleteCommand1
        Me.OleDbDataAdapter1.InsertCommand = Me.OleDbInsertCommand1
        Me.OleDbDataAdapter1.SelectCommand = Me.OleDbSelectCommand1
        Me.OleDbDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "manual", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ДатаПоследнегоСохранения", "ДатаПоследнегоСохранения"), New System.Data.Common.DataColumnMapping("ПапкаИдентификатор", "ПапкаИдентификатор"), New System.Data.Common.DataColumnMapping("СодержаниеФайла", "СодержаниеФайла"), New System.Data.Common.DataColumnMapping("Строка", "Строка"), New System.Data.Common.DataColumnMapping("ТипОбъекта", "ТипОбъекта"), New System.Data.Common.DataColumnMapping("ФайлВПапке", "ФайлВПапке"), New System.Data.Common.DataColumnMapping("ФайлИдентификатор", "ФайлИдентификатор")})})
        Me.OleDbDataAdapter1.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'DataSet11
        '
        Me.DataSet11.DataSetName = "DataSet1"
        Me.DataSet11.Locale = New System.Globalization.CultureInfo("ru-RU")
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "Модуль исходного кода (*.scm)|*.scm"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.Filter = "Модуль исходного кода (*.scm)|*.scm"
        '
        'DataSet21
        '
        Me.DataSet21.DataSetName = "DataSet2"
        Me.DataSet21.Locale = New System.Globalization.CultureInfo("ru-RU")
        '
        'Compiler
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(642, 445)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.StatusBar1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "Compiler"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Somov Compiler v 1.0"
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet21, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim ListFocus As Boolean = False

    Private Sub Compiler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LOAD_OPERATORS()
        TREE_UPDATE()
        RichTextBox1.LoadFile(SomovCompileModule.MyProgramDirectory + "resource\newModule.rtf", RichTextBoxStreamType.RichText)
    End Sub

    Private Sub Compiler_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        'end
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        If (SomovCompileModule.AboutShow = False) Then
            SomovCompileModule.LoadAbout()
            SomovCompileModule.SCAbout.Show()
        End If
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.RichText)
            StatusBarPanel2.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        If (SaveFileDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText)
            StatusBarPanel2.Text = SaveFileDialog1.FileName
        End If
    End Sub

    Public Sub TREE_UPDATE()
        Try
            OleDbConnection1.ConnectionString = SomovCompileModule.StringConnection
            OleDbConnection1.Open()

            DataSet11.Clear()
            DataSet21.Clear()

            Dim CommandSQL As String

            If (MenuItem25.Checked = True) Then CommandSQL = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM manual ORDER BY ПапкаИдентификатор ASC"
            If (MenuItem26.Checked = True) Then CommandSQL = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM manual ORDER BY ПапкаИдентификатор DESC"

            OleDbDataAdapter1.SelectCommand.CommandText = CommandSQL
            OleDbDataAdapter1.Fill(DataSet11, "manual")

            CommandSQL = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM manual ORDER BY Строка ASC"
            OleDbDataAdapter1.SelectCommand.CommandText = CommandSQL
            OleDbDataAdapter1.Fill(DataSet21, "manual")

            'загрузка дерева
            Dim NameGr As String
            Dim ActionGr, ActionEl As Integer
            Dim NextGr As Boolean = False
            TreeView1.Nodes.Clear()
            For i As Integer = 0 To DataSet11.manual.Rows.Count - 1
                If (DataSet11.manual.Item(i)("ТипОбъекта") = "Группа") Then
                    NameGr = DataSet11.manual.Item(i)("ПапкаИдентификатор")
                    TreeView1.Nodes.Add(NameGr)
                    TreeView1.Nodes.Item(ActionGr).ImageIndex = 0
                    TreeView1.Nodes.Item(ActionGr).SelectedImageIndex = 1
                    ActionGr = ActionGr + 1
                    NextGr = True
                End If

                If (NextGr = True) Then
                    ActionEl = 0
                    For iEl As Integer = 0 To DataSet21.manual.Rows.Count - 1
                        If (DataSet21.manual.Item(iEl)("ТипОбъекта") = "Элемент") Then
                            If (DataSet21.manual.Item(iEl)("ФайлВПапке") = NameGr) And (DataSet21.manual.Item(iEl)("ТипОбъекта") = "Элемент") Then
                                TreeView1.Nodes.Item(ActionGr - 1).Nodes.Add(DataSet21.manual.Item(iEl)("ФайлИдентификатор"))
                                TreeView1.Nodes.Item(ActionGr - 1).Nodes.Item(ActionEl).ImageIndex = 2
                                TreeView1.Nodes.Item(ActionGr - 1).Nodes.Item(ActionEl).SelectedImageIndex = 2
                                ActionEl = ActionEl + 1
                            End If
                        End If
                    Next
                    NextGr = False
                End If
            Next
            OleDbConnection1.Close()
            TreeView1.Select()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click
        'Отмена
        RichTextBox1.Undo()
    End Sub

    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click
        'Вырезать
        RichTextBox1.Cut()
    End Sub

    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem19.Click
        'Копировать
        RichTextBox1.Copy()
    End Sub

    Private Sub MenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem20.Click
        'Вставить
        RichTextBox1.Paste()
    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem21.Click
        'Удалить
        RichTextBox1.SelectedText = ""
    End Sub

    Private Sub MenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem23.Click
        'Выделить всё
        RichTextBox1.SelectAll()
    End Sub

    Private Sub MenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem24.Click
        'Найти
        If (SomovCompileModule.WSearchShow = False) Then
            SomovCompileModule.LoadWSearch()
            SomovCompileModule.WSearch.Show()
        End If
    End Sub

    Private Sub MenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem27.Click
        'Отмена
        RichTextBox1.Undo()
    End Sub

    Private Sub MenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem29.Click
        'Вырезать
        RichTextBox1.Cut()
    End Sub

    Private Sub MenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem30.Click
        'Копировать
        RichTextBox1.Copy()
    End Sub

    Private Sub MenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem31.Click
        'Вставить
        RichTextBox1.Paste()
    End Sub

    Private Sub MenuItem32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem32.Click
        'Удалить
        RichTextBox1.SelectedText = ""
    End Sub

    Private Sub MenuItem34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem34.Click
        'Выделить всё
        RichTextBox1.SelectAll()
    End Sub

    Private Sub MenuItem35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem35.Click
        'Найти
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Try
            For i As Integer = 0 To DataSet11.manual.Count - 1
                If (DataSet11.manual.Item(i)("ТипОбъекта") = "Группа") Then
                    If (DataSet11.manual.Item(i)("ПапкаИдентификатор") = TreeView1.SelectedNode.Text) Then
                        RichTextBox2.Text = ""
                        Exit For
                    End If
                End If
                If (DataSet11.manual.Item(i)("ТипОбъекта") = "Элемент") Then
                    If (DataSet11.manual.Item(i)("ФайлИдентификатор") = TreeView1.SelectedNode.Text) Then
                        RichTextBox2.Rtf = DataSet11.manual.Item(i)("СодержаниеФайла")
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged
        Try
            Dim nStr, nChar As Integer
            nStr = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart) + 1
            StatusBarPanel1.Text = "Строка : " + nStr.ToString
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub RichTextBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextBox1.MouseMove
        Try
            Dim nStr, nChar As Integer
            nStr = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart) + 1
            StatusBarPanel1.Text = "Строка : " + nStr.ToString

        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub RichTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RichTextBox1.KeyDown
        Try
            If (e.Control = True) And (e.KeyValue = 32) Then
                Dim WindowH, smallWinH As Integer
                WindowH = Me.Height
                smallWinH = RichTextBox1.GetPositionFromCharIndex(RichTextBox1.SelectionStart).Y + 100 + Panel6.Height
                If (WindowH < smallWinH + 50) Then
                    TextBox2.Left = 0
                    TextBox2.Top = 81
                    ListBox1.Left = 0
                    ListBox1.Top = 0
                    Panel6.Left = RichTextBox1.GetPositionFromCharIndex(RichTextBox1.SelectionStart).X
                    Panel6.Top = RichTextBox1.GetPositionFromCharIndex(RichTextBox1.SelectionStart).Y - 80
                    Panel6.Visible = True
                    TextBox2.Focus()
                Else
                    TextBox2.Left = 0
                    TextBox2.Top = 0
                    ListBox1.Left = 0
                    ListBox1.Top = 21
                    Panel6.Left = RichTextBox1.GetPositionFromCharIndex(RichTextBox1.SelectionStart).X
                    Panel6.Top = RichTextBox1.GetPositionFromCharIndex(RichTextBox1.SelectionStart).Y '+ 40
                    Panel6.Visible = True
                    TextBox2.Focus()
                End If

            End If
            If (e.KeyData = Keys.Escape) Then 'закрытие без выбора
                Panel6.Visible = False
                ListBox1.SelectedIndex = 0
                TextBox2.Text = ""
                ListFocus = False
                RichTextBox1.Focus()
            End If
            If (e.KeyValue > 64) And (e.KeyValue <> 67) And (e.KeyValue <> 86) And (e.KeyValue <> 88) And (e.KeyValue < 91) Then 'символы черным цветом
                RichTextBox1.SelectionColor = Color.Black
            End If
            If (e.KeyData = Keys.Enter) Then
                'ВЫДИЛЕНИЕ СПЕЦСИМВОЛОВ ЦВЕТОМ
                Dim BeginOperator, EndOperator As Integer
                EndOperator = RichTextBox1.SelectionStart
                BeginOperator = EndOperator - 1
                RichTextBox1.SelectionStart = BeginOperator
                RichTextBox1.SelectionLength = 1

                If (RichTextBox1.SelectedText = "(") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "!") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "&") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = ")") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "*") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "-") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "=") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "+") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "|") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "{") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = ";") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "'") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "<") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "}") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = """") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "/") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = ">") Then
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = " ") Then
                    RichTextBox1.SelectionColor = Color.Black
                End If

                RichTextBox1.SelectionStart = EndOperator
                RichTextBox1.SelectionLength = 0
                RichTextBox1.SelectionColor = Color.Black
                '---------------------------------------
            Else
                If (e.KeyData = Keys.Space) Then 'пробел
                    RichTextBox1.SelectionColor = Color.Black
                End If
                'RichTextBox1.SelectionColor = Color.Black
            End If

            'Форматирование вводимых операторов
            If (InputOperator = True) Then
                BeginOperator = RichTextBox1.SelectionStart
                InputOperator = False
            End If
            If (e.KeyData = Keys.Enter) Or (e.KeyData = Keys.Up) Or (e.KeyData = Keys.Down) Then
                RichTextBox1.SelectionColor = Color.Black
            End If
            If (e.KeyData = Keys.Space) Or (e.KeyData = Keys.Enter) Or (e.KeyData = Keys.Tab) And (ComentOperator = False) Then
                DepartsModule()
            ElseIf (ComentOperator = True) Then
                ComentOperator = True 'коментарий начался
                RichTextBox1.SelectionColor = Color.Green
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub RichTextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RichTextBox1.KeyPress
        Try
            If (e.KeyChar.GetHashCode = 851981) Then
                Dim nStr, nChar As Integer
                nStr = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart) + 1
                StatusBarPanel1.Text = "Строка : " + nStr.ToString
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub RichTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RichTextBox1.KeyUp
        Try
            If (e.KeyData <> Keys.Delete) Or (e.KeyData <> Keys.Shift) Then
                If (e.KeyValue > 95) And (e.KeyValue < 112) Or (e.KeyValue = 48) Or (e.KeyValue = 49) Or (e.KeyValue = 50) Or (e.KeyValue = 51) Or (e.KeyValue = 52) Or (e.KeyValue = 53) Or (e.KeyValue = 54) Or (e.KeyValue = 55) Or (e.KeyValue = 56) Or (e.KeyValue = 57) Or (e.KeyValue = 189) Or (e.KeyValue = 187) Or (e.KeyValue = 219) Or (e.KeyValue = 221) Or (e.KeyValue = 186) Or (e.KeyValue = 222) Or (e.KeyValue = 220) Or (e.KeyValue = 188) Or (e.KeyValue = 190) Or (e.KeyValue = 191) Then
                    'ВЫДИЛЕНИЕ СПЕЦСИМВОЛОВ ЦВЕТОМ
                    Dim BeginOperator, EndOperator As Integer
                    EndOperator = RichTextBox1.SelectionStart
                    BeginOperator = EndOperator - 1
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = 1

                    RichTextBox1.SelectionColor = Color.Black
                    RichTextBox1.SelectionColor = Color.Black
                    RichTextBox1.SelectionColor = Color.Black
                    If (RichTextBox1.SelectedText = "(") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "!") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "&") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = ")") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "*") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "-") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "=") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "+") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "|") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "{") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = ";") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "'") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "<") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "}") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = """") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = "/") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = ">") Then
                        RichTextBox1.SelectionColor = Color.Blue
                    End If
                    If (RichTextBox1.SelectedText = " ") Then
                        RichTextBox1.SelectionColor = Color.Black
                    End If

                    RichTextBox1.SelectionStart = EndOperator
                    RichTextBox1.SelectionLength = 0
                    RichTextBox1.SelectionColor = Color.Black
                    '---------------------------------------
                End If
            End If
            Dim nStr, nChar As Integer
            nStr = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart) + 1
            StatusBarPanel1.Text = "Строка : " + nStr.ToString
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        Try
            If (e.KeyData = Keys.Escape) Then 'закрытие без выбора
                Panel6.Visible = False
                TextBox2.Text = ""
                ListBox1.SelectedIndex = 0
                RichTextBox1.SelectionStart = RichTextBox1.SelectionStart - 1
                RichTextBox1.SelectionLength = 1
                RichTextBox1.Focus()
            End If
            If (e.KeyData = Keys.Down) Or (e.KeyData = Keys.Up) Or (e.KeyData = Keys.PageUp) Or (e.KeyData = Keys.PageDown) Then 'переход к списку выбора
                ListFocus = True
                ListBox1.Focus()
            End If
            If (e.KeyData = Keys.Enter) Then
                If (ListBox1.SelectedIndex < 0) Then
                    RichTextBox1.SelectionStart = RichTextBox1.SelectionStart - 1
                    RichTextBox1.SelectionLength = 1
                    FormatText(TextBox2.Text)
                    Panel6.Visible = False
                    ListBox1.SelectedIndex = 0
                    TextBox2.Text = ""
                    RichTextBox1.Focus()
                Else
                    RichTextBox1.SelectionStart = RichTextBox1.SelectionStart - 1
                    RichTextBox1.SelectionLength = 1
                    FormatText(ListBox1.Items.Item(ListBox1.SelectedIndex))
                    Panel6.Visible = False
                    ListBox1.SelectedIndex = 0
                    TextBox2.Text = ""
                    RichTextBox1.Focus()
                End If

            End If
        Catch ex As Exception 'В СЛУЧАЕ ОШИБКИ
            RichTextBox1.SelectionStart = RichTextBox1.SelectionStart - 1
            RichTextBox1.SelectionLength = 1
            FormatText(ListBox1.Items.Item(ListBox1.SelectedIndex))
            Panel6.Visible = False
            ListBox1.SelectedIndex = 0
            TextBox2.Text = ""
            RichTextBox1.Focus()
        End Try
    End Sub

    Private Sub TextBox2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyUp
        If (e.KeyCode <> 16 And e.KeyCode <> 17) Then 'игнорировать SHIFT
            Dim Find As String
            Dim CountChar As Integer
            Dim InputOperator As String
            CountChar = TextBox2.TextLength

            For i As Integer = 0 To ListBox1.Items.Count - 1

                Find = ListBox1.Items.Item(i)
                If (Find.Length >= CountChar) Then
                    For j As Integer = 0 To CountChar - 1 'Find.Length - 1
                        InputOperator = InputOperator + Find.Chars(j)
                    Next
                End If

                Me.Update()
                Panel6.Update()
                ListBox1.Update()
                If (TextBox2.Text = InputOperator) Then 'если находит
                    ListBox1.SelectedIndex = i
                    Me.Update()
                    Panel6.Update()
                    ListBox1.Update()
                    Exit For
                Else
                    InputOperator = ""
                End If
            Next
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        'при выборе значения
        If (ListFocus = True) Then
            TextBox2.Text = ListBox1.Items.Item(ListBox1.SelectedIndex)
        End If
    End Sub

    Private Sub ListBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListBox1.KeyDown
        If (e.KeyData = Keys.Escape) Then
            ListFocus = False
            TextBox2.Focus()
        End If
        If (e.KeyData = Keys.Enter) Then
            RichTextBox1.SelectionStart = RichTextBox1.SelectionStart - 1
            RichTextBox1.SelectionLength = 1
            FormatText(ListBox1.Items.Item(ListBox1.SelectedIndex))
            Panel6.Visible = False
            ListBox1.SelectedIndex = 0
            TextBox2.Text = ""
            ListFocus = False
            RichTextBox1.Focus()
        End If
    End Sub

    Public Sub LOAD_OPERATORS()
        Try
            ListBox1.Items.Clear()
            '--------------------------------------------------
            Dim OpenFileList As File
            If OpenFileList.Exists(SomovCompileModule.MyProgramDirectory + "\Resource\operators.txt") = True Then
                RichTextBox2.LoadFile(SomovCompileModule.MyProgramDirectory + "\Resource\operators.txt", RichTextBoxStreamType.PlainText)

                Dim STROKA As String
                For i As Integer = 0 To RichTextBox2.Lines.Length - 1 'колличество строк
                    ListBox1.Items.Add(RichTextBox2.Lines.GetValue(i)) 'получаем строку
                Next
                RichTextBox2.Clear()
            Else
                MsgBox("ОШИБКА: Программа не смогла найти файл " + SomovCompileModule.MyProgramDirectory + "\Resource\operators.txt", MsgBoxStyle.OKOnly, "Сообщение:")
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Public Sub FormatText(ByVal AddText As String)
        RichTextBox1.SelectionColor = Color.Blue
        Clipboard.SetDataObject(AddText.ToString)
        RichTextBox1.Paste()
        RichTextBox1.SelectionColor = Color.Black
    End Sub

    '--ИНТЕРАКТИВНОЕ ВЫДИЛЕНИЕ СЛОВ------------------------------------------------------
    Dim BeginOperator, EndOperator As Integer
    Dim InputOperator As Boolean = True
    Dim ComentOperator As Boolean = False
    Private Sub DepartsModule()
        Try
            EndOperator = RichTextBox1.SelectionStart
            RichTextBox1.SelectionStart = BeginOperator
            Dim LengthOperator As Integer
            LengthOperator = EndOperator - BeginOperator
            If (LengthOperator > 0) Then
                RichTextBox1.SelectionLength = LengthOperator
                If (RichTextBox1.SelectedText = "Возврат") Or (RichTextBox1.SelectedText = "возврат") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Вопрос") Or (RichTextBox1.SelectedText = "вопрос") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Для") Or (RichTextBox1.SelectedText = "для") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Если") Or (RichTextBox1.SelectedText = "если") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Запрос") Or (RichTextBox1.SelectedText = "запрос") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Иначе") Or (RichTextBox1.SelectedText = "иначе") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Перем") Or (RichTextBox1.SelectedText = "перем") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "По") Or (RichTextBox1.SelectedText = "по") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Пока") Or (RichTextBox1.SelectedText = "пока") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Процедура") Or (RichTextBox1.SelectedText = "процедура") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Сообщить") Or (RichTextBox1.SelectedText = "сообщить") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Строка") Or (RichTextBox1.SelectedText = "строка") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Тип") Or (RichTextBox1.SelectedText = "тип") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Тогда") Or (RichTextBox1.SelectedText = "тогда") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Функция") Or (RichTextBox1.SelectedText = "функция") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "Число") Or (RichTextBox1.SelectedText = "число") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "!=") Or (RichTextBox1.SelectedText = "==") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "<=") Or (RichTextBox1.SelectedText = ">=") Then
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Blue
                End If
                If (RichTextBox1.SelectedText = "/*") Then
                    ComentOperator = True
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Green
                End If
                If (RichTextBox1.SelectedText = "*/") Then
                    ComentOperator = False
                    RichTextBox1.SelectionStart = BeginOperator
                    RichTextBox1.SelectionLength = EndOperator - BeginOperator
                    RichTextBox1.SelectionColor = Color.Green
                End If

                RichTextBox1.SelectionStart = EndOperator
                RichTextBox1.SelectionLength = 0
                RichTextBox1.SelectionColor = Color.Black
            Else
                RichTextBox1.SelectionStart = EndOperator
                RichTextBox1.SelectionLength = 0
                RichTextBox1.SelectionColor = Color.Black
            End If
            InputOperator = True
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        If (SomovCompileModule.WErrorShow = True) Then
            SomovCompileModule.WError.Close()
        End If

        CORE.RUN(RichTextBox1, ListBox1)   'Отправляем модуль на выполнение

    End Sub

    Private Sub MenuItem39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem39.Click
        If (SomovCompileModule.WErrorShow = False) Then
            SomovCompileModule.LoadWError()
            SomovCompileModule.WError.Show()
        End If
    End Sub

    Private Sub Compiler_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If (MsgBox("Завершить работу?", MsgBoxStyle.YesNo, "Вопрос:") = MsgBoxResult.Yes) Then
            End
        Else
            e.Cancel = True
        End If

    End Sub

    Private Sub MenuItem40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem40.Click
        'Найти
        If (SomovCompileModule.WSearchShow = False) Then
            SomovCompileModule.LoadWSearch()
            SomovCompileModule.WSearch.Show()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.RichText)
            StatusBarPanel2.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If (StatusBarPanel2.Text = "Модуль не сохранён!") Then
            If (SaveFileDialog1.ShowDialog = DialogResult.OK) Then
                RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText)
                StatusBarPanel2.Text = SaveFileDialog1.FileName
            End If
        Else
            RichTextBox1.SaveFile(StatusBarPanel2.Text, RichTextBoxStreamType.RichText)
            MsgBox("Файл успешно сохранен!", MsgBoxStyle.OKOnly, "Сохранение файла.")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If (SomovCompileModule.WErrorShow = True) Then
            SomovCompileModule.WError.Close()
        End If

        CORE.RUN(RichTextBox1, ListBox1)   'Отправляем модуль на выполнение

    End Sub

    Private Sub MenuItem41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem41.Click
        'Отмена
        RichTextBox2.Undo()
    End Sub

    Private Sub MenuItem43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem43.Click
        'Вырезать
        RichTextBox2.Cut()
    End Sub

    Private Sub MenuItem44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem44.Click
        'Копировать
        RichTextBox2.Copy()
    End Sub

    Private Sub MenuItem45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem45.Click
        'Вставить
        RichTextBox2.Paste()
    End Sub

    Private Sub MenuItem46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem46.Click
        'Удалить
        RichTextBox2.SelectedText = ""
    End Sub

    Private Sub MenuItem48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem48.Click
        'Выделить всё
        RichTextBox2.SelectAll()
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        If (StatusBarPanel2.Text = "Модуль не сохранён!") Then
            If (SaveFileDialog1.ShowDialog = DialogResult.OK) Then
                RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText)
                StatusBarPanel2.Text = SaveFileDialog1.FileName
            End If
        Else
            RichTextBox1.SaveFile(StatusBarPanel2.Text, RichTextBoxStreamType.RichText)
            MsgBox("Файл успешно сохранен!", MsgBoxStyle.OKOnly, "Сохранение файла.")
        End If

    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        RichTextBox1.LoadFile(SomovCompileModule.MyProgramDirectory + "resource\newModule.rtf", RichTextBoxStreamType.RichText)
        StatusBarPanel2.Text = "Модуль не сохранён!"
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        Close()
    End Sub

    Private Sub MenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem37.Click
        If (ColorDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.SelectionColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub MenuItem36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem36.Click
        If (FontDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.SelectionFont = FontDialog1.Font
        End If
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Dim DIRCOPY As File
        Dim DirSTR, DirPATH As String
        Dim FName As String

        FName = InputBox("Введите имя проекта", "Построение проекта", "")

        If (FName <> "") Then
            If (DIRCOPY.Exists(SomovCompileModule.MyProgramDirectory + "resource\ExeSC.exe") = True) Then
                If (FolderBrowserDialog1.ShowDialog = DialogResult.OK) Then
                    DIRCOPY.Copy(SomovCompileModule.MyProgramDirectory + "resource\ExeSC.exe", FolderBrowserDialog1.SelectedPath + "\" + FName + ".exe")
                    RichTextBox1.SaveFile(FolderBrowserDialog1.SelectedPath + "\mcode.scm", RichTextBoxStreamType.RichText)
                End If
            Else
                MsgBox("Не найден файл ExeSC.exe !!!", MsgBoxStyle.OKCancel, "Ошибка:")
            End If
            MsgBox("Проект успешно построен!", MsgBoxStyle.OKOnly, "Сообщение:")
        End If
    End Sub
End Class
