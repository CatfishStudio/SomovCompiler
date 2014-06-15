Public Class load
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(load))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Transparent
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(352, 245)
        Me.Panel1.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(112, 224)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(280, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "2013 © Сомов Евгений. Все права защищены."
        '
        'Timer1
        '
        Me.Timer1.Interval = 2000
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
        'load
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(352, 246)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "load"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CATFISH STUDIO"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub load_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Update()
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        SomovCompileModule.LoadSCompile()
        'Открыть с помощью моей программы
        If (Environment.GetCommandLineArgs.Length > 1) Then
            SomovCompileModule.FileOpenMyProgram = Environment.GetCommandLineArgs(1)
            SomovCompileModule.SCompile.RichTextBox1.LoadFile(SomovCompileModule.FileOpenMyProgram, RichTextBoxStreamType.RichText)
            SomovCompileModule.SCompile.StatusBarPanel2.Text = Environment.GetCommandLineArgs(1)
        Else
            SomovCompileModule.MyProgramDirectory = Environment.CurrentDirectory + "\"
            Me.Update()
            SomovCompileModule.PathBase = SomovCompileModule.MyProgramDirectory + "resource\help.mdb"
            SomovCompileModule.StringConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SomovCompileModule.PathBase & ";Jet OLEDB:Database Password="
            '----------------------------------------------
            OleDbConnection1.ConnectionString = SomovCompileModule.StringConnection
            Try
                OleDbConnection1.Open()
            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
                End
            End Try
            OleDbConnection1.Close()
            '----------------------------------------------
        End If
        Me.Visible = False
        SomovCompileModule.SCompile.Show()
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class
