Public Class WindowERROR
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
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(WindowERROR))
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.Location = New System.Drawing.Point(0, 0)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.Size = New System.Drawing.Size(448, 95)
        Me.ListBox1.TabIndex = 0
        '
        'WindowERROR
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(448, 96)
        Me.Controls.Add(Me.ListBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WindowERROR"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Окно ошибок:"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub WindowERROR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SomovCompileModule.WErrorShow = True
    End Sub

    Private Sub WindowERROR_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        SomovCompileModule.WErrorShow = False
    End Sub
End Class
