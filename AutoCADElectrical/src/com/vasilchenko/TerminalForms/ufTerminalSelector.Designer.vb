<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ufTerminalSelector
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.cbxLocation = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxTerminal = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbxOrientation = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbxDuctSide = New System.Windows.Forms.ComboBox()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(266, 224)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(90, 25)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'cbxLocation
        '
        Me.cbxLocation.FormattingEnabled = True
        Me.cbxLocation.Location = New System.Drawing.Point(186, 57)
        Me.cbxLocation.Name = "cbxLocation"
        Me.cbxLocation.Size = New System.Drawing.Size(170, 21)
        Me.cbxLocation.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Выберите местонахождение"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 100)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Выберите клеммник"
        '
        'cbxTerminal
        '
        Me.cbxTerminal.FormattingEnabled = True
        Me.cbxTerminal.Location = New System.Drawing.Point(186, 97)
        Me.cbxTerminal.Name = "cbxTerminal"
        Me.cbxTerminal.Size = New System.Drawing.Size(170, 21)
        Me.cbxTerminal.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 140)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(121, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Выберите ориентацию"
        '
        'cbxOrientation
        '
        Me.cbxOrientation.FormattingEnabled = True
        Me.cbxOrientation.Location = New System.Drawing.Point(186, 137)
        Me.cbxOrientation.Name = "cbxOrientation"
        Me.cbxOrientation.Size = New System.Drawing.Size(170, 21)
        Me.cbxOrientation.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 180)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(147, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Выберите сторону к коробу"
        '
        'cbxDuctSide
        '
        Me.cbxDuctSide.FormattingEnabled = True
        Me.cbxDuctSide.Location = New System.Drawing.Point(186, 177)
        Me.cbxDuctSide.Name = "cbxDuctSide"
        Me.cbxDuctSide.Size = New System.Drawing.Size(170, 21)
        Me.cbxDuctSide.TabIndex = 7
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(170, 224)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(90, 25)
        Me.btnApply.TabIndex = 9
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label5.Location = New System.Drawing.Point(12, 20)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(351, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Укажите данные клеммника, который хотите отобразить"
        '
        'ufTerminalSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonShadow
        Me.ClientSize = New System.Drawing.Size(368, 256)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbxDuctSide)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbxOrientation)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbxTerminal)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbxLocation)
        Me.Controls.Add(Me.btnCancel)
        Me.Name = "ufTerminalSelector"
        Me.Text = "Terminal Selector"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents cbxLocation As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents cbxTerminal As Windows.Forms.ComboBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents cbxOrientation As Windows.Forms.ComboBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents cbxDuctSide As Windows.Forms.ComboBox
    Friend WithEvents btnApply As Windows.Forms.Button
    Friend WithEvents Label5 As Windows.Forms.Label
End Class
