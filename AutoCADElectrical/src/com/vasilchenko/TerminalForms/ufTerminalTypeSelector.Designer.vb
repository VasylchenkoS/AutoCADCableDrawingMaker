<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ufTerminalTypeSelector
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
        Me.rbtnSignalisation = New System.Windows.Forms.RadioButton()
        Me.rbtnMeasurement = New System.Windows.Forms.RadioButton()
        Me.rbtnControl = New System.Windows.Forms.RadioButton()
        Me.rbtnPower = New System.Windows.Forms.RadioButton()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'rbtnSignalisation
        '
        Me.rbtnSignalisation.AutoSize = True
        Me.rbtnSignalisation.Location = New System.Drawing.Point(67, 35)
        Me.rbtnSignalisation.Name = "rbtnSignalisation"
        Me.rbtnSignalisation.Size = New System.Drawing.Size(121, 17)
        Me.rbtnSignalisation.TabIndex = 0
        Me.rbtnSignalisation.TabStop = True
        Me.rbtnSignalisation.Text = "Телесигнализация"
        Me.rbtnSignalisation.UseVisualStyleBackColor = True
        '
        'rbtnMeasurement
        '
        Me.rbtnMeasurement.AutoSize = True
        Me.rbtnMeasurement.Location = New System.Drawing.Point(67, 60)
        Me.rbtnMeasurement.Name = "rbtnMeasurement"
        Me.rbtnMeasurement.Size = New System.Drawing.Size(106, 17)
        Me.rbtnMeasurement.TabIndex = 1
        Me.rbtnMeasurement.TabStop = True
        Me.rbtnMeasurement.Text = "Телеизмерение"
        Me.rbtnMeasurement.UseVisualStyleBackColor = True
        '
        'rbtnControl
        '
        Me.rbtnControl.AutoSize = True
        Me.rbtnControl.Location = New System.Drawing.Point(67, 85)
        Me.rbtnControl.Name = "rbtnControl"
        Me.rbtnControl.Size = New System.Drawing.Size(109, 17)
        Me.rbtnControl.TabIndex = 2
        Me.rbtnControl.TabStop = True
        Me.rbtnControl.Text = "Телеуправление"
        Me.rbtnControl.UseVisualStyleBackColor = True
        '
        'rbtnPower
        '
        Me.rbtnPower.AutoSize = True
        Me.rbtnPower.Location = New System.Drawing.Point(67, 110)
        Me.rbtnPower.Name = "rbtnPower"
        Me.rbtnPower.Size = New System.Drawing.Size(68, 17)
        Me.rbtnPower.TabIndex = 3
        Me.rbtnPower.TabStop = True
        Me.rbtnPower.Text = "Питание"
        Me.rbtnPower.UseVisualStyleBackColor = True
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(98, 143)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(75, 23)
        Me.btnApply.TabIndex = 4
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(179, 143)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(203, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Выберие назначение клеммника"
        '
        'ufTerminalTypeSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(267, 178)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.rbtnPower)
        Me.Controls.Add(Me.rbtnControl)
        Me.Controls.Add(Me.rbtnMeasurement)
        Me.Controls.Add(Me.rbtnSignalisation)
        Me.Name = "ufTerminalTypeSelector"
        Me.Text = "Terminal Type Selector"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents rbtnSignalisation As Windows.Forms.RadioButton
    Friend WithEvents rbtnMeasurement As Windows.Forms.RadioButton
    Friend WithEvents rbtnControl As Windows.Forms.RadioButton
    Friend WithEvents rbtnPower As Windows.Forms.RadioButton
    Friend WithEvents btnApply As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
End Class
