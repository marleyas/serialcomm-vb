<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbPorta = New System.Windows.Forms.ComboBox()
        Me.btnConecta = New System.Windows.Forms.Button()
        Me.rbtReceived = New System.Windows.Forms.RichTextBox()
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Porta:"
        '
        'cmbPorta
        '
        Me.cmbPorta.FormattingEnabled = True
        Me.cmbPorta.Location = New System.Drawing.Point(53, 6)
        Me.cmbPorta.Name = "cmbPorta"
        Me.cmbPorta.Size = New System.Drawing.Size(121, 21)
        Me.cmbPorta.TabIndex = 1
        '
        'btnConecta
        '
        Me.btnConecta.Enabled = False
        Me.btnConecta.Location = New System.Drawing.Point(193, 4)
        Me.btnConecta.Name = "btnConecta"
        Me.btnConecta.Size = New System.Drawing.Size(85, 23)
        Me.btnConecta.TabIndex = 2
        Me.btnConecta.Text = "Conectar"
        Me.btnConecta.UseVisualStyleBackColor = True
        '
        'rbtReceived
        '
        Me.rbtReceived.BackColor = System.Drawing.Color.White
        Me.rbtReceived.Location = New System.Drawing.Point(14, 33)
        Me.rbtReceived.Name = "rbtReceived"
        Me.rbtReceived.ReadOnly = True
        Me.rbtReceived.Size = New System.Drawing.Size(598, 397)
        Me.rbtReceived.TabIndex = 3
        Me.rbtReceived.Text = ""
        '
        'SerialPort1
        '
        Me.SerialPort1.DtrEnable = True
        Me.SerialPort1.RtsEnable = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(624, 442)
        Me.Controls.Add(Me.rbtReceived)
        Me.Controls.Add(Me.btnConecta)
        Me.Controls.Add(Me.cmbPorta)
        Me.Controls.Add(Me.Label1)
        Me.MinimumSize = New System.Drawing.Size(350, 350)
        Me.Name = "frmMain"
        Me.Text = "Comunicação serial com medidor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbPorta As System.Windows.Forms.ComboBox
    Friend WithEvents btnConecta As System.Windows.Forms.Button
    Friend WithEvents rbtReceived As System.Windows.Forms.RichTextBox
    Friend WithEvents SerialPort1 As System.IO.Ports.SerialPort

End Class
