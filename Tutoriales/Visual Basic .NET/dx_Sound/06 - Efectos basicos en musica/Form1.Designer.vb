<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
#Region "Código generado por el Diseñador de Windows Forms "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'Llamada necesaria para el Diseñador de Windows Forms.
		InitializeComponent()
	End Sub
	'Form invalida a Dispose para limpiar la lista de componentes.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Requerido por el Diseñador de Windows Forms
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents HScroll3 As System.Windows.Forms.HScrollBar
	Public WithEvents HScroll2 As System.Windows.Forms.HScrollBar
	Public WithEvents HScroll1 As System.Windows.Forms.HScrollBar
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
	'Se puede modificar mediante el Diseñador de Windows Forms.
	'No lo modifique con el editor de código.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.HScroll3 = New System.Windows.Forms.HScrollBar
		Me.HScroll2 = New System.Windows.Forms.HScrollBar
		Me.HScroll1 = New System.Windows.Forms.HScrollBar
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "dx_Sound - Efectos basicos en musica"
		Me.ClientSize = New System.Drawing.Size(304, 204)
		Me.Location = New System.Drawing.Point(3, 23)
		Me.Icon = CType(resources.GetObject("Form1.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form1"
		Me.HScroll3.Size = New System.Drawing.Size(165, 21)
		Me.HScroll3.LargeChange = 50
		Me.HScroll3.Location = New System.Drawing.Point(104, 116)
		Me.HScroll3.Maximum = 249
		Me.HScroll3.Minimum = 25
		Me.HScroll3.SmallChange = 5
		Me.HScroll3.TabIndex = 5
		Me.HScroll3.Value = 100
		Me.HScroll3.CausesValidation = True
		Me.HScroll3.Enabled = True
		Me.HScroll3.Cursor = System.Windows.Forms.Cursors.Default
		Me.HScroll3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HScroll3.TabStop = True
		Me.HScroll3.Visible = True
		Me.HScroll3.Name = "HScroll3"
		Me.HScroll2.Size = New System.Drawing.Size(165, 21)
		Me.HScroll2.LargeChange = 10
		Me.HScroll2.Location = New System.Drawing.Point(104, 92)
		Me.HScroll2.Maximum = 109
		Me.HScroll2.Minimum = -100
		Me.HScroll2.TabIndex = 3
		Me.HScroll2.CausesValidation = True
		Me.HScroll2.Enabled = True
		Me.HScroll2.Cursor = System.Windows.Forms.Cursors.Default
		Me.HScroll2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HScroll2.SmallChange = 1
		Me.HScroll2.TabStop = True
		Me.HScroll2.Value = 0
		Me.HScroll2.Visible = True
		Me.HScroll2.Name = "HScroll2"
		Me.HScroll1.Size = New System.Drawing.Size(165, 21)
		Me.HScroll1.LargeChange = 5
		Me.HScroll1.Location = New System.Drawing.Point(104, 68)
		Me.HScroll1.Maximum = 104
		Me.HScroll1.TabIndex = 1
		Me.HScroll1.Value = 100
		Me.HScroll1.CausesValidation = True
		Me.HScroll1.Enabled = True
		Me.HScroll1.Minimum = 0
		Me.HScroll1.Cursor = System.Windows.Forms.Cursors.Default
		Me.HScroll1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HScroll1.SmallChange = 1
		Me.HScroll1.TabStop = True
		Me.HScroll1.Visible = True
		Me.HScroll1.Name = "HScroll1"
		Me.Label3.Text = "Velocidad"
		Me.Label3.Size = New System.Drawing.Size(61, 17)
		Me.Label3.Location = New System.Drawing.Point(36, 120)
		Me.Label3.TabIndex = 4
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.Text = "Balance"
		Me.Label2.Size = New System.Drawing.Size(61, 17)
		Me.Label2.Location = New System.Drawing.Point(36, 96)
		Me.Label2.TabIndex = 2
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "Volumen"
		Me.Label1.Size = New System.Drawing.Size(61, 17)
		Me.Label1.Location = New System.Drawing.Point(36, 72)
		Me.Label1.TabIndex = 0
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(HScroll3)
		Me.Controls.Add(HScroll2)
		Me.Controls.Add(HScroll1)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class