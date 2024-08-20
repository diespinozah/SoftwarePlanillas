<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.BtnComparar = New System.Windows.Forms.Button()
        Me.TxtArchivos = New System.Windows.Forms.RichTextBox()
        Me.BtnSeleccionarArchivos = New System.Windows.Forms.Button()
        Me.BtnLimpiarCeldas = New System.Windows.Forms.Button()
        Me.lblValidando = New System.Windows.Forms.Label()
        Me.PBComparar = New System.Windows.Forms.ProgressBar()
        Me.lblLimpiar = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BtnComparar
        '
        Me.BtnComparar.Location = New System.Drawing.Point(134, 34)
        Me.BtnComparar.Name = "BtnComparar"
        Me.BtnComparar.Size = New System.Drawing.Size(119, 26)
        Me.BtnComparar.TabIndex = 10
        Me.BtnComparar.Text = "Comparar"
        Me.BtnComparar.UseVisualStyleBackColor = True
        '
        'TxtArchivos
        '
        Me.TxtArchivos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtArchivos.Location = New System.Drawing.Point(12, 85)
        Me.TxtArchivos.Name = "TxtArchivos"
        Me.TxtArchivos.ReadOnly = True
        Me.TxtArchivos.Size = New System.Drawing.Size(1029, 410)
        Me.TxtArchivos.TabIndex = 11
        Me.TxtArchivos.Text = ""
        '
        'BtnSeleccionarArchivos
        '
        Me.BtnSeleccionarArchivos.Location = New System.Drawing.Point(12, 34)
        Me.BtnSeleccionarArchivos.Name = "BtnSeleccionarArchivos"
        Me.BtnSeleccionarArchivos.Size = New System.Drawing.Size(116, 26)
        Me.BtnSeleccionarArchivos.TabIndex = 12
        Me.BtnSeleccionarArchivos.Text = "Seleccionar"
        Me.BtnSeleccionarArchivos.UseVisualStyleBackColor = True
        '
        'BtnLimpiarCeldas
        '
        Me.BtnLimpiarCeldas.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnLimpiarCeldas.Location = New System.Drawing.Point(922, 34)
        Me.BtnLimpiarCeldas.Name = "BtnLimpiarCeldas"
        Me.BtnLimpiarCeldas.Size = New System.Drawing.Size(119, 26)
        Me.BtnLimpiarCeldas.TabIndex = 13
        Me.BtnLimpiarCeldas.Text = "Limpiar"
        Me.BtnLimpiarCeldas.UseVisualStyleBackColor = True
        '
        'lblValidando
        '
        Me.lblValidando.AutoSize = True
        Me.lblValidando.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblValidando.Location = New System.Drawing.Point(319, 36)
        Me.lblValidando.Name = "lblValidando"
        Me.lblValidando.Size = New System.Drawing.Size(91, 20)
        Me.lblValidando.TabIndex = 18
        Me.lblValidando.Text = "Validando"
        Me.lblValidando.Visible = False
        '
        'PBComparar
        '
        Me.PBComparar.Location = New System.Drawing.Point(434, 36)
        Me.PBComparar.Name = "PBComparar"
        Me.PBComparar.Size = New System.Drawing.Size(290, 23)
        Me.PBComparar.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.PBComparar.TabIndex = 19
        Me.PBComparar.Value = 100
        Me.PBComparar.Visible = False
        '
        'lblLimpiar
        '
        Me.lblLimpiar.AutoSize = True
        Me.lblLimpiar.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLimpiar.Location = New System.Drawing.Point(319, 34)
        Me.lblLimpiar.Name = "lblLimpiar"
        Me.lblLimpiar.Size = New System.Drawing.Size(95, 20)
        Me.lblLimpiar.TabIndex = 21
        Me.lblLimpiar.Text = "Limpiando"
        Me.lblLimpiar.Visible = False
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1053, 507)
        Me.Controls.Add(Me.lblLimpiar)
        Me.Controls.Add(Me.PBComparar)
        Me.Controls.Add(Me.lblValidando)
        Me.Controls.Add(Me.BtnLimpiarCeldas)
        Me.Controls.Add(Me.BtnSeleccionarArchivos)
        Me.Controls.Add(Me.BtnComparar)
        Me.Controls.Add(Me.TxtArchivos)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ValidadorPlanillas"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnComparar As Button
    Friend WithEvents TxtArchivos As RichTextBox
    Friend WithEvents BtnSeleccionarArchivos As Button
    Friend WithEvents BtnLimpiarCeldas As Button
    Friend WithEvents lblValidando As Label
    Friend WithEvents PBComparar As ProgressBar
    Friend WithEvents lblLimpiar As Label
End Class
