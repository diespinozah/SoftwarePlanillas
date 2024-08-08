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
        Me.BtnComparar = New System.Windows.Forms.Button()
        Me.TxtArchivos = New System.Windows.Forms.RichTextBox()
        Me.BtnSeleccionarArchivos = New System.Windows.Forms.Button()
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
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1053, 507)
        Me.Controls.Add(Me.BtnSeleccionarArchivos)
        Me.Controls.Add(Me.BtnComparar)
        Me.Controls.Add(Me.TxtArchivos)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ValidadorPlanillas"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnComparar As Button
    Friend WithEvents TxtArchivos As RichTextBox
    Friend WithEvents BtnSeleccionarArchivos As Button
End Class
