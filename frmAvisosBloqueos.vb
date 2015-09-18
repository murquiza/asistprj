Public Class frmAvisosBloqueos
    Inherits System.Windows.Forms.Form

#Region " Código generado por el Diseñador de Windows Forms "

    Public Sub New()
        MyBase.New()

        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()

        'Agregar cualquier inicialización después de la llamada a InitializeComponent()

    End Sub

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms requiere el siguiente procedimiento
    'Puede modificarse utilizando el Diseñador de Windows Forms. 
    'No lo modifique con el editor de código.
    Friend WithEvents txtTexto As System.Windows.Forms.Label
    Friend WithEvents cbVolver As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAvisosBloqueos))
        Me.txtTexto = New System.Windows.Forms.Label
        Me.cbVolver = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtTexto
        '
        Me.txtTexto.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtTexto.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTexto.ForeColor = System.Drawing.Color.DodgerBlue
        Me.txtTexto.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.txtTexto.Location = New System.Drawing.Point(8, 8)
        Me.txtTexto.Name = "txtTexto"
        Me.txtTexto.Size = New System.Drawing.Size(384, 208)
        Me.txtTexto.TabIndex = 1
        Me.txtTexto.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cbVolver
        '
        Me.cbVolver.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cbVolver.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbVolver.Image = CType(resources.GetObject("cbVolver.Image"), System.Drawing.Image)
        Me.cbVolver.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbVolver.Location = New System.Drawing.Point(336, 232)
        Me.cbVolver.Name = "cbVolver"
        Me.cbVolver.Size = New System.Drawing.Size(56, 56)
        Me.cbVolver.TabIndex = 2
        Me.cbVolver.Text = "Volver"
        Me.cbVolver.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'frmAvisosBloqueos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 16)
        Me.BackColor = System.Drawing.Color.LemonChiffon
        Me.CancelButton = Me.cbVolver
        Me.ClientSize = New System.Drawing.Size(400, 298)
        Me.Controls.Add(Me.cbVolver)
        Me.Controls.Add(Me.txtTexto)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmAvisosBloqueos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detalle de Bloqueo Referencia"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmAvisosBloqueos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = Me.Text & " " & frmInstCierres.lvwCierres.SelectedItems(0).Text
        txtTexto.Text = colAvisosBloqueo.Item(frmInstCierres.lvwCierres.SelectedItems(0).Text)
    End Sub

    Private Sub cbVolver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVolver.Click
        Me.Hide()
    End Sub
End Class
