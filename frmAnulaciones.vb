Public Class frmAnulaciones
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
    Friend WithEvents txtReferencia As System.Windows.Forms.Label
    Friend WithEvents lblReferencia As System.Windows.Forms.Label
    Friend WithEvents lblDescripcion As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbVolver As System.Windows.Forms.Button
    Friend WithEvents txtDescripcion As System.Windows.Forms.Label
    Friend WithEvents txtComentarios As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAnulaciones))
        Me.txtReferencia = New System.Windows.Forms.Label
        Me.lblReferencia = New System.Windows.Forms.Label
        Me.lblDescripcion = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDescripcion = New System.Windows.Forms.Label
        Me.txtComentarios = New System.Windows.Forms.Label
        Me.cbVolver = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtReferencia
        '
        Me.txtReferencia.BackColor = System.Drawing.Color.Transparent
        Me.txtReferencia.Font = New System.Drawing.Font("Tahoma", 9.75!)
        Me.txtReferencia.ForeColor = System.Drawing.Color.RoyalBlue
        Me.txtReferencia.Location = New System.Drawing.Point(112, 8)
        Me.txtReferencia.Name = "txtReferencia"
        Me.txtReferencia.Size = New System.Drawing.Size(96, 23)
        Me.txtReferencia.TabIndex = 0
        '
        'lblReferencia
        '
        Me.lblReferencia.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReferencia.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblReferencia.Location = New System.Drawing.Point(8, 8)
        Me.lblReferencia.Name = "lblReferencia"
        Me.lblReferencia.Size = New System.Drawing.Size(80, 23)
        Me.lblReferencia.TabIndex = 1
        Me.lblReferencia.Text = "Referencia:"
        '
        'lblDescripcion
        '
        Me.lblDescripcion.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescripcion.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblDescripcion.Location = New System.Drawing.Point(8, 32)
        Me.lblDescripcion.Name = "lblDescripcion"
        Me.lblDescripcion.Size = New System.Drawing.Size(88, 23)
        Me.lblDescripcion.TabIndex = 2
        Me.lblDescripcion.Text = "Descripción:"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label2.Location = New System.Drawing.Point(8, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 23)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Comentarios:"
        '
        'txtDescripcion
        '
        Me.txtDescripcion.BackColor = System.Drawing.Color.Transparent
        Me.txtDescripcion.Font = New System.Drawing.Font("Tahoma", 9.75!)
        Me.txtDescripcion.ForeColor = System.Drawing.Color.RoyalBlue
        Me.txtDescripcion.Location = New System.Drawing.Point(112, 32)
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.Size = New System.Drawing.Size(424, 23)
        Me.txtDescripcion.TabIndex = 4
        '
        'txtComentarios
        '
        Me.txtComentarios.BackColor = System.Drawing.Color.Transparent
        Me.txtComentarios.Font = New System.Drawing.Font("Tahoma", 9.75!)
        Me.txtComentarios.ForeColor = System.Drawing.Color.RoyalBlue
        Me.txtComentarios.Location = New System.Drawing.Point(112, 64)
        Me.txtComentarios.Name = "txtComentarios"
        Me.txtComentarios.Size = New System.Drawing.Size(424, 168)
        Me.txtComentarios.TabIndex = 5
        '
        'cbVolver
        '
        Me.cbVolver.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbVolver.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbVolver.Image = CType(resources.GetObject("cbVolver.Image"), System.Drawing.Image)
        Me.cbVolver.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbVolver.Location = New System.Drawing.Point(552, 176)
        Me.cbVolver.Name = "cbVolver"
        Me.cbVolver.Size = New System.Drawing.Size(56, 56)
        Me.cbVolver.TabIndex = 6
        Me.cbVolver.Text = "Volver"
        Me.cbVolver.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'frmAnulaciones
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(618, 240)
        Me.Controls.Add(Me.cbVolver)
        Me.Controls.Add(Me.txtComentarios)
        Me.Controls.Add(Me.txtDescripcion)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblDescripcion)
        Me.Controls.Add(Me.lblReferencia)
        Me.Controls.Add(Me.txtReferencia)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmAnulaciones"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Comentarios y Descripciónes de Anulación"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmAnulaciones_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rsLocal As New ADODB.Recordset
        Dim sql As String

        txtReferencia.Text = frmInstCierres.lvwAnulaciones.SelectedItems(0).SubItems(frmInstCierres.T2_REFER.Index).Text

        sql = "Select   T5_Descripcion, T5_Comentarios " & _
              "From     AnulacionesAsistencia " & _
              "Where    T5_Refer = '" & frmInstCierres.lvwAnulaciones.SelectedItems(0).SubItems(frmInstCierres.T2_REFER.Index).Text & _
              "' and    T5_Codcia = '" & strCodCia & "'"

        rsLocal.Open(sql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsLocal.EOF Then
            txtDescripcion.Text = rsLocal.Fields("T5_Descripcion").Value
            txtComentarios.Text = rsLocal.Fields("T5_Comentarios").Value
        End If

        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()

    End Sub

    Private Sub cbVolver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVolver.Click
        Me.Hide()
    End Sub
End Class
