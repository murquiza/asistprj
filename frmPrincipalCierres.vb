Public Class frmPrincipalCierres
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cbBuscar As System.Windows.Forms.Button
    Friend WithEvents cbxBusquedaAvanzada As System.Windows.Forms.CheckBox
    Friend WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbxCompania As System.Windows.Forms.ListBox
    Friend WithEvents lbxProducto As System.Windows.Forms.ListBox
    Friend WithEvents FiltroErrores As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroAviso As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroNoPagados As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroPagados As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroTodos As System.Windows.Forms.RadioButton
    Friend WithEvents gbxFiltro As System.Windows.Forms.GroupBox
    Friend WithEvents ttipAyuda As System.Windows.Forms.ToolTip
    Friend WithEvents cbSalir As System.Windows.Forms.Button
    Friend WithEvents cbAvisos As System.Windows.Forms.Button
    Friend WithEvents prbProgreso As System.Windows.Forms.ProgressBar
    'Friend WithEvents CR2 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents sbPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents cbxCompania As System.Windows.Forms.ComboBox
    Friend WithEvents cbxProducto As System.Windows.Forms.ComboBox
    Friend WithEvents lbCompaniaAsistencia As System.Windows.Forms.Label
    Friend WithEvents T2_CODSIN As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_REFER As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_ESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ANU_T2_CODSIN As System.Windows.Forms.ColumnHeader
    Friend WithEvents ANU_T2_REFER As System.Windows.Forms.ColumnHeader
    Friend WithEvents ANU_FEC_ANUL As System.Windows.Forms.ColumnHeader
    Friend WithEvents ANU_COD As System.Windows.Forms.ColumnHeader
    Friend WithEvents chkFiltroAvisos As System.Windows.Forms.CheckBox
    Friend WithEvents cbTodos As System.Windows.Forms.Button
    Friend WithEvents cbNinguno As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cbxTipoFecha As System.Windows.Forms.ComboBox
    Friend WithEvents chkUltimoPago As System.Windows.Forms.CheckBox
    Friend WithEvents lvwCierres As System.Windows.Forms.ListView
    Friend WithEvents rbSiniestros As System.Windows.Forms.RadioButton
    Friend WithEvents rbAnulaciones As System.Windows.Forms.RadioButton
    Friend WithEvents dtpFechaCierre As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblSiniestro As System.Windows.Forms.Label
    Friend WithEvents txtSiniestro As System.Windows.Forms.TextBox
    Friend WithEvents lblReferencia As System.Windows.Forms.Label
    Friend WithEvents txtReferencia As System.Windows.Forms.TextBox
    Friend WithEvents lblFechaAnulaciones As System.Windows.Forms.Label
    Friend WithEvents rbAnuProvisionales As System.Windows.Forms.RadioButton
    Friend WithEvents rbAnuConfirmados As System.Windows.Forms.RadioButton
    Friend WithEvents lblFechaCierre As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFechaAnulaciones As System.Windows.Forms.DateTimePicker
    Friend WithEvents lvwAnulaciones As System.Windows.Forms.ListView
    Friend WithEvents T2_FPAGO As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECGRA As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_FESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents SITUACION As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblTipoFecha As System.Windows.Forms.Label
    Friend WithEvents lblFechaHasta As System.Windows.Forms.Label
    Friend WithEvents lblFechaDesde As System.Windows.Forms.Label
    Friend WithEvents lblProducto As System.Windows.Forms.Label
    Friend WithEvents cbCerrar As System.Windows.Forms.Button
    Friend WithEvents cbPendientes As System.Windows.Forms.Button
    Friend WithEvents stbEstado As System.Windows.Forms.StatusBar
    Friend WithEvents cbImprimir As System.Windows.Forms.Button
    Friend WithEvents picTest As System.Windows.Forms.PictureBox
    Public WithEvents CR2 As AxCrystal.AxCrystalReport
    Public WithEvents CR3 As AxCrystal.AxCrystalReport
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrincipalCierres))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lbCompaniaAsistencia = New System.Windows.Forms.Label
        Me.cbxCompania = New System.Windows.Forms.ComboBox
        Me.gbxFiltro = New System.Windows.Forms.GroupBox
        Me.chkUltimoPago = New System.Windows.Forms.CheckBox
        Me.FiltroErrores = New System.Windows.Forms.RadioButton
        Me.FiltroAviso = New System.Windows.Forms.RadioButton
        Me.FiltroNoPagados = New System.Windows.Forms.RadioButton
        Me.FiltroPagados = New System.Windows.Forms.RadioButton
        Me.FiltroTodos = New System.Windows.Forms.RadioButton
        Me.chkFiltroAvisos = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpFechaAnulaciones = New System.Windows.Forms.DateTimePicker
        Me.rbAnuConfirmados = New System.Windows.Forms.RadioButton
        Me.rbAnuProvisionales = New System.Windows.Forms.RadioButton
        Me.lblFechaAnulaciones = New System.Windows.Forms.Label
        Me.dtpFechaCierre = New System.Windows.Forms.DateTimePicker
        Me.lblFechaCierre = New System.Windows.Forms.Label
        Me.cbxTipoFecha = New System.Windows.Forms.ComboBox
        Me.lblTipoFecha = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblSiniestro = New System.Windows.Forms.Label
        Me.txtSiniestro = New System.Windows.Forms.TextBox
        Me.lblReferencia = New System.Windows.Forms.Label
        Me.txtReferencia = New System.Windows.Forms.TextBox
        Me.cbxBusquedaAvanzada = New System.Windows.Forms.CheckBox
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker
        Me.cbxProducto = New System.Windows.Forms.ComboBox
        Me.lblFechaHasta = New System.Windows.Forms.Label
        Me.lblFechaDesde = New System.Windows.Forms.Label
        Me.lblProducto = New System.Windows.Forms.Label
        Me.rbSiniestros = New System.Windows.Forms.RadioButton
        Me.rbAnulaciones = New System.Windows.Forms.RadioButton
        Me.cbBuscar = New System.Windows.Forms.Button
        Me.lvwCierres = New System.Windows.Forms.ListView
        Me.T2_CODSIN = New System.Windows.Forms.ColumnHeader
        Me.T2_REFER = New System.Windows.Forms.ColumnHeader
        Me.T2_FPAGO = New System.Windows.Forms.ColumnHeader
        Me.FECGRA = New System.Windows.Forms.ColumnHeader
        Me.T2_ESTADO = New System.Windows.Forms.ColumnHeader
        Me.T2_FESTADO = New System.Windows.Forms.ColumnHeader
        Me.SITUACION = New System.Windows.Forms.ColumnHeader
        Me.stbEstado = New System.Windows.Forms.StatusBar
        Me.sbPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.cbCerrar = New System.Windows.Forms.Button
        Me.cbAvisos = New System.Windows.Forms.Button
        Me.cbImprimir = New System.Windows.Forms.Button
        Me.cbSalir = New System.Windows.Forms.Button
        Me.lbxCompania = New System.Windows.Forms.ListBox
        Me.lbxProducto = New System.Windows.Forms.ListBox
        Me.ttipAyuda = New System.Windows.Forms.ToolTip(Me.components)
        Me.cbTodos = New System.Windows.Forms.Button
        Me.cbNinguno = New System.Windows.Forms.Button
        Me.cbPendientes = New System.Windows.Forms.Button
        Me.picTest = New System.Windows.Forms.PictureBox
        Me.prbProgreso = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lvwAnulaciones = New System.Windows.Forms.ListView
        Me.ANU_T2_CODSIN = New System.Windows.Forms.ColumnHeader
        Me.ANU_T2_REFER = New System.Windows.Forms.ColumnHeader
        Me.ANU_FEC_ANUL = New System.Windows.Forms.ColumnHeader
        Me.ANU_COD = New System.Windows.Forms.ColumnHeader
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.CR2 = New AxCrystal.AxCrystalReport
        Me.CR3 = New AxCrystal.AxCrystalReport
        Me.gbxFiltro.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.CR2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CR3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(48, 48)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'lbCompaniaAsistencia
        '
        Me.lbCompaniaAsistencia.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbCompaniaAsistencia.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lbCompaniaAsistencia.Location = New System.Drawing.Point(64, 32)
        Me.lbCompaniaAsistencia.Name = "lbCompaniaAsistencia"
        Me.lbCompaniaAsistencia.Size = New System.Drawing.Size(152, 24)
        Me.lbCompaniaAsistencia.TabIndex = 1
        Me.lbCompaniaAsistencia.Text = "Compañía Asistencia:"
        Me.lbCompaniaAsistencia.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxCompania
        '
        Me.cbxCompania.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxCompania.Location = New System.Drawing.Point(232, 32)
        Me.cbxCompania.Name = "cbxCompania"
        Me.cbxCompania.Size = New System.Drawing.Size(336, 24)
        Me.cbxCompania.TabIndex = 2
        '
        'gbxFiltro
        '
        Me.gbxFiltro.Controls.Add(Me.chkUltimoPago)
        Me.gbxFiltro.Controls.Add(Me.FiltroErrores)
        Me.gbxFiltro.Controls.Add(Me.FiltroAviso)
        Me.gbxFiltro.Controls.Add(Me.FiltroNoPagados)
        Me.gbxFiltro.Controls.Add(Me.FiltroPagados)
        Me.gbxFiltro.Controls.Add(Me.FiltroTodos)
        Me.gbxFiltro.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxFiltro.ForeColor = System.Drawing.Color.RoyalBlue
        Me.gbxFiltro.Location = New System.Drawing.Point(728, 16)
        Me.gbxFiltro.Name = "gbxFiltro"
        Me.gbxFiltro.Size = New System.Drawing.Size(136, 176)
        Me.gbxFiltro.TabIndex = 3
        Me.gbxFiltro.TabStop = False
        Me.gbxFiltro.Text = "Filtro"
        Me.ttipAyuda.SetToolTip(Me.gbxFiltro, "Filtros sobre el resultado de la consulta")
        '
        'chkUltimoPago
        '
        Me.chkUltimoPago.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUltimoPago.ForeColor = System.Drawing.Color.DodgerBlue
        Me.chkUltimoPago.Location = New System.Drawing.Point(16, 136)
        Me.chkUltimoPago.Name = "chkUltimoPago"
        Me.chkUltimoPago.TabIndex = 5
        Me.chkUltimoPago.Text = "Último Pago"
        '
        'FiltroErrores
        '
        Me.FiltroErrores.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroErrores.ForeColor = System.Drawing.Color.Red
        Me.FiltroErrores.Location = New System.Drawing.Point(16, 112)
        Me.FiltroErrores.Name = "FiltroErrores"
        Me.FiltroErrores.TabIndex = 4
        Me.FiltroErrores.Text = "Errores ( E )"
        Me.ttipAyuda.SetToolTip(Me.FiltroErrores, "Muestra sólo los registros en los que se ha producido error")
        '
        'FiltroAviso
        '
        Me.FiltroAviso.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroAviso.ForeColor = System.Drawing.Color.DarkGoldenrod
        Me.FiltroAviso.Location = New System.Drawing.Point(16, 88)
        Me.FiltroAviso.Name = "FiltroAviso"
        Me.FiltroAviso.TabIndex = 3
        Me.FiltroAviso.Text = "Aviso ( A )"
        Me.ttipAyuda.SetToolTip(Me.FiltroAviso, "Muestra sólo los registros con mensaje de aviso")
        '
        'FiltroNoPagados
        '
        Me.FiltroNoPagados.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroNoPagados.ForeColor = System.Drawing.Color.Green
        Me.FiltroNoPagados.Location = New System.Drawing.Point(16, 64)
        Me.FiltroNoPagados.Name = "FiltroNoPagados"
        Me.FiltroNoPagados.Size = New System.Drawing.Size(112, 24)
        Me.FiltroNoPagados.TabIndex = 2
        Me.FiltroNoPagados.Text = "No Pagados ( X )"
        Me.ttipAyuda.SetToolTip(Me.FiltroNoPagados, "Muestra sólo los pendientes de aperturar")
        '
        'FiltroPagados
        '
        Me.FiltroPagados.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroPagados.ForeColor = System.Drawing.Color.Gray
        Me.FiltroPagados.Location = New System.Drawing.Point(16, 40)
        Me.FiltroPagados.Name = "FiltroPagados"
        Me.FiltroPagados.TabIndex = 1
        Me.FiltroPagados.Text = "Pagados ( P )"
        Me.ttipAyuda.SetToolTip(Me.FiltroPagados, "Muestra sólo los aperturados")
        '
        'FiltroTodos
        '
        Me.FiltroTodos.Checked = True
        Me.FiltroTodos.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroTodos.ForeColor = System.Drawing.Color.Black
        Me.FiltroTodos.Location = New System.Drawing.Point(16, 16)
        Me.FiltroTodos.Name = "FiltroTodos"
        Me.FiltroTodos.TabIndex = 0
        Me.FiltroTodos.TabStop = True
        Me.FiltroTodos.Text = "Todos"
        Me.ttipAyuda.SetToolTip(Me.FiltroTodos, "Muestra todos los registros")
        '
        'chkFiltroAvisos
        '
        Me.chkFiltroAvisos.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFiltroAvisos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.chkFiltroAvisos.Location = New System.Drawing.Point(736, 192)
        Me.chkFiltroAvisos.Name = "chkFiltroAvisos"
        Me.chkFiltroAvisos.Size = New System.Drawing.Size(128, 40)
        Me.chkFiltroAvisos.TabIndex = 6
        Me.chkFiltroAvisos.Text = "Desactivar filtros para avisos"
        Me.ttipAyuda.SetToolTip(Me.chkFiltroAvisos, "Desactiva los Filtros para procesar los Avisos")
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Controls.Add(Me.dtpFechaCierre)
        Me.GroupBox2.Controls.Add(Me.lblFechaCierre)
        Me.GroupBox2.Controls.Add(Me.cbxTipoFecha)
        Me.GroupBox2.Controls.Add(Me.lblTipoFecha)
        Me.GroupBox2.Controls.Add(Me.Panel1)
        Me.GroupBox2.Controls.Add(Me.dtpHasta)
        Me.GroupBox2.Controls.Add(Me.dtpDesde)
        Me.GroupBox2.Controls.Add(Me.cbxProducto)
        Me.GroupBox2.Controls.Add(Me.lblFechaHasta)
        Me.GroupBox2.Controls.Add(Me.lblFechaDesde)
        Me.GroupBox2.Controls.Add(Me.lblProducto)
        Me.GroupBox2.Controls.Add(Me.rbSiniestros)
        Me.GroupBox2.Controls.Add(Me.rbAnulaciones)
        Me.GroupBox2.Controls.Add(Me.cbBuscar)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox2.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(712, 168)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Criterios de Selección"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpFechaAnulaciones)
        Me.GroupBox1.Controls.Add(Me.rbAnuConfirmados)
        Me.GroupBox1.Controls.Add(Me.rbAnuProvisionales)
        Me.GroupBox1.Controls.Add(Me.lblFechaAnulaciones)
        Me.GroupBox1.Location = New System.Drawing.Point(504, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(128, 128)
        Me.GroupBox1.TabIndex = 17
        Me.GroupBox1.TabStop = False
        '
        'dtpFechaAnulaciones
        '
        Me.dtpFechaAnulaciones.Enabled = False
        Me.dtpFechaAnulaciones.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpFechaAnulaciones.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFechaAnulaciones.Location = New System.Drawing.Point(16, 88)
        Me.dtpFechaAnulaciones.Name = "dtpFechaAnulaciones"
        Me.dtpFechaAnulaciones.Size = New System.Drawing.Size(96, 21)
        Me.dtpFechaAnulaciones.TabIndex = 13
        '
        'rbAnuConfirmados
        '
        Me.rbAnuConfirmados.Checked = True
        Me.rbAnuConfirmados.Enabled = False
        Me.rbAnuConfirmados.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.rbAnuConfirmados.ForeColor = System.Drawing.Color.Firebrick
        Me.rbAnuConfirmados.Location = New System.Drawing.Point(24, 16)
        Me.rbAnuConfirmados.Name = "rbAnuConfirmados"
        Me.rbAnuConfirmados.Size = New System.Drawing.Size(88, 24)
        Me.rbAnuConfirmados.TabIndex = 2
        Me.rbAnuConfirmados.TabStop = True
        Me.rbAnuConfirmados.Text = "Confirmados"
        '
        'rbAnuProvisionales
        '
        Me.rbAnuProvisionales.Enabled = False
        Me.rbAnuProvisionales.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.rbAnuProvisionales.ForeColor = System.Drawing.Color.Firebrick
        Me.rbAnuProvisionales.Location = New System.Drawing.Point(24, 40)
        Me.rbAnuProvisionales.Name = "rbAnuProvisionales"
        Me.rbAnuProvisionales.Size = New System.Drawing.Size(88, 24)
        Me.rbAnuProvisionales.TabIndex = 3
        Me.rbAnuProvisionales.Text = "Provisionales"
        '
        'lblFechaAnulaciones
        '
        Me.lblFechaAnulaciones.Enabled = False
        Me.lblFechaAnulaciones.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblFechaAnulaciones.ForeColor = System.Drawing.Color.Firebrick
        Me.lblFechaAnulaciones.Location = New System.Drawing.Point(8, 72)
        Me.lblFechaAnulaciones.Name = "lblFechaAnulaciones"
        Me.lblFechaAnulaciones.Size = New System.Drawing.Size(104, 16)
        Me.lblFechaAnulaciones.TabIndex = 12
        Me.lblFechaAnulaciones.Text = "Fecha Anulaciones:"
        '
        'dtpFechaCierre
        '
        Me.dtpFechaCierre.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpFechaCierre.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFechaCierre.Location = New System.Drawing.Point(96, 134)
        Me.dtpFechaCierre.Name = "dtpFechaCierre"
        Me.dtpFechaCierre.Size = New System.Drawing.Size(104, 21)
        Me.dtpFechaCierre.TabIndex = 16
        Me.ttipAyuda.SetToolTip(Me.dtpFechaCierre, "Abre el calendario de selección")
        '
        'lblFechaCierre
        '
        Me.lblFechaCierre.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblFechaCierre.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblFechaCierre.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblFechaCierre.Location = New System.Drawing.Point(16, 136)
        Me.lblFechaCierre.Name = "lblFechaCierre"
        Me.lblFechaCierre.Size = New System.Drawing.Size(72, 17)
        Me.lblFechaCierre.TabIndex = 15
        Me.lblFechaCierre.Text = "Fecha cierre:"
        Me.lblFechaCierre.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxTipoFecha
        '
        Me.cbxTipoFecha.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbxTipoFecha.Location = New System.Drawing.Point(96, 64)
        Me.cbxTipoFecha.Name = "cbxTipoFecha"
        Me.cbxTipoFecha.Size = New System.Drawing.Size(216, 21)
        Me.cbxTipoFecha.TabIndex = 13
        '
        'lblTipoFecha
        '
        Me.lblTipoFecha.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblTipoFecha.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblTipoFecha.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTipoFecha.Location = New System.Drawing.Point(24, 64)
        Me.lblTipoFecha.Name = "lblTipoFecha"
        Me.lblTipoFecha.Size = New System.Drawing.Size(64, 18)
        Me.lblTipoFecha.TabIndex = 12
        Me.lblTipoFecha.Text = "Tipo fecha:"
        Me.lblTipoFecha.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightGray
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.lblSiniestro)
        Me.Panel1.Controls.Add(Me.txtSiniestro)
        Me.Panel1.Controls.Add(Me.lblReferencia)
        Me.Panel1.Controls.Add(Me.txtReferencia)
        Me.Panel1.Controls.Add(Me.cbxBusquedaAvanzada)
        Me.Panel1.Location = New System.Drawing.Point(320, 40)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(176, 120)
        Me.Panel1.TabIndex = 10
        '
        'lblSiniestro
        '
        Me.lblSiniestro.Enabled = False
        Me.lblSiniestro.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblSiniestro.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblSiniestro.Location = New System.Drawing.Point(24, 74)
        Me.lblSiniestro.Name = "lblSiniestro"
        Me.lblSiniestro.Size = New System.Drawing.Size(56, 17)
        Me.lblSiniestro.TabIndex = 2
        Me.lblSiniestro.Text = "Siniestro:"
        Me.lblSiniestro.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtSiniestro
        '
        Me.txtSiniestro.Enabled = False
        Me.txtSiniestro.Location = New System.Drawing.Point(80, 72)
        Me.txtSiniestro.Name = "txtSiniestro"
        Me.txtSiniestro.Size = New System.Drawing.Size(88, 21)
        Me.txtSiniestro.TabIndex = 4
        Me.txtSiniestro.Text = ""
        '
        'lblReferencia
        '
        Me.lblReferencia.Enabled = False
        Me.lblReferencia.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblReferencia.Location = New System.Drawing.Point(8, 48)
        Me.lblReferencia.Name = "lblReferencia"
        Me.lblReferencia.Size = New System.Drawing.Size(72, 16)
        Me.lblReferencia.TabIndex = 1
        Me.lblReferencia.Text = "Referencia:"
        Me.lblReferencia.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtReferencia
        '
        Me.txtReferencia.Enabled = False
        Me.txtReferencia.Location = New System.Drawing.Point(80, 46)
        Me.txtReferencia.Name = "txtReferencia"
        Me.txtReferencia.Size = New System.Drawing.Size(88, 21)
        Me.txtReferencia.TabIndex = 3
        Me.txtReferencia.Text = ""
        '
        'cbxBusquedaAvanzada
        '
        Me.cbxBusquedaAvanzada.Location = New System.Drawing.Point(8, 6)
        Me.cbxBusquedaAvanzada.Name = "cbxBusquedaAvanzada"
        Me.cbxBusquedaAvanzada.Size = New System.Drawing.Size(136, 24)
        Me.cbxBusquedaAvanzada.TabIndex = 0
        Me.cbxBusquedaAvanzada.Text = "Búsqueda avanzada"
        '
        'dtpHasta
        '
        Me.dtpHasta.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpHasta.Location = New System.Drawing.Point(96, 110)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(104, 21)
        Me.dtpHasta.TabIndex = 7
        Me.ttipAyuda.SetToolTip(Me.dtpHasta, "Abre el calendario de selección")
        '
        'dtpDesde
        '
        Me.dtpDesde.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDesde.Location = New System.Drawing.Point(96, 88)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(104, 21)
        Me.dtpDesde.TabIndex = 6
        Me.ttipAyuda.SetToolTip(Me.dtpDesde, "Abre el calendario de selección")
        '
        'cbxProducto
        '
        Me.cbxProducto.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbxProducto.Location = New System.Drawing.Point(96, 40)
        Me.cbxProducto.Name = "cbxProducto"
        Me.cbxProducto.Size = New System.Drawing.Size(216, 21)
        Me.cbxProducto.TabIndex = 5
        '
        'lblFechaHasta
        '
        Me.lblFechaHasta.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblFechaHasta.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblFechaHasta.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblFechaHasta.Location = New System.Drawing.Point(16, 112)
        Me.lblFechaHasta.Name = "lblFechaHasta"
        Me.lblFechaHasta.Size = New System.Drawing.Size(72, 17)
        Me.lblFechaHasta.TabIndex = 2
        Me.lblFechaHasta.Text = "Fecha hasta:"
        Me.lblFechaHasta.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFechaDesde
        '
        Me.lblFechaDesde.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblFechaDesde.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblFechaDesde.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblFechaDesde.Location = New System.Drawing.Point(16, 90)
        Me.lblFechaDesde.Name = "lblFechaDesde"
        Me.lblFechaDesde.Size = New System.Drawing.Size(72, 16)
        Me.lblFechaDesde.TabIndex = 1
        Me.lblFechaDesde.Text = "Fecha desde:"
        Me.lblFechaDesde.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblProducto
        '
        Me.lblProducto.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lblProducto.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblProducto.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblProducto.Location = New System.Drawing.Point(32, 40)
        Me.lblProducto.Name = "lblProducto"
        Me.lblProducto.Size = New System.Drawing.Size(56, 16)
        Me.lblProducto.TabIndex = 0
        Me.lblProducto.Text = "Producto:"
        Me.lblProducto.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'rbSiniestros
        '
        Me.rbSiniestros.Checked = True
        Me.rbSiniestros.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbSiniestros.ForeColor = System.Drawing.Color.DodgerBlue
        Me.rbSiniestros.Location = New System.Drawing.Point(24, 16)
        Me.rbSiniestros.Name = "rbSiniestros"
        Me.rbSiniestros.Size = New System.Drawing.Size(112, 24)
        Me.rbSiniestros.TabIndex = 0
        Me.rbSiniestros.TabStop = True
        Me.rbSiniestros.Text = "Siniestros"
        '
        'rbAnulaciones
        '
        Me.rbAnulaciones.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbAnulaciones.ForeColor = System.Drawing.Color.Firebrick
        Me.rbAnulaciones.Location = New System.Drawing.Point(504, 16)
        Me.rbAnulaciones.Name = "rbAnulaciones"
        Me.rbAnulaciones.TabIndex = 1
        Me.rbAnulaciones.Text = "Anulaciones"
        '
        'cbBuscar
        '
        Me.cbBuscar.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbBuscar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBuscar.Image = CType(resources.GetObject("cbBuscar.Image"), System.Drawing.Image)
        Me.cbBuscar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbBuscar.Location = New System.Drawing.Point(640, 104)
        Me.cbBuscar.Name = "cbBuscar"
        Me.cbBuscar.Size = New System.Drawing.Size(64, 56)
        Me.cbBuscar.TabIndex = 11
        Me.cbBuscar.Text = "Buscar"
        Me.cbBuscar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbBuscar, "Ejecuta la busqueda de Pagos según los criterios asignados")
        '
        'lvwCierres
        '
        Me.lvwCierres.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvwCierres.CheckBoxes = True
        Me.lvwCierres.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.T2_CODSIN, Me.T2_REFER, Me.T2_FPAGO, Me.FECGRA, Me.T2_ESTADO, Me.T2_FESTADO, Me.SITUACION})
        Me.lvwCierres.FullRowSelect = True
        Me.lvwCierres.GridLines = True
        Me.lvwCierres.HoverSelection = True
        Me.lvwCierres.Location = New System.Drawing.Point(8, 256)
        Me.lvwCierres.MultiSelect = False
        Me.lvwCierres.Name = "lvwCierres"
        Me.lvwCierres.Size = New System.Drawing.Size(536, 304)
        Me.lvwCierres.TabIndex = 5
        Me.lvwCierres.View = System.Windows.Forms.View.Details
        '
        'T2_CODSIN
        '
        Me.T2_CODSIN.Text = "Siniestro"
        Me.T2_CODSIN.Width = 83
        '
        'T2_REFER
        '
        Me.T2_REFER.Text = "Referencia"
        Me.T2_REFER.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_REFER.Width = 106
        '
        'T2_FPAGO
        '
        Me.T2_FPAGO.Text = "Ramo"
        Me.T2_FPAGO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_FPAGO.Width = 50
        '
        'FECGRA
        '
        Me.FECGRA.Text = "Póliza"
        Me.FECGRA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.FECGRA.Width = 70
        '
        'T2_ESTADO
        '
        Me.T2_ESTADO.Text = "Estado"
        Me.T2_ESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_ESTADO.Width = 54
        '
        'T2_FESTADO
        '
        Me.T2_FESTADO.Text = "Situación"
        Me.T2_FESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_FESTADO.Width = 69
        '
        'SITUACION
        '
        Me.SITUACION.Text = "Cierre"
        Me.SITUACION.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.SITUACION.Width = 83
        '
        'stbEstado
        '
        Me.stbEstado.Location = New System.Drawing.Point(0, 644)
        Me.stbEstado.Name = "stbEstado"
        Me.stbEstado.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbPanel1, Me.sbPanel2, Me.sbPanel3})
        Me.stbEstado.ShowPanels = True
        Me.stbEstado.Size = New System.Drawing.Size(880, 22)
        Me.stbEstado.TabIndex = 6
        '
        'sbPanel1
        '
        Me.sbPanel1.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.sbPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.sbPanel1.Icon = CType(resources.GetObject("sbPanel1.Icon"), System.Drawing.Icon)
        Me.sbPanel1.Width = 31
        '
        'sbPanel2
        '
        Me.sbPanel2.Width = 420
        '
        'sbPanel3
        '
        Me.sbPanel3.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.sbPanel3.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.sbPanel3.Width = 413
        '
        'cbCerrar
        '
        Me.cbCerrar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbCerrar.Image = CType(resources.GetObject("cbCerrar.Image"), System.Drawing.Image)
        Me.cbCerrar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbCerrar.Location = New System.Drawing.Point(608, 576)
        Me.cbCerrar.Name = "cbCerrar"
        Me.cbCerrar.Size = New System.Drawing.Size(64, 56)
        Me.cbCerrar.TabIndex = 10
        Me.cbCerrar.Text = "Cerrar"
        Me.cbCerrar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbCerrar, "Ejecución Proceso de Pagos")
        '
        'cbAvisos
        '
        Me.cbAvisos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbAvisos.Image = CType(resources.GetObject("cbAvisos.Image"), System.Drawing.Image)
        Me.cbAvisos.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbAvisos.Location = New System.Drawing.Point(736, 576)
        Me.cbAvisos.Name = "cbAvisos"
        Me.cbAvisos.Size = New System.Drawing.Size(64, 56)
        Me.cbAvisos.TabIndex = 11
        Me.cbAvisos.Text = "Avisos"
        Me.cbAvisos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbAvisos, "Visualizar Histórico de Errores/Avisos")
        '
        'cbImprimir
        '
        Me.cbImprimir.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbImprimir.Image = CType(resources.GetObject("cbImprimir.Image"), System.Drawing.Image)
        Me.cbImprimir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbImprimir.Location = New System.Drawing.Point(672, 576)
        Me.cbImprimir.Name = "cbImprimir"
        Me.cbImprimir.Size = New System.Drawing.Size(64, 56)
        Me.cbImprimir.TabIndex = 12
        Me.cbImprimir.Text = "Imprimir"
        Me.cbImprimir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbImprimir, "Imprime la selección de pantalla")
        '
        'cbSalir
        '
        Me.cbSalir.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbSalir.Image = CType(resources.GetObject("cbSalir.Image"), System.Drawing.Image)
        Me.cbSalir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbSalir.Location = New System.Drawing.Point(800, 576)
        Me.cbSalir.Name = "cbSalir"
        Me.cbSalir.Size = New System.Drawing.Size(64, 56)
        Me.cbSalir.TabIndex = 13
        Me.cbSalir.Text = "Salir"
        Me.cbSalir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbSalir, "Cierra el Gestor de Pagos de Asistencia")
        '
        'lbxCompania
        '
        Me.lbxCompania.Location = New System.Drawing.Point(272, 576)
        Me.lbxCompania.Name = "lbxCompania"
        Me.lbxCompania.Size = New System.Drawing.Size(80, 30)
        Me.lbxCompania.TabIndex = 14
        Me.lbxCompania.Visible = False
        '
        'lbxProducto
        '
        Me.lbxProducto.Location = New System.Drawing.Point(576, 8)
        Me.lbxProducto.Name = "lbxProducto"
        Me.lbxProducto.Size = New System.Drawing.Size(88, 17)
        Me.lbxProducto.TabIndex = 15
        Me.lbxProducto.Visible = False
        '
        'cbTodos
        '
        Me.cbTodos.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbTodos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbTodos.Image = CType(resources.GetObject("cbTodos.Image"), System.Drawing.Image)
        Me.cbTodos.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbTodos.Location = New System.Drawing.Point(8, 16)
        Me.cbTodos.Name = "cbTodos"
        Me.cbTodos.Size = New System.Drawing.Size(64, 56)
        Me.cbTodos.TabIndex = 8
        Me.cbTodos.Tag = "TODOS"
        Me.cbTodos.Text = "Todos"
        Me.cbTodos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbTodos, "Selecciona todos los pagos de la consulta")
        '
        'cbNinguno
        '
        Me.cbNinguno.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbNinguno.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbNinguno.Image = CType(resources.GetObject("cbNinguno.Image"), System.Drawing.Image)
        Me.cbNinguno.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbNinguno.Location = New System.Drawing.Point(80, 16)
        Me.cbNinguno.Name = "cbNinguno"
        Me.cbNinguno.Size = New System.Drawing.Size(64, 56)
        Me.cbNinguno.TabIndex = 8
        Me.cbNinguno.Tag = "NINGUNO"
        Me.cbNinguno.Text = "Ninguno"
        Me.cbNinguno.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbNinguno, "Cancela la selección")
        '
        'cbPendientes
        '
        Me.cbPendientes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbPendientes.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbPendientes.Image = CType(resources.GetObject("cbPendientes.Image"), System.Drawing.Image)
        Me.cbPendientes.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbPendientes.Location = New System.Drawing.Point(152, 16)
        Me.cbPendientes.Name = "cbPendientes"
        Me.cbPendientes.Size = New System.Drawing.Size(72, 56)
        Me.cbPendientes.TabIndex = 9
        Me.cbPendientes.Tag = "PENDIENTES"
        Me.cbPendientes.Text = "Pendientes"
        Me.cbPendientes.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbPendientes, "Cancela la selección")
        '
        'picTest
        '
        Me.picTest.Cursor = System.Windows.Forms.Cursors.Help
        Me.picTest.Image = CType(resources.GetObject("picTest.Image"), System.Drawing.Image)
        Me.picTest.Location = New System.Drawing.Point(8, 8)
        Me.picTest.Name = "picTest"
        Me.picTest.Size = New System.Drawing.Size(48, 48)
        Me.picTest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picTest.TabIndex = 63
        Me.picTest.TabStop = False
        Me.ttipAyuda.SetToolTip(Me.picTest, "La aplicación se esta ejecutando en modo pruebas")
        Me.picTest.Visible = False
        '
        'prbProgreso
        '
        Me.prbProgreso.Location = New System.Drawing.Point(40, 653)
        Me.prbProgreso.Name = "prbProgreso"
        Me.prbProgreso.Size = New System.Drawing.Size(402, 8)
        Me.prbProgreso.TabIndex = 16
        Me.prbProgreso.Visible = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BlanchedAlmond
        Me.Label1.Location = New System.Drawing.Point(8, 240)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(536, 16)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Siniestros"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Firebrick
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(552, 240)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(312, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Anluaciones de Siniestros"
        '
        'lvwAnulaciones
        '
        Me.lvwAnulaciones.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvwAnulaciones.CheckBoxes = True
        Me.lvwAnulaciones.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ANU_T2_CODSIN, Me.ANU_T2_REFER, Me.ANU_FEC_ANUL, Me.ANU_COD})
        Me.lvwAnulaciones.ForeColor = System.Drawing.Color.Black
        Me.lvwAnulaciones.FullRowSelect = True
        Me.lvwAnulaciones.GridLines = True
        Me.lvwAnulaciones.Location = New System.Drawing.Point(552, 256)
        Me.lvwAnulaciones.Name = "lvwAnulaciones"
        Me.lvwAnulaciones.Size = New System.Drawing.Size(312, 304)
        Me.lvwAnulaciones.TabIndex = 19
        Me.lvwAnulaciones.View = System.Windows.Forms.View.Details
        '
        'ANU_T2_CODSIN
        '
        Me.ANU_T2_CODSIN.Text = "Siniestro"
        Me.ANU_T2_CODSIN.Width = 76
        '
        'ANU_T2_REFER
        '
        Me.ANU_T2_REFER.Text = "Referencia"
        Me.ANU_T2_REFER.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ANU_T2_REFER.Width = 68
        '
        'ANU_FEC_ANUL
        '
        Me.ANU_FEC_ANUL.Text = "F.Anul."
        Me.ANU_FEC_ANUL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ANU_FEC_ANUL.Width = 74
        '
        'ANU_COD
        '
        Me.ANU_COD.Text = "Cod"
        Me.ANU_COD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ANU_COD.Width = 63
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cbPendientes)
        Me.GroupBox3.Controls.Add(Me.cbTodos)
        Me.GroupBox3.Controls.Add(Me.cbNinguno)
        Me.GroupBox3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox3.Location = New System.Drawing.Point(8, 560)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(232, 80)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Selección de pagos"
        '
        'CR2
        '
        Me.CR2.Enabled = True
        Me.CR2.Location = New System.Drawing.Point(416, 584)
        Me.CR2.Name = "CR2"
        Me.CR2.OcxState = CType(resources.GetObject("CR2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CR2.Size = New System.Drawing.Size(28, 28)
        Me.CR2.TabIndex = 64
        '
        'CR3
        '
        Me.CR3.Enabled = True
        Me.CR3.Location = New System.Drawing.Point(472, 584)
        Me.CR3.Name = "CR3"
        Me.CR3.OcxState = CType(resources.GetObject("CR3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CR3.Size = New System.Drawing.Size(28, 28)
        Me.CR3.TabIndex = 65
        '
        'frmPrincipalCierres
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(880, 666)
        Me.Controls.Add(Me.CR3)
        Me.Controls.Add(Me.CR2)
        Me.Controls.Add(Me.lvwAnulaciones)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.prbProgreso)
        Me.Controls.Add(Me.lbxProducto)
        Me.Controls.Add(Me.lbxCompania)
        Me.Controls.Add(Me.cbSalir)
        Me.Controls.Add(Me.cbImprimir)
        Me.Controls.Add(Me.cbAvisos)
        Me.Controls.Add(Me.cbCerrar)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.stbEstado)
        Me.Controls.Add(Me.lvwCierres)
        Me.Controls.Add(Me.gbxFiltro)
        Me.Controls.Add(Me.cbxCompania)
        Me.Controls.Add(Me.lbCompaniaAsistencia)
        Me.Controls.Add(Me.chkFiltroAvisos)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.picTest)
        Me.Controls.Add(Me.GroupBox2)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrincipalCierres"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Siniestros: Área de Asistencia  -  Cierres Automáticos"
        Me.gbxFiltro.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.CR2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CR3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub frmPrincipalAperturas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InicioApp()
    End Sub

    Private Sub InicioApp()

        On Error GoTo InicioApp_Error

        Dim strParametro As String

        HoraTopeEjecucion = "21:15:00"

        strIdProceso = "C"

        colAvisosBloqueo = New Collection
        colSiniestrosCerrados = New Collection
        colSiniestrosPendientes = New Collection

        frmInstCierres = Me

        strParametro = Microsoft.VisualBasic.Command
        If strParametro = "PM" Then
            claseBDCierres.ConnexionPruebas()
            claseBDCierres.BDComand.CommandTimeout = 300
            picTest.Show()
            picTest.BringToFront()
        Else
            claseBDCierres.BDWorkConnect.CommandTimeout = 0
            claseBDCierres.BDComand.CommandTimeout = 0
            picTest.Hide()
            picTest.SendToBack()
        End If

        ' Valores globales
        '
        'CodUserApli = objUtiles.CodUser(UsuaApli)   ' Usuario de la aplicación
        If strCodUserApli = "" Then strCodUserApli = strParametro


        'PathReports = clses.GetParam("PathReports") ' Ubicación de los ficheros de impresión
        PathReports = "K:\Reports\"

        ' Creación de objetos  
        PathIconos = "K:\Graficos\"

        ' Identificador de componente
        strIDComp = "PAG"

        ' Inicialización ComboBox de Productos
        '
        If Not LlenarComboProducto(cbxProducto, lbxProducto, claseBDCierres, "TODOS") Then
            Err.Raise(Val(strGlobalNumErr))
        Else
            'JCLopez_i
            cbxProducto.SelectedIndex = cbxProducto.Items.Count - 1
            'cbxProducto.AddItem("Todos los Productos", 0)
            'cbxProducto.Text = cbxProducto.List(0)
            'JCLopez_f

        End If

        ' Inicialización ComboBox de Compañías
        '
        If LlenarComboCias(cbxCompania, lbxCompania) Then
            cbxCompania.SelectedIndex = 0
        Else
            Err.Raise(Val(strGlobalNumErr))
        End If

        ' Asignación de valores iniciales
        '
        cbxTipoFecha.Items.Add("Importación Pagos")
        cbxTipoFecha.Items.Add("Pago Mutua")
        cbxTipoFecha.Items.Add("Ninguna")
        cbxTipoFecha.Text = "Importación Pagos"

        dtpDesde.Value = Today
        dtpHasta.Value = Today
        dtpFechaCierre.Value = Today



        ' Asignación de valores iniciales
        '
        dtpDesde.Value = Today
        dtpHasta.Value = Today

        bwflag = False

        strFiltro = "T"

        FiltroTodos.PerformClick()
        Exit Sub
InicioApp_Error:

        MsgBox("Ha ocurrido un error iniciando la aplicación", MsgBoxStyle.Exclamation)

    End Sub

    Private Sub FiltroTodos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FiltroTodos.Click, FiltroAviso.Click, FiltroErrores.Click, FiltroNoPagados.Click, FiltroPagados.Click
        Dim rbBoton As RadioButton
        Dim lb_ret As Boolean

        rbBoton = sender
        lb_ret = FiltrarRegistros(rbBoton.TabIndex, cbxCompania)

        If bwflag And lb_ret Then
            RefrescarGrid(dtpDesde.Value, dtpHasta.Value, cbxProducto.Text, lbxCompania.Items.Item(cbxCompania.SelectedIndex))
        End If
    End Sub

    Private Sub cbBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBuscar.Click
        If rbSiniestros.Checked Then
            If cbxBusquedaAvanzada.Checked Then
                BusquedaAvanzada()
            Else
                If cbxCompania.Text = "" Then
                    MsgBox("Debe escoger Producto y/o Compañía de Asistencia", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If

                BusquedaSimple()
            End If
        Else
            If ValidacionFecha((dtpFechaAnulaciones.Value)) Then
                CargaAnulaciones()
            End If
        End If

    End Sub
    Public Function ValidacionFecha(ByRef Fecha As DateTime) As Boolean

        ValidacionFecha = True

        If Not IsDate(Fecha) Then
            MsgBox("La fecha introducida no es correcta")
            ValidacionFecha = False
        End If

    End Function
    Private Sub BusquedaSimple()
        On Error GoTo BusquedaSimple_Error

        Dim dtFechaDesde, dtFechaHasta As Date

        dtFechaDesde = dtpDesde.Value
        dtFechaHasta = dtpHasta.Value

        strSigCompa = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))
        strCodProducto = cbxProducto.Text

        If strCodProducto <> "Todos los Productos" Then
            strCodProducto = lbxProducto.Items.Item(cbxProducto.SelectedIndex)
        End If

        RefrescarGrid(dtFechaDesde, dtFechaHasta, strCodProducto, strSigCompa)
        Exit Sub
BusquedaSimple_Error:
        MsgBox("Ha ocurrido un error en la búsqueda", MsgBoxStyle.Critical)
    End Sub

    Private Sub BusquedaAvanzada()
        On Error GoTo BusquedaAvanzada_Error

        If txtSiniestro.Text = "" And txtReferencia.Text = "" Then
            MsgBox("No se ha indicado ningún siniestro / referencia", MsgBoxStyle.Exclamation)
        ElseIf txtSiniestro.Text <> "" Then
            strCampoBuscaAvanzada = "T2_Codsin"
            strValorBuscaAvanzada = Trim(txtSiniestro.Text)
        Else
            strCampoBuscaAvanzada = "T2_Refer"
            strValorBuscaAvanzada = Trim(txtReferencia.Text)
        End If
        Call RefrescarGrid(dtpDesde.Value, dtpHasta.Value, cbxProducto.Text, strSigCompa)
        Exit Sub

BusquedaAvanzada_Error:
        MsgBox("Ha ocurrido un error realizando la búsqueda avanzada.", MsgBoxStyle.Critical)
    End Sub

    Private Sub cbxCompania_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxCompania.SelectedIndexChanged

        strCodcia = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))
        'Label2.Text = strCodCia

        If Not DatosCiaAsistencia(strCodcia) Then
            MsgBox("No hay datos", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Public Sub RefrescarGrid(ByRef FecDes As Date, ByRef FecHas As Date, ByRef produc As String, ByRef compa As String)

        On Error GoTo RefrescaGrid_Error

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql entera
        Dim FecFiltro As Date ' Fecha a filtar
        Dim objlistitem As ListViewItem ' Objeto con los registros del grid
        Dim objListSubItem As ListViewItem.ListViewSubItem ' Objeto con las columnas del grid
        Dim strOrderBy As String ' Parte Order By de la Sql

        ' Reinicializamos la colección de avisos de bloqueo
        '
        colAvisosBloqueo = New Collection
        objCierres = New clsCierres_NET

        Cursor.Current = Cursors.WaitCursor

        ' Establecemos la selección de los campos con los que vamos a
        ' trabajar de la tabla de pagos de asistencia
        strSQLSel = "SELECT DISTINCT Compa = '" & strCodcia & "', " & " Snsinies.Codsin SINIESTRO, " & " Snsinies.Refext REFERENCIA, " & _
                                              " Snsinies.Codram RAMO, Snsinies.Numpol POLIZA, Angel_t1.EstadoProcCierre ESTADO, Snsinies.Estado SITUACION, " & _
                                              " 'Cierre' CIERRE"

        ' Si estamos realizando una busqueda avanzada ( por Siniestro o Referencia )
        ' no se tienen en cuenta los filtros ni ningún otro criterio de busqueda
        '
        If cbxBusquedaAvanzada.CheckState = 1 Then
            strFromMas = ""
            strWhereMas = ""
            strFrom = " From Angel_T2, Angel_T1, Snsinies "
            strWhere = " Where Angel_T2." & strCampoBuscaAvanzada & " = '" & Trim(strValorBuscaAvanzada) & "' " & _
                       " and Angel_T2.T2_Codsin = Snsinies.Codsin " & _
                       " and Angel_T1.T1_Codsin = Snsinies.Codsin " & _
                       " and Angel_t1.T1_Codcia = '" & strCodcia & "' " & _
                       " and Left(Snsinies.Refext,4) = 'AS" & strIdReferCompa & "'"
        Else
            ' Añadimos el From de la Sql
            '
            strFrom = " From Angel_T1, Snsinies "

            ' Añadimos la Where de la Sql
            '

            '/*MUL T-19908 INI
            'If strCodcia = "I" Then
            '    strWhere = " Where (Angel_T1.T1_Codsin = Snsinies.Codsin) and Angel_T1.T1_Codcia ='" & strCodcia & "' and " & "       Left(Snsinies.Refext,4) = 'AS" & strIdReferCompa & "' and Angel_T1.T1_Estado = 'P'"
            'ElseIf strCodcia = "R" Then
            '    strWhere = " Where (Angel_T1.T1_Codsin = Snsinies.Codsin) and Angel_T1.T1_Codcia ='" & strCodcia & "' and " & "       Left(Snsinies.Refext,2) = '" & strIdReferCompa & "' and Angel_T1.T1_Estado = 'P'"
            'End If
            Select Case strCodcia
                Case "I", "M", "E"
                    strWhere = " Where (Angel_T1.T1_Codsin = Snsinies.Codsin) and Angel_T1.T1_Codcia ='" & strCodcia & "' and " & " Left(Snsinies.Refext,4) = 'AS" & strIdReferCompa & "' and Angel_T1.T1_Estado = 'P'"
                Case "R"
                    strWhere = " Where (Angel_T1.T1_Codsin = Snsinies.Codsin) and Angel_T1.T1_Codcia ='" & strCodcia & "' and " & " Left(Snsinies.Refext,2) = '" & strIdReferCompa & "' and Angel_T1.T1_Estado = 'P'"
                Case Else
            End Select
            '/*MUL T-19908 FIN

            ' Añadimos la parte de la Where que filtrará los registros en
            ' función del tipo de fecha que hayamos seleccionado
            '
            Select Case cbxTipoFecha.Text

                Case "Ninguna"
                    strFromMas = ""
                    strWhereMas = ""

                Case "Importación Pagos" ' Fecha en la que se importo el pago
                    strFromMas = " , Angel_T2 "
                    strWhereMas = " And Angel_T2.T2_Fgraba BETWEEN '" & claseUtilidadesCierres.FormatoFechaSQL(FecDes, False, False) & "' AND '" & claseUtilidadesCierres.FormatoFechaSQL(FecHas, False, False) & "' and Angel_T2.T2_Codsin = Angel_T1.T1_Codsin "

                Case "Pago Mutua" ' Fecha en la que Mutua realiza el pago
                    strFromMas = " , Angel_T2 "
                    strWhereMas = " And Angel_T2.T2_Festado BETWEEN '" & claseUtilidadesCierres.FormatoFechaSQL(FecDes, False, False) & "' AND '" & claseUtilidadesCierres.FormatoFechaSQL(FecHas, False, False) & "' and Angel_T2.T2_Codsin = Angel_T1.T1_Codsin "
            End Select

            ' Añadimos la parte de la Where que filtrará los registros en
            ' función del producto ( Codram ) seleccionado
            '
            ' Y dentro del producto en función del estado del pago
            '
            If produc = "Todos los Productos" Or produc = "" Then
                If strFiltro <> "T" Then
                    strWhereMas = strWhereMas & " And Angel_T1.EstadoProcCierre = '" & strFiltro & "'"
                End If
            Else
                If strFiltro <> "T" Then
                    strWhereMas = strWhereMas & " And And Angel_T1.T1_Codram = '" & produc & "' And Angel_T1.EstadoProcCierre = '" & strFiltro & "'"
                End If
            End If


            ' Añadimos la parte de la Where que filtrará los registros en
            ' función de que sean o ono el último pago
            '
            If chkUltimoPago.CheckState Then
                strFromMas = " , Angel_T2 "
                strWhereMas = strWhereMas & " And Angel_T2.T2_Ultpag = '1' and " & " Angel_T2.T2_Codsin = Angel_T1.T1_Codsin "
            End If
        End If

        strOrderBy = " Order By Siniestro, Referencia "

        strsql = strSQLSel & strFrom & strFromMas & strWhere & strWhereMas & strOrderBy

        ' Establece origen de datos para Crystal Reports
        '
        strSQLCR = strsql

        Call CargarListView_cierres(lvwCierres, strsql, "", "SINIESTRO", "REFERENCIA", "RAMO", "POLIZA", "ESTADO", "SITUACION", "CIERRE")

        stbEstado.Panels(2).Text = CStr(Format(lvwCierres.Items.Count, "#,##0")) & " Referencias"
        Cursor.Current = Cursors.WaitCursor

        ' Poner atenuados aquellos siniestros que ya esten aperturados.
        ' o con el color identificativo para aquellos que tengan errores o avisos
        '

        For Each objlistitem In lvwCierres.Items
            Select Case objlistitem.SubItems.Item(frmInstCierres.T2_ESTADO.Index).Text
                Case "P"
                    objlistitem.Tag = "1"
                    ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro ya está Procesado")
                    Call ColorListItem(objlistitem, Color.Gray)
                Case "X"
                    objlistitem.Tag = "0"
                    ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro está pendiente de procesar")
                    Call ColorListItem(objlistitem, Color.Green)
                Case "W", "A"
                    objlistitem.Tag = "0"
                    ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro tiene avisos pendientes de resolución")
                    Call ColorListItem(objlistitem, Color.DarkGoldenrod)
                Case "E"
                    objlistitem.Tag = "0"
                    ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro tiene mensajes de error del proceso")
                    Call ColorListItem(objlistitem, Color.Red)
                Case Else
                    objlistitem.Tag = "0"
                    Call ColorListItem(objlistitem, Color.Gray)
            End Select
            If objlistitem.SubItems.Item(frmInstCierres.T2_FESTADO.Index).Text = "C" And objlistitem.SubItems.Item(frmInstCierres.T2_ESTADO.Index).Text <> "P" Then
                objlistitem.Tag = "1"
                objlistitem.SubItems.Item(frmInstCierres.SITUACION.Index).Text = "Cierre Manual"
            End If
            If objlistitem.SubItems.Item(frmInstCierres.T2_FESTADO.Index).Text <> "C" Then
                objlistitem.SubItems.Item(frmInstCierres.SITUACION.Index).Text = objCierres.Bloqueo(objlistitem)
                colAvisosBloqueo.Add(gstrError, objlistitem.Text)
            End If
            System.Windows.Forms.Application.DoEvents()
        Next objlistitem
        Cursor.Current = Cursors.Default
        lvwCierres.Refresh()
        Exit Sub

RefrescaGrid_Error:
        Cursor.Current = Cursors.Default

    End Sub


    Private Sub ActivaControles(ByVal boolActivo As Boolean)
        FiltroAviso.Enabled = boolActivo
        FiltroErrores.Enabled = boolActivo
        FiltroNoPagados.Enabled = boolActivo
        FiltroPagados.Enabled = boolActivo
        FiltroTodos.Enabled = boolActivo

        lblFechaDesde.Enabled = boolActivo
        lblFechaHasta.Enabled = boolActivo
        lblProducto.Enabled = boolActivo
        lblReferencia.Enabled = boolActivo
        lblSiniestro.Enabled = boolActivo
        lblFechaCierre.Enabled = boolActivo
        dtpFechaCierre.Enabled = boolActivo
        lblReferencia.Enabled = boolActivo
        txtSiniestro.Enabled = boolActivo
        txtReferencia.Enabled = boolActivo
        dtpDesde.Enabled = boolActivo
        cbxTipoFecha.Enabled = boolActivo
        dtpHasta.Enabled = boolActivo
        cbxProducto.Enabled = boolActivo
        chkUltimoPago.Enabled = boolActivo
        cbxBusquedaAvanzada.Enabled = boolActivo
        lblTipoFecha.Enabled = boolActivo

        dtpFechaAnulaciones.Enabled = Not boolActivo
        lblFechaAnulaciones.Enabled = Not boolActivo
        dtpFechaAnulaciones.Enabled = Not boolActivo
        rbAnuConfirmados.Enabled = Not boolActivo
        rbAnuProvisionales.Enabled = Not boolActivo
    End Sub

    Private Sub rbAnulaciones_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbAnulaciones.Click
        ActivaControles(False)
    End Sub

    Private Sub rbSiniestros_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbSiniestros.Click
        ActivaControles(True)
    End Sub
    ' Este procedimiento efectua la carga del ListView de Anulaciones
    '
    Private Sub CargaAnulaciones()

        On Error GoTo CargaAnulaciones_Err

        ' Declaraciones
        '
        Dim strsql As String
        Dim FechaAnul As Date

        FechaAnul = dtpFechaAnulaciones.Value

        Cursor.Current = Cursors.WaitCursor

        ' Volvemos a realizar el cruce de referencias para actualizar
        ' los 'No Existe'

        ' Abrimos la transacción
        '

        claseBDCierres.BDWorkConnect.BeginTrans()
        boolTransaccion = True

        '/*MUL T-19908 INi
        'If strCodcia = "R" Then
        Select Case strCodcia
            Case "R"
                ' Primero asignamos el siniestro a cada referencia
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = SNSINIES.Codsin " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Refer = Snsinies.Refext and " & "       AnulacionesAsistencia.T5_Codcia = '" & strCodcia & "' and " & "       AnulacionesAsistencia.T5_Estado <> 'P' "
                claseBDCierres.BDWorkConnect.Execute(strsql)

                ' Si no tienen siniestro abierto marcamos como 'No Existe'
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = 'No Existe' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin is null and " & "       AnulacionesAsistencia.T5_Codcia = '" & strCodcia & "'"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                ' Por último actualizamos el estado de los siniestros anulados
                ' primero los que ya estan cerrados y luego los abiertos y las
                ' denegaciones
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Estado = 'P' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin = Snsinies.Codsin and " & "       Snsinies.Estado = 'C' and AnulacionesAsistencia.T5_Estado = 'X'"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Estado = 'X', AnulacionesAsistencia.T5_Denega = Snsinies.Denega " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin = Snsinies.Codsin and " & "       Snsinies.Estado <> 'C' and AnulacionesAsistencia.T5_Estado = 'X'"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                ' Eliminamos los null del campo denega
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Denega = 'N' " & "WHERE  AnulacionesAsistencia.T5_Denega is null"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                '/*MUL T-19908 INI
                'ElseIf strCodcia = "I" Then

            Case "I", "E", "M"
                '/*MUL T-19908 FIN
                ' Primero asignamos el siniestro a cada referencia
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = SNSINIES.Codsin " & "FROM   Snsinies " & "WHERE  'AS' + AnulacionesAsistencia.T5_Refer = Snsinies.Refext and " & "       AnulacionesAsistencia.T5_Codcia = '" & strCodcia & "' and " & "       AnulacionesAsistencia.T5_Estado <> 'P' "
                claseBDCierres.BDWorkConnect.Execute(strsql)

                ' Si no tienen siniestro abierto marcamos como 'No Existe'
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = 'No Existe' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin is null and " & "       AnulacionesAsistencia.T5_Codcia = '" & strCodcia & "'"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                ' Por último actualizamos el estado de los siniestros anulados
                ' primero los que ya estan cerrados y luego los abiertos y las
                ' denegaciones
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Estado = 'P' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin = Snsinies.Codsin and " & "       Snsinies.Estado = 'C' and AnulacionesAsistencia.T5_Estado = 'X'"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Estado = 'X', AnulacionesAsistencia.T5_Denega = Snsinies.Denega " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin = Snsinies.Codsin and " & "       Snsinies.Estado <> 'C' and AnulacionesAsistencia.T5_Estado = 'X'"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                ' Eliminamos los null del campo denega
                '
                strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Denega = 'N' " & "WHERE  AnulacionesAsistencia.T5_Denega is null"
                claseBDCierres.BDWorkConnect.Execute(strsql)

                '/*MUL T-19908 INI
                'End If

            Case Else
                'No hacer nada
        End Select
        '/*MUL T-19908 FIN

        ' Cerramos la transacción
        '
        claseBDCierres.BDWorkConnect.CommitTrans()
        boolTransaccion = False

        ' Construimos la sentencia Sql en base a la opción determinada en el menú
        '
        If rbAnuProvisionales.Checked Then
            strsql = "Select T5_Codsin Siniestro, T5_Refer Referencia, T5_FecAnula F_Anul , T5_CodRechazo Código " & "From   AnulacionesAsistencia " & "Where  T5_Codcia = '" & strCodcia & "' and (T5_Estado <> 'P' and T5_Denega <> 'S') and " & "       AnulacionesAsistencia.T5_Fgraba = '" & claseUtilidadesCierres.FormatoFechaSQL(FechaAnul, False, False) & "' and " & "       T5_Tipmov = '5'" & "Order By T5_Codsin"
        Else
            strsql = "Select T5_Codsin Siniestro, T5_Refer Referencia, T5_FecAnula F_Anul , T5_CodRechazo Código " & "From   AnulacionesAsistencia " & "Where  T5_Codcia = '" & strCodcia & "' and (T5_Estado <> 'P' and T5_Denega <> 'S') and " & "       AnulacionesAsistencia.T5_Fgraba = '" & claseUtilidadesCierres.FormatoFechaSQL(FechaAnul, False, False) & "' and " & "       T5_Tipmov = '6' " & "Order By T5_Codsin"
        End If

        strSQLCR = strsql

        Call CargarListView_cierres(lvwAnulaciones, strsql, "REFERENCIA", "SINIESTRO", "REFERENCIA", "F_ANUL", "CÓDIGO")

        'UPGRADE_WARNING: Screen propiedad Screen.MousePointer tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
        Cursor.Current = Cursors.Default

        Exit Sub

CargaAnulaciones_Err:
        If Err.Number = -2147217871 Then
            Resume
        Else
            'UPGRADE_WARNING: Screen propiedad Screen.MousePointer tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Private Sub lvwCierres_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lvwCierres.ItemCheck
        Dim lvPagosAux As ListView

        lvPagosAux = sender
        If lvPagosAux.Items(e.Index).Tag = "1" Then
            e.NewValue = e.CurrentValue
        End If
    End Sub

    Private Sub cbxBusquedaAvanzada_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxBusquedaAvanzada.CheckedChanged
        Dim boolActivo As Boolean
        boolActivo = cbxBusquedaAvanzada.Checked
        'Habilitados
        lblSiniestro.Enabled = boolActivo
        lblReferencia.Enabled = boolActivo
        txtReferencia.Enabled = boolActivo
        txtSiniestro.Enabled = boolActivo


        'Deshabilitados
        FiltroAviso.Enabled = Not boolActivo
        FiltroErrores.Enabled = Not boolActivo
        FiltroNoPagados.Enabled = Not boolActivo
        FiltroPagados.Enabled = Not boolActivo
        FiltroTodos.Enabled = Not boolActivo
        lblProducto.Enabled = Not boolActivo
        lblFechaDesde.Enabled = Not boolActivo
        lblTipoFecha.Enabled = Not boolActivo
        dtpDesde.Enabled = Not boolActivo
        dtpHasta.Enabled = Not boolActivo
        cbxTipoFecha.Enabled = Not boolActivo
        lblFechaDesde.Enabled = Not boolActivo
        lblFechaHasta.Enabled = Not boolActivo
        cbxProducto.Enabled = Not boolActivo
        cbxTipoFecha.Enabled = Not boolActivo
        chkFiltroAvisos.Enabled = Not boolActivo
        chkUltimoPago.Enabled = Not boolActivo
        lblFechaCierre.Enabled = Not boolActivo
        dtpFechaCierre.Enabled = Not boolActivo
    End Sub

    Private Sub cbCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCerrar.Click
        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim strRefer As String
        Dim vntReferencia As Object
        Dim Result As Boolean
        Dim Continua As Boolean

        objCierres = New clsCierres_NET

        prbProgreso.Minimum = 0
        prbProgreso.Value = 1
        Continua = False
        Result = False

        stbEstado.Panels(2).Text = "Comprobando Filtros ..."

        FechaEjecucion = Now
        NombreFichero = "K:\Siniestros\Asistencia\Cierres\" & Format(Now, "yyyyMMdd") & "_Cierres_Asistencia.Log"
        Asunto = "Proceso Cierres Asistencia: Finalizado Correctamente"
        Mensaje = "Se adjunta fichero log con la información del proceso"

        If rbSiniestros.Checked = True Then
            prbProgreso.Maximum = lvwCierres.Items.Count
            For Each objlistitem In lvwCierres.Items
                Call ActualizarPorcentaje(lvwCierres.Items.Count, prbProgreso, stbEstado)
                If objlistitem.Checked And objlistitem.SubItems.Item(frmInstCierres.T2_ESTADO.Index).Text <> "P" Then
                    objCierres.Referencias.Add(objlistitem.Text) ' Añadir siniestros que se pueden procesar a colección.
                End If
            Next objlistitem
        ElseIf rbAnulaciones.Checked Then
            prbProgreso.Maximum = lvwAnulaciones.Items.Count
            For Each objlistitem In lvwAnulaciones.Items
                Call ActualizarPorcentaje(lvwAnulaciones.Items.Count, prbProgreso, stbEstado)
                If objlistitem.Checked And objlistitem.SubItems.Item(frmInstCierres.T2_ESTADO.Index).Text <> "P" Then
                    objCierres.Referencias.Add(objlistitem.Text) ' Añadir siniestros que se pueden procesar a colección.
                End If
            Next objlistitem
        End If

        ' JLL - 19/10/2010
        '
        ' Se añade la creación de un fichero log donde se informa del inicio
        ' y el final del proceso así como de todas las referencias cerradas
        ' Si no se puede crear el fichero .log no se puede ejecutar el proceso
        '
        If objCierres.CabeceraLog Then
            Continua = True
        Else
            MsgBox("No se ha podido crear el fichero .Log, el proceso de cierres no puede ejecutarse", MsgBoxStyle.Exclamation)
            Continua = False
        End If
        '
        ' Fin JLL - 19/10/2010

        If objCierres.Referencias.Count() > 0 And Continua Then

            stbEstado.Panels(2).Text = "Realizando Cierres de Siniestros ..."

            prbProgreso.Minimum = 0
            prbProgreso.Value = 0
            prbProgreso.Maximum = objCierres.Referencias.Count()
            For Each vntReferencia In objCierres.Referencias
                If rbSiniestros.Checked Then
                    Call ActualizarPorcentaje(lvwCierres.Items.Count, prbProgreso, stbEstado)
                    objlistitem = BuscarItemPorNombreListView(vntReferencia, lvwCierres)
                    If Not objlistitem Is Nothing Then Call objCierres.Cerrar(objlistitem, "Siniestros")
                    If Str(TimeOfDay.ToOADate) > HoraTopeEjecucion Then
                        Asunto = "Proceso Cierres Asistencia: Referencias Pendientes"
                        Mensaje = "El proceso no pudo cerrar todas las referencias seleccionadas. Consultar el fichero adjunto para ver el informe"
                        Result = objCierres.BorraSiniestrosPendientes
                        Result = objCierres.GrabaCierrePendiente
                        Exit For
                    End If

                ElseIf rbAnulaciones.Checked Then
                    Call ActualizarPorcentaje(lvwAnulaciones.Items.Count, prbProgreso, stbEstado)
                    objlistitem = BuscarItemPorNombreListView(vntReferencia, lvwAnulaciones)
                    If Not objlistitem Is Nothing Then Call objCierres.Cerrar(objlistitem, "Anulaciones")
                End If
            Next vntReferencia
            If colSiniestrosPendientes.Count() = 0 Then
                Result = objCierres.BorraSiniestrosPendientes
            End If
            Result = objCierres.InsertarLog
            Result = objCierres.PieLog
            Result = EnviaEmail()
            stbEstado.Panels(2).Text = "Proceso Finalizado"
        Else
            stbEstado.Panels(2).Text = "No se han realizado los Cierres"
        End If

        prbProgreso.Visible = False
        strSigCompa = lbxCompania.Items(cbxCompania.SelectedIndex)
        strFiltro = "T"
        FiltroTodos.PerformClick()
        Call RefrescarGrid(dtpDesde.Value, dtpHasta.Value, cbxProducto.Text, strSigCompa)

        If Not Result Then
            MsgBox("El Proceso de Cierres no ha cerrado todas las referencias. El envio del email ha fallado.", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Function BuscarItemPorNombreListView(ByVal strTexto As String, ByVal lvwAux As ListView) As ListViewItem

        For Each lvi As ListViewItem In lvwAux.Items
            If lvi.Text.Equals(strTexto) Then Return lvi
            For Each si As ListViewItem.ListViewSubItem In lvi.SubItems
                If si.Text.Equals(strTexto) Then Return lvi
            Next
        Next
        Return Nothing

    End Function


    Private Function EnviaEmail() As Boolean

        On Error GoTo EnviaEmail_Error

        If claseBDCierres.BDAuxRecord.State = 1 Then claseBDCierres.BDAuxRecord.Close()

        With claseBDCierres.BDAuxRecord
            .Open("email_programado", claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            .AddNew()
            .Fields("De").Value = "MDP MAIL"
            .Fields("Para").Value = "Asistencia@mutuadepropietarios.es"
            .Fields("Cc").Value = "araceli.frances@mutuadepropietarios.es,jl.lacalle@mutuadepropietarios.es"
            .Fields("Bcc").Value = ""
            .Fields("Asunto").Value = Asunto
            .Fields("adjunto").Value = NombreFichero
            .Fields("Texto").Value = Mensaje
            .Fields("Fecha").Value = FechaEjecucion
            .Fields("enviado").Value = "N"
            .Update()
            .Close()
        End With
        EnviaEmail = True

        Exit Function

EnviaEmail_Error:
        EnviaEmail = False
    End Function

    Private Sub lvwCierres_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvwCierres.SelectedIndexChanged
        If IsNothing(frmInstanciaAnulaciones) Then
            frmInstanciaAnulaciones = New frmAnulaciones
        End If
        If lvwCierres.SelectedItems.Count > 0 Then
            If lvwCierres.SelectedItems(0).SubItems(6).Text = "Bloqueado" Then
                frmInstanciaAvisoBloqueos.Show()
            End If
        End If
    End Sub

    Private Sub lvwAnulaciones_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvwAnulaciones.SelectedIndexChanged
        lvwCierres.SelectedItems.Clear()
        If lvwAnulaciones.SelectedItems.Count > 0 Then
            frmInstanciaAnulaciones.Show()
        End If
    End Sub

    Private Sub cbTodos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTodos.Click, cbNinguno.Click, cbPendientes.Click
        Dim strOpcion As String
        Dim cbBotonAux As Button

        cbBotonAux = CType(sender, Button)

        strOpcion = cbBotonAux.Tag
        If strOpcion <> "" Then
            SeleccionarRegistros(strOpcion)
        End If
    End Sub

    Private Sub SeleccionarRegistros(ByVal strOpcion As String)
        ' Declaraciones
        '
        Dim blnCheck As Boolean
        Dim objlistitem As ListViewItem

        Select Case strOpcion

            Case "TODOS"
                blnCheck = True

            Case "NINGUNO" ' Fecha en la que se importo el pago
                blnCheck = False

            Case "PENDIENTES" ' Fecha en la que Mutua realiza el pago
                blnCheck = True

            Case Else
                blnCheck = False

        End Select


        ' Seleccionar o deseleccionar todos los elementos de la lista
        ' Teniendo en cuenta que aquellos siniestros esten aperturados
        ' no se pueden seleccionar
        If rbSiniestros.Checked Then
            If lvwCierres.Items.Count > 0 Then
                For Each objlistitem In lvwCierres.Items
                    If objlistitem.Tag <> "1" Then
                        If objlistitem.Text <> "" And objlistitem.Text <> "No Existe" And objlistitem.SubItems(frmInstCierres.SITUACION.Index).Text <> "Bloqueado" And objlistitem.SubItems(frmInstCierres.SITUACION.Index).Text <> "Error" And strOpcion = "TODOS" Then
                            Call ColorListItem(objlistitem, Color.Green)
                            objlistitem.Checked = blnCheck
                        ElseIf objlistitem.Text <> "" And objlistitem.Text <> "No Existe" And objlistitem.SubItems(frmInstCierres.SITUACION.Index).Text <> "Bloqueado" And objlistitem.SubItems(frmInstCierres.SITUACION.Index).Text <> "Error" And strOpcion = "PENDIENTES" And EstaPendiente(objlistitem) Then
                            objlistitem.Checked = blnCheck
                            Call ColorListItem(objlistitem, Color.Green)

                        ElseIf objlistitem.Text <> "" And objlistitem.Text <> "No Existe" And objlistitem.SubItems(frmInstCierres.SITUACION.Index).Text <> "Bloqueado" And objlistitem.SubItems(frmInstCierres.SITUACION.Index).Text <> "Error" And strOpcion = "NINGUNO" Then
                            Call ColorListItem(objlistitem, Color.Green)
                            objlistitem.Checked = blnCheck
                        End If
                    End If
                Next
            End If
        Else
            ' Seleccionar o deseleccionar todos los elementos de la lista
            ' Teniendo en cuenta que aquellos siniestros esten aperturados
            ' no se pueden seleccionar
            If lvwAnulaciones.Items.Count > 0 Then
                For Each objlistitem In lvwAnulaciones.Items
                    If objlistitem.Tag <> "1" Then
                        If objlistitem.Text <> "" And objlistitem.Text <> "No Existe" Then
                            Call ColorListItem(objlistitem, Color.Green)
                            objlistitem.Checked = blnCheck
                        End If
                    End If
                Next objlistitem
            End If
        End If

    End Sub

    Private Sub dtpFechaCierre_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFechaCierre.ValueChanged
        'FechaCierre = claseUtilidadesCierres.FormatoFechaSQL(dtpFechaCierre.Value, False, True)
        FechaCierre = sender.Value
    End Sub

    Private Sub cbSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSalir.Click
        Application.Exit()
    End Sub

    Private Sub cbAvisos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAvisos.Click
        On Error GoTo cbAvisos_Error

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim strReferencia As String
        Dim claseSiniestros As clsSiniestro_NET
        Dim frmInstanciaErrores As New frmVisorErrores

        If rbSiniestros.Checked Then
            ' Comprobamos que se haya seleccionados
            If lvwCierres.CheckedItems.Count > 0 Then
                ' Comprobamos el estado de la referencia seleccionada
                If lvwCierres.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "W" Or lvwCierres.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "A" Or lvwCierres.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "E" Then
                    If lvwCierres.CheckedItems(0).Tag = "1" Then Exit Sub
                    frmInstanciaErrores.Show()
                    objlistitem = lvwCierres.CheckedItems(0)
                    strReferencia = lvwCierres.CheckedItems(0).SubItems(frmInstCierres.T2_REFER.Index).Text
                    frmInstanciaErrores.MostrarErrores(strReferencia)
                ElseIf lvwCierres.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "P" Then
                    claseBDCierres.BDAuxRecord = claseSiniestros.Siniestro(lvwCierres.CheckedItems(0).SubItems(frmInstCierres.T2_CODSIN.Index).Text, True, strUsuaApli)
                Else
                    MsgBox("La referencia seleccionada no tiene avisos / errores.", MsgBoxStyle.Information)
                End If
            Else
                'MsgBox("No ha seleccionado ninguna referencia", MsgBoxStyle.Information)
                '/* MUL  si no se selecciona ninguno se muestra el historial.
                'strError = "4007"
                'objError.Tipo = Pantalla
                'objError.Ver(IdProceso, gstrError, , Codcia)
                frmInstanciaErrores.dtpFechaInicio.Value = dtpDesde.Value
                frmInstanciaErrores.dtpFechaFin.Value = dtpHasta.Value
                frmInstanciaErrores.MostrarErrores("")
                frmInstanciaErrores.Show()
            End If
        Else
            ' Comprobamos que se haya seleccionados
            If lvwAnulaciones.CheckedItems.Count > 0 Then
                ' Comprobamos el estado de la referencia seleccionada
                If lvwAnulaciones.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "W" Or _
                   lvwAnulaciones.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "A" Or _
                   lvwAnulaciones.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "E" Then
                    If lvwAnulaciones.CheckedItems(0).Tag = "1" Then Exit Sub
                    frmInstanciaErrores.Show()
                    objlistitem = lvwAnulaciones.CheckedItems(0)
                    strReferencia = lvwAnulaciones.CheckedItems(0).SubItems(frmInstCierres.T2_REFER.Index).Text
                    frmInstanciaErrores.MostrarErrores(strReferencia)
                ElseIf lvwAnulaciones.CheckedItems(0).SubItems(frmInstCierres.T2_ESTADO.Index).Text = "P" Then
                    claseBDCierres.BDAuxRecord = claseSiniestros.Siniestro(lvwAnulaciones.CheckedItems(0).SubItems(frmInstCierres.T2_CODSIN.Index).Text, True, strUsuaApli)
                End If
            Else
                'MsgBox("No ha seleccionado ninguna referencia", MsgBoxStyle.Information)
                '/* MUL  si no se selecciona ninguno se muestra el historial.
                'strError = "4007"
                'objError.Tipo = Pantalla
                'objError.Ver(IdProceso, gstrError, , Codcia)
                frmInstanciaErrores.dtpFechaInicio.Value = dtpDesde.Value
                frmInstanciaErrores.dtpFechaFin.Value = dtpHasta.Value
                frmInstanciaErrores.MostrarErrores("")
                frmInstanciaErrores.Show()
            End If

        End If
        Exit Sub

cbAvisos_Error:
        MsgBox("Ha ocurrido un error mostrando el aviso", MsgBoxStyle.Critical)
    End Sub

    Private Sub cbImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbImprimir.Click
        Imprimir_cierres()
    End Sub
End Class
