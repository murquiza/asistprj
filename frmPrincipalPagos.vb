Public Class frmPrincipalPagos
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
    Friend WithEvents chkDtoReparalia As System.Windows.Forms.CheckBox
    Friend WithEvents chkFiltroAvisos As System.Windows.Forms.CheckBox
    Friend WithEvents chkUltimoPago As System.Windows.Forms.CheckBox
    Friend WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbxTipoFecha As System.Windows.Forms.ComboBox
    Friend WithEvents cbxTipoPago As System.Windows.Forms.ComboBox
    Friend WithEvents stbEstado As System.Windows.Forms.StatusBar
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents RTotal As System.Windows.Forms.Label
    Public WithEvents RIva As System.Windows.Forms.Label
    Public WithEvents RImporte As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lbTotal As System.Windows.Forms.Label
    Friend WithEvents lbIVA As System.Windows.Forms.Label
    Friend WithEvents lbImporte As System.Windows.Forms.Label
    Friend WithEvents lbResumenReferencias As System.Windows.Forms.Label
    Friend WithEvents lbResumenSiniestros As System.Windows.Forms.Label
    Friend WithEvents lbxCompania As System.Windows.Forms.ListBox
    Friend WithEvents lbxProducto As System.Windows.Forms.ListBox
    Friend WithEvents FiltroErrores As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroAviso As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroNoPagados As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroPagados As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroTodos As System.Windows.Forms.RadioButton
    Friend WithEvents gbxFiltro As System.Windows.Forms.GroupBox
    Friend WithEvents lvwPagos As System.Windows.Forms.ListView
    Friend WithEvents ttipAyuda As System.Windows.Forms.ToolTip
    Friend WithEvents T2_CODSIN As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_REFER As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_FPAGO As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECGRA As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_ESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_FESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents SITUACION As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_CAUSPER As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_IMPOR As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_IMPTVA As System.Windows.Forms.ColumnHeader
    Friend WithEvents TOTAL As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_PAGADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_TIPGAS As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_ULTPAG As System.Windows.Forms.ColumnHeader
    Friend WithEvents PERITO As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_POLIZA As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_CODRAM As System.Windows.Forms.ColumnHeader
    Friend WithEvents T2_NUMORD As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECHAPROC As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECHACAUSA As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECHAEXPORT As System.Windows.Forms.ColumnHeader
    Friend WithEvents FACTURA As System.Windows.Forms.ColumnHeader
    Friend WithEvents MODO_GAR As System.Windows.Forms.ColumnHeader
    Friend WithEvents GRUPO_GAR As System.Windows.Forms.ColumnHeader
    Friend WithEvents cbSalir As System.Windows.Forms.Button
    Friend WithEvents cbBorrar As System.Windows.Forms.Button
    Friend WithEvents cbPagos As System.Windows.Forms.Button
    Friend WithEvents cbAvisos As System.Windows.Forms.Button
    Friend WithEvents cbImprimir As System.Windows.Forms.Button
    Friend WithEvents cbTodos As System.Windows.Forms.Button
    Friend WithEvents cbNinguno As System.Windows.Forms.Button
    Friend WithEvents prbProgreso As System.Windows.Forms.ProgressBar
    Friend WithEvents sbPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents tbSiniestro As System.Windows.Forms.TextBox
    Friend WithEvents tbReferencia As System.Windows.Forms.TextBox
    Friend WithEvents lbSiniestro As System.Windows.Forms.Label
    Friend WithEvents lbReferencia As System.Windows.Forms.Label
    Friend WithEvents cbxCompania As System.Windows.Forms.ComboBox
    Friend WithEvents cbxProducto As System.Windows.Forms.ComboBox
    Friend WithEvents lbTipoPago As System.Windows.Forms.Label
    Friend WithEvents lbTipoFecha As System.Windows.Forms.Label
    Friend WithEvents lbFechaHasta As System.Windows.Forms.Label
    Friend WithEvents lbFechaDesde As System.Windows.Forms.Label
    Friend WithEvents lbProducto As System.Windows.Forms.Label
    Friend WithEvents lbCompaniaAsistencia As System.Windows.Forms.Label
    Public WithEvents CR2 As AxCrystal.AxCrystalReport
    Friend WithEvents picTest As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrincipalPagos))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lbCompaniaAsistencia = New System.Windows.Forms.Label
        Me.cbxCompania = New System.Windows.Forms.ComboBox
        Me.gbxFiltro = New System.Windows.Forms.GroupBox
        Me.chkDtoReparalia = New System.Windows.Forms.CheckBox
        Me.chkFiltroAvisos = New System.Windows.Forms.CheckBox
        Me.chkUltimoPago = New System.Windows.Forms.CheckBox
        Me.FiltroErrores = New System.Windows.Forms.RadioButton
        Me.FiltroAviso = New System.Windows.Forms.RadioButton
        Me.FiltroNoPagados = New System.Windows.Forms.RadioButton
        Me.FiltroPagados = New System.Windows.Forms.RadioButton
        Me.FiltroTodos = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cbBuscar = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.tbSiniestro = New System.Windows.Forms.TextBox
        Me.tbReferencia = New System.Windows.Forms.TextBox
        Me.lbSiniestro = New System.Windows.Forms.Label
        Me.lbReferencia = New System.Windows.Forms.Label
        Me.cbxBusquedaAvanzada = New System.Windows.Forms.CheckBox
        Me.cbxTipoPago = New System.Windows.Forms.ComboBox
        Me.cbxTipoFecha = New System.Windows.Forms.ComboBox
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker
        Me.cbxProducto = New System.Windows.Forms.ComboBox
        Me.lbTipoPago = New System.Windows.Forms.Label
        Me.lbTipoFecha = New System.Windows.Forms.Label
        Me.lbFechaHasta = New System.Windows.Forms.Label
        Me.lbFechaDesde = New System.Windows.Forms.Label
        Me.lbProducto = New System.Windows.Forms.Label
        Me.lvwPagos = New System.Windows.Forms.ListView
        Me.T2_CODSIN = New System.Windows.Forms.ColumnHeader
        Me.T2_REFER = New System.Windows.Forms.ColumnHeader
        Me.T2_FPAGO = New System.Windows.Forms.ColumnHeader
        Me.FECGRA = New System.Windows.Forms.ColumnHeader
        Me.T2_ESTADO = New System.Windows.Forms.ColumnHeader
        Me.T2_FESTADO = New System.Windows.Forms.ColumnHeader
        Me.SITUACION = New System.Windows.Forms.ColumnHeader
        Me.T2_CAUSPER = New System.Windows.Forms.ColumnHeader
        Me.T2_IMPOR = New System.Windows.Forms.ColumnHeader
        Me.T2_IMPTVA = New System.Windows.Forms.ColumnHeader
        Me.TOTAL = New System.Windows.Forms.ColumnHeader
        Me.T2_PAGADO = New System.Windows.Forms.ColumnHeader
        Me.T2_TIPGAS = New System.Windows.Forms.ColumnHeader
        Me.T2_ULTPAG = New System.Windows.Forms.ColumnHeader
        Me.PERITO = New System.Windows.Forms.ColumnHeader
        Me.T2_POLIZA = New System.Windows.Forms.ColumnHeader
        Me.T2_CODRAM = New System.Windows.Forms.ColumnHeader
        Me.T2_NUMORD = New System.Windows.Forms.ColumnHeader
        Me.FECHAPROC = New System.Windows.Forms.ColumnHeader
        Me.FECHACAUSA = New System.Windows.Forms.ColumnHeader
        Me.FECHAEXPORT = New System.Windows.Forms.ColumnHeader
        Me.FACTURA = New System.Windows.Forms.ColumnHeader
        Me.MODO_GAR = New System.Windows.Forms.ColumnHeader
        Me.GRUPO_GAR = New System.Windows.Forms.ColumnHeader
        Me.stbEstado = New System.Windows.Forms.StatusBar
        Me.sbPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cbTodos = New System.Windows.Forms.Button
        Me.cbNinguno = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.lbTotal = New System.Windows.Forms.Label
        Me.lbIVA = New System.Windows.Forms.Label
        Me.lbImporte = New System.Windows.Forms.Label
        Me.lbResumenReferencias = New System.Windows.Forms.Label
        Me.lbResumenSiniestros = New System.Windows.Forms.Label
        Me.RTotal = New System.Windows.Forms.Label
        Me.RIva = New System.Windows.Forms.Label
        Me.RImporte = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.cbBorrar = New System.Windows.Forms.Button
        Me.cbPagos = New System.Windows.Forms.Button
        Me.cbAvisos = New System.Windows.Forms.Button
        Me.cbImprimir = New System.Windows.Forms.Button
        Me.cbSalir = New System.Windows.Forms.Button
        Me.lbxCompania = New System.Windows.Forms.ListBox
        Me.lbxProducto = New System.Windows.Forms.ListBox
        Me.ttipAyuda = New System.Windows.Forms.ToolTip(Me.components)
        Me.prbProgreso = New System.Windows.Forms.ProgressBar
        Me.CR2 = New AxCrystal.AxCrystalReport
        Me.picTest = New System.Windows.Forms.PictureBox
        Me.gbxFiltro.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.CR2, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.lbCompaniaAsistencia.Location = New System.Drawing.Point(64, 16)
        Me.lbCompaniaAsistencia.Name = "lbCompaniaAsistencia"
        Me.lbCompaniaAsistencia.Size = New System.Drawing.Size(152, 24)
        Me.lbCompaniaAsistencia.TabIndex = 1
        Me.lbCompaniaAsistencia.Text = "Compañía Asistencia:"
        Me.lbCompaniaAsistencia.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxCompania
        '
        Me.cbxCompania.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxCompania.Location = New System.Drawing.Point(224, 16)
        Me.cbxCompania.Name = "cbxCompania"
        Me.cbxCompania.Size = New System.Drawing.Size(552, 24)
        Me.cbxCompania.TabIndex = 1
        '
        'gbxFiltro
        '
        Me.gbxFiltro.Controls.Add(Me.chkDtoReparalia)
        Me.gbxFiltro.Controls.Add(Me.chkFiltroAvisos)
        Me.gbxFiltro.Controls.Add(Me.chkUltimoPago)
        Me.gbxFiltro.Controls.Add(Me.FiltroErrores)
        Me.gbxFiltro.Controls.Add(Me.FiltroAviso)
        Me.gbxFiltro.Controls.Add(Me.FiltroNoPagados)
        Me.gbxFiltro.Controls.Add(Me.FiltroPagados)
        Me.gbxFiltro.Controls.Add(Me.FiltroTodos)
        Me.gbxFiltro.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxFiltro.ForeColor = System.Drawing.Color.RoyalBlue
        Me.gbxFiltro.Location = New System.Drawing.Point(568, 56)
        Me.gbxFiltro.Name = "gbxFiltro"
        Me.gbxFiltro.Size = New System.Drawing.Size(208, 184)
        Me.gbxFiltro.TabIndex = 3
        Me.gbxFiltro.TabStop = False
        Me.gbxFiltro.Text = "Filtro"
        Me.ttipAyuda.SetToolTip(Me.gbxFiltro, "Filtros sobre el resultado de la consulta")
        '
        'chkDtoReparalia
        '
        Me.chkDtoReparalia.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.chkDtoReparalia.Location = New System.Drawing.Point(32, 164)
        Me.chkDtoReparalia.Name = "chkDtoReparalia"
        Me.chkDtoReparalia.Size = New System.Drawing.Size(120, 16)
        Me.chkDtoReparalia.TabIndex = 18
        Me.chkDtoReparalia.Text = "10% de descuento"
        '
        'chkFiltroAvisos
        '
        Me.chkFiltroAvisos.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.chkFiltroAvisos.Location = New System.Drawing.Point(32, 144)
        Me.chkFiltroAvisos.Name = "chkFiltroAvisos"
        Me.chkFiltroAvisos.Size = New System.Drawing.Size(168, 16)
        Me.chkFiltroAvisos.TabIndex = 17
        Me.chkFiltroAvisos.Text = "Desactivar filtros para avisos"
        Me.ttipAyuda.SetToolTip(Me.chkFiltroAvisos, "Desactiva los Filtros para procesar los Avisos")
        '
        'chkUltimoPago
        '
        Me.chkUltimoPago.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.chkUltimoPago.Location = New System.Drawing.Point(32, 124)
        Me.chkUltimoPago.Name = "chkUltimoPago"
        Me.chkUltimoPago.Size = New System.Drawing.Size(104, 16)
        Me.chkUltimoPago.TabIndex = 16
        Me.chkUltimoPago.Text = "Último pago"
        Me.ttipAyuda.SetToolTip(Me.chkUltimoPago, "Muestra sólo aquellos siniestros cuyo movimiento de pago sea el último")
        '
        'FiltroErrores
        '
        Me.FiltroErrores.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroErrores.ForeColor = System.Drawing.Color.Red
        Me.FiltroErrores.Location = New System.Drawing.Point(16, 101)
        Me.FiltroErrores.Name = "FiltroErrores"
        Me.FiltroErrores.Size = New System.Drawing.Size(112, 16)
        Me.FiltroErrores.TabIndex = 4
        Me.FiltroErrores.Text = "Errores ( E )"
        Me.ttipAyuda.SetToolTip(Me.FiltroErrores, "Muestra sólo los registros en los que se ha producido error")
        '
        'FiltroAviso
        '
        Me.FiltroAviso.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroAviso.ForeColor = System.Drawing.Color.DarkGoldenrod
        Me.FiltroAviso.Location = New System.Drawing.Point(16, 80)
        Me.FiltroAviso.Name = "FiltroAviso"
        Me.FiltroAviso.Size = New System.Drawing.Size(112, 16)
        Me.FiltroAviso.TabIndex = 3
        Me.FiltroAviso.Text = "Aviso ( A )"
        Me.ttipAyuda.SetToolTip(Me.FiltroAviso, "Muestra sólo los registros con mensaje de aviso")
        '
        'FiltroNoPagados
        '
        Me.FiltroNoPagados.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroNoPagados.ForeColor = System.Drawing.Color.Green
        Me.FiltroNoPagados.Location = New System.Drawing.Point(16, 59)
        Me.FiltroNoPagados.Name = "FiltroNoPagados"
        Me.FiltroNoPagados.Size = New System.Drawing.Size(112, 16)
        Me.FiltroNoPagados.TabIndex = 2
        Me.FiltroNoPagados.Text = "No Pagados ( X )"
        Me.ttipAyuda.SetToolTip(Me.FiltroNoPagados, "Muestra sólo los pendientes de aperturar")
        '
        'FiltroPagados
        '
        Me.FiltroPagados.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroPagados.ForeColor = System.Drawing.Color.Gray
        Me.FiltroPagados.Location = New System.Drawing.Point(16, 38)
        Me.FiltroPagados.Name = "FiltroPagados"
        Me.FiltroPagados.Size = New System.Drawing.Size(112, 16)
        Me.FiltroPagados.TabIndex = 1
        Me.FiltroPagados.Text = "Pagados ( P )"
        Me.ttipAyuda.SetToolTip(Me.FiltroPagados, "Muestra sólo los aperturados")
        '
        'FiltroTodos
        '
        Me.FiltroTodos.Checked = True
        Me.FiltroTodos.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroTodos.ForeColor = System.Drawing.Color.Black
        Me.FiltroTodos.Location = New System.Drawing.Point(16, 17)
        Me.FiltroTodos.Name = "FiltroTodos"
        Me.FiltroTodos.Size = New System.Drawing.Size(112, 16)
        Me.FiltroTodos.TabIndex = 0
        Me.FiltroTodos.TabStop = True
        Me.FiltroTodos.Text = "Todos"
        Me.ttipAyuda.SetToolTip(Me.FiltroTodos, "Muestra todos los registros")
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbBuscar)
        Me.GroupBox2.Controls.Add(Me.Panel1)
        Me.GroupBox2.Controls.Add(Me.cbxTipoPago)
        Me.GroupBox2.Controls.Add(Me.cbxTipoFecha)
        Me.GroupBox2.Controls.Add(Me.dtpHasta)
        Me.GroupBox2.Controls.Add(Me.dtpDesde)
        Me.GroupBox2.Controls.Add(Me.cbxProducto)
        Me.GroupBox2.Controls.Add(Me.lbTipoPago)
        Me.GroupBox2.Controls.Add(Me.lbTipoFecha)
        Me.GroupBox2.Controls.Add(Me.lbFechaHasta)
        Me.GroupBox2.Controls.Add(Me.lbFechaDesde)
        Me.GroupBox2.Controls.Add(Me.lbProducto)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox2.Location = New System.Drawing.Point(8, 70)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(552, 170)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Criterios de Selección"
        '
        'cbBuscar
        '
        Me.cbBuscar.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbBuscar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBuscar.Image = CType(resources.GetObject("cbBuscar.Image"), System.Drawing.Image)
        Me.cbBuscar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbBuscar.Location = New System.Drawing.Point(464, 24)
        Me.cbBuscar.Name = "cbBuscar"
        Me.cbBuscar.Size = New System.Drawing.Size(72, 56)
        Me.cbBuscar.TabIndex = 10
        Me.cbBuscar.Text = "Buscar"
        Me.cbBuscar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbBuscar, "Ejecuta la busqueda de Pagos según los criterios asignados")
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightGray
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.tbSiniestro)
        Me.Panel1.Controls.Add(Me.tbReferencia)
        Me.Panel1.Controls.Add(Me.lbSiniestro)
        Me.Panel1.Controls.Add(Me.lbReferencia)
        Me.Panel1.Controls.Add(Me.cbxBusquedaAvanzada)
        Me.Panel1.Location = New System.Drawing.Point(8, 120)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(536, 40)
        Me.Panel1.TabIndex = 10
        '
        'tbSiniestro
        '
        Me.tbSiniestro.Enabled = False
        Me.tbSiniestro.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.tbSiniestro.Location = New System.Drawing.Point(392, 8)
        Me.tbSiniestro.Name = "tbSiniestro"
        Me.tbSiniestro.Size = New System.Drawing.Size(88, 21)
        Me.tbSiniestro.TabIndex = 9
        Me.tbSiniestro.Text = ""
        '
        'tbReferencia
        '
        Me.tbReferencia.Enabled = False
        Me.tbReferencia.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.tbReferencia.Location = New System.Drawing.Point(232, 8)
        Me.tbReferencia.Name = "tbReferencia"
        Me.tbReferencia.Size = New System.Drawing.Size(88, 21)
        Me.tbReferencia.TabIndex = 8
        Me.tbReferencia.Text = ""
        '
        'lbSiniestro
        '
        Me.lbSiniestro.Enabled = False
        Me.lbSiniestro.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbSiniestro.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbSiniestro.Location = New System.Drawing.Point(328, 10)
        Me.lbSiniestro.Name = "lbSiniestro"
        Me.lbSiniestro.Size = New System.Drawing.Size(56, 17)
        Me.lbSiniestro.TabIndex = 2
        Me.lbSiniestro.Text = "Siniestro:"
        Me.lbSiniestro.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lbReferencia
        '
        Me.lbReferencia.Enabled = False
        Me.lbReferencia.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbReferencia.Location = New System.Drawing.Point(152, 10)
        Me.lbReferencia.Name = "lbReferencia"
        Me.lbReferencia.Size = New System.Drawing.Size(72, 16)
        Me.lbReferencia.TabIndex = 1
        Me.lbReferencia.Text = "Referencia:"
        Me.lbReferencia.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cbxBusquedaAvanzada
        '
        Me.cbxBusquedaAvanzada.Location = New System.Drawing.Point(8, 6)
        Me.cbxBusquedaAvanzada.Name = "cbxBusquedaAvanzada"
        Me.cbxBusquedaAvanzada.Size = New System.Drawing.Size(136, 24)
        Me.cbxBusquedaAvanzada.TabIndex = 5
        Me.cbxBusquedaAvanzada.Text = "Búsqueda avanzada"
        '
        'cbxTipoPago
        '
        Me.cbxTipoPago.Enabled = False
        Me.cbxTipoPago.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbxTipoPago.Location = New System.Drawing.Point(80, 88)
        Me.cbxTipoPago.Name = "cbxTipoPago"
        Me.cbxTipoPago.Size = New System.Drawing.Size(160, 21)
        Me.cbxTipoPago.TabIndex = 4
        '
        'cbxTipoFecha
        '
        Me.cbxTipoFecha.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbxTipoFecha.Location = New System.Drawing.Point(80, 56)
        Me.cbxTipoFecha.Name = "cbxTipoFecha"
        Me.cbxTipoFecha.Size = New System.Drawing.Size(160, 21)
        Me.cbxTipoFecha.TabIndex = 3
        '
        'dtpHasta
        '
        Me.dtpHasta.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpHasta.Location = New System.Drawing.Point(336, 56)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(104, 21)
        Me.dtpHasta.TabIndex = 7
        Me.ttipAyuda.SetToolTip(Me.dtpHasta, "Abre el calendario de selección")
        '
        'dtpDesde
        '
        Me.dtpDesde.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDesde.Location = New System.Drawing.Point(336, 24)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(104, 21)
        Me.dtpDesde.TabIndex = 6
        Me.ttipAyuda.SetToolTip(Me.dtpDesde, "Abre el calendario de selección")
        '
        'cbxProducto
        '
        Me.cbxProducto.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbxProducto.Location = New System.Drawing.Point(80, 24)
        Me.cbxProducto.Name = "cbxProducto"
        Me.cbxProducto.Size = New System.Drawing.Size(160, 21)
        Me.cbxProducto.TabIndex = 2
        '
        'lbTipoPago
        '
        Me.lbTipoPago.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbTipoPago.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbTipoPago.Location = New System.Drawing.Point(8, 88)
        Me.lbTipoPago.Name = "lbTipoPago"
        Me.lbTipoPago.Size = New System.Drawing.Size(64, 19)
        Me.lbTipoPago.TabIndex = 4
        Me.lbTipoPago.Text = "Tipo Pago:"
        Me.lbTipoPago.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbTipoFecha
        '
        Me.lbTipoFecha.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbTipoFecha.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbTipoFecha.Location = New System.Drawing.Point(8, 57)
        Me.lbTipoFecha.Name = "lbTipoFecha"
        Me.lbTipoFecha.Size = New System.Drawing.Size(64, 18)
        Me.lbTipoFecha.TabIndex = 3
        Me.lbTipoFecha.Text = "Tipo fecha:"
        Me.lbTipoFecha.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbFechaHasta
        '
        Me.lbFechaHasta.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbFechaHasta.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbFechaHasta.Location = New System.Drawing.Point(256, 58)
        Me.lbFechaHasta.Name = "lbFechaHasta"
        Me.lbFechaHasta.Size = New System.Drawing.Size(72, 17)
        Me.lbFechaHasta.TabIndex = 2
        Me.lbFechaHasta.Text = "Fecha hasta:"
        Me.lbFechaHasta.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbFechaDesde
        '
        Me.lbFechaDesde.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbFechaDesde.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbFechaDesde.Location = New System.Drawing.Point(256, 26)
        Me.lbFechaDesde.Name = "lbFechaDesde"
        Me.lbFechaDesde.Size = New System.Drawing.Size(72, 16)
        Me.lbFechaDesde.TabIndex = 1
        Me.lbFechaDesde.Text = "Fecha desde:"
        Me.lbFechaDesde.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbProducto
        '
        Me.lbProducto.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbProducto.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbProducto.Location = New System.Drawing.Point(16, 26)
        Me.lbProducto.Name = "lbProducto"
        Me.lbProducto.Size = New System.Drawing.Size(56, 16)
        Me.lbProducto.TabIndex = 0
        Me.lbProducto.Text = "Producto:"
        Me.lbProducto.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvwPagos
        '
        Me.lvwPagos.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvwPagos.CheckBoxes = True
        Me.lvwPagos.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.T2_CODSIN, Me.T2_REFER, Me.T2_FPAGO, Me.FECGRA, Me.T2_ESTADO, Me.T2_FESTADO, Me.SITUACION, Me.T2_CAUSPER, Me.T2_IMPOR, Me.T2_IMPTVA, Me.TOTAL, Me.T2_PAGADO, Me.T2_TIPGAS, Me.T2_ULTPAG, Me.PERITO, Me.T2_POLIZA, Me.T2_CODRAM, Me.T2_NUMORD, Me.FECHAPROC, Me.FECHACAUSA, Me.FECHAEXPORT, Me.FACTURA, Me.MODO_GAR, Me.GRUPO_GAR})
        Me.lvwPagos.FullRowSelect = True
        Me.lvwPagos.GridLines = True
        Me.lvwPagos.Location = New System.Drawing.Point(8, 248)
        Me.lvwPagos.Name = "lvwPagos"
        Me.lvwPagos.Size = New System.Drawing.Size(768, 243)
        Me.lvwPagos.TabIndex = 19
        Me.lvwPagos.View = System.Windows.Forms.View.Details
        '
        'T2_CODSIN
        '
        Me.T2_CODSIN.Text = "Siniestro"
        Me.T2_CODSIN.Width = 72
        '
        'T2_REFER
        '
        Me.T2_REFER.Text = "Referencia"
        Me.T2_REFER.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_REFER.Width = 68
        '
        'T2_FPAGO
        '
        Me.T2_FPAGO.Text = "Fecha Pago"
        Me.T2_FPAGO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_FPAGO.Width = 74
        '
        'FECGRA
        '
        Me.FECGRA.Text = "F. Importación"
        Me.FECGRA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.FECGRA.Width = 89
        '
        'T2_ESTADO
        '
        Me.T2_ESTADO.Text = "Estado"
        Me.T2_ESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_FESTADO
        '
        Me.T2_FESTADO.Text = "Fecha Estado"
        Me.T2_FESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_FESTADO.Width = 81
        '
        'SITUACION
        '
        Me.SITUACION.Text = "Situación"
        Me.SITUACION.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_CAUSPER
        '
        Me.T2_CAUSPER.Text = "Caus./Perj."
        Me.T2_CAUSPER.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_CAUSPER.Width = 70
        '
        'T2_IMPOR
        '
        Me.T2_IMPOR.Text = "Importe"
        Me.T2_IMPOR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_IMPTVA
        '
        Me.T2_IMPTVA.Text = "IVA"
        Me.T2_IMPTVA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T2_IMPTVA.Width = 41
        '
        'TOTAL
        '
        Me.TOTAL.Text = "TOTAL"
        Me.TOTAL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_PAGADO
        '
        Me.T2_PAGADO.Text = "Pagado"
        Me.T2_PAGADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_TIPGAS
        '
        Me.T2_TIPGAS.Text = "Gas / Ind"
        Me.T2_TIPGAS.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_ULTPAG
        '
        Me.T2_ULTPAG.Text = "Ult. Pago"
        Me.T2_ULTPAG.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PERITO
        '
        Me.PERITO.Text = "Otro Per."
        Me.PERITO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_POLIZA
        '
        Me.T2_POLIZA.Text = "Poliza"
        Me.T2_POLIZA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_CODRAM
        '
        Me.T2_CODRAM.Text = "Ramo"
        Me.T2_CODRAM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T2_NUMORD
        '
        Me.T2_NUMORD.Text = "Núm. Orden"
        Me.T2_NUMORD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FECHAPROC
        '
        Me.FECHAPROC.Text = "Fecha Proceso"
        Me.FECHAPROC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FECHACAUSA
        '
        Me.FECHACAUSA.Text = "Fecha Causa"
        Me.FECHACAUSA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FECHAEXPORT
        '
        Me.FECHAEXPORT.Text = "Fecha Exportación"
        Me.FECHAEXPORT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FACTURA
        '
        Me.FACTURA.Text = "Factura"
        Me.FACTURA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'MODO_GAR
        '
        Me.MODO_GAR.Text = "Modo Gar."
        Me.MODO_GAR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GRUPO_GAR
        '
        Me.GRUPO_GAR.Text = "Grupo Gar."
        Me.GRUPO_GAR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'stbEstado
        '
        Me.stbEstado.Location = New System.Drawing.Point(0, 582)
        Me.stbEstado.Name = "stbEstado"
        Me.stbEstado.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbPanel1, Me.sbPanel2, Me.sbPanel3})
        Me.stbEstado.ShowPanels = True
        Me.stbEstado.Size = New System.Drawing.Size(784, 22)
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
        Me.sbPanel3.Width = 317
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cbTodos)
        Me.GroupBox3.Controls.Add(Me.cbNinguno)
        Me.GroupBox3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox3.Location = New System.Drawing.Point(8, 496)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(152, 80)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Selección de pagos"
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
        Me.cbTodos.TabIndex = 20
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
        Me.cbNinguno.TabIndex = 21
        Me.cbNinguno.Tag = "NINGUNO"
        Me.cbNinguno.Text = "Ninguno"
        Me.cbNinguno.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbNinguno, "Cancela la selección")
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label14)
        Me.GroupBox4.Controls.Add(Me.Label13)
        Me.GroupBox4.Controls.Add(Me.Label12)
        Me.GroupBox4.Controls.Add(Me.lbTotal)
        Me.GroupBox4.Controls.Add(Me.lbIVA)
        Me.GroupBox4.Controls.Add(Me.lbImporte)
        Me.GroupBox4.Controls.Add(Me.lbResumenReferencias)
        Me.GroupBox4.Controls.Add(Me.lbResumenSiniestros)
        Me.GroupBox4.Controls.Add(Me.RTotal)
        Me.GroupBox4.Controls.Add(Me.RIva)
        Me.GroupBox4.Controls.Add(Me.RImporte)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Controls.Add(Me.Label10)
        Me.GroupBox4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox4.Location = New System.Drawing.Point(168, 496)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(280, 80)
        Me.GroupBox4.TabIndex = 8
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Resumen"
        '
        'Label14
        '
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(264, 49)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(9, 14)
        Me.Label14.TabIndex = 89
        Me.Label14.Text = "€"
        '
        'Label13
        '
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(264, 32)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(9, 16)
        Me.Label13.TabIndex = 88
        Me.Label13.Text = "€"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label12
        '
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(264, 17)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(9, 14)
        Me.Label12.TabIndex = 87
        Me.Label12.Text = "€"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lbTotal
        '
        Me.lbTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbTotal.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbTotal.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(0, Byte), CType(128, Byte))
        Me.lbTotal.Location = New System.Drawing.Point(160, 49)
        Me.lbTotal.Name = "lbTotal"
        Me.lbTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbTotal.Size = New System.Drawing.Size(96, 14)
        Me.lbTotal.TabIndex = 86
        Me.lbTotal.Text = "0,00"
        Me.lbTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lbIVA
        '
        Me.lbIVA.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbIVA.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbIVA.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(0, Byte), CType(128, Byte))
        Me.lbIVA.Location = New System.Drawing.Point(176, 32)
        Me.lbIVA.Name = "lbIVA"
        Me.lbIVA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbIVA.Size = New System.Drawing.Size(82, 16)
        Me.lbIVA.TabIndex = 85
        Me.lbIVA.Text = "0,00"
        Me.lbIVA.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbImporte
        '
        Me.lbImporte.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbImporte.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbImporte.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(0, Byte), CType(128, Byte))
        Me.lbImporte.Location = New System.Drawing.Point(176, 17)
        Me.lbImporte.Name = "lbImporte"
        Me.lbImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbImporte.Size = New System.Drawing.Size(82, 14)
        Me.lbImporte.TabIndex = 84
        Me.lbImporte.Text = "0,00"
        Me.lbImporte.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbResumenReferencias
        '
        Me.lbResumenReferencias.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbResumenReferencias.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbResumenReferencias.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(0, Byte), CType(128, Byte))
        Me.lbResumenReferencias.Location = New System.Drawing.Point(64, 32)
        Me.lbResumenReferencias.Name = "lbResumenReferencias"
        Me.lbResumenReferencias.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbResumenReferencias.Size = New System.Drawing.Size(47, 16)
        Me.lbResumenReferencias.TabIndex = 83
        Me.lbResumenReferencias.Text = "0,00"
        Me.lbResumenReferencias.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbResumenSiniestros
        '
        Me.lbResumenSiniestros.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbResumenSiniestros.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbResumenSiniestros.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(0, Byte), CType(128, Byte))
        Me.lbResumenSiniestros.Location = New System.Drawing.Point(64, 17)
        Me.lbResumenSiniestros.Name = "lbResumenSiniestros"
        Me.lbResumenSiniestros.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbResumenSiniestros.Size = New System.Drawing.Size(47, 14)
        Me.lbResumenSiniestros.TabIndex = 82
        Me.lbResumenSiniestros.Text = "0,00"
        Me.lbResumenSiniestros.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'RTotal
        '
        Me.RTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.RTotal.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.RTotal.ForeColor = System.Drawing.Color.RoyalBlue
        Me.RTotal.Location = New System.Drawing.Point(120, 49)
        Me.RTotal.Name = "RTotal"
        Me.RTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RTotal.Size = New System.Drawing.Size(53, 14)
        Me.RTotal.TabIndex = 81
        Me.RTotal.Text = "Total:"
        '
        'RIva
        '
        Me.RIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.RIva.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.RIva.ForeColor = System.Drawing.Color.RoyalBlue
        Me.RIva.Location = New System.Drawing.Point(120, 32)
        Me.RIva.Name = "RIva"
        Me.RIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RIva.Size = New System.Drawing.Size(53, 16)
        Me.RIva.TabIndex = 80
        Me.RIva.Text = "I.V.A.:"
        Me.RIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RImporte
        '
        Me.RImporte.Cursor = System.Windows.Forms.Cursors.Default
        Me.RImporte.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.RImporte.ForeColor = System.Drawing.Color.RoyalBlue
        Me.RImporte.Location = New System.Drawing.Point(120, 17)
        Me.RImporte.Name = "RImporte"
        Me.RImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RImporte.Size = New System.Drawing.Size(53, 14)
        Me.RImporte.TabIndex = 79
        Me.RImporte.Text = "Importe:"
        Me.RImporte.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label11.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label11.Location = New System.Drawing.Point(8, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(72, 16)
        Me.Label11.TabIndex = 78
        Me.Label11.Text = "Referencias:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label10
        '
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label10.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label10.Location = New System.Drawing.Point(8, 17)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(72, 14)
        Me.Label10.TabIndex = 77
        Me.Label10.Text = "Siniestros:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cbBorrar
        '
        Me.cbBorrar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBorrar.Image = CType(resources.GetObject("cbBorrar.Image"), System.Drawing.Image)
        Me.cbBorrar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbBorrar.Location = New System.Drawing.Point(456, 517)
        Me.cbBorrar.Name = "cbBorrar"
        Me.cbBorrar.Size = New System.Drawing.Size(64, 56)
        Me.cbBorrar.TabIndex = 22
        Me.cbBorrar.Text = "Borrar"
        Me.cbBorrar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbBorrar, "Elimina la selección actual de datos")
        '
        'cbPagos
        '
        Me.cbPagos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbPagos.Image = CType(resources.GetObject("cbPagos.Image"), System.Drawing.Image)
        Me.cbPagos.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbPagos.Location = New System.Drawing.Point(520, 517)
        Me.cbPagos.Name = "cbPagos"
        Me.cbPagos.Size = New System.Drawing.Size(64, 56)
        Me.cbPagos.TabIndex = 23
        Me.cbPagos.Text = "Pagos"
        Me.cbPagos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbPagos, "Ejecución Proceso de Pagos")
        '
        'cbAvisos
        '
        Me.cbAvisos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbAvisos.Image = CType(resources.GetObject("cbAvisos.Image"), System.Drawing.Image)
        Me.cbAvisos.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbAvisos.Location = New System.Drawing.Point(648, 517)
        Me.cbAvisos.Name = "cbAvisos"
        Me.cbAvisos.Size = New System.Drawing.Size(64, 56)
        Me.cbAvisos.TabIndex = 25
        Me.cbAvisos.Text = "Avisos"
        Me.cbAvisos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbAvisos, "Visualizar Histórico de Errores/Avisos")
        '
        'cbImprimir
        '
        Me.cbImprimir.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbImprimir.Image = CType(resources.GetObject("cbImprimir.Image"), System.Drawing.Image)
        Me.cbImprimir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbImprimir.Location = New System.Drawing.Point(584, 517)
        Me.cbImprimir.Name = "cbImprimir"
        Me.cbImprimir.Size = New System.Drawing.Size(64, 56)
        Me.cbImprimir.TabIndex = 24
        Me.cbImprimir.Text = "Imprimir"
        Me.cbImprimir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbImprimir, "Imprime la selección de pantalla")
        '
        'cbSalir
        '
        Me.cbSalir.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbSalir.Image = CType(resources.GetObject("cbSalir.Image"), System.Drawing.Image)
        Me.cbSalir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbSalir.Location = New System.Drawing.Point(712, 517)
        Me.cbSalir.Name = "cbSalir"
        Me.cbSalir.Size = New System.Drawing.Size(64, 56)
        Me.cbSalir.TabIndex = 26
        Me.cbSalir.Text = "Salir"
        Me.cbSalir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbSalir, "Cierra el Gestor de Pagos de Asistencia")
        '
        'lbxCompania
        '
        Me.lbxCompania.Location = New System.Drawing.Point(792, 16)
        Me.lbxCompania.Name = "lbxCompania"
        Me.lbxCompania.Size = New System.Drawing.Size(80, 95)
        Me.lbxCompania.TabIndex = 14
        '
        'lbxProducto
        '
        Me.lbxProducto.Location = New System.Drawing.Point(792, 120)
        Me.lbxProducto.Name = "lbxProducto"
        Me.lbxProducto.Size = New System.Drawing.Size(80, 95)
        Me.lbxProducto.TabIndex = 15
        '
        'prbProgreso
        '
        Me.prbProgreso.Location = New System.Drawing.Point(40, 592)
        Me.prbProgreso.Name = "prbProgreso"
        Me.prbProgreso.Size = New System.Drawing.Size(402, 8)
        Me.prbProgreso.TabIndex = 16
        Me.prbProgreso.Visible = False
        '
        'CR2
        '
        Me.CR2.Enabled = True
        Me.CR2.Location = New System.Drawing.Point(576, 456)
        Me.CR2.Name = "CR2"
        Me.CR2.OcxState = CType(resources.GetObject("CR2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CR2.Size = New System.Drawing.Size(28, 28)
        Me.CR2.TabIndex = 52
        Me.CR2.Visible = False
        '
        'picTest
        '
        Me.picTest.Cursor = System.Windows.Forms.Cursors.Help
        Me.picTest.Image = CType(resources.GetObject("picTest.Image"), System.Drawing.Image)
        Me.picTest.Location = New System.Drawing.Point(8, 8)
        Me.picTest.Name = "picTest"
        Me.picTest.Size = New System.Drawing.Size(48, 48)
        Me.picTest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picTest.TabIndex = 53
        Me.picTest.TabStop = False
        Me.picTest.Visible = False
        '
        'frmPrincipalPagos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(784, 604)
        Me.Controls.Add(Me.lvwPagos)
        Me.Controls.Add(Me.CR2)
        Me.Controls.Add(Me.prbProgreso)
        Me.Controls.Add(Me.lbxProducto)
        Me.Controls.Add(Me.lbxCompania)
        Me.Controls.Add(Me.cbSalir)
        Me.Controls.Add(Me.cbImprimir)
        Me.Controls.Add(Me.cbAvisos)
        Me.Controls.Add(Me.cbPagos)
        Me.Controls.Add(Me.cbBorrar)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.stbEstado)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.gbxFiltro)
        Me.Controls.Add(Me.cbxCompania)
        Me.Controls.Add(Me.lbCompaniaAsistencia)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.picTest)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrincipalPagos"
        Me.Text = "Siniestros: Área de Asistencia  -  Pagos Automáticos"
        Me.gbxFiltro.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.CR2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub frmPrincipalPagos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InicioApp()
    End Sub

    Private Sub InicioApp()
        On Error GoTo InicioApp_Err

        Dim parametroApp As String

        parametroApp = Microsoft.VisualBasic.Command
        strIdProceso = "P"

        If parametroApp = "PM" Then
            claseBDPagos.ConnexionPruebas()
            claseBDPagos.BDComand.CommandTimeout = 300
            picTest.Show()
            picTest.BringToFront()
        Else
            claseBDPagos.BDWorkConnect.CommandTimeout = 0
            claseBDPagos.BDComand.CommandTimeout = 0
            picTest.Hide()
            picTest.SendToBack()
        End If


        frmInstPagos = Me

        ' Valores globales
        '
        'CodUserApli = objUtiles.CodUser(UsuaApli)   ' Usuario de la aplicación
        If strCodUserAplicacion = "" Then strCodUserAplicacion = Microsoft.VisualBasic.Command()

        'PathIconos = clses.GetParam("PathGraficos") ' Ubicación de objetos gráficos
        strPathIconos = "K:\Graficos\"

        'Identificador de componente
        '
        strIDComp = "PAG"

        ' Inicialización ComboBox de Productos
        '
        If Not LlenarComboProducto(cbxProducto, lbxProducto, claseBDPagos, "TODOS") Then
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
        cbxTipoFecha.Items.Add("Importación")
        cbxTipoFecha.Items.Add("Pago Mutua")
        cbxTipoFecha.Text = "Importación"

        cbxTipoPago.Items.Add("Indemnización")
        cbxTipoPago.Items.Add("Gastos")
        cbxTipoPago.Text = "Indemnización"

        dtpDesde.Value = Now
        dtpHasta.Value = Now

        If strCodCia = "R" Then
            chkDtoReparalia.Visible = True
        End If

        bwflag = False

        '/* MUL INI filtro todos
        'FiltroTodos.PerformClick()
        '/* MUL FIN
        strFiltro = "T"

        ' Destrucción de objetos
        '
        'Set clses = Nothing

        ' En la carga del formulario efectuamos un cruce de referencias
        ' para actualizar el código de siniestro
        bVuelta = claseAsistenciaPagos.CruceReferenciasSiniestros("AS" + strIdReferCompa, strCodCia)
        Exit Sub

InicioApp_Err:
        'JCLopez_i
        MsgBox("Ha ocurrido un error iniciando la aplicación", MsgBoxStyle.Exclamation)
        End
        'objError.Tipo = Pantalla
        'objError.Ver(IdProceso, strGlobalNumErr, , Codcia)
        'JCLopez_f
    End Sub

    Private Sub cbBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBuscar.Click
        If cbxBusquedaAvanzada.Checked Then
            BusquedaAvanzada()
        Else
            If cbxCompania.Text = "" Then
                MsgBox("Debe escoger Producto y/o Compañía de Asistencia", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            BusquedaSimple()
        End If
        '/* MUL INI filtro todos
        'FiltroTodos.Checked = True
        '/* MUL FIN filtro todos
        bwflag = True
    End Sub

    Private Sub BusquedaAvanzada()
        On Error GoTo BusquedaAvanzada_Error

        If tbSiniestro.Text = "" And tbReferencia.Text = "" Then
            MsgBox("No se ha indicado ningún siniestro / referencia", MsgBoxStyle.Exclamation)
        ElseIf tbSiniestro.Text <> "" Then
            strCampoBuscaAvanzada = "T2_Codsin"
            strValorBuscaAvanzada = Trim(tbSiniestro.Text)
        Else
            strCampoBuscaAvanzada = "T2_Refer"
            strValorBuscaAvanzada = Trim(tbReferencia.Text)
        End If
        Call RefrescarGrid(dtpDesde.Value, dtpHasta.Value, cbxProducto.Text, strSigCompa)
        Exit Sub

BusquedaAvanzada_Error:
        MsgBox("Ha ocurrido un error realizando la búsqueda avanzada.", MsgBoxStyle.Critical)
    End Sub

    Private Sub BusquedaSimple()
        On Error GoTo BusquedaSimple_Error

        Dim dtFechaDesde, dtFechaHasta As Date

        dtFechaDesde = dtpDesde.Value
        dtFechaHasta = dtpHasta.Value

        strSigCompa = lbxCompania.Items.Item(cbxCompania.SelectedIndex)

        If cbxProducto.Text <> "Todos los Productos" Then
            strCodProducto = lbxProducto.Items.Item(cbxProducto.SelectedIndex)
        End If

        RefrescarGrid(dtFechaDesde, dtFechaHasta, strCodProducto, strSigCompa)
        Exit Sub
BusquedaSimple_Error:
        MsgBox("Ha ocurrido un error en la búsqueda", MsgBoxStyle.Critical)
    End Sub

    Public Function RefrescarGrid(ByRef FecDes As Date, ByRef FecHas As Date, ByRef strProducto As String, ByRef strCompa As String) As Short

        ' Declaraciones
        Dim strSQL As String                                ' Instrucción SQL entera
        Dim dtFecFiltro As Date                             ' Fecha a filtar
        Dim objListItem As ListViewItem                     ' Objeto con los registros del grid
        Dim objListSubItem As ListViewItem.ListViewSubItem  ' Objeto con las columnas del grid
        Dim strFrom As String                               ' Parte From de la SQL
        Dim strOrderBy As String                            ' Parte Order By de la SQL
        Dim numFila As Long

        ' Establecemos la selección de los campos con los que vamos a
        ' trabajar de la tabla de pagos de asistencia

        'T-3669 LMORILLA, Se toman las 10 primeras posiciones del fichero recibido
        'con la aprobación de asistenci y Cristina Pérez para que siempre coincida
        'con la tabla en la que graba.
        strSQLSel = "SELECT Compa = '" & strCodCia & "', " & _
                    "       Angel_t2.t2_refer, " & _
                    "       Angel_t2.t2_numord, " & _
                    "       Angel_t2.t2_fpago, " & _
                    "       Round((Angel_t2.t2_impor - Angel_t2.t2_imptva),2) as T2_IMPOR, " & _
                    "       Angel_t2.t2_ultpag, " & _
                    "       Left(Angel_t2.t2_poliza,10) as t2_poliza, " & _
                    "       Angel_t2.t2_causper, " & _
                    "       Angel_t2.t2_imptva, " & _
                    "       Angel_t2.t2_fgraba, " & _
                    "       Angel_t2.t2_codsin, " & _
                    "       Angel_t2.t2_pagado, " & _
                    "       Angel_t2.t2_tipgas, " & _
                    "       Angel_t2.t2_estado, " & _
                    "       Angel_t2.t2_festado, " & _
                    "       Situacion = Snsinies.estado, " & _
                    "       Perito = (select count(*) from snsinges  where snsinges.codsin = Angel_T2.T2_Codsin and snsinges.numper <> '" & strNumCompa & "'), " & _
                    "       Angel_t2.t2_Codram, " & _
                    "       Total = ISNULL(round(angel_t2.t2_impor,2),0), " & _
                    "       FechaProceso, " & _
                    "       Snsinies.Feccas, " & _
                    "       Fexport = (Select Min(Fservig) From Polizahist Where Polizahist.Numpol = Angel_T2.T2_Poliza and " & _
                    "                  Polizahist.Codram = Angel_T2.T2_Codram)," & _
                    "       Angel_T2.t2_factura, Angel_T2.t2_Codmod, Angel_T2.t2_Codgru"

        ' Si estamos realizando una busqueda avanzada ( por Siniestro o Referencia )
        ' no se tienen en cuenta los filtros ni ningún otro criterio de busqueda
        '
        If cbxBusquedaAvanzada.Checked Then
            strFrom = " From Angel_T2, Snsinies "
            strWhere = " Where Angel_T2." & strCampoBuscaAvanzada & " = '" & Trim(strValorBuscaAvanzada) & "' AND " & _
                       " (Angel_T2.T2_Codsin *= Snsinies.Codsin) and Angel_T2.T2_Codcia = '" & strCodCia & "'"
        Else
            ' Añadimos el From de la Sql
            '
            strFrom = " From Angel_t2, Snsinies "

            ' Añadimos la Where de la Sql
            '
            strWhere = " Where (Angel_T2.T2_Codsin *= Snsinies.Codsin) and Angel_T2.T2_Codcia ='" & strCodCia & "'"

            ' Añadimos la parte de la Where que filtrará los registros en
            ' función del tipo de fecha que hayamos seleccionado
            '
            Select Case cbxTipoFecha.Text
                Case "Importación"    ' Fecha en la que se importo el pago
                    strWhereMas = " And Angel_T2.T2_Fgraba BETWEEN '" & claseUtilidadesPagos.FormatoFechaSQL(FecDes, False, False) & "' AND '" & claseUtilidadesPagos.FormatoFechaSQL(FecHas, False, False) & "'"

                Case "Pago Mutua"          ' Fecha en la que Mutua realiza el pago
                    strWhereMas = " And Angel_T2.T2_Festado BETWEEN '" & claseUtilidadesPagos.FormatoFechaSQL(FecDes, False, False) & "' AND '" & claseUtilidadesPagos.FormatoFechaSQL(FecHas, False, False) & "'"
            End Select

            ' Añadimos la parte de la Where que filtrará los registro en
            ' función del tipo de pago que hayamos seleccionado.
            '
            Select Case cbxTipoPago.Text
                Case "Todos"          ' No se establece ningún filtro, entran todos los registros
                    strWhereMas = strWhereMas

                Case "Indemnización"  ' Sólo pagos de indemnización
                    strWhereMas = strWhereMas & " And Angel_T2.T2_tipgas = 'I'"

                Case "Gastos"         ' Sólo pagos de gatos
                    strWhereMas = strWhereMas & " And Angel_T2.T2_tipgas = 'G'"
            End Select

            ' Añadimos la parte de la Where que filtrará los registros en
            ' función del producto ( Codram ) seleccionado
            '
            ' Y dentro del producto en función del estado del pago
            '
            If strProducto = "Todos los Productos" Or strProducto = "" Then
                If strFiltro <> "T" Then
                    strWhereMas = strWhereMas & " And Angel_T2.T2_Estado = '" & strFiltro & "'"
                End If
            Else
                If strFiltro <> "T" Then
                    strWhereMas = strWhereMas & " And Angel_T2.T2_Codram = '" & strProducto & "' And Angel_T2.T2_ESTADO = '" & strFiltro & "'"
                End If
            End If

            ' Añadimos la parte de la Where que filtrará los registros en
            ' función de que sean o ono el último pago
            '
            If chkUltimoPago.Checked Then
                strWhereMas = strWhereMas & " And Angel_T2.T2_Ultpag = '1'"
            End If
        End If

        strOrderBy = " Order By Angel_T2.T2_Codsin, Angel_T2.T2_Refer"

        strSQL = strSQLSel & strFrom & strWhere & strWhereMas & strOrderBy

        ' Establece origen de datos para Crystal Reports
        '
        strSQLCR = strSQL

        ' /* MUL INI para optimizar al contar
        'Call CargarListView(lvwPagos, strSQL, "", "T2_CODSIN", "T2_REFER", "T2_FPAGO", "T2_FGRABA", "T2_ESTADO", "T2_FESTADO", "SITUACION", "T2_CAUSPER", "T2_IMPOR", "T2_IMPTVA", "TOTAL", "T2_PAGADO", "T2_TIPGAS", "T2_ULTPAG", "PERITO", "T2_POLIZA", "T2_CODRAM", "T2_NUMORD", "FECHAPROCESO", "FECCAS", "FEXPORT", "T2_FACTURA", "T2_CODMOD", "T2_CODGRU")
        Call CargarListView_pagos(lvwPagos, strSQLSel, strFrom & strWhere & strWhereMas, strOrderBy, "", _
                                  "T2_CODSIN", "T2_REFER", "T2_FPAGO", "T2_FGRABA", "T2_ESTADO", "T2_FESTADO", "SITUACION", "T2_CAUSPER", "T2_IMPOR", "T2_IMPTVA", "TOTAL", "T2_PAGADO", "T2_TIPGAS", "T2_ULTPAG", "PERITO", "T2_POLIZA", "T2_CODRAM", "T2_NUMORD", "FECHAPROCESO", "FECCAS", "FEXPORT", "T2_FACTURA", "T2_CODMOD", "T2_CODGRU")
        '/* MUL FIN
        stbEstado.Panels(2).Text = CStr(Format(lvwPagos.Items.Count, "##,##0")) & " Pagos"

        ' Poner atenuados aquellos siniestros que ya esten aperturados.
        ' o con el color identificativo para aquellos que tengan errores o avisos
        If lvwPagos.Items.Count > 0 Then
            numFila = 0
            For Each objListItem In lvwPagos.Items
                objListSubItem = objListItem.SubItems.Item(frmInstPagos.T2_ESTADO.Index)

                Select Case objListSubItem.Text
                    Case "P"
                        objListItem.Tag = "1"
                        'ttipAyuda.SetToolTip(lvwPagos.Items.Item(numFila)., "Este Siniestro ya está Procesado")
                        Call ColorListItem(objListItem, Color.Gray)
                    Case "X"
                        objListItem.Tag = "0"
                        'ttipAyuda.SetToolTip(objListItem.ListView, "Este Siniestro está pendiente de procesar")
                        Call ColorListItem(objListItem, Color.Green)
                    Case "W", "A"
                        objListItem.Tag = "0"
                        'ttipAyuda.SetToolTip(objListItem.ListView, "Este Siniestro tiene avisos pendientes de resolución")
                        Call ColorListItem(objListItem, Color.DarkGoldenrod)
                    Case "E"
                        objListItem.Tag = "0"
                        'ttipAyuda.SetToolTip(objListItem.ListView, "Este Siniestro tiene mensajes de error del proceso")
                        Call ColorListItem(objListItem, Color.Red)
                    Case Else
                        objListItem.Tag = "0"
                        Call ColorListItem(objListItem, Color.Black)
                End Select

                'JCLopez_i
                ' En la columna de ültimo Pago substituimos el 1 o el 0
                ' por la S o la N
                If Int(objListItem.SubItems.Item(frmInstPagos.T2_ULTPAG.Index).Text) = 1 Then
                    objListItem.SubItems.Item(frmInstPagos.T2_ULTPAG.Index).Text = "S"
                Else
                    objListItem.SubItems.Item(frmInstPagos.T2_ULTPAG.Index).Text = "N"
                End If

                ' En la columna de Perito substituimos el 0 o el >=1
                ' por la S o la N.
                '
                If objListItem.SubItems.Item(frmInstPagos.PERITO.Index).Text <> "0" Then
                    objListItem.SubItems.Item(frmInstPagos.T2_ULTPAG.Index).Text = "S"
                Else
                    objListItem.SubItems.Item(frmInstPagos.T2_ULTPAG.Index).Text = "N"
                End If
                'JCLopez_f
            Next
        End If
        Exit Function
RefrescarGrid_Error:
        MsgBox("Error refrescando datos", MsgBoxStyle.Critical)
    End Function


    Private Sub cbxCompania_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxCompania.SelectedIndexChanged

        strCodCia = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))
        'Label2.Text = strCodCia

        If DatosCiaAsistencia(strCodCia) Then
            'cbxCompania.Visible = False
            If strCodCia = "R" Then
                chkDtoReparalia.Visible = True
            Else
                chkDtoReparalia.Visible = False
            End If
        Else
            MsgBox("No hay datos", MsgBoxStyle.Exclamation)
        End If

        'cbxCompania_Click_Err:
        '        MsgBox("Se ha producido un error al intentar crear la lista de selección. Por favor avise a Informática.", MsgBoxStyle.Exclamation)

    End Sub

    Private Sub cbSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSalir.Click
        Application.Exit()
    End Sub

    Private Sub cbBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBorrar.Click
        On Error GoTo BorrarError

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        'Dim lstrSQL As String
        Dim objPagos As New clsPagos_NET
        Dim objlista As ListViewItem
        Dim retornoMsgBox As MsgBoxResult
        ' Creación de objetos
        '
        Me.Cursor = Cursors.Hand

        ' Confirmación de la orden de eliminación de registros
        retornoMsgBox = MsgBox("¿Esta seguro de querer eliminar los datos seleccionados?", MsgBoxStyle.YesNo)

        'objError.Ver(IdProceso, gstrError, " ¿ Esta seguro de querer eliminar los datos seleccionados ? ", Codcia)

        If retornoMsgBox = MsgBoxResult.Yes Then
            For Each objlistitem In lvwPagos.Items
                Call ActualizarPorcentaje(lvwPagos.Items.Count, prbProgreso, stbEstado)
                If objlistitem.Checked Then

                    objPagos.Referencias.Add(objlistitem.Text)
                    objPagos.NumeroOrden.Add(objlistitem.SubItems(frmInstPagos.T2_NUMORD.Index).Text)
                End If
            Next
        End If

        If Not IsNothing(objPagos) Then
            If objPagos.Referencias.Count > 0 Then
                If objPagos.DeletePagos(objPagos.Referencias, objPagos.NumeroOrden) Then
                    MsgBox("La eliminación de los registros seleccionados se ha procesado correctamente", MsgBoxStyle.Information)
                    'gstrError = "108"
                    'objError.Tipo = Pantalla
                    'objError.Ver(IdProceso, gstrError, , Codcia)
                Else
                    Err.Raise(1)
                End If
            End If
        End If

        stbEstado.Panels(2).Text = ""
        prbProgreso.Visible = False
        Call RefrescarGrid(dtpDesde.Value, dtpHasta.Value, cbxProducto.Text, strSigCompa)
        Me.Cursor = Cursors.Hand
        Exit Sub

BorrarError:
        'objPagos = Nothing
        Me.Cursor = Cursors.Default
        MsgBox("Se ha producido un error en la Base de Datos. Los registros seleccionados no han sido borrados", MsgBoxStyle.Exclamation)
        'gstrError = "107"
        'objError.Tipo = Pantalla
        'objError.Ver(IdProceso, gstrError, , Codcia)

    End Sub

    Private Sub cbPagos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPagos.Click
        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim objPagar As New clsPagos_NET
        Dim strRefer As String
        Dim vntReferencia As Object

        prbProgreso.Minimum = 0
        prbProgreso.Value = 1
        prbProgreso.Maximum = lvwPagos.Items.Count


        stbEstado.Panels(2).Text = "Comprobando Filtros ..."

        For Each objlistitem In lvwPagos.Items
            Call ActualizarPorcentaje(lvwPagos.Items.Count, prbProgreso, stbEstado)
            If objlistitem.Checked And objlistitem.SubItems.Item(frmInstPagos.T2_ESTADO.Index).Text <> "P" Then
                If chkFiltroAvisos.Checked And objlistitem.SubItems.Item(frmInstPagos.T2_ESTADO.Index).Text = "A" Then
                    objPagar.Referencias.Add(objlistitem.Text)
                ElseIf objPagar.Filtros(objlistitem) Then
                    objPagar.Referencias.Add(objlistitem.Text)
                End If

                stbEstado.Panels(2).Text = "Realizando Pagos de Siniestros ..."
                prbProgreso.Minimum = 0
                prbProgreso.Value = 0
                If Not IsNothing(objPagar) Then
                    If objPagar.Referencias.Count > 0 Then
                        prbProgreso.Maximum = objPagar.Referencias.Count
                        Call ActualizarPorcentaje(lvwPagos.Items.Count, prbProgreso, stbEstado)
                        If Not objlistitem Is Nothing Then Call objPagar.Pagar(objlistitem)
                        objPagar.Referencias.Remove(1)
                    End If
                End If
            End If
        Next

        stbEstado.Panels(2).Text = ""
        prbProgreso.Visible = False

        strSigCompa = lbxCompania.Items.Item(cbxCompania.SelectedIndex)
        'Option1(0).Value = True
        '/* MUL INI filtro todos    
        'FiltroTodos.Checked = True
        'strFiltro = "T"
        '/* MUL FIN filtro todos
        Call RefrescarGrid(dtpDesde.Value, dtpHasta.Value, cbxProducto.Text, strSigCompa)
    End Sub


    Private Sub lvwPagos_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lvwPagos.ItemCheck
        Dim lvPagosAux As ListView

        lvPagosAux = sender
        If lvPagosAux.Items(e.Index).Tag = "1" Then
            e.NewValue = e.CurrentValue
        End If
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

    Private Sub cbTodos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTodos.Click, cbNinguno.Click
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

        If strOpcion = "TODOS" Then
            blnCheck = True
        ElseIf strOpcion = "NINGUNO" Then
            blnCheck = False
        Else
            blnCheck = False
        End If

        ' Seleccionar o deseleccionar todos los elementos de la lista
        ' Teniendo en cuenta que aquellos siniestros esten aperturados
        ' no se pueden seleccionar
        If lvwPagos.Items.Count > 0 Then
            For Each objlistitem In lvwPagos.Items
                If objlistitem.Tag <> "1" Then
                    If objlistitem.Text <> "" And objlistitem.Text <> "No Existe" And objlistitem.SubItems.Item(frmInstPagos.SITUACION.Index).Text <> "C" Then
                        Call ColorListItem(objlistitem, Color.Green)
                        objlistitem.Checked = blnCheck
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub cbxBusquedaAvanzada_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxBusquedaAvanzada.CheckedChanged

        Dim boolActivo As Boolean
        boolActivo = cbxBusquedaAvanzada.Checked
        'Habilitados
        lbSiniestro.Enabled = boolActivo
        lbReferencia.Enabled = boolActivo
        tbReferencia.Enabled = boolActivo
        tbSiniestro.Enabled = boolActivo


        'Deshabilitados
        FiltroAviso.Enabled = Not boolActivo
        FiltroErrores.Enabled = Not boolActivo
        FiltroNoPagados.Enabled = Not boolActivo
        FiltroPagados.Enabled = Not boolActivo
        FiltroTodos.Enabled = Not boolActivo

        dtpDesde.Enabled = Not boolActivo
        dtpHasta.Enabled = Not boolActivo
        cbxTipoFecha.Enabled = Not boolActivo
        cbxTipoPago.Enabled = Not boolActivo
        lbFechaDesde.Enabled = Not boolActivo
        lbFechaHasta.Enabled = Not boolActivo
        cbxProducto.Enabled = Not boolActivo
        cbxTipoFecha.Enabled = Not boolActivo
        cbxTipoPago.Enabled = Not boolActivo
        chkFiltroAvisos.Enabled = Not boolActivo
        chkUltimoPago.Enabled = Not boolActivo
    End Sub

    Private Sub cbAvisos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAvisos.Click
        On Error GoTo cbAvisos_Error

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim strReferencia As String
        Dim frmInstanciaErrores As New frmVisorErrores

        ' Comprobamos que se haya seleccionado
        If lvwPagos.CheckedItems.Count > 0 Then
            ' Comprobamos el estado de la referencia seleccionada
            If lvwPagos.CheckedItems(0).SubItems(frmInstPagos.T2_ESTADO.Index).Text = "W" Or _
               lvwPagos.CheckedItems(0).SubItems(frmInstPagos.T2_ESTADO.Index).Text = "A" Or _
               lvwPagos.CheckedItems(0).SubItems(frmInstPagos.T2_ESTADO.Index).Text = "E" Then
                If lvwPagos.CheckedItems(0).Tag = "1" Then Exit Sub

                frmInstanciaErrores.Show()
                objlistitem = lvwPagos.CheckedItems(0)
                strReferencia = lvwPagos.CheckedItems(0).SubItems(frmInstPagos.T2_REFER.Index).Text
                frmInstanciaErrores.MostrarErrores(strReferencia)
            ElseIf lvwPagos.CheckedItems(0).SubItems(frmInstPagos.T2_ESTADO.Index).Text = "P" Then
                claseBDPagos.BDAuxRecord = claseSiniestroPagos.Siniestro(lvwPagos.CheckedItems(0).SubItems(frmInstPagos.T2_CODSIN.Index).Text, True, strUsuarioAplicacion)
            Else
                MsgBox("La referencia seleccionada no tiene Avisos/Errores", MsgBoxStyle.Information)
            End If
        Else
            '/* MUL  si no se selecciona ninguno se muestra el historial.
            strError = "4007"
            'objError.Tipo = Pantalla
            'objError.Ver(IdProceso, gstrError, , Codcia)
            frmInstanciaErrores.dtpFechaInicio.Value = dtpDesde.Value
            frmInstanciaErrores.dtpFechaFin.Value = dtpHasta.Value
            frmInstanciaErrores.MostrarErrores("")
            frmInstanciaErrores.Show()
        End If

        Exit Sub

cbAvisos_Error:
        MsgBox("Ha ocurrido un error mostrando el aviso", MsgBoxStyle.Critical)
    End Sub

    Private Sub cbImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbImprimir.Click
        Imprimir_pagos()
    End Sub

End Class
