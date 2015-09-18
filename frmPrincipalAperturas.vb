Public Class frmPrincipalAperturas
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
    Friend WithEvents cbBuscar As System.Windows.Forms.Button
    Friend WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkFiltroAvisos As System.Windows.Forms.CheckBox
    Friend WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Friend WithEvents stbEstado As System.Windows.Forms.StatusBar
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lbxCompania As System.Windows.Forms.ListBox
    Friend WithEvents lbxProducto As System.Windows.Forms.ListBox
    Friend WithEvents FiltroErrores As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroAviso As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroNoPagados As System.Windows.Forms.RadioButton
    Friend WithEvents FiltroTodos As System.Windows.Forms.RadioButton
    Friend WithEvents gbxFiltro As System.Windows.Forms.GroupBox
    Friend WithEvents ttipAyuda As System.Windows.Forms.ToolTip
    Friend WithEvents cbSalir As System.Windows.Forms.Button
    Friend WithEvents cbBorrar As System.Windows.Forms.Button
    Friend WithEvents cbAvisos As System.Windows.Forms.Button
    Friend WithEvents cbImprimir As System.Windows.Forms.Button
    Friend WithEvents cbTodos As System.Windows.Forms.Button
    Friend WithEvents cbNinguno As System.Windows.Forms.Button
    Friend WithEvents prbProgreso As System.Windows.Forms.ProgressBar
    ' Friend WithEvents CR2 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents sbPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents cbxCompania As System.Windows.Forms.ComboBox
    Friend WithEvents lbFechaHasta As System.Windows.Forms.Label
    Friend WithEvents lbFechaDesde As System.Windows.Forms.Label
    Friend WithEvents lbCompaniaAsistencia As System.Windows.Forms.Label
    Friend WithEvents lvwAperturas As System.Windows.Forms.ListView
    Friend WithEvents T1_REFER As System.Windows.Forms.ColumnHeader
    Friend WithEvents T1_CIA As System.Windows.Forms.ColumnHeader
    Friend WithEvents T1_CODSIN As System.Windows.Forms.ColumnHeader
    Friend WithEvents T1_POLIZA As System.Windows.Forms.ColumnHeader
    Friend WithEvents AVISO As System.Windows.Forms.ColumnHeader
    Friend WithEvents T1_FAPER As System.Windows.Forms.ColumnHeader
    Friend WithEvents T1_FSINI As System.Windows.Forms.ColumnHeader
    Friend WithEvents T1_DESCR As System.Windows.Forms.ColumnHeader
    Friend WithEvents T3_CCAUSA As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECGRA As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECHAPROCESO As System.Windows.Forms.ColumnHeader
    Friend WithEvents T1_PERTURAPOR As System.Windows.Forms.ColumnHeader
    Friend WithEvents cbxProducto As System.Windows.Forms.ComboBox
    Friend WithEvents lbProducto As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents tbSiniestro As System.Windows.Forms.TextBox
    Friend WithEvents tbReferencia As System.Windows.Forms.TextBox
    Friend WithEvents lbSiniestro As System.Windows.Forms.Label
    Friend WithEvents lbReferencia As System.Windows.Forms.Label
    Friend WithEvents cbxBusquedaAvanzada As System.Windows.Forms.CheckBox
    Friend WithEvents T1_ESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents cbAperturar As System.Windows.Forms.Button
    Friend WithEvents picTest As System.Windows.Forms.PictureBox
    Public WithEvents CR2 As AxCrystal.AxCrystalReport
    Friend WithEvents FiltroPagados As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrincipalAperturas))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lbCompaniaAsistencia = New System.Windows.Forms.Label
        Me.cbxCompania = New System.Windows.Forms.ComboBox
        Me.gbxFiltro = New System.Windows.Forms.GroupBox
        Me.chkFiltroAvisos = New System.Windows.Forms.CheckBox
        Me.FiltroErrores = New System.Windows.Forms.RadioButton
        Me.FiltroAviso = New System.Windows.Forms.RadioButton
        Me.FiltroNoPagados = New System.Windows.Forms.RadioButton
        Me.FiltroPagados = New System.Windows.Forms.RadioButton
        Me.FiltroTodos = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.tbSiniestro = New System.Windows.Forms.TextBox
        Me.tbReferencia = New System.Windows.Forms.TextBox
        Me.lbSiniestro = New System.Windows.Forms.Label
        Me.lbReferencia = New System.Windows.Forms.Label
        Me.cbxBusquedaAvanzada = New System.Windows.Forms.CheckBox
        Me.cbxProducto = New System.Windows.Forms.ComboBox
        Me.lbProducto = New System.Windows.Forms.Label
        Me.cbBuscar = New System.Windows.Forms.Button
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker
        Me.lbFechaHasta = New System.Windows.Forms.Label
        Me.lbFechaDesde = New System.Windows.Forms.Label
        Me.lvwAperturas = New System.Windows.Forms.ListView
        Me.T1_REFER = New System.Windows.Forms.ColumnHeader
        Me.T1_CIA = New System.Windows.Forms.ColumnHeader
        Me.T1_CODSIN = New System.Windows.Forms.ColumnHeader
        Me.T1_POLIZA = New System.Windows.Forms.ColumnHeader
        Me.AVISO = New System.Windows.Forms.ColumnHeader
        Me.T1_FAPER = New System.Windows.Forms.ColumnHeader
        Me.T1_FSINI = New System.Windows.Forms.ColumnHeader
        Me.T1_DESCR = New System.Windows.Forms.ColumnHeader
        Me.T3_CCAUSA = New System.Windows.Forms.ColumnHeader
        Me.FECGRA = New System.Windows.Forms.ColumnHeader
        Me.FECHAPROCESO = New System.Windows.Forms.ColumnHeader
        Me.T1_PERTURAPOR = New System.Windows.Forms.ColumnHeader
        Me.T1_ESTADO = New System.Windows.Forms.ColumnHeader
        Me.stbEstado = New System.Windows.Forms.StatusBar
        Me.sbPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cbTodos = New System.Windows.Forms.Button
        Me.cbNinguno = New System.Windows.Forms.Button
        Me.cbBorrar = New System.Windows.Forms.Button
        Me.cbAperturar = New System.Windows.Forms.Button
        Me.cbAvisos = New System.Windows.Forms.Button
        Me.cbImprimir = New System.Windows.Forms.Button
        Me.cbSalir = New System.Windows.Forms.Button
        Me.lbxCompania = New System.Windows.Forms.ListBox
        Me.lbxProducto = New System.Windows.Forms.ListBox
        Me.ttipAyuda = New System.Windows.Forms.ToolTip(Me.components)
        Me.picTest = New System.Windows.Forms.PictureBox
        Me.prbProgreso = New System.Windows.Forms.ProgressBar
        Me.CR2 = New AxCrystal.AxCrystalReport
        Me.gbxFiltro.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
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
        Me.cbxCompania.Size = New System.Drawing.Size(336, 24)
        Me.cbxCompania.TabIndex = 2
        '
        'gbxFiltro
        '
        Me.gbxFiltro.Controls.Add(Me.chkFiltroAvisos)
        Me.gbxFiltro.Controls.Add(Me.FiltroErrores)
        Me.gbxFiltro.Controls.Add(Me.FiltroAviso)
        Me.gbxFiltro.Controls.Add(Me.FiltroNoPagados)
        Me.gbxFiltro.Controls.Add(Me.FiltroPagados)
        Me.gbxFiltro.Controls.Add(Me.FiltroTodos)
        Me.gbxFiltro.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxFiltro.ForeColor = System.Drawing.Color.RoyalBlue
        Me.gbxFiltro.Location = New System.Drawing.Point(568, 0)
        Me.gbxFiltro.Name = "gbxFiltro"
        Me.gbxFiltro.Size = New System.Drawing.Size(192, 200)
        Me.gbxFiltro.TabIndex = 3
        Me.gbxFiltro.TabStop = False
        Me.gbxFiltro.Text = "Filtro"
        Me.ttipAyuda.SetToolTip(Me.gbxFiltro, "Filtros sobre el resultado de la consulta")
        '
        'chkFiltroAvisos
        '
        Me.chkFiltroAvisos.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.chkFiltroAvisos.Location = New System.Drawing.Point(16, 152)
        Me.chkFiltroAvisos.Name = "chkFiltroAvisos"
        Me.chkFiltroAvisos.Size = New System.Drawing.Size(168, 24)
        Me.chkFiltroAvisos.TabIndex = 6
        Me.chkFiltroAvisos.Text = "Desactivar filtros para avisos"
        Me.ttipAyuda.SetToolTip(Me.chkFiltroAvisos, "Desactiva los Filtros para procesar los Avisos")
        '
        'FiltroErrores
        '
        Me.FiltroErrores.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroErrores.ForeColor = System.Drawing.Color.Red
        Me.FiltroErrores.Location = New System.Drawing.Point(40, 120)
        Me.FiltroErrores.Name = "FiltroErrores"
        Me.FiltroErrores.Size = New System.Drawing.Size(104, 16)
        Me.FiltroErrores.TabIndex = 4
        Me.FiltroErrores.Text = "(E) Errores"
        Me.ttipAyuda.SetToolTip(Me.FiltroErrores, "Muestra sólo los registros en los que se ha producido error")
        '
        'FiltroAviso
        '
        Me.FiltroAviso.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroAviso.ForeColor = System.Drawing.Color.DarkGoldenrod
        Me.FiltroAviso.Location = New System.Drawing.Point(40, 96)
        Me.FiltroAviso.Name = "FiltroAviso"
        Me.FiltroAviso.Size = New System.Drawing.Size(104, 16)
        Me.FiltroAviso.TabIndex = 3
        Me.FiltroAviso.Text = "(A) Avisos"
        Me.ttipAyuda.SetToolTip(Me.FiltroAviso, "Muestra sólo los registros con mensaje de aviso")
        '
        'FiltroNoPagados
        '
        Me.FiltroNoPagados.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroNoPagados.ForeColor = System.Drawing.Color.Green
        Me.FiltroNoPagados.Location = New System.Drawing.Point(40, 72)
        Me.FiltroNoPagados.Name = "FiltroNoPagados"
        Me.FiltroNoPagados.Size = New System.Drawing.Size(128, 16)
        Me.FiltroNoPagados.TabIndex = 2
        Me.FiltroNoPagados.Text = "(X) No Aperturados"
        Me.ttipAyuda.SetToolTip(Me.FiltroNoPagados, "Muestra sólo los pendientes de aperturar")
        '
        'FiltroPagados
        '
        Me.FiltroPagados.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroPagados.ForeColor = System.Drawing.Color.Gray
        Me.FiltroPagados.Location = New System.Drawing.Point(40, 48)
        Me.FiltroPagados.Name = "FiltroPagados"
        Me.FiltroPagados.Size = New System.Drawing.Size(104, 16)
        Me.FiltroPagados.TabIndex = 1
        Me.FiltroPagados.Text = "(A) Aperturados "
        Me.ttipAyuda.SetToolTip(Me.FiltroPagados, "Muestra sólo los aperturados")
        '
        'FiltroTodos
        '
        Me.FiltroTodos.Checked = True
        Me.FiltroTodos.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroTodos.ForeColor = System.Drawing.Color.Black
        Me.FiltroTodos.Location = New System.Drawing.Point(40, 24)
        Me.FiltroTodos.Name = "FiltroTodos"
        Me.FiltroTodos.Size = New System.Drawing.Size(104, 16)
        Me.FiltroTodos.TabIndex = 0
        Me.FiltroTodos.TabStop = True
        Me.FiltroTodos.Text = "Todos"
        Me.ttipAyuda.SetToolTip(Me.FiltroTodos, "Muestra todos los registros")
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Panel1)
        Me.GroupBox2.Controls.Add(Me.cbxProducto)
        Me.GroupBox2.Controls.Add(Me.lbProducto)
        Me.GroupBox2.Controls.Add(Me.cbBuscar)
        Me.GroupBox2.Controls.Add(Me.dtpHasta)
        Me.GroupBox2.Controls.Add(Me.dtpDesde)
        Me.GroupBox2.Controls.Add(Me.lbFechaHasta)
        Me.GroupBox2.Controls.Add(Me.lbFechaDesde)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox2.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(552, 136)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Criterios de Selección"
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
        Me.Panel1.Location = New System.Drawing.Point(8, 88)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(536, 40)
        Me.Panel1.TabIndex = 14
        '
        'tbSiniestro
        '
        Me.tbSiniestro.Enabled = False
        Me.tbSiniestro.Location = New System.Drawing.Point(392, 8)
        Me.tbSiniestro.Name = "tbSiniestro"
        Me.tbSiniestro.Size = New System.Drawing.Size(88, 21)
        Me.tbSiniestro.TabIndex = 4
        Me.tbSiniestro.Text = ""
        '
        'tbReferencia
        '
        Me.tbReferencia.Enabled = False
        Me.tbReferencia.Location = New System.Drawing.Point(232, 8)
        Me.tbReferencia.Name = "tbReferencia"
        Me.tbReferencia.Size = New System.Drawing.Size(88, 21)
        Me.tbReferencia.TabIndex = 3
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
        Me.cbxBusquedaAvanzada.TabIndex = 0
        Me.cbxBusquedaAvanzada.Text = "Búsqueda avanzada"
        '
        'cbxProducto
        '
        Me.cbxProducto.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbxProducto.Location = New System.Drawing.Point(296, 32)
        Me.cbxProducto.Name = "cbxProducto"
        Me.cbxProducto.Size = New System.Drawing.Size(160, 21)
        Me.cbxProducto.TabIndex = 13
        '
        'lbProducto
        '
        Me.lbProducto.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbProducto.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbProducto.Location = New System.Drawing.Point(232, 32)
        Me.lbProducto.Name = "lbProducto"
        Me.lbProducto.Size = New System.Drawing.Size(56, 16)
        Me.lbProducto.TabIndex = 12
        Me.lbProducto.Text = "Producto:"
        Me.lbProducto.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbBuscar
        '
        Me.cbBuscar.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbBuscar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBuscar.Image = CType(resources.GetObject("cbBuscar.Image"), System.Drawing.Image)
        Me.cbBuscar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbBuscar.Location = New System.Drawing.Point(480, 24)
        Me.cbBuscar.Name = "cbBuscar"
        Me.cbBuscar.Size = New System.Drawing.Size(64, 56)
        Me.cbBuscar.TabIndex = 11
        Me.cbBuscar.Text = "Buscar"
        Me.cbBuscar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbBuscar, "Ejecuta la busqueda de Pagos según los criterios asignados")
        '
        'dtpHasta
        '
        Me.dtpHasta.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpHasta.Location = New System.Drawing.Point(96, 62)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(104, 21)
        Me.dtpHasta.TabIndex = 7
        Me.ttipAyuda.SetToolTip(Me.dtpHasta, "Abre el calendario de selección")
        '
        'dtpDesde
        '
        Me.dtpDesde.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDesde.Location = New System.Drawing.Point(96, 30)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(104, 21)
        Me.dtpDesde.TabIndex = 6
        Me.ttipAyuda.SetToolTip(Me.dtpDesde, "Abre el calendario de selección")
        '
        'lbFechaHasta
        '
        Me.lbFechaHasta.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbFechaHasta.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lbFechaHasta.Location = New System.Drawing.Point(16, 64)
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
        Me.lbFechaDesde.Location = New System.Drawing.Point(16, 32)
        Me.lbFechaDesde.Name = "lbFechaDesde"
        Me.lbFechaDesde.Size = New System.Drawing.Size(72, 16)
        Me.lbFechaDesde.TabIndex = 1
        Me.lbFechaDesde.Text = "Fecha desde:"
        Me.lbFechaDesde.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lvwAperturas
        '
        Me.lvwAperturas.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvwAperturas.CheckBoxes = True
        Me.lvwAperturas.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.T1_REFER, Me.T1_CIA, Me.T1_CODSIN, Me.T1_POLIZA, Me.AVISO, Me.T1_FAPER, Me.T1_FSINI, Me.T1_DESCR, Me.T3_CCAUSA, Me.FECGRA, Me.FECHAPROCESO, Me.T1_PERTURAPOR, Me.T1_ESTADO})
        Me.lvwAperturas.FullRowSelect = True
        Me.lvwAperturas.GridLines = True
        Me.lvwAperturas.Location = New System.Drawing.Point(8, 208)
        Me.lvwAperturas.Name = "lvwAperturas"
        Me.lvwAperturas.Size = New System.Drawing.Size(768, 256)
        Me.lvwAperturas.TabIndex = 5
        Me.lvwAperturas.View = System.Windows.Forms.View.Details
        '
        'T1_REFER
        '
        Me.T1_REFER.Text = "Referencia"
        Me.T1_REFER.Width = 83
        '
        'T1_CIA
        '
        Me.T1_CIA.Text = "Ramo"
        Me.T1_CIA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T1_CIA.Width = 50
        '
        'T1_CODSIN
        '
        Me.T1_CODSIN.Text = "Siniestro"
        Me.T1_CODSIN.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T1_CODSIN.Width = 74
        '
        'T1_POLIZA
        '
        Me.T1_POLIZA.Text = "Póliza"
        Me.T1_POLIZA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T1_POLIZA.Width = 75
        '
        'AVISO
        '
        Me.AVISO.Text = "Aviso/Error"
        Me.AVISO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.AVISO.Width = 68
        '
        'T1_FAPER
        '
        Me.T1_FAPER.Text = "F.Apertura"
        Me.T1_FAPER.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T1_FAPER.Width = 73
        '
        'T1_FSINI
        '
        Me.T1_FSINI.Text = "F.Siniestro"
        Me.T1_FSINI.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T1_FSINI.Width = 71
        '
        'T1_DESCR
        '
        Me.T1_DESCR.Text = "Descripción"
        Me.T1_DESCR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T1_DESCR.Width = 110
        '
        'T3_CCAUSA
        '
        Me.T3_CCAUSA.Text = "Cód. Causa"
        Me.T3_CCAUSA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T3_CCAUSA.Width = 70
        '
        'FECGRA
        '
        Me.FECGRA.Text = "F.Importación"
        Me.FECGRA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.FECGRA.Width = 71
        '
        'FECHAPROCESO
        '
        Me.FECHAPROCESO.Text = "Fecha Proceso"
        '
        'T1_PERTURAPOR
        '
        Me.T1_PERTURAPOR.Text = "Perito"
        '
        'T1_ESTADO
        '
        Me.T1_ESTADO.Text = "Estado"
        '
        'stbEstado
        '
        Me.stbEstado.Location = New System.Drawing.Point(0, 556)
        Me.stbEstado.Name = "stbEstado"
        Me.stbEstado.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbPanel1, Me.sbPanel2, Me.sbPanel3})
        Me.stbEstado.ShowPanels = True
        Me.stbEstado.Size = New System.Drawing.Size(786, 24)
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
        Me.sbPanel3.Width = 319
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cbTodos)
        Me.GroupBox3.Controls.Add(Me.cbNinguno)
        Me.GroupBox3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox3.Location = New System.Drawing.Point(8, 472)
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
        'cbBorrar
        '
        Me.cbBorrar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBorrar.Image = CType(resources.GetObject("cbBorrar.Image"), System.Drawing.Image)
        Me.cbBorrar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbBorrar.Location = New System.Drawing.Point(456, 488)
        Me.cbBorrar.Name = "cbBorrar"
        Me.cbBorrar.Size = New System.Drawing.Size(64, 56)
        Me.cbBorrar.TabIndex = 9
        Me.cbBorrar.Text = "Borrar"
        Me.cbBorrar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbBorrar, "Elimina la selección actual de datos")
        '
        'cbAperturar
        '
        Me.cbAperturar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbAperturar.Image = CType(resources.GetObject("cbAperturar.Image"), System.Drawing.Image)
        Me.cbAperturar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbAperturar.Location = New System.Drawing.Point(520, 488)
        Me.cbAperturar.Name = "cbAperturar"
        Me.cbAperturar.Size = New System.Drawing.Size(64, 56)
        Me.cbAperturar.TabIndex = 10
        Me.cbAperturar.Text = "Aperturar"
        Me.cbAperturar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbAperturar, "Ejecución Proceso de Pagos")
        '
        'cbAvisos
        '
        Me.cbAvisos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbAvisos.Image = CType(resources.GetObject("cbAvisos.Image"), System.Drawing.Image)
        Me.cbAvisos.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbAvisos.Location = New System.Drawing.Point(648, 488)
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
        Me.cbImprimir.Location = New System.Drawing.Point(584, 488)
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
        Me.cbSalir.Location = New System.Drawing.Point(712, 488)
        Me.cbSalir.Name = "cbSalir"
        Me.cbSalir.Size = New System.Drawing.Size(64, 56)
        Me.cbSalir.TabIndex = 13
        Me.cbSalir.Text = "Salir"
        Me.cbSalir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbSalir, "Cierra el Gestor de Pagos de Asistencia")
        '
        'lbxCompania
        '
        Me.lbxCompania.Location = New System.Drawing.Point(128, 248)
        Me.lbxCompania.Name = "lbxCompania"
        Me.lbxCompania.Size = New System.Drawing.Size(80, 95)
        Me.lbxCompania.TabIndex = 14
        Me.lbxCompania.Visible = False
        '
        'lbxProducto
        '
        Me.lbxProducto.Location = New System.Drawing.Point(224, 248)
        Me.lbxProducto.Name = "lbxProducto"
        Me.lbxProducto.Size = New System.Drawing.Size(80, 95)
        Me.lbxProducto.TabIndex = 15
        Me.lbxProducto.Visible = False
        '
        'picTest
        '
        Me.picTest.Cursor = System.Windows.Forms.Cursors.Help
        Me.picTest.Image = CType(resources.GetObject("picTest.Image"), System.Drawing.Image)
        Me.picTest.Location = New System.Drawing.Point(8, 8)
        Me.picTest.Name = "picTest"
        Me.picTest.Size = New System.Drawing.Size(48, 48)
        Me.picTest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picTest.TabIndex = 37
        Me.picTest.TabStop = False
        Me.ttipAyuda.SetToolTip(Me.picTest, "La aplicación se esta ejecutando en modo pruebas")
        Me.picTest.Visible = False
        '
        'prbProgreso
        '
        Me.prbProgreso.Location = New System.Drawing.Point(40, 568)
        Me.prbProgreso.Name = "prbProgreso"
        Me.prbProgreso.Size = New System.Drawing.Size(402, 8)
        Me.prbProgreso.TabIndex = 16
        Me.prbProgreso.Visible = False
        '
        'CR2
        '
        Me.CR2.Enabled = True
        Me.CR2.Location = New System.Drawing.Point(304, 504)
        Me.CR2.Name = "CR2"
        Me.CR2.OcxState = CType(resources.GetObject("CR2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CR2.Size = New System.Drawing.Size(28, 28)
        Me.CR2.TabIndex = 65
        '
        'frmPrincipalAperturas
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(786, 580)
        Me.Controls.Add(Me.CR2)
        Me.Controls.Add(Me.prbProgreso)
        Me.Controls.Add(Me.lbxProducto)
        Me.Controls.Add(Me.lbxCompania)
        Me.Controls.Add(Me.cbSalir)
        Me.Controls.Add(Me.cbImprimir)
        Me.Controls.Add(Me.cbAvisos)
        Me.Controls.Add(Me.cbAperturar)
        Me.Controls.Add(Me.cbBorrar)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.stbEstado)
        Me.Controls.Add(Me.lvwAperturas)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.cbxCompania)
        Me.Controls.Add(Me.lbCompaniaAsistencia)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.picTest)
        Me.Controls.Add(Me.gbxFiltro)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrincipalAperturas"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Siniestros: Area de Asistencia  -   Aperturas Automáticas"
        Me.gbxFiltro.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.CR2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub frmPrincipalSuplidos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InicioApp()
    End Sub

    Private Sub InicioApp()
        On Error GoTo InicioApp_Err

        'Dim parametroApp As String
        'parametroApp = Microsoft.VisualBasic.Command

        'strIdProceso = "A"
        frmInstAperturas = Me

        'parametroApp = Microsoft.VisualBasic.Command
        If parametroApp = "PM" Then
            claseBDAperturas.ConnexionPruebas()
            claseBDAperturas.BDComand.CommandTimeout = 300
            picTest.Show()
            picTest.BringToFront()
        Else
            claseBDAperturas.BDWorkConnect.CommandTimeout = 0
            claseBDAperturas.BDComand.CommandTimeout = 0
            picTest.Hide()
            picTest.SendToBack()
        End If

        'MUL ini no se recogia el parametro del usuario
        'strCodUserAplicacion = parametroApp
        'MUL fin no se recogia el parametro del usuario

        ' Creación de objetos  
        strPathIconos = "K:\Graficos\"

        ' Identificador de componente
        strIDComp$ = "APE"

        ' Establecer ToolTipText de los Botones
        '
        'Selector.ToolTipText = " Ejecuta la busqueda de Aperturas según los criterios asignados "
        'dtpDesde.ToolTipText = " Abre el calendario de selección "
        'dtpHasta.ToolTipText = " Abre el calendario de selección "
        'Salir.ToolTipText = " Cierra el Gestor de Aperturas de Asistencia "
        'Imprimir.ToolTipText = " Imprime la selección de pantalla "
        'Proceso.ToolTipText = " Ejecución Proceso de Aperturas "
        'ErrAvi.ToolTipText = " Visualizar Histórico de Errores "
        'Image1.ToolTipText = " Visualizar Histórico de Errores "
        'SelTodos.ToolTipText = " Selecciona todos los expedientes de la consulta "
        'SelNinguno.ToolTipText = " Cancela la selección "
        'Borrar.ToolTipText = " Elimina la selección actual de datos "
        'BorrarBis.ToolTipText = " Elimina la selección actual de datos "

        ' Establecer ToolTipText para etiquetas
        '
        'Label1.ToolTipText = " Producto seleccionado "
        'Label2.ToolTipText = " Compañia de Asistencia seleccionada para el proceso "
        'Label3.ToolTipText = " Ejecuta la busqueda de Aperturas según los criterios asignados "
        'Frame1.ToolTipText = " Filtros sobre el resultado de la consulta "
        'Label4.ToolTipText = " Ejecución Proceso de Aperturas "
        'Label5.ToolTipText = " Visualizar Histórico de Errores "
        'Label6.ToolTipText = " Imprime la selección de pantalla "
        'Label7.ToolTipText = " Cierra el Gestor de Aperturas de Asistencia "
        'Label8.ToolTipText = " Selecciona todos los expedientes de la consulta "
        'Label9.ToolTipText = " Cancela la Selección "
        'txtDesde.ToolTipText = " Fecha de inicio seleccionada "
        'txtHasta.ToolTipText = " Fecha final seleccionada "
        'Label15.ToolTipText = " Elimina la selección actual de datos "

        ' Inicialización ComboBox de Productos
        If Not LlenarComboProducto(cbxProducto, lbxProducto, claseBDAperturas, "TODOS") Then
            Err.Raise(Val(strGlobalNumErr))
        Else
            cbxProducto.SelectedIndex = cbxProducto.Items.Count - 1
        End If

        ' Inicialización ComboBox de Compañías
        If LlenarComboCias(cbxCompania, lbxCompania) Then
            cbxCompania.SelectedIndex = 0
        Else
            MsgBox("Ha ocurrido un error iniciando la aplicación", MsgBoxStyle.Critical)
        End If

        ' Asignación de valores iniciales
        '
        dtpDesde.Value = Today
        dtpHasta.Value = Today

        bwflag = False

        FiltroTodos.PerformClick()
        strFiltro = "T"
        Exit Sub

InicioApp_Err:
        'JCLopez_i
        MsgBox("Ha ocurrido un error iniciando la aplicación", MsgBoxStyle.Exclamation)
        'objError.Tipo = Pantalla
        'objError.Ver(IdProceso, strGlobalNumErr, , Codcia)
        'JCLopez_f
    End Sub

    Private Sub cbBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBuscar.Click

        On Error GoTo Buscar_Error

        If cbxBusquedaAvanzada.Checked = False Then
            If cbxCompania.Text <> "" Then
                strSigCompa = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))
                strCodProducto = cbxProducto.Text

                If strCodProducto <> "Todos los Productos" Then
                    strCodProducto = lbxProducto.Items.Item(cbxProducto.SelectedIndex)
                End If

                'FiltroTodos.Select()
                Busqueda()
            Else
                MsgBox("Debe escoger Producto y/o Compañía de Asistencia", MsgBoxStyle.Information)
            End If
            'sigcompa = VB6.GetItemString(lstCompania, VB6.GetItemData(cboCompania, cboCompania.SelectedIndex))
            '''''''
        Else
            busquedaAvanzada()
        End If

        Exit Sub

Buscar_Error:
        MsgBox("Ha ocurrido un error al buscar", MsgBoxStyle.Exclamation)
    End Sub

    Private Sub Busqueda()
        On Error GoTo Busqueda_Error

        Dim dtFechaDesde, dtFechaHasta As Date

        dtFechaDesde = dtpDesde.Value
        dtFechaHasta = dtpHasta.Value

        RefrescarGrid(dtFechaDesde, dtFechaHasta, strCodProducto, strSigCompa)
        Exit Sub
Busqueda_Error:
        MsgBox("Ha ocurrido un error en la búsqueda", MsgBoxStyle.Critical)
    End Sub

    Private Sub BusquedaAvanzada()
        On Error GoTo BusquedaAvanzada_Error

        If tbSiniestro.Text = "" And tbReferencia.Text = "" Then
            MsgBox("No se ha indicado ningún siniestro / referencia", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If tbSiniestro.Text <> "" Then
            strCampoBuscaAvanzada = "T1_Codsin"
            strValorBuscaAvanzada = Trim(tbSiniestro.Text)
        Else
            strCampoBuscaAvanzada = "T1_Refer"
            strValorBuscaAvanzada = Trim(tbReferencia.Text)
        End If

        RefrescarGrid(dtpDesde.Value, dtpHasta.Value, cbxProducto.Text, strSigCompa)
        Exit Sub

BusquedaAvanzada_Error:
        MsgBox("Ha ocurrido un error realizando la búsqueda avanzada.", MsgBoxStyle.Critical)
    End Sub


    Private Sub cbxCompania_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxCompania.SelectedIndexChanged

        strCodCia = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))

        If Not DatosCiaAsistencia(strCodCia) Then
            MsgBox("No hay datos", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub cbSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSalir.Click
        Application.Exit()
    End Sub

    Private Sub cbBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBorrar.Click
        On Error GoTo Borra_Err

        ' Declaraciones
        '
        Dim objListItem As ListViewItem
        Dim strSQL As String
        Dim objAperturas As clsAperturar_NET
        Dim msgRetorno As MsgBoxResult

        ' Creación de objetos
        '
        objAperturas = New clsAperturar_NET

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        ' Confirmación de la orden de eliminación de registros
        '
        msgRetorno = MsgBox("¿Está seguro de querer eliminar los datos seleccionados?", MsgBoxStyle.YesNo)

        'If CBool(gstrError) Then
        If msgRetorno = MsgBoxResult.Yes Then

            For Each objListItem In lvwAperturas.Items
                Call ActualizarPorcentaje(lvwAperturas.Items.Count)
                If objListItem.Checked Then
                    objAperturas.Referencias.Add(objListItem.Text)
                End If
            Next objListItem
        End If
        If objAperturas.Referencias.Count() > 0 Then
            If objAperturas.DeleteAperturasAsistencia((objAperturas.Referencias)) Then
                MsgBox("La eliminación de los registros seleccionados se ha procesado correctamente.", MsgBoxStyle.Information)
            Else
                Err.Raise(1)
            End If
        End If
        'stbEstado.Panels(2).Text = ""
        prbProgreso.Visible = False
        Call RefrescarGrid((dtpDesde.Value), (dtpHasta.Value), (cbxProducto.Text), strSigCompa)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

Borra_Err:
        Me.Cursor = System.Windows.Forms.Cursors.Default
        MsgBox("Se ha producido un error en la Base de Datos. Los registros seleccionados no han sido borrados", MsgBoxStyle.Critical)
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cbAperturar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAperturar.Click
        ProcesoAperturar()
    End Sub

    Private Sub ProcesoAperturar()
        If strCodCia = "A" Then
            MsgBox("El contrato de la Compañía Angel ha sido resuelto. No es posible aperturar mas expedientes de esta compañía", MsgBoxStyle.Information)
            Exit Sub
        End If

        ' Declaraciones
        Dim objListItem As ListViewItem
        Dim objAperturar As clsAperturar_NET
        Dim vntReferencia As Object

        ' Inicialización Barra de Progreso
        prbProgreso.Minimum = 0
        prbProgreso.Value = 1
        prbProgreso.Maximum = lvwAperturas.Items.Count

        objAperturar = New clsAperturar_NET
        Call objAperturar.Inicializar(strCodCia)

        stbEstado.Panels(2).Text = "Comprobando Referencias (Filtros) ..."

        For Each objListItem In lvwAperturas.Items
            Call ActualizarPorcentaje(lvwAperturas.Items.Count)
            If objListItem.Checked And objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text <> "P" Then
                If chkFiltroAvisos.CheckState And objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text = "A" Then
                    objAperturar.Referencias.Add(objListItem.Text) ' Warnings que se pueden procesar
                ElseIf chkFiltroAvisos.CheckState = 0 And objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text <> "A" Then
                    If objAperturar.Filtros(objListItem) Then
                        objAperturar.Referencias.Add(objListItem.Text) ' Añadir referencias que se pueden procesar a colección.
                    End If
                End If
            End If
        Next objListItem

        If objAperturar.Referencias.Count > 0 Then
            stbEstado.Panels(2).Text = "Aperturando Siniestros ..."

            prbProgreso.Minimum = 0
            prbProgreso.Value = 0
            prbProgreso.Maximum = objAperturar.Referencias.Count()

            For Each vntReferencia In objAperturar.Referencias
                Call ActualizarPorcentaje(lvwAperturas.Items.Count)
                'objListItem = lvwAperturas.FindItem(vntReferencia, MSComctlLib.ListFindItemWhereConstants.lvwText)
                objListItem = BuscarItemPorNombreListView(vntReferencia)
                If Not objListItem Is Nothing Then Call objAperturar.Aperturar(objListItem)
            Next vntReferencia
            stbEstado.Panels(2).Text = "Proceso Finalizado"
        Else
            stbEstado.Panels(2).Text = "No se ha realizado la Apertura"
        End If
        prbProgreso.Visible = False
        'strSigCompa = VB6.GetItemString(lstCompania, VB6.GetItemData(cboCompania, cboCompania.SelectedIndex))
        strSigCompa = Trim(lbxCompania.GetItemText(cbxCompania.SelectedIndex))
        strFiltro = "T"
        FiltroTodos.PerformClick()
        Call RefrescarGrid((dtpDesde.Value), (dtpHasta.Value), (cbxProducto.Text), strSigCompa)

    End Sub

    Private Function BuscarItemPorNombreListView(ByVal strTexto As String) As ListViewItem

        For Each lvi As ListViewItem In lvwAperturas.Items

            If lvi.Text.Equals(strTexto) Then Return lvi

            For Each si As ListViewItem.ListViewSubItem In lvi.SubItems

                If si.Text.Equals(strTexto) Then Return lvi

            Next

        Next

        Return Nothing

    End Function

    Public Sub ActualizarPorcentaje(ByRef Total As Integer)

        Dim intPorcentaje As Short

        On Error Resume Next

        If Not Total = -1 Then ' Actualizar barra de estado, de progreso y porcentaje
            prbProgreso.Visible = True
            prbProgreso.Value = prbProgreso.Value + 1
            intPorcentaje = System.Math.Round((prbProgreso.Value * 100) / Total, 0)
            If CStr(intPorcentaje) & " %" <> stbEstado.Panels(2).Text Then
                stbEstado.Panels(2).Text = CStr(intPorcentaje) & " %"
            End If
        Else
            prbProgreso.Visible = False
        End If

    End Sub

    Private Sub FiltroTodos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FiltroTodos.Click, FiltroAviso.Click, FiltroErrores.Click, FiltroNoPagados.Click, FiltroPagados.Click
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
        If lvwAperturas.Items.Count > 0 Then
            For Each objlistitem In lvwAperturas.Items
                If objlistitem.Tag <> "1" Then
                    If objlistitem.Text <> "" And objlistitem.Text <> "No Existe" Then
                        Call ColorListItem(objlistitem, Color.Green)
                        objlistitem.Checked = blnCheck
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub RefrescarGrid(ByRef FecDes As Date, ByRef FecHas As Date, ByRef produc As String, ByRef compa As String)

        ' Declaraciones
        '
        Dim strSQL As String
        Dim objListItem As ListViewItem
        Dim objListSubItem As ListViewItem.ListViewSubItem
        Dim Index As Short
        Dim strOrderBy As String

        ' 1/12/2004 JLL
        '
        ' Aprovechando el cambio de compañia de aistencia ( de Angel a Reparalia ) se realizan
        ' cambios estructurales en las tablas de asistencia para mejorar y simplificar el acceso
        ' a los datos, esto provoca que a partir de esta fecha se tenga que diferenciar en la
        ' consulta de aperturas de siniestros cuando ésta es de aperturas de Angel y cuando no.
        '
        If cbxBusquedaAvanzada.CheckState Then
            If strCodCia = "A" Then
                strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T3_CCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1, Angel_T3 " & "WHERE  Angel_T1." & strCampoBuscaAvanzada & " = '" & Trim(strValorBuscaAvanzada) & "' AND " & "       Angel_T1.T1_Refer = Angel_T3.T3_Refer and Angel_T1.T1_Codcia = '" & strCodCia & "'"
            Else
                strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T1_CODCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1 " & "WHERE  Angel_T1." & strCampoBuscaAvanzada & " = '" & Trim(strValorBuscaAvanzada) & "' AND " & "       Angel_T1.T1_Codcia = '" & strCodCia & "'"
            End If
        Else
            ' Falta cargar dato ASEGURADO
            ' e investigar propiedad ghosted del Listitem
            '
            If strCodCia = "A" Then
                If produc = "Todos los Productos" Or produc = "" Then
                    If strFiltro = "T" Then
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T3_CCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1, Angel_T3 " & "WHERE  Angel_T1.T1_REFER = Angel_T3.T3_REFER AND (Angel_T1.T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND " & "       Angel_T1.T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    Else
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T3_CCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1, Angel_T3 " & "WHERE  Angel_T1.T1_REFER = Angel_T3.T3_REFER AND (Angel_T1.T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND Angel_T1.T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') AND " & "       T1_ESTADO = '" & strFiltro & "' and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    End If
                Else
                    If strFiltro = "T" Then
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T3_CCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1, Angel_T3 " & "WHERE  Angel_T1.T1_REFER = Angel_T3.T3_REFER AND (T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') AND " & "       Angel_T1.T1_CIA = '" & lbxProducto.GetItemText(cbxProducto.SelectedIndex - 1) & "' and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    Else
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T3_CCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1, Angel_T3 " & "WHERE  Angel_T1.T1_REFER = Angel_T3.T3_REFER AND (T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') AND " & "       T1_CIA = '" & lbxProducto.GetItemText(cbxProducto.SelectedIndex - 1) & "' AND T1_ESTADO = '" & strFiltro & "' and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    End If
                End If
            Else
                If produc = "Todos los Productos" Or produc = "" Then
                    If strFiltro = "T" Then
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T1_CODCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1 " & "WHERE  (Angel_T1.T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND " & "       Angel_T1.T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    Else
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T1_CODCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1 " & "WHERE  (Angel_T1.T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND Angel_T1.T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') AND " & "       T1_ESTADO = '" & strFiltro & "' and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    End If
                Else
                    If strFiltro = "T" Then
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T1_CODCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1 " & "WHERE  (T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') AND " & "       Angel_T1.T1_CIA = '" & produc & "' and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    Else
                        strSQL = "SELECT T1_REFER, T1_CIA, T1_CODSIN, T1_POLIZA, ' ' AS T1_AVISO, T1_FAPER, T1_FSINI, T1_DESCR, " & "       T1_CODCAUSA, T1_FGRABA AS FECGRA, T1_ESTADO, FECHAPROCESO, ISNULL(T1_Aperturapor,'') AS T1_APERTURAPOR " & "FROM   Angel_T1 " & "WHERE  (T1_FGRABA >= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecDes, False, False) & "' AND T1_FGRABA <= '" & claseUtilidadesAperturas.FormatoFechaSQL(FecHas, False, False) & "') AND " & "       T1_CIA = '" & produc & "' AND T1_ESTADO = '" & strFiltro & "' and Angel_T1.T1_Codcia = '" & strCodCia & "'"
                    End If

                End If
            End If
        End If

        strOrderBy = " Order By Angel_T1.T1_Codsin, Angel_T1.T1_Refer"
        strSQL = strSQL & strOrderBy

        'Establece origen de datos para Crystal Reports
        '
        strSQLCR = strSQL
        strSQLCR = Mid(strSQLCR, 8, Len(strSQLCR) - 7)
        strSQLCR = "SELECT '" & strIDComp & "', " & strSQLCR

        ' Carga de datos en el objeto ListView
        '
        If strCodCia = "A" Then
            Call CargarListView_aperturas(lvwAperturas, strSQL, "", _
                                          "T1_REFER", "T1_CIA", "T1_CODSIN", "T1_POLIZA", "T1_AVISO", "T1_FAPER", "T1_FSINI", "T1_DESCR", "T3_CCAUSA", "FECGRA", "FECHAPROCESO", "T1_APERTURAPOR", "T1_ESTADO")
        Else
            Call CargarListView_aperturas(lvwAperturas, strSQL, "", _
                                          "T1_REFER", "T1_CIA", "T1_CODSIN", "T1_POLIZA", "T1_AVISO", "T1_FAPER", "T1_FSINI", "T1_DESCR", "T1_CODCAUSA", "FECGRA", "FECHAPROCESO", "T1_APERTURAPOR", "T1_ESTADO")
        End If

        stbEstado.Panels(2).Text = Format(lvwAperturas.Items.Count, "#,##0") & " Referencias"

        ' Poner atenuados aquellos siniestros que ya esten aperturados, o con el
        ' color identificativo para aquellos que tengan errores o avisos
        '
        'If lvwAperturas.ListItems.Count > 0 Then
        Index = 0
        For Each objListItem In lvwAperturas.Items
            'If Not (objListItem.ListSubItems.Item(2) = "" Or IsNull(objListItem.ListSubItems.Item(2)) Or objListItem.ListSubItems(2) = "No Existe") Then
            Select Case objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text
                Case "P"
                    'objListItem.ListView.Enabled = False
                    objListItem.Tag = "1"
                    ttipAyuda.SetToolTip(objListItem.ListView, "Este Siniestro ya está Procesado")
                    Call ColorListItem(objListItem, Color.Gray)
                Case "X"
                    objListItem.Tag = "0"
                    ttipAyuda.SetToolTip(objListItem.ListView, "Este Siniestro está pendiente de procesar")
                    Call ColorListItem(objListItem, Color.Green)
                Case "W", "A"
                    objListItem.Tag = "0"
                    ttipAyuda.SetToolTip(objListItem.ListView, "Este Siniestro tiene avisos pendientes de resolución")
                    Call ColorListItem(objListItem, Color.DarkGoldenrod)
                    objListItem.SubItems.Item(frmInstAperturas.AVISO.Index).Text = CodigoAviso(lvwAperturas.Items(Index).Text)
                Case "E"
                    objListItem.Tag = "0"
                    ttipAyuda.SetToolTip(objListItem.ListView, "Este Siniestro tiene mensajes de error del proceso")
                    Call ColorListItem(objListItem, Color.Red)
                    objListItem.SubItems.Item(frmInstAperturas.AVISO.Index).Text = CodigoAviso(lvwAperturas.Items(Index).Text)
                Case Else
                    objListItem.Tag = "0"
                    Call ColorListItem(objListItem, Color.Gray)
            End Select
            'End If

            ' En la columna de Perito substituimos el 0 o el >=1
            ' por la S o la N.
            '
            If objListItem.SubItems.Item(frmInstAperturas.T1_PERTURAPOR.Index).Text = "1" Then
                objListItem.SubItems.Item(frmInstAperturas.T1_PERTURAPOR.Index).Text = "S"
            End If
            Index = Index + 1
        Next objListItem
        'Else

        'End If
        bwflag = True
        Exit Sub
RefrescarGrid_Error:
        MsgBox("Error refrescando datos", MsgBoxStyle.Critical)
    End Sub


    Private Sub lvwAperturas_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lvwAperturas.ItemCheck
        Dim lvSuplidosAux As ListView

        lvSuplidosAux = sender
        If lvSuplidosAux.Items(e.Index).Tag = "1" Then
            e.NewValue = e.CurrentValue
        End If
    End Sub

    Private Sub cbAvisos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAvisos.Click

        'Private Sub ErrAvi_Click()

        '        On Error GoTo ErrAvi_Click_Err

        '        ' Declaraciones
        '        '
        '        Dim objListItem As ListItem
        '            ' Comprobamos que se haya seleccionado
        '            '
        '            If lvwAperturas.SelectedItem Is Nothing Then
        '                gstrError = "4007"
        '                Err.Raise(1)
        '            End If

        '            ' Comprobamos el estado de la referencia seleccionada
        '            '
        '            If lvwAperturas.SelectedItem.ListSubItems.Item("T1_ESTADO").Text = "W" Or lvwAperturas.SelectedItem.ListSubItems.Item("T1_ESTADO") = "A" Or lvwAperturas.SelectedItem.ListSubItems.Item("T1_ESTADO") = "E" Then
        '                If lvwAperturas.SelectedItem.Ghosted Then Exit Sub
        '                objListItem = lvwAperturas.SelectedItem
        '                objError.Tipo = BD
        '                objError.Ver(IdProceso, lvwAperturas.SelectedItem, , Codcia)
        '            ElseIf lvwAperturas.SelectedItem.ListSubItems.Item("T1_ESTADO").Text = "P" Then
        '                mdpbd.BDAuxRecord = objSiniestro.Siniestro(lvwAperturas.SelectedItem.ListSubItems.Item("T1_CODSIN").Text, True, UsuaApli)
        '            Else
        '                gstrError = "4008"
        '                Err.Raise(1)
        '            End If


        '        objListItem = Nothing
        '        Exit Sub

        'ErrAvi_Click_Err:
        '        objError.Tipo = Pantalla
        '        objError.Ver(IdProceso, gstrError, , Codcia)
        '    End Sub




        On Error GoTo cbAvisos_Error

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim strReferencia As String
        Dim claseSiniestros As clsSiniestro_NET
        Dim frmInstanciaErrores As New frmVisorErrores


        ' Comprobamos que se haya seleccionado
        If lvwAperturas.CheckedItems.Count > 0 Then
            ' Comprobamos el estado de la referencia seleccionada
            If lvwAperturas.CheckedItems(0).SubItems(frmInstAperturas.T1_ESTADO.Index).Text = "W" Or _
               lvwAperturas.CheckedItems(0).SubItems(frmInstAperturas.T1_ESTADO.Index).Text = "A" Or _
               lvwAperturas.CheckedItems(0).SubItems(frmInstAperturas.T1_ESTADO.Index).Text = "E" Then

                If lvwAperturas.CheckedItems(0).Tag = "1" Then
                    Exit Sub
                End If
                objlistitem = lvwAperturas.CheckedItems(0)
                strReferencia = lvwAperturas.CheckedItems(0).SubItems(frmInstAperturas.T1_REFER.Index).Text
                frmInstanciaErrores.MostrarErrores(strReferencia)
                frmInstanciaErrores.Show()
            ElseIf lvwAperturas.CheckedItems(0).SubItems(frmInstAperturas.T1_ESTADO.Index).Text = "P" Then
                claseBDAperturas.BDAuxRecord = claseSiniestros.Siniestro(lvwAperturas.CheckedItems(0).SubItems(frmInstAperturas.T1_CODSIN.Index).Text, True, strUsuarioAplicacion)
            Else
                strError = "4008"
                Err.Raise(1)
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
        Imprimir_aperturas()
    End Sub

    Private Sub cbxBusquedaAvanzada_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxBusquedaAvanzada.CheckedChanged
        Dim boolActivo As Boolean

        boolActivo = cbxBusquedaAvanzada.Checked

        'Habilitados
        lbSiniestro.Enabled = boolActivo
        lbReferencia.Enabled = boolActivo
        tbReferencia.Enabled = boolActivo
        tbSiniestro.Enabled = boolActivo
        If Not boolActivo Then
            tbReferencia.Text = ""
            tbSiniestro.Text = ""
        End If

        'Deshabilitados
        FiltroAviso.Enabled = Not boolActivo
        FiltroErrores.Enabled = Not boolActivo
        FiltroNoPagados.Enabled = Not boolActivo
        FiltroPagados.Enabled = Not boolActivo
        FiltroTodos.Enabled = Not boolActivo

        lbFechaDesde.Enabled = Not boolActivo
        dtpDesde.Enabled = Not boolActivo
        lbFechaHasta.Enabled = Not boolActivo
        dtpHasta.Enabled = Not boolActivo

        cbxProducto.Enabled = Not boolActivo
        chkFiltroAvisos.Enabled = Not boolActivo
    End Sub


End Class
