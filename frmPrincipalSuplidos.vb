Public Class frmPrincipalSuplidos
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
    Friend WithEvents FiltroPagados As System.Windows.Forms.RadioButton
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
    Friend WithEvents sbPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents cbxCompania As System.Windows.Forms.ComboBox
    Friend WithEvents lbFechaHasta As System.Windows.Forms.Label
    Friend WithEvents lbFechaDesde As System.Windows.Forms.Label
    Friend WithEvents lbCompaniaAsistencia As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSuplido As System.Windows.Forms.TextBox
    Friend WithEvents txtSuplidoAdicional As System.Windows.Forms.TextBox
    Friend WithEvents T7_CODSIN As System.Windows.Forms.ColumnHeader
    Friend WithEvents T7_REFER As System.Windows.Forms.ColumnHeader
    Friend WithEvents T7_CODRAM As System.Windows.Forms.ColumnHeader
    Friend WithEvents T7_NUMPOL As System.Windows.Forms.ColumnHeader
    Friend WithEvents T7_ESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents T7_FESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents Fichero As System.Windows.Forms.ColumnHeader
    Friend WithEvents FechaProceso As System.Windows.Forms.ColumnHeader
    Friend WithEvents NUMFACTURA As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECHA_FACT As System.Windows.Forms.ColumnHeader
    Friend WithEvents lvwSuplidos As System.Windows.Forms.ListView
    Friend WithEvents cbSuplidos As System.Windows.Forms.Button
    Public WithEvents CR2 As AxCrystal.AxCrystalReport
    Friend WithEvents picTest As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrincipalSuplidos))
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
        Me.txtSuplidoAdicional = New System.Windows.Forms.TextBox
        Me.txtSuplido = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbBuscar = New System.Windows.Forms.Button
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker
        Me.lbFechaHasta = New System.Windows.Forms.Label
        Me.lbFechaDesde = New System.Windows.Forms.Label
        Me.lvwSuplidos = New System.Windows.Forms.ListView
        Me.T7_CODSIN = New System.Windows.Forms.ColumnHeader
        Me.T7_REFER = New System.Windows.Forms.ColumnHeader
        Me.T7_CODRAM = New System.Windows.Forms.ColumnHeader
        Me.T7_NUMPOL = New System.Windows.Forms.ColumnHeader
        Me.T7_ESTADO = New System.Windows.Forms.ColumnHeader
        Me.T7_FESTADO = New System.Windows.Forms.ColumnHeader
        Me.Fichero = New System.Windows.Forms.ColumnHeader
        Me.FechaProceso = New System.Windows.Forms.ColumnHeader
        Me.NUMFACTURA = New System.Windows.Forms.ColumnHeader
        Me.FECHA_FACT = New System.Windows.Forms.ColumnHeader
        Me.stbEstado = New System.Windows.Forms.StatusBar
        Me.sbPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cbTodos = New System.Windows.Forms.Button
        Me.cbNinguno = New System.Windows.Forms.Button
        Me.cbBorrar = New System.Windows.Forms.Button
        Me.cbSuplidos = New System.Windows.Forms.Button
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
        Me.gbxFiltro.Location = New System.Drawing.Point(568, 8)
        Me.gbxFiltro.Name = "gbxFiltro"
        Me.gbxFiltro.Size = New System.Drawing.Size(208, 160)
        Me.gbxFiltro.TabIndex = 3
        Me.gbxFiltro.TabStop = False
        Me.gbxFiltro.Text = "Filtro"
        Me.ttipAyuda.SetToolTip(Me.gbxFiltro, "Filtros sobre el resultado de la consulta")
        '
        'chkFiltroAvisos
        '
        Me.chkFiltroAvisos.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.chkFiltroAvisos.Location = New System.Drawing.Point(16, 128)
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
        Me.FiltroErrores.Location = New System.Drawing.Point(16, 104)
        Me.FiltroErrores.Name = "FiltroErrores"
        Me.FiltroErrores.Size = New System.Drawing.Size(104, 16)
        Me.FiltroErrores.TabIndex = 4
        Me.FiltroErrores.Text = "Errores ( E )"
        Me.ttipAyuda.SetToolTip(Me.FiltroErrores, "Muestra sólo los registros en los que se ha producido error")
        '
        'FiltroAviso
        '
        Me.FiltroAviso.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroAviso.ForeColor = System.Drawing.Color.DarkGoldenrod
        Me.FiltroAviso.Location = New System.Drawing.Point(16, 84)
        Me.FiltroAviso.Name = "FiltroAviso"
        Me.FiltroAviso.Size = New System.Drawing.Size(104, 16)
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
        Me.FiltroNoPagados.Size = New System.Drawing.Size(112, 16)
        Me.FiltroNoPagados.TabIndex = 2
        Me.FiltroNoPagados.Text = "No Pagados ( X )"
        Me.ttipAyuda.SetToolTip(Me.FiltroNoPagados, "Muestra sólo los pendientes de aperturar")
        '
        'FiltroPagados
        '
        Me.FiltroPagados.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroPagados.ForeColor = System.Drawing.Color.Gray
        Me.FiltroPagados.Location = New System.Drawing.Point(16, 44)
        Me.FiltroPagados.Name = "FiltroPagados"
        Me.FiltroPagados.Size = New System.Drawing.Size(104, 16)
        Me.FiltroPagados.TabIndex = 1
        Me.FiltroPagados.Text = "Pagados ( P )"
        Me.ttipAyuda.SetToolTip(Me.FiltroPagados, "Muestra sólo los aperturados")
        '
        'FiltroTodos
        '
        Me.FiltroTodos.Checked = True
        Me.FiltroTodos.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FiltroTodos.ForeColor = System.Drawing.Color.Black
        Me.FiltroTodos.Location = New System.Drawing.Point(16, 24)
        Me.FiltroTodos.Name = "FiltroTodos"
        Me.FiltroTodos.Size = New System.Drawing.Size(104, 16)
        Me.FiltroTodos.TabIndex = 0
        Me.FiltroTodos.TabStop = True
        Me.FiltroTodos.Text = "Todos"
        Me.ttipAyuda.SetToolTip(Me.FiltroTodos, "Muestra todos los registros")
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtSuplidoAdicional)
        Me.GroupBox2.Controls.Add(Me.txtSuplido)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
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
        Me.GroupBox2.Size = New System.Drawing.Size(552, 104)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Criterios de Selección"
        '
        'txtSuplidoAdicional
        '
        Me.txtSuplidoAdicional.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.txtSuplidoAdicional.Location = New System.Drawing.Point(352, 62)
        Me.txtSuplidoAdicional.Name = "txtSuplidoAdicional"
        Me.txtSuplidoAdicional.Size = New System.Drawing.Size(112, 21)
        Me.txtSuplidoAdicional.TabIndex = 15
        Me.txtSuplidoAdicional.Text = ""
        Me.txtSuplidoAdicional.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSuplido
        '
        Me.txtSuplido.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.txtSuplido.Location = New System.Drawing.Point(352, 30)
        Me.txtSuplido.Name = "txtSuplido"
        Me.txtSuplido.Size = New System.Drawing.Size(112, 21)
        Me.txtSuplido.TabIndex = 14
        Me.txtSuplido.Text = ""
        Me.txtSuplido.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label2.Location = New System.Drawing.Point(208, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 19)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Importe Suplido Adicional:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label1.Location = New System.Drawing.Point(216, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 16)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Importe Suplido:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
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
        'lvwSuplidos
        '
        Me.lvwSuplidos.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvwSuplidos.CheckBoxes = True
        Me.lvwSuplidos.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.T7_CODSIN, Me.T7_REFER, Me.T7_CODRAM, Me.T7_NUMPOL, Me.T7_ESTADO, Me.T7_FESTADO, Me.Fichero, Me.FechaProceso, Me.NUMFACTURA, Me.FECHA_FACT})
        Me.lvwSuplidos.FullRowSelect = True
        Me.lvwSuplidos.GridLines = True
        Me.lvwSuplidos.Location = New System.Drawing.Point(8, 176)
        Me.lvwSuplidos.Name = "lvwSuplidos"
        Me.lvwSuplidos.Size = New System.Drawing.Size(768, 288)
        Me.lvwSuplidos.TabIndex = 5
        Me.lvwSuplidos.View = System.Windows.Forms.View.Details
        '
        'T7_CODSIN
        '
        Me.T7_CODSIN.Text = "Siniestro"
        Me.T7_CODSIN.Width = 83
        '
        'T7_REFER
        '
        Me.T7_REFER.Text = "Referencia"
        Me.T7_REFER.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T7_REFER.Width = 76
        '
        'T7_CODRAM
        '
        Me.T7_CODRAM.Text = "Producto"
        Me.T7_CODRAM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T7_CODRAM.Width = 74
        '
        'T7_NUMPOL
        '
        Me.T7_NUMPOL.Text = "Póliza"
        Me.T7_NUMPOL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T7_NUMPOL.Width = 89
        '
        'T7_ESTADO
        '
        Me.T7_ESTADO.Text = "Estado"
        Me.T7_ESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'T7_FESTADO
        '
        Me.T7_FESTADO.Text = "Fecha Grab."
        Me.T7_FESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.T7_FESTADO.Width = 81
        '
        'Fichero
        '
        Me.Fichero.Text = "Fichero"
        Me.Fichero.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FechaProceso
        '
        Me.FechaProceso.Text = "Fecha Proceso"
        Me.FechaProceso.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.FechaProceso.Width = 82
        '
        'NUMFACTURA
        '
        Me.NUMFACTURA.Text = "Nº Factura"
        Me.NUMFACTURA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NUMFACTURA.Width = 70
        '
        'FECHA_FACT
        '
        Me.FECHA_FACT.Text = "Fecha Fra."
        Me.FECHA_FACT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.FECHA_FACT.Width = 71
        '
        'stbEstado
        '
        Me.stbEstado.Location = New System.Drawing.Point(0, 556)
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
        'cbSuplidos
        '
        Me.cbSuplidos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbSuplidos.Image = CType(resources.GetObject("cbSuplidos.Image"), System.Drawing.Image)
        Me.cbSuplidos.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbSuplidos.Location = New System.Drawing.Point(520, 488)
        Me.cbSuplidos.Name = "cbSuplidos"
        Me.cbSuplidos.Size = New System.Drawing.Size(64, 56)
        Me.cbSuplidos.TabIndex = 10
        Me.cbSuplidos.Text = "Suplidos"
        Me.cbSuplidos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ttipAyuda.SetToolTip(Me.cbSuplidos, "Ejecución Proceso de Pagos")
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
        Me.prbProgreso.Location = New System.Drawing.Point(40, 564)
        Me.prbProgreso.Name = "prbProgreso"
        Me.prbProgreso.Size = New System.Drawing.Size(402, 8)
        Me.prbProgreso.TabIndex = 16
        Me.prbProgreso.Visible = False
        '
        'CR2
        '
        Me.CR2.Enabled = True
        Me.CR2.Location = New System.Drawing.Point(384, 488)
        Me.CR2.Name = "CR2"
        Me.CR2.OcxState = CType(resources.GetObject("CR2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CR2.Size = New System.Drawing.Size(28, 28)
        Me.CR2.TabIndex = 24
        '
        'picTest
        '
        Me.picTest.Cursor = System.Windows.Forms.Cursors.Help
        Me.picTest.Image = CType(resources.GetObject("picTest.Image"), System.Drawing.Image)
        Me.picTest.Location = New System.Drawing.Point(8, 8)
        Me.picTest.Name = "picTest"
        Me.picTest.Size = New System.Drawing.Size(48, 48)
        Me.picTest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picTest.TabIndex = 39
        Me.picTest.TabStop = False
        Me.picTest.Visible = False
        '
        'frmPrincipalSuplidos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(784, 578)
        Me.Controls.Add(Me.CR2)
        Me.Controls.Add(Me.prbProgreso)
        Me.Controls.Add(Me.lbxProducto)
        Me.Controls.Add(Me.lbxCompania)
        Me.Controls.Add(Me.cbSalir)
        Me.Controls.Add(Me.cbImprimir)
        Me.Controls.Add(Me.cbAvisos)
        Me.Controls.Add(Me.cbSuplidos)
        Me.Controls.Add(Me.cbBorrar)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.stbEstado)
        Me.Controls.Add(Me.lvwSuplidos)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.gbxFiltro)
        Me.Controls.Add(Me.cbxCompania)
        Me.Controls.Add(Me.lbCompaniaAsistencia)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.picTest)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrincipalSuplidos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Siniestros: Area de Asistencia  -  Pago Automáticos de Suplidos"
        Me.gbxFiltro.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
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

        ' parametroApp = Microsoft.VisualBasic.Command

        frmInstSuplidos = Me

        Arranque = True
        'IdProceso = "S"

        Select Case UCase(parametroApp)

            Case "P" '  Ejecución Programada
                TipoEjecucion = "P"

            Case "PP" ' Ejecución Programa en Pruebas
                claseBDSuplidos.ConnexionPruebas()
                TipoEjecucion = "P"
                picTest.Show()
                picTest.BringToFront()

            Case "PM" ' Ejecución Pruebas Manual
                claseBDSuplidos.ConnexionPruebas()
                TipoEjecucion = "M"
                picTest.Show()
                picTest.BringToFront()

            Case Else ' Ejecución Manual Real
                TipoEjecucion = "M"

        End Select

        claseBDSuplidos.BDWorkConnect.CommandTimeout = 0
        ' Valores globales
        '
        'CodUserApli = objUtiles.CodUser(UsuaApli)   ' Usuario de la aplicación
        If CodUserApli = "" Then CodUserApli = Microsoft.VisualBasic.Command
        'PathIconos = clses.GetParam("PathGraficos") ' Ubicación de objetos gráficos
        strPathIconos = "K:\Graficos\"

        'Identificador de componente
        '
        strIDComp = "PAG"

        ' Inicialización ComboBox de Compañías
        '
        If LlenarComboCias(cbxCompania, lbxCompania) Then
            cbxCompania.SelectedIndex = 0
        Else
            Err.Raise(Val(GlobalNumErr))
            'cbxCompania.Text = VB6.GetItemString(cbxCompania, CInt(strCiaDefault))
        End If


        ' Asignación de valores iniciales
        '
        dtpDesde.Value = Today
        dtpHasta.Value = Today

        bwflag = False

        CargarImportes()

        FiltroTodos.PerformClick()
        strFiltro = "T"

        ' Destrucción de objetos
        '
        'Set clses = Nothing

        ' En la carga del formulario efectuamos un cruce de referencias
        ' para actualizar el código de siniestro
        '
        CruceReferencias("AS" & strIdReferCompa, strCodcia)

        Arranque = False

        Exit Sub
        'JCLopez_i
        'objError.Pruebas = False
        'JCLopez_f

InicioApp_Err:
        'JCLopez_i
        MsgBox("Ha ocurrido un error iniciando la aplicación", MsgBoxStyle.Exclamation)
        'objError.Tipo = Pantalla
        'objError.Ver(IdProceso, strGlobalNumErr, , Codcia)
        'JCLopez_f
    End Sub

    Private Sub cbBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBuscar.Click

        On Error GoTo Buscar_Error

        If cbxCompania.Text <> "" Then

            Busqueda()
        Else
            MsgBox("Debe escoger Producto y/o Compañía de Asistencia", MsgBoxStyle.Information)
        End If
        'sigcompa = VB6.GetItemString(lstCompania, VB6.GetItemData(cboCompania, cboCompania.SelectedIndex))

        Exit Sub

Buscar_Error:
        MsgBox("Ha ocurrido un error al buscar", MsgBoxStyle.Exclamation)

    End Sub

    Private Sub CargarImportes()

        ' Tratamiento de errores
        '
        On Error GoTo CargarImportes_Err

        ' Declaraciones
        '
        Dim NumError As String
        Dim RsLocal As New ADODB.Recordset
        Dim strsql As String

        NumError = "4067"
        SuplidosExisten = False

        ' Establecemos la sql a ejecutar
        '
        strsql = "Select * From SuplidosAsistenciaImportes Where Codcia = '" & strCodcia & "'"

        ' Ejecutamos y cargamos consulta sql en objeto recordset
        '
        RsLocal.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        ' Leemos valores y mostramos en objetos de pantalla
        '
        If Not RsLocal.EOF Then
            txtSuplido.Text = Format(RsLocal.Fields("ImportePrincipal").Value, "##,##0.00")
            txtSuplidoAdicional.Text = Format(RsLocal.Fields("importeAdicional").Value, "##,##0.00")
            NumError = ""
            SuplidosExisten = True
        Else
            NumError = "4067"
            gstrError = "No se han podido cargar los valores para los importes de los suplidos de la compañia de asistencia seleccionada. El valor actual de liquidación será 0."

            Err.Raise(1)
        End If

        If Val(txtSuplido.Text) <= 0 Then
            SuplidosExisten = True
            gstrError = "El importe del suplido de la compañía de asistencia seleccionada es 0"
            'NumError = "4068"
            Err.Raise(1)
        ElseIf Val(txtSuplidoAdicional.Text) <= 0 Then
            SuplidosExisten = True
            'NumError = "4069"
            gstrError = "El importe del suplido adicional de la compañía de asistencia seleccionada es 0"
            Err.Raise(1)
        End If

        ' Cerramos y desinstanciamos objeto recordset local
        '
        RsLocal.Close()
        'UPGRADE_NOTE: El objeto RsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'RsLocal = Nothing
        Exit Sub

CargarImportes_Err:
        MsgBox("Ha ocurrido un error cargando los importes: " + gstrError)
        'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
        'objError.Ver(IdProceso, NumError, , Codcia)
    End Sub

    Private Sub Busqueda()
        On Error GoTo Busqueda_Error

        Dim dtFechaDesde, dtFechaHasta As Date

        dtFechaDesde = dtpDesde.Value
        dtFechaHasta = dtpHasta.Value

        strSigCompa = lbxCompania.Items.Item(cbxCompania.SelectedIndex)

        RefrescarGrid(dtFechaDesde, dtFechaHasta, strSigCompa)
        Exit Sub
Busqueda_Error:
        MsgBox("Ha ocurrido un error en la búsqueda", MsgBoxStyle.Critical)
    End Sub


    Private Sub cbxCompania_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxCompania.SelectedIndexChanged
        strCodcia = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))

        If Not DatosCiaAsistencia(strCodcia) Then
            MsgBox("No hay datos", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub cbSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSalir.Click
        Application.Exit()
    End Sub

    Private Sub cbBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBorrar.Click
        On Error GoTo Borrar_Error

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim strsql As String
        Dim objSuplidos As New clsSuplidos_NET
        Dim objLista As ListViewItem.ListViewSubItem

        Dim msgRetorno As MsgBoxResult

        ' Creación de objetos

        Me.Cursor = Cursors.WaitCursor

        ' Confirmación de la orden de eliminación de registros
        '
        gstrError = ""
        'objError.Tipo = mdpErroresMensajes.Tipo.Pregunta
        'objError.Ver(IdProceso, gstrError, " ¿ Esta seguro de querer eliminar los datos seleccionados ? ", strCodcia)

        msgRetorno = MsgBox("¿Está seguro de querer eliminar los datos seleccionados?", MsgBoxStyle.YesNo)

        'If CBool(gstrError) Then
        If msgRetorno = MsgBoxResult.Yes Then
            For Each objlistitem In lvwSuplidos.Items
                Call ActualizarPorcentaje(lvwSuplidos.Items.Count, prbProgreso, stbEstado)
                If objlistitem.Checked Then
                    objSuplidos.Referencias.Add(objlistitem.Text)
                    objSuplidos.Fichero.Add(objlistitem.SubItems(6))
                End If
            Next objlistitem
        End If
        If objSuplidos.Referencias.Count() > 0 Then
            If objSuplidos.DeleteSuplidos((objSuplidos.Referencias), (objSuplidos.Fichero)) Then
                MsgBox("La eliminación de los registros seleccionados se ha procesado correctamente.")
                'gstrError = "108"
                'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                'objError.Ver(IdProceso, gstrError, , strCodcia)
            Else
                Err.Raise(1)
            End If
        End If
        'stbEstado.Panels(4).Text = ""
        prbProgreso.Visible = False
        Call RefrescarGrid((dtpDesde.Value), (dtpHasta.Value), strSigCompa)
        Me.Cursor = Cursors.Default
        Exit Sub

Borrar_Error:
        MsgBox("Se ha producido un error en la Base de Datos. Los registros seleccionados no han sido borrados", MsgBoxStyle.Critical)
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cbSuplidos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSuplidos.Click
        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim objSuplidos As New clsSuplidos_NET
        Dim strRefer As String
        Dim vntReferencia As String
        Dim decSuplido, decSuplidoAdicional As Decimal

        If IsNumeric(txtSuplido.Text) Then
            decSuplido = Convert.ToDouble(txtSuplido.Text)
        End If

        If IsNumeric(txtSuplido.Text) Then
            decSuplidoAdicional = Convert.ToDouble(txtSuplido.Text)
        End If

        If (decSuplido <= 0 Or decSuplidoAdicional <= 0) And (txtSuplido.Text <> "" Or txtSuplidoAdicional.Text <> "") Then
            MsgBox("El importe del suplido y del suplido adicional han de ser mayores de 0 para ejecutar el proceso de pagos de suplidos", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        prbProgreso.Minimum = 0
        prbProgreso.Value = 1
        prbProgreso.Maximum = lvwSuplidos.Items.Count

        stbEstado.Panels(2).Text = "Comprobando Filtros ..."

        For Each objlistitem In lvwSuplidos.Items
            Call ActualizarPorcentaje(lvwSuplidos.Items.Count, prbProgreso, stbEstado)
            If objlistitem.Checked And objlistitem.SubItems.Item(frmInstSuplidos.T7_FESTADO.Index).Text <> "P" Then
                If chkFiltroAvisos.CheckState And objlistitem.SubItems.Item(frmInstSuplidos.T7_FESTADO.Index).Text = "A" Then
                    objSuplidos.Referencias.Add(objlistitem.Text) ' Warnings que se pueden procesar
                ElseIf objSuplidos.Filtros(objlistitem) Then
                    objSuplidos.Referencias.Add(objlistitem.Text) ' Añadir siniestros que se pueden procesar a colección.
                End If

                ' 27/10/2005 Mercedes
                stbEstado.Panels(2).Text = "Realizando Suplidos de Siniestros ..."
                prbProgreso.Minimum = 0
                prbProgreso.Value = 0
                If objSuplidos.Referencias.Count() > 0 Then
                    prbProgreso.Maximum = objSuplidos.Referencias.Count()
                    Call ActualizarPorcentaje(lvwSuplidos.Items.Count, prbProgreso, stbEstado)
                    If Not objlistitem Is Nothing Then Call objSuplidos.PagarSuplidos(objlistitem)
                    objSuplidos.Referencias.Remove(1)
                End If
                ' Fin 27/10/2005 Mercedes
            End If
        Next objlistitem

        'stbEstado.Panels(4).Text = ""
        prbProgreso.Visible = False
        strSigCompa = lbxCompania.Items.Item(cbxCompania.SelectedIndex)
        FiltroTodos.Checked = True
        strFiltro = "T"
        Call RefrescarGrid((dtpDesde.Value), (dtpHasta.Value), strSigCompa)

    End Sub


    Private Sub lvwSuplidos_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lvwSuplidos.ItemCheck
        Dim lvSuplidosAux As ListView

        lvSuplidosAux = sender
        If lvSuplidosAux.Items(e.Index).Tag = "1" Then
            e.NewValue = e.CurrentValue
        End If
    End Sub

    Private Sub FiltroTodos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FiltroTodos.Click, FiltroAviso.Click, FiltroErrores.Click, FiltroNoPagados.Click, FiltroPagados.Click
        Dim rbBoton As RadioButton
        Dim lb_ret As Boolean

        rbBoton = sender

        lb_ret = FiltrarRegistros(rbBoton.TabIndex, cbxCompania)
        If bwflag And lb_ret Then
            strSigCompa = cbxCompania.Items.Item(cbxCompania.SelectedIndex)
            RefrescarGrid(frmInstSuplidos.dtpDesde.Value, dtpHasta.Value, strSigCompa)
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
        If lvwSuplidos.Items.Count > 0 Then
            For Each objlistitem In lvwSuplidos.Items
                If objlistitem.Tag <> "1" Then
                    If objlistitem.Text <> "" And objlistitem.Text <> "No Existe" Then
                        Call ColorListItem(objlistitem, Color.Green)
                        objlistitem.Checked = blnCheck
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub RefrescarGrid(ByRef FecDes As Date, ByRef FecHas As Date, ByRef compa As String)

        'JCLopez_i

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql entera
        Dim FecFiltro As Date ' Fecha a filtar
        Dim objlistitem As ListViewItem ' Objeto con los registros del grid
        Dim objListSubItem As ListViewItem.ListViewSubItem ' Objeto con las columnas del grid
        Dim strFrom As String ' Parte From de la Sql
        Dim strOrderBy As String ' Parte Order By de la Sql

        ' Establecemos la selección de los campos con los que vamos a
        ' trabajar de la tabla de pagos de asistencia
        '
        strSqlSel = "SELECT Compa = '" & strCodcia & "', " & "       SuplidosAsistencia.T7_codsin, " & "       SuplidosAsistencia.T7_refer, " & "       SuplidosAsistencia.T7_codram, " & "       SuplidosAsistencia.T7_numpol, " & "       SuplidosAsistencia.T7_estado, " & "       SuplidosAsistencia.T7_fgraba, " & "       SuplidosAsistencia.Fichero, " & "       SuplidosAsistencia.FechaProceso, " & "       SuplidosAsistencia.T7_FacturaIP, " & "       SuplidosAsistencia.T7_FechaFactura"

        ' Añadimos el From de la Sql
        '
        strFrom = " From SuplidosAsistencia, Snsinies "

        ' Añadimos la Where de la Sql
        '
        strWhere = " Where (SuplidosAsistencia.T7_Codsin *= Snsinies.Codsin) and SuplidosAsistencia.T7_Codcia ='" & strCodcia & "'"

        ' Añadimos la parte de la Where que filtrará los registros en
        ' función del tipo de fecha que hayamos seleccionado
        '
        strWhereMas = " And SuplidosAsistencia.T7_Fgraba BETWEEN '" & objUtiles.FormatoFechaSQL(FecDes, False, False) & "' AND '" & objUtiles.FormatoFechaSQL(FecHas, False, False) & "'"

        ' Añadimos la parte de la where que filtrará según el estado que el usuario
        ' haya podido seleccionar.
        '
        If strFiltro <> "T" Then
            strWhereMas = strWhereMas & " And SuplidosAsistencia.T7_ESTADO = '" & strFiltro & "'"
        End If

        ' Añadimos la ordenación
        '
        strOrderBy = " Order By SuplidosAsistencia.T7_Codsin, SuplidosAsistencia.T7_Refer"

        strsql = strSqlSel & strFrom & strWhere & strWhereMas & strOrderBy

        ' Establece origen de datos para Crystal Reports
        '
        strSQLCR = strsql

        Call CargarListView_suplidos(frmInstSuplidos.lvwSuplidos, strsql, "", "T7_CODSIN", "T7_REFER", "T7_CODRAM", "T7_NUMPOL", "T7_ESTADO", "T7_FGRABA", "FICHERO", "FECHAPROCESO", "T7_FACTURAIP", "T7_FECHAFACTURA")


        frmInstanciaPrincipal.stbEstado.Panels(2).Text = Format(frmInstSuplidos.lvwSuplidos.Items.Count, "##,##0") & " Suplidos"

        ' Poner atenuados aquellos siniestros que ya esten procesados
        ' o con el color identificativo para aquellos que tengan errores o avisos
        '
        For Each objlistitem In frmInstSuplidos.lvwSuplidos.Items
            objListSubItem = objlistitem.SubItems.Item(frmInstSuplidos.T7_ESTADO.Index)

            Select Case objListSubItem.Text
                Case "P"
                    'objListItem.ListView.Enabled = False
                    objlistitem.Tag = "1"
                    frmInstSuplidos.ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro ya está Procesado")
                    Call ColorListItem(objlistitem, Color.Gray)
                Case "X"
                    objlistitem.Tag = "0"
                    frmInstSuplidos.ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro está pendiente de procesar")
                    Call ColorListItem(objlistitem, Color.Green)
                Case "W", "A"
                    objlistitem.Tag = "0"
                    frmInstSuplidos.ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro tiene avisos pendientes de resolución")
                    Call ColorListItem(objlistitem, Color.DarkGoldenrod)
                Case "E"
                    objlistitem.Tag = "0"
                    frmInstSuplidos.ttipAyuda.SetToolTip(objlistitem.ListView, "Este Siniestro tiene mensajes de error del proceso")
                    Call ColorListItem(objlistitem, Color.Red)
                Case Else
                    objlistitem.Tag = "0"
                    Call ColorListItem(objlistitem, Color.Gray)
            End Select

        Next objlistitem

        Exit Sub
RefrescarGrid_Error:
        MsgBox("Error refrescando datos", MsgBoxStyle.Critical)
    End Sub

    Private Sub cbAvisos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAvisos.Click
        On Error GoTo cbAvisos_Error

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim strReferencia As String
        Dim claseSiniestros As clsSiniestro_NET
        Dim frmInstanciaErrores As New frmVisorErrores

        ' Comprobamos que se haya seleccionados
        If lvwSuplidos.CheckedItems.Count > 0 Then
            ' Comprobamos el estado de la referencia seleccionada
            If lvwSuplidos.CheckedItems(0).SubItems(frmInstSuplidos.T7_ESTADO.Index).Text = "W" Or _
               lvwSuplidos.CheckedItems(0).SubItems(frmInstSuplidos.T7_ESTADO.Index).Text = "A" Or _
               lvwSuplidos.CheckedItems(0).SubItems(frmInstSuplidos.T7_ESTADO.Index).Text = "E" Then

                If lvwSuplidos.CheckedItems(0).Tag = "1" Then Exit Sub

                frmInstanciaErrores.Show()
                objlistitem = lvwSuplidos.CheckedItems(0)
                strReferencia = lvwSuplidos.CheckedItems(0).SubItems(frmInstSuplidos.T7_REFER.Index).Text
                frmInstanciaErrores.MostrarErrores(strReferencia)
            ElseIf lvwSuplidos.CheckedItems(0).SubItems(frmInstSuplidos.T7_ESTADO.Index).Text = "P" Then
                claseBDSuplidos.BDAuxRecord = claseSiniestros.Siniestro(lvwSuplidos.CheckedItems(0).SubItems(frmInstSuplidos.T7_CODSIN.Index).Text, True, UsuaApli)
            Else
                MsgBox("La referencia seleccionada no tiene Avisos/Errores", MsgBoxStyle.Information)
            End If
        Else
            '/* MUL  si no se selecciona ninguno se muestra el historial.
            'strError = "4007"
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
        Imprimir_suplidos()
    End Sub

End Class
