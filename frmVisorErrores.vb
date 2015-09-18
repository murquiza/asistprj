Imports System.data
Imports System.Data.SqlClient

Public Class frmVisorErrores


    Inherits System.Windows.Forms.Form

    Public strReferencia As String
    'Public strCodCompania As String


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
    Friend WithEvents cbxTipoErrores As System.Windows.Forms.ComboBox
    Friend WithEvents s As System.Windows.Forms.Label
    Friend WithEvents dtpFechaInicio As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFechaFin As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFechaFin As System.Windows.Forms.Label
    Friend WithEvents lblTipoError As System.Windows.Forms.Label
    Friend WithEvents dbBusqueda As System.Windows.Forms.GroupBox
    Friend WithEvents cbImprimir As System.Windows.Forms.Button
    Friend WithEvents cbVolver As System.Windows.Forms.Button
    Friend WithEvents cbBuscar As System.Windows.Forms.Button
    Friend WithEvents cbEliminar As System.Windows.Forms.Button
    Public WithEvents CR2 As AxCrystal.AxCrystalReport
    Friend WithEvents lvwErrores As System.Windows.Forms.ListView
    Friend WithEvents Coderr As System.Windows.Forms.ColumnHeader
    Friend WithEvents Referencia As System.Windows.Forms.ColumnHeader
    Friend WithEvents Codsin As System.Windows.Forms.ColumnHeader
    Friend WithEvents Errores As System.Windows.Forms.ColumnHeader
    Friend WithEvents texto As System.Windows.Forms.ColumnHeader
    Friend WithEvents Fecgra As System.Windows.Forms.ColumnHeader
    Friend WithEvents RamoRel As System.Windows.Forms.ColumnHeader
    Friend WithEvents PolizaRel As System.Windows.Forms.ColumnHeader
    Friend WithEvents siniestroRel As System.Windows.Forms.ColumnHeader
    Friend WithEvents TipoObjetoRel As System.Windows.Forms.ColumnHeader
    Friend WithEvents Numero As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmVisorErrores))
        Me.cbxTipoErrores = New System.Windows.Forms.ComboBox
        Me.s = New System.Windows.Forms.Label
        Me.dtpFechaInicio = New System.Windows.Forms.DateTimePicker
        Me.dtpFechaFin = New System.Windows.Forms.DateTimePicker
        Me.lblFechaFin = New System.Windows.Forms.Label
        Me.lblTipoError = New System.Windows.Forms.Label
        Me.cbBuscar = New System.Windows.Forms.Button
        Me.dbBusqueda = New System.Windows.Forms.GroupBox
        Me.cbEliminar = New System.Windows.Forms.Button
        Me.cbImprimir = New System.Windows.Forms.Button
        Me.cbVolver = New System.Windows.Forms.Button
        Me.CR2 = New AxCrystal.AxCrystalReport
        Me.lvwErrores = New System.Windows.Forms.ListView
        Me.Coderr = New System.Windows.Forms.ColumnHeader
        Me.Referencia = New System.Windows.Forms.ColumnHeader
        Me.Codsin = New System.Windows.Forms.ColumnHeader
        Me.Errores = New System.Windows.Forms.ColumnHeader
        Me.texto = New System.Windows.Forms.ColumnHeader
        Me.Fecgra = New System.Windows.Forms.ColumnHeader
        Me.RamoRel = New System.Windows.Forms.ColumnHeader
        Me.PolizaRel = New System.Windows.Forms.ColumnHeader
        Me.siniestroRel = New System.Windows.Forms.ColumnHeader
        Me.TipoObjetoRel = New System.Windows.Forms.ColumnHeader
        Me.Numero = New System.Windows.Forms.ColumnHeader
        Me.dbBusqueda.SuspendLayout()
        CType(Me.CR2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbxTipoErrores
        '
        Me.cbxTipoErrores.Location = New System.Drawing.Point(224, 40)
        Me.cbxTipoErrores.Name = "cbxTipoErrores"
        Me.cbxTipoErrores.TabIndex = 0
        '
        's
        '
        Me.s.ForeColor = System.Drawing.Color.RoyalBlue
        Me.s.Location = New System.Drawing.Point(16, 24)
        Me.s.Name = "s"
        Me.s.Size = New System.Drawing.Size(80, 16)
        Me.s.TabIndex = 1
        Me.s.Text = "Fecha Inicio:"
        '
        'dtpFechaInicio
        '
        Me.dtpFechaInicio.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFechaInicio.Location = New System.Drawing.Point(16, 40)
        Me.dtpFechaInicio.Name = "dtpFechaInicio"
        Me.dtpFechaInicio.Size = New System.Drawing.Size(88, 21)
        Me.dtpFechaInicio.TabIndex = 2
        '
        'dtpFechaFin
        '
        Me.dtpFechaFin.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFechaFin.Location = New System.Drawing.Point(120, 40)
        Me.dtpFechaFin.Name = "dtpFechaFin"
        Me.dtpFechaFin.Size = New System.Drawing.Size(88, 21)
        Me.dtpFechaFin.TabIndex = 3
        '
        'lblFechaFin
        '
        Me.lblFechaFin.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblFechaFin.Location = New System.Drawing.Point(120, 24)
        Me.lblFechaFin.Name = "lblFechaFin"
        Me.lblFechaFin.Size = New System.Drawing.Size(72, 16)
        Me.lblFechaFin.TabIndex = 4
        Me.lblFechaFin.Text = "Fecha Fin:"
        '
        'lblTipoError
        '
        Me.lblTipoError.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblTipoError.Location = New System.Drawing.Point(224, 24)
        Me.lblTipoError.Name = "lblTipoError"
        Me.lblTipoError.Size = New System.Drawing.Size(100, 16)
        Me.lblTipoError.TabIndex = 5
        Me.lblTipoError.Text = "Tipo Error:"
        '
        'cbBuscar
        '
        Me.cbBuscar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBuscar.Image = CType(resources.GetObject("cbBuscar.Image"), System.Drawing.Image)
        Me.cbBuscar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbBuscar.Location = New System.Drawing.Point(360, 16)
        Me.cbBuscar.Name = "cbBuscar"
        Me.cbBuscar.Size = New System.Drawing.Size(64, 56)
        Me.cbBuscar.TabIndex = 7
        Me.cbBuscar.Text = "Buscar"
        Me.cbBuscar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'dbBusqueda
        '
        Me.dbBusqueda.Controls.Add(Me.cbBuscar)
        Me.dbBusqueda.Controls.Add(Me.lblFechaFin)
        Me.dbBusqueda.Controls.Add(Me.lblTipoError)
        Me.dbBusqueda.Controls.Add(Me.cbxTipoErrores)
        Me.dbBusqueda.Controls.Add(Me.s)
        Me.dbBusqueda.Controls.Add(Me.dtpFechaInicio)
        Me.dbBusqueda.Controls.Add(Me.dtpFechaFin)
        Me.dbBusqueda.ForeColor = System.Drawing.Color.RoyalBlue
        Me.dbBusqueda.Location = New System.Drawing.Point(0, 464)
        Me.dbBusqueda.Name = "dbBusqueda"
        Me.dbBusqueda.Size = New System.Drawing.Size(440, 80)
        Me.dbBusqueda.TabIndex = 8
        Me.dbBusqueda.TabStop = False
        Me.dbBusqueda.Text = "Busqueda avanzada"
        '
        'cbEliminar
        '
        Me.cbEliminar.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbEliminar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbEliminar.Image = CType(resources.GetObject("cbEliminar.Image"), System.Drawing.Image)
        Me.cbEliminar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbEliminar.Location = New System.Drawing.Point(512, 480)
        Me.cbEliminar.Name = "cbEliminar"
        Me.cbEliminar.Size = New System.Drawing.Size(64, 56)
        Me.cbEliminar.TabIndex = 9
        Me.cbEliminar.Text = "Eliminar"
        Me.cbEliminar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbImprimir
        '
        Me.cbImprimir.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbImprimir.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbImprimir.Image = CType(resources.GetObject("cbImprimir.Image"), System.Drawing.Image)
        Me.cbImprimir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbImprimir.Location = New System.Drawing.Point(576, 480)
        Me.cbImprimir.Name = "cbImprimir"
        Me.cbImprimir.Size = New System.Drawing.Size(64, 56)
        Me.cbImprimir.TabIndex = 10
        Me.cbImprimir.Text = "Imprimir"
        Me.cbImprimir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbVolver
        '
        Me.cbVolver.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cbVolver.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbVolver.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbVolver.Image = CType(resources.GetObject("cbVolver.Image"), System.Drawing.Image)
        Me.cbVolver.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbVolver.Location = New System.Drawing.Point(640, 480)
        Me.cbVolver.Name = "cbVolver"
        Me.cbVolver.Size = New System.Drawing.Size(64, 56)
        Me.cbVolver.TabIndex = 11
        Me.cbVolver.Text = "Volver"
        Me.cbVolver.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CR2
        '
        Me.CR2.Enabled = True
        Me.CR2.Location = New System.Drawing.Point(472, 488)
        Me.CR2.Name = "CR2"
        Me.CR2.OcxState = CType(resources.GetObject("CR2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CR2.Size = New System.Drawing.Size(28, 28)
        Me.CR2.TabIndex = 66
        '
        'lvwErrores
        '
        Me.lvwErrores.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvwErrores.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Coderr, Me.Referencia, Me.Codsin, Me.Errores, Me.texto, Me.Fecgra, Me.RamoRel, Me.PolizaRel, Me.siniestroRel, Me.TipoObjetoRel, Me.Numero})
        Me.lvwErrores.FullRowSelect = True
        Me.lvwErrores.GridLines = True
        Me.lvwErrores.HideSelection = False
        Me.lvwErrores.Location = New System.Drawing.Point(0, 0)
        Me.lvwErrores.Name = "lvwErrores"
        Me.lvwErrores.Size = New System.Drawing.Size(712, 456)
        Me.lvwErrores.TabIndex = 71
        Me.lvwErrores.View = System.Windows.Forms.View.Details
        '
        'Coderr
        '
        Me.Coderr.Text = "Error"
        Me.Coderr.Width = 40
        '
        'Referencia
        '
        Me.Referencia.Text = "Referencia"
        Me.Referencia.Width = 80
        '
        'Codsin
        '
        Me.Codsin.Text = "Expediente"
        Me.Codsin.Width = 80
        '
        'Errores
        '
        Me.Errores.Text = "Tipo"
        Me.Errores.Width = 40
        '
        'texto
        '
        Me.texto.Text = "Descripción"
        Me.texto.Width = 300
        '
        'Fecgra
        '
        Me.Fecgra.Text = "F. Grabación"
        Me.Fecgra.Width = 80
        '
        'RamoRel
        '
        Me.RamoRel.Text = "Ramo Rel."
        Me.RamoRel.Width = 80
        '
        'PolizaRel
        '
        Me.PolizaRel.Text = "Poliza Rel."
        Me.PolizaRel.Width = 80
        '
        'siniestroRel
        '
        Me.siniestroRel.Text = "Siniestro Rel."
        Me.siniestroRel.Width = 80
        '
        'TipoObjetoRel
        '
        Me.TipoObjetoRel.Text = "Tipo Objeto Rel."
        Me.TipoObjetoRel.Width = 100
        '
        'Numero
        '
        Me.Numero.Text = "Numero"
        Me.Numero.Width = 0
        '
        'frmVisorErrores
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.CancelButton = Me.cbVolver
        Me.ClientSize = New System.Drawing.Size(712, 546)
        Me.ControlBox = False
        Me.Controls.Add(Me.lvwErrores)
        Me.Controls.Add(Me.CR2)
        Me.Controls.Add(Me.cbVolver)
        Me.Controls.Add(Me.cbImprimir)
        Me.Controls.Add(Me.cbEliminar)
        Me.Controls.Add(Me.dbBusqueda)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmVisorErrores"
        Me.Text = "Històrico de Avisos/Errores de aperturas de Asistencia"
        Me.TopMost = True
        Me.dbBusqueda.ResumeLayout(False)
        CType(Me.CR2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmVisorErrores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InicioApp()
    End Sub

    Private Sub InicioApp()
        'MostrarErrores("")
        CargarComboErrores()
    End Sub

    Public Sub CargarComboErrores()
        On Error GoTo CargarComboErrores_ERR
        ' Declaraciones
        '
        Dim strSqlAux As String ' Instrucción Sql
        Dim lngResult As Integer ' Número de registros devueltos poa la consulta

        ' Establecemos el primer valor que tendra el combo
        '
        cbxTipoErrores.Items.Add("Todos")

        strSqlAux = "Select Distinct coderr From mpasihisterror Where Proceso = '" & strIdProceso & "'"

        claseBDAperturas.BDAuxRecord.Open(strSqlAux, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        Do While Not claseBDAperturas.BDAuxRecord.EOF
            cbxTipoErrores.Items.Add(claseBDAperturas.BDAuxRecord.Fields("Coderr").Value)
            claseBDAperturas.BDAuxRecord.MoveNext()
        Loop

        cbxTipoErrores.Text = "Todos"
        claseBDAperturas.BDAuxRecord.Close()
        Exit Sub

CargarComboErrores_ERR:
        If claseBDAperturas.BDWorkConnect.Errors.Count > 0 Then
            If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        End If

        MsgBox("Se ha producido un error en la carga del histórico de errores", MsgBoxStyle.Critical)
    End Sub


    Public Sub MostrarErrores(Optional ByVal strReferenciaAux As String = "")
        On Error GoTo MostrarErrores_Error
        ' Declaraciones
        '
        Dim strSQL, Donde, Orden As String ' Instrucción Sql
        Dim dtFechaInicio, dtFechaFin As DateTime

        If IsNothing(strReferenciaAux) Then
            strReferenciaAux = ""
        End If

        strReferencia = strReferenciaAux
        claseBDAperturas.BDSystemConnect.Errors.Clear()

        dtFechaInicio = dtpFechaInicio.Value
        dtFechaFin = dtpFechaFin.Value

        If Not IsDate(dtFechaInicio) Then Exit Sub

        If dtpFechaInicio.Value > dtpFechaFin.Value Then
            MsgBox("El rango de fechas especificado es erroneo. La fecha de inicio no puede ser mayor que la fecha final.", MsgBoxStyle.Information)
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        ' Si el recordset está abierto lo cerramos
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()

        If strReferencia = "" Then
            strSQL = "Select Error = Coderr, Referencia = Referencia, Expediente = Codsin, Tipo = Errores, Descripcion = Texto, FGrabación = Fecgra, RamoRel = isnull(RamoRel,0), PolizaRel = isnull(PolizaRel,''), SiniestroRel = isnull(siniestroRel,''), TipoObjetoRel = isnull(TipoObjetoRel,''), Numero = isnull(numero,0) " & _
                     "From mpAsiHistError"
            Donde = " Where Fecgra Between '" & claseUtilidadesAperturas.FormatoFechaSQL(dtFechaInicio, False, False) & "' and '" & claseUtilidadesAperturas.FormatoFechaSQL(dtFechaFin, False, False) & _
                    IIf(cbxTipoErrores.Text <> "Todos" And cbxTipoErrores.Text <> "", "' and Coderr = '" & Trim(cbxTipoErrores.Text), "") & _
                    "' And Proceso = '" & strIdProceso & _
                    "' and Codcia = '" & strCodCia & "'"
            Orden = " Order By Fecgra, Expediente, Referencia "
        Else
            strSQL = "Select Error = Coderr, Referencia = '', Expediente = Codsin, Tipo = Errores, Descripcion = Texto, F_Grabación = Fecgra, RamoRel = isnull(RamoRel,0), PolizaRel = isnull(PolizaRel,''), SiniestroRel = isnull(siniestroRel,''), TipoObjetoRel = isnull(TipoObjetoRel,''), Numero = isnull(numero,0) " & _
                     "From mpAsiHistError"
            Donde = " Where Referencia = '" & strReferencia & "' And Proceso = '" & strIdProceso & "' and Codcia ='" & strCodCia & "'"
            Orden = " Order By Fecgra, Expediente, Referencia "
        End If
        strSQLVisor = strSQL & Donde & Orden

        'claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        'dgErrores.SetDataBinding(claseBDAperturas.BDWorkRecord.DataSource, claseBDAperturas.BDWorkRecord.DataMember)
        CargaGrid(strSQLVisor, "Coderr", "Referencia", "Expediente", "Tipo", "Texto", "Fecgra", "RamoRel", "PolizaRel", "siniestroRel", "TipoObjetoRel", "Numero")

        Cursor.Current = Cursors.Default
        Exit Sub

MostrarErrores_Error:
        Cursor.Current = Cursors.Default
        If claseBDAperturas.BDWorkConnect.Errors.Count > 0 Then
            If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        End If
        Cursor = System.Windows.Forms.Cursors.Default
        MsgBox("Se ha producido un error en la carga del hisórico de errores de aperturas de asistencia", MsgBoxStyle.Critical)
    End Sub

    Private Sub CargaGrid(ByVal strCadenaSQL As String, ByVal ParamArray Campos() As Object)
        '''''Dim Conexion As SqlConnection
        '''''Dim daAdaptador As SqlDataAdapter
        '''''Dim dsDatos As New DataSet
        ''''''Dim strParametro As String

        ''''''/*MUL T-19908 INI
        ''''''strParametro = Trim(Microsoft.VisualBasic.Command)

        ''''''If strParametro = "PM" Then
        ''''''    Conexion = New SqlConnection("Data Source=INTEGRIX;Initial Catalog=BDMantenimiento;User ID=Usuarios;Password=All")
        ''''''Else
        ''''''    Conexion = New SqlConnection("Data Source=OBELIX;Initial Catalog=BDMutuae;User ID=Usuarios;Password=All")
        ''''''End If
        '''''Conexion = New SqlConnection("Data Source=" & claseBDAperturas.BDServer & _
        '''''               ";Initial Catalog=" & claseBDAperturas.BDName & _
        '''''               ";User ID=" & claseBDAperturas.UserApp() & _
        '''''               ";Password=" & claseBDAperturas.pwdApp)
        ''''''/*MUL T-19908 FIN

        '''''daAdaptador = New SqlDataAdapter(strCadenaSQL, Conexion)
        '''''daAdaptador.Fill(dsDatos)
        '''''dgErrores.DataSource = dsDatos.Tables(0)
        Dim objListItem As ListViewItem
        Dim intContador As Integer
        Dim numFila As Long
        Dim strCodSiniestro As String
        Dim maxCampos As Long
        Dim minCampos As Long

        On Error GoTo CargaGrid_ERROR

        Cursor.Current = Cursors.WaitCursor

        numFila = 0
        minCampos = LBound(Campos)
        maxCampos = UBound(Campos)

        lvwErrores.Items.Clear()
        claseBDAperturas.BDWorkConnect.Errors.Clear()
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        claseBDAperturas.BDWorkRecord.Open(strCadenaSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDAperturas.BDWorkRecord.EOF Then
            Do
                For intContador = minCampos To maxCampos
                    If intContador = minCampos Then
                        lvwErrores.Items.Add("" & claseBDAperturas.BDWorkRecord.Fields(intContador).Value)
                    Else
                        lvwErrores.Items(numFila).SubItems.Add("" & claseBDAperturas.BDWorkRecord.Fields(intContador).Value)
                    End If
                Next
                lvwErrores.Items(numFila).Selected = False
                If Not claseBDAperturas.BDWorkRecord.EOF Then claseBDAperturas.BDWorkRecord.MoveNext()

                numFila = numFila + 1
            Loop Until claseBDAperturas.BDWorkRecord.EOF

        Else
            MsgBox("No existen datos para mostrar con los criterios de selección asignados.", MsgBoxStyle.Exclamation)
        End If

        claseBDAperturas.BDWorkRecord.Close()
        Cursor.Current = Cursors.Default
        Exit Sub

CargaGrid_ERROR:
        Cursor.Current = Cursors.Default
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
    End Sub

    Private Sub cbVolver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVolver.Click
        Me.Close()
    End Sub

    Private Sub cbBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBuscar.Click
        MostrarErrores("")
    End Sub

    Private Sub cbImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbImprimir.Click
        On Error GoTo ImprimirClick_Err

        ' Declaraciones
        '
        Dim Acum As String
        Dim rsLocal As ADODB.Recordset
        Dim booltransaccion As Boolean
        Dim s_strSql As String

        booltransaccion = False
        Cursor.Current = Cursors.WaitCursor

        'Generación de nuevos registros temporales en mdpImpAperAsis
        '
        With claseBDAperturas.BDWorkConnect
            booltransaccion = True
            If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
            s_strSql = Replace(strSQLVisor, ", Numero = isnull(numero,0)", "")
            .BeginTrans()
            .Execute("DELETE FROM mdpImpAvisosErrores")
            '.Execute("INSERT INTO mdpImpAvisosErrores (Coderr,Referencia,Codsin,Errores,Texto,Fecgra,RamoRel,PolizaRel,SiniestroRel,TipoObjetorel)" & strSQLVisor)
            .Execute("INSERT INTO mdpImpAvisosErrores (Coderr,Referencia,Codsin,Errores,Texto,Fecgra,RamoRel,PolizaRel,SiniestroRel,TipoObjetorel)" & s_strSql)
            .CommitTrans()
            booltransaccion = False
        End With

        ' Fijamos los parametros y contenidos de impresión y ejecutamos
        '
        'CR2.ReportFileName = PathReports & "SiniMensajes.rpt"
        CR2.set_ParameterFields(1, "FecDesde;" & CStr(dtpFechaInicio.Value) & ";TRUE")
        CR2.set_ParameterFields(2, "FecHasta;" & CStr(dtpFechaFin.Value) & ";TRUE")
        CR2.set_ParameterFields(3, "Cia;" & Trim(strNombreCompa) & ";TRUE")

        CR2.Connect = "DSN = ariesSqlServer.dsn;UID = allUsers;PWD = all;DSQ = Aries"
        CR2.Destination = Crystal.DestinationConstants.crptToPrinter
        CR2.Action = 1

        Cursor.Current = Cursors.Default
        Exit Sub

ImprimirClick_Err:
        If booltransaccion = True Then
            claseBDAperturas.BDWorkConnect.RollbackTrans()
        End If
        Cursor.Current = Cursors.Default
        MsgBox("Error en la impresión del Informe. Avise a Informática", vbCritical + vbOKOnly, "Impresión Errores de Importación")
    End Sub


    Private Sub cbEliminar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbEliminar.Click
        On Error GoTo cbEliminar_Err

        ' Declaraciones
        '
        Dim strsql As String
        Dim lngResult As Integer
        Dim lviewSelect As ListView.SelectedListViewItemCollection
        Dim item As ListViewItem
        Dim ls_Borrar As String

        If Me.lvwErrores.SelectedItems.Count <= 0 Then
            MsgBox("Ha de seleccionar el aviso / error que desea eliminar", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If MsgBox("¿ Esta seguro de querer eliminar los datos seleccionados ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.Yes Then
            Cursor.Current = Cursors.WaitCursor

            lviewSelect = Me.lvwErrores.SelectedItems
            'me.lvwErrores.SelectedListViewItemCollection.
            ls_Borrar = ""
            For Each item In lviewSelect
                'referencia, numero ,proceso
                If ls_Borrar.Length = 0 Then
                    ls_Borrar = "( referencia = '" & item.SubItems(1).Text & "' and numero = " & item.SubItems(10).Text & " and proceso = '" & strIdProceso & "')"
                Else
                    ls_Borrar += " or ( referencia = '" & item.SubItems(1).Text & "' and numero = " & item.SubItems(10).Text & " and proceso = '" & strIdProceso & "')"
                End If
            Next
            ls_Borrar = " where (" & ls_Borrar & ")"
            strsql = "Delete From mpAsihistError " + ls_Borrar

            claseBDAperturas.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
            claseBDAperturas.BDComand.CommandText = strsql
            claseBDAperturas.BDComand.ActiveConnection = claseBDAperturas.BDWorkConnect
            claseBDAperturas.BDComand.Execute(lngResult)
        End If
        MostrarErrores("")
        Cursor.Current = Cursors.Default
        Exit Sub

cbEliminar_Err:
        Cursor.Current = Cursors.Default
        'If claseBDSuplidos.BDWorkConnect.Errors.Count = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkConnect.RollbackTrans()
        MsgBox("Se ha producido un error en la Base de Datos. Los registros seleccionados no han sido borrados", MsgBoxStyle.Exclamation)
    End Sub
End Class
