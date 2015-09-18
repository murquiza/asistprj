Public Class frmPrincipalExportacion
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
    Friend WithEvents cbxCompania As System.Windows.Forms.ComboBox
    Friend WithEvents lbCompaniaAsistencia As System.Windows.Forms.Label
    Friend WithEvents pbLogoMutua As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tipAyuda As System.Windows.Forms.ToolTip
    Friend WithEvents cbSalir As System.Windows.Forms.Button
    Friend WithEvents stbEstado As System.Windows.Forms.StatusBar
    Friend WithEvents sbPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents prbProgreso As System.Windows.Forms.ProgressBar
    Friend WithEvents cbAcumulado As System.Windows.Forms.Button
    Friend WithEvents cbDiario As System.Windows.Forms.Button
    Friend WithEvents dtpFechaEfecto As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDirDestino As System.Windows.Forms.Label
    Friend WithEvents fbdDirectorio As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents cbSelDirectorio As System.Windows.Forms.Button
    Friend WithEvents lbxCompania As System.Windows.Forms.ListBox
    Friend WithEvents sbPanel4 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents picTest As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrincipalExportacion))
        Me.cbxCompania = New System.Windows.Forms.ComboBox
        Me.lbCompaniaAsistencia = New System.Windows.Forms.Label
        Me.pbLogoMutua = New System.Windows.Forms.PictureBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpFechaEfecto = New System.Windows.Forms.DateTimePicker
        Me.cbSelDirectorio = New System.Windows.Forms.Button
        Me.lblDirDestino = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.tipAyuda = New System.Windows.Forms.ToolTip(Me.components)
        Me.cbAcumulado = New System.Windows.Forms.Button
        Me.cbDiario = New System.Windows.Forms.Button
        Me.cbSalir = New System.Windows.Forms.Button
        Me.stbEstado = New System.Windows.Forms.StatusBar
        Me.sbPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel4 = New System.Windows.Forms.StatusBarPanel
        Me.prbProgreso = New System.Windows.Forms.ProgressBar
        Me.fbdDirectorio = New System.Windows.Forms.FolderBrowserDialog
        Me.lbxCompania = New System.Windows.Forms.ListBox
        Me.picTest = New System.Windows.Forms.PictureBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbxCompania
        '
        Me.cbxCompania.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxCompania.Location = New System.Drawing.Point(232, 16)
        Me.cbxCompania.Name = "cbxCompania"
        Me.cbxCompania.Size = New System.Drawing.Size(264, 24)
        Me.cbxCompania.TabIndex = 6
        '
        'lbCompaniaAsistencia
        '
        Me.lbCompaniaAsistencia.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbCompaniaAsistencia.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lbCompaniaAsistencia.Location = New System.Drawing.Point(72, 16)
        Me.lbCompaniaAsistencia.Name = "lbCompaniaAsistencia"
        Me.lbCompaniaAsistencia.Size = New System.Drawing.Size(152, 24)
        Me.lbCompaniaAsistencia.TabIndex = 5
        Me.lbCompaniaAsistencia.Text = "Compañía Asistencia:"
        Me.lbCompaniaAsistencia.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pbLogoMutua
        '
        Me.pbLogoMutua.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pbLogoMutua.Image = CType(resources.GetObject("pbLogoMutua.Image"), System.Drawing.Image)
        Me.pbLogoMutua.Location = New System.Drawing.Point(8, 8)
        Me.pbLogoMutua.Name = "pbLogoMutua"
        Me.pbLogoMutua.Size = New System.Drawing.Size(48, 48)
        Me.pbLogoMutua.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbLogoMutua.TabIndex = 4
        Me.pbLogoMutua.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpFechaEfecto)
        Me.GroupBox1.Controls.Add(Me.cbSelDirectorio)
        Me.GroupBox1.Controls.Add(Me.lblDirDestino)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.GroupBox1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox1.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(560, 136)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Datos Generación Fichero"
        '
        'dtpFechaEfecto
        '
        Me.dtpFechaEfecto.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.dtpFechaEfecto.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFechaEfecto.Location = New System.Drawing.Point(136, 24)
        Me.dtpFechaEfecto.Name = "dtpFechaEfecto"
        Me.dtpFechaEfecto.Size = New System.Drawing.Size(88, 21)
        Me.dtpFechaEfecto.TabIndex = 4
        '
        'cbSelDirectorio
        '
        Me.cbSelDirectorio.Image = CType(resources.GetObject("cbSelDirectorio.Image"), System.Drawing.Image)
        Me.cbSelDirectorio.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.cbSelDirectorio.Location = New System.Drawing.Point(528, 48)
        Me.cbSelDirectorio.Name = "cbSelDirectorio"
        Me.cbSelDirectorio.Size = New System.Drawing.Size(24, 24)
        Me.cbSelDirectorio.TabIndex = 3
        Me.tipAyuda.SetToolTip(Me.cbSelDirectorio, "Seleccione el directorio de destino")
        '
        'lblDirDestino
        '
        Me.lblDirDestino.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDirDestino.Location = New System.Drawing.Point(136, 48)
        Me.lblDirDestino.Name = "lblDirDestino"
        Me.lblDirDestino.Size = New System.Drawing.Size(384, 72)
        Me.lblDirDestino.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 23)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Directorio de destino:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Fecha Efecto:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbAcumulado
        '
        Me.cbAcumulado.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbAcumulado.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbAcumulado.Image = CType(resources.GetObject("cbAcumulado.Image"), System.Drawing.Image)
        Me.cbAcumulado.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbAcumulado.Location = New System.Drawing.Point(80, 208)
        Me.cbAcumulado.Name = "cbAcumulado"
        Me.cbAcumulado.Size = New System.Drawing.Size(72, 56)
        Me.cbAcumulado.TabIndex = 8
        Me.cbAcumulado.Text = "Acumulado"
        Me.cbAcumulado.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbDiario
        '
        Me.cbDiario.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbDiario.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbDiario.Image = CType(resources.GetObject("cbDiario.Image"), System.Drawing.Image)
        Me.cbDiario.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbDiario.Location = New System.Drawing.Point(8, 208)
        Me.cbDiario.Name = "cbDiario"
        Me.cbDiario.Size = New System.Drawing.Size(72, 56)
        Me.cbDiario.TabIndex = 9
        Me.cbDiario.Text = "Diario"
        Me.cbDiario.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbSalir
        '
        Me.cbSalir.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbSalir.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbSalir.Image = CType(resources.GetObject("cbSalir.Image"), System.Drawing.Image)
        Me.cbSalir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbSalir.Location = New System.Drawing.Point(504, 208)
        Me.cbSalir.Name = "cbSalir"
        Me.cbSalir.Size = New System.Drawing.Size(64, 56)
        Me.cbSalir.TabIndex = 10
        Me.cbSalir.Text = "Salir"
        Me.cbSalir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'stbEstado
        '
        Me.stbEstado.Location = New System.Drawing.Point(0, 274)
        Me.stbEstado.Name = "stbEstado"
        Me.stbEstado.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbPanel1, Me.sbPanel2, Me.sbPanel3, Me.sbPanel4})
        Me.stbEstado.ShowPanels = True
        Me.stbEstado.Size = New System.Drawing.Size(576, 24)
        Me.stbEstado.TabIndex = 14
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
        Me.sbPanel3.Width = 99
        '
        'sbPanel4
        '
        Me.sbPanel4.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.sbPanel4.Width = 10
        '
        'prbProgreso
        '
        Me.prbProgreso.Location = New System.Drawing.Point(460, 283)
        Me.prbProgreso.Name = "prbProgreso"
        Me.prbProgreso.Size = New System.Drawing.Size(80, 8)
        Me.prbProgreso.TabIndex = 16
        Me.prbProgreso.Visible = False
        '
        'fbdDirectorio
        '
        Me.fbdDirectorio.RootFolder = System.Environment.SpecialFolder.MyComputer
        '
        'lbxCompania
        '
        Me.lbxCompania.Location = New System.Drawing.Point(224, 40)
        Me.lbxCompania.Name = "lbxCompania"
        Me.lbxCompania.Size = New System.Drawing.Size(312, 56)
        Me.lbxCompania.TabIndex = 17
        Me.lbxCompania.Visible = False
        '
        'picTest
        '
        Me.picTest.Cursor = System.Windows.Forms.Cursors.Help
        Me.picTest.Image = CType(resources.GetObject("picTest.Image"), System.Drawing.Image)
        Me.picTest.Location = New System.Drawing.Point(8, 8)
        Me.picTest.Name = "picTest"
        Me.picTest.Size = New System.Drawing.Size(48, 48)
        Me.picTest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picTest.TabIndex = 38
        Me.picTest.TabStop = False
        Me.picTest.Visible = False
        '
        'frmPrincipalExportacion
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(576, 298)
        Me.Controls.Add(Me.lbxCompania)
        Me.Controls.Add(Me.prbProgreso)
        Me.Controls.Add(Me.stbEstado)
        Me.Controls.Add(Me.cbSalir)
        Me.Controls.Add(Me.cbDiario)
        Me.Controls.Add(Me.cbAcumulado)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cbxCompania)
        Me.Controls.Add(Me.lbCompaniaAsistencia)
        Me.Controls.Add(Me.pbLogoMutua)
        Me.Controls.Add(Me.picTest)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrincipalExportacion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Siniestros: Area de Asistencia  -  Exportación Datos"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmPrincipalExportacion_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InicioApp()
    End Sub

    Private Sub InicioApp()
        'Dim ConecPruebas As String
        Dim strParametro As String
        Dim objExportar As New clsExportar_NET
        'Dim i As Short
        Dim lResultado As Boolean
        Dim directori As String
        Dim li_posAsterisco As Integer

        ' Creación de Objetos
        '
        'Set clses = New clsmdpSesion
        Cabecera = New ADODB.Recordset

        frmInstanciaPrincipal = Me
        '        IdProceso = "E"

        FicheroLog = objExportar.RutaLog()

        'strParametro = Microsoft.VisualBasic.Command
        'li_posAsterisco = InStr(strParametro, "*")
        'If li_posAsterisco > 0 Then
        '   strParametro = Trim(Mid(strParametro, 1, li_posAsterisco - 1))
        'End If


        claseBDExportacion.BDComand.CommandTimeout = 0
        'Select Case UCase(Mid$(Command$, 1, PosCom - 1))
        Select Case parametroApp  'strParametro
            Case "P"
                TipoEjecucion = "P"
            Case "PP"
                claseBDExportacion.ConnexionPruebas()
                TipoEjecucion = "P"
                picTest.Show()
                picTest.BringToFront()
                'mdpbd.BDComand.CommandTimeout = 300
            Case "PM"
                claseBDExportacion.ConnexionPruebas()
                TipoEjecucion = "M"
                picTest.Show()
                picTest.BringToFront()
                'mdpbd.BDComand.CommandTimeout = 300
                lblDirDestino.Text = claseBDExportacion.PathExport
            Case Else
                TipoEjecucion = "M"
        End Select
        ' Fin de la Modificación

        ' Asignación de valores y parametros iniciales
        '
        '@m001_i 07/07/2015  MAN-3451
        'dtpFechaEfecto.Value = System.DateTime.FromOADate(Today.ToOADate - 1)
        dtpFechaEfecto.Value = System.DateTime.FromOADate(Today.ToOADate)
        '@m001_f
        LlenarComboCias(cbxCompania, lbxCompania)

        cbxCompania.SelectedIndex = 0

        'directori = clses.GetParam("PathExport")
        directori = "K:\Siniestros\Asistencia\Export"

        'PathIconos = clses.GetParam("PathGraficos")
        PathIconos = "K:\Graficos\"

        NomIniFichero = "\AngelDiario" & fectxt

        'fectxt = dtpFechaEfecto.Value.Day & "" , "ddmmyyyy")
        fectxt = dtpFechaEfecto.Value.Day & dtpFechaEfecto.Value.Month & dtpFechaEfecto.Value.Year

        Call ObtenerLote()

        If TipoEjecucion = "P" Then
            If claseBDExportacion.BDNameSys = "Error" Then
                lResultado = objExportar.InsertarLog("Error crítico en la conexión a la BD.", True)
                End
            Else
                ProcesoProgramado()
                Me.Close()
                End
            End If
        End If
    End Sub

    Private Sub ProcesoProgramado()

        On Error GoTo ProcesoProgramado_Err

        ' Primero leemos de la tabla el modo de ejecución ( Acumulado/Diario )
        '
        Dim NombreFicheroZip As String
        Dim NombreFicheroZip1 As String
        Dim NombreFicheroZip3 As String
        Dim Comando As String
        Dim Ruta As String
        Dim v_zip As String
        Dim Path As String
        Dim PassWordZip As String
        Dim v_estado As String
        Dim idx_asistencia As Integer
        '/* MUL INI
        Dim spath As String
        Dim nomfich As String = " "
        Dim sext As String = " "
        '/* MUL FIN

        ' 25-6-2007. Eloi. Creación de variables

        Dim lobjExportar As clsExportar_NET
        Dim lResultado As Boolean

        ' 25-6-2007. Final Eloi

        lobjExportar = New clsExportar_NET
        'Rslocal = New ADODB.Recordset

        ' Asignacion de valores iniciales
        '
        v_estado = "N"
        UsuaApli = "Automata"

        ' Obtenemos los datos del tipo de proceso (Acumulado o Diaro) a ejecutar
        ' Si es día 2 de cada mes se ejecuta un Acumulado por petición de InterPartner
        ' si el dia 2 es domingo tengo que ejecutarlo en dia 1 sabado
        '
        If Now.Day = 2 Or (Now.Day = 1 And Now.DayOfWeek = DayOfWeek.Saturday) Then
            Modo = UCase("Acumulado")
        Else
            '/* MUL INI
            'claseBDExportacion.BDComand.ActiveConnection.ConnectionString = claseBDExportacion.BDWorkConnect.ConnectionString
            'claseBDExportacion.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
            'claseBDExportacion.BDComand.CommandText = "Select * From EjecucionProgramadaAsistencia"
            'Rslocal = claseBDExportacion.BDComand.Execute(lngResult)

            'Rslocal.MoveFirst()
            'Modo = Rslocal.Fields("Modo").Value
            'Rslocal.Close()

            Modo = lobjExportar.ObtenerModoEjecucion()
            '/* MUL FIN
        End If


        ''''------borrar ini!!!

        Modo = UCase("Acumulado")

        '''''------borrar fin!!!


        ' 25-6-2007. Eloi. En los procesos batch se guarda el inicio del proceso
        lResultado = lobjExportar.InsertarLog("------------------------------------------------------------------------------", False)
        lResultado = lobjExportar.InsertarLog("El proceso de Exportación de Asistencia " & UCase(Modo) & " ha empezado: " & Now.ToShortDateString & " " & Now.ToShortTimeString, False)
        lResultado = lobjExportar.InsertarLog(" ", False)

        ' PARA PRUEBAS -----------------------------------------------------------------------
        ' Modo = "ACUMULADO"
        'TipoEjecucion = "P"
        ' -------------------------------------------------------------------------------------

        ' /*MUL INI
        ' Se ha de ejecutar para las asistencias de InterParner, Europ Assitance  y Multiasistencia
        For idx_asistencia = 1 To lbxCompania.Items.Count
            If Trim(lbxCompania.Items(idx_asistencia - 1)) = "I" Or _
               Trim(lbxCompania.Items(idx_asistencia - 1)) = "E" Or _
               Trim(lbxCompania.Items(idx_asistencia - 1)) = "M" Then

                'Seleccionar la compañia de asistencia
                Codcia = Trim(lbxCompania.Items(idx_asistencia - 1))

                If DatosCiaAsistencia(Codcia) Then
                    Call ObtenerLote()
                Else
                    lResultado = lobjExportar.InsertarLog("No hay datos para la compañía de Asistencia " & Trim(UCase(NombreCompa)) & " " & Now.ToShortTimeString, False)
                    Exit Sub
                End If

                lResultado = lobjExportar.InsertarLog("Exportación Asistencia empresa " & Trim(UCase(NombreCompa)) & " ha empezado a las " & Now.ToShortTimeString, False)
                '/* MUL FIN

                Path = claseBDExportacion.PathExport
                If UCase(Modo) = UCase("ACUMULADO") Then
                    ProcesoExportar(("Acumulado"))
                ElseIf UCase(Modo) = UCase("DIARIO") Then
                    ProcesoExportar(("Diario"))
                End If

                PassWordZip = Format(Today, "yyyyMMdd") & LoteEnviado

                ' Una vez ejecutado el proceso generamos un fichero zip para su envio
                '
                Select Case Codcia
                    Case "I" ' Interpartner 
                        If Len(FicheroFusion) > 0 Then
                            NombreFicheroZip1 = Trim(Descia) & "_" & Modo & "_" & "Global_" & LoteEnviado & ".zip"
                            '/* MUL INI
                            'Comando = "C:\Archivos de programa\WinZip\wzzip -a " & Path & "\" & NombreFicheroZip1 & " " & FicheroFusion ' + " -s" + PassWordZip
                            Comando = claseBDExportacion.getZipCommand() + " -a " & FtpCiaExport & "\" & NombreFicheroZip1 & " " & FicheroFusion ' + " -s" + PassWordZip
                            '//Shell(Comando)
                            Shell(Comando, AppWinStyle.Hide, True, 18000)
                            '/* MUL FIN
                            'v_zip = LTrim(Path + "\" + NombreFicheroZip1)
                            '/* MUL INI
                            'CopiarFichero(Path & "\" & NombreFicheroZip1, FtpCiaExport & "\" & NombreFicheroZip1)
                            lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & NombreFicheroZip1, False)

                            'objUtilidades.SplitPath(FicheroFusion, spath, nomfich, sext)
                            'CopiarFichero(FicheroFusion, FtpCiaExport & "\" & nomfich)
                            '/* MUL FIN
                        End If

                        If HayPeritajes Then
                            If Len(Fich4) > 0 Then
                                NombreFicheroZip = Trim(Descia) & "_" & "Diario" & "_" & "Peritajes_" & LoteEnviado & ".zip"
                                '/* MUL INI
                                'Comando = "C:\Archivos de programa\WinZip\wzzip -a " & Path & "\" & NombreFicheroZip & " " + Fich4 ' + " -s" + PassWordZip
                                Comando = claseBDExportacion.getZipCommand() + " -a " & FtpCiaExport & "\" & NombreFicheroZip & " " + Fich4 ' + " -s" + PassWordZip
                                '/* MUL FIN
                                '//Shell(Comando)
                                Shell(Comando, AppWinStyle.Hide, True, 18000)
                                'v_zip = ""
                                'v_zip = LTrim(Path + "\" + NombreFicheroZip1) + LTrim(",") + LTrim(Path + "\" + NombreFicheroZip)
                                '/* MUL INI
                                'CopiarFichero(Path & "\" & NombreFicheroZip, FtpCiaExport & "\" & NombreFicheroZip)
                                lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & NombreFicheroZip, False)
                                'objUtilidades.SplitPath(Fich4, spath, nomfich, sext)
                                'CopiarFichero(Fich4, FtpCiaExport & "\" & nomfich)
                                '/* MUL FIN
                            End If
                        End If

                        If HayCruce Then
                            If Len(Fich5) > 0 Then
                                NombreFicheroZip3 = Trim(Descia) & "_" & "Diario" & "_" & "CruceReferencias_" & LoteEnviado & ".zip"
                                '/* MUL INI
                                'Comando = "C:\Archivos de programa\WinZip\wzzip -a " & Path & "\" & NombreFicheroZip3 & " " & Fich5 ' + " -s" + PassWordZip
                                Comando = claseBDExportacion.getZipCommand() + " -a " & FtpCiaExport & "\" & NombreFicheroZip3 & " " & Fich5 ' + " -s" + PassWordZip
                                'Shell(Comando)
                                Shell(Comando, AppWinStyle.Hide, True, 18000)
                                '/* MUL FIN
                                'v_zip = ""
                                'v_zip = LTrim(Path + "\" + NombreFicheroZip1) + LTrim(",") + LTrim(Path + "\" + NombreFicheroZip) + LTrim(",") + LTrim(Path + "\" + NombreFicheroZip3)
                                '/* MUL INI
                                'CopiarFichero(Path & "\" & NombreFicheroZip3, FtpCiaExport & "\" & NombreFicheroZip3)
                                lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & NombreFicheroZip3, False)
                                'objUtilidades.SplitPath(Fich5, spath, nomfich, sext)
                                'CopiarFichero(Fich5, FtpCiaExport & "\" & nomfich)
                                '/* MUL FIN
                            End If
                        End If
                        '/* MUL INI solo se copian los ficheros zip 
                        'copiar los ficheros a exportar
                        '''objUtilidades.SplitPath(Fich1, spath, nomfich, sext)
                        '''CopiarFichero(Fich1, FtpCiaExport & "\" & nomfich)

                        '''objUtilidades.SplitPath(Fich2, spath, nomfich, sext)
                        '''CopiarFichero(Fich2, FtpCiaExport & "\" & nomfich)

                        '''objUtilidades.SplitPath(Fich3, spath, nomfich, sext)
                        '''CopiarFichero(Fich3, FtpCiaExport & "\" & nomfich)
                        '/* MUL FIN

                    Case "E" 'Europ Assistance
                        '/* MUL INI copiar los ficheros a exportar
                        NombreFicheroZip = "\MDP_polizas_" & LoteEnviado & "_" & Format(Today, "yyyyMMdd") & ".txt"
                        CopiarFichero(Path & NombreFicheroZip, FtpCiaExport & NombreFicheroZip)
                        lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & NombreFicheroZip, False)

                        NombreFicheroZip1 = "\MDP_garantias_" & LoteEnviado & "_" & Format(Today, "yyyyMMdd") & ".txt"
                        CopiarFichero(Path & NombreFicheroZip1, FtpCiaExport & NombreFicheroZip1)
                        lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & NombreFicheroZip1, False)

                        If HayPeritajes Then
                            objUtilidades.SplitPath(Fich4, spath, nomfich, sext)
                            CopiarFichero(Fich4, FtpCiaExport & "\" & nomfich)
                            lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & nomfich, False)
                        End If

                        If HayCruce Then
                            objUtilidades.SplitPath(Fich5, spath, nomfich, sext)
                            CopiarFichero(Fich5, FtpCiaExport & "\" & nomfich)
                            lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & nomfich, False)
                        End If
                        '/* MUL FIN

                    Case "M" 'Multiasistencia
                        '/* MUL INI copiar los ficheros a exportar
                        If Len(FicheroFusion) > 0 Then
                            objUtilidades.SplitPath(FicheroFusion, spath, nomfich, sext)
                            CopiarFichero(FicheroFusion, FtpCiaExport & "\" & nomfich)
                            lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & nomfich, False)
                        End If
                        If HayPeritajes Then
                            objUtilidades.SplitPath(Fich4, spath, nomfich, sext)
                            CopiarFichero(Fich4, FtpCiaExport & "\" & nomfich)
                            lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & nomfich, False)
                        End If
                        If HayCruce Then
                            objUtilidades.SplitPath(Fich5, spath, nomfich, sext)
                            CopiarFichero(Fich5, FtpCiaExport & "\" & nomfich)
                            lResultado = lobjExportar.InsertarLog("*- Se ha generado -> '" & FtpCiaExport & "\" & nomfich, False)
                        End If
                        '/* MUL FIN

                    Case "R", "A" ' Reparalia y Angel envio por email

                        If Len(Fich1) > 0 Then
                            NombreFicheroZip1 = Trim(Descia) & "_" & Modo & "_" & "DatosCabecera_" & LoteEnviado & ".zip"
                            Comando = "C:\Archivos de programa\WinZip\wzzip -a " & Path & "\" & NombreFicheroZip1 & " " + Fich1 + " "
                            Shell(Comando)
                        End If

                        If Len(Fich2) > 0 Then
                            NombreFicheroZip = Trim(Descia) & "_" & Modo & "_" & "DatosGarantias_" & LoteEnviado & ".zip"
                            Comando = "C:\Archivos de programa\WinZip\wzzip -a " & Path & "\" & NombreFicheroZip & " " + Fich2 + " "
                            Shell(Comando)
                        End If
                        v_zip = LTrim(Path & "\" & NombreFicheroZip1) & LTrim(",") & LTrim(Path & "\" & NombreFicheroZip) & LTrim(",") & LTrim(Path & "\" & NombreFicheroZip3)

                    Case Else
                End Select

                ' Copiamos al disco FTP el fichero zip con los datos de pólizas, riesgos, garantías, peritajes, aperturas

                '/*MUL INI  ya no se envia por FTP, ahora lo recogen las empresas de Asistencia

                'If ComandosFTP(Path, NombreFicheroZip1, NombreFicheroZip, NombreFicheroZip3) Then
                '    System.Windows.Forms.Application.DoEvents()
                '    Comando = "C:\psftp mprop@195.77.230.7 -pw ultpg.2 -b C:\comandos.scr"
                '    Shell(Comando)
                'Else
                '    Err.Raise(1)
                'End If
                '/* MUL FIN

                If Not lobjExportar.RegistroComunicaciones Then Err.Raise(50)

                '/*MUL INI
                lResultado = lobjExportar.InsertarLog("Exportación Asistencia empresa " & Trim(UCase(NombreCompa)) & " ha acabado a las " & Now.ToShortTimeString, False)
                lResultado = lobjExportar.InsertarLog(" ", False)
            End If
        Next
        '/*MUL FIN

        lResultado = lobjExportar.InsertarLog(" ", False)
        lResultado = lobjExportar.InsertarLog("El proceso de Exportación Asistencia " & UCase(Modo) & " ha finalizado correctamente", False)
        lResultado = lobjExportar.InsertarLog("------------------------------------------------------------------------------", False)
        Exit Sub

        ' -------------------------------------------------------------------------------------
        ' 21/09/2009 JLL - Fin
        ' -------------------------------------------------------------------------------------

ProcesoProgramado_Err:

        ' Quitar el mensaje. Escribir fichero de texto con error
        ' 25-6-2007. Eloi. En los procesos batch se guarda el error en un fichero log
        '
        If TipoEjecucion <> "P" Then
            If Err.Number = 50 Then
                MsgBox("El proceso de registro de los datos enviados a la compañía de asistencia ha dado un error. El registro deberá realizarse manualmente. Llame a informática", MsgBoxStyle.Critical, "Notificación de Errores:")
            Else
                MsgBox("El proceso de envio automatizado ha devuelto un error. Los ficheros a la compañía de Asistencia no han sido enviados", MsgBoxStyle.Critical, "Notificación de Errores:")
            End If
        Else
            If Err.Number = 50 Then
                lResultado = lobjExportar.InsertarLog("El proceso de registro de datos enviados a la compañía de asistencia ha devuelto un error. El registro habrá que realizarlo manualmente. Llame a Informática", True)
            Else
                lResultado = lobjExportar.InsertarLog("El proceso de envio via FTP ha devuelto un error. Los ficheros a la compañía de Asistencia no han sido enviados", True)
            End If
        End If
    End Sub

    Public Sub ProcesoExportar(ByRef TipoExp As String)

        On Error GoTo ProcesoExportar_Err

        ' Declaraciones
        '
        Dim objExportar As clsExportar_NET
        Dim lngTiempo As Integer
        Dim strParametro As String
        Dim lResultado As Boolean
        Dim resultadoMsg As MsgBoxResult
        '

        ' Asignación de valores iniciales a objetos y variables
        dteFechaIni = Now
        'Label5.Text = " Procesando exportación de datos en modo " & TipoExp
        stbEstado.Panels(1).Text = " Procesando exportación de datos en modo " & TipoExp
        strParametro = Microsoft.VisualBasic.Command

        ' Llamada al proceso de validación de los parámetros
        ' establecidos para la exportación
        '
        If ValidarDatos() <> True Then
            Exit Sub
        End If

        ' Pedimos la confirmación de la orden de exportar
        '
        If TipoEjecucion <> "P" Then
            strErr = ""
            resultadoMsg = MsgBox("Se va a iniciar el proceso de exportación de datos en modo " & TipoExp & " ¿Desea confirmarlo?", MsgBoxStyle.YesNo)
            If resultadoMsg = MsgBoxResult.Yes Then
                System.Windows.Forms.Application.DoEvents()
                Me.Cursor = Cursors.WaitCursor
            Else
                Exit Sub
            End If
        End If

        ' Creamos el objeto de exportación a la compañia y
        ' asignamos los valores a sus propiedades
        '
        objExportar = New clsExportar_NET
        objExportar.FecEfe = dtpFechaEfecto.Value
        If strParametro = "" Or Strings.Left(strParametro, 1) <> "P" Then
            objExportar.Archivo = lblDirDestino.Text
        ElseIf Strings.Left(strParametro, 1) = "P" Then
            objExportar.Archivo = claseBDExportacion.PathExport
        End If
        objExportar.TipoExportacion = TipoExp

        If Not objExportar.SeleccionPolizas Then
            Err.Raise(1)
        End If

        ' Llamada al objeto de Exportación según la compañía
        '
        Select Case Codcia

            Case "A", "R" ' Angel, Reparalia
                ' Llamada al proceso que confecciona los registros de exportación para el fichero de cabecera
                If objExportar.DatosCabecera Then
                    ' Llamada al proceso que confecciona los registros de exportación para el fichero de Garantías
                    If objExportar.DatosGarantias Then
                        ' llamada al proceso que actualiza los datos en el histórico de movimientos de pólizas
                        System.Windows.Forms.Application.DoEvents()
                        If objExportar.ActualizaHistorico Then
                            If objExportar.ActualizaLote(Codcia) Then
                                strErr = "4033"
                            Else
                                Err.Raise(1)
                            End If
                        Else
                            Err.Raise(1)
                        End If
                    Else
                        Err.Raise(1)
                    End If
                Else
                    Err.Raise(1)
                End If

                ''/*MUL INI */
            Case "I"
                ''/*MUL FIN*/

                ''/*MUL INI */
                ' quito los else ....
                'If objExportar.DatosCabecera_IP Then
                '    If objExportar.DatosRiesgo_IP Then
                '        If objExportar.DatosGarantias_IP Then
                '            If objExportar.FusionFicheros_IP Then
                '                If objExportar.DatosPeritajes_IP Then
                '                    If objExportar.DatosCruceReferencias_IP Then
                '                        If objExportar.ActualizaHistorico Then
                '                            If objExportar.ActualizaLote Then
                '                                strErr = "El Proceso de Exportacion a la compañía de asistencia ha terminado con exito."
                '                            Else
                '                                strErr = "No se ha podido generar el número de lote. El envio de ficheros a la compañía de asistencia no se ha  realizado. Avise a informática"
                '                                Err.Raise(1)
                '                            End If
                '                        Else
                '                            strErr = "Se ha producido un error critico en la actualización del historico de movimientos y estados de pólizas"
                '                            Err.Raise(1)
                '                        End If
                '                    Else
                '                        strErr = "Se ha producido un error critico al generar el fichero de cruce de referencias.Avise a informática"
                '                        Err.Raise(1)
                '                    End If
                '                Else
                '                    strErr = "Se ha producido un error crítico al generar el fichero de exportación de datos de peritajes. Avise a Informática"
                '                    Err.Raise(1)
                '                End If
                '            Else
                '                strErr = "Se ha producido un error en el procedimeinto de fusión de los diferentes ficheros. El proceso de envio no se ha producido."
                '                Err.Raise(1)
                '            End If
                '        Else
                '            strErr = "Se ha producido un error critico al generar el fichero de garantías en la exportación de datos. Avise a Informática."
                '            Err.Raise(1)
                '        End If
                '    Else
                '        strErr = "Se ha producido un critico al generar el fichero de riesgos en la exportacion de datos. Avise a Informatica"
                '        Err.Raise(1)
                '    End If
                'Else
                '    strErr = "Se ha producido un error critico al generar el fichero de datos de cabecera en la exportación de datos. Avise a Informática"
                '    Err.Raise(1)
                'End If
                If objExportar.DatosCabecera_IP = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de datos de cabecera en la exportación de datos. Avise a Informática"
                    Err.Raise(1)
                End If
                If objExportar.DatosRiesgo_IP = False Then
                    strErr = "Se ha producido un critico al generar el fichero de riesgos en la exportacion de datos. Avise a Informatica"
                    Err.Raise(1)
                End If
                If objExportar.DatosGarantias_IP = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de garantías en la exportación de datos. Avise a Informática."
                    Err.Raise(1)
                End If
                ''/*MUL INI */
                ''If objExportar.FusionFicheros_IP = False Then
                If objExportar.FusionFicherosCia(Codcia) = False Then
                    ''/*MUL FIN */
                    strErr = "Se ha producido un error en el procedimiento de fusión de los diferentes ficheros. El proceso de envio no se ha producido."
                    Err.Raise(1)
                End If
                If objExportar.DatosPeritajes_IP = False Then
                    strErr = "Se ha producido un error crítico al generar el fichero de exportación de datos de peritajes. Avise a Informática"
                    Err.Raise(1)
                End If
                If objExportar.DatosCruceReferencias_IP = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de cruce de referencias.Avise a informática"
                    Err.Raise(1)
                End If
                If objExportar.ActualizaHistorico = False Then
                    strErr = "Se ha producido un error critico en la actualización del historico de movimientos y estados de pólizas"
                    Err.Raise(1)
                End If
                If objExportar.ActualizaLote(Codcia) = False Then
                    strErr = "No se ha podido generar el número de lote. El envio de ficheros a la compañía de asistencia no se ha  realizado. Avise a informática"
                    Err.Raise(1)
                End If
                strErr = "El Proceso de Exportacion a la compañía de asistencia ha terminado con exito."

            Case "E"  ' Europe Assistance 
                If objExportar.DatosPolizasCia(Codcia) = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de datos de cabecera en la exportación de datos. Avise a Informática"
                    Err.Raise(1)
                End If
                'If objExportar.DatosRiesgo_IP = False Then
                '    strErr = "Se ha producido un critico al generar el fichero de riesgos en la exportacion de datos. Avise a Informatica"
                '    Err.Raise(1)
                'End If
                If objExportar.DatosGarantiasCia(Codcia) = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de garantías en la exportación de datos. Avise a Informática."
                    Err.Raise(1)
                End If
                If objExportar.FusionFicherosCia(Codcia) = False Then
                    strErr = "Se ha producido un error en el procedimiento de fusión de los diferentes ficheros. El proceso de envio no se ha producido."
                    Err.Raise(1)
                End If
                'fusionar la cabecera de garantias
                If objExportar.FusionFicherosCia("EG") = False Then
                    strErr = "Se ha producido un error en el procedimiento de fusión de los ficheros de garantias. El proceso de envio no se ha producido."
                    Err.Raise(1)
                End If
                If objExportar.DatosPeritajesCia(Codcia) = False Then
                    strErr = "Se ha producido un error crítico al generar el fichero de exportación de datos de peritajes. Avise a Informática"
                    Err.Raise(1)
                End If
                If objExportar.DatosCruceReferenciasCia(Codcia) = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de cruce de referencias.Avise a informática"
                    Err.Raise(1)
                End If
                If objExportar.ActualizaHistorico = False Then
                    strErr = "Se ha producido un error critico en la actualización del historico de movimientos y estados de pólizas"
                    Err.Raise(1)
                End If
                If objExportar.ActualizaLote(Codcia) = False Then
                    strErr = "No se ha podido generar el número de lote. El envio de ficheros a la compañía de asistencia no se ha  realizado. Avise a informática"
                    Err.Raise(1)
                End If
                strErr = "El Proceso de Exportacion a la compañía de asistencia ha terminado con exito."

            Case "M"   'Multiasistencia
                If objExportar.DatosPolizasCia(Codcia) = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de datos de cabecera en la exportación de datos. Avise a Informática"
                    Err.Raise(1)
                End If
                'If objExportar.DatosRiesgo_IP = False Then
                '    strErr = "Se ha producido un critico al generar el fichero de riesgos en la exportacion de datos. Avise a Informatica"
                '    Err.Raise(1)
                'End If
                'If objExportar.DatosGarantiasCia(Codcia) = False Then
                '    strErr = "Se ha producido un error critico al generar el fichero de garantías en la exportación de datos. Avise a Informática."
                '    Err.Raise(1)
                'End If
                If objExportar.FusionFicherosCia(Codcia) = False Then
                    strErr = "Se ha producido un error en el procedimiento de fusión de los diferentes ficheros. El proceso de envio no se ha producido."
                    Err.Raise(1)
                End If
                If objExportar.DatosPeritajesCia(Codcia) = False Then
                    strErr = "Se ha producido un error crítico al generar el fichero de exportación de datos de peritajes. Avise a Informática"
                    Err.Raise(1)
                End If
                If objExportar.DatosCruceReferenciasCia(Codcia) = False Then
                    strErr = "Se ha producido un error critico al generar el fichero de cruce de referencias.Avise a informática"
                    Err.Raise(1)
                End If
                If objExportar.ActualizaHistorico = False Then
                    strErr = "Se ha producido un error critico en la actualización del historico de movimientos y estados de pólizas"
                    Err.Raise(1)
                End If
                If objExportar.ActualizaLote(Codcia) = False Then
                    strErr = "No se ha podido generar el número de lote. El envio de ficheros a la compañía de asistencia no se ha  realizado. Avise a informática"
                    Err.Raise(1)
                End If
                strErr = "El Proceso de Exportacion a la compañía de asistencia ha terminado con exito."

                ''/*MUL FIN */
        End Select

        ' Ocultar y dejar barra estado como al principio
        '
        Me.prbProgreso.Visible = False
        Me.stbEstado.Panels(1).Text = ""
        Me.Cursor = Cursors.Default

        If TipoEjecucion <> "P" Then
            MsgBox(strErr, MsgBoxStyle.Exclamation)
        End If

        Exit Sub

ProcesoExportar_Err:
        ' 25-6-2007. Eloi. En los procesos batch se guarda el error en un fichero log
        Dim lstrErr As String

        Me.Cursor = Cursors.Default

        If TipoEjecucion = "P" Then
            lstrErr = strErr
            lResultado = objExportar.InsertarLog(lstrErr, True)
        Else
            MsgBox(strErr)
        End If
    End Sub

    ' Function:   ValidarDatos
    ' Objetivo:   Valida que los datos seleccionados en el formulario sean
    '             correctos. Talos como compañia, fecha y fichero.
    ' Parametros:
    ' Retorno:    Verdadero o Falso.
    '
    Public Function ValidarDatos() As Boolean

        On Error GoTo ValidarDatos_Err

        Dim lResultado As Boolean
        Dim lstrErr As String
        Dim lobjExportar As clsExportar_NET

        ValidarDatos = True

        ' Validación de la compañia de asistencia ( si no se ha introducido )
        If cbxCompania.SelectedIndex = -1 Then
            Err.Raise(2500)
        End If

        ' Validación de la fecha de efecto ( si no se ha introducido )
        If Not IsDate(dtpFechaEfecto.Value) Then
            Err.Raise(2501)
        End If

        ' Validación de la fecha de efecto
        '
        If Not IsDate(dtpFechaEfecto.Value) Then
            Err.Raise(2502)
        End If

        ' Validación del nombre del fichero
        '
        If objUtilidades.ValidaNombreFichero((frmInstanciaPrincipal.lblDirDestino.Text)) <> True Then
            Err.Raise(2503)
        End If

        ' Validación número de lote
        '
        If LoteEnviado = "Error" Then
            Err.Raise(2504)
        End If

        Exit Function

ValidarDatos_Err:
        ' 25-6-2007. Eloi. En los procesos batch se guarda el error en un fichero log
        ' Quitar pantalla cuando error. Escribir fichero de texto con error
        Select Case Err.Number

            Case 2500 ' No se ha introducido el código de compañia de asistencia
                strErr = "No se ha seleccionado ninguna compañía de asistencia"

            Case 2501 ' No se ha introducido la fecha de efecto
                strErr = "No se ha introducido la fecha de efecto para la exportación de datos"

            Case 2502 ' La fecha de efecto intrducida no es válida
                strErr = "La fecha de efecto introducida no es correcta"

            Case 2503 ' El Nombre de fichero no es válido
                strErr = "El nombre de fichero no es válido. Puede haber carácteres no permitidos."

            Case 2504 ' No se ha generado un número de lote correcto
                strErr = "No se ha podido generar el número de lote. El envio de ficheros a la compañía de asistencia no se ha  realizado. Avise a informática"

            Case Else ' Otros
                strErr = "Error no catalogado en el proceso de validación."

        End Select
        ValidarDatos = False
        If TipoEjecucion = "P" Then
            lobjExportar = New clsExportar_NET
            lstrErr = strErr
            lResultado = lobjExportar.InsertarLog(lstrErr, True)
        Else
            MsgBox(strErr, MsgBoxStyle.Exclamation)
        End If
    End Function

    Private Sub cbSelDirectorio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSelDirectorio.Click
        Dim resultadoDirectorio As DialogResult
        resultadoDirectorio = fbdDirectorio.ShowDialog()

        If resultadoDirectorio = DialogResult.OK Then
            lblDirDestino.Text = fbdDirectorio.SelectedPath
        End If

    End Sub

    Private Sub cbDiario_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDiario.Click
        ProcesoExportar("Diario")
    End Sub

    Private Sub cbAcumulado_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAcumulado.Click
        ProcesoExportar("Acumulado")
    End Sub

    Private Sub cbxCompania_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxCompania.SelectedIndexChanged
        Codcia = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))

        If DatosCiaAsistencia(Codcia) Then
            Call ObtenerLote()
        Else
            MsgBox("No hay datos de compañias de asistencia", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub cbSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSalir.Click
        Application.Exit()
    End Sub


    Private Function CopiarFichero(ByVal fichorigen As String, ByVal fichdestino As String) As Boolean

        On Error GoTo CopiarFichero_Err

        FileSystem.FileCopy(fichorigen, fichdestino)
        CopiarFichero = True
        Exit Function

CopiarFichero_Err:
        Dim lobjExportar As clsExportar_NET

        lobjExportar = New clsExportar_NET
        lobjExportar.InsertarLog("Exportación Asistencia no se ha podido copiar " & fichorigen & " a " & fichdestino & "  Error: " & Err.Description & " " & Now.ToShortTimeString, False)
        CopiarFichero = False
    End Function


End Class
