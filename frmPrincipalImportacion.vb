Public Class frmPrincipalImportacion
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
    Friend WithEvents lbCompaniaAsistencia As System.Windows.Forms.Label
    Friend WithEvents cbxCompania As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lbxFicheros As System.Windows.Forms.ListBox
    Friend WithEvents cbBuscar As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents cbImportar As System.Windows.Forms.Button
    Friend WithEvents pbLogoMutua As System.Windows.Forms.PictureBox
    Friend WithEvents cbBajarAperturas As System.Windows.Forms.Button
    Friend WithEvents stbEstado As System.Windows.Forms.StatusBar
    Friend WithEvents sbPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbPanel4 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents lbxCompania As System.Windows.Forms.ListBox
    Friend WithEvents prbProgreso As System.Windows.Forms.ProgressBar
    Friend WithEvents cbBajarPagos As System.Windows.Forms.Button
    Friend WithEvents dlgFicheros As System.Windows.Forms.OpenFileDialog
    Public WithEvents File1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Public WithEvents Dir1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents cbxProceso As System.Windows.Forms.ComboBox
    Friend WithEvents gbFTP As System.Windows.Forms.GroupBox
    Friend WithEvents pbFTP As System.Windows.Forms.PictureBox
    Friend WithEvents cbAvisoErrores As System.Windows.Forms.Button
    Friend WithEvents picTest As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrincipalImportacion))
        Me.pbLogoMutua = New System.Windows.Forms.PictureBox
        Me.lbCompaniaAsistencia = New System.Windows.Forms.Label
        Me.cbxCompania = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbxProceso = New System.Windows.Forms.ComboBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cbBuscar = New System.Windows.Forms.Button
        Me.lbxFicheros = New System.Windows.Forms.ListBox
        Me.gbFTP = New System.Windows.Forms.GroupBox
        Me.cbBajarPagos = New System.Windows.Forms.Button
        Me.cbBajarAperturas = New System.Windows.Forms.Button
        Me.pbFTP = New System.Windows.Forms.PictureBox
        Me.cbImportar = New System.Windows.Forms.Button
        Me.cbAvisoErrores = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.stbEstado = New System.Windows.Forms.StatusBar
        Me.sbPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.sbPanel4 = New System.Windows.Forms.StatusBarPanel
        Me.lbxCompania = New System.Windows.Forms.ListBox
        Me.prbProgreso = New System.Windows.Forms.ProgressBar
        Me.dlgFicheros = New System.Windows.Forms.OpenFileDialog
        Me.File1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
        Me.Dir1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
        Me.picTest = New System.Windows.Forms.PictureBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.gbFTP.SuspendLayout()
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbPanel4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbLogoMutua
        '
        Me.pbLogoMutua.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pbLogoMutua.Image = CType(resources.GetObject("pbLogoMutua.Image"), System.Drawing.Image)
        Me.pbLogoMutua.Location = New System.Drawing.Point(8, 8)
        Me.pbLogoMutua.Name = "pbLogoMutua"
        Me.pbLogoMutua.Size = New System.Drawing.Size(48, 48)
        Me.pbLogoMutua.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbLogoMutua.TabIndex = 0
        Me.pbLogoMutua.TabStop = False
        '
        'lbCompaniaAsistencia
        '
        Me.lbCompaniaAsistencia.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbCompaniaAsistencia.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lbCompaniaAsistencia.Location = New System.Drawing.Point(72, 16)
        Me.lbCompaniaAsistencia.Name = "lbCompaniaAsistencia"
        Me.lbCompaniaAsistencia.Size = New System.Drawing.Size(152, 24)
        Me.lbCompaniaAsistencia.TabIndex = 2
        Me.lbCompaniaAsistencia.Text = "Compañía Asistencia:"
        Me.lbCompaniaAsistencia.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxCompania
        '
        Me.cbxCompania.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxCompania.Location = New System.Drawing.Point(232, 16)
        Me.cbxCompania.Name = "cbxCompania"
        Me.cbxCompania.Size = New System.Drawing.Size(344, 24)
        Me.cbxCompania.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Proceso:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbxProceso)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox1.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(568, 56)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Selección de procesos"
        '
        'cbxProceso
        '
        Me.cbxProceso.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbxProceso.Location = New System.Drawing.Point(80, 24)
        Me.cbxProceso.Name = "cbxProceso"
        Me.cbxProceso.Size = New System.Drawing.Size(216, 21)
        Me.cbxProceso.TabIndex = 5
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbBuscar)
        Me.GroupBox2.Controls.Add(Me.lbxFicheros)
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox2.Location = New System.Drawing.Point(8, 128)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(568, 144)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Selección de ficheros"
        '
        'cbBuscar
        '
        Me.cbBuscar.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cbBuscar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBuscar.Image = CType(resources.GetObject("cbBuscar.Image"), System.Drawing.Image)
        Me.cbBuscar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbBuscar.Location = New System.Drawing.Point(496, 72)
        Me.cbBuscar.Name = "cbBuscar"
        Me.cbBuscar.Size = New System.Drawing.Size(64, 56)
        Me.cbBuscar.TabIndex = 1
        Me.cbBuscar.Text = "Buscar"
        Me.cbBuscar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lbxFicheros
        '
        Me.lbxFicheros.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.lbxFicheros.Location = New System.Drawing.Point(8, 24)
        Me.lbxFicheros.Name = "lbxFicheros"
        Me.lbxFicheros.Size = New System.Drawing.Size(472, 108)
        Me.lbxFicheros.TabIndex = 0
        '
        'gbFTP
        '
        Me.gbFTP.Controls.Add(Me.cbBajarPagos)
        Me.gbFTP.Controls.Add(Me.cbBajarAperturas)
        Me.gbFTP.Controls.Add(Me.pbFTP)
        Me.gbFTP.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFTP.ForeColor = System.Drawing.Color.RoyalBlue
        Me.gbFTP.Location = New System.Drawing.Point(8, 280)
        Me.gbFTP.Name = "gbFTP"
        Me.gbFTP.Size = New System.Drawing.Size(256, 96)
        Me.gbFTP.TabIndex = 7
        Me.gbFTP.TabStop = False
        Me.gbFTP.Text = "FTP"
        Me.gbFTP.Visible = False
        '
        'cbBajarPagos
        '
        Me.cbBajarPagos.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cbBajarPagos.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbBajarPagos.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBajarPagos.Image = CType(resources.GetObject("cbBajarPagos.Image"), System.Drawing.Image)
        Me.cbBajarPagos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbBajarPagos.Location = New System.Drawing.Point(64, 56)
        Me.cbBajarPagos.Name = "cbBajarPagos"
        Me.cbBajarPagos.Size = New System.Drawing.Size(176, 32)
        Me.cbBajarPagos.TabIndex = 2
        Me.cbBajarPagos.Text = "Bajar archivos de Pagos"
        Me.cbBajarPagos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cbBajarAperturas
        '
        Me.cbBajarAperturas.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cbBajarAperturas.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbBajarAperturas.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbBajarAperturas.Image = CType(resources.GetObject("cbBajarAperturas.Image"), System.Drawing.Image)
        Me.cbBajarAperturas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbBajarAperturas.Location = New System.Drawing.Point(64, 16)
        Me.cbBajarAperturas.Name = "cbBajarAperturas"
        Me.cbBajarAperturas.Size = New System.Drawing.Size(176, 32)
        Me.cbBajarAperturas.TabIndex = 1
        Me.cbBajarAperturas.Text = "Bajar archivos de Aperturas"
        Me.cbBajarAperturas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pbFTP
        '
        Me.pbFTP.Image = CType(resources.GetObject("pbFTP.Image"), System.Drawing.Image)
        Me.pbFTP.Location = New System.Drawing.Point(16, 24)
        Me.pbFTP.Name = "pbFTP"
        Me.pbFTP.Size = New System.Drawing.Size(32, 32)
        Me.pbFTP.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbFTP.TabIndex = 0
        Me.pbFTP.TabStop = False
        '
        'cbImportar
        '
        Me.cbImportar.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbImportar.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbImportar.Image = CType(resources.GetObject("cbImportar.Image"), System.Drawing.Image)
        Me.cbImportar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbImportar.Location = New System.Drawing.Point(392, 320)
        Me.cbImportar.Name = "cbImportar"
        Me.cbImportar.Size = New System.Drawing.Size(64, 56)
        Me.cbImportar.TabIndex = 8
        Me.cbImportar.Text = "Ejecutar"
        Me.cbImportar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbAvisoErrores
        '
        Me.cbAvisoErrores.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAvisoErrores.ForeColor = System.Drawing.Color.DodgerBlue
        Me.cbAvisoErrores.Image = CType(resources.GetObject("cbAvisoErrores.Image"), System.Drawing.Image)
        Me.cbAvisoErrores.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cbAvisoErrores.Location = New System.Drawing.Point(456, 320)
        Me.cbAvisoErrores.Name = "cbAvisoErrores"
        Me.cbAvisoErrores.Size = New System.Drawing.Size(64, 56)
        Me.cbAvisoErrores.TabIndex = 9
        Me.cbAvisoErrores.Text = "Errores"
        Me.cbAvisoErrores.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.DodgerBlue
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button5.Location = New System.Drawing.Point(520, 320)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(56, 56)
        Me.Button5.TabIndex = 10
        Me.Button5.Text = "Salir"
        Me.Button5.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'stbEstado
        '
        Me.stbEstado.Location = New System.Drawing.Point(0, 386)
        Me.stbEstado.Name = "stbEstado"
        Me.stbEstado.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbPanel1, Me.sbPanel2, Me.sbPanel3, Me.sbPanel4})
        Me.stbEstado.ShowPanels = True
        Me.stbEstado.Size = New System.Drawing.Size(584, 24)
        Me.stbEstado.TabIndex = 15
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
        Me.sbPanel3.Width = 107
        '
        'sbPanel4
        '
        Me.sbPanel4.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.sbPanel4.Width = 10
        '
        'lbxCompania
        '
        Me.lbxCompania.Location = New System.Drawing.Point(632, 48)
        Me.lbxCompania.Name = "lbxCompania"
        Me.lbxCompania.Size = New System.Drawing.Size(24, 17)
        Me.lbxCompania.TabIndex = 16
        '
        'prbProgreso
        '
        Me.prbProgreso.Location = New System.Drawing.Point(458, 396)
        Me.prbProgreso.Name = "prbProgreso"
        Me.prbProgreso.Size = New System.Drawing.Size(96, 8)
        Me.prbProgreso.TabIndex = 17
        Me.prbProgreso.Visible = False
        '
        'File1
        '
        Me.File1.BackColor = System.Drawing.SystemColors.Window
        Me.File1.Cursor = System.Windows.Forms.Cursors.Default
        Me.File1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.File1.Location = New System.Drawing.Point(312, 328)
        Me.File1.Name = "File1"
        Me.File1.Pattern = "*.zip"
        Me.File1.Size = New System.Drawing.Size(48, 17)
        Me.File1.TabIndex = 24
        Me.File1.Visible = False
        '
        'Dir1
        '
        Me.Dir1.BackColor = System.Drawing.SystemColors.Window
        Me.Dir1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Dir1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Dir1.IntegralHeight = False
        Me.Dir1.Location = New System.Drawing.Point(296, 296)
        Me.Dir1.Name = "Dir1"
        Me.Dir1.Size = New System.Drawing.Size(64, 21)
        Me.Dir1.TabIndex = 23
        Me.Dir1.Visible = False
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
        'frmPrincipalImportacion
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(584, 410)
        Me.Controls.Add(Me.File1)
        Me.Controls.Add(Me.Dir1)
        Me.Controls.Add(Me.prbProgreso)
        Me.Controls.Add(Me.lbxCompania)
        Me.Controls.Add(Me.stbEstado)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.cbAvisoErrores)
        Me.Controls.Add(Me.cbImportar)
        Me.Controls.Add(Me.gbFTP)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cbxCompania)
        Me.Controls.Add(Me.lbCompaniaAsistencia)
        Me.Controls.Add(Me.pbLogoMutua)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.picTest)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrincipalImportacion"
        Me.Text = "Siniestros: Area de Asistencia  -  Importación Datos"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.gbFTP.ResumeLayout(False)
        CType(Me.sbPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbPanel4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    'Public directori As String
    Private ObjExterno As Object

    Private Sub frmPrincipalImportacion_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InicioApp()
    End Sub

    Private Sub InicioApp()
        Dim ConecPruebas, strParametro As String

        'IdProceso = "I"
        rsProcesos = New ADODB.Recordset
        'strParametro = Microsoft.VisualBasic.Command

        Select Case parametroApp 'UCase(strParametro)

            Case "P" '  Ejecución Programada
                TipoEjecucion = "P"

            Case "PP" ' Ejecución Programa en Pruebas
                claseBDImportar.ConnexionPruebas()
                TipoEjecucion = "P"
                picTest.Show()
                picTest.BringToFront()

            Case "PM" ' Ejecución Pruebas Manual
                claseBDImportar.ConnexionPruebas()
                TipoEjecucion = "M"
                picTest.Show()
                picTest.BringToFront()

            Case Else ' Ejecución Manual Real
                TipoEjecucion = "M"

        End Select

        HaySuplidos = False
        frmInstImportacion = Me

        Dim s As String

        ' Creación de objetos
        '
        'Set clses = New clsmdpSesion

        ' Asignación de valores y parametros iniciales
        '
        If CodUserApli = "" Then CodUserApli = Microsoft.VisualBasic.Command

        'directori = clses.GetParam("PathImport")
        'directori = "K:\Siniestros\Asistencia\Import"
        'PathIconos = clses.GetParam("PathGraficos")
        PathIconos = "K:\Graficos\"
        'PathImportacion = clses.GetParam("PathImport")
        PathImportacion = "K:\Siniestros\Asistencia\Import"
        'PathReports = clses.GetParam("PathReports")
        PathReports = "K:\Reports\"
        'DiscoFTP = clses.GetParam("DiscoFTP")
        DiscoFTP = "195.77.230.7"
        'UsuarioFTP = clses.GetParam("UsuarioFTP")
        UsuarioFTP = "mprop@"
        'PasswordFTP = clses.GetParam("PasswordFTP")
        PasswordFTP = "-pw ultpg.2"
        'ConfigFTP = clses.GetParam("ConfigFTP")
        ConfigFTP = "K:\Siniestros\Asistencia\FTP"
        'DatosFTPApe = clses.GetParam("DatosApe")
        DatosFTPApe = "K:\Siniestros\Asistencia\FTP\Aperturas"
        'DatosFTPPag = clses.GetParam("DatosPag")
        DatosFTPPag = "K:\Siniestros\Asistencia\FTP\Pagos"

        LlenarComboCias(cbxCompania, lbxCompania)
        LlenarComboProcesos(cbxProceso)
        cbxProceso.SelectedIndex = 0
        cbxCompania.SelectedIndex = 0

        ' Asignación de objetos gráficos
        '/*MUL T-19908 INI
        'If Codcia <> "I" Then
        If Codcia <> "I" And Codcia <> "M" And Codcia <> "E" Then
            '/*MUL T-19908 FIN
            gbFTP.Enabled = False
            cbBajarAperturas.Enabled = False
            cbBajarPagos.Enabled = False
            pbFTP.Enabled = False
        End If
    End Sub

    Private Sub cbBajarAperturas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBajarAperturas.Click
        Dim resultadoMensaje As MsgBoxResult
        Dim objAngelImportacion As New clsImportar_NET
        Dim strParametro As String

        ' En pruebas no conectamos al ftp
        '/*MUL T-19908 INI
        strParametro = Microsoft.VisualBasic.Command

        'If strParametro = "PM" Then Return
        '/*MUL T-19908 FIN

        resultadoMensaje = MsgBox("Por favor, confirme que desea bajar los archivos de aperturas y anulaciones del disco externo FTP", MsgBoxStyle.YesNo)

        If resultadoMensaje = MsgBoxResult.Yes Then

            ' Declaraciones de objetos y variables
            '
            '/*MUL T-19908 INI
            'If Codcia = "I" Then
            If Codcia = "I" Or Codcia = "M" Or Codcia = "E" Then
                '/*MUL T-19908 FIN
                If Not objAngelImportacion.BajarFTPApe Then
                    MsgBox("Se ha producido un error en el proceso de bajada de ficheros del disco FTP. Avise a Informática")
                Else
                    globalNumerr = ""
                    resultadoMensaje = MsgBox("Desea eliminar en el disco FTP (origen) los ficheros acumulados ya bajados ?", MsgBoxStyle.YesNo)
                    System.Windows.Forms.Application.DoEvents()
                    If resultadoMensaje = MsgBoxResult.Yes Then
                        If Not objAngelImportacion.DeleteFTP("Aperturas") Then
                            MsgBox("Se ha producido un error al eliminar los ficheros acumulados del disco origen FTP, la operación no se ha realizado.")
                        End If
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub cbBajarPagos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBajarPagos.Click
        Dim resultadoMensaje As MsgBoxResult
        Dim objAngelImportacion As clsImportar_NET
        Dim strParametro As String

        globalNumerr = ""

        ' En pruebas no conectamos al ftp
        '/*MUL T-19908 INI
        strParametro = Microsoft.VisualBasic.Command

        If strParametro = "PM" Then Return
        '/*MUL T-19908 FIN

        resultadoMensaje = MsgBox("Por favor, confirme que desea bajar los archivos de pagos del disco externo FTP", MsgBoxStyle.YesNo)
        'System.Windows.Forms.Application.DoEvents()

        If resultadoMensaje = MsgBoxResult.Yes Then

            ' Declaraciones de objetos y variables
            '

            ' Creación de instancias e inicialización de propiedades
            '
            objAngelImportacion = New clsImportar_NET

            '/*MUL T-19908 INI
            'If Codcia = "I" Then
            If Codcia = "I" Or Codcia = "M" Or Codcia = "E" Then
                '/*MUL T-19908 FIN 
                If Not objAngelImportacion.BajarFTPPag Then
                    MsgBox("Se ha producido un error en el proceso de bajada de ficheros del disco FTP. Avise a Informática", MsgBoxStyle.Critical)
                Else
                    resultadoMensaje = MsgBox("¿Desea eliminar en el disco FTP (origen) los ficheros acumulados ya bajados?", MsgBoxStyle.YesNo)
                    globalNumerr = ""
                    'objError.Tipo = mdpErroresMensajes.Tipo.Pregunta
                    'objError.Ver(IdProceso, globalNumerr, "Desea eliminar en el disco FTP (origen) los ficheros acumulados ya bajados ?", Codcia)
                    'System.Windows.Forms.Application.DoEvents()
                    If resultadoMensaje = MsgBoxResult.Yes Then
                        If Not objAngelImportacion.DeleteFTP("Pagos") Then
                            MsgBox("Se ha producido un error al eliminar los ficheros acumulados del disco origen FTP, la operación no se ha realizado.", MsgBoxStyle.Critical)
                            'globalNumerr = "4064"
                            'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                            'objError.Ver(IdProceso, globalNumerr, , Codcia)
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub cbBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBuscar.Click
        On Error GoTo Selector_Err

        ' Declaraciones
        '
        Dim StrNomFich As String
        Dim PathNomFich As String
        Dim NumFicheros As Long
        Dim numIndice As Long
        Dim i, x As Integer

        ' Establece los parametros y valores iniciales para el box de dialogo
        ' de busqueda de ficheros, tipo fichero, path etc.
        '
        'dlgFicheros.InitialDirectory = objUtilidades.PathFromFileName(lbxFicheros.List(lstFicheros.ListIndex))
        'dlgFicheros.FileName = objUtilidades.NameFromFileName(lbxFicheros.ite.List(lstFicheros.ListIndex))
        dlgFicheros.Filter = "Texto (*.txt)|*.txt"
        dlgFicheros.FilterIndex = 1
        dlgFicheros.Multiselect = True
        dlgFicheros.CheckFileExists = True
        dlgFicheros.CheckPathExists = True
        dlgFicheros.InitialDirectory = PathImportacion


        If (dlgFicheros.ShowDialog() <> DialogResult.OK) Then Return
        PathNomFich = System.IO.Path.GetDirectoryName(dlgFicheros.FileName)
        For numIndice = 0 To UBound(dlgFicheros.FileNames)
            StrNomFich = System.IO.Path.GetFileName(dlgFicheros.FileNames(numIndice))


            'PathNomFich = objUtilidades.PathFromFileName(dlgFicheros.FileName) + "\"
            'StrNomFich = Trim(objUtilidades.NameFromFileName(dlgFicheros.FileName))
            lbxFicheros.Items.Add(PathNomFich & "\" & StrNomFich)
            'lbxFicheros.AddItem(PathNomFich + Mid(StrNomFich, x))
            'StrNomFich = Trim(Mid(StrNomFich, 1, x - 1))
        Next

        Exit Sub

Selector_Err:
        If Err.Number <> 32755 Then
            globalNumerr = "4020"
        End If
    End Sub

    Private Sub cbImportar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbImportar.Click

        On Error GoTo Importar_Err

        ' Declaraciones
        '
        Dim objAngelImportacion As clsImportar_NET
        Dim i As Short

        ' Validación de los parametros de importación
        '
        If Not ValidarDatos() Then
            Exit Sub
        End If
        '/*MUL T-19908 INI*/
        Select Case Codcia
            Case "A", "R", "I", "M", "E"
                '/*MUL T-19908 FIN*/
                objAngelImportacion = New clsImportar_NET
                If objAngelImportacion.Inicializar(cbxProceso.SelectedIndex, lbxFicheros) Then
                    globalNumerr = "4030"
                    MsgBox("El proceso de importación de datos de compañías de asistencia ha finalizado con exito.", MsgBoxStyle.Information)
                Else
                    MsgBox("Se han producido errores en el proceso de importación de datos de compañías de Asistencia")
                End If

        End Select
        lbxFicheros.Items.Clear()
        Exit Sub

Importar_Err:
        MsgBox("El proceso de Importación de datos de Asistencia ha devuelto un error.", MsgBoxStyle.Critical)
        globalNumerr = "4019"
    End Sub

    Public Function ValidarDatos() As Boolean

        On Error GoTo ValidarDatos_Error

        ValidarDatos = True

        If cbxCompania.SelectedIndex = -1 Then
            Err.Raise(2500)
        End If

        If cbxProceso.SelectedIndex = -1 Then
            Err.Raise(2501)
        End If

        If lbxFicheros.Items.Count = 0 And Not cbxProceso.SelectedIndex = 6 Then
            Err.Raise(4003)
        End If

        Exit Function

ValidarDatos_Error:

        Select Case Err.Number

            Case 2500 ' No se ha escogido compañia de aistencia
                globalNumerr = "4011"
                strError = "No se ha seleccionado ninguna compañía de asistencia"

            Case 2501 ' No se ha escogido ningún proceso de importación
                globalNumerr = "4023"
                strError = "No se ha seleccionado ningún proceso de importación a ejecutar"

            Case 4003 ' No se han seleccionado ficheros de importación
                globalNumerr = "4003"
                strError = "No se ha seleccionado ningún fichero para importar"

            Case Else
                globalNumerr = "4015"
                strError = "Error no catalogado en el proceso de validación."


        End Select

        MsgBox(strError, MsgBoxStyle.Critical)
        'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
        'objError.Ver(IdProceso, "4003", , Codcia)
        ValidarDatos = False

    End Function

    Private Sub cbxCompania_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxCompania.SelectedIndexChanged
        Codcia = Trim(lbxCompania.Items.Item(cbxCompania.SelectedIndex))

        If DatosCiaAsistencia(Codcia) Then
            '/*MUL T-19908 INI*/
            'If Codcia = "I" THEN
            If Codcia = "I" Or Codcia = "E" Or Codcia = "M" Then
                '/*MUL T-19908 FIN*/ 
                gbFTP.Enabled = True
                pbFTP.Enabled = True
                cbBajarAperturas.Enabled = True
                cbBajarPagos.Enabled = True
            Else
                gbFTP.Enabled = False
                pbFTP.Enabled = False
                cbBajarAperturas.Enabled = False
                cbBajarPagos.Enabled = False
            End If
        Else
            MsgBox("No hay datos", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub cbAvisoErrores_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAvisoErrores.Click
        On Error GoTo cbAvisos_Error

        ' Declaraciones
        '
        Dim objlistitem As ListViewItem
        Dim strReferencia As String
        Dim frmInstanciaErrores As New frmVisorErrores

        frmInstanciaErrores.Show()

        frmInstanciaErrores.MostrarErrores("")

        Exit Sub

cbAvisos_Error:
        MsgBox("Ha ocurrido un error mostrando el aviso", MsgBoxStyle.Critical)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Application.Exit()
    End Sub

End Class
