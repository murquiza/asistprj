Public Class clsBD_NET
    ' -------------------------------------------------------------------------------
    '
    '  Nombre......: clsmdpBD
    '  Tipo........: Módulo de clase
    '  Autor.......: Jose Luis de Lacalle
    '  Fecha.......: 1 de Marzo de 2002
    '  Descripción.: Esta clase reporta por un lado 2 objetos de conexión a la BD
    '                una a la BD de Sql Server y otro a AS400, también crea un
    '                RecordSet interno de acceso a las tablas de systemadel apli-
    '                cativo.Asi mismo suministra dos objetos recordset uno para
    '                cada BD
    '
    ' ----------------------------------------------------------------------------------


    ' Declaración de objetos y variables locales para asignaciones
    '
    Private WithEvents pBdSystem As ADODB.Connection ' Conexión BD de systema
    Private WithEvents pBdSystemRecord As ADODB.Recordset ' Acceso a las tablas de systema
    Private WithEvents pbdWork As ADODB.Connection ' Conexión a la BD de datos
    Private WithEvents pbdWorkRecord As ADODB.Recordset ' Acceso a registros BD de datos
    Private WithEvents pbdAuxRecord As ADODB.Recordset ' Objeto RecordSet auxiliar para acceso a datos
    Private pbdComand As ADODB.Command ' Objeto de comandos de ADO
    Private pCircuitoBd As String ' Entorno o sistema de la BD
    Private pBd As String ' Base de datos de trabajo
    Private pServer As String ' Servidor de la BD de trabajo
    Private pUserApp As String ' Usuario de la aplicación
    '/*MUL T-19908 INI 
    Private pUserPwd As String 'pwd usuario de la aplicación
    Private ZipCommand As String 'path Donde encontrar el winzip.exe
    Private pPathReports As String  'path donde encontrar los ficheros de reports
    '/*MUL T-19908 FIN 
    Private BDSystem As String ' Base de Datos de Sistema
    Private ServerSystem As String ' Servidor de la BD de sistema
    ' 28022007 Eloi JLL. Modificaciones para incluir modo de pruebas
    Private pBDPruebas As String 'BD de datos de pruebas
    Private pServerPruebas As String 'Servidor de BD de datos de pruebas
    Private pPath As String 'Path donde se guardan los ficheros
    
    Private Sub InicioApp()
        ' 28022007 Eloi JLL. Modificaciones para incluir modo de pruebas
        Dim TextoError As String
        ' Fin modificación
        On Error GoTo InicioApp_Error

        ' Cargamos los parametros de arranque de la BD
        '
        If Not ParametrosSistema() Then Err.Raise(1)

        ' Creación de objeto de acceso a Base de datos de Systema (SQL Server 7.0)
        '
        pBdSystem = New ADODB.Connection
        pBdSystem.ConnectionString = "Provider=SQLOLEDB.1;Password=admalfa;Persist Security Info=False;User ID=MdpPlus;Initial Catalog=" & BDSystem & ";" & "Data Source=" & ServerSystem

        'pBdSystem.ConnectionString = "Provider=SQLOLEDB.1;Password=1607;Persist Security Info=False;User ID=JLL;Initial Catalog=" & BDSystem & ";" & _
        ''                             "Data Source=" & ServerSystem

        ' Establecemos el tipo de acceso y ejecutamos la apertura de la BD
        '
        pBdSystem.Mode = ADODB.ConnectModeEnum.adModeReadWrite
        pBdSystem.Open(pBdSystem.ConnectionString)

        ' Creación objeto Recordset para lectura de registros BD de systema
        '
        pBdSystemRecord = New ADODB.Recordset

        ' Llamamos a la función que nos devuelve el nombre de la
        ' Base de Datos de trabajo
        '
        pCircuitoBd = BuscaBd(pBdSystemRecord, pBdSystem)
        If pCircuitoBd = "N" Or pCircuitoBd = "" Then Err.Raise(1000)

        ' Llamamos a la función que nos devuelve el nombre del
        ' Servidor de la Base de Datos de trabajo
        '
        Call BuscaServer(pBdSystemRecord, pBdSystem)

        ' Creación de objetos de acceso a datos de...
        '
        pbdWork = New ADODB.Connection

        '/*MUL T-19908 INI 
        pUserApp = "Usuarios"
        pwdApp = "All"
        '/*MUL T-19908 FIN

        Select Case pCircuitoBd

            Case "MdpPlus"

                ' A Orsis ( Sql 7 )
                '
                '/*MUL T-19908 INI 
                'pbdWork.ConnectionString = "Provider=SQLOLEDB.1;Password=All;Persist Security Info=False;User ID=Usuarios;Initial Catalog=" & pBd & ";" & "Data Source=" & pServer
                pbdWork.ConnectionString = "Provider=SQLOLEDB.1;Password=" & pwdApp & ";Persist Security Info=False;User ID=" & pUserApp & ";Initial Catalog=" & pBd & ";" & "Data Source=" & pServer
                '/*MUL T-19908 FIN 
            Case "Aries"

                '  A Aries ( DB2 AS400/IBM )
                '
                pbdWork.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Mode=ReadWrite;Extended Properties=" & "DSN=QDSN_AS400;SYSTEM=AS400;CMT=0;DBQ=ARI199E;NAM=0;DFT=5;DSP=1;TFT=0;TSP=0;" & "DEC=0;XDYNAMIC=1;RECBLOCK=2;BLOCKSIZE=32;SCROLLABLE=0;TRANSLATE=0;LAZYCLOSE=1;" & "LIBVIEW=0;REMARKS=0;CONNTYPE=0;SORTTYPE=0;PREFETCH=1;DFTPKGLIB=ARI199E;LANGUAGEID=ENU;" & "SORTWEIGHT=0;MAXFIELDLEN=32;COMPRESSION=1;ALLOWUNSCHAR=0;SEARCHPATTERN=1;MGDSN=0;"

            Case Else

                ' Desconocida
                '
                Err.Raise(1000)

        End Select

        ' Establecemos el modo de conexión y la abrimos
        '
        pbdWork.Mode = ADODB.ConnectModeEnum.adModeReadWrite
        pbdWork.Open(pbdWork.ConnectionString)

        ' Creamos los obetos RecordSet para acceso a datos
        '
        pbdWorkRecord = New ADODB.Recordset
        pbdWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseServer
        pbdAuxRecord = New ADODB.Recordset
        pbdAuxRecord.CursorLocation = ADODB.CursorLocationEnum.adUseServer
        pbdComand = New ADODB.Command
        ' 28022007 Eloi JLL. Modificaciones para incluir modo de pruebas
        TextoError = BuscaPathExport(pBdSystemRecord, pBdSystem)
        If TextoError = "N" Or TextoError = "" Then Err.Raise(1000)
        ' Fin Modificación
        Exit Sub

InicioApp_Error:
        'MsgBox "Se ha producido un error crítico a leer las tablas de sistema." & Chr$(10) & Chr$(13) + "MdpPlus no puede ejecutarse. Por favor, llame a Informática.", vbCritical, "Error crítico en tablas de sistema " & Err.Number
        BDSystem = "Error"
    End Sub
    Public Sub New()
        MyBase.New()
        InicioApp()
    End Sub

    ' Devuelve el usuario de la aplicación
    '

    ' Asigna el usuario de la aplicación
    '
    Public Property UserApp() As String
        Get
            UserApp = pUserApp
        End Get
        Set(ByVal Value As String)
            pUserApp = Value
        End Set
    End Property

    '/*MUL T-19908 INI 
    Public Property pwdApp() As String
        Get
            pwdApp = pUserPwd
        End Get
        Set(ByVal Value As String)
            pUserPwd = Value
        End Set
    End Property

    Public ReadOnly Property getZipCommand() As String
        Get
            getZipCommand = ZipCommand
        End Get
    End Property

    '/*MUL T-19908 FIN

    ' Propiedad que devuelve el nombre de la BD de Trabajo ( datos )
    '
    Public ReadOnly Property BDManagement() As String
        Get
            BDManagement = pCircuitoBd
        End Get
    End Property

    ' Propiedad que devuelve el nombre de la BD de Trabajo ( datos )
    '
    Public ReadOnly Property BDName() As String
        Get
            BDName = pBd
        End Get
    End Property

    ' Propiedad que devuelve el nombre del Servidor de  la BD de Trabajo ( datos )
    '
    Public ReadOnly Property BDServer() As String
        Get
            BDServer = pServer
        End Get
    End Property

    ' Propiedad que devuelve el nombre del servidor de la BD de sistema
    '
    Public ReadOnly Property ServerSys() As String
        Get
            ServerSys = ServerSystem
        End Get
    End Property

    ' Propiedad que devuelve el nombre de la BD de sistema
    '
    Public ReadOnly Property BDNameSys() As String
        Get
            BDNameSys = BDSystem
        End Get
    End Property

    ' Propiedad que devuelve el objeto de conexión a la BD de Sistema
    '
    Public ReadOnly Property BDSystemConnect() As ADODB.Connection
        Get
            BDSystemConnect = pBdSystem
        End Get
    End Property

    ' Propiedad que devuelve un RecordSet a la BD de Sistema
    '
    Public ReadOnly Property BDSystemRecord() As ADODB.Recordset
        Get
            BDSystemRecord = pBdSystemRecord
        End Get
    End Property

    ' Propiedad que devuelve el objeto de conexión a la BD de Trabajo
    '

    Public Property BDWorkConnect() As ADODB.Connection
        Get
            BDWorkConnect = pbdWork
        End Get
        Set(ByVal Value As ADODB.Connection)
            BDWorkConnect = Value
        End Set
    End Property

    ' Propiedad que devuelve un RecordSet a la BD de Trabajo
    '

    ' Propiedad que asigna un RecordSet al RecordSet de Trabajo
    '
    Public Property BDWorkRecord() As ADODB.Recordset
        Get
            BDWorkRecord = pbdWorkRecord
        End Get
        Set(ByVal Value As ADODB.Recordset)
            pbdWorkRecord = Value
        End Set
    End Property

    ' Propiedad que devuelve un RecordSet Auxiliar a la BD de Trabajo
    '

    ' Propiedad que asigna un RecordSet al ReordSet Auxiliar
    '
    Public Property BDAuxRecord() As ADODB.Recordset
        Get
            BDAuxRecord = pbdAuxRecord
        End Get
        Set(ByVal Value As ADODB.Recordset)
            pbdAuxRecord = Value
        End Set
    End Property


    ' Propiedad que devuelve un objeto Command sobre la BD
    '
    Public ReadOnly Property BDComand() As ADODB.Command
        Get
            BDComand = pbdComand
        End Get
    End Property

    Public ReadOnly Property PathExport() As String
        Get
            PathExport = pPath

        End Get
    End Property

    Private Sub FinApp()
        ' Destruimos todos los objetos locales
        '
        'pBdSystem = Nothing
        'pBdSystemRecord = Nothing
        'pbdWork = Nothing
        'pbdWorkRecord = Nothing
    End Sub

    Protected Overrides Sub Finalize()
        FinApp()
        MyBase.Finalize()
    End Sub

    ' Este procedimiento carga en un objeto RecordSet una Query existente en la
    ' la tabla mdpQuerys ( Tabla de Consultas )
    '
    Public Function QueryLoad(ByRef CodQuery As String) As ADODB.Recordset

        On Error GoTo QueryLoad_Error

        ' Declaraciones
        '
        Dim mdpSelect As String ' Parte Select de la Query
        Dim mdpFrom As String ' Parte From de la Query
        Dim mdpWhere As String ' Parte Where de la Query
        Dim mdpresto As String ' Resto de comandos de la Query ( Group By, Order, etc..)
        Dim sQuery As String ' Instrucción SQL completa
        Dim sQueryResult As String ' Instrucción SQL depurada

        ' Creamos la cadena pra la busqueda de la Sql en la BD de sistema
        ' La ejecutamos y depuramos
        '
        sQuery = "Select * From mdpQuerys Where Codigo = '" & CodQuery & "'"
        pBdSystemRecord.Open(sQuery, pBdSystem, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        pBdSystemRecord.MoveFirst()

        With pBdSystemRecord.Fields
            If .Item("Select").Value <> VariantType.Null Then
                sQueryResult = "Select " & .Item("Select").Value & " "
                If .Item("From").Value <> VariantType.Null Then
                    sQueryResult = sQueryResult & "From " & .Item("From").Value & " "
                    If .Item("Where").Value <> VariantType.Null Then
                        sQueryResult = sQueryResult & "Where " & .Item("Where").Value & " "
                    End If
                    If .Item("Resto").Value <> VariantType.Null Then
                        sQueryResult = sQueryResult & .Item("Resto").Value & " "
                    End If
                Else
                    Err.Raise(5000)
                End If
            Else
                Err.Raise(5000)
            End If
        End With
        pBdSystemRecord.Close()

        ' Ejecutamos la instrucción Sql sobre la Bd de trabajo
        '
        pbdWorkRecord.Open(sQueryResult, BDWorkConnect, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        QueryLoad = pbdWorkRecord

        Exit Function

QueryLoad_Error:

        If Err.Number = 5000 Then
            MsgBox("Hay un error en la confección de la SQL", MsgBoxStyle.Critical, " Error en definición SQL")
        Else
            MsgBox(Err.Description, MsgBoxStyle.OKOnly, " Error en consulta a la Base de datos ")
        End If

    End Function

    ' Esta función devuelve el nombre de la Base de Datos con la que se va a
    ' trabajar, y que está especificada en la tabla de systema mdpSysParametros
    '
    Private Function BuscaBd(ByRef SystemRecord As ADODB.Recordset, ByRef Conexion As ADODB.Connection) As String

        On Error GoTo BuscaBD_Error

        Dim Sql As String

        Sql = "Select Valor From mdpSysParametros Where Parametro = 'BaseDatos' or Parametro = 'CircuitoBD'"

        SystemRecord.Open(Sql, Conexion)
        SystemRecord.MoveFirst()
        pBd = SystemRecord.Fields(0).Value

        SystemRecord.MoveNext()
        BuscaBd = SystemRecord.Fields(0).Value

        SystemRecord.Close()
        Exit Function

BuscaBD_Error:
        BuscaBd = "N"
    End Function

    ' Esta función devuelve el nombre de la Base de Datos con la que se va a
    ' trabajar, y que está especificada en la tabla de systema mdpSysParametros
    '
    Private Sub BuscaServer(ByRef SystemRecord As ADODB.Recordset, ByRef Conexion As ADODB.Connection)
        On Error GoTo BuscaServer_Error

        Dim Sql As String

        Sql = "Select Valor From mdpSysParametros Where Parametro = 'ServidorDatos'"

        SystemRecord.Open(Sql, Conexion)
        SystemRecord.MoveFirst()
        pServer = SystemRecord.Fields(0).Value

        SystemRecord.Close()
        Exit Sub

BuscaServer_Error:
        pServer = "N"
    End Sub

    ' Esta función prueba la conexión con la BD con el usuario y password
    ' pasados como parámetros
    '
    Public Function TestConnect(ByRef Usuario As String, ByRef Password As String, ByRef BD As String, ByRef Server As String) As Boolean

        On Error GoTo TestConnect_Err

        ' Declaraciones
        '
        Dim bdConn As ADODB.Connection

        bdConn = New ADODB.Connection

        bdConn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Password & ";Persist Security Info=False;User ID=" & Usuario & ";Initial Catalog=" & BD & ";" & "Data Source=" & Server
        bdConn.Mode = ADODB.ConnectModeEnum.adModeReadWrite
        bdConn.Open(bdConn.ConnectionString)
        TestConnect = True
        'JCLopez_i
        'bdConn = Nothing
        'JCLopez_f

        Exit Function

TestConnect_Err:
        bdConn.Errors.Clear()
        'bdConn = Nothing
        TestConnect = False
    End Function

    Public Function ConfigurarImpesion() As Boolean
        On Error GoTo ConfigurarImpesion_err

        Dim l_idfitx As Integer
        Dim ls_reportFile As String
        Dim ls_lin As String
        Dim ls_rptdatabase As String
        Dim ls_rptserver As String
        Dim ls_linfitx() As String

        ConfigurarImpesion = False
        l_idfitx = FreeFile()
        If BuscaPathReports(pBdSystemRecord, pBdSystem) Then
            ls_reportFile = pPathReports & "AriesSQLServer_PM.dsn"
            If System.IO.File.Exists(ls_reportFile) = False Then
                Exit Function
            End If
            FileOpen(l_idfitx, ls_reportFile, OpenMode.Output)
        End If

        ' Cambiamos el dsn de conexion segun si estamos en pruebas o producción
        PrintLine(l_idfitx, "[Odbc]")
        PrintLine(l_idfitx, "DRIVER=SQL Server")
        PrintLine(l_idfitx, "UID=usuarios")
        PrintLine(l_idfitx, "DATABASE=" & pBd)
        PrintLine(l_idfitx, "WSID=JL_LACALLE")
        PrintLine(l_idfitx, "APP=Microsoft Open Database Connectivity")
        PrintLine(l_idfitx, "SERVER=" & pServer)

        FileClose(l_idfitx)
        ConfigurarImpesion = True
        Exit Function

ConfigurarImpesion_err:
        ConfigurarImpesion = True
    End Function

    ' Este procedimiento carga la Bd de sistema y el Servidor de sistema
    ' que figuran en el archivo mdpPlus.ini
    '
    Private Function ParametrosSistema() As Boolean

        On Error GoTo ParametrosSistema_Err

        ' Declaraciones
        '
        Dim Canal As Short
        Dim Linea As String
        Dim strParametro As String
        Dim ls_iniFile As String
        Dim ls_appPath As String
        Dim ls_appFileName As String
        Dim ls_appFileExt As String
        Dim lo_clsutil As clsUtilidades_NET
        Dim li_posAsterisco As Integer

        ' Obtenemos el numero de canal por el que abriremos el fichero
        '
        Canal = FreeFile()

        ' Apertura del fichero
        '
        '/*MUL INI 
        lo_clsutil = New clsUtilidades_NET
        ls_appPath = " "
        ls_appFileName = " "
        ls_appFileExt = " "

        strParametro = Microsoft.VisualBasic.Command
        'If strParametro = "P" Then
        '    'Producción
        '    FileOpen(Canal, ls_iniFile, OpenMode.Input)
        'Else
        li_posAsterisco = InStr(strParametro, "*")
        If li_posAsterisco > 0 Then
            strParametro = Trim(Mid(strParametro, 1, li_posAsterisco - 1))
        End If

        lo_clsutil.SplitPath(System.Windows.Forms.Application.ExecutablePath(), ls_appPath, ls_appFileName, ls_appFileExt)
        If strParametro = "PM" Or strParametro = "PP" Then
            'pruebas
            ls_iniFile = ls_appPath & "\MdpPlus_PM.ini"
        Else
            'Produccion
            ls_iniFile = ls_appPath & "\MdpPlus.ini"
        End If

        If System.IO.File.Exists(ls_iniFile) Then
            FileOpen(Canal, ls_iniFile, OpenMode.Input)
        Else
            'Por defecto vamos a este ini, el de producción !!
            FileOpen(Canal, "K:\Inicio\MdpPlus.ini", OpenMode.Input)
        End If
        '/*MUL FIN

        ' Bucle de lectura
        '
        '/*MUL INI 
        ZipCommand = "C:\Archivos de programa\WinZip\wzzip "
        '/*MUL FIN
        Do While Not EOF(Canal)
            Linea = LineInput(Canal)
            If Mid(Linea, 1, 8) = "DBSystem" Then
                BDSystem = Mid(Linea, 10, Len(Linea) - 9)
            End If
            If Mid(Linea, 1, 12) = "ServerSystem" Then
                ServerSystem = Mid(Linea, 14, Len(Linea) - 13)
            End If
            '/*MUL INI 
            If UCase(Mid(Linea, 1, 10)) = "ZIPCOMMAND" Then
                ZipCommand = Mid(Linea, 12, Len(Linea) - 10)
            End If
            '/*MUL FIN
        Loop

        If BDSystem = "" Or ServerSystem = "" Then Err.Raise(1)

        FileClose(Canal)
        ParametrosSistema = True
        Exit Function

ParametrosSistema_Err:
        ParametrosSistema = False
    End Function

    ' Este procedimiento elimnina el objeto de Base de Datos
    '
    Public Sub BdObjectClear()
        'pbdWork = Nothing
        'BDWorkConnect = Nothing
    End Sub

    ' 28022007 Eloi JLL. Modificaciones para incluir modo de pruebas
    Public Function ConnexionPruebas() As Object

        Dim TextoError As String

        'pbdWork = Nothing
        'pbdWorkRecord = Nothing
        'pbdAuxRecord = Nothing
        pbdWork = New ADODB.Connection


        TextoError = BuscarServidorPruebas(pBdSystemRecord, pBdSystem)
        If TextoError = "N" Or TextoError = "" Then Err.Raise(1000)
        TextoError = BuscaBdPruebas(pBdSystemRecord, pBdSystem)
        If TextoError = "N" Or TextoError = "" Then Err.Raise(1000)

        '/*MUL T-19908 INI 
        'pbdWork.ConnectionString = "Provider=SQLOLEDB.1;Password=All;Persist Security Info=False;User ID=Usuarios;Initial Catalog=" & pBDPruebas & ";" & "Data Source=" & pServerPruebas
        pbdWork.ConnectionString = "Provider=SQLOLEDB.1;Password=" & pwdApp & ";Persist Security Info=False;User ID=" & pUserApp & ";Initial Catalog=" & pBDPruebas & ";" & "Data Source=" & pServerPruebas
        '/*MUL T-19908 FIN 

        pbdWork.Mode = ADODB.ConnectModeEnum.adModeReadWrite
        pbdWork.Open(pbdWork.ConnectionString)

        pServer = pServerPruebas
        pBd = pBDPruebas

        ' Creamos los obetos RecordSet para acceso a datos
        '
        pbdWorkRecord = New ADODB.Recordset
        pbdWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseServer
        pbdAuxRecord = New ADODB.Recordset
        pbdAuxRecord.CursorLocation = ADODB.CursorLocationEnum.adUseServer

        TextoError = BuscaPathExportPruebas(pBdSystemRecord, pBdSystem)
        If TextoError = "N" Or TextoError = "" Then Err.Raise(1000)

        Exit Function
Initialize_Error:
        'MsgBox "Se ha producido un error crítico a leer las tablas de sistema." & Chr$(10) & Chr$(13) + "MdpPlus no puede ejecutarse. Por favor, llame a Informática.", vbCritical, "Error crítico en tablas de sistema " & Err.Number
        BDSystem = "Error, no se han podido leer las tablas del sistema"
    End Function

    Private Function BuscarServidorPruebas(ByRef SystemRecord As ADODB.Recordset, ByRef Conexion As ADODB.Connection) As String

        On Error GoTo BuscaServer_Error

        Dim Sql As String

        Sql = "Select Valor From mdpSysParametros Where Parametro = 'PrServidorDatos'"

        SystemRecord.Open(Sql, Conexion)
        SystemRecord.MoveFirst()
        pServerPruebas = SystemRecord.Fields(0).Value

        SystemRecord.Close()
        BuscarServidorPruebas = "S"
        Exit Function

BuscaServer_Error:
        BuscarServidorPruebas = "N"

    End Function

    Private Function BuscaBdPruebas(ByRef SystemRecord As ADODB.Recordset, ByRef Conexion As ADODB.Connection) As String

        On Error GoTo BuscaBD_Error

        Dim Sql As String

        Sql = "Select Valor From mdpSysParametros Where Parametro = 'PrBaseDatos'"

        SystemRecord.Open(Sql, Conexion)
        SystemRecord.MoveFirst()
        pBDPruebas = SystemRecord.Fields(0).Value

        SystemRecord.Close()
        BuscaBdPruebas = "S"
        Exit Function

BuscaBD_Error:
        BuscaBdPruebas = "N"
    End Function

    Private Function BuscaPathExport(ByRef SystemRecord As ADODB.Recordset, ByRef Conexion As ADODB.Connection) As String

        On Error GoTo BuscaBD_Error

        Dim Sql As String

        Sql = "Select Valor From mdpSysParametros Where Parametro = 'PathExport'"

        SystemRecord.Open(Sql, Conexion)
        SystemRecord.MoveFirst()
        pPath = SystemRecord.Fields(0).Value


        SystemRecord.Close()
        BuscaPathExport = "S"
        Exit Function

BuscaBD_Error:
        BuscaPathExport = "N"
    End Function

    Private Function BuscaPathExportPruebas(ByRef SystemRecord As ADODB.Recordset, ByRef Conexion As ADODB.Connection) As String

        On Error GoTo BuscaBD_Error

        Dim Sql As String

        Sql = "Select Valor From mdpSysParametros Where Parametro = 'PrPathExport'"

        SystemRecord.Open(Sql, Conexion)
        SystemRecord.MoveFirst()
        pPath = SystemRecord.Fields(0).Value


        SystemRecord.Close()
        BuscaPathExportPruebas = "S"

        Exit Function

BuscaBD_Error:
        BuscaPathExportPruebas = "N"
    End Function


    Private Function BuscaPathReports(ByRef SystemRecord As ADODB.Recordset, ByRef Conexion As ADODB.Connection) As Boolean
        On Error GoTo BuscaPathReports_Error

        Dim ls_sql As String

        ls_sql = "Select Valor From mdpSysParametros Where Parametro = 'PathReports'"

        SystemRecord.Open(ls_sql, Conexion)
        SystemRecord.MoveFirst()
        pPathReports = Trim(SystemRecord.Fields(0).Value)

        SystemRecord.Close()
        BuscaPathReports = True
        Exit Function

BuscaPathReports_Error:
        BuscaPathReports = False
    End Function

End Class
   