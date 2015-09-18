Module mdlVariables
    ''' exportar
    Friend claseBDLibrerias As New clsBD_NET
    Friend claseBDExportacion As New clsBD_NET
    Friend frmInstanciaPrincipal As frmPrincipalExportacion

    Friend ObjExterno As Object
    Public texte As Object
    Public fectxt As String
    Public Cabecera As ADODB.Recordset
    Public dteFechaIni As Date
    Public PathIconos As String
    Public strRiesgo As String
    Public NomIniFichero As String ' Numero de ficheros en los quer se ha
    ' de dividir la exportación
    Public Fich1, Fich2 As String
    Public Fich3 As String ' Nombres de fichero definidos en la
    ' tabal mpasicias para cada compañia
    ' de asistencia
    Public Fich4 As String
    Public Fich5 As String ' Nombres de fichero definidos en la
    ' tabal mpasicias para cada compañia
    ' de asistencia
    Public FechaProceso As Date
    Public SelectFich1 As String
    Public SelectFich2 As String
    Public SelectFich3 As String ' Cada una de las Select que devuelven los datos para los ficheros
    Public SelectFich4 As String ' Cada una de las Select que devuelven los datos para los ficheros
    Public SelectFich5 As String ' Cada una de las Select que devuelven los datos para los ficheros
    'Public objSiniestros As mdpCapaNegocioSiniestros.clsAsistencia
    Public objSiniestros As New clsAsistencia_NET
    'Public objUtilidades As mdpUtilidades.clsmdpUtilidades
    Public objUtilidades As New clsUtilidades_NET
    'Public objErr As mdpErroresMensajes.clsVisorLog
    Public strErr As String = "" ' cadena que contiene el código de error
    Public IdProceso As String
    Public Codcia As String ' Contiene el código de la compañía de asistencia seleccionada
    Public Descia As String ' Nombre de la Compañia de Asistencia
    Public NIFCompa As String ' NIF compañía asistencia
    Public DirecCompa As String ' Dirección Compañía asistencia
    Public PoblaCompa As String ' Población Compañia asistencia
    Public CodPobCompa As String ' Código Postal compañia asistencia
    Public NombreCompa As String ' Nombre Compañia asistencia
    Public Numcompa As String ' Numero Compañia asistencia
    Public GlobalNumErr As String ' Código de error
    Public CiaDefault As String ' Compañia Asist.de arranque por defecto
    Public TipoEjecucion As String ' Indica si la aplicación se ejecuta en modo Manual o Programado
    Public Modo As String ' Indica el modo del proceso ( Acumulado/Diario )
    Public UsuaApli As String
    Public LoteLeido As Double ' Número ordinal del último Lote leído de MpAsicias
    Public LoteEnviado As String ' Lote enviado en el email
    Public NumRegPolizas As Double ' Número de registros de polizas grabados
    Public NumRegRiesgos As Double ' Número de registros de riesgos grabados
    Public NumRegGarantias As Double ' Número de registros de garantias grabados
    Public NumRegPeritajes As Double ' Número de registros de peritajes grabados
    Public FicheroFusion As String ' Nombre del fichero que fusiona el resto de ficheros en uno solo para IP
    Public HayPeritajes As Boolean ' Indica si hay peritajes en la fecha del proceso
    Public HayCruce As Boolean ' Inidca si hay referencia spara cruzar en la fecha del proceso
    Public Pruebas As Boolean
    Public HaySPA As Boolean ' Indica si hay pólizas SPA
    Public FicheroLog As String ' Ruta y nombre del fichero de log
    Public PosCom As String ' Indica la posición del asterisco en el command$

    '/* MUL INI
    Public FtpCiaExport As String 'ruta del directorio donde realizar el export de ficheros
    Public FtpCiaImport As String 'ruta del directorio donde realizar el import de ficheros
    Public strSQLVisor As String
    Public strNombreCompa As String
    '/* MUL FIN

    '''' IMPORTAR
    Friend claseBDImportar As New clsBD_NET
    Friend frmInstImportacion As frmPrincipalImportacion
    Public PathImportacion As String
    Public PathReports As String
    Public DiscoFTP As String
    Public UsuarioFTP As String
    Public PasswordFTP As String
    Public ConfigFTP As String
    Public DatosFTPApe As String
    Public DatosFTPPag As String
    'Public objSiniestros As mdpCapaNegocioSiniestros.clsAsistencia
    'Public objUtilidades As mdpUtilidades.clsmdpUtilidades
    Public rsProcesos As ADODB.Recordset
    Public ProcesoIni As String
    Public strError As String
    Public IdReferCompa As String
    Public Transaccion As Boolean
    Public HaySuplidos As Boolean
    Public CodUserApli As String


    '''''''''''' aperturas
    'Variables Globales
    Friend claseBDAperturas As New clsBD_NET
    Friend clasePolizaAperturas As New clsPoliza_NET
    Friend claseUtilidadesAperturas As New clsUtilidades_NET
    Friend claseSiniestroAperturas As New clsSiniestro_NET
    Friend claseAsistenciaAperturas As New clsAsistencia_NET

    'Variable de instancia de formulario
    Friend frmInstAperturas As frmPrincipalAperturas

    Friend strFiltro As String
    Friend strNIFCompa As String
    Friend strSigCompa As String
    Friend strCodProducto As String
    Friend dbProvis As Double
    Friend bwflag As Boolean
    Friend strNumRec As String
    Friend strAgente As String
    Friend strNomProd As String
    Friend strGlobalNumErr As String
    Friend strDirecCompa As String
    Friend strPoblaCompa As String
    Friend strCodPobCompa As String
    Friend strNumCompa As String
    Friend strCodCia As String
    Friend strIdReferCompa As String
    Friend strUsuarioAplicacion As String
    Friend strCodUserAplicacion As String
    Friend strCampoBuscaAvanzada As String
    Friend strValorBuscaAvanzada As String
    Friend strIdProceso As String
    Friend boolTransaccion As Boolean ' Indicador de Transacción activa
    Friend strSQLSel As String ' Parte Select de la Sql
    Friend strWhere As String ' Parte Where de la Sql
    Friend strWhereMas As String ' Parte Where variable de la Sql
    Friend strCiaDefault As String
    Friend dbRPDtoTotal As Double ' Incluye descuento especial Reparalia
    Friend dbRPDtoBase As Double ' Incluye descuento especial Reparalia
    Friend dbRPDtoIVA As Double ' Incluye descuento especial Reparalia
    Friend strSQLCR As String ' Asume la instrucción SQL de generación de grid para generar registros temporales en mdpINFSiniestros
    Friend strIDComp As String ' Identificador de componente
    Friend strPathIconos As String ' Ruta donde se encuentran los archivos graficos
    Friend bVuelta As Boolean
    '''''''''
    ''pagos
    'Variables Globales
    Friend claseBDPagos As New clsBD_NET
    Friend clasePolizaPagos As New clsPoliza_NET
    Friend claseUtilidadesPagos As New clsUtilidades_NET
    Friend claseSiniestroPagos As New clsSiniestro_NET
    Friend claseAsistenciaPagos As New clsAsistencia_NET

    'Variable de instancia de formulario
    Friend frmInstPagos As frmPrincipalPagos

    ''Friend strError As String
    ''Friend strFiltro As String
    ''Friend strNIFCompa As String
    ''Friend strNombreCompa As String
    ''Friend strSigCompa As String
    ''Friend strCodProducto As String
    ''Friend dbProvis As Double
    ''Friend bwflag As Boolean
    ''Friend strNumRec As String
    ''Friend strAgente As String
    ''Friend strNomProd As String
    ''Friend strGlobalNumErr As String
    ''Friend strDirecCompa As String
    ''Friend strPoblaCompa As String
    ''Friend strCodPobCompa As String
    ''Friend strNumCompa As String
    ''Friend strCodCia As String
    ''Friend strIdReferCompa As String
    ''Friend strUsuarioAplicacion As String
    ''Friend strCodUserAplicacion As String
    ''Friend strCampoBuscaAvanzada As String
    ''Friend strValorBuscaAvanzada As String
    ''Friend strIdProceso As String
    ''Friend boolTransaccion As Boolean ' Indicador de Transacción activa
    ''Friend strSQLSel As String ' Parte Select de la Sql
    ''Friend strWhere As String ' Parte Where de la Sql
    ''Friend strWhereMas As String ' Parte Where variable de la Sql
    ''Friend strCiaDefault As String
    ''Friend dbRPDtoTotal As Double ' Incluye descuento especial Reparalia
    ''Friend dbRPDtoBase As Double ' Incluye descuento especial Reparalia
    ''Friend dbRPDtoIVA As Double ' Incluye descuento especial Reparalia
    ''Friend strSQLCR As String ' Asume la instrucción SQL de generación de grid para generar registros temporales en mdpINFSiniestros
    ''Friend strSQLVisor As String
    ''Friend strIDComp As String ' Identificador de componente
    ''Friend strPathIconos As String ' Ruta donde se encuentran los archivos graficos
    ''Friend bVuelta As Boolean
    '''''''''' suplidos
    'JCLopez_i
    'Codigo antiguo
    'Public objError As mdpErroresMensajes.clsVisorLog
    'JCLopez_f

    'Variables globales
    Friend claseBDSuplidos As New clsBD_NET
    Friend objSiniestro As New clsSiniestro_NET
    Friend objUtiles As New clsUtilidades_NET
    Friend objAsistencia As New clsAsistencia_NET

    'Variable de instancia de formulario
    Friend frmInstSuplidos As frmPrincipalSuplidos
    Friend frmInstanciaErrores As frmVisorErrores

    'Public strPathIconos As String
    Public gstrError As String
    'Public strFiltro As String
    'Public strNIFCompa As String
    'Public strNombreCompa As String
    'Public strSigCompa As String
    Public Provis As Double
    'Public bwflag As Boolean
    Public Numrec As String
    Public Agente As String
    'Public strNomProd As String
    'Public GlobalNumErr As String
    'Public strDirecCompa As String
    'Public strPoblaCompa As String
    'Public strCodPobCompa As String
    'Public strNumCompa As String
    'Public strCodcia As String
    'Public strIdReferCompa As String
    'Public UsuaApli As String
    'Public CodUserApli As String
    Public CampoBuscaAvanzada As String
    Public ValorBuscaAvanzada As String
    'Public IdProceso As String
    'Public Transaccion As Boolean ' Indicador de Transacción activa
    'Public strSqlSel As String ' Parte Select de la Sql
    'Public strWhere As String ' Parte Where de la Sql
    'Public strWhereMas As String ' Parte Where variable de la Sql
    'Public strCiaDefault As String
    Public Arranque As Boolean ' Indica si se está ejecutando el arranque
    Public SuplidosExisten As Boolean
    'Public TipoEjecucion As String
    Public ImporteIva As Double
    Public Iva As Short

    'Variables de los reports
    'Public strSQLCR As String ' Asume la instrucción SQL de generación de grid para generar registros temporales en mdpINFSiniestros
    'Public strIDComp As String ' Identificador de componente
    'Public strSQLVisor As String

    '''' cierres
    'Variables Globales
    Friend claseBDCierres As New clsBD_NET
    Friend clasePolizaCierres As New clsPoliza_NET
    Friend claseUtilidadesCierres As New clsUtilidades_NET
    Friend claseSiniestroCierres As New clsSiniestro_NET
    Friend claseAsistenciaCierres As New clsAsistencia_NET

    'Variable de instancia de formulario
    Friend frmInstCierres As frmPrincipalCierres
    Friend frmInstanciaAvisoBloqueos As New frmAvisosBloqueos
    Friend frmInstanciaAnulaciones As New frmAnulaciones

    'Public gstrError As String
    'Public strFiltro As String
    'Public strNIFCompa As String
    'Public strNombreCompa As String
    'Public strSigCompa As String
    Public strProvis As Double
    'Public bwflag As Boolean
    'Public strNumrec As String
    'Public strAgente As String
    'Public strNomProd As String
    'Public strGlobalNumErr As String
    'Public strDirecCompa As String
    'Public strPoblaCompa As String
    'Public strCodPobCompa As String
    'Public strNumcompa As String
    'Public strCodcia As String
    'Public strIdReferCompa As String
    Public strUsuaApli As String
    Public strCodUserApli As String
    'Public strCampoBuscaAvanzada As String
    'Public strValorBuscaAvanzada As String
    'Public strIdProceso As String
    'Public boolTransaccion As Boolean ' Indicador de Transacción activa
    'Public strSQLSel As String ' Parte Select de la Sql
    'Public strWhere As String ' Parte Where de la Sql
    'Public strWhereMas As String ' Parte Where variable de la Sql
    Public strFrom As String ' Parte From de la Sql
    Public strFromMas As String ' Parte From variable de la Sql
    'Public strCiaDefault As String
    Public sCierre As String ' Indica si un Siniestro se puede cerrar
    Public FechaCierre As DateTime
    Public objCierres As clsCierres_NET
    Public colAvisosBloqueo As Collection ' Es la colección de mensajes de bloqueo
    'Public PathIconos As String ' Ruta donde se encuentran los archivos graficos
    'Public PathReports As String ' Ruta donde se encuentran los ficheros de impresión
    'Public Modo As String ' Indica si se trabaja con siniestros o con anulaciones
    Public FechaEjecucion As Date ' Fecha de ejecuion del proceso
    Public Asunto As String ' Descripción para el asunto del email
    Public Mensaje As String ' Descripción para el texto del cuerpo del email
    Public colSiniestrosCerrados As Collection ' Colección de siniestros cerrados en el procso
    Public colSiniestrosPendientes As Collection ' Colección de siniestros que han quedado pendientes desde el último proceso lanzado
    Public NombreFichero As String ' Nombre del fichero de Log
    Public HoraTopeEjecucion As Object ' Hora máxima hasta la que se puede estar ejecutando
    'Public strCodProducto As String
    'Friend strIDComp As String ' Identificador de componente
    'Friend strSQLCR As String
    'Public strSQLVisor As String

End Module
