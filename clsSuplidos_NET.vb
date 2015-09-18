Public Class clsSuplidos_NET

    Private mvarReferencias As Collection ' Lista de pagos a procesar
    Private mvarFichero As Collection
    Private mstrCia As String ' Compañía
    Private TipoErr As String ' Tipo de error producido ( 'Aviso' o 'Error Severo')
    Private strError As String ' Mensaje de error a grabar en tabla de errores
    Private PrimaPendiente As String ' Indicador de pago de recibo pendiente

    ' Añade a la colección de referencias a procesar
    '

    ' Lee de la colección de Referencias a procesar
    '
    Public Property Referencias() As Collection
        Get
            Referencias = mvarReferencias
        End Get
        Set(ByVal Value As Collection)
            mvarReferencias = Value
        End Set
    End Property

    ' Lee de la colección de numero de orden a procesar
    '
    Public Property Fichero() As Collection
        Get
            Fichero = mvarFichero
        End Get
        Set(ByVal Value As Collection)
            Fichero = Value
        End Set
    End Property

    Public Function Filtros(ByRef objlistitem As ListViewItem) As Boolean

        ' Declaraciones
        '
        Dim strRamo As String
        Dim strPoliza As String
        Dim strNexp As String
        Dim strRefer As String
        Dim strPago As String
        Dim dteFechaSiniestro As Date
        Dim objCmd As ADODB.Command
        Dim strsql As String
        Dim lngSiniestros As Integer
        Dim lngNumPago As Integer
        Dim strCoderr As String
        Dim dteFechaPago As Date
        Dim errTipoObjeto As String
        Dim errRamo As String
        Dim errPoliza As String
        Dim errSiniestro As String
        Dim Archivo As String
        Dim FacturaIP As String
        Dim FechaFraIP As String

        On Error GoTo Filtros_Err

        Filtros = True
        '/* MUL INI
        TipoErr = ""
        strError = ""
        '/* MUL FIN

        ' Obtener Siniestro y Poliza
        '
        strPoliza = IIf(objlistitem.SubItems.Item(frmInstSuplidos.T7_NUMPOL.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstSuplidos.T7_NUMPOL.Index).Text), "0")
        strRamo = IIf(objlistitem.SubItems.Item(frmInstSuplidos.T7_CODRAM.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstSuplidos.T7_CODRAM.Index).Text), "0")
        strNexp = IIf(objlistitem.Text <> "", Trim(objlistitem.Text), "0")
        strRefer = IIf(objlistitem.SubItems.Item(frmInstSuplidos.T7_REFER.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstSuplidos.T7_REFER.Index).Text), "0")
        dteFechaPago = CDate(Now)
        strPago = Trim(objlistitem.SubItems.Item(frmInstSuplidos.T7_ESTADO.Index).Text)
        Archivo = Trim(objlistitem.SubItems.Item(frmInstSuplidos.Fichero.Index).Text)
        FacturaIP = IIf(objlistitem.SubItems.Item(frmInstSuplidos.NUMFACTURA.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstSuplidos.NUMFACTURA.Index).Text), "0")
        FechaFraIP = IIf(objlistitem.SubItems.Item(frmInstSuplidos.FECHA_FACT.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstSuplidos.FECHA_FACT.Index).Text), "0")

        ' Comprobación de filtros

        ' 1er Filtro
        ' ---------------------------------------------------------------------
        '  Este filtro comprueba que la compañía de asistencia  haya comunicado
        '  algún pago del siniestro del que vamos a pagar el suplido, aunque el
        '  importe esté aún pendiente de pago
        ' ---------------------------------------------------------------------

        ' Si el objeto recordset está abierto lo cerramos
        '
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()

        strsql = "Select Count(*) " & "From   Angel_T2 " & "Where  Angel_T2.T2_Refer = '" & strRefer & "' and Angel_T2.T2_Codsin = '" & strNexp & "'"

        claseBDSuplidos.BDWorkRecord.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDSuplidos.BDWorkRecord.EOF Then
            strCoderr = "SE001"
            strError = "El Siniestro " & strNexp & " no tiene registrado ningún pago, ni tiene pagos retenidos"
            TipoErr = "E"
            Err.Raise(1)
        End If
        claseBDSuplidos.BDWorkRecord.Close()

        ' 2do. Filtro
        ' ---------------------------------------------------------------------
        '  Este filtro detecta si ya se ha realizado el pago del suplido de la
        '  referencia de siniestro.
        ' ---------------------------------------------------------------------

        ' Si el objeto recordset está abierto lo cerramos
        '
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()

        strsql = "Select Count(*) " & "From   SuplidosAsistencia " & "Where  SuplidosAsistencia.T7_codsin = '" & strNexp & "' and SuplidosAsistencia.T7_Refer = '" & strRefer & "' and SuplidosAsistencia.Fichero <>'" & Archivo & "'"

        claseBDSuplidos.BDWorkRecord.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDSuplidos.BDWorkRecord.EOF Then
            If claseBDSuplidos.BDWorkRecord.Fields(0).Value > 0 Then
                strError = "El Siniestro " & strNexp & " tiene ya " & claseBDSuplidos.BDWorkRecord.Fields(0).Value - 1 & " pagos de suplidos"
                strCoderr = "SA001"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If

        ' 3º Filtro
        ' --------------------------------------------------------------------
        '  Reconocer aquellos siniestros que esten cerrados
        ' --------------------------------------------------------------------

        ' Si el objeto recordset está abierto lo cerramos
        '
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()

        strsql = "SELECT snsinies.estado, snsinies.feccie FROM snsinies WHERE snsinies.codsin = '" & strNexp & "'"
        claseBDSuplidos.BDWorkRecord.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If Not claseBDSuplidos.BDWorkRecord.EOF Then
            If claseBDSuplidos.BDWorkRecord.Fields("estado").Value = "C" Then
                strError = "El expediente " & strNexp & " ya está cerrado. La fecha de cierre es " & claseBDSuplidos.BDWorkRecord.Fields("feccie").Value & "'"
                strCoderr = "SP001"
                TipoErr = "E"
                Err.Raise(1)
            End If
        End If

        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()

        Exit Function

Filtros_Err:
        Filtros = False
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()
        If TipoErr = "E" Then
            Call ColorListItem(objlistitem, Color.Red)
        Else
            Call ColorListItem(objlistitem, Color.Goldenrod)
        End If
        If Err.Number = 55 Then
            Call InsertarError(strCoderr, (objlistitem.SubItems(frmInstSuplidos.T7_REFER.Index).Text), (objlistitem.Text), TipoErr, strError, "S", errTipoObjeto, errRamo, errPoliza, errSiniestro)
        Else
            Call InsertarError(strCoderr, (objlistitem.SubItems(frmInstSuplidos.T7_REFER.Index).Text), (objlistitem.Text), TipoErr, strError, "S")
        End If
        Call ActualizarEstado(TipoErr, strRefer, Archivo)
        objlistitem.SubItems.Item(frmInstSuplidos.T7_ESTADO.Index).Text = TipoErr
    End Function

    ' Procedure:  Pagar
    ' Objetivo:   Realiza la apertura del siniestro, inserta registro en las tablas de
    '             siniestros.
    ' Parametros: objListItem = Objeto ListItem de un ListView, con datos apertura.
    '             Retorno = Verdadero o Falso si pasa o no pasa el filtro
    '
    Public Function PagarSuplidos(ByRef objlistitem As ListViewItem) As Boolean

        On Error GoTo Pagar_Err

        ' Declaraciones
        '
        Dim strRamo As String
        Dim strPoliza As String
        Dim strPago As String
        Dim strRefer As String
        Dim Codsin As String
        Dim strsql As String
        Dim AcumPagos As Double
        Dim AcumProvisiones As Double
        Dim AcumProvPdte As Double
        Dim rsSiniestro As New ADODB.Recordset
        Dim strCoderr As String
        Dim Archivo As String
        Dim ImporteSuplido As Double
        Dim ImporteSumado As Double
        Dim FacturaIP As String
        Dim FechaFraIP As Date

        PagarSuplidos = True
        '/* MUL INI
        Transaccion = False
        TipoErr = ""
        '/* MUL FIN

        ' Verificar que Siniestro no este cerrado
        '
        Codsin = Trim(objlistitem.Text)
        If VerificaEstado(Codsin) Then

            ' Obtener Póliza, Fechas, Referencia, Ramo...
            '
            strPoliza = Trim(objlistitem.SubItems.Item(frmInstSuplidos.T7_NUMPOL.Index).Text)
            strPago = Trim(objlistitem.SubItems.Item(frmInstSuplidos.T7_ESTADO.Index).Text)
            strRefer = Trim(objlistitem.SubItems.Item(frmInstSuplidos.T7_REFER.Index).Text)
            Codsin = Trim(objlistitem.Text)
            Archivo = Trim(objlistitem.SubItems.Item(frmInstSuplidos.Fichero.Index).Text)
            FacturaIP = IIf(objlistitem.SubItems.Item(frmInstSuplidos.NUMFACTURA.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstSuplidos.NUMFACTURA.Index).Text), "0")
            FechaFraIP = IIf(objlistitem.SubItems.Item(frmInstSuplidos.FECHA_FACT.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstSuplidos.FECHA_FACT.Index).Text), "0")

            ' Obtener Acumulado de Pagos y Provisiones de Siniestros
            '
            rsSiniestro = objSiniestro.Siniestro(Codsin, False, UsuaApli)

            AcumPagos = IIf(IsDBNull(rsSiniestro.Fields("pte_pago").Value), 0, rsSiniestro.Fields("pte_pago").Value)
            AcumProvPdte = IIf(IsDBNull(rsSiniestro.Fields("propen_p").Value), 0, rsSiniestro.Fields("propen_p").Value)

            ' Obtener el porcentaje de IVA aplicable
            '
            Iva = IvaAplicableNuevo(strNumCompa, FechaFraIP)

            ' Establecemos el importe del suplido a pagar
            '
            ImporteIva = 0
            If strPago = "A" And frmInstSuplidos.chkFiltroAvisos.CheckState Then
                ImporteSuplido = CDec(frmInstSuplidos.txtSuplidoAdicional.Text)
            Else
                ImporteSuplido = CDec(frmInstSuplidos.txtSuplido.Text)
            End If
            If Iva > 0 Then
                ImporteIva = (ImporteSuplido * Iva) / 100
            End If
            ImporteIva = System.Math.Round(CDec(ImporteIva), 2)
            ImporteSuplido = System.Math.Round(CDec(ImporteSuplido), 2)
            ImporteSumado = System.Math.Round(CDec(ImporteSuplido + ImporteIva), 2)

            ' Inicio de la Transacción
            '
            claseBDSuplidos.BDWorkConnect.BeginTrans()
            Transaccion = True

            ' Actualizar Snsincta ( Tabla de Pagos )
            '
            If Not ActualizarSnSincta(objlistitem, Codsin, ImporteSuplido) Then
                strCoderr = "S0001"
                TipoErr = "E"
                Err.Raise(1)
            End If

            ' Actualizar Acumulados de Pagos y Provisión Pendiente en Snsinies
            '
            ' JLL - 11/4/2011
            ' salta el timeout debido a lentitud en la BBDD, por razones que desconozco
            ' no me deja compilar el módulo de mdpCapaNegocioSiniestro para establecer el TimeOut a 0
            ' y que no salte. Al compilar da error MEMORIA INSUFICIENTE
            ' Como los acumulados tambien se actualizan por la noche y en la ventande
            ' Orsis se calculan interactivamente es un tema que de momento podemos
            ' no ejecutar para salir del paso. Habrá que ver mas adelante el tema
            ' del errro de memoria para poder compilar
            '
            '    If Not ActualizarAcumulados(ImporteSuplido + ImporteIva, Codsin, AcumProvPdte, AcumPagos, rsSiniestro) Then
            '        strCoderr = "SA002"
            '        TipoErr = "A"
            '        Err.Raise 1
            '    End If

            ' Actualiza el estado del siniestro en la tabal de Pagos ( Angel_t2 )
            '
            If Not ActualizarEstado("P", (objlistitem.SubItems(frmInstSuplidos.T7_REFER.Index).Text), Archivo) Then
                strCoderr = "SA003"
                strError = "4072"
                TipoErr = "E"
                Err.Raise(1)
            End If

            If rsSiniestro.State = 1 Then rsSiniestro.Close()

            ' Final de la Transacción
            '
            claseBDSuplidos.BDWorkConnect.CommitTrans()
        Else
            Err.Raise(1)
            Transaccion = False
            TipoErr = "A"
        End If

        Transaccion = False
        'UPGRADE_NOTE: El objeto rsSiniestro no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'rsSiniestro = Nothing
        Exit Function

Pagar_Err:
        PagarSuplidos = False
        If Transaccion Then
            claseBDSuplidos.BDWorkConnect.RollbackTrans()
        End If
        If TipoErr = "E" Then
            Call ColorListItem(objlistitem, Color.Red)
        Else
            Call ColorListItem(objlistitem, Color.Goldenrod)
        End If
        objlistitem.SubItems.Item(frmInstSuplidos.T7_ESTADO.Index).Text = TipoErr
        If Not InsertarError(strCoderr, (objlistitem.SubItems(frmInstSuplidos.T7_REFER.Index).Text), Codsin, TipoErr, strError, IdProceso) Then
            strError = "Se ha producido un error crítico en el registro de la tabla de errores y avisos" & Chr(13) & Chr(10) & "El proceso de suplidos no puede continuar."
            MsgBox(strError, MsgBoxStyle.Exclamation)
            'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
            'objError.Ver(IdProceso, , strError, strCodcia)
        End If
        If strError <> "4041" Then
            Call ActualizarEstado(TipoErr, strRefer, Archivo)
        End If
        If claseBDSuplidos.BDWorkRecord.State = 1 Then claseBDSuplidos.BDWorkRecord.Close()

    End Function

    ' Procedure:  InsertarError
    ' Objetivo:   Inserta registro en la tabla MPASIHIST.
    ' Parametros: Referencia = Referencia del siniestros
    '             CodSin = Codigo de siniestro
    '             Error = Tipo de Error A/E
    '             Texto = Texto del error (descripcion)
    '             Proceso = Tipo proceso Aperturas/Pagos/...
    '
    'UPGRADE_NOTE: Error se actualizó a Error_Renamed. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
    Private Function InsertarError(ByRef CodError As String, ByRef Referencia As String, ByRef Codsin As String, ByRef Error_Renamed As String, ByRef Texto As String, ByRef Proceso As String, Optional ByRef strObjeto As Object = Nothing, Optional ByRef strRamo As Object = Nothing, Optional ByRef strPoliza As Object = Nothing, Optional ByRef strSiniestro As Object = Nothing) As Boolean

        On Error GoTo InsertarError_Err

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim lngNumero As Integer ' Número de Error
        Dim lngCero As Integer ' Variable long para ontener el código de ramo

        ' Valores iniciales
        '
        'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
        If Not IsNothing(strRamo) Then
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto strRamo. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            lngCero = Val(strRamo)
        End If

        ' Si el recordset está abierto lo cerramos
        '
        If claseBDSuplidos.BDWorkRecord.State = 1 Then claseBDSuplidos.BDWorkRecord.Close()

        ' Obtener el numero maximo de errores de una referencia
        '
        strsql = "SELECT IsNull(Max(numero),0) AS NUMERO " & "FROM   mpAsiHistError " & "WHERE  referencia = '" & Referencia & "' and proceso = '" & Proceso & "'"

        claseBDSuplidos.BDWorkRecord.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDSuplidos.BDWorkRecord.EOF Then
            lngNumero = 0
        Else
            lngNumero = claseBDSuplidos.BDWorkRecord.Fields("Numero").Value
        End If
        claseBDSuplidos.BDWorkRecord.Close()

        lngNumero = lngNumero + 1

        With claseBDSuplidos.BDWorkRecord
            .Open("mpAsiHistError", claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            .AddNew()
            .Fields("Referencia").Value = Referencia
            .Fields("Codsin").Value = Codsin
            .Fields("Numero").Value = lngNumero
            .Fields("Errores").Value = Error_Renamed
            .Fields("Texto").Value = Texto
            .Fields("Fecgra").Value = Today
            .Fields("Proceso").Value = Proceso
            .Fields("Cia").Value = strNombreCompa
            'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
            .Fields("TipoObjetoRel").Value = IIf(IsNothing(strObjeto), "", strObjeto)
            .Fields("RamoRel").Value = lngCero
            'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
            .Fields("PolizaRel").Value = IIf(IsNothing(strPoliza), "", strPoliza)
            'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
            .Fields("SiniestroRel").Value = IIf(IsNothing(strSiniestro), "", strSiniestro)
            .Fields("CodErr").Value = IIf(CodError = "", "0", CodError)
            .Fields("Codcia").Value = strCodcia
            .Update()
            .Close()
        End With
        InsertarError = True
        Exit Function

InsertarError_Err:
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()
        InsertarError = False
    End Function

    Public Sub New()
        MyBase.New()
        mvarReferencias = New Collection
        mvarFichero = New Collection
    End Sub

    ' Esta Función actualiza el estado de la referencia en la tabla de datos
    '
    Public Function ActualizarEstado(ByRef TipoErr As String, ByRef refererr As String, ByRef Fichero As String) As Boolean

        On Error GoTo ActualizarEstado_Err

        ' Declaraciones
        '
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim strsql As String ' Instrucción Sql
        Dim lngResult As Integer

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strsql = "Update SuplidosAsistencia Set T7_Estado = '" & TipoErr & "'," & "FechaProceso = '" & Now.Month & "/" & Now.Day & "/" & Now.Year & "' " & " Where T7_REFER = '" & refererr & "' and Fichero = '" & Fichero & "'"

        claseBDSuplidos.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDSuplidos.BDComand.CommandText = strsql
        claseBDSuplidos.BDComand.ActiveConnection = claseBDSuplidos.BDWorkConnect
        claseBDSuplidos.BDComand.Execute(lngResult)

        ActualizarEstado = True
        Exit Function

ActualizarEstado_Err:
        ActualizarEstado = False
    End Function

    ' Esta función graba los datos del pago en la tabla de Pagos Snsincta
    '
    Private Function ActualizarSnSincta(ByRef objlistitem As ListViewItem, ByRef Codsin As String, ByRef ImporteSuplido As Double) As Boolean

        On Error GoTo ActualizarSnSincta_Err

        ' Declaraciones
        '
        Dim ProvisionDbl As Double
        Dim Perjudicado As String
        Dim DiferenciaDbl As Double
        Dim ImporteTotal As Double

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto ImporteSuplido. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        ImporteTotal = System.Math.Round(CDec(ImporteSuplido + ImporteIva), 2)

        ' Comprobación de la suficiencia de la provisión pendiente para
        ' poder realizar el pago. Si no hubiera provisión suficiente re-
        ' lizamos una provisión por Ajuste Técnico por la diferencia +1
        '
        ProvisionDbl = CDbl(objSiniestro.ProvisionDisponible("", Codsin))
        DiferenciaDbl = System.Math.Round(CDec(ImporteTotal - ProvisionDbl), 2)

        ' Si la Provisión disponible queda a 0 se han de comprobar
        ' si quedan gestiones a 0 antes de grabar el pago
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto DiferenciaDbl. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If DiferenciaDbl = 0 Then
            If NumeroGestiones(Codsin) > 0 Then
                Err.Raise(1000)
            End If
        End If

        ' Si la provision disponible no es suficiente para realizar el pago
        ' se deberá realizar un ajuste de provisión por el deficit +1, para
        ' que el proceso nocturno no cierre el siniestro
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto DiferenciaDbl. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If DiferenciaDbl > 0 Then
            ' 3/12/2007 JLL No hace falta cambiar el sigo, ya que la falta de importe
            ' nos viene en positivo DiferenciaDbl = (DiferenciaDbl * -1) + 1
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto DiferenciaDbl. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            DiferenciaDbl = DiferenciaDbl + 1
        End If
        If ImporteTotal > ProvisionDbl Then
            If Not objSiniestro.AjusteTecnico(Codsin, DiferenciaDbl, "P", "AT", CodUserApli) Then
                ActualizarSnSincta = False
                Exit Function
            End If
        End If

        ' El objeto RecordSet debe estar cerrado antes de ejecutar un operacion 'Open'
        '
        If claseBDSuplidos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDSuplidos.BDWorkRecord.Close()

        ' Abrimos la tabla en la que vamos a grabar
        '
        claseBDSuplidos.BDWorkRecord.Open("SnSincta", claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        With claseBDSuplidos.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = Codsin
            .Fields("nummov").Value = objSiniestro.NumeroMovimientoTabla("Nummov", "Snsincta", "Codsin", Codsin)
            If .Fields("nummov").Value = "-1" Or .Fields("nummov").Value = "" Then Err.Raise(1)
            .Fields("Fecmov").Value = CDate(Now)
            .Fields("Tipgas").Value = "G"
            .Fields("SubTipGas").Value = "GV"
            .Fields("Pagado").Value = "N"
            .Fields("Contab").Value = "N"
            .Fields("Reaseg").Value = "N"
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto ImporteSuplido. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            .Fields("import").Value = ImporteSuplido
            .Fields("impiva").Value = ImporteIva
            .Fields("poriva").Value = Iva
            .Fields("numper").Value = strNumCompa
            If objlistitem.SubItems.Item(frmInstSuplidos.T7_ESTADO.Index).Text = "A" Then
                .Fields("Otdest").Value = Left("Suplido Ad: " & objlistitem.SubItems.Item(frmInstSuplidos.Fichero.Index).Text, 30)
            Else
                .Fields("Otdest").Value = Left("Suplido: " & objlistitem.SubItems.Item(frmInstSuplidos.Fichero.Index).Text, 30)
            End If
            .Fields("Codram").Value = objlistitem.SubItems.Item(frmInstSuplidos.T7_CODRAM.Index).Text
            .Fields("Fecgra").Value = CDate(Now)
            .Fields("Usuari").Value = CodUserApli
            .Fields("Numdoc").Value = Left(objUtiles.NameFromFileName(objlistitem.SubItems.Item(frmInstSuplidos.Fichero.Index).Text), 15)
            .Update()
            .Close()
        End With
        ActualizarSnSincta = True
        Exit Function

ActualizarSnSincta_Err:
        If strError = "" Then
            strError = "El registro de la tabla de pagos de Siniestros ha dado el error: " & Err.Description
        End If
        ActualizarSnSincta = False
    End Function

    ' Esta función devuelve el IVA aplicable
    '
    Private Function IvaAplicable(ByRef numper As String) As Short

        On Error GoTo IvaAplicable_Err

        ' Declaraciones
        '
        Dim strsql As String
        Dim rsIva As New ADODB.Recordset

        strsql = "Select iva From Proveedor Where Cod_Prov = '" & numper & "'"
        rsIva.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not rsIva.EOF Then
            IvaAplicable = rsIva.Fields("Iva").Value
        End If
        rsIva.Close()
        Exit Function

IvaAplicable_Err:
        rsIva.Close()
        IvaAplicable = 0
    End Function

    ' Esta función devuelve el IVA aplicable
    '
    Private Function IvaAplicableNuevo(ByRef numper As String, ByRef FechaFra As Date) As Short

        On Error GoTo IvaAplicableNuevo_Err

        If FechaFra = CDate("0:00:00") Then
            IvaAplicableNuevo = 16
            Exit Function
        End If

        Dim Fec As String

        'Fec = Format(FechaFra, "yyyy/mm/dd")
        Fec = FechaFra.Year & "/" & FechaFra.Month & "/" & FechaFra.Day

        ' Declaraciones
        '
        Dim strsql As String
        Dim rsIva As New ADODB.Recordset

        strsql = "Select  GrupIva.PorIva From Proveedor, GrupIva " & "Where   Proveedor.CodIva = GrupIva.Codiva and " & "Proveedor.Cod_Prov = '" & numper & "' and " & "(isNull(GrupIva.FecAlta,GetDate()) <= '" & Fec & "' and IsNull(GrupIva.FecBaja,'2100-01-01') > '" & Fec & "')"

        rsIva.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not rsIva.EOF Then
            IvaAplicableNuevo = rsIva.Fields("poriva").Value
        End If
        rsIva.Close()
        Exit Function

IvaAplicableNuevo_Err:
        Resume
        rsIva.Close()
        IvaAplicableNuevo = 0
    End Function

    ' Esta función actualiza los acumulados de pagos y la provisión
    ' pendiente del siniestro
    '
    Private Function ActualizarAcumulados(ByRef Importe As Double, ByRef Codsin As String, ByRef ProvPdte As Double, ByRef AcumPagos As Double, ByRef rsSinies As ADODB.Recordset) As Boolean

        On Error GoTo ActualizarAcumulados_Err

        ' Cálculos
        '
        ProvPdte = ProvPdte - Importe
        AcumPagos = AcumPagos + Importe

        ' Actualización de la Tabla
        '
        'rsSinies!propen_p = ProvPdte
        rsSinies.Fields("pte_pago").Value = AcumPagos
        rsSinies.Update()
        rsSinies.Close()
        ActualizarAcumulados = True
        Exit Function

ActualizarAcumulados_Err:
        ActualizarAcumulados = False
        strError = "Se ha producido un error en la actualización de los acumulados de pagos y provisión pendiente. " & Err.Description
    End Function

    ' Esta función borra las referencias pasadas en la colección de las tablas
    ' de Pagos
    '
    Public Function DeleteSuplidos(ByRef ColRefers As Collection, ByRef colrefer2 As Collection) As Boolean

        On Error GoTo DeleteSuplidos_Err

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim i As Short ' Contador para bucles


        ' Inicio de la transsacción
        '
        claseBDSuplidos.BDWorkConnect.BeginTrans()

        For i = 1 To ColRefers.Count()

            ' Primero borramos el expediente de la tabla Angel_t2
            '
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto colrefer2(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto ColRefers(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            strsql = "Delete From SuplidosAsistencia Where t7_codsin = '" & Trim(ColRefers.Item(i)) & "' and Fichero = '" & Trim(colrefer2.Item(i).Text) & "'"
            claseBDSuplidos.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
            claseBDSuplidos.BDComand.CommandText = strsql
            claseBDSuplidos.BDComand.ActiveConnection = claseBDSuplidos.BDWorkConnect
            claseBDSuplidos.BDComand.Execute()
        Next i

        claseBDSuplidos.BDWorkConnect.CommitTrans()
        DeleteSuplidos = True
        Exit Function

DeleteSuplidos_Err:
        claseBDSuplidos.BDWorkConnect.RollbackTrans()
        DeleteSuplidos = False
    End Function


    ' Esta función devuelve el numero de gestiones abiertas sin contar la gestión
    ' de reparación de la propia compañía de Asistencia
    '
    Private Function NumeroGestiones(ByRef scodsin As String) As Short

        On Error GoTo NumeroGestiones_Err

        ' Declaraciones
        '
        Dim strsql As String
        Dim localRs As ADODB.Recordset

        ' Creación de objetos
        '
        localRs = New ADODB.Recordset

        NumeroGestiones = 0

        ' Instrucción Sql de busqueda
        '
        strsql = "Select * " & "From   SnSinges " & "Where  Codsin = '" & scodsin & "'"

        ' Construimos el objeto RecodSet
        '
        claseBDSuplidos.BDWorkConnect.Errors.Clear()
        localRs.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If claseBDSuplidos.BDWorkConnect.Errors.Count > 0 Then Err.Raise(-1000)

        ' Incio bucle de comprobación
        '
        NumeroGestiones = 0
        With localRs
            Do While Not .EOF
                If .Fields("Numper").Value <> strNumCompa Then
                    'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                    If IsDBNull(.Fields("Fecrec").Value) Or .Fields("Fecrec").Value = "" Then
                        NumeroGestiones = NumeroGestiones + 1
                    End If
                Else
                    If .Fields("Tipges").Value <> "RE" Then
                        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                        If IsDBNull(.Fields("Fecrec").Value) Or .Fields("Fecrec").Value = "" Then
                            NumeroGestiones = NumeroGestiones + 1
                        End If
                    End If
                End If
                .MoveNext()
            Loop
        End With
        localRs.Close()

        If NumeroGestiones > 0 Then strError = "la disponibilidad quedará a 0 y existen gestiones abiertas "

        Exit Function

NumeroGestiones_Err:
        If Err.Number = 1000 Then
            strError = "No se han podido verificar las gestiones abiertas "
        Else
            strError = "la disponibilidad quedará a 0 y existen gestiones abiertas "
        End If
    End Function

    ' Función que devuelve True si el siniestro está abierto
    ' y False si esta cerrado
    '
    Public Function VerificaEstado(ByRef scodsin As String) As Boolean

        On Error GoTo VerificaEstado_Err

        ' Declaraciones
        '
        Dim strsql As String
        Dim localRs As ADODB.Recordset

        ' Creación de objetos
        '
        localRs = New ADODB.Recordset

        ' Instrucción Sql de busqueda
        '
        strsql = "Select Estado " & "From   SnSinies " & "Where  Codsin = '" & scodsin & "'"

        ' Construimos el objeto RecodSet
        '
        claseBDSuplidos.BDWorkConnect.Errors.Clear()
        localRs.Open(strsql, claseBDSuplidos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If claseBDSuplidos.BDWorkConnect.Errors.Count > 0 Then Err.Raise(-1000)

        If localRs.Fields.Item("Estado").Value = "P" Then
            VerificaEstado = True
        Else
            VerificaEstado = False
        End If

        localRs.Close()

        Exit Function

VerificaEstado_Err:
        VerificaEstado = False
    End Function
End Class
