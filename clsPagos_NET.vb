Friend Class clsPagos_NET

    Private mvarReferencias As New Collection ' Lista de pagos a procesar
    Private numorden As New Collection
    Private mstrCia As String ' Compañía
    Private TipoErr As String ' Tipo de error producido ( 'Aviso' o 'Error Severo')
    Private strError As String ' Mensaje de error a grabar en tabla de errores
    Private PrimaPendiente As String ' Indicador de pago de recibo pendiente

    ' Añade a la colección de referencias a procesar
    '

    ' Lee de la colección de Referencias a procesar
    Public Property Referencias() As Collection
        Get
            Referencias = mvarReferencias
        End Get
        Set(ByVal Value As Collection)
            mvarReferencias = Value
        End Set
    End Property
    ' Lee de la colección de numero de orden a procesar
    Public Property NumeroOrden() As Collection
        Get
            NumeroOrden = numorden
        End Get
        Set(ByVal Value As Collection)
            numorden = Value
        End Set
    End Property

    Public Function CountReferencias() As Long
        Return Referencias.Count
    End Function

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
        Dim NumOrd As String
        Dim dtFechaBaja As Date

        On Error GoTo Filtros_Err

        Filtros = True
        '/* MUL INI
        TipoErr = ""
        strCoderr = ""
        strError = ""
        '/* MUL FIN

        ' Obtener Siniestro y Poliza
        '
        strPoliza = IIf(objlistitem.SubItems.Item(frmInstPagos.T2_POLIZA.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstPagos.T2_POLIZA.Index).Text), "0")
        strRamo = IIf(objlistitem.SubItems.Item(frmInstPagos.T2_CODRAM.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstPagos.T2_CODRAM.Index).Text), "0")
        strNexp = IIf(objlistitem.Text <> "", Trim(objlistitem.Text), "0")
        strRefer = IIf(objlistitem.SubItems.Item(frmInstPagos.T2_REFER.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstPagos.T2_REFER.Index).Text), "0")
        dteFechaPago = CDate(Now)
        strPago = Trim(objlistitem.SubItems.Item(frmInstPagos.T2_ESTADO.Index).Text)
        NumOrd = Trim(objlistitem.SubItems.Item(frmInstPagos.T2_NUMORD.Index).Text)
        PrimaPendiente = "N"
        ' Comprobación de filtros

        ' 1er Filtro
        ' --------------------------------------------------------------------
        '  Este filtro comprueba que la póliza esté en vigor
        ' --------------------------------------------------------------------

        ' Si el objeto recordset está abierto lo cerramos
        '
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()
        strsql = "Select Polizaca.FecPol, Polizaca.FecEfe, Polizaca.Polanu, Polizaca.Fecbaj " & "From   Polizaca " & "Where  Polizaca.NumPol = '" & strPoliza & "' and Polizaca.Codram = '" & strRamo & "'"

        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDPagos.BDWorkRecord.EOF Then
            strCoderr = "PE001"
            strError = "La Póliza " & strPoliza & " no existe"
            TipoErr = "E"
            Err.Raise(1)
        Else

            If Not IsDBNull(claseBDPagos.BDWorkRecord.Fields("Fecbaj").Value) Then
                If claseBDPagos.BDWorkRecord.Fields("Polanu").Value = "S" And dteFechaPago >= dtFechaBaja Then
                    strError = "La Póliza " & strPoliza & " está anulada con efecto " & claseBDPagos.BDWorkRecord.Fields("Fecbaj").Value
                    strCoderr = "PE002"
                    TipoErr = "A"
                    Err.Raise(1)
                End If
            End If
        End If
        claseBDPagos.BDWorkRecord.Close()

        ' 2do. Filtro
        ' --------------------------------------------------------------------
        '  Detecta la duplicidad de pagos fijos de la gestión y de
        '  indemnización
        ' --------------------------------------------------------------------

        ' Primero comprobamos los gastos de gestión
        '
        If Trim(objlistitem.SubItems.Item(frmInstPagos.T2_REFER.Index).Text) = "G" Then
            strsql = "SELECT Count(*) FROM snsincta WHERE snsincta.codsin = '" & strNexp & "' and snsincta.tipgas = 'G'"
            claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If Not claseBDPagos.BDWorkRecord.EOF Then
                strError = "El expediente " & strNexp & " tiene ya un pago por gastos de  gestión"
                strCoderr = "PA002"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()

        ' ...y luego los de Indemnización
        '
        If Trim(objlistitem.SubItems.Item(frmInstPagos.T2_REFER.Index).Text) = "G" Then
            strsql = "SELECT Count(*) FROM snsincta WHERE snsincta.codsin = '" & strNexp & "' and snsincta.tipgas = 'I'"
            claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If Not claseBDPagos.BDWorkRecord.EOF Then
                strError = "El expediente " & strNexp & " tiene ya un pago por gastos de Indemnización"
                strCoderr = "PA003"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()

        ' 3to. Filtro
        ' --------------------------------------------------------------------
        '  Detectar aquellos siniestros denegados por Mutua
        ' --------------------------------------------------------------------

        strsql = "SELECT snsinies.denega, snsinies.fecdng, snsinies.usuden, isnull(snsinies.cadeim,'N') as cadeim,  isnull(nombre + ' ' + apell1 + ' ' + apell2,space(1)) as nombreEmp " & _
                 "FROM snsinies, empleado WHERE snsinies.usuden *= empleado.num_empl " & _
                 "and snsinies.codsin = '" & strNexp & "' " & _
                 "and denega = 'S' "

        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDPagos.BDWorkRecord.EOF Then
            '/* MUL no hace falta, la query solo extrae los denegados ....
            'If claseBDPagos.BDWorkRecord.Fields("denega").Value = "S" Then
            If claseBDPagos.BDWorkRecord.Fields("cadeim").Value = "S" Then
                strError = "El expediente " & strNexp & " se denegó en fecha " & claseBDPagos.BDWorkRecord.Fields("fecdng").Value & _
                            " por el usuario " & claseBDPagos.BDWorkRecord.Fields("nombreEmp").Value & _
                            ". Si se efectúa el pago, el siniestro pasará a situación de Reconsiderado."
                '" por el usuario " & claseSiniestroPagos.NombreEmpleado(claseBDPagos.BDWorkRecord.Fields("usuden").Value) & _

            Else
                strError = "El expediente " & strNexp & " se denegó en fecha " & claseBDPagos.BDWorkRecord.Fields("fecdng").Value & _
                            " por el usuario " & claseBDPagos.BDWorkRecord.Fields("nombreEmp").Value & _
                            ". Si se efectúa el pago, se anulará la denegación del siniestro."
                ' " por el usuario " & claseSiniestroPagos.NombreEmpleado(claseBDPagos.BDWorkRecord.Fields("usuden").Value) & _
            End If
            strCoderr = "PA004"
            TipoErr = "A"
            Err.Raise(1)
            ' End If
        End If
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()

        ' 4º Filtro
        ' --------------------------------------------------------------------
        '  Si la prima estuviera pendiente, retener dicho pago
        ' --------------------------------------------------------------------

        strsql = "SELECT carterac.numrec as recibo From snsinies, carterac " & "WHERE  snsinies.numrec = carterac.numrec and " & "       snsinies.codsin = '" & strNexp & "' and " & "       carterac.fesuvt <= snsinies.feccas and" & "       carterac.estado <> 'C'"
        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If Not claseBDPagos.BDWorkRecord.EOF Then
            strError = "El recibo " & claseBDPagos.BDWorkRecord.Fields("recibo").Value & " que da cobertura a este expediente esta pendiente de pago"
            PrimaPendiente = "S"
            strCoderr = "PA005"
            TipoErr = "A"
            Err.Raise(1)
        End If
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()

        ' 5º Filtro
        ' --------------------------------------------------------------------
        '  Reconocer aquellos siniestros que contengan Franquicias
        ' --------------------------------------------------------------------

        strsql = "Select isnull(Polizaca.franrc,0) as franrc " & "From   Polizaca " & "Where  Polizaca.Numpol = '" & strPoliza & "' And Polizaca.Codram = '" & strRamo & "'"
        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If Not claseBDPagos.BDWorkRecord.EOF Then
            If claseBDPagos.BDWorkRecord.Fields("franrc").Value > 0 Then
                strError = "La Póliza " & strPoliza & " contiene una franquicia de " & claseBDPagos.BDWorkRecord.Fields("franrc").Value & "€"
                strCoderr = "PA006"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()

        ' 6º Filtro
        ' --------------------------------------------------------------------
        '  Reconocer aquellos siniestros que esten cerrados
        ' --------------------------------------------------------------------

        strsql = "SELECT snsinies.estado, snsinies.feccie FROM snsinies WHERE snsinies.codsin = '" & strNexp & "'"
        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If Not claseBDPagos.BDWorkRecord.EOF Then
            If claseBDPagos.BDWorkRecord.Fields("estado").Value = "C" Then
                strError = "El expediente " & strNexp & " ya está cerrado. La fecha de cierre es " & claseBDPagos.BDWorkRecord.Fields("feccie").Value & "'"
                strCoderr = "PE003"
                TipoErr = "E"
                Err.Raise(1)
            End If
        End If
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()


        Exit Function

Filtros_Err:
        Filtros = False

        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    If claseBDPagos.BDWorkRecord.State <> ADODB.ObjectStateEnum.adStateClosed Then
            claseBDPagos.BDWorkRecord.Close()
            '    End If
        End If

        If TipoErr = "E" Then
            'Call ColorListItem(objlistitem, System.Drawing.Color.Red)
            Call ColorListItem(objlistitem, Color.Red)
        Else
            'Call ColorListItem(objlistitem, System.Drawing.ColorTranslator.FromOle(&H1F8EC5))
            Call ColorListItem(objlistitem, Color.Goldenrod)
        End If
        If Err.Number = 55 Then
            Call InsertarError(strCoderr, (objlistitem.SubItems(frmInstPagos.T2_REFER.Index).Text), (objlistitem.Text), TipoErr, strError, "P", errTipoObjeto, errRamo, errPoliza, errSiniestro)
        Else
            Call InsertarError(strCoderr, (objlistitem.SubItems(frmInstPagos.T2_REFER.Index).Text), (objlistitem.Text), TipoErr, strError, "P")
        End If
        Call ActualizarEstado(TipoErr, strRefer, NumOrd)
        objlistitem.SubItems.Item(frmInstPagos.T2_ESTADO.Index).Text = TipoErr

    End Function

    ' Procedure:  Pagar
    ' Objetivo:   Realiza la apertura del siniestro, inserta registro en las tablas de
    '             siniestros.
    ' Parametros: objListItem = Objeto ListItem de un ListView, con datos apertura.
    '             Retorno = Verdadero o Falso si pasa o no pasa el filtro
    '
    Public Function Pagar(ByRef objlistitem As ListViewItem) As Boolean

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
        Dim Iva As Short
        Dim strCoderr As String
        Dim NumOrd As String

        Pagar = True
        '/* MUL INI
        TipoErr = ""
        boolTransaccion = False
        strCoderr = ""
        strError = ""
        '/* MUL FIN

        ' Obtener Póliza, Fechas, Referencia, Ramo...
        '
        strPoliza = Trim(objlistitem.SubItems.Item(frmInstPagos.T2_POLIZA.Index).Text)
        strPago = Trim(objlistitem.SubItems.Item(frmInstPagos.T2_ESTADO.Index).Text)
        strRefer = Trim(objlistitem.SubItems.Item(frmInstPagos.T2_REFER.Index).Text)
        Codsin = Trim(objlistitem.Text)
        NumOrd = Trim(objlistitem.SubItems.Item(frmInstPagos.T2_NUMORD.Index).Text)

        ' Obtener Acumulado de Pagos y Provisiones de Siniestros
        '
        rsSiniestro = claseSiniestroPagos.Siniestro(Codsin, False, strUsuarioAplicacion)

        AcumPagos = IIf(IsDBNull(rsSiniestro.Fields("pte_pago").Value), 0, rsSiniestro.Fields("pte_pago").Value)
        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
        AcumProvPdte = IIf(IsDBNull(rsSiniestro.Fields("propen_p").Value), 0, rsSiniestro.Fields("propen_p").Value)

        ' Obtener el porcentaje de IVA aplicable
        '
        Iva = IvaAplicable(strNumCompa)

        ' Inicio de la Transacción
        '
        claseBDPagos.BDWorkConnect.BeginTrans()
        boolTransaccion = True

        ' Actualizar Snsincta ( Tabla de Pagos )
        '
        If Not ActualizarSnSincta(objlistitem, Codsin) Then
            strCoderr = "P0001"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Actualizar Acumulados de Pagos y Provisión Pendiente en Snsinies
        '
        If Not ActualizarAcumulados(objlistitem, Codsin, AcumProvPdte, AcumPagos, rsSiniestro) Then
            strCoderr = "PA001"
            TipoErr = "A"
            Err.Raise(1)
        End If

        ' Actualiza el estado del siniestro en la tabal de Pagos ( Angel_t2 )
        '
        If Not ActualizarEstado("P", (objlistitem.SubItems(frmInstPagos.T2_REFER.Index).Text), NumOrd) Then
            strCoderr = "PA007"
            strError = "4041"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Cambio, si procede de denegado a reconsiderado
        '
        If DenegadosReconsiderados(objlistitem, Codsin) = True Then
            Call HistoricoDenegaciones(objlistitem, Codsin)
        End If


        If rsSiniestro.State = 1 Then rsSiniestro.Close()

        ' Final de la Transacción
        '
        claseBDPagos.BDWorkConnect.CommitTrans()

        boolTransaccion = False
        'UPGRADE_NOTE: El objeto rsSiniestro no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'rsSiniestro = Nothing
        Exit Function

Pagar_Err:
        Pagar = False
        If boolTransaccion Then
            claseBDPagos.BDWorkConnect.RollbackTrans()
        End If
        If TipoErr = "E" Then
            Call ColorListItem(objlistitem, Color.Red)
        Else
            Call ColorListItem(objlistitem, Color.Goldenrod)
        End If
        objlistitem.SubItems.Item(frmInstPagos.T2_ESTADO.Index).Text = TipoErr
        If Not InsertarError(strCoderr, (objlistitem.SubItems(frmInstPagos.T2_REFER.Index).Text), Codsin, TipoErr, strError, strIdProceso) Then
            'strError = 
            'objError.TipoMensaje = mdpErroresMensajes_NET.clsVisorLog.Tipo.Pantalla
            MsgBox("Se ha producido un error crítico en el registro de la tabla de errores y avisos" & Chr(13) & Chr(10) & "El proceso de aperturas no puede continuar.")
            'objError.Ver(IdProceso, "", strError, Codcia)
        End If
        If strError <> "4041" Then
            Call ActualizarEstado(TipoErr, strRefer, NumOrd)
        End If
        If claseBDPagos.BDWorkRecord.State = 1 Then claseBDPagos.BDWorkRecord.Close()

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
    Private Function InsertarError(ByRef CodError As String, ByRef Referencia As String, ByRef Codsin As String, ByRef Error_Renamed As String, ByRef Texto As String, ByRef Proceso As String, Optional ByRef strObjeto As Object = Nothing, Optional ByRef strRamo As String = Nothing, Optional ByRef strPoliza As Object = Nothing, Optional ByRef strSiniestro As Object = Nothing) As Boolean

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
        If claseBDPagos.BDWorkRecord.State = 1 Then claseBDPagos.BDWorkRecord.Close()

        ' Obtener el numero maximo de errores de una referencia
        '
        strsql = "SELECT IsNull(Max(numero),0) AS NUMERO " & "FROM   mpAsiHistError " & "WHERE  referencia = '" & Referencia & "' and proceso = '" & Proceso & "'"

        claseBDPagos.BDWorkRecord.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDPagos.BDWorkRecord.EOF Then
            lngNumero = 0
        Else
            lngNumero = claseBDPagos.BDWorkRecord.Fields("Numero").Value
        End If
        claseBDPagos.BDWorkRecord.Close()

        lngNumero = lngNumero + 1

        If IsNothing(Error_Renamed) Then
            Error_Renamed = ""
        End If

        If IsNothing(Texto) Then
            Texto = ""
        End If

        If IsNothing(strNombreCompa) Then
            strNombreCompa = ""
        End If

        If IsNothing(Referencia) Then
            Referencia = ""
        End If

        If IsNothing(Codsin) Then
            Codsin = ""
        End If

        If IsNothing(strCodCia) Then
            Codsin = ""
        End If

        With claseBDPagos.BDWorkRecord
            .Open("mpAsiHistError", claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            .AddNew()
            .Fields("Referencia").Value = Referencia
            .Fields("Codsin").Value = Codsin
            .Fields("Numero").Value = lngNumero
            .Fields("Errores").Value = Error_Renamed
            .Fields("Texto").Value = Texto
            .Fields("Fecgra").Value = Today '' Araceli pide que sea el mismo que el de la carga -> t2_fgraba 
            .Fields("Proceso").Value = Proceso
            .Fields("Cia").Value = strNombreCompa
            .Fields("TipoObjetoRel").Value = IIf(IsNothing(strObjeto), "", strObjeto)
            .Fields("RamoRel").Value = lngCero
            .Fields("PolizaRel").Value = IIf(IsNothing(strPoliza), "", strPoliza)
            .Fields("SiniestroRel").Value = IIf(IsNothing(strSiniestro), "", strSiniestro)
            .Fields("CodErr").Value = IIf(CodError = "", "0", CodError)
            .Fields("Codcia").Value = strCodCia
            .Update()
            .Close()
        End With
        InsertarError = True
        Exit Function

InsertarError_Err:
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()
        InsertarError = False
    End Function

    ' Esta Función actualiza el estado de la referencia en la tabla de aperturas
    '
    Public Function ActualizarEstado(ByRef TipoErr As String, ByRef refererr As String, ByRef NumeroOrden As String) As Boolean

        On Error GoTo ActualizarEstado_Err

        ' Declaraciones
        '
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim strsql As String ' Instrucción Sql
        Dim lngResult As Integer

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strsql = "Update Angel_T2 Set T2_Estado = '" & TipoErr & "'," & "FechaProceso = '" & Format(CDate(Now), "MM/dd/yyyy") & "' " & " Where T2_REFER = '" & refererr & "' and T2_NUMORD = '" & NumeroOrden & "'"

        claseBDPagos.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDPagos.BDComand.CommandText = strsql
        claseBDPagos.BDComand.ActiveConnection = claseBDPagos.BDWorkConnect
        claseBDPagos.BDComand.Execute(lngResult)

        ActualizarEstado = True
        Exit Function

ActualizarEstado_Err:
        ActualizarEstado = False
    End Function

    ' Esta función graba los datos del pago en la tabla de Pagos Snsincta
    '
    Private Function ActualizarSnSincta(ByRef objlistitem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo ActualizarSnSincta_Err

        ' Declaraciones
        '
        Dim ImporteDbl As Double
        Dim ProvisionDbl As Double
        Dim Perjudicado As String
        Dim DiferenciaDbl As Double

        If strCodCia = "R" And frmInstPagos.chkDtoReparalia.Checked Then
            If DescuentoReparalia((objlistitem.SubItems.Item(frmInstPagos.TOTAL.Index).Text)) Then
                objlistitem.SubItems.Item(frmInstPagos.TOTAL.Index).Text = CStr(dbRPDtoTotal)
            End If
        End If

        ' Comprobación de la suficiencia de la provisión pendiente para
        ' poder realizar el pago. Si no hubiera provisión suficiente re-
        ' lizamos una provisión por Ajuste Técnico por la diferencia +1
        '
        ImporteDbl = CDbl(objlistitem.SubItems.Item(frmInstPagos.TOTAL.Index).Text)
        ProvisionDbl = CDbl(claseSiniestroPagos.ProvisionDisponible("", Codsin))
        DiferenciaDbl = System.Math.Round(CDec(ImporteDbl - ProvisionDbl), 2)

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
            ' nos viene en positivo
            'DiferenciaDbl = (DiferenciaDbl * -1) + 1
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto DiferenciaDbl. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            DiferenciaDbl = DiferenciaDbl + 1
        End If
        If ImporteDbl > ProvisionDbl Then
            If Not claseSiniestroPagos.AjusteTecnico(Codsin, DiferenciaDbl, "P", "AT", strCodUserAplicacion) Then
                ActualizarSnSincta = False
                Exit Function
            End If
        End If

        ' El objeto RecordSet debe estar cerrado antes de ejecutar un operacion 'Open'
        '
        If claseBDPagos.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDPagos.BDWorkRecord.Close()

        ' Abrimos la tabla en la que vamos a grabar
        '
        claseBDPagos.BDWorkRecord.Open("SnSincta", claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        With claseBDPagos.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = Codsin
            .Fields("nummov").Value = claseSiniestroPagos.NumeroMovimientoTabla("Nummov", "Snsincta", "Codsin", Codsin)
            If .Fields("nummov").Value = "-1" Or .Fields("nummov").Value = "" Then Err.Raise(1)
            .Fields("Fecmov").Value = CDate(Now)
            .Fields("Tipgas").Value = "I"
            .Fields("SubTipGas").Value = "II"
            .Fields("Pagado").Value = "N"
            .Fields("Contab").Value = "N"
            .Fields("Reaseg").Value = "N"
            .Fields("import").Value = objlistitem.SubItems.Item(frmInstPagos.TOTAL.Index).Text
            .Fields("numper").Value = strNumCompa
            Perjudicado = claseAsistenciaPagos.NumeroPerjudicadoAsistencia(Codsin, strNumCompa)
            If Perjudicado <> "Error" Then
                .Fields("Numprj").Value = Perjudicado
            Else
                strError = "No ha sido creado el perjudicado de Asistencia"
                Err.Raise(1)
            End If
            'JCLopez_i
            'No se encuentra el campo T2_FACTURA en el ListView pero en VB6 no petaba. Se pone el campo de la factura.
            .Fields("Numdoc").Value = objlistitem.SubItems.Item(frmInstPagos.FACTURA.Index).Text
            '.Fields("Numdoc").Value = objlistitem.SubItems.Item(frmInstanciaPrincipal.T2_FACTURA.Index).Text
            'JCLopez_f
            .Fields("Codram").Value = objlistitem.SubItems.Item(frmInstPagos.T2_CODRAM.Index).Text
            If strCodCia <> "R" Then
                '.Fields("Codmod").Value = objlistitem.SubItems.Item("T2_CODMOD").Text
                '.Fields("Codgru").Value = objlistitem.SubItems.Item("T2_CODGRU").Text
                .Fields("Codmod").Value = objlistitem.SubItems.Item(frmInstPagos.MODO_GAR.Index).Text
                .Fields("Codgru").Value = objlistitem.SubItems.Item(frmInstPagos.GRUPO_GAR.Index).Text
            End If
            .Fields("Fecgra").Value = CDate(Now)
            .Fields("Usuari").Value = strCodUserAplicacion
            If strCodCia = "R" And frmInstPagos.chkDtoReparalia.Checked Then
                .Fields("otdest").Value = "Dto.Reparalia 10%"
            End If
            .Update()
            .Close()
        End With
        ActualizarSnSincta = True
        Call CierraReparacion(Codsin)
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
        rsIva.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not rsIva.EOF Then
            IvaAplicable = rsIva.Fields("Iva").Value
        End If
        rsIva.Close()
        Exit Function

IvaAplicable_Err:
        rsIva.Close()
        IvaAplicable = 0
    End Function

    ' Esta función actualiza los acumulados de pagos y la provisión
    ' pendiente del siniestro
    '
    Private Function ActualizarAcumulados(ByRef objlistitem As ListViewItem, ByRef Codsin As String, ByRef ProvPdte As Double, ByRef AcumPagos As Double, ByRef rsSinies As ADODB.Recordset) As Boolean

        On Error GoTo ActualizarAcumulados_Err

        ' Cálculos
        '
        ProvPdte = ProvPdte - Val(objlistitem.SubItems.Item(frmInstPagos.TOTAL.Index).Text)
        AcumPagos = AcumPagos + Val(objlistitem.SubItems.Item(frmInstPagos.TOTAL.Index).Text)

        ' Declaraciones
        '
        Dim strsql As String
        Dim Rslocal As ADODB.Recordset

        ' Creación de objetos
        '
        Rslocal = New ADODB.Recordset

        ' Construimos el objeto RecodSet de apertura
        '
        ' Instrucción Sql de busqueda
        '
        strsql = "Select * " & "From   SnSinies " & "Where  Codsin = '" & Codsin & "'"

        claseBDPagos.BDWorkConnect.Errors.Clear()
        Rslocal.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDPagos.BDWorkConnect.Errors.Count > 0 Then Err.Raise(-1000)

        ' Actualización de la Tabla
        '
        Rslocal.Fields("propen_p").Value = ProvPdte
        Rslocal.Fields("pte_pago").Value = AcumPagos
        Rslocal.Update()
        Rslocal.Close()
        ActualizarAcumulados = True
        Exit Function

ActualizarAcumulados_Err:
        ActualizarAcumulados = False
        strError = "Se ha producido un error en la actualización de los acumulados de pagos y provisión pendiente. " & Err.Description
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
        claseBDPagos.BDWorkConnect.Errors.Clear()
        localRs.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If claseBDPagos.BDWorkConnect.Errors.Count > 0 Then Err.Raise(-1000)

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

    ' Este procedimiento realiza el cierre de la gestión de reparación de la
    ' compañía de Asistencia
    '
    Private Sub CierraReparacion(ByRef scodsin As String)

        On Error GoTo CierraReparacion_Err

        Dim strsql As String
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim lngResult As Integer

        'strsql = "Update Snsinges Set Fecrec = '" & VB6.Format(CDate(Now), "mm/dd/yyyy") & "' " & "Where  Codsin = '" & scodsin & "' and Tipges = 'RE' and " & "       Numper = '" & Numcompa & "' and Fecrec is null"
        strsql = "Update Snsinges Set Fecrec = '" & Now.Month & "/" & Now.Day & "/" & Now.Year & "' " & "Where  Codsin = '" & scodsin & "' and Tipges = 'RE' and " & "       Numper = '" & strNumCompa & "' and Fecrec is null"

        claseBDPagos.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDPagos.BDComand.CommandText = strsql
        claseBDPagos.BDComand.ActiveConnection = claseBDPagos.BDWorkConnect
        claseBDPagos.BDComand.Execute(lngResult)

        Exit Sub

CierraReparacion_Err:
    End Sub
    ' Esta función borra las referencias pasadas en la colección de las tablas
    ' de Pagos
    Public Function DeletePagos(ByRef ColRefers As Collection, ByRef colrefer2 As Collection) As Boolean

        On Error GoTo Deletepagos_Err

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim i As Short ' Contador para bucles


        ' Inicio de la transsacción
        '
        claseBDPagos.BDWorkConnect.BeginTrans()

        For i = 1 To ColRefers.Count()

            ' Primero borramos el expediente de la tabla Angel_t2
            '
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto colrefer2(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto ColRefers(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            strsql = "Delete From  Angel_T2 Where t2_codsin = '" & Trim(ColRefers.Item(i)) & "' and t2_numord = '" & Trim(colrefer2.Item(i)) & "'"
            claseBDPagos.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
            claseBDPagos.BDComand.CommandText = strsql
            claseBDPagos.BDComand.ActiveConnection = claseBDPagos.BDWorkConnect
            claseBDPagos.BDComand.Execute()
        Next i

        claseBDPagos.BDWorkConnect.CommitTrans()
        DeletePagos = True
        Exit Function

Deletepagos_Err:
        claseBDPagos.BDWorkConnect.RollbackTrans()
        DeletePagos = False
    End Function

    '' Comprueba si el siniestro está denegado para pasarlo a reconsiderado
    '
    Private Function DenegadosReconsiderados(ByRef objlistitem As ListViewItem, ByRef scodsin As String) As Boolean

        ' Declaraciones
        '
        Dim strsql As String
        Dim Rslocal As New ADODB.Recordset

        DenegadosReconsiderados = False

        strsql = "Select Denega, Comode, Desden, fecdng, (select vcp_descri from valorconcepto where vcp_codvcp='SNTTRAMOTREC01') AS DesRec, " & _
                 "isnull(Cadeim,''), feimcd, decade, cadeco, isnull((select centralizado from ramos where ramos.codram=snsinies.codram),'') AS centralizado," & _
                 "isnull((select email from empleado where num_empl=(select tramitador from ramos where ramos.codram=snsinies.codram)),'') AS tramitador_ramo, " & _
                 "isnull((select email from empleado where num_empl=(select tramitador from agentes where agentes.codage=snsinies.codage)),'') AS tramitador_agente, " & _
                 "isnull((select email from empleado where num_empl=(SELECT s_user FROM SnDenega_Hist WHERE codsin='" & scodsin & "' AND numhist=1)),'') AS tramitador_denegacion " & _
                 "From SnSinies Where Codsin ='" & scodsin & "'"
        Rslocal.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        'si esta denegado
        Dim ls_email As String
        If Rslocal.Fields(0).Value = "S" Then
            'si la carta de denegacion ha sido enviada
            If Rslocal.Fields(5).Value = "S" Then
                Rslocal.Fields(0).Value = "R"
                Rslocal.Fields(1).Value = "SNTTRAMOTREC01"
                Rslocal.Fields(2).Value = Rslocal.Fields(4).Value
                Rslocal.Fields(3).Value = CDate(Now)
            Else
                Rslocal.Fields(0).Value = "N"
                Rslocal.Fields(1).Value = vbNullString
                Rslocal.Fields(2).Value = vbNullString
                Rslocal.Fields(3).Value = VariantType.Null
                'Rslocal.Fields(5).Value = vbnullstring
                Rslocal.Fields(6).Value = VariantType.Null
                Rslocal.Fields(7).Value = vbNullString
                Rslocal.Fields(8).Value = vbNullString


                'Si el email del usuario que denegó el siniestro se usa el email del tramitador
                'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                If Trim(Rslocal.Fields(12).Value) = "" Or IsDBNull(Rslocal.Fields(12).Value) Then
                    'para buscar el tramitador dependerá si el ramo es centralizado o no
                    If Trim(Rslocal.Fields(9).Value) = "S" Then
                        ls_email = Trim(Rslocal.Fields(10).Value)
                    Else
                        ls_email = Trim(Rslocal.Fields(11).Value)
                    End If
                Else
                    'email del usuario que denegó el siniestro
                    ls_email = Trim(Rslocal.Fields(12).Value)
                End If

                'si no se encontró ningún email se envía a control gestión
                'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
                If Trim(ls_email) = "" Or IsDBNull(ls_email) Then
                    ls_email = "cgsiniestros"
                End If

                If Right(Trim(ls_email), 23) <> "@mutuadepropietarios.es" Then
                    ls_email = Trim(ls_email) & "@mutuadepropietarios.es"
                End If

                Call enviarMail(ls_email, "Anulación Siniestro " & scodsin, "Aviso automático: Se ha anulado automáticamente la denegación del siniestro " & scodsin & " tras efectuarse un pago de indemnización.")

            End If
            Rslocal.Update()
            Rslocal.Close()
            DenegadosReconsiderados = True
        End If

    End Function


    Private Function HistoricoDenegaciones(ByRef objlistitem As ListViewItem, ByRef scodsin As String) As Boolean
        Dim strsql As String
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim lngResult As Integer
        Dim Rslocal As New ADODB.Recordset



        strsql = "Select comode,Desden,Decade,Cadeco,fecdng,isnull(cadeim,'') From SnSinies Where Codsin ='" & scodsin & "'"
        Rslocal.Open(strsql, claseBDPagos.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Rslocal.Fields(5).Value = "S" Then
            'strsql = "INSERT INTO SnDenega_Hist (codsin,numhist,tipo,comode,desmot,s_user,fechagra,decade,cadeco,fecdng) SELECT '" & scodsin & "',(SELECT MAX(numhist)+1 FROM SnDenega_Hist WHERE codsin='" & scodsin & "'),'R','" & Rslocal.Fields(0).Value & "','" & Rslocal.Fields(1).Value & "','" & CodUserApli & "','" & VB6.Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & Rslocal.Fields(2).Value & "','" & Rslocal.Fields(3).Value & "','" & VB6.Format(Rslocal.Fields(4).Value, "yyyy-mm-dd hh:mm:ss") & "'"
            strsql = "INSERT INTO SnDenega_Hist (codsin,numhist,tipo,comode,desmot,s_user,fechagra,decade,cadeco,fecdng) SELECT '" & scodsin & _
                     "',(SELECT MAX(numhist)+1 FROM SnDenega_Hist WHERE codsin='" & scodsin & "'),'R','" & Rslocal.Fields(0).Value & "','" & _
                    Rslocal.Fields(1).Value & "','" & strCodUserAplicacion & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & _
                    Rslocal.Fields(2).Value & "','" & Rslocal.Fields(3).Value & "','" & Format(Rslocal.Fields(4).Value, "yyyy-MM-dd hh:mm:ss") & "'"
        Else
            strsql = "DELETE FROM SnDenega_Hist WHERE codsin='" & scodsin & "'"
        End If

        claseBDPagos.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDPagos.BDComand.CommandText = strsql
        claseBDPagos.BDComand.ActiveConnection = claseBDPagos.BDWorkConnect
        claseBDPagos.BDComand.Execute(lngResult)

        Rslocal.Close()

    End Function


    Private Function enviarMail(ByRef strSendTo As String, ByRef strSubject As String, ByRef strBody As String) As Object
        Const APP_NAME As String = "LPSendMail.DLL"

        'Error Messages
        Const ERR_SENDTO As String = "Send to email address has not been set."
        Const ERR_SUBJECT As String = "Subject of email has not been set."
        Const ERR_FILE As String = "The following attachment file does not exist."

        'Error Numbers
        Const ERR_NO_SENDTO As Integer = 10101
        Const ERR_NO_SUBJECT As Integer = 10102
        Const ERR_NO_FILE As Integer = 10103

        'Object Declaration
        Dim mobjSession As Object
        Dim mobjNotesDB As Object
        Dim mobjMailDoc As Object
        Dim mobjAttachment As Object

        'Variable Declaration
        Dim mstrSendTo As String
        Dim mstrCopyTo As String
        Dim mstrSubject As String
        Dim mstrBody As String
        Dim mstrAttachFile As String
        Dim mstrSendToArray() As String

        '-----------------------------
        'Instantiate objects
        mobjSession = CreateObject("Notes.Notessession")
        'try setting the server and maybe the db
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjSession.GETDATABASE. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        mobjNotesDB = mobjSession.GETDATABASE("", "")
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjNotesDB.OPENMAIL. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Call mobjNotesDB.OPENMAIL()
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjNotesDB.CREATEDOCUMENT. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        mobjMailDoc = mobjNotesDB.CREATEDOCUMENT
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjMailDoc.Form. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        mobjMailDoc.Form = "Memo"
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjMailDoc.CREATERICHTEXTITEM. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        mobjAttachment = mobjMailDoc.CREATERICHTEXTITEM("BODY")

        '--------------------------
        Dim lCount As Integer

        If mstrSendTo <> "" Then
            lCount = UBound(mstrSendToArray)
            ReDim Preserve mstrSendToArray(lCount + 1)
            mstrSendToArray(lCount + 1) = strSendTo
        Else
            mstrSendTo = strSendTo
            ReDim mstrSendToArray(1)
            mstrSendToArray(1) = strSendTo
        End If
        '--------------------------

        mstrSubject = strSubject
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjMailDoc.Subject. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        mobjMailDoc.Subject = mstrSubject
        '--------------------------
        mstrBody = strBody
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjAttachment.APPENDTEXT. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Call mobjAttachment.APPENDTEXT(mstrBody)
        '--------------------------

        On Error GoTo ErrorHandler

        'Ensure the Sendto has been set.
        If UBound(mstrSendToArray) < 1 And mstrSendTo = "" Then
            'SendMail = False
            Err.Raise(ERR_NO_SENDTO, APP_NAME, ERR_SENDTO)
            Exit Function
        Else
            If UBound(mstrSendToArray) > 1 Then
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjMailDoc.SendTo. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                mstrSendToArray.CopyTo(mobjMailDoc.SendTo, 0)
            Else
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjMailDoc.SendTo. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                mobjMailDoc.SendTo = mstrSendTo
            End If
        End If

        'Ensure the Subject line has been set.
        If mstrSubject = "" Then
            'SendMail = False
            Err.Raise(ERR_NO_SUBJECT, APP_NAME, ERR_SUBJECT)
            Exit Function
        End If

        'Save the message in the users sent folder
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjMailDoc.SAVEMESSAGEONSEND. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        mobjMailDoc.SAVEMESSAGEONSEND = True

        'Send the email
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mobjMailDoc.SEND. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        '/* MUL INI si estamos en pruebas no se envia correo
        Dim parametroApp As String

        parametroApp = Microsoft.VisualBasic.Command
        If parametroApp = "PM" Then
            Exit Function
        End If
        '/* MUL FIN
        Call mobjMailDoc.SEND(False)
        'SendMail = True

        Exit Function

ErrorHandler:
        'SendMail = False
        Err.Raise(Err.Number, APP_NAME, Err.Description)

    End Function

    ' Función que calcula un 10% de descuento especial só aplicable a Reparalia
    '
    Private Function DescuentoReparalia(ByRef Total As Object) As Boolean

        On Error GoTo DescuentoReparalia_Err

        Dim tmpDto As Double
        Dim tmpTotal As Double

        If Total <= 0 Then Err.Raise(1)

        tmpDto = System.Math.Round((Total * 10) / 100, 2)
        tmpTotal = System.Math.Round(Total - tmpDto, 2)

        dbRPDtoTotal = tmpTotal

        DescuentoReparalia = True
        Exit Function

DescuentoReparalia_Err:
        dbRPDtoTotal = 0
        DescuentoReparalia = False
    End Function
End Class