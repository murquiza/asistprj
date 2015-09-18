Friend Class clsCierres_NET

    Private mvarReferencias As Collection ' Lista de pagos a procesar
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

    ' Esta función determina si un siniestro se puede cerra en base a los
    ' criterios previamente especificados y que se detallan a continuación
    '
    Public Function Bloqueo(ByRef objlistitem As ListViewItem) As String

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
        Dim Numperitos As String
        Dim UltimoPago As String
        Dim ProvisionTotal As Double
        Dim ProvisionFecha As Double

        Dim hora As DateTime

        hora = CDate("01/01/2001 23:59:59")

        On Error GoTo Bloqueo_Err

        Bloqueo = "Cierre"
        gstrError = ""

        ' Obtener Siniestro y Poliza
        '
        strPoliza = IIf(objlistitem.SubItems.Item(frmInstCierres.FECGRA.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstCierres.FECGRA.Index).Text), "0")
        strRamo = IIf(objlistitem.SubItems.Item(frmInstCierres.T2_FPAGO.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstCierres.T2_FPAGO.Index).Text), "0")
        strNexp = IIf(objlistitem.Text <> "", Trim(objlistitem.Text), "0")
        strRefer = IIf(objlistitem.SubItems.Item(frmInstCierres.T2_REFER.Index).Text <> "", Trim(objlistitem.SubItems.Item(frmInstCierres.T2_REFER.Index).Text), "0")

        '/*MUL T-19908 INI
        'If strCodcia = "I" Then
        '    strRefer = Mid(strRefer, 5, Len(strRefer) - 4)
        'ElseIf strCodcia = "R" Then
        '    strRefer = Mid(strRefer, 3, Len(strRefer) - 2)
        'End If
        Select Case strCodcia
            Case "I", "E", "M"
                strRefer = Mid(strRefer, 5, Len(strRefer) - 4)
            Case "R"
                strRefer = Mid(strRefer, 3, Len(strRefer) - 2)
            Case Else
                strRefer = ""
        End Select
        '/*MUL T-19908 FIN

        strPago = Trim(objlistitem.SubItems.Item(frmInstCierres.T2_ESTADO.Index).Text)
        dteFechaPago = CDate(Now)
        PrimaPendiente = "N"

        ' 1er Filtro
        ' --------------------------------------------------------------------
        '  Si hay otro périto ( ademas de la cia.de asistencia )
        ' --------------------------------------------------------------------

        ' Si la desactivación de Avisos esta seleccionada no ejecutamos
        ' el proceso de Avisos

        If Not frmInstCierres.chkFiltroAvisos.CheckState Then

            ' Si el objeto recordset está abierto lo cerramos
            '
            If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

            strsql = "Select count(*) Numperitos From Snsinges Where Snsinges.codsin = '" & strNexp & "' and " & "Snsinges.numper <> '" & strNumCompa & "'"
            claseBDCierres.BDWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If claseBDCierres.BDWorkRecord.EOF Then
                If claseBDCierres.BDWorkRecord.Fields("Numperitos").Value > 0 Then
                    Numperitos = claseBDCierres.BDWorkRecord.Fields("Numperitos").Value
                    gstrError = "* El Siniestro " & strNexp & " tiene, ademas de la Cia. de Asistencia, " & Numperitos & " peritos"
                    Call InsertarError("AC001", strRefer, strNexp, "A", gstrError, "C", , strRamo, strPoliza, strNexp)
                    Call ActualizarEstado("A", strRefer)
                    If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
                End If
            End If
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        ' 1er. Bloqueo
        ' --------------------------------------------------------------------
        '  Si hay pagos posteriores a la fecha de cierre indicada
        ' --------------------------------------------------------------------
        'strsql = "Select Count(*) Numpagos From Snsincta Where Snsincta.Codsin = '" & strNexp & "' and " & "Snsincta.Fecpag > '" & claseUtilidadesCierres.FormatoFechaSQL(CDate(FechaCierre), False, False) & "'"
        strsql = "Select Count(*) Numpagos From Snsincta Where Snsincta.Codsin = '" & strNexp & "' and " & "Snsincta.Fecpag > '" & claseUtilidadesCierres.FormatoFechaSQL(FechaCierre, False, False) & "'"
        claseBDCierres.BDWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If claseBDCierres.BDWorkRecord.EOF Then
            If claseBDCierres.BDWorkRecord.Fields("Numpagos").Value > 0 Then
                gstrError = gstrError & "* Existen pagos con fecha posterior a la fecha de cierre especificada. " & Chr(10) & Chr(10)
                Bloqueo = "Bloqueado"
            End If
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        ' 2do. Bloqueo
        ' --------------------------------------------------------------------
        '  Si no es el último pago
        ' --------------------------------------------------------------------

        strsql = "Select Count(*) UltimoPago From Angel_T2 Where Angel_T2.T2_Refer = '" & strRefer & "' and " & "Angel_T2.T2_Ultpag = 1 and Angel_T2.T2_Estado = 'P'"
        claseBDCierres.BDWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If claseBDCierres.BDWorkRecord.Fields("UltimoPago").Value = 0 Then
            gstrError = gstrError & "* No se ha realizado el último pago." & Chr(10) & Chr(10)
            Bloqueo = "Bloqueado"
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        ' 3er.. Bloqueo
        ' --------------------------------------------------------------------
        '  Si la Provisión Pendiente Absoluta es superior a la Provisión
        '  final a la fecha de cierre indicada
        ' --------------------------------------------------------------------

        ProvisionTotal = claseSiniestroCierres.ProvisionPendiente("", strNexp)
        'If strNexp = "41413909" Then Stop
        'ProvisionFecha = claseSiniestroCierres.ProvisionPendiente(claseUtilidadesCierres.FormatoFechaSQL(FechaCierre, False, False) & " 23:59:59", strNexp)
        'ProvisionFecha = claseSiniestroCierres.ProvisionPendiente(claseUtilidadesCierres.FormatoFechaSQL(FechaCierre.ToString & " 23:59:59", False, False), strNexp)
        ProvisionFecha = claseSiniestroCierres.ProvisionPendiente(FechaCierre, strNexp)

        If ProvisionTotal > ProvisionFecha Then
            gstrError = gstrError & "* Existen movimientos de provisión después de la fecha de cierre. " & Chr(10) & Chr(10)
            Bloqueo = "Bloqueado"
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        ' 4to. Bloqueo
        ' --------------------------------------------------------------------
        '  Si el siniestro tiene provisiones de recobro
        ' --------------------------------------------------------------------

        strsql = "SELECT Count(*) NumProvrecobros From Snprovis Where Codsin = '" & strNexp & "' and Tipprv = 'R'"
        claseBDCierres.BDWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not claseBDCierres.BDWorkRecord.EOF Then
            If claseBDCierres.BDWorkRecord.Fields("NumProvrecobros").Value > 0 Then
                gstrError = gstrError & "* El siniestro " & strNexp & " tiene provisiones de recobro. " & Chr(10) & Chr(10)
                Bloqueo = "Bloqueado"
            End If
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        ' 5to. Bloqueo
        ' --------------------------------------------------------------------
        '  Si existen recobros en el Siniestro
        ' --------------------------------------------------------------------
        strsql = "Select Count(*) Numrecobros From Snsincta Where Codsin = '" & strNexp & "' and Tipgas in ('C', 'D')"
        claseBDCierres.BDWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not claseBDCierres.BDWorkRecord.EOF Then
            If claseBDCierres.BDWorkRecord.Fields("Numrecobros").Value > 0 Then
                gstrError = gstrError & "* El siniestro " & strNexp & " tiene recobros. " & Chr(10) & Chr(10)
                Bloqueo = "Bloqueado"
            End If
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        ' 6to. Bloqueo
        ' --------------------------------------------------------------------
        '  Si existen pagos pendientes en el siniestro
        ' --------------------------------------------------------------------
        strsql = "Select Count(*) NumpagosPen From Snsincta Where Snsincta.Codsin = '" & strNexp & "' and " & "Snsincta.Pagado = 'N'"
        claseBDCierres.BDWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not claseBDCierres.BDWorkRecord.EOF Then
            If claseBDCierres.BDWorkRecord.Fields("NumpagosPen").Value > 0 Then
                gstrError = gstrError & "* Existen pagos pendientes. " & Chr(10) & Chr(10)
                Bloqueo = "Bloqueado"
            End If
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        ' 7mo. Bloqueo
        ' --------------------------------------------------------------------
        '  Si existen gestiones abiertas que no sean la de la Cia.de Asistencia
        ' --------------------------------------------------------------------
        strsql = "Select Count(*) Numges From Snsinges Where Snsinges.Codsin = '" & strNexp & "' and " & "(Snsinges.Tipges <> 'RE' and Snsinges.Numper <> '" & strNumCompa & "' and Snsinges.Fecrec is null)"
        claseBDCierres.BDWorkRecord.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not claseBDCierres.BDWorkRecord.EOF Then
            If claseBDCierres.BDWorkRecord.Fields("Numges").Value > 0 Then
                gstrError = gstrError & "* Existen " & claseBDCierres.BDWorkRecord.Fields("Numges").Value & " gestiones abiertas. " & Chr(10) & Chr(10)
                Bloqueo = "Bloqueado"
            End If
        End If

        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        Exit Function

Bloqueo_Err:
        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
        Bloqueo = "Error"
    End Function

    ' Procedure:  Pagar
    ' Objetivo:   Realiza la apertura del siniestro, inserta registro en las tablas de
    '             siniestros.
    ' Parametros: objListItem = Objeto ListItem de un ListView, con datos apertura.
    '             Retorno = Verdadero o Falso si pasa o no pasa el filtro
    '
    Public Function Cerrar(ByRef objlistitem As ListViewItem, Optional ByRef Origen As Object = Nothing) As Boolean
        On Error GoTo Cerrar_Err

        ' Declaraciones
        '
        Dim strRamo As String
        Dim strPoliza As String
        Dim strPago As String
        Dim strRefer As String
        Dim Codsin As String
        Dim strsql As String
        Dim Iva As Short
        Dim strCoderr As String

        '/* MUL INI
        boolTransaccion = False
        strError = ""
        TipoErr = ""
        '/* MUL FIN
        Cerrar = True

        ' Obtener Póliza, Fechas, Referencia, Ramo...
        '
        'strPoliza = Trim(objlistitem.SubItems.Item("POLIZA").Text)
        strPago = Trim(objlistitem.SubItems.Item(frmInstCierres.T2_ESTADO.Index).Text)
        strRefer = Trim(objlistitem.SubItems.Item(frmInstCierres.T2_REFER.Index).Text)
        Codsin = Trim(objlistitem.Text)

        ' Inicio de la Transacción
        '
        claseBDCierres.BDWorkConnect.BeginTrans()
        boolTransaccion = True

        ' Cierre de las Provisiones
        '
        If Not CierreSnProvis(objlistitem, Codsin) Then
            If gstrError = "" Then
                strCoderr = "CE001"
                strError = "4046"
                TipoErr = "E"
                Err.Raise(1)
            Else
                strCoderr = "PRV01"
                strError = CStr(4073)
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If

        ' Cierre de las Gestiones
        '
        If Not CierreSnsinges(objlistitem, Codsin) Then
            strCoderr = "CE002"
            strError = "4047"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Cierre del registro de cabecera del siniestro
        '
        If Not CierreSnsinies(objlistitem, Codsin) Then
            strCoderr = "CE003"
            strError = "4048"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Cierre de las anotaciones de agenda
        '
        If Not CierreAgenda(objlistitem, Codsin) Then
            strCoderr = "CE003"
            strError = "4048"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Actualiza el hisórico de estados del siniestro
        '
        If Not ActualizarHistoricoEstados(Codsin) Then
            strCoderr = "CE004"
            strError = "4049"
            TipoErr = "E"
            Err.Raise(1)
        End If

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Origen. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If Origen = "Anulaciones" Then
            If Not ActualizarObservaciones(Codsin) Then
                strCoderr = "C008"
                strError = "4051"
                TipoErr = "E"
                Err.Raise(1)
            End If
        End If

        ' Actualiza el estado del siniestro en la tabal de Aperturas ( Angel_t1 )
        '
        If Not ActualizarEstado("P", (objlistitem.Text)) Then
            strCoderr = "CE007"
            strError = "4045"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Final de la Transacción
        '
        claseBDCierres.BDWorkConnect.CommitTrans()

        colSiniestrosCerrados.Add(Codsin, Codsin)
        boolTransaccion = False

        If claseBDCierres.BDWorkRecord.State = 1 Then claseBDCierres.BDWorkRecord.Close()
        Exit Function

Cerrar_Err:
        Cerrar = False
        If boolTransaccion Then
            claseBDCierres.BDWorkConnect.RollbackTrans()
        End If
        If TipoErr = "E" Then
            Call ColorListItem(objlistitem, Color.Red)
        Else
            Call ColorListItem(objlistitem, Color.DarkGoldenrod)
        End If
        objlistitem.SubItems.Item(frmInstCierres.T2_ESTADO.Index).Text = TipoErr
        If Not InsertarError(strCoderr, (objlistitem.Text), Codsin, TipoErr, strError, strIdProceso) Then
            strError = "Se ha producido un error crítico en el registro de la tabla de errores y avisos" & Chr(13) & Chr(10) & "El proceso de cierres no puede continuar."
            Asunto = "Proceso Cierres Asistencia: Error Critico No Recuperable"
            Mensaje = "Se ha producido un error critico no recuperable que ha impedido que finalice el procesos de cierres. Consulte el archivo adjunto"
            MsgBox(Mensaje, MsgBoxStyle.Exclamation)
        End If
        If strError <> "4041" Then
            Call ActualizarEstado(TipoErr, strRefer)
        End If
        If claseBDCierres.BDWorkRecord.State = 1 Then claseBDCierres.BDWorkRecord.Close()
    End Function

    ' Procedure:  InsertarError
    ' Objetivo:   Inserta registro en la tabla MPASIHIST.
    ' Parametros: Referencia = Referencia del siniestros
    '             CodSin = Codigo de siniestro
    '             Error = Tipo de Error A/E
    '             Texto = Texto del error (descripcion)
    '             Proceso = Tipo proceso Aperturas/Pagos/...
    '
    Private Function InsertarError(ByRef CodError As String, ByRef Referencia As String, ByRef Codsin As String, ByRef strError As String, ByRef Texto As String, ByRef Proceso As String, Optional ByRef strObjeto As Object = Nothing, Optional ByRef strRamo As Object = Nothing, Optional ByRef strPoliza As Object = Nothing, Optional ByRef strSiniestro As Object = Nothing) As Boolean

        On Error GoTo InsertarError_Err

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim lngNumero As Integer ' Número de Error
        Dim lngCero As Integer ' Variable long para ontener el código de ramo

        ' Valores iniciales
        '
        If Not IsNothing(strRamo) Then
            lngCero = Val(strRamo)
        End If

        If IsNothing(strError) Then
            strError = ""
        End If

        If IsNothing(Codsin) Then
            Codsin = ""
        End If

        If IsNothing(Texto) Then
            Texto = ""
        End If

        ' Si el recordset está abierto lo cerramos
        '
        If claseBDCierres.BDWorkRecord.State = 1 Then claseBDCierres.BDWorkRecord.Close()

        ' Obtener el numero maximo de errores de una referencia
        '
        strsql = "SELECT IsNull(Max(numero),0) AS NUMERO " & "FROM   mpAsiHistError " & "WHERE  referencia = '" & Referencia & "' and proceso = '" & Proceso & "'"

        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDCierres.BDWorkRecord.EOF Then
            lngNumero = 0
        Else
            lngNumero = claseBDCierres.BDWorkRecord.Fields("Numero").Value
        End If
        claseBDCierres.BDWorkRecord.Close()

        lngNumero = lngNumero + 1

        With claseBDCierres.BDWorkRecord
            .Open("mpAsiHistError", claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            .AddNew()
            .Fields("Referencia").Value = Referencia
            .Fields("Codsin").Value = Codsin
            .Fields("Numero").Value = lngNumero
            .Fields("Errores").Value = strError
            .Fields("Texto").Value = Texto
            .Fields("Fecgra").Value = Today
            .Fields("Proceso").Value = Proceso
            .Fields("Cia").Value = strNombreCompa
            .Fields("TipoObjetoRel").Value = IIf(IsNothing(strObjeto), "", strObjeto)
            .Fields("RamoRel").Value = lngCero
            .Fields("PolizaRel").Value = IIf(IsNothing(strPoliza), "", strPoliza)
            .Fields("SiniestroRel").Value = IIf(IsNothing(strSiniestro), "", strSiniestro)
            .Fields("CodErr").Value = IIf(CodError = "", "0", CodError)
            .Fields("Codcia").Value = strCodcia
            .Update()
            .Close()
        End With
        InsertarError = True
        Exit Function

InsertarError_Err:
        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
        InsertarError = False
    End Function

    Public Sub New()
        MyBase.New()
        mvarReferencias = New Collection
    End Sub

    ' Esta Función actualiza el estado de la referencia en la tabla de aperturas
    '
    Private Function ActualizarEstado(ByRef TipoErr As String, ByRef refererr As String) As Boolean

        On Error GoTo ActualizarEstado_Err

        ' Declaraciones
        '
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim strsql As String ' Instrucción Sql
        Dim lngResult As Integer

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strsql = "Update Angel_T1 Set EstadoProcCierre = '" & TipoErr & "'," & "FechaProcCierre = '" & Format(CDate(Now), "MM/dd/yyyy") & "' " & " Where T1_codsin = '" & refererr & "'"

        With claseBDCierres.BDWorkConnect
            If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
            .Execute(strsql)
        End With

        ActualizarEstado = True
        Exit Function

ActualizarEstado_Err:
        ActualizarEstado = False
    End Function


    ' Esta función cierra la tabla de provisiones para el cierre del siniestro
    '
    Private Function CierreSnProvis(ByRef objlistitem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo CierreSnProvis_Err

        Dim CierreProvision As Double
        Dim strsql As String

        ' Obtenemos el importe para el asiento de provisión de cierre
        '
        CierreProvision = claseSiniestroCierres.ProvisionDisponible("", Codsin)

        ' CCheca. 03/04/2009. T-8009 Siniestros cerrados con provisión pendiente.
        '
        If CierreProvision = -1 Then Err.Raise(1)

        ' JLL - 19/04/2010 Si el importe con el que se va a crear el movimiento
        '                  de ajuste de provisión es 0, se marca como aviso. Si
        '                  la marca de desactivación de avisos está establecida
        '                  no se graba el asiento de ajuste de provisión.                  más.

        If CierreProvision = 0 Then
            If frmInstCierres.chkFiltroAvisos.CheckState Then
                CierreSnProvis = True
                Exit Function
            Else
                CierreSnProvis = False
                gstrError = "Aviso"
                Err.Raise(1)
            End If
        End If

        ' Fin JLL - 19/04/2010

        ' Cambiamos el signo
        '
        CierreProvision = CierreProvision / -1

        ' El objeto RecordSet debe estar cerrado antes de ejecutar un operacion 'Open'
        '
        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()

        strsql = "Select * from SnProvis Where Codsin = '" & Codsin & "'"

        ' Abrimos la tabla en la que vamos a grabar
        '
        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)


        With claseBDCierres.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = Codsin
            .Fields("Numprv").Value = claseSiniestroCierres.NumeroMovimientoTabla("Numprv", "Snprovis", "Codsin", Codsin)
            .Fields("Impprv").Value = CierreProvision
            'MUL ini MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            '.Fields("Fecprv").Value = claseUtilidadesCierres.FormatoFechaSQL(FechaCierre, False, True)
            .Fields("Fecprv").Value = FechaCierre.ToString("dd/MM/yyyy")
            'MUL fin MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            .Fields("Tipprv").Value = "P"
            .Fields("Comprv").Value = ""
            .Fields("Motpro").Value = "AT"
            'MUL ini MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            '.Fields("Fecmot").Value = claseUtilidadesCierres.FormatoFechaSQL(FechaCierre, False, True)
            .Fields("Fecmot").Value = FechaCierre.ToString("dd/MM/yyyy")
            'MUL fin MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            .Fields("IMPORT").Value = CierreProvision
            .Fields("Gasnpr").Value = 0
            .Fields("Usuprov").Value = strCodUserApli
            .Update()
            .Close()
        End With
        CierreSnProvis = True
        Exit Function

CierreSnProvis_Err:
        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
        strError = "Se ha producido un error al grabar el ajuste técnico para el cierre de las provisiones"
        gstrError = ""
        CierreSnProvis = False
    End Function


    ' Esta función cierra las gestiones abiertas para el cierre del siniestro
    '
    Private Function CierreSnsinges(ByRef objlistitem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo CierreSnsinges_Err

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim rsLocal As ADODB.Recordset ' Objeto RecordSet a BD

        ' Creación de objetos
        '
        rsLocal = New ADODB.Recordset

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strsql = "Select * From Snsinges Where Snsinges.Codsin ='" & Codsin & "' and " & "Snsinges.Tipges = 'RE' and Snsinges.Numper = '" & strNumcompa & "'"

        rsLocal.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        rsLocal.MoveFirst()
        Do While Not rsLocal.EOF
            With rsLocal
                'MUL ini MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
                '.Fields("Fecrec").Value = claseUtilidadesCierres.FormatoFechaSQL(FechaCierre, False, True)
                .Fields("Fecrec").Value = FechaCierre.ToString("dd/MM/yyyy")
                'MUL fin MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
                .Update()
                .MoveNext()
            End With
        Loop
        rsLocal.Close()

        CierreSnsinges = True
        Exit Function

CierreSnsinges_Err:
        If claseBDCierres.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
        strError = "Se ha producido un error al cerrar las gestiones pendientes"
        CierreSnsinges = False
    End Function

    ' Esta función actualiza los datos de cabecera para el cierre del siniestro
    '
    Private Function CierreSnsinies(ByRef objlistitem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo CierreSnsinies_Err

        ' Declaraciones
        '
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim strsql As String ' Instrucción Sql
        Dim lngResult As Integer
        Dim rsLocal As ADODB.Recordset ' Objeto RecordSet BD

        ' Creación de objetos
        '
        rsLocal = New ADODB.Recordset

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strsql = "Select * From Snsinies Where Codsin = '" & Codsin & "'"

        rsLocal.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        With rsLocal
            'MUL ini MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            '.Fields("Feccie").Value = claseUtilidadesCierres.FormatoFechaSQL(FechaCierre, False, True)
            .Fields("Feccie").Value = FechaCierre.ToString("dd/MM/yyyy")
            'MUL fin MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            .Fields("Estado").Value = "C"
            .Fields("Propen_P").Value = 0
            .Fields("Propen_R").Value = 0
            .Fields("Pte_Pago").Value = 0
            .Update()
        End With
        rsLocal.Close()

        CierreSnsinies = True
        Exit Function

CierreSnsinies_Err:

        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
        strError = "Se ha producido un error en la actualización del estado del siniestro"
        CierreSnsinies = False
    End Function

    Private Function CierreAgenda(ByRef objlistitem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo CierreAgenda_Err

        ' Declaraciones
        '
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim strsql As String ' Instrucción Sql
        Dim lngResult As Integer
        Dim rsLocal As ADODB.Recordset ' Objeto RecordSet BD
        Dim PerfilUs As String
        Dim AreaUs As String

        ' Creación de objetos
        '
        rsLocal = New ADODB.Recordset

        'PerfilUs = PerfilUsuario(CodUserApli)
        'AreaUs = AreaPerfil(PerfilUs)

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strsql = "Select * From Agenda, Empleado, Perfiles Where Agenda.ClaveCadena = '" & Codsin & "'" & " and Agenda.UsuarioReceptor = Empleado.Clave and Empleado.Perfil = Perfiles.Perfil " & " and Perfiles.Area = 'Asistencia' and CodigoAnotacion not in ('OBSER', 'ENPER', 'ENMIL')"

        rsLocal.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If rsLocal.RecordCount > 0 Then
            rsLocal.MoveFirst()
            Do While Not rsLocal.EOF
                With rsLocal
                    .Fields("Estado").Value = "ANU01"
                    .Update()
                    .MoveNext()
                End With
            Loop
        End If
        rsLocal.Close()

        CierreAgenda = True
        Exit Function

CierreAgenda_Err:

        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDCierres.BDWorkRecord.Close()
        strError = "Se ha producido un error en la actualización del estado del siniestro"
        CierreAgenda = False
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
        rsIva.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not rsIva.EOF Then
            IvaAplicable = rsIva.Fields("Iva").Value
        End If
        rsIva.Close()
        Exit Function

IvaAplicable_Err:
        rsIva.Close()
        IvaAplicable = 0
    End Function

    ' Esta función actualiza el histórico de estados de siniestro
    '
    Private Function ActualizarHistoricoEstados(ByRef sCodsin As String) As Boolean

        On Error GoTo ActualizarHistoricoEstados_Err

        Dim strsql As String

        If claseBDCierres.BDWorkRecord.State = 1 Then claseBDCierres.BDWorkRecord.Close()

        strsql = "Select * From Sn_EstadoHist Where Codsin = '" & sCodsin & "'"

        claseBDCierres.BDWorkRecord.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDCierres.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = sCodsin
            .Fields("Nummov").Value = claseSiniestroCierres.NumeroMovimientoTabla("Nummov", "Sn_Estadohist", "Codsin", sCodsin)
            .Fields("Estado").Value = "C"
            'MUL ini MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            '.Fields("Fecest").Value = claseUtilidadesCierres.FormatoFechaSQL(FechaCierre, False, True)
            .Fields("Fecest").Value = FechaCierre.ToString("dd/MM/yyyy")
            'MUL fin MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
            .Fields("Usuari").Value = strCodUserApli
            .Fields("Fecgra").Value = CDate(Now)
            .Update()
        End With
        claseBDCierres.BDWorkRecord.Close()

        ActualizarHistoricoEstados = True
        Exit Function

ActualizarHistoricoEstados_Err:
        ActualizarHistoricoEstados = False
        If claseBDCierres.BDWorkRecord.State = 1 Then claseBDCierres.BDWorkRecord.Close()
    End Function

    ' Esta función graba en el campo de observaciones de la agenda la
    ' descripción y el comentario de la anulación
    '
    Private Function ActualizarObservaciones(ByRef sCodsin As String) As Boolean

        On Error GoTo ActualizarObservaciones_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim sql1 As String

        sql1 = "Select T5_Descripcion, T5_Comentarios " & "From   AnulacionesAsistencia " & "Where  T5_Codsin = '" & sCodsin & "'"

        rsLocal.Open(sql1, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsLocal.EOF Then
            If claseBDCierres.BDWorkRecord.State = 1 Then claseBDCierres.BDWorkRecord.Close()
            claseBDCierres.BDWorkRecord.Open("Snagenda", claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            With claseBDCierres.BDWorkRecord
                .AddNew()
                .Fields("Codsin").Value = sCodsin
                .Fields("Fecha").Value = CDate(Now)
                .Fields("Descri").Value = "Descripción Rechazo Asistencia: " & rsLocal.Fields("T5_Descripcion").Value
                .Fields("Fecest").Value = CDate(Now)
                .Fields("Tipo").Value = "O"
                .Fields("Observ").Value = rsLocal.Fields("T5_Comentarios").Value
                .Fields("Usuari").Value = strCodUserApli
                .Update()
            End With
            claseBDCierres.BDWorkRecord.Close()
            ActualizarObservaciones = True
        End If
        If rsLocal.State = 1 Then rsLocal.Close()

        Exit Function

ActualizarObservaciones_Err:
        ActualizarObservaciones = False
        If claseBDCierres.BDWorkRecord.State = 1 Then claseBDCierres.BDWorkRecord.Close()
    End Function

    ' Esta función graba en el campo de observaciones de la agenda la
    ' descripción y el comentario de la anulación
    '
    Private Function ActualizarEstadoAnulaciones(ByRef sCodsin As String) As Boolean

        On Error GoTo ActualizarEstadoAnulaciones_Err

        ' Declaraciones
        '
        Dim strsql As String

        strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Estado = 'P' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codcia = '" & strCodcia & "' and " & "       AnulacionesAsistencia.T5_Codsin = '" & sCodsin & "'"

        claseBDCierres.BDWorkConnect.Execute(strsql)

        Exit Function

ActualizarEstadoAnulaciones_Err:
        ActualizarEstadoAnulaciones = False
    End Function

    ' Esta función inserta un registro en el fichero log por cada una de las
    ' referencias de siniestro que se hayan cerrado en el proceso
    '
    Public Function InsertarLog() As Boolean

        Dim Canal As Short ' Número de canal de apertura
        Dim i As Short ' contador para bucle del número de siniestro grabado

        On Error GoTo InsertarLog_Err

        ' Abrimos el fichero log
        '
        Canal = AbreFichero(NombreFichero)

        ' Si la creación del fichero se ejecutado correctamente, entramos
        ' en un bucle que grabará todas las referecias de siniestros que
        ' se han cerrado
        '
        If Canal > -1 Then
            For i = 1 To colSiniestrosCerrados.Count()
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto colSiniestrosCerrados.Item(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                If Not GrabaFichero(colSiniestrosCerrados.Item(i), Canal) Then
                    Err.Raise(1)
                End If
            Next i
        End If
        If Not CierraFichero(Canal) Then
            Err.Raise(1)
        End If

        InsertarLog = True
        FileClose(Canal)
        Exit Function

InsertarLog_Err:
        InsertarLog = False
        FileClose(Canal)
        Kill(NombreFichero)
    End Function

    ' Esta función abre un fichero secuencial y devuelve el número de canal
    ' por el que ha sido abierto
    '
    Private Function AbreFichero(ByVal Fichero As String) As Short

        On Error GoTo Abrefichero_Err

        ' Declaraciones
        '
        Dim Canal As Short

        ' Obtenemos el numero de canal por el que abriremos el fichero
        '
        Canal = FreeFile()

        ' Apertura del fichero
        '
        FileOpen(Canal, Fichero, OpenMode.Append, OpenAccess.Write)

        AbreFichero = Canal

        Exit Function

Abrefichero_Err:
        AbreFichero = -1
    End Function

    ' Esta función graba una linea de datos, pasada en el parámetro Vectordatos,
    ' en el fichero abierto en el número de canal pasado en el parametro Canal
    '
    Private Function GrabaFichero(ByRef VectorDatos As String, ByRef Canal As Short) As Boolean

        On Error GoTo GrabaFichero_Err

        ' Grabamos en el fichero especificado el vector de datos
        '
        PrintLine(Canal, VectorDatos)
        GrabaFichero = True

        Exit Function

GrabaFichero_Err:
        GrabaFichero = False
    End Function

    ' Esta función cierra el fichero abierto por el canal especificado
    '
    Private Function CierraFichero(ByRef Canal As Short) As Boolean

        On Error GoTo CierraFichero_Err

        FileClose(Canal)
        CierraFichero = True

        Exit Function

CierraFichero_Err:
        CierraFichero = False
    End Function

    ' Esta función crea el fichero de log y graba el registro de cabecera
    '
    Public Function CabeceraLog() As Boolean

        Dim Canal As Short ' Número de canal de apertura
        Dim i As Short ' contador para bucle del número de siniestro grabado
        Dim Cadena As String ' Texto a grabar en la linea de cabecera

        On Error GoTo CabeceraLog_Err

        ' Si existe lo borramos
        '
        Kill(NombreFichero)

        ' Abrimos el fichero log
        '
        Canal = AbreFichero(NombreFichero)

        ' Si la creación del fichero se ejecutado correctamente, entramos
        ' en un bucle que grabará todas las referecias de siniestros que
        ' se han cerrado

        If Canal > -1 Then
            Cadena = "Inicio Proceso Cierres Asistencia" & Chr(13) & Chr(10)
            Cadena = Cadena & "Siniestros Seleccionados: " & Str(Referencias.Count()) & Chr(13) & Chr(10)
            Cadena = Cadena & "Fecha Ejecución: " & FechaEjecucion & Chr(13) & Chr(10)
            Cadena = Cadena & "Fecha Procesada: " & Format(frmInstCierres.dtpHasta.Value, "yyyy/MM/dd hh:mm") & Chr(13) & Chr(10)
            Cadena = Cadena & " " & Chr(13) & Chr(10) & ""
            If Not GrabaFichero(Cadena, Canal) Then
                Err.Raise(1)
            End If
        End If
        If Not CierraFichero(Canal) Then
            Err.Raise(1)
        End If

        CabeceraLog = True
        FileClose(Canal)
        Exit Function

CabeceraLog_Err:
        If Err.Number = 53 Then
            Resume Next
        End If
        CabeceraLog = False
        FileClose(Canal)
    End Function

    ' Esta función graba el registro de cierre del fichero log
    '
    Public Function PieLog() As Boolean

        Dim Canal As Short ' Número de canal de apertura
        Dim i As Short ' contador para bucle del número de siniestro grabado
        Dim Cadena As String ' Texto a grabar en la linea de cabecera
        Dim DifProcesados As Integer ' Siniestros que quedan pendientes de procesar

        On Error GoTo PieLog_Err

        DifProcesados = Referencias.Count() - colSiniestrosCerrados.Count()

        ' Abrimos el fichero log
        '
        Canal = AbreFichero(NombreFichero)

        ' Si la creación del fichero se ejecutado correctamente, entramos
        ' en un bucle que grabará todas las referecias de siniestros que
        ' se han cerrado

        If Canal > -1 Then
            Cadena = " " & Chr(13) & Chr(10) & ""
            Cadena = Cadena & "Fin Proceso Cierres Asistencia" & Chr(13) & Chr(10)
            Cadena = Cadena & "Siniestros Cerrados: " & Str(colSiniestrosCerrados.Count()) & Chr(13) & Chr(10)
            If DifProcesados > 0 Then
                Cadena = Cadena & "Siniestros Seleccionados Pendientes de Procesar: " & Str(DifProcesados) & Chr(13) & Chr(10)
            End If
            Cadena = Cadena & "Fecha Final: " & Format(Now, "yyyy/MM/dd hh:mm") & Chr(13) & Chr(10)
            Cadena = Cadena & "Fecha Procesada: " & Format(frmInstCierres.dtpDesde.Value, "yyyy/MM/dd hh:mm") & Chr(13) & Chr(10)
            If Not GrabaFichero(Cadena, Canal) Then
                Err.Raise(1)
            End If
        End If
        If Not CierraFichero(Canal) Then
            Err.Raise(1)
        End If

        PieLog = True
        FileClose(Canal)
        Exit Function

PieLog_Err:
        PieLog = False
        FileClose(Canal)
    End Function

    ' Esta función graba en la tabla mdpCierresPendientes, todas las referencias
    ' de siniestros que no se hayan podido cerrar durante el proceso porque
    ' se ha sobrepasado el límite de hora de ejecución
    '
    Public Function GrabaCierrePendiente() As Boolean

        On Error GoTo GrabaCierrePendiente_Err

        Dim i As Short
        Dim rsLocal As ADODB.Recordset
        Dim strsql As String

        ' Creación de objetos
        '
        rsLocal = New ADODB.Recordset

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strsql = "Select * From mdpCierresPendientes"
        rsLocal.Open(strsql, claseBDCierres.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If colSiniestrosCerrados.Count() < Me.Referencias.Count() Then
            For i = 1 To Me.Referencias.Count()
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Me.Referencias.Item(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                If Not BuscaRefCerrada(Me.Referencias.Item(i)) Then
                    With rsLocal
                        .AddNew()
                        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Me.Referencias.Item(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                        .Fields("Codsin").Value = Me.Referencias.Item(i)
                        'MUL ini MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
                        '.Fields("FechaPendiente").Value = claseUtilidadesCierres.FormatoFechaSQL(Now, False, True)
                        .Fields("FechaPendiente").Value = Now.ToString("dd/MM/yyyy")
                        'MUL fin MAN-481 fecha de cierre de siniestro es anterior a fecha de siniestro y pagos efectuados 
                        .Update()
                    End With
                End If
            Next
            rsLocal.Close()
        End If

        Exit Function

GrabaCierrePendiente_Err:
        GrabaCierrePendiente = False
    End Function

    ' Esta función busca la referencia de siniestro pasada en el parámetro
    ' REF en la colección de siniestros cerrados
    '
    Private Function BuscaRefCerrada(ByRef Ref As String) As Boolean

        On Error GoTo BuscaRefCerrada_Err

        Dim i As Short

        BuscaRefCerrada = False

        For i = 1 To colSiniestrosCerrados.Count()
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto colSiniestrosCerrados.Item(i). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            If Ref = colSiniestrosCerrados.Item(i) Then
                BuscaRefCerrada = True
                Exit For
            End If
        Next i
        Exit Function

BuscaRefCerrada_Err:
        BuscaRefCerrada = False
    End Function

    ' Esta función elimina de la tabla mdpSiniestrosPendientes todos
    ' los registros
    '
    Public Function BorraSiniestrosPendientes() As Boolean

        On Error GoTo BorraSiniestrosPendientes_Err

        ' Declaraciones
        '
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim strsql As String ' Instrucción Sql
        Dim lngResult As Integer

        If colSiniestrosCerrados.Count() < Me.Referencias.Count() Then

            ' Construimos la sentencia Sql para actualizar el estado
            '
            strsql = "Delete From mdpCierresPendientes"

            With claseBDCierres.BDWorkConnect
                If .State = ADODB.ObjectStateEnum.adStateClosed Then .Open()
                .Execute(strsql)
            End With

        End If
        BorraSiniestrosPendientes = True
        Exit Function

BorraSiniestrosPendientes_Err:
        BorraSiniestrosPendientes = False
    End Function
End Class
