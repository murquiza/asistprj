Friend Class clsAperturar_NET

    ' Variables locales para almacenar los valores de las propiedades
    '
    Private mvarReferencias As Collection ' Colección local
    Private mstrCia As String ' Código de la compañia de asistencia
    Private strError As String ' Mensaje de error a grabar en tabla de errores
    Private lProvision As Integer ' Contiene la provisión inicial para el siniestro
    Private sRecibo As String ' Contiene el número de recibo que da cobertura al siniestro
    Private strRiesgo As String ' Contiene el código del riesgo de la póliza del siniestro
    Private sMutualista As String ' Contiene el código del mutualista de la póliza del siniestro
    Private sAgente As String ' Contiene el código del agente de la póliza del siniestro
    Private TipoErr As String ' Tipo de error producido ( 'Aviso' o 'Error Severo')
    Private Codsin As String ' Código del siniestro

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

    ' Procedure:  Inicializar
    ' Objetivo:   Inicializa el objeto, y llama a la funcion de importacion que toca
    ' Parametros: Proceso = Tipo de proceso de Importación
    '
    Public Sub Inicializar(ByRef Cia As String)
        mstrCia = Cia
    End Sub

    Public Function Filtros(ByRef objListItem As ListViewItem) As Boolean

        On Error GoTo Filtros_Err

        ' Declaraciones
        '
        Dim strRamo As String
        Dim strPoliza As String
        Dim strApertura As String
        Dim dteFechaSiniestro As Date
        Dim objCmd As ADODB.Command
        Dim strSQL As String
        Dim lngSiniestros As Integer
        Dim lngNumEstadoRecibo As Integer
        Dim errTipoObjeto As String
        Dim errRamo As String
        Dim errPoliza As String
        Dim errSiniestro As String
        Dim strCoderr As String
        Dim Codsin As String

        Dim strSiniRefExt As String 'Ripolles T-3686 17/06/2008

        Filtros = True

        ' Obtener poliza, ramo, fecha siniestro y estado
        '
        strPoliza = Trim(objListItem.SubItems.Item(frmInstAperturas.T1_POLIZA.Index).Text)
        strRamo = CInt(Trim(objListItem.SubItems.Item(frmInstAperturas.T1_CIA.Index).Text))
        dteFechaSiniestro = CDate(objListItem.SubItems.Item(frmInstAperturas.T1_FSINI.Index).Text)
        strApertura = Trim(objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text)

        ' Comprobación de filtros

        ' 1er Filtro
        ' ----------------------------------------------------------------------------
        '   Este filtro comprueba:
        '           Que la Póliza exista
        '           Que la Póliza este en vigor en el momento del siniestro
        '           Que no este anulada
        ' ----------------------------------------------------------------------------
        '
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        strSQL = "Select Polizaca.FecPol, Polizaca.FecEfe, Polizaca.Polanu, Polizaca.Fecbaj " & "From   Polizaca " & "Where  Polizaca.NumPol = '" & strPoliza & "' and Polizaca.Codram = '" & strRamo & "'"

        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDAperturas.BDWorkRecord.EOF Then
            strCoderr = "E0001"
            strError = "La Póliza " & strPoliza & " no existe"
            TipoErr = "E"
            Err.Raise(1)
        Else

            If Not IsDBNull(claseBDAperturas.BDWorkRecord.Fields("FecPol").Value) Then

                If dteFechaSiniestro < claseBDAperturas.BDWorkRecord.Fields("FecPol").Value Then
                    strError = "La Póliza " & strPoliza & " está fuera de cobertura"
                    strCoderr = "E0003"
                    TipoErr = "E"
                    Err.Raise(1)
                ElseIf Not IsDBNull(claseBDAperturas.BDWorkRecord.Fields("Fecbaj").Value) Then
                    If claseBDAperturas.BDWorkRecord.Fields("Polanu").Value = "S" And dteFechaSiniestro >= CDate(claseBDAperturas.BDWorkRecord.Fields("Fecbaj").Value) Then
                        strError = "La Póliza " & strPoliza & " está anulada con efecto " & claseBDAperturas.BDWorkRecord.Fields("Fecbaj").Value
                        strCoderr = "E0002"
                        TipoErr = "E"
                        Err.Raise(1)
                    End If
                End If

            End If
        End If

        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()

        ' 2º Filtro
        ' ---------------------------------------------------------------------------------
        ' Comprobar poliza y fecha de siniestro con un margen de +/- 1 semana, si dichos
        ' datos coinciden con nuestra base de datos, no se apertura el siniestro.
        ' ---------------------------------------------------------------------------------
        '
        lngSiniestros = SiniestrosPoliza(strPoliza, strRamo, dteFechaSiniestro, Codsin)

        If lngSiniestros = 1 Then
            strError = "La Poliza " & strPoliza & " tiene siniestros con fecha de +/- 1 semana" & " El siniestro relacionado es el " & Codsin & " Descripción: " & claseBDAperturas.BDWorkRecord.Fields("Dessin").Value
            strCoderr = "A0001"
            TipoErr = "A"
            errTipoObjeto = "Siniestro"
            errRamo = ""
            errPoliza = ""
            errSiniestro = Codsin
            Codsin = ""
            Err.Raise(55)
        ElseIf lngSiniestros > 1 Then
            strError = "La Poliza " & strPoliza & " tiene " & lngSiniestros & " siniestros con fecha de +/- 1 semana"
            strCoderr = "A0001"
            TipoErr = "A"
            errTipoObjeto = "Siniestro"
            errRamo = ""
            errPoliza = ""
            errSiniestro = Codsin
            Codsin = ""
            Err.Raise(55)
        End If
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()

        ' 3er Filtro
        ' ----------------------------------------------------------------------------
        ' Si el recibo está pendiente desde hace > de 1 mes de la fecha del siniestro,
        ' se paraliza la entrada del parte para su posterior estudio.
        ' ----------------------------------------------------------------------------
        '
        '    If mdpbd.BDWorkRecord.State = adStateOpen Then mdpbd.BDWorkRecord.Close
        '    lngNumEstadoRecibo = 0
        '    strsql = "SELECT IsNull(Count(*),0) AS NumEstadoRecibo " & _
        ''                 "FROM Snsinies, Carterac " & _
        ''                 "WHERE Snsinies.Numrec = Carterac.Numrec and " & _
        ''                 "Snsinies.Numpol = '" & strPoliza & "' AND " & _
        ''                 "Snsinies.Codram = '" & strRamo & "' And " & _
        ''                 "Carterac.Fesuvt < '" & objUtiles.FormatoFechaSQL(DateAdd("m", -1, dteFechaSiniestro), False, False) & "' AND Carterac.Estado <> 'C'"
        '
        '    mdpbd.BDWorkRecord.Open strsql, mdpbd.BDWorkConnect, adOpenDynamic, adLockOptimistic
        '    If mdpbd.BDWorkRecord.EOF Then
        '        lngNumEstadoRecibo = 0
        '    Else
        '        lngNumEstadoRecibo = mdpbd.BDWorkRecord.Fields(0).Value
        '    End If
        '    mdpbd.BDWorkRecord.Close
        '
        '    If lngNumEstadoRecibo > 0 Then
        '        strCoderr = "A0002"
        '        Call InsertarError(strCoderr, objListItem.Text, objListItem.SubItems("T1_CODSIN").Text, "A", "La poliza tiene siniestros con recibos pendientes de pagos desde > 1 mes.", "A")
        '        Call ColorListItem(objListItem, &H1F8EC5)
        '        objListItem.SubItems.Item("T1_ESTADO").Text = "A"  '// Aviso
        '        Call ActualizarEstado(objListItem.SubItems.Item("T1_ESTADO").Text, objListItem.Text)
        '        Filtros = False
        '    End If

        ' 4º Filtro
        ' ----------------------------------------------------------------------------
        ' Si el Siniestro se apertura por orden pericial, generar Aviso
        ' ----------------------------------------------------------------------------
        '
        If Trim(objListItem.SubItems.Item(frmInstAperturas.T1_PERTURAPOR.Index).Text) = "S" Then
            strCoderr = "A004"
            Call InsertarError(strCoderr, (objListItem.Text), (objListItem.SubItems(frmInstAperturas.T1_ESTADO.Index).Text), "A", "El siniestro tiene Perito", "A")
            Call ColorListItem(objListItem, System.Drawing.ColorTranslator.FromOle(&H1F8EC5))
            objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text = "A" '// Aviso
            Call ActualizarEstado((objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text), (objListItem.Text))
            Filtros = False
        End If

        'Inicio Ripolles T-3686 17/06/2008

        ' 5º Filtro
        ' ----------------------------------------------------------------------------
        ' Validar que la referencia externa no exista ya en un siniestro
        ' ----------------------------------------------------------------------------
        '
        strSiniRefExt = objListItem.Text

        If ReferenciaExternaDuplicada(strSiniRefExt, Codsin) >= 1 Then
            strError = "La referencia externa " & strSiniRefExt & " ya existe en el siniestro " & Codsin
            strCoderr = "E0004"
            TipoErr = "E"
            errTipoObjeto = "Siniestro"
            errRamo = ""
            errPoliza = ""
            errSiniestro = Codsin
            Err.Raise(1)
        End If
        'If mdpbd.BDWorkRecord.State = adStateOpen Then mdpbd.BDWorkRecord.Close

        'Fin Ripolles T-3686

        Exit Function


Filtros_Err:
        Filtros = False
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        If TipoErr = "E" Then
            Call ColorListItem(objListItem, System.Drawing.Color.Red)
        Else
            Call ColorListItem(objListItem, System.Drawing.ColorTranslator.FromOle(&H1F8EC5))
        End If
        If Err.Number = 55 Then
            Call InsertarError(strCoderr, (objListItem.Text), (objListItem.Text), TipoErr, strError, "A", errTipoObjeto, errRamo, errPoliza, errSiniestro)
        Else
            Call InsertarError(strCoderr, (objListItem.Text), (objListItem.Text), TipoErr, strError, "A")
        End If
        Call ActualizarEstado(TipoErr, (objListItem.Text))
        objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text = TipoErr
    End Function

    ' Procedure:  Aperturar
    ' Objetivo:   Realiza la apertura del siniestro, inserta registro en las tablas de
    '             siniestros.
    ' Parametros: objListItem = Objeto ListItem de un ListView, con datos apertura.
    ' Retorno:    Booleano
    '
    Public Function Aperturar(ByRef objListItem As ListViewItem) As Boolean

        On Error GoTo Aperturar_Err

        ' Declaraciones
        '
        Dim strRamo As String
        Dim strPoliza As String
        Dim strApertura As String
        Dim strRefer As String
        Dim dteFechaSiniestro As Date
        Dim dteYear As String
        Dim strSQL As String
        Dim lngSiniestros As Integer
        Dim lngNumEstadoRecibo As Integer
        Dim strCoderr As String
        Dim errTipoObjeto As String
        Dim errRamo As String
        Dim errPoliza As String
        Dim errSiniestro As String
        'JCLopez_i
        Dim strTramitador As String
        'JCLopez_f

        Aperturar = True
        '/* MUL INI
        TipoErr = ""
        boolTransaccion = False
        '/* MUL FIN

        ' Obtener Póliza, Fechas, Referencia, Ramo...
        '
        strPoliza = Trim(objListItem.SubItems.Item(frmInstAperturas.T1_POLIZA.Index).Text)
        dteFechaSiniestro = CDate(objListItem.SubItems.Item(frmInstAperturas.T1_FSINI.Index).Text)
        dteYear = CStr(Year(dteFechaSiniestro))
        dteYear = Mid(dteYear, 3, 2)
        strApertura = Trim(objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text)
        strRefer = Trim(objListItem.Text)
        strRamo = Trim(objListItem.SubItems.Item(frmInstAperturas.T1_CIA.Index).Text)

        ' Debido a que la compañia de Asistencia nos  puede enviar n veces
        ' el mismo siniestro con diferentes  referencias en un mismo dia o
        ' fichero, nos vemos obligados a comprobar el filtro de duplicidad
        ' de +/- 1 semana antes de procesar la apertura de cada referencia

        ' Comprobamos que la comprobación de filtros esta activada
        '
        If Not frmInstAperturas.chkFiltroAvisos.Checked Then

            lngSiniestros = SiniestrosPoliza(strPoliza, strRamo, dteFechaSiniestro, Codsin)

            If lngSiniestros = 1 Then
                strError = "La Poliza " & strPoliza & " tiene siniestros con fecha de +/- 1 semana" & " El siniestro relacionado es el " & Codsin & " Descripción: " & claseBDAperturas.BDWorkRecord.Fields("Dessin").Value
                strCoderr = "A0001"
                TipoErr = "A"
                errTipoObjeto = "Siniestro"
                errRamo = ""
                errPoliza = ""
                errSiniestro = Codsin
                Codsin = ""
                Err.Raise(55)
            ElseIf lngSiniestros > 1 Then
                strError = "La Poliza " & strPoliza & " tiene " & lngSiniestros & " siniestros con fecha de +/- 1 semana"
                strCoderr = "A0001"
                TipoErr = "A"
                errTipoObjeto = "Siniestro"
                errRamo = ""
                errPoliza = ""
                errSiniestro = Codsin
                Codsin = ""
                Err.Raise(55)
            End If
            If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()
        End If

        ' Obtener Codigo del Riesgo
        '
        strRiesgo = RiesgoPoliza(strPoliza, strRamo)
        If strRiesgo = vbNullString Then
            If Not frmInstAperturas.chkFiltroAvisos.CheckState Then
                strError = "No se ha podido obtener el código de riesgo de la póliza"
                strCoderr = "A0007"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If

        ' Obtener Provisión Inicial
        '
        lProvision = claseSiniestroAperturas.ProvisionInicialSiniestro(strRamo, "", True)
        If lProvision = 0 Then
            strError = "No ha sido posible obtener provisión inicial para el siniestro"
            strCoderr = "E004"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Obtener Mutualista
        '
        sMutualista = MutualistaPoliza(strPoliza, strRamo)
        If sMutualista = vbNullString Then
            If Not frmInstAperturas.chkFiltroAvisos.CheckState Then
                strError = "No se ha podido obtener el mutualista de la póliza del siniestro"
                strCoderr = "A0004"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If

        ' Obtener Número de Recibo
        '
        sRecibo = claseSiniestroAperturas.ReciboSiniestro(strPoliza, strRamo, dteFechaSiniestro)
        If sRecibo = vbNullString Then
            If Not frmInstAperturas.chkFiltroAvisos.CheckState Then
                strError = "No ha sido posible obtener el recibo que da cobertura al siniestro"
                strCoderr = "A0003"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If

        ' Obtener Agente de la Póliza
        '
        sAgente = AgentePoliza(strPoliza, strRamo)
        If sAgente = vbNullString Then
            If Not frmInstAperturas.chkFiltroAvisos.CheckState Then
                strError = "No ha sido posible obtener el Agente de la póliza"
                strCoderr = "A0005"
                TipoErr = "A"
                Err.Raise(1)
            End If
        End If

        ' Obtener Número de Siniestro
        '
        Codsin = claseSiniestroAperturas.ObtenerNumeroSiniestro(Trim(objListItem.SubItems.Item(frmInstAperturas.T1_CIA.Index).Text))

        If Codsin = "" Then
            strError = "No se ha podido construir el código de expediente. Consulte al Departamento de Infomática"
            strCoderr = "E0005"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Inicio de la Transacción
        '
        claseBDAperturas.BDWorkConnect.BeginTrans()
        boolTransaccion = True

        ' Graba el registro de cabecera de siniestro ( Tabla SnSinies)
        '
        If Not ActualizarSnSinies(objListItem, Codsin) Then
            strCoderr = "S0001"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Graba la provisión inicial del siniestro ( Tabla SnProvis )
        '
        If Not ActualizarSnProvis(objListItem, Codsin) Then
            strCoderr = "S0002"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Graba el perjudicado del siniestro ( Tabla SnSinper )
        '
        If Not ActualizarSnSinper(objListItem, Codsin) Then
            strCoderr = "S0003"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Graba la gestión de reparación del siniestro ( Tabla SnSinges )
        '
        If Not ActualizarSnSinges(objListItem, Codsin) Then
            strCoderr = "S0004"
            TipoErr = "E"
            Err.Raise(1)
        End If

        If Not ActualizarHistoricoEstados(Codsin) Then
            strCoderr = "S0005"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Graba un recordatorio en la Agenda
        '
        '    If Not ActualizarSnAgenda(Codsin) Then
        '        strCoderr = "S0008"
        '        TipoErr = "E"
        '        Err.Raise 1
        '    End If

        ' Registra el nuevo número de siniestro en la tabla de numeradores ( Tabla_FicApl )
        '
        If Not ActualizarNumerador(Codsin) Then
            strCoderr = "S0006"
            TipoErr = "E"
            Err.Raise(1)
        End If

        ' Actualiza el estado del siniestro en la tabal de Aperturas ( Angel_t1 )
        '
        If Not ActualizarEstado("P", (objListItem.Text)) Then
            strCoderr = "S0007"
            strError = "4028"
            TipoErr = "E"
            Err.Raise(1)
        End If

        'JCLopez_i
        'Se coge el tramitador asignado para el siniestro en los siniestros del grupo de hogar
        If GetGrupoRamo(strRamo) = 6 Then
            strTramitador = SiguienteTramitadorGrupo("GENGHOG1")

            'Actualiza el campo Tramitador en siniatrib
            If Not ActualizarSiniatrib(Codsin, strTramitador) Then
                strCoderr = ""
                strError = "Ha ocurrido un error actualizando el tramitador en la tabla siniatrib"
                TipoErr = "E"
                Err.Raise(1)
            End If

        End If
        'JCLopez_f

        ' Llamada a nueva función para la tabla nueva de Personas por la migración a iAxis
        '
        Call ActualizarPersonas2(objListItem, Codsin)


        ' Final de la Transacción
        '
        claseBDAperturas.BDWorkConnect.CommitTrans()

        Exit Function

Aperturar_Err:
        Aperturar = False
        If boolTransaccion And Err.Number <> 55 Then
            claseBDAperturas.BDWorkConnect.RollbackTrans()
        End If
        If TipoErr = "E" Then
            Call ColorListItem(objListItem, System.Drawing.Color.Red)
        Else
            Call ColorListItem(objListItem, System.Drawing.ColorTranslator.FromOle(&H1F8EC5))
        End If
        objListItem.SubItems.Item(frmInstAperturas.T1_ESTADO.Index).Text = TipoErr
        If Not InsertarError(strCoderr, (objListItem.Text), Codsin, TipoErr, strError, strIdProceso) Then
            strError = "Se ha producido un error crítico en el registro de la tabla de errores y avisos" & Chr(13) & Chr(10) & "El proceso de aperturas no puede continuar."
            MsgBox(strError, MsgBoxStyle.Critical)
            'objError.Ver(IdProceso, , strError, Codcia)
        End If
        If strError <> "4028" Then
            Call ActualizarEstado(TipoErr, (objListItem.Text))
        End If
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()
    End Function

    ' Procedure:  InsertarError
    ' Objetivo:   Inserta registro en la tabla MPASIHIST.
    ' Parametros: Referencia = Referencia del siniestros
    '             CodSin = Codigo de siniestro
    '             Error = Tipo de Error A/E
    '             Texto = Texto del error (descripcion)
    '             Proceso = Tipo proceso Aperturas/Pagos/...
    Private Function InsertarError(ByRef CodError As String, ByRef Referencia As String, ByRef Codsin As String, ByRef strMensajeError As String, ByRef Texto As String, ByRef Proceso As String, Optional ByRef strObjeto As Object = Nothing, Optional ByRef strRamo As Object = Nothing, Optional ByRef strPoliza As Object = Nothing, Optional ByRef strSiniestro As Object = Nothing) As Boolean

        On Error GoTo InsertarError_Err

        ' Declaraciones
        '
        Dim strSQL As String ' Instrucción Sql
        Dim lngNumero As Integer ' Número de Error
        Dim lngCero As Integer ' Variable long para ontener el código de ramo

        ' Valores iniciales
        If Not IsNothing(strRamo) Then
            lngCero = Val(strRamo)
        End If

        If IsNothing(strMensajeError) Then
            strMensajeError = ""
        End If

        ' Si el recordset está abierto lo cerramos
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()

        ' Obtener el numero maximo de errores de una referencia
        strSQL = "SELECT IsNull(Max(numero),0) AS NUMERO " & "FROM   mpAsiHistError " & "WHERE  referencia = '" & Referencia & "' and proceso = '" & Proceso & "'"

        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDAperturas.BDWorkRecord.EOF Then
            lngNumero = 0
        Else
            lngNumero = claseBDAperturas.BDWorkRecord.Fields("Numero").Value
        End If
        claseBDAperturas.BDWorkRecord.Close()

        lngNumero = lngNumero + 1

        With claseBDAperturas.BDWorkRecord
            .Open("mpAsiHistError", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            .AddNew()
            .Fields("Referencia").Value = Referencia
            .Fields("Codsin").Value = "No Existe"
            .Fields("Numero").Value = lngNumero
            .Fields("Errores").Value = strMensajeError
            .Fields("Texto").Value = Texto
            .Fields("Fecgra").Value = Today
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
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()
        InsertarError = False
    End Function

    Public Sub New()
        MyBase.New()
        mvarReferencias = New Collection
        boolTransaccion = False
    End Sub

    ' Esta Función actualiza el estado de la referencia en la tabla de aperturas
    '
    Public Function ActualizarEstado(ByRef TipoErr As String, ByRef refererr As String) As Boolean

        On Error GoTo ActualizarEstado_Err

        ' Declaraciones
        '
        Dim objCmd As ADODB.Command ' Objeto command para ejecución de instrucciones Sql
        Dim strSQL As String ' Instrucción Sql

        ' Construimos la sentencia Sql para actualizar el estado
        '
        strSQL = "Update Angel_T1 Set T1_Estado = '" & TipoErr & "'" & ", T1_Codsin = '" & Codsin & "' , Estado_snt = 'P', Prov_Inic = '" & lProvision & "', " & " FechaProceso = '" & Now.Month.ToString & "/" & Now.Day & "/" & Now.Year.ToString & "'" & " Where T1_REFER = '" & refererr & "'"

        claseBDAperturas.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDAperturas.BDComand.CommandText = strSQL
        claseBDAperturas.BDComand.ActiveConnection = claseBDAperturas.BDWorkConnect
        claseBDAperturas.BDComand.Execute()

        ActualizarEstado = True
        Exit Function

ActualizarEstado_Err:
        ActualizarEstado = False
    End Function


    ' Esta función graba en la tabla SnSinies un nuevo siniestro. Devuelve un booleano
    '
    Public Function ActualizarSnSinies(ByRef objListItem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo AcualizarSnSinies_Err

        claseBDAperturas.BDWorkRecord.Open("SnSinies", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        With claseBDAperturas.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = Codsin
            .Fields("fecrec").Value = CDate(Now)
            .Fields("fecden").Value = CDate(objListItem.SubItems.Item(frmInstAperturas.T1_FAPER.Index).Text)
            .Fields("feccas").Value = CDate(objListItem.SubItems.Item(frmInstAperturas.T1_FSINI.Index).Text)
            .Fields("codram").Value = Trim(objListItem.SubItems.Item(frmInstAperturas.T1_CIA.Index).Text)
            .Fields("Numpol").Value = Trim(objListItem.SubItems.Item(frmInstAperturas.T1_POLIZA.Index).Text)
            .Fields("Codrie").Value = strRiesgo
            .Fields("Codmut").Value = sMutualista
            .Fields("Codcau").Value = Trim(objListItem.SubItems.Item(frmInstAperturas.T3_CCAUSA.Index).Text)
            .Fields("Estado").Value = "P"
            .Fields("Codage").Value = sAgente
            .Fields("dessin").Value = objListItem.SubItems.Item(frmInstAperturas.T1_DESCR.Index).Text
            .Fields("Canals").Value = "M"
            .Fields("Usuari").Value = strCodUserAplicacion
            .Fields("Numrec").Value = sRecibo
            .Fields("RecExt").Value = "R"
            .Fields("RefExt").Value = "AS" & strIdReferCompa & Trim(objListItem.Text)
            .Fields("Fecpdt").Value = CDate(Now)
            .Fields("Propen_P").Value = lProvision
            .Fields("Provis").Value = lProvision
            'JCLopez_i
            .Fields("Codcan").Value = "AA"
            'JCLopez_f
            .Update()
            .Close()
        End With
        ActualizarSnSinies = True
        Exit Function

AcualizarSnSinies_Err:
        ActualizarSnSinies = False
        strError = "El registro de la tabla de cabecera de Siniestros ha dado el error: " & Err.Description
    End Function

    ' Graba la provisión inicial en la aperetura de partes nuevos
    '
    Public Function ActualizarSnProvis(ByRef objListItem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo ActualizarSnProvis_Err

        claseBDAperturas.BDWorkRecord.Open("Snprovis", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = Codsin
            .Fields("Numprv").Value = "001"
            .Fields("Impprv").Value = lProvision
            .Fields("Fecprv").Value = CDate(Now)
            .Fields("Tipprv").Value = "P"
            .Fields("Comprv").Value = ""
            .Fields("Motpro").Value = "IN"
            .Fields("Fecmot").Value = CDate(Now)
            .Fields("Import").Value = lProvision
            .Fields("Gasnpr").Value = 0
            .Fields("Usuprov").Value = strCodUserAplicacion
            .Update()
            .Close()
        End With

        ActualizarSnProvis = True
        Exit Function

ActualizarSnProvis_Err:
        ActualizarSnProvis = False
        strError = "El registro de la tabla de Provisiones ha dado el error: " & Err.Description
    End Function

    ' Graba los datos del perjudicado en la aprtura de partes nuevos
    '
    Public Function ActualizarSnSinper(ByRef objListItem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo ActualizarSnSinper_Err

        claseBDAperturas.BDWorkRecord.Open("SnSinper", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = Codsin
            .Fields("Numper").Value = "001"
            .Fields("Apell1").Value = strNumCompa
            .Fields("Apell2").Value = ""
            .Fields("Nombre").Value = Mid(strNombreCompa, 1, 60)
            .Fields("Domici").Value = strDirecCompa
            .Fields("Codpos").Value = strCodPobCompa
            .Fields("Poblac").Value = strPoblaCompa
            .Fields("Provin").Value = ""
            .Fields("Pais").Value = "España"
            .Fields("Nifper").Value = strNIFCompa
            .Fields("Fecgra").Value = CDate(Now)
            .Update()
            .Close()
        End With
        ActualizarSnSinper = True

        Exit Function

ActualizarSnSinper_Err:
        ActualizarSnSinper = False
        strError = "El registro de la tabla de Personas ha producido el error: " & Err.Description
    End Function

    ' Graba los datos del perjudicado en la aprtura de partes nuevos
    '
    Public Function ActualizarSnAgenda(ByRef sCodsin As String) As Boolean

        On Error GoTo ActualizarSnAgenda_Err

        claseBDAperturas.BDWorkRecord.Open("SnAgenda", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = sCodsin
            .Fields("Fecha").Value = CDate(Now)
            .Fields("Descri").Value = "Recordatorio Evento"
            .Fields("Fecest").Value = CDate(Now)
            .Fields("Tipo").Value = "M"
            .Fields("Observ").Value = "Recordatorio Inicial"
            .Fields("Usuari").Value = strCodUserAplicacion
            .Fields("Codcar").Value = "105"
            .Fields("Fecrecordatorio").Value = System.DateTime.FromOADate(CDate(Now).ToOADate + 60)
            .Update()
            .Close()
        End With
        ActualizarSnAgenda = True

        Exit Function

ActualizarSnAgenda_Err:
        ActualizarSnAgenda = False
        strError = "El registro de la tabla de Agenda ha producido el error: " & Err.Description
    End Function

    Public Function ActualizarSnSinges(ByRef objListItem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo ActualizarSnSinges_Err

        claseBDAperturas.BDWorkRecord.Open("SnSinges", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = Codsin
            .Fields("Numges").Value = "001"
            .Fields("Tipges").Value = "RE"
            .Fields("Fecgra").Value = CDate(Now)
            .Fields("fecord").Value = CDate(Now)
            .Fields("Usuari").Value = strCodUserAplicacion
            .Fields("Numper").Value = strNumCompa
            .Update()
            .Close()
        End With
        ActualizarSnSinges = True

        Exit Function

ActualizarSnSinges_Err:
        ActualizarSnSinges = False
        strError = "El registro de apertura de gestión de reparación ha producido el error: " & Err.Description
    End Function

    ' Función que devuelve el código de riesgo de la poliza/ramo pasada como parámetro
    '
    Public Function RiesgoPoliza(ByRef sPoliza As String, ByRef sCodram As String) As String

        On Error GoTo RiesgoPoliza_Err

        ' Declaraciones
        '
        Dim strSQL As String ' Instrucción Sql

        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()

        RiesgoPoliza = vbNullString

        ' Instrucción Sql para buscar la póliza y leer el riesgo
        '
        strSQL = "Select Codris From Polizaca Where Numpol = '" & sPoliza & "' and Codram = '" & sCodram & "'"
        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If claseBDAperturas.BDWorkRecord.EOF Then
            Err.Raise(1)
        Else
            claseBDAperturas.BDWorkRecord.MoveFirst()
            RiesgoPoliza = claseBDAperturas.BDWorkRecord.Fields("Codris").Value
        End If
        claseBDAperturas.BDWorkRecord.Close()
        Exit Function

RiesgoPoliza_Err:
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()
    End Function

    ' Esta función devuelve el código del mutualista de la póliza especificada
    '
    Public Function MutualistaPoliza(ByRef sPoliza As String, ByRef sCodram As String) As String

        On Error GoTo MutualistaPoliza_Err

        ' Declaraciones
        '
        Dim strSQL As String ' Instruccion Sql

        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()

        MutualistaPoliza = vbNullString

        strSQL = "Select Numtom From Polizaca Where Numpol = '" & sPoliza & "' and Codram = '" & sCodram & "'"
        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If claseBDAperturas.BDWorkRecord.EOF Then
            MutualistaPoliza = "-1"
            Exit Function
        Else
            claseBDAperturas.BDWorkRecord.MoveFirst()
            MutualistaPoliza = claseBDAperturas.BDWorkRecord.Fields("Numtom").Value
        End If
        claseBDAperturas.BDWorkRecord.Close()

        Exit Function

MutualistaPoliza_Err:
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()
    End Function

    ' Esta función devuelve el código de agente de la póliza del siniestro
    '
    Public Function AgentePoliza(ByRef sPoliza As String, ByRef sCodram As String) As String

        On Error GoTo AgentePoliza_Err

        ' Declraciones
        '
        Dim strSQL As String ' Instrucción Sql

        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()

        AgentePoliza = vbNullString

        strSQL = "Select Codage From PolizaAg Where Numpol = '" & sPoliza & "' and Codram = '" & sCodram & "'"
        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If claseBDAperturas.BDWorkRecord.EOF Then
            Err.Raise(1)
        Else
            claseBDAperturas.BDWorkRecord.MoveFirst()
            AgentePoliza = claseBDAperturas.BDWorkRecord.Fields("Codage").Value
        End If
        claseBDAperturas.BDWorkRecord.Close()
        Exit Function

AgentePoliza_Err:
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()
    End Function

    ' Esta función graba en la tabla de numeradores el último siniestro
    ' creado
    '
    Public Function ActualizarNumerador(ByRef Codsin As String) As Boolean

        On Error GoTo ActualizarNumerador_Err

        ' Declaraciones
        '
        Dim strSQL As String

        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()

        strSQL = "Select Ultimo_Codigo From fic_apl Where nom_fic_apl = 'SNSINIES'"
        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDAperturas.BDWorkRecord.EOF Then
            Err.Raise(1)
        Else
            claseBDAperturas.BDWorkRecord.Fields("Ultimo_Codigo").Value = Mid(Codsin, 2, Len(Codsin) - 1)
            claseBDAperturas.BDWorkRecord.Update()
            claseBDAperturas.BDWorkRecord.Close()
        End If
        ActualizarNumerador = True
        Exit Function

ActualizarNumerador_Err:
        ActualizarNumerador = False
        strError = "El registro de numeradores ha devuelto el error: " & Err.Description
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()
    End Function

    Public Function ActualizarHistoricoEstados(ByRef sCodsin As String) As Boolean

        On Error GoTo ActualizarHistoricoEstados_Err

        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()

        claseBDAperturas.BDWorkRecord.Open("Sn_EstadoHist", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = sCodsin
            .Fields("Nummov").Value = "001"
            .Fields("Estado").Value = "A"
            .Fields("Fecest").Value = CDate(Now)
            .Fields("Usuari").Value = strCodUserAplicacion
            .Fields("Fecgra").Value = CDate(Now)
            .Update()
        End With
        claseBDAperturas.BDWorkRecord.Close()

        ActualizarHistoricoEstados = True
        Exit Function

ActualizarHistoricoEstados_Err:
        ActualizarHistoricoEstados = False
        If claseBDAperturas.BDWorkRecord.State = 1 Then claseBDAperturas.BDWorkRecord.Close()
    End Function

    ' Esta función borra las referencias pasadas en la colección de las tablas
    ' de aperturas de asistencia
    '
    Public Function DeleteAperturasAsistencia(ByRef ColRefers As Collection) As Boolean

        On Error GoTo DeleteAperturasAsistencia_Err

        ' Declaraciones
        '
        Dim strSQL As String ' Instrucción Sql
        Dim i As Short ' Contador para bucles

        ' Inicio de la transsacción
        '
        claseBDAperturas.BDWorkConnect.BeginTrans()

        For i = 1 To ColRefers.Count()

            ' Primero borramos el expediente de la tabla Angel_t1
            '
            strSQL = "Delete From Angel_t1 Where T1_Refer = '" & Trim(ColRefers.Item(i)) & "'"
            claseBDAperturas.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
            claseBDAperturas.BDComand.CommandText = strSQL
            claseBDAperturas.BDComand.ActiveConnection = claseBDAperturas.BDWorkConnect
            claseBDAperturas.BDComand.Execute()

            ' Después, y si la compañia es Angel,  lo borramos de la tabla Angel_t3
            '
            If strCodCia = "A" Then
                strSQL = "Delete From Angel_t3 Where T3_Refer = '" & Trim(ColRefers.Item(i)) & "'"
                claseBDAperturas.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
                claseBDAperturas.BDComand.CommandText = strSQL
                claseBDAperturas.BDComand.ActiveConnection = claseBDAperturas.BDWorkConnect
                claseBDAperturas.BDComand.Execute()
            End If
        Next i

        claseBDAperturas.BDWorkConnect.CommitTrans()
        DeleteAperturasAsistencia = True
        Exit Function

DeleteAperturasAsistencia_Err:
        claseBDAperturas.BDWorkConnect.RollbackTrans()
        DeleteAperturasAsistencia = False
    End Function

    ' Graba los datos del perjudicado en la tabla nueva de personas
    '
    Public Function ActualizarPersonas2(ByRef objListItem As ListViewItem, ByRef Codsin As String) As Boolean

        On Error GoTo ActualizarPersonas2_Err

        claseBDAperturas.BDWorkRecord.Open("Personas2", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            'TABLA PERSONAS2:
            .Fields("Per_Codigo").Value = "PER001" & Codsin
            .Fields("Per_TipoPersona").Value = "F"
            .Fields("Per_TipoDoc").Value = "C"
            .Fields("Per_NumDoc").Value = strNIFCompa
            .Fields("Per_Nombre").Value = Mid(strNombreCompa, 1, 30)
            .Fields("Per_Apell1").Value = strNumCompa
            .Fields("Per_Apell2").Value = ""
            .Fields("Per_RazSoc").Value = ""
            .Fields("Per_ApeNom").Value = strNumCompa & "," & Mid(strNombreCompa, 1, 30)
            .Fields("Per_NifOKSn").Value = ""
            .Fields("Per_NoResidente").Value = "N"
            .Fields("Per_Pasaporte").Value = ""
            .Fields("Per_Nombre_Trat").Value = Mid(strNombreCompa, 1, 30)
            .Fields("Per_Apell1_Trat").Value = strNumCompa
            .Fields("Per_Apell2_Trat").Value = ""
            .Fields("Per_RazSoc_Trat").Value = ""
            .Fields("Per_ApeNom_Trat").Value = strNumCompa & Mid(strNombreCompa, 1, 30)
            .Fields("Per_NomComercial").Value = ""
            .Fields("Per_NomComercial_Trat").Value = ""
            .Fields("Per_SDate").Value = CDate(Now)
            .Fields("Per_SUser").Value = strCodUserAplicacion
            .Update()
            .Close()
        End With

        claseBDAperturas.BDWorkRecord.Open("PersonasRelacion", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            'TABLA PERSONASRELACION:
            .Fields("ClaveRol").Value = "PER001" & Codsin
            .Fields("Per_Codigo").Value = "PER001" & Codsin
            .Update()
            .Close()
        End With

        claseBDAperturas.BDWorkRecord.Open("TMP_Direcciones", claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With claseBDAperturas.BDWorkRecord
            .AddNew()
            'TABLA TMP_DIRECCIONES
            .Fields("Tmp_CodDir").Value = "PER001" & Codsin
            .Fields("Tmp_TipoVia").Value = "CARRER"
            .Fields("Tmp_Calle").Value = "TARRAGONA"
            .Fields("Tmp_Poblacion").Value = "BARCELONA"
            .Fields("Tmp_Provincia").Value = "BARCELONA"
            .Fields("Tmp_Pais").Value = "ESPAÑA"
            .Fields("Coddir").Value = ""
            .Fields("DireccionEditada").Value = "TARRAGONA, 16"
            .Fields("DireccionOrigen").Value = "TARRAGONA, 16"
            .Fields("NumeroDesde").Value = "16"
            .Fields("NumeroHasta").Value = "16"
            .Fields("TipoVia").Value = "C"
            .Fields("Resto").Value = ""
            .Fields("CodigoPostal").Value = "08014"
            .Fields("CodigoPoblacion").Value = "08019"
            .Fields("CodigoProvincia").Value = 8
            .Fields("CodigoPais").Value = "ES"
            .Fields("CodigoCalle").Value = 2251
            .Fields("Numeracion_Confirmada").Value = "S"
            .Fields("FechaAlta").Value = CDate(Now)
            .Fields("UsuarioAlta").Value = strCodUserAplicacion
            .Update()
            .Close()
        End With

        ActualizarPersonas2 = True

        Exit Function

ActualizarPersonas2_Err:
        claseBDAperturas.BDWorkRecord.Close()
        ActualizarPersonas2 = False
        strError = "El registro de la tabla de Personas2 ha producido el error: " & Err.Description
    End Function

    'Funcion que devuelve el siguiente tramitador del grupo que le pasemos
    Public Function SiguienteTramitadorGrupo(ByRef sGrupo As String) As String
        On Error GoTo SiguienteTramitadorGrupo_Err

        ' Declaraciones

        Dim strSQL As String ' instrucción sql
        Dim iIndex As Integer

        SiguienteTramitadorGrupo = "0"


        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()

        'Se recupera el tramitador del siniestro según orden de preferencia
        strSQL = "SELECT empleado.num_empl FROM empleado INNER JOIN tramitadoresgrupossnt ON tramitadoresgrupossnt.tgs_clave = empleado.clave WHERE tramitadoresgrupossnt.tgs_grupo = '" & sGrupo & "' AND tgs_ordenasig = 1"

        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDAperturas.BDWorkRecord.BOF = True And claseBDAperturas.BDWorkRecord.EOF = True Then
            MsgBox("No se han encontrado tramitadores para el grupo " & sGrupo)
            Exit Function
        End If

        ' Si estamos antes del primero. mover al primero
        If claseBDAperturas.BDWorkRecord.BOF Then
            claseBDAperturas.BDWorkRecord.MoveFirst()
            ' Si estamos después del último, mover al último
        ElseIf claseBDAperturas.BDWorkRecord.EOF Then
            Exit Function
        End If

        If claseBDAperturas.BDWorkRecord.EOF Then
            SiguienteTramitadorGrupo = "0"
        Else
            SiguienteTramitadorGrupo = claseBDAperturas.BDWorkRecord.Fields("num_empl").Value
            If IsDBNull(SiguienteTramitadorGrupo) Or Trim(SiguienteTramitadorGrupo) = "" Then
                SiguienteTramitadorGrupo = "0"
            End If
        End If

        claseBDAperturas.BDWorkRecord.Close()

        'Se cogen los tramitadores para actualizar el orden de asignacion
        strSQL = "SELECT tgs_ordenasig FROM tramitadoresgrupossnt WHERE tramitadoresgrupossnt.tgs_grupo = '" & sGrupo & "' ORDER BY tgs_ordenasig"

        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDAperturas.BDWorkRecord.BOF = True And claseBDAperturas.BDWorkRecord.EOF = True Then
            MsgBox("No se han encontrado tramitadores para el grupo " & sGrupo)
            Exit Function
        End If

        iIndex = claseBDAperturas.BDWorkRecord.RecordCount

        claseBDAperturas.BDWorkRecord.Fields("tgs_ordenasig").Value = iIndex
        claseBDAperturas.BDWorkRecord.Update()
        claseBDAperturas.BDWorkRecord.MoveNext()

        iIndex = 0

        Do While Not claseBDAperturas.BDWorkRecord.EOF
            claseBDAperturas.BDWorkRecord.Fields("tgs_ordenasig").Value = 1 + iIndex
            iIndex = iIndex + 1
            claseBDAperturas.BDWorkRecord.Update()
            claseBDAperturas.BDWorkRecord.MoveNext()
        Loop

        claseBDAperturas.BDWorkRecord.Close()

        Exit Function
SiguienteTramitadorGrupo_Err:
        claseBDAperturas.BDWorkRecord.Close()
        SiguienteTramitadorGrupo = "0"
        strError = "Ha ocurrido un error al obtener el tramitador para el siniestro en el grupo " & sGrupo
    End Function


    Public Function ActualizarSiniatrib(ByRef strCodigoSiniestro As String, ByRef strTramitador As String) As Boolean
        'Inserta en la tabla siniatrib el numero de tramitador
        Dim strSQL As String

        strSQL = "SELECT * FROM siniatrib"

        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        With claseBDAperturas.BDWorkRecord
            .AddNew()
            .Fields("Codsin").Value = strCodigoSiniestro
            .Fields("codatr").Value = "TRAMITGRUPOHOGAR"
            .Fields("valor").Value = strTramitador
            .Fields("fecalta").Value = CDate(Now)
            .Fields("fecbaj").Value = System.DBNull.Value
            .Fields("s_date").Value = CDate(Now)
            .Fields("s_user").Value = strCodUserAplicacion
            .Update()
        End With

        claseBDAperturas.BDWorkRecord.Close()
        ActualizarSiniatrib = True
        Exit Function
ActualizarSiniatrib_Err:
        claseBDAperturas.BDWorkRecord.Close()
        ActualizarSiniatrib = False

    End Function

    Public Function GetGrupoRamo(ByRef strRamo As String) As Integer

        Dim strSQL As String

        GetGrupoRamo = 0

        'Recoge el grupo del ramo que le pasamos
        If claseBDAperturas.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDAperturas.BDWorkRecord.Close()

        strSQL = "SELECT gruram FROM ramos where codram = '" & strRamo & "'"

        claseBDAperturas.BDWorkRecord.Open(strSQL, claseBDAperturas.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If claseBDAperturas.BDWorkRecord.BOF = True And claseBDAperturas.BDWorkRecord.EOF = True Then
            Exit Function
        End If

        ' Si estamos antes del primero. mover al primero
        If claseBDAperturas.BDWorkRecord.BOF Then
            claseBDAperturas.BDWorkRecord.MoveFirst()
            ' Si estamos después del último, mover al último
        ElseIf claseBDAperturas.BDWorkRecord.EOF Then
            Exit Function
        End If

        If claseBDAperturas.BDWorkRecord.EOF Then
            GetGrupoRamo = 0
        Else
            GetGrupoRamo = claseBDAperturas.BDWorkRecord.Fields("gruram").Value
            If IsDBNull(GetGrupoRamo) Then
                GetGrupoRamo = 0
            End If
        End If

        claseBDAperturas.BDWorkRecord.Close()
    End Function

End Class
