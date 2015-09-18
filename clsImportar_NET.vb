Public Class clsImportar_NET

    Private mlngRegistros As Integer '
    Private mstrLog As String ' Fichero de log


    ' Lee el archivo de errores de la importación activa
    '

    ' Asigna el nombre del fichero de errores de la importación activa
    '
    Public Property FileLog() As String
        Get
            FileLog = mstrLog
        End Get
        Set(ByVal Value As String)
            mstrLog = Value
        End Set
    End Property

    ' Procedure:    Inicializar
    ' Objetivo:     Inicializa el objeto, y llama a la funcion de importacion que toca
    ' Parametros:   Proceso = Tipo de proceso de Importación
    '               ListaFicheros = Objeto ListBox con la lista de ficheroa a importar
    '
    Public Function Inicializar(ByRef Proceso As Short, ByRef ListaFicheros As System.Windows.Forms.ListBox) As Boolean

        On Error GoTo Inicializar_Err

        ' No realizar el contar registros del fichero si se escoge
        ' solamente proceso Referencias cruzadas
        '
        If Not Proceso = 6 Then
            mlngRegistros = ContarLineasFichero(ListaFicheros.Items.Item(0))


            If mlngRegistros = -1 Then
                MsgBox("No se ha seleccionado ningún proceso de importación a ejecutar.", MsgBoxStyle.Information)
                'globalNumerr = "4023"
                'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                'objError.Ver(IdProceso, globalNumerr, , Codcia)
            Else
                frmInstImportacion.prbProgreso.Visible = True
                frmInstImportacion.prbProgreso.Maximum = mlngRegistros
                frmInstImportacion.prbProgreso.Value = 1
                frmInstImportacion.stbEstado.Panels(1).Text = "0 %"
            End If
        End If

        If Not Procesar(Proceso, ListaFicheros) Then
            Inicializar = False
        Else
            Inicializar = True
        End If

        frmInstImportacion.stbEstado.Panels(1).Text = ""
        frmInstImportacion.prbProgreso.Visible = False

        Exit Function

Inicializar_Err:
        Inicializar = False
    End Function

    ' Funcion:      EsFicheroCia 
    ' Objetivo:     Determina si el fichero pertenece a una compañía de asistencia determinada
    ' Parametros:   pathfich -> path y nombre del fichero 
    '               cia -> compañía de asistencia
    '
    Private Function EsFicheroCia(ByVal pathfich As String, ByVal cia As String) As Boolean
        Dim spath As String
        Dim nomfich As String = " "
        Dim sext As String = " "

        EsFicheroCia = False
        On Error GoTo validarFicheroCia_Err

        objUtilidades.SplitPath(pathfich, spath, nomfich, sext)
        Select Case cia
            Case "I"
                If Mid(nomfich, 1, 2) = "MP" Then
                    EsFicheroCia = True
                End If
            Case "E"
                If Mid(nomfich, 1, 2) = "EA" Then
                    EsFicheroCia = True
                End If
            Case "M"
                If Mid(nomfich, 1, 2) = "MA" Then
                    EsFicheroCia = True
                End If
        End Select

validarFicheroCia_Err:

    End Function

    ' Procedure:    Procesar
    ' Objetivo:     Procesa el archivo a importar y llama su respectivo tipo de proceso
    ' Parametros:   Proceso = Tipo de proceso de Importación
    '
    Private Function Procesar(ByRef Proceso As Short, ByRef ListaFicheros As System.Windows.Forms.ListBox) As Boolean

        On Error GoTo Procesar_Err

        ' Declaraciones
        '
        Dim intFichero As Short
        Dim strLinea As String
        Dim NumLinea As Integer
        Dim strTipo As String
        Dim nFile As Integer

        Procesar = True
        NumLinea = 0
        nFile = 0
        mstrLog = objUtilidades.AddBackSlash(PathImportacion) & "Importacion_" & Format(Now, "yyyymmdd HHMMSS") & ".log"

        Do While nFile <= ListaFicheros.Items.Count - 1

            ' Realizar esta parte si no se a escogido solo Referencias cruzadas de siniestros
            '

            ''/*MUL INI */
            ' Validar que la empresa seleccionada carga el fichero correcto, solo nos podemos fijar en el nombre
            If EsFicheroCia(ListaFicheros.Items.Item(nFile), Codcia) = False Then
                strError = "El fichero no pertenece a la compañía seleccionada"
                Err.Raise(4001)
                Exit Do
            End If
            ''/*MUL FIN */
            If Not Proceso = 6 Then

                ' Comienzo de la transacción
                '
                claseBDImportar.BDWorkConnect.BeginTrans()
                Transaccion = True

                intFichero = FreeFile()

                FileOpen(intFichero, ListaFicheros.Items.Item(nFile), OpenMode.Input, OpenAccess.Read)

                Do While Not EOF(intFichero)
                    strLinea = LineInput(intFichero)

                    Call ActualizarPorcentaje(mlngRegistros)

                    strTipo = Left(strLinea, 1)
                    NumLinea = NumLinea + 1

                    ' El caso de Proceso = 0 es para cuando se escoge la opción de Todos.

                    Select Case strTipo
                        Case "1"
                            If Proceso = 1 Or Proceso = 0 Then ' Apertura de Siniestro
                                If Not Aperturas(strLinea, NumLinea, ListaFicheros.Items.Item(nFile)) Then
                                    globalNumerr = CStr(4001)
                                    Err.Raise(4001)
                                    Exit Do
                                End If
                            End If
                        Case "2"
                            If Proceso = 2 Or Proceso = 0 Then ' Pagos de Siniestro
                                'If NumLinea = 1077 Then
                                '    NumLinea = NumLinea
                                'End If
                                If Not Pagos(strLinea, NumLinea, ListaFicheros.Items.Item(nFile)) Then
                                    Err.Raise(4001)
                                    Exit Do
                                End If
                            End If
                        Case "3"
                            If Proceso = 3 Or Proceso = 0 Then ' Datos de los Causantes
                                If Not DatosCausantes(strLinea, NumLinea, ListaFicheros.Items.Item(nFile)) Then
                                    Err.Raise(4002)
                                    Exit Do
                                End If
                            End If
                        Case "4"
                            If Proceso = 4 Or Proceso = 0 Then ' Datos Fiscales
                                If Not DatosFiscales(strLinea, NumLinea, ListaFicheros.Items.Item(nFile)) Then
                                    Err.Raise(4003)
                                    Exit Do
                                End If
                            End If
                        Case "5", "6"
                            If Proceso = 7 Or Proceso = 0 Then ' Datos Anulaciones
                                If Not Anulaciones(strLinea, NumLinea, ListaFicheros.Items.Item(nFile)) Then
                                    Err.Raise(4006)
                                    Exit Do
                                End If
                            End If
                        Case "7"
                            If Proceso = 8 Or Proceso = 0 Then
                                If Not Suplidos(strLinea, NumLinea, ListaFicheros.Items.Item(nFile)) Then
                                    HaySuplidos = False
                                    Err.Raise(4001)
                                    Exit Do
                                Else
                                    HaySuplidos = True
                                End If
                            End If
                        Case "A"
                            If Proceso = 5 Or Proceso = 0 Then ' Seguimientros de Siniestros
                                If Not Seguimientos(strLinea, NumLinea, ListaFicheros.Items.Item(nFile)) Then
                                    Err.Raise(4004)
                                    Exit Do
                                End If
                            End If

                        Case Else
                            ' Error de log tipo linea no valida
                            'Err.Raise 4005
                            'Exit Do
                    End Select
                Loop
                claseBDImportar.BDWorkConnect.CommitTrans()
                Transaccion = False
                FileClose(intFichero)
            End If
            nFile = nFile + 1
        Loop

        ' Si se han importado suplidos de asistencia haceos el cruce de referencias
        '
        If HaySuplidos Then CruceReferenciasSuplidos()

        ' Referencias cruzadas de siniestros
        '
        If Proceso = 6 Or Proceso = 0 Then
            If Not ReferenciasCruzadasSiniestros() Then
                Err.Raise(4006)
                Exit Function
            End If
        End If

        Procesar = True
        Exit Function

Procesar_Err:
        If Err.Number = 4006 Then
            globalNumerr = "4021"
        ElseIf Err.Number = 4007 Then
            globalNumerr = "4022"
        Else
            MsgBox(strError, MsgBoxStyle.Critical)
            Call Log(mstrLog, strError, NumLinea, ListaFicheros.Items.Item(nFile))
        End If
        Procesar = False
        claseBDImportar.BDWorkConnect.RollbackTrans()
        '/* MUL INI si el fichero estaba abierto lo cerramos
        If intFichero > 0 Then
            FileClose(intFichero)
        End If
        '/* MUL FIN si el fichero estaba abierto lo cerramos
    End Function

    ' Procedure:    Aperturas
    ' Objetivo:     Procesa linea a importar del tipo Aperturas siniestro
    ' Parametros:   Linea = Linea a importar del tipo 1
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Private Function Aperturas(ByRef Linea As String, ByRef NumLinea As Integer, ByRef Archivo As String) As Boolean

        On Error GoTo AperturasError

        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Importando Aperturas siniestro ..."

        claseBDImportar.BDWorkConnect.Errors.Clear() ' Eliminar coleccion de errores

        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
        With claseBDImportar.BDWorkRecord
            .Open("angel_t1", claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic) ' Tabla de Aperturas de siniestros
            .AddNew()
            .Fields("T1_Codcia").Value = Codcia ' Código compañía de Asistencia en Mutua
            .Fields("T1_Tipmov").Value = Left(Linea, 1) ' Tipo de Registro
            '/* MUL INI
            '.Fields("T1_Poliza").Value = Mid(Linea, 2, 14) ' Poliza
            .Fields("T1_Poliza").Value = Mid(Linea, 2, 9) ' Poliza
            '/* MUL FIN
            .Fields("T1_FSini").Value = Mid(Linea, 16, 2) & "/" & Mid(Linea, 18, 2) & "/" & Mid(Linea, 20, 4) ' Fecha de siniestro (AAAAMMDD)
            .Fields("T1_FAper").Value = Mid(Linea, 24, 2) & "/" & Mid(Linea, 26, 2) & "/" & Mid(Linea, 28, 4) ' Fecha de apertura (AAAAMMDD)
            .Fields("T1_Descr").Value = Mid(Linea, 32, 255) ' Descripcion
            .Fields("T1_Refer").Value = Mid(Linea, 288, 10) ' Referencia Compañia Asistencia
            .Fields("T1_Aperturapor").Value = IIf(Mid(Linea, 298, 1) = "1", Mid(Linea, 298, 1), "") ' Indica a petición de quién intervine asistencia ( P = Perito )
            .Fields("T1_Cia").Value = Mid(Linea, 299, 5) ' Codram o Producto
            .Fields("T1_Cau_Per").Value = Mid(Linea, 304, 1) ' (C) Causante o (P) Perjudicado
            .Fields("T1_Refmutua").Value = Mid(Linea, 305, 8) ' Referencia de Mutua si existe
            '!T1_Codmod = Mid(Linea, 313, 3)      ' Código Modo Garantúa
            '!T1_Codgru = Mid(Linea, 316, 3)      ' Código de Grupo Garantía
            .Fields("T1_Codcausa").Value = Mid(Linea, 319, 5) ' Código de causa
            .Fields("T1_FGraba").Value = Today ' Fecha de Grabación
            .Fields("T1_Estado").Value = "X" ' Estado del siniestro
            .Fields("T1_FEstado").Value = Today ' Fecha del Estado
            .Fields("EstadoProcCierre").Value = "X" ' Estado del proceso de cierre
        End With
        claseBDImportar.BDWorkRecord.Update()
        claseBDImportar.BDWorkRecord.Close()

        Aperturas = True
        Exit Function

AperturasError:
        'If mdpbd.BDWorkRecord.State = 1 Then mdpbd.BDWorkRecord.Close
        strError = Err.Description
        Aperturas = False

        ''/*MUL INI */
        claseBDImportar.BDWorkRecord.CancelUpdate()
        claseBDImportar.BDWorkRecord.Close()
        ''/*MUL FIN */
    End Function

    ' Procedure:    Aperturas
    ' Objetivo:     Procesa linea a importar del tipo Aperturas siniestro
    ' Parametros:   Linea = Linea a importar del tipo 1
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Private Function Anulaciones(ByRef Linea As String, ByRef NumLinea As Integer, ByRef Archivo As String) As Boolean

        On Error GoTo Anulaciones_Err

        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Importando Anulaciones de siniestro ..."

        claseBDImportar.BDWorkConnect.Errors.Clear() ' Eliminar coleccion de errores

        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
        With claseBDImportar.BDWorkRecord
            .Open("anulacionesasistencia", claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic) ' Tabla de Anulaciones de siniestros
            .AddNew()
            .Fields("T5_Tipmov").Value = Left(Linea, 1) ' Tipo de Registro
            .Fields("T5_Refer").Value = IdReferCompa & Mid(Linea, 2, 10) ' Referencia Compañia Asistencia
            .Fields("T5_FecAnula").Value = Mid(Linea, 12, 2) & "/" & Mid(Linea, 14, 2) & "/" & Mid(Linea, 16, 4) ' Fecha de anulación
            .Fields("T5_CodRechazo").Value = Mid(Linea, 20, 3) ' Código de Motivo Anulación
            .Fields("T5_Descripcion").Value = objSiniestros.DescripcionAnulacion(Mid(Linea, 20, 3))
            .Fields("T5_Comentarios").Value = Mid(Linea, 23, 255) ' Comentarios
            .Fields("T5_FGraba").Value = Today ' Fecha de Grabación
            .Fields("T5_Codcia").Value = Codcia ' Código compañía de Asistencia en Mutua
            .Fields("T5_Estado").Value = "X" ' Estado del proceso de anulación
            .Fields("Fichero").Value = Archivo ' Nombre del fichero del que se importan los datos
        End With
        claseBDImportar.BDWorkRecord.Update()
        claseBDImportar.BDWorkRecord.Close()

        Anulaciones = True
        Exit Function

Anulaciones_Err:
        If Err.Number = -2147217873 Then
            claseBDImportar.BDWorkRecord.CancelUpdate()
            Resume Next
        Else
            strError = Err.Description
            Anulaciones = False
        End If
    End Function

    ' Procedure:  Pagos
    ' Objetivo:   Procesa linea a importar del tipo Pagos siniestro
    ' Parametros: Linea = Linea a importar del tipo 2
    ' Retorno:    True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    ' Cambio signo Importacion Pagos Asistencia
    '   01. Usuario / Fecha : Murquiza 21/11/2014
    '	   Doc. Relacionado : MAN-134
    '	   Descripción		: Cambio signo Importacion Pagos Asistencia
    Private Function Pagos(ByRef Linea As String, ByRef NumLinea As Integer, ByRef Archivo As String) As Boolean

        On Error GoTo Pagos_Err

        Dim sql As String
        '//@m001_i
        Dim li_esnegativo As Integer

        li_esnegativo = 1
        '//@m001_f
        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Importando Pagos de siniestro ..."

        '/*MUL T-19908 INI
        'If Codcia = "I" Then
        If Codcia = "I" Or Codcia = "M" Or Codcia = "E" Then
            '/*MUL T-19908 FIN
            sql = "SELECT COUNT(T2_REFER) AS REFER FROM Angel_t2 WHERE T2_REFER = '" & Mid(Linea, 2, 10) & "' and T2_NUMORD = '" & Mid(Linea, 12, 6) & "'"
        Else
            sql = "SELECT COUNT(T2_REFER) AS REFER FROM Angel_t2 WHERE T2_REFER = '" & Mid(Linea, 2, 10) & "' and T2_NUMORD = " & Mid(Linea, 12, 6)
        End If

        ' Eliminar coleccion de errores
        '
        claseBDImportar.BDWorkConnect.Errors.Clear()
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
        With claseBDImportar.BDWorkRecord
            .Open(sql, claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If .Fields(0).Value > 0 Then
                Err.Raise(1)
            End If
            .Close()
            .Open("angel_t2", claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic) ' Tabla de Aperturas de siniestros
            .AddNew()
            .Fields("T2_Codcia").Value = Codcia ' Código Compañía Asistencia en Mutua
            .Fields("T2_TipMov").Value = Left(Linea, 1) ' Tipo de registro
            .Fields("T2_REFER").Value = Mid(Linea, 2, 10) ' Referencia
            .Fields("T2_NumOrd").Value = Mid(Linea, 12, 6) ' Numero de orden interno
            .Fields("T2_FPago").Value = Mid(Linea, 18, 2) & "/" & Mid(Linea, 20, 2) & "/" & Mid(Linea, 22, 4) ' Fecha de orden (AAAAMMDD)
            .Fields("T2_TiMovi").Value = Mid(Linea, 26, 1) ' (1) Pago o (2) Cobro
            '//@m001_i
            If .Fields("T2_TiMovi").Value = 2 Then
                li_esnegativo = -1
            End If
            '.Fields("T2_impor").Value = objUtilidades.TextToNumeric(Mid(Linea, 27, 12), 2) ' Importe con IVA Incluido
            .Fields("T2_impor").Value = li_esnegativo * objUtilidades.TextToNumeric(Mid(Linea, 27, 12), 2) ' Importe con IVA Incluido
            '//@m001_f
            .Fields("T2_UltPag").Value = Mid(Linea, 39, 1) ' Ultimo pago
            .Fields("T2_Poliza").Value = Mid(Linea, 40, 14) ' Poliza
            .Fields("T2_Codmod").Value = Val(Mid(Linea, 54, 3)) ' Código Modo Garantía
            .Fields("T2_Codgru").Value = Val(Mid(Linea, 57, 3)) ' Código Grupo Garantía
            .Fields("T2_Codram").Value = Mid(Linea, 60, 5) ' Codram o Producto
            '//@m001_i
            '.Fields("T2_imptva").Value = objUtilidades.TextToNumeric(Mid(Linea, 65, 9), 2) ' Importe IVA
            .Fields("T2_imptva").Value = li_esnegativo * objUtilidades.TextToNumeric(Mid(Linea, 65, 9), 2) ' Importe IVA
            '//@m001_f
            .Fields("T2_factura").Value = Mid(Linea, 74, 11) ' Número de factura de la compañía de Asistencia
            .Fields("T2_Estado").Value = "X" ' X -> Pendiente
            .Fields("T2_Fgraba").Value = Today ' Fecha Grabación
            .Fields("T2_Codsin").Value = "" ' Codigo Siniestro
            .Fields("T2_Tipgas").Value = "I" ' Tipo G/I (Gasto/Indemnización)
            .Fields("T2_Pagomutua").Value = .Fields("T2_impor").Value
            .Fields("T2_Ivamutua").Value = .Fields("T2_imptva").Value
            .Fields("T2_Pagado").Value = "N" ' Indica si Mutua ha liquidado el pago con Angel

            .Update()
            .Close()
        End With
        Pagos = True
        Exit Function

Pagos_Err:
        If Err.Number = 1 Then
            strError = "Clave duplicada. El registro que intenta añadir ya existe"
        Else
            strError = Err.Description
        End If
        Pagos = False
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' Procedure:  DatosCausantes
    ' Objetivo:   Procesa linea a importar del tipo Datos de los causantes
    ' Parametros: Linea = Linea a importar del tipo 3
    ' Retorno:    True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Private Function DatosCausantes(ByRef Linea As String, ByRef NumLinea As Integer, ByRef Archivo As String) As Boolean

        On Error GoTo DatosCausantesError

        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Importando Datos de los Causantes ..."
        claseBDImportar.BDWorkConnect.Errors.Clear() ' Eliminar coleccion de errores
        With claseBDImportar.BDWorkRecord
            .Open("angel_t3", claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic) ' Tabla de Aperturas de siniestros
            .AddNew()

            .Fields("T3_TipMov").Value = Left(Linea, 1) ' Tipo de registro
            .Fields("T3_Poliza").Value = Mid(Linea, 2, 14) ' Poliza
            .Fields("T3_Refer").Value = Mid(Linea, 16, 10) ' Referencia
            .Fields("T3_Garant1").Value = Mid(Linea, 26, 2) ' Garantia afectada
            .Fields("T3_Garant2").Value = Mid(Linea, 28, 5) ' Garantia principal
            .Fields("T3_CausPer").Value = Mid(Linea, 33, 1) ' Causante (C) o Perjudicado (P)
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto objUtilidades.TextToNumeric(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            .Fields("T3_Valor").Value = objUtilidades.TextToNumeric(Mid(Linea, 34, 6), 2) ' Valoración del siniestro
            .Fields("T3_Riesgo").Value = Mid(Linea, 40, 3) ' Riesgo asegurado
            .Fields("T3_CCausa").Value = Mid(Linea, 43, 5) ' Codigo de causa
            .Fields("T3_SuRefe").Value = Mid(Linea, 48, 10) ' Referencia de la compañia cliente
            .Fields("T3_Sucur").Value = Mid(Linea, 58, 3) ' Sucursal
            .Fields("T3_Agente").Value = Mid(Linea, 61, 10) ' Agente
            .Fields("T3_FCausan").Value = Mid(Linea, 71, 1) ' No utilizado
            .Fields("T3_FCausad").Value = Mid(Linea, 72, 1) ' No utilizado
            .Fields("T3_HayDan").Value = Mid(Linea, 73, 1) ' No utilizado
            .Fields("T3_Danos").Value = Mid(Linea, 74, 40) ' Nombre del asegurado
            .Fields("T3_CCaus").Value = Mid(Linea, 114, 4) ' Codigo causa compañia cliente
            .Fields("T3_Cesti").Value = Mid(Linea, 118, 4) ' Codigo estimacion compañia cliente
            .Fields("T3_Garasis").Value = Mid(Linea, 122, 1) ' Gestion (G) o Asistencia (A)
            .Fields("T3_Biengar").Value = Mid(Linea, 123, 6) '
            .Fields("T3_Libre").Value = Mid(Linea, 129, 5)
            .Fields("T3_Fgraba").Value = Today ' Fecha de grbación de los datos

            .Update()
            .Close()
        End With
        DatosCausantes = True
        Exit Function

DatosCausantesError:
        strError = Err.Description
        DatosCausantes = False
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' Procedure:    DatosFiscales
    ' Objetivo:     Procesa linea a importar del tipo Datos fiscales
    ' Parametros:   Linea = Linea a importar del tipo 4
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Private Function DatosFiscales(ByRef Linea As String, ByRef NumLinea As Integer, ByRef Archivo As String) As Boolean

        On Error GoTo DatosFiscalesError

        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Importando Datos fiscales de siniestro ..."

        claseBDImportar.BDWorkConnect.Errors.Clear() ' Eliminar coleccion de errores
        With claseBDImportar.BDWorkRecord
            .Open("angel_t4", claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic) ' Tabla de Aperturas de siniestros
            .AddNew()
            .Fields("T4_TipMov").Value = Left(Linea, 1) ' Tipo de registro
            .Fields("T4_Poliza").Value = Mid(Linea, 2, 14) ' Poliza
            .Fields("T4_NumOrd").Value = Mid(Linea, 16, 6) ' Numero de orden interno angel
            .Fields("T4_Nombre").Value = Mid(Linea, 22, 36) ' Nombre o Razon Social
            .Fields("T4_Dni").Value = Mid(Linea, 58, 10) ' DNI, NIf o CIF
            .Fields("T4_Via").Value = Mid(Linea, 68, 2) ' Tipo de Via
            .Fields("T4_Domicil").Value = Mid(Linea, 70, 25) ' Domicilio
            .Fields("T4_Numero").Value = Mid(Linea, 95, 5) ' Número
            .Fields("T4_CPostal").Value = Mid(Linea, 100, 5) ' Codigo Postal
            .Fields("T4_Pobla").Value = Mid(Linea, 105, 24) ' Poblacion
            .Fields("T4_refer").Value = Mid(Linea, 129, 10) ' Referencia de Angel
            .Fields("T4_Fgraba").Value = Today ' Fecha de Grabación

            .Update()
            .Close()
        End With
        DatosFiscales = True
        Exit Function

DatosFiscalesError:
        strError = Err.Description
        DatosFiscales = False
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' Procedure:    Seguimientos
    ' Objetivo:     Procesa linea a importar del tipo Seguimientos de siniestros
    ' Parametros:   Linea = Linea a importar del tipo A
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Private Function Seguimientos(ByRef Linea As String, ByRef NumLinea As Integer, ByRef Archivo As String) As Boolean

        ' Declaraciones
        '
        Dim objRec As ADODB.Recordset

        On Error GoTo SeguimientosError

        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Importando Seguimientos de siniestro ..."
        claseBDImportar.BDWorkConnect.Errors.Clear() ' Eliminar coleccion de errores
        With claseBDImportar.BDWorkRecord
            .Open("angel_seguimientos", claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic) ' Tabla de Aperturas de siniestros
            .AddNew()

            .Fields("Tipo").Value = Left(Linea, 1) ' Tipo de registro
            .Fields("Referencia").Value = Mid(Linea, 2, 10) ' Referencia
            .Fields("Entrada").Value = DateSerial(CInt(Mid(Linea, 12, 4)), CInt(Mid(Linea, 16, 2)), CInt(Mid(Linea, 19, 2))) ' Fecha de entrada (AAAAMMDD)
            .Fields("Accion").Value = DateSerial(CInt(Mid(Linea, 25, 4)), CInt(Mid(Linea, 29, 2)), CInt(Mid(Linea, 31, 2))) ' Fecha de accion (AAAAMMDD)
            .Fields("Cod_ACC").Value = 0 ' Código de Seguimiento (No usado en Angel)
            .Fields("Cia").Value = Mid(Linea, 46, 7) ' Codigo de contrato
            .Fields("Descripcio").Value = Mid(Linea, 53, 76) ' Descripcion de la nota
            .Fields("Nota").Value = Mid(Linea, 129, 2) ' Nº de nota dentro del seguimiento
            .Fields("Libre").Value = "" ' No utilizado
            .Fields("Fecha").Value = Today

            .Update()
            .Close()
        End With
        Seguimientos = True
        Exit Function

SeguimientosError:
        strError = Err.Description
        Seguimientos = True
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' Procedure:    ReferenciasCruzadasSiniestros
    ' Objetivo:     Procesa las referencias cruzadas entre las aperturas importadas
    '               y las tablas de siniestros
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Private Function ReferenciasCruzadasSiniestros() As Boolean

        On Error GoTo ReferenciasCruzadasSiniestros_Err

        Call CruceReferencias(IdReferCompa, Codcia)
        'If Not CruceReferencias(IdReferCompa, Codcia) Then
        '   Err.Raise 1
        'End If

        ReferenciasCruzadasSiniestros = True
        Exit Function

ReferenciasCruzadasSiniestros_Err:
        strError = "El proceso de cruce de referencias de siniestros ha producido un error"
        ReferenciasCruzadasSiniestros = False
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ' Procedure:  ActualizarPorcentaje
    ' Objetivo:   Actualiza barra de progreso y barra de estado del frmInstanciaPrincipal
    '             con porcentajes.
    ' Parametros: Total = cantidad maxima que va a ver la barra de progreso.
    '
    Private Sub ActualizarPorcentaje(ByRef total As Integer)

        Dim intPorcentaje As Short

        On Error Resume Next

        If Not total = -1 Or total = 0 Then ' Actualizar barra de estado, de progreso y porcentaje
            frmInstanciaPrincipal.prbProgreso.Value = frmInstanciaPrincipal.prbProgreso.Value + 1
            intPorcentaje = System.Math.Round((frmInstanciaPrincipal.prbProgreso.Value * 100) / total, 0)
            If CStr(intPorcentaje) & " %" <> frmInstanciaPrincipal.stbEstado.Panels(1).Text Then
                frmInstanciaPrincipal.stbEstado.Panels(1).Text = CStr(intPorcentaje) & " %"
            End If
        End If

    End Sub

    ' Obtiene la compañia asistencia con la que se trabaja
    '
    'Private Function ObtenerCiaForm() As String
    '    ObtenerCiaForm = Trim(frmInstanciaPrincipal.lbxCompania.Items.Item(frmInstanciaPrincipal.cbxCompania.SelectedIndex))
    'End Function


    ' Procedure:  ObtenerCodigoSiniestro
    ' Objetivo:   Obtiene codigo de siniestros a partir de una referencia.
    ' Parametros: RefExt = Referencia externa compañia asistencia
    ' Retorno:    Codigo de Siniestro
    '
    Private Function ObtenerCodigoSiniestro(ByRef RefExt As String) As String

        On Error GoTo ObtenerCodigoSiniestro_Err

        ' Declaraciones
        '
        Dim strAplicacion As String
        Dim strsql As String
        Dim rsLocal As ADODB.Recordset

        ' Creación de objetos
        '
        rsLocal = New ADODB.Recordset

        ' Instrucción SQL
        '
        strsql = "Select Snsinies.Codsin From Snsinies Where Snsinies.RefExt = '" & RefExt & "'"

        ' Ejecución de la SQL
        '
        rsLocal.Open(strsql, claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        ' Si el resultado de la busqueda es nulo o nada devolvemos una
        ' cadena vacia, de lo contrario devolvemos el número sniestro.
        '
        If rsLocal.BOF And rsLocal.EOF Then
            ObtenerCodigoSiniestro = ""
        Else
            rsLocal.MoveFirst()
            'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
            If Not IsDBNull(rsLocal.Fields("Codsin").Value) Then
                ObtenerCodigoSiniestro = rsLocal.Fields("Codsin").Value
            Else
                ObtenerCodigoSiniestro = ""
            End If
        End If
        rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        rsLocal = Nothing

        Exit Function

ObtenerCodigoSiniestro_Err:
        ObtenerCodigoSiniestro = "Error"
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' Procedure:  ObtenerPerito
    ' Objetivo:   Obtiene si un Siniestro tiene o no Perito, a partir de una referencia.
    ' Parametros: Refer = Referencia externa compañia asistencia
    ' Retorno:    S / N
    '
    Private Function ObtenerPerito(ByRef Codsin As String) As String

        On Error GoTo ObtenerPerito_Err

        ' Declaraciones
        '
        Dim strsql As String
        Dim rsLocal As ADODB.Recordset

        ' Creación de objetos
        '
        rsLocal = New ADODB.Recordset

        ' Instrucción SQL de busqueda de peritos del Siniestros
        '
        strsql = "Select Count(*) From Snsinges Where Codsin = '" & Codsin & "' and TipGes = 'P'"

        ' Ejecución de la SQL
        '
        rsLocal.Open(strsql, claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

        If rsLocal.BOF Or rsLocal.EOF Then
            ObtenerPerito = "N"
        Else
            rsLocal.MoveFirst()
            If rsLocal.Fields(0).Value >= 0 Then
                ObtenerPerito = "S"
            Else
                ObtenerPerito = "N"
            End If
        End If
        rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        rsLocal = Nothing

        Exit Function

ObtenerPerito_Err:
        ObtenerPerito = "Error"
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' Une los registros tipo T1 y T3 en la tabla Angel_Aper
    ' Devuelve un booleano
    '
    Private Function Fusion() As Boolean

        On Error GoTo Fusion_Err

        ' Declaraciones
        '
        Dim strsql As String
        Dim lngresult As Integer

        ' Instrucción Sql para inserción de registros
        '
        strsql = "Insert Into Angel_Aper Select * " & "From   Angel_T1, Angel_T3 " & "Where  Angel_T1.T1_poliza = Angel_T3.T3_poliza and " & "       Angel_T1.T1_refer = Angel_T3.T3_refer and " & "       Not Exists (Select * From Angel_Aper Where Angel_Aper.T1_Refer = Angel_T1.T1_Refer)"

        ' Ejecución de la Instrucción SQL
        '
        claseBDImportar.BDComand.ActiveConnection.ConnectionString = claseBDImportar.BDWorkConnect.ConnectionString
        claseBDImportar.BDComand.CommandType = ADODB.CommandTypeEnum.adCmdText
        claseBDImportar.BDComand.CommandText = strsql
        claseBDImportar.BDComand.Execute(lngresult)

        Fusion = True
        Exit Function

Fusion_Err:
        Fusion = False
    End Function

    ' Procedure:  ContarRegistros
    ' Objetivo:   Calcula la cantidad de registros tiene un archivo.
    ' Parametros: Fichero = Ruta/Nombre fichero a calcular longitud
    ' Retorno:    Retorna la cantidad de registros que tiene el archivo.
    '
    Private Function ContarLineasFichero(ByRef Fichero As String) As Integer

        On Error GoTo ContarLineasFichero_Err

        Dim lngLongitud As Integer
        Dim intFichero As Short
        Dim strLinea As String

        intFichero = FreeFile()

        FileOpen(intFichero, Fichero, OpenMode.Input, OpenAccess.Read)

        Do While Not EOF(intFichero)
            strLinea = LineInput(intFichero)
            lngLongitud = lngLongitud + 1
        Loop

        FileClose(intFichero)

        ContarLineasFichero = lngLongitud

        Exit Function

ContarLineasFichero_Err:
        ContarLineasFichero = -1
    End Function

    ' Esta función devuelve el código de ramo de la póliza especificada como
    ' parametro.
    '
    Private Function CodramFromPoliza(ByRef Poliza As String) As String

        On Error GoTo CodramFromPoliza_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset ' Objeto recordset para lectura
        ' de registros en la BD
        Dim StrSqlMan As String ' Cadena con la instrucción Sql

        ' Establecemos la Sql de busqueda
        '
        StrSqlMan = "Select Codram From Polizaca Where Numpol = '" & Poliza & "'"
        rsLocal.Open(StrSqlMan, claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)

        If rsLocal.EOF Then
            CodramFromPoliza = "No existe"
        Else
            CodramFromPoliza = rsLocal.Fields(0).Value
        End If
        rsLocal.Close()
        Exit Function

CodramFromPoliza_Err:
        CodramFromPoliza = "Error"
        If rsLocal.State = 1 Then rsLocal.Close()
    End Function

    ' Procedure:    ReferenciasCruzadasSiniestros
    ' Objetivo:     Procesa las referencias cruzadas entre las aperturas importadas
    '               y las tablas de siniestros
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Public Function CruceReferencias(ByRef IdReferCompa As String, ByRef Codcia As String) As Boolean

        On Error GoTo CruceReferencias_Err

        CruceReferencias = True

        ' Declaraciones
        '
        Dim strsql As String
        Dim objCmd As ADODB.Command
        Dim strSiniestro As String
        Dim numRegis As Integer
        Dim lngRegistros As Integer
        Dim strCia As String

        '/*MUL T-19908 INI
        'If Codcia = "I" Then
        If Codcia = "I" Or Codcia = "M" Or Codcia = "E" Then
            '/*MUL T-19908 FIN

            ' Actualizamos el código de siniestro de la tabla de aperturas cruzando
            ' los campos de referencia externa con la tabla maestro de siniestro
            '
            strsql = "UPDATE angel_t1 " & "SET    T1_CODSIN = snsinies.CODSIN, T1_ESTADO = 'P' " & _
                     "FROM   Snsinies " & _
                     "WHERE  Snsinies.Refext = 'AS" & IdReferCompa & "' + angel_t1.T1_REFER AND " & _
                     "       (angel_t1.T1_CODSIN = '' OR angel_t1.T1_CODSIN IS NULL or angel_t1.T1_CODSIN = 'No Existe') and " & _
                     "       T1_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' Marcamos los que no han sido encontrados como ' No Existe'
            '
            strsql = "UPDATE angel_t1 " & "SET    T1_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = 'AS" & IdReferCompa & "' +angel_t1.T1_REFER AND " & "       ((angel_t1.T1_CODSIN = '' OR angel_t1.T1_CODSIN IS NULL) and Angel_t1.T1_Codsin <> 'No Existe' ) and " & "       T1_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' --------------------------------------------------------------------------
            '  JLL - 26/03/2004 Modificación
            ' --------------------------------------------------------------------------

            '   Despues de comentarlo con Araceli decidimos realizar la siguiente
            '   modificación:  Actualizamos el código de siniestro en la tabla de pagos
            '                  Angel_T2 independientemente de que exista la apertura en
            '                  la tabla de aperturas

            strsql = "UPDATE angel_t2 " & "SET    T2_CODSIN = snsinies.CODSIN " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = 'AS" & IdReferCompa & "' +angel_t2.T2_REFER AND " & "       (angel_t2.T2_CODSIN = '' OR angel_t2.T2_CODSIN IS NULL OR angel_t2.T2_CODSIN = 'No Existe') and " & "       T2_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' Actualizamos marcamos los que no han sido encontrados como 'No Existe'
            '
            strsql = "UPDATE angel_t2 " & "SET    T2_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = 'AS" & IdReferCompa & "' +angel_t2.T2_REFER AND " & "       ((angel_t2.T2_CODSIN = '' OR angel_t2.T2_CODSIN IS NULL) and angel_t2.T2_CODSIN <> 'No Existe') and " & "       T2_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' ----------------------------------------------------------------------
            '  Fín de la modificación de JLL - 26/03/2004
            ' ----------------------------------------------------------------------

            ' Realizamos el cruce entre Referencias y Siniestros para la
            ' tabla de Anulaciones de Siniestros
            '
            ' Cuándo tenga una rato modificaré todo el código anterior
            ' porque da pena mirarlo y ademas puede ser mucho más rápido
            '

            ' Abrimos la transacción
            '
            claseBDImportar.BDWorkConnect.BeginTrans()
            Transaccion = True

            ' Primero asignamos el siniestro a cada referencia
            '
            strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = SNSINIES.Codsin " & "FROM   Snsinies " & "WHERE  'AS' + AnulacionesAsistencia.T5_Refer = Snsinies.Refext and " & "       AnulacionesAsistencia.T5_Codcia = '" & Codcia & "'"
            claseBDImportar.BDWorkConnect.Execute(strsql)

            ' Si no tienen siniestro abierto marcamos como 'No Existe'
            '
            strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = 'No ExiSte' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin is null and " & "       AnulacionesAsistencia.T5_Codcia = '" & Codcia & "'"
            claseBDImportar.BDWorkConnect.Execute(strsql)
        Else
            ' Actualizamos el código de siniestro de la tabla de aperturas cruzando
            ' los campos de referencia externa con la tabla maestro de siniestro
            '
            strsql = "UPDATE angel_t1 " & "SET    T1_CODSIN = snsinies.CODSIN, T1_ESTADO = 'P' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' + angel_t1.T1_REFER AND " & "       (angel_t1.T1_CODSIN = '' OR angel_t1.T1_CODSIN IS NULL or angel_t1.T1_CODSIN = 'No Existe') and " & "       T1_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' Marcamos los que no han sido encontrados como ' No Existe'
            '
            strsql = "UPDATE angel_t1 " & "SET    T1_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' +angel_t1.T1_REFER AND " & "       ((angel_t1.T1_CODSIN = '' OR angel_t1.T1_CODSIN IS NULL) and Angel_t1.T1_Codsin <> 'No Existe' ) and " & "       T1_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' --------------------------------------------------------------------------
            '  JLL - 26/03/2004 Modificación
            ' --------------------------------------------------------------------------

            '   Despues de comentarlo con Araceli decidimos realizar la siguiente
            '   modificación:  Actualizamos el código de siniestro en la tabla de pagos
            '                  Angel_T2 independientemente de que exista la apertura en
            '                  la tabla de aperturas

            strsql = "UPDATE angel_t2 " & "SET    T2_CODSIN = snsinies.CODSIN " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' +angel_t2.T2_REFER AND " & "       (angel_t2.T2_CODSIN = '' OR angel_t2.T2_CODSIN IS NULL OR angel_t2.T2_CODSIN = 'No Existe') and " & "       T2_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' Actualizamos marcamos los que no han sido encontrados como 'No Existe'
            '
            strsql = "UPDATE angel_t2 " & "SET    T2_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' +angel_t2.T2_REFER AND " & "       ((angel_t2.T2_CODSIN = '' OR angel_t2.T2_CODSIN IS NULL) and angel_t2.T2_CODSIN <> 'No Existe') and " & "       T2_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' ----------------------------------------------------------------------
            '  Fín de la modificación de JLL - 26/03/2004
            ' ----------------------------------------------------------------------

            ' Realizamos el cruce entre Referencias y Siniestros para la
            ' tabla de Anulaciones de Siniestros
            '
            ' Cuándo tenga una rato modificaré todo el código anterior
            ' porque da pena mirarlo y ademas puede ser mucho más rápido
            '

            ' Abrimos la transacción
            '
            claseBDImportar.BDWorkConnect.BeginTrans()
            Transaccion = True

            ' Primero asignamos el siniestro a cada referencia
            '
            strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = SNSINIES.Codsin " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Refer = Snsinies.Refext and " & "       AnulacionesAsistencia.T5_Codcia = '" & Codcia & "'"
            claseBDImportar.BDWorkConnect.Execute(strsql)

            ' Si no tienen siniestro abierto marcamos como 'No Existe'
            '
            strsql = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = 'No ExiSte' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin is null and " & "       AnulacionesAsistencia.T5_Codcia = '" & Codcia & "'"
            claseBDImportar.BDWorkConnect.Execute(strsql)
        End If

        ' Cerramos la transacción
        '
        claseBDImportar.BDWorkConnect.CommitTrans()
        Transaccion = False

        'UPGRADE_NOTE: El objeto objCmd no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        objCmd = Nothing
        CruceReferencias = True

        Exit Function

CruceReferencias_Err:
        CruceReferencias = False
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' 04/02/2010 - JLL
    '
    ' Esta función baja los ficheros de aperturas y anulaciones de la compañía
    ' de asistencia desde un disco FTP
    '
    Public Function BajarFTPApe() As Boolean

        On Error GoTo BajarFTPApe_Err

        Dim Comando As String
        Dim NumFicherosAntes, NumFicherosDespues As Integer
        Dim Sigue As Boolean
        Dim Cadena As String
        Dim RetVal As Object
        Dim CompruebaFinal As String
        Dim Canal As Short

        ' Eliminamos el fichero de marca final del proceso anterior
        '
        Kill("C:\Final.txt")

        ' Establecemos la unidad y directorio activos para los ficheros de aperturas
        '
        ChDrive("K")
        ChDir(DatosFTPApe)
        frmInstImportacion.Dir1.Path = CurDir()

        ' Contamos el número de ficheros existentes en la carpeta destino antes de
        ' efectuar la bajada de archivos desde FTP
        '

        frmInstImportacion.File1.Refresh()
        frmInstImportacion.File1.Path = CurDir()
        NumFicherosAntes = 0
        NumFicherosDespues = 0
        NumFicherosAntes = frmInstImportacion.File1.Items.Count

        ' Configuramos la llamada parametrizada a la aplicación de bajada de archivos
        ' desde el disco FTP
        '
        Comando = ConfigFTP & "\psftp " & UsuarioFTP & DiscoFTP & " " & PasswordFTP & " -b " & ConfigFTP & "\ComandosGetApe.bat"

        ' Eliminamos el archivo .bat existente
        '
        Kill(ConfigFTP & "\FTPFile.bat")

        ' Creamos el arvhivo Bat
        '
        Canal = FreeFile()
        FileOpen(Canal, ConfigFTP & "\FTPFile.Bat", OpenMode.Output)
        PrintLine(Canal, Comando)
        PrintLine(Canal, "Dir > C:\Final.txt")
        FileClose(Canal)

        ' Configuramos y realizamos ejecución fichero .bat
        '
        Comando = ConfigFTP & "\FTPFile.bat"
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto RetVal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        RetVal = Shell(Comando)

        ' Obtenemos el número de fichero copiados y entramos en bucle en espera de
        ' encontrar el fichero Final.txt que indica que la copia ha terminado
        '
        Sigue = True
        Do While Sigue = True
            'UPGRADE_WARNING: Dir tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            CompruebaFinal = Dir("C:\Final.txt")
            If CompruebaFinal = "Final.txt" Then
                Sigue = False
                System.Windows.Forms.Application.DoEvents()
                ChDrive("K")
                ChDir(DatosFTPApe)
                frmInstImportacion.Dir1.Path = CurDir()
                frmInstImportacion.File1.Path = CurDir()
                frmInstImportacion.File1.Refresh()
                NumFicherosDespues = frmInstImportacion.File1.Items.Count
                NumFicherosDespues = NumFicherosDespues - NumFicherosAntes
                If NumFicherosDespues > 0 Then
                    Cadena = "Se han bajado " & CStr(NumFicherosDespues) & " ficheros de Aperturas y/o Anulaciones desde el disco FTP"
                    MsgBox(Cadena, MsgBoxStyle.Information)
                    'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                    'objError.Ver(IdProceso, , Cadena, Codcia)
                Else
                    globalNumerr = "4062"
                    MsgBox("No se han bajado ficheros desde el disco FTP", MsgBoxStyle.Critical)
                    'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                    'objError.Ver(IdProceso, globalNumerr, , Codcia)
                End If
            End If
        Loop
        ChDrive("C")
        ChDir("\")

        BajarFTPApe = True

        Exit Function

BajarFTPApe_Err:
        If Err.Number = 53 Then
            Resume Next
        End If
        If Err.Number = 70 Then
            Resume
        End If
        frmInstanciaPrincipal.Cursor = System.Windows.Forms.Cursors.Default
        BajarFTPApe = False
    End Function

    ' 04/02/2010 - JLL
    '
    ' Esta función baja los ficheros de aperturas y anulaciones de la compañía
    ' de asistencia desde un disco FTP
    '
    Public Function BajarFTPPag() As Boolean

        On Error GoTo BajarFTPPag_Err

        Dim Comando As String
        Dim NumFicherosAntes, NumFicherosDespues As Integer
        Dim Sigue As Boolean
        Dim Cadena As String
        Dim RetVal As Object
        Dim CompruebaFinal As String
        Dim Canal As Short

        ' Eliminamos el fichero de marca final del proceso anterior
        '
        Kill("C:\Final.txt")

        ' Establecemos la unidad y directorio activos para los ficheros de aperturas
        '
        ChDrive("K")
        ChDir(DatosFTPPag)
        frmInstImportacion.Dir1.Path = CurDir()
        frmInstImportacion.File1.Path = CurDir()

        ' Contamos el número de ficheros existentes en la carpeta destino antes de
        ' efectuar la bajada de archivos desde FTP
        '

        frmInstImportacion.File1.Refresh()
        NumFicherosAntes = 0
        NumFicherosDespues = 0
        NumFicherosAntes = frmInstImportacion.File1.Items.Count

        ' Configuramos la llamada parametrizada a la aplicación de bajada de archivos
        ' desde el disco FTP
        '
        Comando = ConfigFTP & "\psftp " & UsuarioFTP & DiscoFTP & " " & PasswordFTP & " -b " & ConfigFTP & "\ComandosGetPag.bat"

        ' Eliminamos el archivo .bat existente
        '
        Kill(ConfigFTP & "\FTPFile.bat")

        ' Creamos el arvhivo Bat
        '
        Canal = FreeFile()
        FileOpen(Canal, ConfigFTP & "\FTPFile.Bat", OpenMode.Output)
        PrintLine(Canal, Comando)
        PrintLine(Canal, "Dir > C:\Final.txt")
        FileClose(Canal)

        ' Configuramos y realizamos ejecución fichero .bat
        '
        Comando = ConfigFTP & "\FTPFile.bat"
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto RetVal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        RetVal = Shell(Comando)

        ' Obtenemos el número de fichero copiados y entramos en bucle en espera de
        ' encontrar el fichero Final.txt que indica que la copia ha terminado
        '
        Sigue = True
        Do While Sigue = True
            CompruebaFinal = Dir("C:\Final.txt")
            If CompruebaFinal = "Final.txt" Then
                System.Windows.Forms.Application.DoEvents()
                ChDrive("K")
                ChDir(DatosFTPPag)
                frmInstImportacion.Dir1.Path = CurDir()
                frmInstImportacion.File1.Refresh()
                NumFicherosDespues = frmInstImportacion.File1.Items.Count
                NumFicherosDespues = NumFicherosDespues - NumFicherosAntes
                If NumFicherosDespues > 0 Then
                    Cadena = "Se han bajado " & NumFicherosDespues & " ficheros de Pagos desde el disco FTP"
                    MsgBox(Cadena, MsgBoxStyle.Information)
                    'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                    'objError.Ver(IdProceso, , Cadena, Codcia)
                    Sigue = False
                Else
                    Sigue = False
                    globalNumerr = "4062"
                    MsgBox("No se han bajado ficheros desde el disco FTP", MsgBoxStyle.Information)
                    'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                    'objError.Ver(IdProceso, globalNumerr, , Codcia)
                End If
            End If
        Loop
        ChDrive("C")
        ChDir("\")

        BajarFTPPag = True

        Exit Function

BajarFTPPag_Err:
        If Err.Number = 53 Then
            Resume Next
        End If
        If Err.Number = 70 Then
            Resume
        End If
        frmInstanciaPrincipal.Cursor = System.Windows.Forms.Cursors.Default
        BajarFTPPag = False
    End Function

    ' 08/02/2010 - JLL
    '
    ' Esta función elimina los ficheros ya bajados ubicados en el disco FTP
    '
    Public Function DeleteFTP(ByRef TipoOP As String) As Boolean

        On Error GoTo DeleteFTP_Err

        Dim Comando As String
        Dim Sigue As Boolean
        Dim Cadena As String
        Dim RetVal As Object
        Dim CompruebaFinal As String
        Dim Canal As Short
        Dim FicheroBat As String
        Dim DatosFTP As String

        ' Eliminamos el fichero de marca final del proceso anterior
        '
        Kill("C:\Final.txt")

        ' Establecemos la unidad y directorio activos para los ficheros a borrar
        '
        If TipoOP = "Aperturas" Then
            ChDrive("K")
            DatosFTP = DatosFTPApe
            frmInstImportacion.Dir1.Path = CurDir()
            FicheroBat = "ComandosGetApeDel.bat"
        ElseIf TipoOP = "Pagos" Then
            ChDrive("K")
            DatosFTP = DatosFTPPag
            frmInstImportacion.Dir1.Path = CurDir()
            FicheroBat = "ComandosGetPagDel.bat"
        End If
        ChDir(DatosFTP)

        ' Configuramos la llamada parametrizada a la aplicación de bajada de archivos
        ' desde el disco FTP
        '
        Comando = ConfigFTP & "\psftp " & UsuarioFTP & DiscoFTP & " " & PasswordFTP & " -b " & ConfigFTP & "\" & FicheroBat

        ' Eliminamos el archivo .bat existente
        '
        Kill(ConfigFTP & "\FTPFile.bat")

        ' Creamos el arvhivo Bat
        '
        Canal = FreeFile()
        FileOpen(Canal, ConfigFTP & "\FTPFile.Bat", OpenMode.Output)
        PrintLine(Canal, Comando)
        PrintLine(Canal, "Dir > C:\Final.txt")
        FileClose(Canal)

        ' Configuramos y realizamos ejecución fichero .bat
        '
        Comando = ConfigFTP & "\FTPFile.bat"
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto RetVal. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        RetVal = Shell(Comando)

        ' Obtenemos el número de fichero copiados y entramos en bucle en espera de
        ' encontrar el fichero Final.txt que indica que la copia ha terminado
        '
        Sigue = True
        Do While Sigue = True
            'UPGRADE_WARNING: Dir tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            CompruebaFinal = Dir("C:\Final.txt")
            If CompruebaFinal = "Final.txt" Then
                Cadena = "Se han eliminado los ficheros " & TipoOP & " ubicados en el disco origen FTP"
                MsgBox(Cadena, MsgBoxStyle.Information)
                'objError.Tipo = mdpErroresMensajes.Tipo.Pantalla
                'objError.Ver(IdProceso, , Cadena, Codcia)
                Sigue = False
            End If
        Loop
        ChDrive("C")
        ChDir("\")
        DeleteFTP = True
        Exit Function

DeleteFTP_Err:
        DeleteFTP = False
    End Function

    ' Procedure:  Pagos
    ' Objetivo:   Procesa linea a importar del tipo Pagos siniestro
    ' Parametros: Linea = Linea a importar del tipo 2
    ' Retorno:    True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Private Function Suplidos(ByRef Linea As String, ByRef NumLinea As Integer, ByRef Archivo As String) As Boolean

        On Error GoTo Suplidos_Err

        Dim sql As String

        frmInstanciaPrincipal.stbEstado.Panels(1).Text = "Importando Suplidos de Asistencia ..."

        ' Eliminar coleccion de errores
        '
        claseBDImportar.BDWorkConnect.Errors.Clear()
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
        With claseBDImportar.BDWorkRecord
            .Open("SuplidosAsistencia", claseBDImportar.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic) ' Tabla de Suplidos de Asistencia
            .AddNew()
            .Fields("T7_TipMov").Value = Left(Linea, 1) ' Tipo de registro
            .Fields("T7_REFER").Value = Mid(Linea, 2, 10) ' Referencia
            .Fields("T7_Codsin").Value = "" ' Codigo Siniestro
            .Fields("T7_Codram").Value = Mid(Linea, 26, 5) ' Codram o Producto
            .Fields("T7_Numpol").Value = Mid(Linea, 12, 14) ' Poliza
            .Fields("T7_Codcia").Value = Codcia ' Código Compañía Asistencia en Mutua
            .Fields("T7_Estado").Value = "X" ' X -> Pendiente
            .Fields("T7_Fgraba").Value = Today ' Fecha Grabación
            .Fields("Fichero").Value = objUtilidades.NameFromFileName(Archivo)
            .Fields("T7_FacturaIP").Value = Mid(Linea, 31, 20)
            .Fields("T7_FechaFactura").Value = Mid(Linea, 51, 2) & "/" & Mid(Linea, 53, 2) & "/" & Mid(Linea, 55, 4)
            .Update()
            .Close()
        End With
        Suplidos = True
        Exit Function

Suplidos_Err:
        strError = Err.Description
        Suplidos = False
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()
    End Function

    ' Procedure:    ReferenciasCruzadasSiniestros por Suplidos
    ' Objetivo:     Procesa las referencias cruzadas entre las aperturas importadas
    '               y las tablas de siniestros
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Public Function CruceReferenciasSuplidos() As Object

        On Error GoTo CruceReferenciasSuplidos_Err

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CruceReferenciasSuplidos. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CruceReferenciasSuplidos = True

        ' Declaraciones
        '
        Dim strsql As String
        Dim objCmd As ADODB.Command
        Dim strSiniestro As String
        Dim numRegis As Integer
        Dim lngRegistros As Integer
        Dim strCia As String

        claseBDImportar.BDWorkConnect.BeginTrans()

        '/*MUL T-19908 INI
        'If Codcia = "I" Then
        If Codcia = "I" Or Codcia = "M" Or Codcia = "E" Then
            '/*MUL T-19908 FIN

            ' Actualizamos el código de siniestro de la tabla de suplidos cruzando
            ' los campos de referencia externa con la tabla maestro de siniestro
            '
            strsql = "UPDATE SuplidosAsistencia " & "SET    T7_CODSIN = SnSinies.CODSIN " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = 'AS" & IdReferCompa & "' + SuplidosAsistencia.T7_REFER AND " & "       (SuplidosAsistencia.T7_CODSIN = '' OR SuplidosAsistencia.T7_CODSIN IS NULL or SuplidosAsistencia.T7_CODSIN = 'No Existe') and " & "       T7_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

            ' Marcamos los que no han sido encontrados como ' No Existe'
            '
            strsql = "UPDATE SuplidosAsistencia " & "SET    T7_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = 'AS" & IdReferCompa & "' + SuplidosAsistencia.T7_REFER AND " & "       ((SuplidosAsistencia.T7_CODSIN = '' OR SuplidosAsistencia.T7_CODSIN IS NULL) and SuplidosAsistencia.T7_Codsin <> 'No Existe' ) and " & "       T7_Codcia = '" & Codcia & "'"

            objCmd = New ADODB.Command
            objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            objCmd.CommandText = strsql
            objCmd.ActiveConnection = claseBDImportar.BDWorkConnect
            objCmd.Execute(lngRegistros)

        End If

        ' Cerramos la transacción
        '
        claseBDImportar.BDWorkConnect.CommitTrans()
        Transaccion = False

        'UPGRADE_NOTE: El objeto objCmd no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        objCmd = Nothing
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CruceReferenciasSuplidos. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CruceReferenciasSuplidos = True

        Exit Function

CruceReferenciasSuplidos_Err:
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto CruceReferenciasSuplidos. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CruceReferenciasSuplidos = False
        If claseBDImportar.BDWorkRecord.State = 1 Then claseBDImportar.BDWorkRecord.Close()

    End Function
End Class
