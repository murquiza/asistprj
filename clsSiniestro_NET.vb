Public Class clsSiniestro_NET

    Public Observaciones As New ADODB.Recordset
    Public strUsuario As String
    ' ----------------------------------------------------------------------------
    ' Desarrollo:         Juan Carlos López - JCLopez
    ' Fecha Creación:     17/04/2014mpExcluyeGarantias
    ' Descripción:        Libreria que contiene las Clase que construye los objetos
    '                     de negocio de siniestros.
    '                     La clase clsAsistencia es para la gestión especifica del
    '                     area de Compañías de asistencia.
    '                     La clase clsSiniestros es general para todo el departa-
    '                     mento de siniestros
    '
    ' -----------------------------------------------------------------------------
    Public Function Siniestro(ByRef sCodsin As String, ByRef Ver As Boolean, ByRef sUser As String) As ADODB.Recordset

        On Error GoTo Siniestro_Err

        ' Declaraciones
        '
        Dim strsql As String ' Cadena que contiene la instrucción Sql

        ' Si el recordset esta abierto lo cerramos
        '
        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()
        If Observaciones.State = 1 Then Observaciones.Close()

        strUsuario = sUser

        ' Construimos la select para la busqueda del registro correspondiente
        ' y creamos el recordset con los datos del siniestro
        '
        strsql = "Select Snsinies.* From Snsinies Where Codsin = '" & sCodsin & "'"
        claseBDLibrerias.BDWorkRecord.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If claseBDLibrerias.BDWorkRecord.EOF Then
            Err.Raise(1)
        End If

        ' Construimos la select para la busqueda de las observaciones
        ' y creamos el recordset con los datos
        '
        strsql = "Select 'Fecha            ' =  ' ' + Fecha + '   ', 'Tramitador   ' = '  ' + Empleado.Nombre + ' ' + Apell1, 'Texto' = snagenda.observ  From Snagenda, Empleado Where codsin = '" & sCodsin & "' and Snagenda.Usuari *= Empleado.Num_empl and Snagenda.Observ is not null"
        Observaciones.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        Siniestro = claseBDLibrerias.BDWorkRecord
        '  If Ver Then frmSiniestro.Show vbModal
        Exit Function

Siniestro_Err:
        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()
        If Observaciones.State = 1 Then Observaciones.Close()
        'UPGRADE_NOTE: El objeto Siniestro no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'Siniestro = Nothing
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ' Esta función devuelve el numero de empleado correspondiente a la
    ' clave especificada en el parametro
    '
    Friend Function NumeroEmpleado(ByRef sClave As String) As String

        On Error GoTo NumeroEmpleado_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strsql As String

        ' Instrucción sql en la que buscamos en la tabla de empleados la
        ' clave que nos han pasado para obtener el número empleado
        '
        strsql = "Select Num_Empl From Empleado Where clave = '" & sClave & " '"

        ' Ejecución de la Sql
        '
        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not rsLocal.EOF Then
            rsLocal.MoveFirst()
            NumeroEmpleado = rsLocal.Fields("num_Empl").Value
        Else
            Err.Raise(1)
        End If
        rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'rsLocal = Nothing
        Exit Function

NumeroEmpleado_Err:
        NumeroEmpleado = "Error"
        rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'rsLocal = Nothing
    End Function

    ' Esta función devuelve la clave de empleado correspondiente al
    ' numero de empleado especificado en el parametro
    '
    Public Function ClaveEmpleado(ByRef sNumeroEmpleado As String) As String

        On Error GoTo ClaveEmpleado_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strsql As String

        ' Instrucción sql en la que buscamos en la tabla de empleados el
        ' numero empleado que nos han pasado para obtener el número empleado
        '
        strsql = "Select Clave From Empleado Where clave = '" & sNumeroEmpleado & " '"

        ' Ejecución de la Sql
        '
        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not rsLocal.EOF Then
            rsLocal.MoveFirst()
            ClaveEmpleado = rsLocal.Fields("Clave").Value
        Else
            Err.Raise(1)
        End If
        rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'rsLocal = Nothing
        Exit Function

ClaveEmpleado_Err:
        ClaveEmpleado = "Error"
        rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'rsLocal = Nothing
    End Function

    ''/* MUL INI  La he anulado porque daba problemas con el recordset de trabajo de la clase
    '''
    ''' Esta función devuelve el nombre del empleado a partir de la clave del
    ''' empleado
    '''
    '''    Public Function NombreEmpleado(ByRef sClave As String) As String

    '''        On Error GoTo NombreEmpleado_Err

    '''        ' Declaraciones
    '''        '
    '''        Dim rsLocal As New ADODB.Recordset
    '''        Dim strsql As String

    '''        ' Instrucción sql en la que buscamos en la tabla de empleados el
    '''        ' nombre empleado asociado a la clave que nos han pasado
    '''        '
    '''        strsql = "Select nombre + ' ' + apell1 + ' ' + apell2 as nombre From Empleado Where Num_empl = '" & sClave & " '"

    '''        ' Ejecución de la Sql
    '''        '
    '''        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

    '''        If Not rsLocal.EOF Then
    '''            rsLocal.MoveFirst()
    '''            NombreEmpleado = rsLocal.Fields("nombre").Value
    '''        Else
    '''            Err.Raise(1)
    '''        End If
    '''        rsLocal.Close()
    '''        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
    '''        'rsLocal = Nothing
    '''        Exit Function

    '''NombreEmpleado_Err:
    '''        NombreEmpleado = "Error"
    '''        rsLocal.Close()
    '''        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
    '''        'rsLocal = Nothing
    '''    End Function
    ''/* MUL FIN  La he anulado porque daba problemas con el recordset de trabajo de la clase

    ' Función que devuelve la provisión inicial a realizar en el nuevo siniestro según
    ' su causa y ramo
    '
    Public Function ProvisionInicialSiniestro(ByRef sCodram As String, ByRef sCodcau As String, ByRef bAsistencia As Boolean) As Integer

        On Error GoTo ProvisionInicialSiniestro_Err

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim sRamo As String

        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()

        ProvisionInicialSiniestro = 0

        ' Construimos la sentencia Sql según el siniestro sea o no de Asistencia...
        '
        If bAsistencia Then
            strsql = "Select Provision From Sncausas_Ramo_Asistencia " & "Where Codram = '" & sCodram & "' and Uso = 'S'"
        Else
            ' En este caso primero debemos obtener el ramo que corresponde al
            ' producto o Codram que deseamos tratar
            '
            strsql = "Select Ramo1 From Ramos Where CCodram = '" & sCodram & "'"
            claseBDLibrerias.BDWorkRecord.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If claseBDLibrerias.BDWorkRecord.EOF Then
                Err.Raise(1)
                Exit Function
            Else
                sRamo = claseBDLibrerias.BDWorkRecord.Fields("Ramo1").Value
            End If
            claseBDLibrerias.BDWorkRecord.Close()
            strsql = "Select Provis From Sncausas_Ramo " & "Where Ramo1 = '" & sRamo & "' and Codcau = '" & sCodcau & "'"
        End If
        claseBDLibrerias.BDWorkRecord.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If claseBDLibrerias.BDWorkRecord.EOF Then
            Err.Raise(1)
        Else
            claseBDLibrerias.BDWorkRecord.MoveFirst()
            ProvisionInicialSiniestro = claseBDLibrerias.BDWorkRecord.Fields(0).Value
        End If
        claseBDLibrerias.BDWorkRecord.Close()
        Exit Function

ProvisionInicialSiniestro_Err:
        If claseBDLibrerias.BDWorkRecord.State = ADODB.ObjectStateEnum.adStateOpen Then claseBDLibrerias.BDWorkRecord.Close()
    End Function

    ' Esta función devuelve el número de recibo que da cobertura al siniestro
    '
    Public Function ReciboSiniestro(ByRef sPoliza As String, ByRef sCodram As String, ByRef dFecCas As Date) As String

        On Error GoTo ReciboSiniestro_Err

        ' Declraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim objutiles As clsUtilidades_NET

        ' Creación de objetos
        '
        objutiles = New clsUtilidades_NET ' Objeto de funciuones y utilidades genéricas

        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()

        ReciboSiniestro = vbNullString

        strsql = "Select Numrec, Numsub, Fesuvt, Fesuvt From Carterac " & "Where Numpol = '" & sPoliza & "' and Codram = '" & sCodram & "' and Fesure < '" & objutiles.FormatoFechaSQL(dFecCas, False, False) & "'" & "and Anldor = 'N' Order By Fesure Desc"

        claseBDLibrerias.BDWorkRecord.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If claseBDLibrerias.BDWorkRecord.EOF Then
            Err.Raise(1)
        Else
            claseBDLibrerias.BDWorkRecord.MoveFirst()
            ReciboSiniestro = claseBDLibrerias.BDWorkRecord.Fields("Numrec").Value
        End If
        claseBDLibrerias.BDWorkRecord.Close()
        'UPGRADE_NOTE: El objeto objutiles no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'objutiles = Nothing
        Exit Function

ReciboSiniestro_Err:
        'UPGRADE_NOTE: El objeto objutiles no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'objutiles = Nothing
        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()
    End Function

    ' Esta función devuelve el número de siniestro a grabar
    '
    Public Function ObtenerNumeroSiniestro(ByRef Codram As String) As String

        On Error GoTo ObtenerNumeroSiniestro_Err

        ' Declaraciones
        '
        Dim strsql As String ' Instrucción Sql
        Dim lcodsin As Object 'Codigo del siniestros
        Dim lany As Integer ' Tratamiento del año y del número de siniestro
        Dim sCodsin As String ' Alamacenamiento número de siniestro
        Dim Ramo As Object
        Dim sany As String ' Almacenamiento ramo y año + 1
        Dim IniAper As String

        ObtenerNumeroSiniestro = ""

        ' Sql para obtener el último ordinal de siniestros
        '
        strsql = "Select Ultimo_Codigo From fic_apl Where nom_fic_apl = 'SNSINIES'"
        claseBDLibrerias.BDWorkRecord.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDLibrerias.BDWorkRecord.EOF Then
            claseBDLibrerias.BDWorkRecord.MoveFirst()
            lcodsin = Val(claseBDLibrerias.BDWorkRecord.Fields("Ultimo_Codigo").Value) + 1
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto lcodsin. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            sCodsin = Trim(Str(lcodsin))
        Else
            Err.Raise(1)
        End If
        claseBDLibrerias.BDWorkRecord.Close()

        ' Sql para obtener el ramo al que pertenece el producto ( codram )
        '
        strsql = "Select Ramo1 From Ramos Where Codram ='" & Codram & "'"
        claseBDLibrerias.BDWorkRecord.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not claseBDLibrerias.BDWorkRecord.EOF Then
            claseBDLibrerias.BDWorkRecord.MoveFirst()
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Ramo. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            Ramo = claseBDLibrerias.BDWorkRecord.Fields("Ramo1").Value
        Else
            Err.Raise(1)
        End If
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Ramo. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        ObtenerNumeroSiniestro = Ramo + sCodsin
        claseBDLibrerias.BDWorkRecord.Close()
        Exit Function

ObtenerNumeroSiniestro_Err:
        ObtenerNumeroSiniestro = vbNullString
    End Function

    ' Esta función devuelve el importe de la Provisión pendiente de pagos de un
    ' siniestro o todos a una fecha o todas ( Pagado = S )
    '
    Public Function ProvisionPendiente(ByRef Fecha As String, Optional ByRef sCodsin As String = "") As Double

        On Error GoTo ProvisionPendiente_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strsql As String
        Dim dtFechaAux As New Date

        If Fecha = "" Or IsDBNull(Fecha) Then
            Fecha = "12/31/2100"
        End If

        dtFechaAux = dtFechaAux.ParseExact(Fecha, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture)


        ' Construimos la instrucción Sql
        '

        'strsql = "SELECT " & "      (SELECT ROUND(ISNULL(SUM(IMPORT),0),2) " & "       FROM SNPROVIS " & "       WHERE  SNPROVIS.CODSIN = SNSINIES.CODSIN AND " & "              SNPROVIS.FECPRV <= '" & VB6.Format(Fecha, "mm/dd/yyyy") & "' AND " & "              SNPROVIS.TIPPRV = 'P' ) - " & "      (SELECT isnull(ROUND(SUM(isnull(IMPORT,0) + isnull(IMPGAS,0) + isnull(IMPIVA,0)),2),0) " & "       FROM SNSINCTA " & "       WHERE SNSINCTA.CODSIN = SNSINIES.CODSIN AND " & "             SNSINCTA.FECPAG < '" & VB6.Format(Fecha, "mm/dd/yyyy") & "' AND " & "             SNSINCTA.PAGADO = 'S' AND " & "             SNSINCTA.TIPGAS in('I','P','R','G')) as ProvPdte " & "FROM  SNSINIES "
        strsql = "SELECT " & "      (SELECT ROUND(ISNULL(SUM(IMPORT),0),2) " & "       FROM SNPROVIS " & "       WHERE  SNPROVIS.CODSIN = SNSINIES.CODSIN AND " & "              SNPROVIS.FECPRV <= '" & dtFechaAux.Month.ToString & "/" & dtFechaAux.Day.ToString & "/" & dtFechaAux.Year.ToString & "' AND " & "              SNPROVIS.TIPPRV = 'P' ) - " & "      (SELECT isnull(ROUND(SUM(isnull(IMPORT,0) + isnull(IMPGAS,0) + isnull(IMPIVA,0)),2),0) " & "       FROM SNSINCTA " & "       WHERE SNSINCTA.CODSIN = SNSINIES.CODSIN AND " & "             SNSINCTA.FECPAG < '" & dtFechaAux.Month.ToString & "/" & dtFechaAux.Day.ToString & "/" & dtFechaAux.Year.ToString & "' AND " & "             SNSINCTA.PAGADO = 'S' AND " & "             SNSINCTA.TIPGAS in('I','P','R','G')) as ProvPdte " & "FROM  SNSINIES "

        'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
        strsql = strsql & IIf(IsNothing(sCodsin), "", "WHERE SNSINIES.CODSIN = '" & sCodsin & "'")

        ' Ejecutamos la Sql
        '
        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        ' Si se ha ejecutado correctamente asignamos el valor
        '

        If Not rsLocal.EOF Then
            ProvisionPendiente = rsLocal.Fields("ProvPdte").Value
        End If
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        rsLocal = Nothing
        Exit Function

ProvisionPendiente_Err:
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        ProvisionPendiente = -1
    End Function

    Public Function ProvisionPendiente(ByRef Fecha As Date, Optional ByRef sCodsin As String = "") As Double

        On Error GoTo ProvisionPendiente_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strsql As String
        Dim dtFechaAux As String

        If Not IsDate(Fecha) Or IsDBNull(Fecha) Or IsNothing(Fecha) Then
            Fecha = CDate("12/31/2100")
        End If

        ' Construimos la instrucción Sql
        '

        'strsql = "SELECT " & "      (SELECT ROUND(ISNULL(SUM(IMPORT),0),2) " & "       FROM SNPROVIS " & "       WHERE  SNPROVIS.CODSIN = SNSINIES.CODSIN AND " & "              SNPROVIS.FECPRV <= '" & VB6.Format(Fecha, "mm/dd/yyyy") & "' AND " & "              SNPROVIS.TIPPRV = 'P' ) - " & "      (SELECT isnull(ROUND(SUM(isnull(IMPORT,0) + isnull(IMPGAS,0) + isnull(IMPIVA,0)),2),0) " & "       FROM SNSINCTA " & "       WHERE SNSINCTA.CODSIN = SNSINIES.CODSIN AND " & "             SNSINCTA.FECPAG < '" & VB6.Format(Fecha, "mm/dd/yyyy") & "' AND " & "             SNSINCTA.PAGADO = 'S' AND " & "             SNSINCTA.TIPGAS in('I','P','R','G')) as ProvPdte " & "FROM  SNSINIES "
        strsql = "SELECT " & "      (SELECT ROUND(ISNULL(SUM(IMPORT),0),2) " & "       FROM SNPROVIS " & "       WHERE  SNPROVIS.CODSIN = SNSINIES.CODSIN AND " & "              SNPROVIS.FECPRV <= '" & Fecha.Month.ToString & "/" & Fecha.Day.ToString & "/" & Fecha.Year.ToString & "' AND " & "              SNPROVIS.TIPPRV = 'P' ) - " & "      (SELECT isnull(ROUND(SUM(isnull(IMPORT,0) + isnull(IMPGAS,0) + isnull(IMPIVA,0)),2),0) " & "       FROM SNSINCTA " & "       WHERE SNSINCTA.CODSIN = SNSINIES.CODSIN AND " & "             SNSINCTA.FECPAG < '" & Fecha.Month.ToString & "/" & Fecha.Day.ToString & "/" & Fecha.Year.ToString & "' AND " & "             SNSINCTA.PAGADO = 'S' AND " & "             SNSINCTA.TIPGAS in('I','P','R','G')) as ProvPdte " & "FROM  SNSINIES "

        'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
        strsql = strsql & IIf(IsNothing(sCodsin), "", "WHERE SNSINIES.CODSIN = '" & sCodsin & "'")

        ' Ejecutamos la Sql
        '
        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        ' Si se ha ejecutado correctamente asignamos el valor
        '

        If Not rsLocal.EOF Then
            ProvisionPendiente = rsLocal.Fields("ProvPdte").Value
        End If
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        Exit Function

ProvisionPendiente_Err:
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        ProvisionPendiente = -1
    End Function


    ' Esta función devuelve el importe de la Provisión pendiente de pagos de un
    ' siniestro o todos a una fecha o todas ( Pagado = a todo )
    '
    Public Function ProvisionDisponible(ByRef Fecha As Object, Optional ByRef sCodsin As String = "") As Double

        ' 10/06/2004  JLL
        ' -------------------------------------------------------------------
        '    Esta función se modifica para que calcule el total de provision
        '    independientemente de la fecha.
        '    De momento no elimino el parametro para no afectar a todas las
        '    llamadas. Lo que hago es modificar la Sql
        ' -------------------------------------------------------------------

        On Error GoTo ProvisionDisponible_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strsql As String

        'If Fecha = "" Or IsNull(Fecha) Then
        '    Fecha = "12/31/2100 23:59:59"
        'End If

        ' Construimos la instrucción Sql
        '
        strsql = "SELECT " & "      (SELECT ROUND(ISNULL(SUM(IMPORT),0),2) " & "       FROM SNPROVIS " & "       WHERE  SNPROVIS.CODSIN = SNSINIES.CODSIN AND " & "              SNPROVIS.TIPPRV = 'P' ) - " & "      (SELECT isnull(ROUND(SUM(isnull(IMPORT,0) + isnull(IMPGAS,0) + isnull(IMPIVA,0)),2),0) " & "       FROM SNSINCTA " & "       WHERE SNSINCTA.CODSIN = SNSINIES.CODSIN AND " & "             SNSINCTA.TIPGAS in('I','P','R','G')) as ProvPdte " & "FROM  SNSINIES "

        'UPGRADE_NOTE: IsMissing() ha cambiado a IsNothing(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1021"'
        strsql = strsql & IIf(IsNothing(sCodsin), "", "WHERE SNSINIES.CODSIN = '" & sCodsin & "'")

        ' Ejecutamos la Sql
        '
        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        ' Si se ha ejecutado correctamente asignamos el valor
        '

        If Not rsLocal.EOF Then
            ProvisionDisponible = rsLocal.Fields("ProvPdte").Value
        End If
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        rsLocal = Nothing
        Exit Function

ProvisionDisponible_Err:
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        ProvisionDisponible = -1
    End Function

    ' Esta función devuelve el imnporte de pagos de un siniestro ( pagado = s )
    ' o todos a una fecha o todas
    '
    Public Function PagosSiniestros(ByRef Fecha As Object, Optional ByRef sCodsin As String = "") As Double

        On Error GoTo PagosSiniestros_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strsql As String

        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fecha. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If Fecha = "" Or IsDBNull(Fecha) Then
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fecha. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            Fecha = "12/31/2100 23:59:59"
        End If

        ' Construimos la instrucción Sql
        '
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Fecha. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        strsql = "Select isnull(round(sum(isnull(import,0)+isnull(impgas,0)+isnull(impiva,0)),2),0) Pagos " & "From   Snsincta " & "Where  snsincta.codsin = '" & sCodsin & "' and " & "       snsincta.pagado = 'S' and " & "       snsincta.tipgas in ('I','P','R','G') and " & "       snsincta.fecpag <= '" & Fecha & "'"

        ' Ejecutamos la Sql
        '
        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        ' Si se ha ejecutado correctamente asignamos el valor
        '
        If Not rsLocal.EOF Then
            PagosSiniestros = rsLocal.Fields("Pagos").Value
        End If
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        rsLocal = Nothing

        Exit Function

PagosSiniestros_Err:
        Resume
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        PagosSiniestros = -1
    End Function

    ' Esta función devuelve el siguiente número de movimiento valido a
    ' grabar de una tabla especificada en el parámetro sTabla
    '
    Public Function NumeroMovimientoTabla(ByRef CampoNum As String, ByRef sTabla As String, ByRef sCampoKey As String, ByRef sValCampoKey As String) As String

        On Error GoTo NumeroMovimientoTabla_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strsql As String
        Dim Cadena, SubCadena As String

        ' Construimos la instrucción Sql
        '
        strsql = "Select Max(" & CampoNum & ") " & "From " & sTabla & " " & "Where " & sCampoKey & " = '" & sValCampoKey & "'"

        ' Ejecutamos la Sql
        '
        rsLocal.Open(strsql, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        ' Si se ha ejecutado correctamente asignamos el valor
        '
        'UPGRADE_WARNING: Se detectó el uso de Null o IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1049"'
        If Not IsDBNull(rsLocal.Fields(0).Value) Then
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Cadena. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            Cadena = rsLocal.Fields(0).Value + 1
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto SubCadena. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            SubCadena = ""

            'SubCadena = New String("0", 3 - Len(Cadena)) & Cadena
            SubCadena = New String("0", 3 - Len(Cadena)) & Cadena

        Else
            SubCadena = "001"
        End If

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto SubCadena. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        NumeroMovimientoTabla = SubCadena

        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()

        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        rsLocal = Nothing

        Exit Function

NumeroMovimientoTabla_Err:
        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()
        NumeroMovimientoTabla = CStr(-1)
    End Function

    ' Esta Función realiza una provisión en el siniestro por el importe y
    ' el tipo de provisión especificada
    '
    Public Function AjusteTecnico(ByRef sCodsin As String, ByRef Importe As Object, ByRef Tipo As String, ByRef Motivo As String, ByRef Usuario As String) As Boolean

        On Error GoTo AjusteTecnico_err

        ' Declaraciones
        '
        Dim sql As String
        Dim Numero As String
        Dim rsLocal As New ADODB.Recordset

        Numero = NumeroMovimientoTabla("Numprv", "Snprovis", "Codsin", sCodsin)

        rsLocal.Open("Snprovis", claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        With rsLocal
            .AddNew()
            .Fields("Codsin").Value = sCodsin
            .Fields("Numprv").Value = Numero
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Importe. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            .Fields("Impprv").Value = Importe
            .Fields("Fecprv").Value = CDate(Now)
            .Fields("Tipprv").Value = Tipo
            .Fields("Comprv").Value = ""
            .Fields("Motpro").Value = Motivo
            .Fields("Fecmot").Value = CDate(Now)
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto Importe. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            .Fields("Import").Value = Importe
            .Fields("Gasnpr").Value = 0
            .Fields("Usuprov").Value = Usuario
            .Update()
            .Close()
        End With

        If rsLocal.State = ADODB.ObjectStateEnum.adStateOpen Then rsLocal.Close()

        'UPGRADE_NOTE: El objeto rsLocal no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'rsLocal = Nothing

        AjusteTecnico = True

        Exit Function

AjusteTecnico_err:
        AjusteTecnico = False
    End Function
End Class
