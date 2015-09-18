Public Class clsAsistencia_NET

    Public Transaccion As Boolean
    ' ----------------------------------------------------------------------------
    ' Desarrollo:         Jose Luis de Lacalle
    ' Fecha Creaci�n:     19/02/2003
    ' Descripci�n:        Libreria que contiene las Clase que construye los objetos
    '                     de negocio de siniestros.
    '                     La clase clsAsistencia es para la gesti�n especifica del
    '                     area de Compa��as de asistencia.
    '                     La clase clsSiniestros es general para todo el departa-
    '                     mento de siniestros
    '
    ' Fecha Modificaci�n: 03/06/2003
    ' -----------------------------------------------------------------------------


    Private pCia As String ' Varibale local propiedad CodCiaAsist
    Private pColLng As Collection ' Colecci�n privada para la propiedad LongitudCampos
    Private pColPos As Collection ' Colecci�n privada para la propiedad PosicionInicial
    Private pRelleno As Collection ' Colecci�n privada para la propiedad Relleno
    Private pRecordSet As ADODB.Recordset ' RecordSet privado con los datos de formato de lcia de asistencia
    Private pNumCmp As Integer ' N�mero de campos a exportar
    Private pColAli As Collection ' Colecci�n privada para la propiedad Alineacion
    Private pFichero As String ' Fichero del cual se obtuene los datos de formato

    ' Esta propiedad asigna el valor de la compa�ia de asistencia de
    ' la cual se extraeran los datos de formato
    '

    ' Esta propiedad devuelve el valor de la compa�ia de asistencia de
    ' la cual se han extraido los datos de formato
    '
    Public Property CodCiaAsist() As String
        Get
            CodCiaAsist = pCia
        End Get
        Set(ByVal Value As String)
            pCia = Value

            pColLng = New Collection
            pColPos = New Collection
            pColAli = New Collection
            pRelleno = New Collection

            ' Llamamos a la funci�n que carga los valores de formato de la
            ' compa��a especificada en las colecciones
            '
            FormatoCia()

        End Set
    End Property

    ' Esta propiedad asigna el valor de la compa�ia de asistencia de
    ' la cual se extraeran los datos de formato
    '
    Public WriteOnly Property ParaFichero() As String
        Set(ByVal Value As String)
            pFichero = Value
        End Set
    End Property

    ' Esta propiedad devuelve una colecci�n con las longitudes de todos
    ' los campos a exportar
    '
    Public ReadOnly Property LongitudCampos() As Collection
        Get
            LongitudCampos = pColLng
        End Get
    End Property

    ' Esta propiedad devuelve una colecci�n con lss caracteres de relleno de todos
    ' los campos a exportar
    '
    Public ReadOnly Property Relleno() As Collection
        Get
            Relleno = pRelleno
        End Get
    End Property
    ' Esta propiedad devuelve una colecci�n con las posiciones iniciales de todos
    ' los campos a exportar
    '
    Public ReadOnly Property PosicionInicial() As Collection
        Get
            PosicionInicial = pColPos
        End Get
    End Property

    ' Esta propiedad devuelve el numero de campos a exportar que
    ' forman el registro de exportaci�n
    '
    Public ReadOnly Property NumeroCampos() As Integer
        Get
            NumeroCampos = pNumCmp
        End Get
    End Property

    ' Esta propiedad devuelve el c�digo de alineaci�n de los campos a exportar que
    ' forman el registro de exportaci�n
    '
    Public ReadOnly Property Alineacion() As Collection
        Get
            Alineacion = pColAli
        End Get
    End Property

    ' Esta funci�n rellena las colecciones privadas con los datos de
    ' formato de la compa�ia asiganada en la propiedad CodCiaAsist
    '
    Private Sub FormatoCia()

        On Error GoTo FormatoCia_Err

        ' Declaraciones
        '
        Dim strSqlman As String ' Cadena que contiene la instrucci�n Sql
        Dim lngResult As Integer ' N�mero de filas afectadas por la ejecuci�n
        ' de command de Ado

        strSqlman = "Select * from mpAsiCiasFormatos Where CodCia = '" & pCia & "' and Archivo = '" & pFichero & "'"

        pRecordSet.Open(strSqlman, claseBDLibrerias.BDSystemConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        pRecordSet.MoveFirst()
        With pRecordSet.Fields
            Do While Not pRecordSet.EOF
                pColLng.Add(.Item("Longitud").Value, Str(.Item("Orden").Value))
                pColPos.Add(.Item("PosicionInicial").Value, Str(.Item("Orden").Value))
                pColAli.Add(.Item("Alinea").Value, Str(.Item("Orden").Value))
                pRelleno.Add(.Item("AsciiRelleno").Value, Str(.Item("Orden").Value))
                pNumCmp = .Item("Orden").Value
                pRecordSet.MoveNext()
            Loop
        End With

        pRecordSet.Close()
        Exit Sub

FormatoCia_Err:
        MsgBox("No existen datos de formato para la compa��a especificada", MsgBoxStyle.Critical, "Error en captaci�n datos Exportaci�n")
    End Sub

    ' En la creaci�n de la clase...
    '
    'UPGRADE_NOTE: Class_Initialize se actualiz� a Class_Initialize_Renamed. Haga clic aqu� para obtener m�s informaci�n: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
    Private Sub Class_Initialize_Renamed()

        ' Creamos los objetos locales
        '
        pRecordSet = New ADODB.Recordset

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    ' En la destrucci�n de la clase...
    '
    'UPGRADE_NOTE: Class_Terminate se actualiz� a Class_Terminate_Renamed. Haga clic aqu� para obtener m�s informaci�n: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
    Private Sub Class_Terminate_Renamed()

        ' Destruimos los objetos locales
        '
        'UPGRADE_NOTE: El objeto pColLng no se puede destruir hasta que no se realice la recolecci�n de los elementos no utilizados. Haga clic aqu� para obtener m�s informaci�n: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        pColLng = Nothing
        'UPGRADE_NOTE: El objeto pColPos no se puede destruir hasta que no se realice la recolecci�n de los elementos no utilizados. Haga clic aqu� para obtener m�s informaci�n: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        pColPos = Nothing
        'UPGRADE_NOTE: El objeto pRecordSet no se puede destruir hasta que no se realice la recolecci�n de los elementos no utilizados. Haga clic aqu� para obtener m�s informaci�n: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        pRecordSet = Nothing
        'UPGRADE_NOTE: El objeto pColAli no se puede destruir hasta que no se realice la recolecci�n de los elementos no utilizados. Haga clic aqu� para obtener m�s informaci�n: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        pColAli = Nothing

    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub

    ' Esta funci�n devuelve la descripci�n del c�digo de rechazo
    ' pasado como parametro
    '
    Public Function DescripcionAnulacion(ByRef sCodigo As String) As String

        On Error GoTo DescripcionAnulacion_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strSQL As String

        ' Construimos la sentencia Sql
        '
        strSQL = "Select Descripcion From CodigosAnulacionesAsistencia Where CodigoAnulacion = '" & sCodigo & "'"

        ' La Ejecutamos
        '
        rsLocal.Open(strSQL, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        ' Devolvemos el valor
        '
        If Not rsLocal.EOF Then
            DescripcionAnulacion = rsLocal.Fields("Descripcion").Value
        End If
        rsLocal.Close()

        Exit Function

DescripcionAnulacion_Err:
        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()
        DescripcionAnulacion = "Error"
    End Function

    ' Esta funci�n devuelve el n�mero de perjudicado relativo al siniestro, adjudicado a la
    ' compa�ia de asistencia cuando esta se registra como perjudicado
    '
    Public Function NumeroPerjudicadoAsistencia(ByRef sCodsin As String, ByRef sCia As String) As String

        On Error GoTo NumeroPerjudicadoAsistencia_Err

        ' Declaraciones
        '
        Dim rsLocal As New ADODB.Recordset
        Dim strSQL As String

        ' Construimos la sentencia Sql
        '
        strSQL = "Select Numper From SnSinper Where Apell1 = '" & sCia & "' and Codsin ='" & sCodsin & "'"

        ' La Ejecutamos
        '
        rsLocal.Open(strSQL, claseBDLibrerias.BDWorkConnect, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        ' Devolvemos el valor
        '
        If Not rsLocal.EOF Then
            NumeroPerjudicadoAsistencia = rsLocal.Fields("numper").Value
        Else
            Err.Raise(1)
        End If
        rsLocal.Close()

        Exit Function

NumeroPerjudicadoAsistencia_Err:
        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()
        NumeroPerjudicadoAsistencia = "Error"
    End Function

    ' Procedure:    ReferenciasCruzadasSiniestros
    ' Objetivo:     Procesa las referencias cruzadas entre las aperturas importadas
    '               y las tablas de siniestros
    ' Retorno:      True o False si se provoca error (excluye adRecIntegrityViolation)
    '
    Public Function CruceReferenciasSiniestros(ByRef IdReferCompa As String, ByRef Codcia As String) As Boolean

        On Error GoTo CruceReferenciasSiniestros_Err

        CruceReferenciasSiniestros = True

        ' Declaraciones
        '
        Dim strSQL As String
        Dim objCmd As ADODB.Command
        Dim strSiniestro As String
        Dim numRegis As Integer
        Dim lngRegistros As Integer
        Dim strCia As String

        ' Actualizamos el c�digo de siniestro de la tabla de aperturas cruzando
        ' los campos de referencia externa con la tabla maestro de siniestro
        '
        strSQL = "UPDATE angel_t1 " & "SET    T1_CODSIN = snsinies.CODSIN, T1_ESTADO = 'P' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' +angel_t1.T1_REFER AND " & "       (angel_t1.T1_CODSIN = '' OR angel_t1.T1_CODSIN IS NULL or angel_t1.T1_CODSIN = 'No Existe') and " & "       T1_Codcia = '" & Codcia & "'"

        objCmd = New ADODB.Command
        objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        objCmd.CommandText = strSQL
        objCmd.ActiveConnection = claseBDLibrerias.BDWorkConnect
        objCmd.Execute(lngRegistros)

        ' Marcamos los que no han sido encontrados como ' No Existe'
        '
        strSQL = "UPDATE angel_t1 " & "SET    T1_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' +angel_t1.T1_REFER AND " & "       ((angel_t1.T1_CODSIN = '' OR angel_t1.T1_CODSIN IS NULL) and Angel_t1.T1_Codsin <> 'No Existe' ) and " & "       T1_Codcia = '" & Codcia & "'"

        objCmd = New ADODB.Command
        objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        objCmd.CommandText = strSQL
        objCmd.ActiveConnection = claseBDLibrerias.BDWorkConnect
        objCmd.Execute(lngRegistros)

        ' --------------------------------------------------------------------------
        '  JLL - 26/03/2004 Modificaci�n
        ' --------------------------------------------------------------------------

        '   Despues de comentarlo con Araceli decidimos realizar la siguiente
        '   modificaci�n:  Actualizamos el c�digo de siniestro en la tabla de pagos
        '                  Angel_T2 independientemente de que exista la apertura en
        '                  la tabla de aperturas

        strSQL = "UPDATE angel_t2 " & "SET    T2_CODSIN = snsinies.CODSIN " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' +angel_t2.T2_REFER AND " & "       (angel_t2.T2_CODSIN = '' OR angel_t2.T2_CODSIN IS NULL OR angel_t2.T2_CODSIN = 'No Existe') and " & "       T2_Codcia = '" & Codcia & "'"

        objCmd = New ADODB.Command
        objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        objCmd.CommandText = strSQL
        objCmd.ActiveConnection = claseBDLibrerias.BDWorkConnect
        objCmd.Execute(lngRegistros)

        ' Actualizamos marcamos los que no han sido encontrados como 'No Existe'
        '
        strSQL = "UPDATE angel_t2 " & "SET    T2_CODSIN = 'No Existe' " & "FROM   Snsinies " & "WHERE  Snsinies.Refext = '" & IdReferCompa & "' +angel_t2.T2_REFER AND " & "       ((angel_t2.T2_CODSIN = '' OR angel_t2.T2_CODSIN IS NULL) and angel_t2.T2_CODSIN <> 'No Existe') and " & "       T2_Codcia = '" & Codcia & "'"

        objCmd = New ADODB.Command
        objCmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        objCmd.CommandText = strSQL
        objCmd.ActiveConnection = claseBDLibrerias.BDWorkConnect
        objCmd.Execute(lngRegistros)

        ' ----------------------------------------------------------------------
        '  F�n de la modificaci�n de JLL - 26/03/2004
        ' ----------------------------------------------------------------------

        ' Realizamos el cruce entre Referencias y Siniestros para la
        ' tabla de Anulaciones de Siniestros
        '
        ' Cu�ndo tenga una rato modificar� todo el c�digo anterior
        ' porque da pena mirarlo y ademas puede ser mucho m�s r�pido
        '

        ' Abrimos la transacci�n
        '
        claseBDLibrerias.BDWorkConnect.BeginTrans()
        Transaccion = True

        ' Primero asignamos el siniestro a cada referencia
        '
        strSQL = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = SNSINIES.Codsin " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Refer = Snsinies.Refext and " & "       AnulacionesAsistencia.T5_Codcia = '" & Codcia & "'"
        claseBDLibrerias.BDWorkConnect.Execute(strSQL)

        ' Si no tienen siniestro abierto marcamos como 'No Existe'
        '
        strSQL = "UPDATE AnulacionesAsistencia " & "SET    AnulacionesAsistencia.T5_Codsin = 'No ExiSte' " & "FROM   Snsinies " & "WHERE  AnulacionesAsistencia.T5_Codsin is null and " & "       AnulacionesAsistencia.T5_Codcia = '" & Codcia & "'"
        claseBDLibrerias.BDWorkConnect.Execute(strSQL)

        ' Cerramos la transacci�n
        '
        claseBDLibrerias.BDWorkConnect.CommitTrans()
        Transaccion = False

        'UPGRADE_NOTE: El objeto objCmd no se puede destruir hasta que no se realice la recolecci�n de los elementos no utilizados. Haga clic aqu� para obtener m�s informaci�n: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        objCmd = Nothing
        CruceReferenciasSiniestros = True

        Exit Function

CruceReferenciasSiniestros_Err:
        CruceReferenciasSiniestros = False
        If claseBDLibrerias.BDWorkRecord.State = 1 Then claseBDLibrerias.BDWorkRecord.Close()
    End Function
End Class